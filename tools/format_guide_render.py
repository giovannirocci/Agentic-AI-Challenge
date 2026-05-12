"""Format-guide presentation renderer.

Extracts content from a source PPTX and rebuilds it following layout/format rules
from a brand guidelines PDF. Preserves all original text while applying new structure.

Run locally:
    python tools/unified_format_guide_render.py <source.pptx> <guidelines.pdf> <output.pptx>

Returns a new PPTX with source content restructured per guidelines.
"""
from ibm_watsonx_orchestrate.agent_builder.tools import tool, WXOFile, MultiFileConstraints

import argparse
import io
import json
import logging
import re
import xml.etree.ElementTree as ET
from pathlib import Path
from typing import Annotated, List

_log = logging.getLogger("format_guide_render")

# ─── XML namespaces ──────────────────────────────────────────────────────

NS = {
    'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
    'p': 'http://schemas.openxmlformats.org/presentationml/2006/main',
    'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
}


# ─── Content extraction from source PPTX (v2 Logic) ──────────────────────

def _emu_to_inches(value):
    """Convert PowerPoint EMU units to inches."""
    return value / 914400


def _safe_font_size(run, paragraph=None):
    """Try to get font size from run or paragraph."""
    try:
        if run.font.size:
            return run.font.size.pt
    except Exception:
        pass

    try:
        if paragraph and paragraph.font.size:
            return paragraph.font.size.pt
    except Exception:
        pass

    return None


def _shape_geometry(shape):
    """Return shape position and size in inches."""
    return {
        "x": _emu_to_inches(shape.left),
        "y": _emu_to_inches(shape.top),
        "w": _emu_to_inches(shape.width),
        "h": _emu_to_inches(shape.height),
    }

def _max_font_size_from_paragraphs(paragraphs):
    """Return the largest font size found in extracted paragraphs."""
    sizes = []

    for p in paragraphs:
        for r in p.get("runs", []):
            size = r.get("font_size")
            if size:
                sizes.append(size)

    return max(sizes) if sizes else None

def _extract_paragraphs_from_text_frame(text_frame):
    """
    Extract paragraphs while preserving:
    - paragraph text
    - bullet level
    - run-level formatting
    """
    paragraphs = []

    for p in text_frame.paragraphs:
        runs = []
        full_text_parts = []

        for r in p.runs:
            txt = r.text or ""

            if txt:
                full_text_parts.append(txt)

            runs.append({
                "text": txt,
                "bold": bool(r.font.bold) if r.font.bold is not None else None,
                "italic": bool(r.font.italic) if r.font.italic is not None else None,
                "font_size": _safe_font_size(r, p),
                "font_name": r.font.name,
            })

        paragraph_text = (p.text or "".join(full_text_parts)).strip()

        if paragraph_text:
            paragraphs.append({
                "text": paragraph_text,
                "level": p.level,
                "runs": runs,
            })

    return paragraphs

def _classify_text_shape(shape, paragraphs, slide_width, slide_height):
    """
    Classify a text shape as title, subtitle, body, caption, footer, etc.
    Uses placeholder metadata first, then position/font heuristics.
    """
    text = " ".join(p["text"] for p in paragraphs).strip()

    if not text:
        return "empty"

    # 1. Prefer actual PowerPoint placeholder information
    try:
        if shape.is_placeholder:
            ph_type = str(shape.placeholder_format.type).lower()

            if "title" in ph_type:
                return "title"
            if "subtitle" in ph_type:
                return "subtitle"
            if "footer" in ph_type:
                return "footer"
            if "body" in ph_type or "object" in ph_type:
                return "body"
    except Exception:
        pass

    # 2. Use position and font size as fallback
    geom = _shape_geometry(shape)
    slide_h = _emu_to_inches(slide_height)

    font_sizes = []
    for p in paragraphs:
        for r in p["runs"]:
            if r.get("font_size"):
                font_sizes.append(r["font_size"])

    max_font_size = max(font_sizes) if font_sizes else None

    # Top area + large font usually means title
    if geom["y"] < 1.2 and max_font_size and max_font_size >= 22:
        return "title"

    # Bottom area is often footer
    if geom["y"] > slide_h - 0.8:
        return "footer"

    # Very small text is often caption/legal/footer
    if max_font_size and max_font_size <= 11:
        return "caption"

    return "body"

def _extract_text_shape_v2(shape, slide_width, slide_height):
    """Extract a text shape with text, paragraph structure, role and geometry."""
    if not hasattr(shape, "text_frame"):
        return None

    try:
        paragraphs = _extract_paragraphs_from_text_frame(shape.text_frame)
    except Exception:
        return None

    if not paragraphs:
        return None

    role = _classify_text_shape(shape, paragraphs, slide_width, slide_height)
    max_font_size = _max_font_size_from_paragraphs(paragraphs)

    placeholder_type = ""
    try:
        if shape.is_placeholder:
            placeholder_type = str(shape.placeholder_format.type)
    except Exception:
        placeholder_type = ""

    return {
        "type": "text",
        "role": role,
        "text": "\n".join(p["text"] for p in paragraphs),
        "paragraphs": paragraphs,
        "max_font_size": max_font_size,
        "placeholder_type": placeholder_type,
        **_shape_geometry(shape),
    }

def _extract_table_shape_v2(shape):
    """Extract table content separately instead of turning it into bullets."""
    try:
        if not shape.has_table:
            return None
    except Exception:
        return None

    rows = []

    for row in shape.table.rows:
        row_values = []
        for cell in row.cells:
            row_values.append(cell.text.strip())
        rows.append(row_values)

    return {
        "type": "table",
        "rows": rows,
        **_shape_geometry(shape),
    }

def _extract_image_shape_v2(shape, image_counter):
    """
    Extract image metadata and bytes.
    For now we keep the image bytes in the intermediate structure.
    Rendering images can be improved later.
    """
    try:
        image = shape.image
    except Exception:
        return None

    image_name = f"image_{image_counter}.{image.ext}"

    return {
        "type": "image",
        "image_name": image_name,
        "image_bytes": image.blob,
        **_shape_geometry(shape),
    }

def _extract_shape_v2(shape, slide_width, slide_height, image_counter):
    """
    Extract one shape.
    Order matters:
    - table first
    - image second
    - text third
    """
    table = _extract_table_shape_v2(shape)
    if table:
        return table, image_counter

    image = _extract_image_shape_v2(shape, image_counter)
    if image:
        return image, image_counter + 1

    text = _extract_text_shape_v2(shape, slide_width, slide_height)
    if text:
        return text, image_counter

    return None, image_counter

def _extract_slide_content_v2(slide, slide_index, prs):
    """Extract one slide as a structured collection of elements."""
    elements = []
    image_counter = 1

    for shape in slide.shapes:
        element, image_counter = _extract_shape_v2(
            shape=shape,
            slide_width=prs.slide_width,
            slide_height=prs.slide_height,
            image_counter=image_counter,
        )

        if element:
            elements.append(element)

    # Find best title candidate
    title_candidates = [
        e for e in elements
        if e.get("type") == "text" and e.get("role") == "title"
    ]

    if title_candidates:
        title = sorted(title_candidates, key=lambda e: (e["y"], -e["w"]))[0]["text"]
    else:
        # fallback: first text element near top
        text_elements = [e for e in elements if e.get("type") == "text"]
        if text_elements:
            title = sorted(text_elements, key=lambda e: e["y"])[0]["text"]
        else:
            title = ""

    notes = ""
    try:
        if slide.has_notes_slide and slide.notes_slide.notes_text_frame:
            notes = slide.notes_slide.notes_text_frame.text.strip()
    except Exception:
        notes = ""

    return {
        "slide_index": slide_index,
        "title": title,
        "layout_name": slide.slide_layout.name if slide.slide_layout else "",
        "elements": elements,
        "notes": notes,
    }

def _extract_all_content_v2(pptx_bytes):
    """Extract full presentation structure from PPTX."""
    from pptx import Presentation

    prs = Presentation(io.BytesIO(pptx_bytes))

    slides = []

    for idx, slide in enumerate(prs.slides, start=1):
        slide_data = _extract_slide_content_v2(slide, idx, prs)
        slides.append(slide_data)

    return {
        "slide_width": _emu_to_inches(prs.slide_width),
        "slide_height": _emu_to_inches(prs.slide_height),
        "slide_count": len(slides),
        "slides": slides,
    }


# ─── Guidelines PDF parsing ──────────────────────────────────────────────

def _parse_guidelines_pdf(pdf_bytes):
    """Parse brand guidelines PDF to extract layout rules and style specs.

    Returns: {
        "layouts": [
            { "name": "title", "description": "...", "max_bullets": 5 },
            ...
        ],
        "colors": { "primary": "#...", ... },
        "fonts": { "title": "Arial", ... },
        "spacing": { "margin_top": 0.5, ... },
    }
    """
    import pdfplumber

    text_content = ""
    try:
        with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
            for page in pdf.pages:
                text_content += page.extract_text() or ""
    except Exception as e:
        raise ValueError(f"Failed to parse PDF: {e}")

    # Extract color specifications (look for hex or RGB patterns)
    colors = {}
    hex_matches = re.findall(r'#([0-9A-Fa-f]{6})', text_content)
    if hex_matches:
        colors['primary'] = f"#{hex_matches[0]}"
        if len(hex_matches) > 1:
            colors['accent'] = f"#{hex_matches[1]}"

    # Extract font names (look for common patterns)
    fonts = {}
    font_patterns = re.findall(r'\b(Arial|Helvetica|Times|Calibri|Verdana|Georgia|Roboto|Ubuntu)\b', text_content, re.I)
    if font_patterns:
        fonts['body'] = font_patterns[0]
        fonts['title'] = font_patterns[0]

    # Infer layout patterns from PDF text
    layouts = _infer_layouts_from_text(text_content)

    # Extract brand logos (full wordmark + icon variant) from the PDF
    logos = _extract_logos_from_pdf(pdf_bytes)

    return {
        "layouts": layouts,
        "colors": colors,
        "fonts": fonts,
        "spacing": {},  # Can be enhanced with more sophisticated parsing
        "logos": logos,
        "raw_text": text_content[:1000],  # Store excerpt for debugging
    }


def _extract_logos_from_pdf(pdf_bytes):
    """Pull brand logos out of the guidelines PDF.

    Scans the first few pages, ranks raster images, and picks:
      - "full": wide wordmark (aspect >= 1.6, or tall <= 0.6 treated as wide)
      - "icon": square-ish mark (0.6 <= aspect < 1.6)

    If no raster logos are found (common — many brand PDFs use vector logos),
    falls back to vector-path clustering and finally to rendering the top half
    of page 1.

    Returns:
        {"full": (bytes, ext)|None, "icon": (bytes, ext)|None}
    """
    report = {
        "pymupdf_available": False,
        "pdf_opened": False,
        "page_count": 0,
        "pages_scanned": 0,
        "raster_images_found": 0,
        "candidates": [],   # list of dicts: page, w, h, aspect, ext, reason
        "selected_full": None,
        "selected_icon": None,
        "used_vector_fallback": False,
        "errors": [],
    }

    import fitz  # PyMuPDF

    MAX_PAGES = 30
    MIN_DIM = 60
    MAX_AREA = 1_500_000

    candidates = []  # tuples used for selection

    try:
        doc = fitz.open(stream=pdf_bytes, filetype="pdf")
        report["pdf_opened"] = True
        report["page_count"] = doc.page_count
    except Exception as e:
        report["errors"].append(f"failed to open PDF: {e}")
        return {"full": None, "icon": None}

    try:
        pages_to_scan = min(MAX_PAGES, doc.page_count)
        report["pages_scanned"] = pages_to_scan
        for page_idx in range(pages_to_scan):
            page = doc[page_idx]
            page_images = page.get_images(full=True)
            report["raster_images_found"] += len(page_images)
            for img_info in page_images:
                xref = img_info[0]
                try:
                    base = doc.extract_image(xref)
                except Exception as e:
                    report["candidates"].append({
                        "page": page_idx + 1, "w": 0, "h": 0, "aspect": 0,
                        "ext": "?", "reason": f"extract_image failed: {e}",
                    })
                    continue
                img_bytes = base.get("image")
                ext = base.get("ext", "png")
                w = base.get("width", 0)
                h = base.get("height", 0)
                aspect = (w / h) if h else 0

                reason = None
                if not img_bytes:
                    reason = "no bytes returned"
                elif w < MIN_DIM or h < MIN_DIM:
                    reason = f"too small (min dim {MIN_DIM})"
                elif w * h > MAX_AREA:
                    reason = f"too large (area > {MAX_AREA})"

                report["candidates"].append({
                    "page": page_idx + 1, "w": w, "h": h, "aspect": round(aspect, 2),
                    "ext": ext, "size_bytes": len(img_bytes) if img_bytes else 0,
                    "reason": reason or "accepted",
                })

                if reason:
                    continue

                score = (MAX_PAGES - page_idx) * 1000 + (w * h) / 1000
                candidates.append((score, aspect, img_bytes, ext, w, h, page_idx))

        # Vector clustering: most charter PDFs store logos as outlined vector paths.
        if len(candidates) < 3:
            candidates.extend(_find_vector_logo_candidates(doc, pages_to_scan, report))
    finally:
        doc.close()

    if not candidates:
        _log.warning("logo extraction: no candidates. report=%s", report)
        return {"full": None, "icon": None}

    wide = [c for c in candidates if c[1] >= 1.6 or c[1] <= 0.6]
    square = [c for c in candidates if 0.6 < c[1] < 1.4]
    wide.sort(key=lambda c: c[0], reverse=True)
    square.sort(key=lambda c: c[0], reverse=True)

    full = (wide[0][2], wide[0][3]) if wide else None
    icon = (square[1][2], square[1][3]) if square else None
    if full and not icon:
        icon = full
    elif icon and not full:
        full = icon

    if wide:
        report["selected_full"] = {"page": wide[0][6] + 1, "w": wide[0][4], "h": wide[0][5], "ext": wide[0][3]}
    if square:
        report["selected_icon"] = {"page": square[0][6] + 1, "w": square[0][4], "h": square[0][5], "ext": square[0][3]}

    _log.warning("logo extraction report: %s", report)
    return {"full": full, "icon": icon}


def _find_vector_logo_candidates(doc, pages_to_scan, report):
    """Find logos by clustering vector paths on early pages."""
    import fitz

    MIN_PATHS = 8
    MIN_DIM_PT = 30
    MAX_PAGE_AREA_FRACTION = 0.5
    CLUSTER_GAP_PT = 10
    MAX_DRAWINGS_PER_PAGE = 4000
    PAD_PT = 6
    TARGET_LONG_EDGE_PX = 1200

    candidates = []

    for page_idx in range(min(pages_to_scan, doc.page_count)):
        page = doc[page_idx]
        try:
            drawings = page.get_drawings()
        except Exception as e:
            report["errors"].append(f"get_drawings page {page_idx+1}: {e}")
            continue

        if not drawings or len(drawings) > MAX_DRAWINGS_PER_PAGE:
            continue

        rects = []
        for d in drawings:
            r = d.get("rect")
            if r is None or r.width < 1 or r.height < 1:
                continue
            rects.append(fitz.Rect(r))

        if len(rects) < MIN_PATHS:
            continue

        clusters = _cluster_rects(rects, gap=CLUSTER_GAP_PT)
        page_area = page.rect.width * page.rect.height

        for cluster_idxs in clusters:
            n_paths = len(cluster_idxs)
            if n_paths < MIN_PATHS:
                continue

            bbox = fitz.Rect(rects[cluster_idxs[0]])
            for i in cluster_idxs[1:]:
                bbox |= rects[i]

            if bbox.width < MIN_DIM_PT or bbox.height < MIN_DIM_PT:
                continue
            if bbox.width * bbox.height > page_area * MAX_PAGE_AREA_FRACTION:
                continue

            aspect = bbox.width / bbox.height
            if aspect < 0.2 or aspect > 6.0:
                continue

            padded = fitz.Rect(
                max(page.rect.x0, bbox.x0 - PAD_PT),
                max(page.rect.y0, bbox.y0 - PAD_PT),
                min(page.rect.x1, bbox.x1 + PAD_PT),
                min(page.rect.y1, bbox.y1 + PAD_PT),
            )

            long_edge_pt = max(padded.width, padded.height)
            dpi = min(300, max(150, int(TARGET_LONG_EDGE_PX * 72 / long_edge_pt))) if long_edge_pt else 200

            try:
                pix = page.get_pixmap(clip=padded, dpi=dpi, alpha=False)
                png_bytes = pix.tobytes("png")
            except Exception as e:
                report["errors"].append(f"render vector cluster on page {page_idx+1}: {e}")
                continue

            score = (30 - page_idx) * 1000 + n_paths * 10
            candidates.append((score, aspect, png_bytes, "png", pix.width, pix.height, page_idx))

    if candidates:
        report["used_vector_fallback"] = True

    return candidates


def _cluster_rects(rects, gap=10.0):
    """Union-find clustering: rects whose expanded bboxes overlap merge."""
    import fitz

    n = len(rects)
    if n == 0:
        return []

    parent = list(range(n))

    def find(i):
        while parent[i] != i:
            parent[i] = parent[parent[i]]
            i = parent[i]
        return i

    expanded = [fitz.Rect(r.x0 - gap, r.y0 - gap, r.x1 + gap, r.y1 + gap) for r in rects]

    for i in range(n):
        for j in range(i + 1, n):
            if expanded[i].intersects(expanded[j]):
                ri, rj = find(i), find(j)
                if ri != rj:
                    parent[ri] = rj

    groups = {}
    for i in range(n):
        groups.setdefault(find(i), []).append(i)

    return list(groups.values())


def _infer_layouts_from_text(text):
    """Infer slide layout patterns from guideline text."""
    layouts = []

    if re.search(r'title.*slide|slide.*title', text, re.I):
        layouts.append({
            "name": "title",
            "description": "Title slide with main title and subtitle",
            "max_bullets": 1,
        })

    if re.search(r'content.*slide|bullet|list', text, re.I):
        layouts.append({
            "name": "content",
            "description": "Content slide with title and bulleted list",
            "max_bullets": 5,
        })

    if re.search(r'two.*column|column|compare|side.?by.?side', text, re.I):
        layouts.append({
            "name": "two_column",
            "description": "Two-column comparison layout",
            "max_bullets": 3,
        })

    if re.search(r'closing|conclusion|thank|end', text, re.I):
        layouts.append({
            "name": "closing",
            "description": "Closing slide",
            "max_bullets": 3,
        })

    if not layouts:
        layouts = [
            {"name": "title", "description": "Title slide", "max_bullets": 1},
            {"name": "content", "description": "Content slide", "max_bullets": 5},
            {"name": "closing", "description": "Closing slide", "max_bullets": 3},
        ]

    return layouts


# ─── Rendering (Structure Preserving Logic) ──────────────────────────────

def _hex_to_rgb(hex_color):
    """Convert hex color to RGBColor."""
    from pptx.dml.color import RGBColor

    hex_color = hex_color.lstrip('#')
    if len(hex_color) != 6:
        hex_color = "000000"
    try:
        return RGBColor(int(hex_color[0:2], 16), int(hex_color[2:4], 16), int(hex_color[4:6], 16))
    except ValueError:
        return RGBColor(0, 0, 0)


def _clamp(value, minimum, maximum):
    """Clamp a numeric value into a safe range."""
    if value is None:
        return minimum
    return max(minimum, min(maximum, value))


def _get_text_color_for_role(role, colors):
    """Choose brand-aware text color based on element role."""
    if role == "title":
        return _hex_to_rgb(colors.get("primary", "#24135F"))

    if role in ["footer", "caption"]:
        return _hex_to_rgb(colors.get("grey", "#B1B3B3"))

    return _hex_to_rgb(colors.get("text", "#1D1D1D"))


def _place_full_logo(slide, prs, logo):
    """Place a large logo near the top right corner of a title/section/closing slide."""
    if not logo:
        return
    from pptx.util import Inches

    img_bytes, _ext = logo
    target_w = Inches(3.2)
    slide_w = prs.slide_width
    left = slide_w - target_w - Inches(0.3)
    top = Inches(0.5)
    try:
        slide.shapes.add_picture(io.BytesIO(img_bytes), left, top, width=target_w)
    except Exception as e:
        _log.warning("failed to place full logo: %s", e)


def _place_icon_logo(slide, prs, logo):
    """Place a small icon in the bottom-left corner of a content slide."""
    if not logo:
        return
    from pptx.util import Inches

    img_bytes, _ext = logo
    target_w = Inches(0.7)
    slide_w = prs.slide_width
    left = Inches(0.3)
    bottom = prs.slide_height - Inches(1.0)

    try:
        slide.shapes.add_picture(io.BytesIO(img_bytes), left, bottom, width=target_w)
    except Exception as e:
        _log.warning("failed to place icon logo: %s", e)


def _render_text_element_preserve(slide, element, colors, fonts):
    """
    Render one extracted text element at approximately the same position
    as in the source slide, while applying brand font/color.
    """
    from pptx.util import Inches, Pt

    role = element.get("role", "body")
    paragraphs = element.get("paragraphs", [])

    if not paragraphs:
        return

    x = element.get("x", 0)
    y = element.get("y", 0)
    w = element.get("w", 1)
    h = element.get("h", 1)

    # Avoid invalid boxes
    if w <= 0 or h <= 0:
        return

    box = slide.shapes.add_textbox(
        Inches(x),
        Inches(y),
        Inches(w),
        Inches(h),
    )

    tf = box.text_frame
    tf.clear()
    tf.word_wrap = True

    # Small margins make extracted layouts less cramped
    tf.margin_left = Inches(0.05)
    tf.margin_right = Inches(0.05)
    tf.margin_top = Inches(0.03)
    tf.margin_bottom = Inches(0.03)

    title_font = fonts.get("title", "Arial")
    body_font = fonts.get("body", "Arial")

    base_font_size = element.get("max_font_size")

    if base_font_size is None:
        if role == "title":
            base_font_size = 30
        elif role in ["footer", "caption"]:
            base_font_size = 9
        else:
            base_font_size = 16

    # Keep original font-size logic, but avoid extreme values
    font_size = _clamp(base_font_size, 7, 46)

    for i, para in enumerate(paragraphs):
        if i == 0:
            p = tf.paragraphs[0]
        else:
            p = tf.add_paragraph()

        text = para.get("text", "").strip()
        if not text:
            continue

        p.text = text

        try:
            p.level = int(para.get("level", 0))
        except Exception:
            p.level = 0

        p.font.name = title_font if role == "title" else body_font
        p.font.size = Pt(font_size)
        p.font.color.rgb = _get_text_color_for_role(role, colors)

        if role == "title":
            p.font.bold = True


def _render_image_element_preserve(slide, element):
    """Render extracted image at approximately the same position."""
    from pptx.util import Inches

    image_bytes = element.get("image_bytes")
    if not image_bytes:
        return

    x = element.get("x", 0)
    y = element.get("y", 0)
    w = element.get("w", 1)
    h = element.get("h", 1)

    if w <= 0 or h <= 0:
        return

    try:
        slide.shapes.add_picture(
            io.BytesIO(image_bytes),
            Inches(x),
            Inches(y),
            width=Inches(w),
            height=Inches(h),
        )
    except Exception:
        # If an image fails, do not crash the entire deck generation
        pass


def _render_table_element_preserve(slide, element, colors, fonts):
    """Render extracted table content at approximately the same position."""
    from pptx.util import Inches, Pt

    rows = element.get("rows", [])
    if not rows:
        return

    row_count = len(rows)
    col_count = max(len(row) for row in rows) if rows else 0

    if row_count <= 0 or col_count <= 0:
        return

    x = element.get("x", 0)
    y = element.get("y", 0)
    w = element.get("w", 1)
    h = element.get("h", 1)

    if w <= 0 or h <= 0:
        return

    try:
        table_shape = slide.shapes.add_table(
            row_count,
            col_count,
            Inches(x),
            Inches(y),
            Inches(w),
            Inches(h),
        )

        table = table_shape.table
        body_font = fonts.get("body", "Arial")
        text_color = _hex_to_rgb(colors.get("text", "#1D1D1D"))

        for r_idx, row in enumerate(rows):
            for c_idx in range(col_count):
                value = row[c_idx] if c_idx < len(row) else ""
                cell = table.cell(r_idx, c_idx)
                cell.text = value

                for paragraph in cell.text_frame.paragraphs:
                    paragraph.font.name = body_font
                    paragraph.font.size = Pt(10)
                    paragraph.font.color.rgb = text_color

    except Exception:
        pass


def _render_slides_preserve_structure(extracted_deck, guidelines):
    """
    Render one output slide per source slide.

    This avoids:
    - unwanted '(cont'd)' slides
    - bullet chunk splitting
    - dropping images
    - dropping tables
    - destroying source slide structure
    """
    from pptx import Presentation
    from pptx.util import Inches
    from pptx.dml.color import RGBColor

    prs = Presentation()

    # Keep the original deck dimensions if available
    source_w = extracted_deck.get("slide_width", 13.333)
    source_h = extracted_deck.get("slide_height", 7.5)

    prs.slide_width = Inches(source_w)
    prs.slide_height = Inches(source_h)

    colors = guidelines.get("colors", {})
    fonts = guidelines.get("fonts", {})

    # Extract parsed logos
    logos = guidelines.get("logos") or {}
    full_logo = logos.get("full")  # (bytes, ext) | None
    icon_logo = logos.get("icon")  # (bytes, ext) | None

    # Provide stronger defaults if PDF parsing misses labels
    colors.setdefault("primary", "#24135F")
    colors.setdefault("accent", "#D0006F")
    colors.setdefault("grey", "#B1B3B3")
    colors.setdefault("text", "#1D1D1D")

    fonts.setdefault("title", "Arial")
    fonts.setdefault("body", "Arial")

    slides = extracted_deck.get("slides", [])
    total_slides = len(slides)

    for i, source_slide in enumerate(slides):
        blank_layout = prs.slide_layouts[6]
        slide = prs.slides.add_slide(blank_layout)

        # White background
        background = slide.background
        fill = background.fill
        fill.solid()
        fill.fore_color.rgb = RGBColor(255, 255, 255)

        elements = source_slide.get("elements", [])

        # Render images first so text can appear above them
        for element in elements:
            if element.get("type") == "image":
                _render_image_element_preserve(slide, element)

        # Render tables second
        for element in elements:
            if element.get("type") == "table":
                _render_table_element_preserve(slide, element, colors, fonts)

        # Render text last
        for element in elements:
            if element.get("type") == "text":
                _render_text_element_preserve(slide, element, colors, fonts)

        # Apply Logo Placement
        # Treat the first slide and the last slide (if there's more than 1) as title/closing
        is_title_or_closing = (i == 0) or (i == total_slides - 1 and total_slides > 1)
        if is_title_or_closing:
            _place_full_logo(slide, prs, full_logo)
        else:
            _place_icon_logo(slide, prs, icon_logo)

    buf = io.BytesIO()
    prs.save(buf)
    return buf.getvalue()


def format_guide_render_internal(source_pptx_bytes, guidelines_pdf_bytes):
    """Core formatting logic."""

    # Extract content from source using improved extractor
    content = _extract_all_content_v2(source_pptx_bytes)

    # Parse guidelines (incorporating the enhanced logic from format_guide_render)
    guidelines = _parse_guidelines_pdf(guidelines_pdf_bytes)

    # Render directly from rich extracted structure
    output_bytes = _render_slides_preserve_structure(content, guidelines)

    return output_bytes


# ─── Tool wrapper ───────────────────────────────────────────────────────

@tool(
    name="format_guide_render",
    description=(
        "Reformats a presentation to match brand guidelines. "
        "Upload exactly two files: one .pptx (source deck) and one .pdf (guidelines). "
        "Order does not matter — the tool detects which is which. "
        "Returns the reformatted .pptx as a downloadable file."
    ),
)
def format_guide_render(
    files: Annotated[
        List[WXOFile],
        MultiFileConstraints(
            min_files=2,
            max_files=2,
            accepted_file_extensions=["pptx", "pdf"],
        ),
    ],
) -> bytes:
    """Reformat presentation per brand guidelines.

    Accepts the two uploaded files in either order. Discriminates by magic bytes
    so the LLM cannot break the call by swapping or duplicating refs.
    """
    if len(files) != 2:
        raise ValueError(f"Expected exactly 2 files, got {len(files)}.")

    blobs = []
    for i, ref in enumerate(files):
        data = WXOFile.get_content(ref)
        _log.warning(
            "file[%d]: type=%s len=%s head=%r",
            i, type(data).__name__,
            len(data) if hasattr(data, "__len__") else "?",
            data[:16] if isinstance(data, (bytes, bytearray)) else str(data)[:120],
        )
        if not isinstance(data, (bytes, bytearray)):
            raise TypeError(
                f"file[{i}]: WXOFile.get_content returned {type(data).__name__}, not bytes"
            )
        blobs.append(bytes(data))

    def kind(buf: bytes) -> str:
        if buf.startswith(b"%PDF"):
            return "pdf"
        if buf.startswith(b"PK\x03\x04"):
            return "pptx"
        return "unknown"

    kinds = [kind(b) for b in blobs]
    if sorted(kinds) != ["pdf", "pptx"]:
        raise ValueError(
            f"Need one distinct .pptx and one distinct .pdf. Got kinds={kinds}, "
            f"sizes={[len(b) for b in blobs]}. Please re-upload one of each."
        )

    pptx_bytes = blobs[kinds.index("pptx")]
    pdf_bytes = blobs[kinds.index("pdf")]
    return format_guide_render_internal(pptx_bytes, pdf_bytes)


# ─── CLI for local testing ──────────────────────────────────────────────

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Reformat PPTX per brand guidelines PDF")
    parser.add_argument("source_pptx", help="Path to source presentation")
    parser.add_argument("guidelines_pdf", help="Path to brand guidelines PDF")
    parser.add_argument("output_pptx", help="Where to write reformatted .pptx")
    args = parser.parse_args()

    source_bytes = Path(args.source_pptx).read_bytes()
    guidelines_bytes = Path(args.guidelines_pdf).read_bytes()

    output_bytes = format_guide_render_internal(source_bytes, guidelines_bytes)

    out_path = Path(args.output_pptx)
    out_path.parent.mkdir(parents=True, exist_ok=True)
    out_path.write_bytes(output_bytes)
    print(f"Wrote {out_path} ({len(output_bytes):,} bytes)")