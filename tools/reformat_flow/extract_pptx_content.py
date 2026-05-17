"""
extract_pptx_content.py --- Extract structured content from PPTX files.
─────────────────────

Given a PPTX file, extract a structured representation of its content, including:
- Slide-level information (title, layout, notes)
- Shape-level information (type, role, text, geometry)
- Image content (as bytes, with metadata)
"""

import argparse
import io
import json
from pathlib import Path


# ─── Content extraction from source PPTX  ──────────────────────

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

def _extract_text_shape(shape, slide_width, slide_height):
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

def _extract_table_shape(shape):
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

def _extract_image_shape(shape, image_counter):
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

def _extract_shape(shape, slide_width, slide_height, image_counter):
    """
    Extract one shape.
    Order matters:
    - table first
    - image second
    - text third
    """
    table = _extract_table_shape(shape)
    if table:
        return table, image_counter

    image = _extract_image_shape(shape, image_counter)
    if image:
        return image, image_counter + 1

    text = _extract_text_shape(shape, slide_width, slide_height)
    if text:
        return text, image_counter

    return None, image_counter

def _extract_slide_content(slide, slide_index, prs):
    """Extract one slide as a structured collection of elements."""
    elements = []
    image_counter = 1

    for shape in slide.shapes:
        element, image_counter = _extract_shape(
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

def extract_all_content(pptx_bytes):
    """Extract full presentation structure from PPTX."""
    from pptx import Presentation

    prs = Presentation(io.BytesIO(pptx_bytes))

    slides = []

    for idx, slide in enumerate(prs.slides, start=1):
        slide_data = _extract_slide_content(slide, idx, prs)
        slides.append(slide_data)

    return {
        "slide_width": _emu_to_inches(prs.slide_width),
        "slide_height": _emu_to_inches(prs.slide_height),
        "slide_count": len(slides),
        "slides": slides,
    }


# ──────────── CLI ─────────────────────────────────────────────────────

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Extract content from PPTX file.")
    parser.add_argument("pptx_path", type=Path, help="Path to the PPTX file")
    parser.add_argument("output_json", type=Path, help="Path to output JSON file")

    args = parser.parse_args()

    pptx_bytes = args.pptx_path.read_bytes()
    content = extract_all_content(pptx_bytes)

    with open(args.output_json, "w", encoding="utf-8") as f:
        json.dump(content, f, indent=2)