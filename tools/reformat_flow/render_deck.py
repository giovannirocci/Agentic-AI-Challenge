"""
Combine extracted brand guidelines with extracted PPTX content to render a reformatted deck
that preserves the original structure and incorporates brand elements.
"""

import argparse
import io
import logging
import os
from pathlib import Path

_log = logging.getLogger("format_guide_render")


# ─── Color / font key normalisation ─────────────────────────────────────────
#
# extract_guidelines() returns:
#   guidelines["colors"] = {"primary": "#hex", "secondary": "#hex", "others": [...]}
#   guidelines["fonts"] = {"primary": "FontName", "fallback": "FontName"}

def _normalise_guidelines(guidelines: dict) -> tuple[dict, dict]:
    """Return (colors, fonts) with the keys the renderer expects."""
    raw_colors = guidelines.get("colors") or {}
    raw_fonts= guidelines.get("fonts")  or {}

    colors = {
        "primary": raw_colors.get("primary",   "#7364AC"),
        "secondary": raw_colors.get("secondary", "#A15F83"),
        "text":"#1D1D1D",
        "grey": raw_colors["others"][0] if raw_colors.get("others") else "#B1B3B3",
    }

    fonts = {
        "title": raw_fonts.get("primary",  raw_fonts.get("fallback", "Arial")),
        "body":  raw_fonts.get("fallback", "Arial"),
    }

    return colors, fonts


# ─── Helpers ─────────────────────────────────────────────────────────────────

def _hex_to_rgb(hex_color):
    from pptx.dml.color import RGBColor
    hex_color = (hex_color or "000000").lstrip("#")
    if len(hex_color) != 6:
        hex_color = "000000"
    try:
        return RGBColor(int(hex_color[0:2], 16),
                        int(hex_color[2:4], 16),
                        int(hex_color[4:6], 16))
    except ValueError:
        return RGBColor(0, 0, 0)


def _crop_image_bytes(img_bytes, left_frac=0.06, right_frac=0.06):
    """Crop a fraction off the left/right sides of an image (bytes -> bytes).

    left_frac and right_frac are proportions of the image width to remove.
    Falls back to returning the original bytes if Pillow isn't available or on error.
    """
    try:
        from PIL import Image
        buf = io.BytesIO(img_bytes)
        img = Image.open(buf)
        w, h = img.size
        left = int(w * left_frac)
        right = w - int(w * right_frac)
        if left >= right:
            return img_bytes
        cropped = img.crop((left, 0, right, h))
        out = io.BytesIO()
        fmt = img.format or "PNG"
        # Preserve alpha when present by saving to PNG
        if fmt.upper() not in ("PNG", "JPEG", "JPG", "WEBP"):
            fmt = "PNG"
        if fmt.upper() in ("JPEG", "JPG") and cropped.mode in ("RGBA", "LA"):
            cropped = cropped.convert("RGB")
        cropped.save(out, format=fmt)
        return out.getvalue()
    except Exception:
        return img_bytes


def _clamp(value, minimum, maximum):
    if value is None:
        return minimum
    return max(minimum, min(maximum, value))


def _get_text_color_for_role(role, colors):
    if role == "title":
        return _hex_to_rgb(colors.get("primary", "#24135F"))
    if role in ("footer", "caption"):
        return _hex_to_rgb(colors.get("grey", "#B1B3B3"))
    return _hex_to_rgb(colors.get("text", "#1D1D1D"))


# ─── Brand chrome ─────────────────────────────────────────────────────────────

def _add_accent_bar(slide, prs, colors):
    """Thin secondary-colour bar at the very top of content slides."""
    from pptx.util import Inches, Emu
    import lxml.etree as etree

    bar = slide.shapes.add_shape(
        1,          # MSO_SHAPE_TYPE.RECTANGLE
        Emu(0), Emu(0),
        prs.slide_width,
        Inches(0.12),
    )
    bar.fill.solid()
    bar.fill.fore_color.rgb = _hex_to_rgb(colors["secondary"])
    bar.line.fill.background()

    # Push the bar behind all other shapes so it doesn't cover text
    sp_tree = slide.shapes._spTree
    sp_tree.remove(bar._element)
    sp_tree.insert(2, bar._element)          # index 2 = behind placeholders


def _add_footer_band(slide, prs, colors, fonts, slide_number):
    """Draw a small filled square with the slide number centered inside it."""
    from pptx.util import Inches, Pt
    from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
    from pptx.dml.color import RGBColor

    box_size = Inches(0.44)
    margin = Inches(0.12)

    left = prs.slide_width - box_size - margin
    top = prs.slide_height - box_size - margin

    try:
        sq = slide.shapes.add_shape(1, left, top, box_size, box_size)
        sq.fill.solid()
        sq.fill.fore_color.rgb = _hex_to_rgb(colors.get("secondary", "#D0006F"))
        sq.line.fill.background()

        tb = sq.text_frame
        tb.clear()
        tb.margin_left = Inches(0)
        tb.margin_right = Inches(0)
        tb.margin_top = Inches(0)
        tb.margin_bottom = Inches(0)
        tb.vertical_anchor = MSO_ANCHOR.MIDDLE

        p = tb.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER
        run = p.add_run()
        run.text = str(slide_number)
        run.font.size = Pt(12)
        run.font.bold = True
        run.font.color.rgb = RGBColor(255, 255, 255)
        run.font.name = fonts.get("title", "Arial")
    except Exception:
        # Non-critical: if drawing fails, don't crash the renderer
        pass


# ─── Logo placement ───────────────────────────────────────────────────────────

def _logo_bytes(logo_candidate):
    """
    extract_guidelines() returns logo candidates as dicts with an "img_bytes" key,
    NOT as (bytes, ext) tuples.  Accept both shapes so the renderer is robust.
    """
    if logo_candidate is None:
        return None
    if isinstance(logo_candidate, dict):
        return logo_candidate.get("img_bytes")
    # legacy (bytes, ext) tuple
    if isinstance(logo_candidate, (tuple, list)) and len(logo_candidate) >= 1:
        return logo_candidate[0]
    return None


def _place_titles_logo(slide, prs, logo, crop: bool = False):
    """Large logo near the top-right corner (dark/title slides).

    If `crop` is True, apply a slight crop to the inserted picture
    (useful for dark-background full-wordmark variants).
    """
    img = _logo_bytes(logo)
    if not img:
        return
    from pptx.util import Inches
    target_w = Inches(2.8)
    left = prs.slide_width - target_w - Inches(0.35)
    top = Inches(0.3)
    try:
        pic = slide.shapes.add_picture(io.BytesIO(img), left, top, width=target_w)
        if crop:
            try:
                pic.crop_left = 0.05
                pic.crop_right = 0.05
                pic.crop_top = 0.02
                pic.crop_bottom = 0.02
            except Exception:
                pass
    except Exception as e:
        _log.warning("failed to place full logo: %s", e)


def _place_full_logo(slide, prs, logo):
    """Small full logo in the bottom-left corner (content slides)."""
    img = _logo_bytes(logo)
    if not img:
        return
    from pptx.util import Inches
    target_w = Inches(1.4)
    left = Inches(0.3)
    top = prs.slide_height - Inches(0.3) - target_w * 0.3
    try:
        slide.shapes.add_picture(io.BytesIO(img), left, top, width=target_w)
    except Exception as e:
        _log.warning("failed to place full logo: %s", e)


def _place_icon_logo(slide, prs, logo):
    """Small icon in the top-right corner (content slides)."""
    img = _logo_bytes(logo)
    if not img:
        return
    from pptx.util import Inches
    target_w = Inches(0.5)
    left = prs.slide_width  - target_w - Inches(0.25)
    top = Inches(0.3)
    try:
        slide.shapes.add_picture(io.BytesIO(img), left, top, width=target_w)
    except Exception as e:
        _log.warning("failed to place icon logo: %s", e)


def _place_center_logo(slide, prs, logo, crop: bool = False):
    """Large centered logo for the explicit closing slide.

    If `crop` is True, apply a slight crop to the inserted picture.
    """
    img = _logo_bytes(logo)
    if not img:
        return
    from pptx.util import Inches

    target_w = Inches(4.2)
    left = (prs.slide_width - target_w) / 2
    top = (prs.slide_height - Inches(1.2)) / 2 - Inches(0.3)
    try:
        pic = slide.shapes.add_picture(io.BytesIO(img), left, top, width=target_w)
        if crop:
            try:
                pic.crop_left = 0.05
                pic.crop_right = 0.05
                pic.crop_top = 0.03
                pic.crop_bottom = 0.03
            except Exception:
                pass
    except Exception as e:
        _log.warning("failed to place centered logo: %s", e)


def _add_closing_placeholder(slide, prs, colors, fonts):
    """Add the closing-slide placeholder text box."""
    from pptx.util import Inches, Pt
    from pptx.enum.text import PP_ALIGN, MSO_ANCHOR

    box_w = Inches(6.0)
    box_h = Inches(0.7)
    left = (prs.slide_width - box_w) / 2
    top = prs.slide_height - Inches(1.1)

    box = slide.shapes.add_textbox(left, top, box_w, box_h)
    tf = box.text_frame
    tf.clear()
    tf.word_wrap = True
    tf.margin_left = Inches(0.05)
    tf.margin_right = Inches(0.05)
    tf.margin_top = Inches(0.02)
    tf.margin_bottom = Inches(0.02)
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE

    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    run = p.add_run()
    run.text = "Add closing text here"
    run.font.size = Pt(14)
    run.font.bold = False
    run.font.color.rgb = _hex_to_rgb(colors.get("text", "#1D1D1D"))
    run.font.name = fonts.get("body", "Arial")


# ─── Slide-type helpers ───────────────────────────────────────────────────────

def _is_section_title_slide(source_slide, slide_index):
    """True when the slide has only a title element and nothing else."""
    if slide_index == 0:
        return False
    elements = source_slide.get("elements", [])
    non_title = [
        e for e in elements
        if not (e.get("type") == "text" and e.get("role") == "title")
    ]
    has_title = any(
        e.get("type") == "text" and e.get("role") == "title"
        for e in elements
    )
    return has_title and not non_title


def _dark_slide_colors(colors):
    """White-on-dark palette for title / closing / section slides."""
    return {
        **colors,
        "primary":   "#FFFFFF",
        "text":      "#FFFFFF",
        "secondary": "#F2F2F2",
        "grey":      "#CCCCCC",
    }


# ─── Element renderers ────────────────────────────────────────────────────────

def _render_text_element_preserve(slide, element, colors, fonts):
    from pptx.util import Inches, Pt

    role = element.get("role", "body")
    paragraphs = element.get("paragraphs", [])
    if not paragraphs:
        return

    x, y, w, h = (element.get(k, d) for k, d in
                   [("x", 0), ("y", 0), ("w", 1), ("h", 1)])
    if w <= 0 or h <= 0:
        return

    box = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = box.text_frame
    tf.clear()
    tf.word_wrap = True
    tf.margin_left = Inches(0.05)
    tf.margin_right = Inches(0.05)
    tf.margin_top = Inches(0.03)
    tf.margin_bottom = Inches(0.03)

    title_font = fonts.get("title", "Arial")
    body_font = fonts.get("body",  "Arial")

    base_size = element.get("max_font_size")
    if base_size is None:
        base_size = 30 if role == "title" else (9 if role in ("footer", "caption") else 16)
    font_size = _clamp(base_size, 7, 46)

    for i, para in enumerate(paragraphs):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
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
    from pptx.util import Inches
    img = element.get("image_bytes")
    if not img:
        return
    x, y, w, h = (element.get(k, d) for k, d in
                   [("x", 0), ("y", 0), ("w", 1), ("h", 1)])
    if w <= 0 or h <= 0:
        return
    try:
        slide.shapes.add_picture(
            io.BytesIO(img), Inches(x), Inches(y),
            width=Inches(w), height=Inches(h),
        )
    except Exception:
        pass


def _render_table_element_preserve(slide, element, colors, fonts):
    from pptx.util import Inches, Pt
    rows = element.get("rows", [])
    if not rows:
        return
    row_count = len(rows)
    col_count = max(len(r) for r in rows)
    if row_count <= 0 or col_count <= 0:
        return
    x, y, w, h = (element.get(k, d) for k, d in
                   [("x", 0), ("y", 0), ("w", 1), ("h", 1)])
    if w <= 0 or h <= 0:
        return
    try:
        tbl = slide.shapes.add_table(
            row_count, col_count,
            Inches(x), Inches(y), Inches(w), Inches(h),
        ).table
        body_font = fonts.get("body", "Arial")
        text_color = _hex_to_rgb(colors.get("text", "#1D1D1D"))
        for r_idx, row in enumerate(rows):
            for c_idx in range(col_count):
                cell = tbl.cell(r_idx, c_idx)
                cell.text = row[c_idx] if c_idx < len(row) else ""
                for para in cell.text_frame.paragraphs:
                    para.font.name = body_font
                    para.font.size = Pt(10)
                    para.font.color.rgb = text_color
    except Exception:
        pass


# ─── Main renderer ────────────────────────────────────────────────────────────

def render_slides_preserve_structure(extracted_deck, guidelines):
    """
    Render one output slide per source slide, applying brand chrome
    (accent bar, footer band, logo, background colour) from guidelines.
    """
    from pptx import Presentation
    from pptx.util import Inches
    from pptx.dml.color import RGBColor

    prs = Presentation()
    prs.slide_width = Inches(extracted_deck.get("slide_width",  13.333))
    prs.slide_height = Inches(extracted_deck.get("slide_height",  7.5))

    # ── Normalise guideline keys ──────────────────────────────────────────
    colors, fonts = _normalise_guidelines(guidelines)

    # ── Logo candidates (dicts from extract_guidelines) ───────────────────
    full_logo = guidelines.get("full", None)
    full_logo_dark = guidelines.get("full_darkback", None)
    icon_logo = guidelines.get("icon", None)

    slides = extracted_deck.get("slides", [])

    for i, source_slide in enumerate(slides):
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        slide_number = i + 1

        is_title = (i == 0)
        is_section = _is_section_title_slide(source_slide, i)
        is_dark = is_title or is_section

        # ── Background ────────────────────────────────────────────────────
        bg = slide.background.fill
        bg.solid()
        bg.fore_color.rgb = (
            _hex_to_rgb(colors["primary"]) if is_dark else RGBColor(255, 255, 255)
        )

        slide_colors = _dark_slide_colors(colors) if is_dark else colors

        # ── Brand chrome (content slides only) ───────────────────────────
        if not is_dark:
            _add_accent_bar(slide, prs, colors)
            _add_footer_band(slide, prs, colors, fonts, slide_number)

        # ── Content ───────────────────────────────────────────────────────
        elements = source_slide.get("elements", [])

        for element in elements:
            if element.get("type") == "image":
                _render_image_element_preserve(slide, element)

        for element in elements:
            if element.get("type") == "table":
                _render_table_element_preserve(slide, element, slide_colors, fonts)

        for element in elements:
            if element.get("type") == "text":
                _render_text_element_preserve(slide, element, slide_colors, fonts)

        # ── Logo ──────────────────────────────────────────────────────────
        if is_dark:
            if full_logo_dark:
                _place_titles_logo(slide, prs, full_logo_dark, crop=True)
            else:
                _place_titles_logo(slide, prs, full_logo, crop=False)
        else:
            _place_full_logo(slide, prs, full_logo)
            _place_icon_logo(slide, prs, icon_logo)

    # ── Explicit closing slide ───────────────────────────────────────────
    closing = prs.slides.add_slide(prs.slide_layouts[6])
    closing_bg = closing.background.fill
    closing_bg.solid()
    closing_bg.fore_color.rgb = _hex_to_rgb(colors.get("primary", "#24135F"))

    if full_logo_dark:
        _place_center_logo(closing, prs, full_logo_dark, crop=True)
    else:
        _place_center_logo(closing, prs, full_logo, crop=False)
    _add_closing_placeholder(closing, prs, _dark_slide_colors(colors), fonts)

    buf = io.BytesIO()
    prs.save(buf)
    return buf.getvalue()


# ─── CLI for local testing ────────────────────────────────────────────────────

if __name__ == "__main__":
    from extract_guidelines import extract_guidelines, build_model
    from extract_pptx_content import extract_all_content
    from dotenv import load_dotenv

    load_dotenv()

    parser = argparse.ArgumentParser(description="Reformat PPTX per brand guidelines PDF")
    parser.add_argument("source_pptx",     help="Path to source presentation")
    parser.add_argument("guidelines_pdf",  help="Path to brand guidelines PDF")
    parser.add_argument("output_pptx",     help="Where to write reformatted .pptx")
    args = parser.parse_args()

    source_bytes = Path(args.source_pptx).read_bytes()
    guidelines_bytes = Path(args.guidelines_pdf).read_bytes()

    api_key = os.getenv("WATSONX_API_KEY")
    project_id = os.getenv("WATSONX_PROJECT_ID")
    url = os.getenv("WATSONX_URL")

    text_model = None
    vision_model = None
    try:
        text_model = build_model(api_key, project_id, url)
    except Exception as e:
        print(f"Error occurred while building text model: {e}")
    try:
        vision_model = build_model(
            api_key, project_id, url,
            model_id="meta-llama/llama-3-2-11b-vision-instruct",
        )
    except Exception as e:
        print(f"Error occurred while building vision model: {e}")

    extracted_deck = extract_all_content(source_bytes)
    guidelines = extract_guidelines(guidelines_bytes, text_model, vision_model)

    output_bytes = render_slides_preserve_structure(extracted_deck, guidelines)

    out_path = Path(args.output_pptx)
    out_path.parent.mkdir(parents=True, exist_ok=True)
    out_path.write_bytes(output_bytes)
    print(f"Wrote {out_path} ({len(output_bytes):,} bytes)")