"""PDF-styled deck generator.

Sibling of brand_render, but the brand source is a PDF guidelines document
instead of a reference .pptx. A PDF has no PowerPoint master to inherit from,
so this flow EXTRACTS brand attributes (colors, fonts, logos) from the PDF and
PAINTS them onto a freshly-built deck — reusing the team's reformat_flow:

  - extract_guidelines()                (PDF  -> colors / fonts / logos)
  - render_slides_preserve_structure()  (deck + guidelines -> painted .pptx)

The only NEW piece is _build_deck_from_slides(): it turns the agent's simple
prompt-generated slide specs into the structured deck (with synthesized element
geometry) that render_slides_preserve_structure expects. The renderer lays
nothing out itself — it paints each element at the x/y/w/h we give it — so this
adapter is where slide layout actually happens.

Used by the rolando_pdf_style_agent.

Run locally (text/vision models optional; without them extraction falls back
to default colors/fonts and skips logo detection):
    python tools/reformat_flow/pdf_brand_render.py <guidelines.pdf> <spec.json> <out.pptx>
"""
from ibm_watsonx_orchestrate.agent_builder.tools import tool, WXOFile
from ibm_watsonx_orchestrate.agent_builder.connections import ConnectionType
from ibm_watsonx_orchestrate.run import connections

import argparse
import io
import json
from pathlib import Path

# Flat imports — same pattern as reformat_deck_flow.py; resolves because this
# file lives alongside the reused modules in tools/reformat_flow/.
from extract_guidelines import extract_guidelines, build_model
from render_deck import render_slides_preserve_structure


# ─── Default slide geometry (16:9, inches) ─────────────────────────────────
_SLIDE_W = 13.333
_SLIDE_H = 7.5
_MARGIN = 0.7


def _para(text, level=0):
    """One paragraph in the extracted_deck shape. `runs` is unused by the
    renderer (it applies its own brand fonts) but kept for shape parity."""
    return {"text": str(text), "level": int(level), "runs": []}


def _text_element(role, lines, x, y, w, h, font_size):
    """Build one text element in the extracted_deck shape."""
    paragraphs = [_para(t) for t in lines if str(t).strip()]
    return {
        "type": "text",
        "role": role,
        "text": "\n".join(p["text"] for p in paragraphs),
        "paragraphs": paragraphs,
        "max_font_size": font_size,
        "placeholder_type": "",
        "x": x, "y": y, "w": w, "h": h,
    }


def _normalize_lines(value):
    """Coerce a bullets/body field into a list of non-empty strings."""
    if value is None:
        return []
    if isinstance(value, str):
        return [value]
    out = []
    for item in value:
        if isinstance(item, list):
            out.extend(str(x) for x in item)
        elif str(item).strip():
            out.append(str(item))
    return out


def _build_deck_from_slides(slides, slide_w=_SLIDE_W, slide_h=_SLIDE_H):
    """Convert prompt-generated slide specs into the structured deck that
    render_slides_preserve_structure expects, synthesizing element geometry
    per role. A PDF has no layouts, so we lay out by role:

      title       -> big title + optional subtitle (renderer paints dark bg
                     + places the logo, because it's slide index 0)
      section     -> title-only slide (renderer treats title-only non-first
                     slides as dark section dividers)
      content     -> title + single body column
      two_column  -> title + two body columns

    The renderer auto-appends its own closing slide, so we do NOT emit one."""
    content_w = slide_w - 2 * _MARGIN
    out_slides = []

    for spec in (slides or []):
        # Defensive: LLMs sometimes pass each slide as a JSON-encoded string.
        if isinstance(spec, str):
            try:
                spec = json.loads(spec)
            except (ValueError, TypeError):
                continue
        if not isinstance(spec, dict):
            continue

        role = (spec.get("role") or "content").lower()
        title = spec.get("title", "")
        elements = []

        if role in ("title", "cover"):
            elements.append(_text_element(
                "title", [title], _MARGIN, 2.7, content_w, 1.6, 40))
            sub = spec.get("subtitle")
            if sub:
                elements.append(_text_element(
                    "body", [sub], _MARGIN, 4.4, content_w, 1.0, 20))

        elif role in ("section", "divider"):
            elements.append(_text_element(
                "title", [title], _MARGIN, 3.1, content_w, 1.4, 34))

        elif role in ("two_column", "comparison"):
            elements.append(_text_element(
                "title", [title], _MARGIN, 0.5, content_w, 1.0, 28))
            col_w = (content_w - 0.4) / 2
            left = _normalize_lines(spec.get("left") or spec.get("leftBullets"))
            right = _normalize_lines(spec.get("right") or spec.get("rightBullets"))
            elements.append(_text_element(
                "body", left, _MARGIN, 1.8, col_w, 4.8, 16))
            elements.append(_text_element(
                "body", right, _MARGIN + col_w + 0.4, 1.8, col_w, 4.8, 16))

        else:  # content / bullets
            elements.append(_text_element(
                "title", [title], _MARGIN, 0.5, content_w, 1.0, 28))
            bullets = _normalize_lines(spec.get("bullets") or spec.get("body"))
            elements.append(_text_element(
                "body", bullets, _MARGIN, 1.8, content_w, 4.8, 16))

        out_slides.append({"elements": elements})

    return {
        "slide_width": slide_w,
        "slide_height": slide_h,
        "slide_count": len(out_slides),
        "slides": out_slides,
    }


def _build_models():
    """Build watsonx text + vision models from the watsonx_model_creds
    connection. Returns (text_model, vision_model); either may be None if
    creds are missing — extract_guidelines degrades to default colors/fonts
    and skips logo detection."""
    try:
        creds = connections.key_value("watsonx_model_creds")
        api_key = creds.get("api_key")
        project_id = creds.get("project_id")
        url = creds.get("watsonx_url")
    except Exception:
        return None, None
    text_model = vision_model = None
    try:
        text_model = build_model(api_key, project_id, url)
    except Exception:
        pass
    try:
        vision_model = build_model(
            api_key, project_id, url,
            model_id="meta-llama/llama-3-2-11b-vision-instruct")
    except Exception:
        pass
    return text_model, vision_model


def render_pdf_styled_deck(pdf_bytes, slides, text_model=None, vision_model=None):
    """Extract brand attributes from the PDF, build a deck from the slide
    specs, and paint it. Models are optional for local testing; in Orchestrate
    they come from the watsonx_model_creds connection via _build_models()."""
    if text_model is None and vision_model is None:
        text_model, vision_model = _build_models()
    guidelines = extract_guidelines(pdf_bytes, text_model, vision_model)
    deck = _build_deck_from_slides(slides)
    return render_slides_preserve_structure(deck, guidelines)


# ─── Orchestrate tool wrapper ─────────────────────────────────────────────

@tool(
    name="pdf_brand_render",
    description=(
        "Generates a brand-styled PowerPoint from a brand-guidelines PDF and a "
        "list of slide specs. Extracts colors, fonts, and logos from the PDF "
        "and paints them onto a newly built deck (a PDF has no PowerPoint "
        "master to inherit, so styling is extracted, not inherited). "
        "Inputs: pdf_file (the brand-guidelines .pdf the user uploaded) and "
        "slides (a list of slide dicts). Each slide has a 'role' field — one of "
        "'title', 'section', 'content', 'two_column' — plus content fields: "
        "'title', 'subtitle' (title slides), 'bullets' (content), and "
        "'left'/'right' (two_column). The first slide is rendered as the dark "
        "title slide and a closing slide is appended automatically, so do not "
        "add your own closing slide. Returns the rendered .pptx as a "
        "downloadable file."
    ),
    expected_credentials=[
        {"app_id": "watsonx_model_creds", "type": ConnectionType.KEY_VALUE}
    ],
)
def pdf_brand_render(pdf_file: WXOFile, slides: list) -> bytes:
    pdf_bytes = WXOFile.get_content(pdf_file)
    return render_pdf_styled_deck(pdf_bytes, slides)


# ─── CLI for local testing ────────────────────────────────────────────────

if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("guidelines_pdf", help="Path to the brand guidelines .pdf")
    parser.add_argument("spec_json", help="JSON file with a 'slides' array")
    parser.add_argument("output_pptx", help="Where to write the rendered .pptx")
    args = parser.parse_args()

    pdf_bytes = Path(args.guidelines_pdf).read_bytes()
    spec = json.loads(Path(args.spec_json).read_text(encoding="utf-8"))
    slides = spec["slides"] if isinstance(spec, dict) and "slides" in spec else spec

    # Local testing: try to build models from env vars; fall back to None
    # (default styling) if unavailable.
    text_model = vision_model = None
    try:
        import os
        from dotenv import load_dotenv
        load_dotenv()
        text_model = build_model(
            os.getenv("WATSONX_API_KEY"),
            os.getenv("WATSONX_PROJECT_ID"),
            os.getenv("WATSONX_URL"),
        )
        vision_model = build_model(
            os.getenv("WATSONX_API_KEY"),
            os.getenv("WATSONX_PROJECT_ID"),
            os.getenv("WATSONX_URL"),
            model_id="meta-llama/llama-3-2-11b-vision-instruct",
        )
    except Exception as e:
        print(f"(no watsonx models — using default styling) {e}")

    out_bytes = render_pdf_styled_deck(pdf_bytes, slides, text_model, vision_model)
    out_path = Path(args.output_pptx)
    out_path.parent.mkdir(parents=True, exist_ok=True)
    out_path.write_bytes(out_bytes)
    print(f"Wrote {out_path} ({len(out_bytes):,} bytes)")
