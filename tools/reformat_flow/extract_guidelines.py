"""
Extract brand colors, fonts, and logos from a guidelines PDF.
Optionally uses a watsonx vision model to pick the primary logos.

Result keys:
    colors  – {"primary": "#hex", "secondary": "#hex", "others": [...]}
    fonts   – {"primary": "FontName", "fallback": "FontName"}
    full    – img_bytes, ext, kind, page, is_primary, score, reason for selected full logo
    full_darkback – same for dark-background variant
    icon    – same for icon variant
    icon_darkback – same for dark-background icon variant
"""

from __future__ import annotations

import base64
import json
import logging
import math
import re
import random

log = logging.getLogger(__name__)

# ── Constants ──────────────────────────────────────────────────────────────────

# Logo image size limits (pixels)
MIN_IMG_DIM  = 60
MAX_IMG_AREA = 1_500_000

# How far from a logo (PDF points) a text label can be to count as its caption
MAX_LABEL_DIST = 90

# Aspect ratio: >= 1.6 → wide wordmark ("full"), <= 0.6 → tall/square mark ("icon")
WIDE_HIGH = 1.6
WIDE_LOW  = 0.6

# Luma threshold for "dark background" detection (0–1)
DARK_LUMA = 0.45

# Page title/label patterns that flag a candidate as NOT the primary logo
_FORBIDDEN_RE = re.compile(
    r"\b(forbidden|incorrect|wrong|do\s*not|don.t|misuse|avoid|never|"
    r"prohibited|minimum.size|protection.area|clear.space|example)\b", re.I
)

# Patterns for dark-background and light-background logo variants
_DARK_RE  = re.compile(
    r"\b(revers\w+|negative|inverted|knockout|on.dark|dark.back|on.black|"
    r"solid.color|solid.back|white.version|night.mode)\b", re.I
)
_LIGHT_RE = re.compile(r"\b(on.white|on.light|white.back|light.back|positive)\b", re.I)

# Keywords that identify logo pages and font pages
_LOGO_KWS = ("logo", "logotype", "logomark", "wordmark", "brand mark", "graphical charter")
_FONT_KWS = ("typography", "typeface", "font family", "brand font", "type system")

# Regex to pull hex colors from text
_HEX_RE = re.compile(r"#([0-9A-Fa-f]{6})\b")

# A set of common system fonts to fall back to if we can't identify any from the PDF text
SYSTEM_FONTS = {
        "helvetica", "arial", "calibri", "times new roman", "georgia",
        "verdana", "trebuchet ms", "tahoma", "courier new", "garamond",
        "palatino", "bookman", "century gothic", "franklin gothic"
}
    

# ── Colors ─────────────────────────────────────────────────────────────────────

def _extract_colors(text: str) -> dict:
    seen = []
    for m in _HEX_RE.findall(text):
        h = f"#{m.upper()}"
        if h not in seen:
            seen.append(h)
    if not seen:
        return {}
    return {
        "primary":   seen[0],
        "secondary": seen[1] if len(seen) > 1 else None,
        "others":    seen[2:] if len(seen) > 2 else [],
    }


# ── Fonts ──────────────────────────────────────────────────────────────────────

def _extract_fonts(doc: fitz.Document, font_pages: list[int], model=None) -> dict:
    """
    Call a watsonx language model to identify the primary and fallback fonts used in the document.
    If model doesn't find any or model is None, fall back to a predefined font from a list of common fonts.
    """
    pages_to_scan = font_pages or list(range(doc.page_count))

    text = ""
    for idx in pages_to_scan:
        page   = doc[idx]
        text  += (page.get_text() or "").lower()

    prompt = """Extract the MAIN and the SECONDARY brand font name from this text excerpt from a 
                brand guidelines PDF. Reply with just: 'font name, font name' NOTHING ELSE. 
                For each font you can't find, reply with null.
                Text:""" + text[:4000]

    primary, fallback = None, None
    
    if model is not None:
        try:
            response = model.chat(messages=[
                {"role": "user", "content": prompt}
            ])
            result = response["choices"][0]["message"]["content"].strip()
            primary, fallback = (s.strip() for s in result.split(",")[:2]) if "," in result else (result.strip(), None)
        except Exception as e:
            print(f"Error occurred while calling the language model: {e}")

    if primary == "null" or not primary:
        primary = random.choice(list(SYSTEM_FONTS))
    if fallback == "null" or not fallback:
        fallback = random.choice(list(SYSTEM_FONTS - {primary}))

    return {"primary": primary.capitalize(), "fallback": fallback.capitalize()}


# ── Page classification ────────────────────────────────────────────────────────

def _page_title(page: fitz.Page) -> str:
    """Return the text of the topmost text block on the page."""
    import fitz
    best, best_y = "", float("inf")
    for block in page.get_text("dict").get("blocks", []):
        if block.get("type") != 0:
            continue
        y = block["bbox"][1]
        text = " ".join(
            s.get("text", "") for l in block.get("lines", []) for s in l.get("spans", [])
        ).strip()
        if text and y < best_y:
            best, best_y = text, y
    return best


def _classify_pages(doc: fitz.Document) -> tuple[list[int], list[int]]:
    """Return (logo_page_indices, font_page_indices)."""
    import fitz
    logo_pages, font_pages = [], []
    for i in range(doc.page_count):
        page = doc[i]
        text = (page.get_text() or "").lower()
        title = _page_title(page).lower()
        combined = title + " " + text

        if any(kw in combined for kw in _LOGO_KWS):
            logo_pages.append(i)
        if any(kw in combined for kw in _FONT_KWS):
            font_pages.append(i)

    return logo_pages, font_pages


# ── Logo candidate collection ──────────────────────────────────────────────────

def _nearest_label(page: fitz.Page, bbox: fitz.Rect) -> str:
    import fitz
    best, best_d = "", float("inf")
    for b in page.get_text("blocks"):
        if len(b) >= 7 and b[6] == 1:  # skip image blocks
            continue
        text = (b[4] or "").strip().splitlines()[0][:120]
        if not text:
            continue
        tr = fitz.Rect(b[:4])
        if tr.contains(bbox):
            continue
        dx = max(bbox.x0 - tr.x1, tr.x0 - bbox.x1, 0)
        dy = max(bbox.y0 - tr.y1, tr.y0 - bbox.y1, 0)
        d  = math.hypot(dx, dy)
        if d < best_d and d <= MAX_LABEL_DIST:
            best, best_d = text, d
    return best


def _luma_dark(pix: fitz.Pixmap) -> bool:
    """True if the median luma of sampled pixels is below the dark threshold."""
    import fitz
    w, h   = pix.width, pix.height
    step   = max(1, min(w, h) // 12)
    lumas  = []
    for y in range(0, h, step):
        for x in range(0, w, step):
            try:
                r, g, b = pix.pixel(x, y)[:3]
                lumas.append((0.2126*r + 0.7152*g + 0.0722*b) / 255)
            except Exception:
                pass
    if not lumas:
        return False
    lumas.sort()
    return lumas[len(lumas) // 2] < DARK_LUMA


def _collect_logos(doc: fitz.Document, page_indices: list[int]) -> list[dict]:
    import fitz
    candidates = []

    for idx in page_indices:
        page  = doc[idx]
        title = _page_title(page)

        # ── Raster images embedded in the PDF ─────────────────────────────
        for img_info in page.get_images(full=True):
            xref = img_info[0]
            try:
                img   = doc.extract_image(xref)
                w, h  = img.get("width", 0), img.get("height", 0)
            except Exception:
                continue

            if w < MIN_IMG_DIM or h < MIN_IMG_DIM or w * h > MAX_IMG_AREA:
                continue

            rects = page.get_image_rects(xref) or []
            bbox  = rects[0] if rects else None
            label = _nearest_label(page, bbox) if bbox else ""

            candidates.append({
                "img_bytes":  img["image"],
                "ext":        img.get("ext", "png"),
                "aspect":     w / h if h else 0,
                "page":       idx,
                "page_title": title,
                "label":      label,
                "bbox":       bbox,
                "source":     "raster",
            })

        # ── Vector clusters (draw commands grouped by proximity) ───────────
        try:
            drawings = page.get_drawings()
        except Exception:
            drawings = []

        if not drawings or len(drawings) > 2000:
            continue

        rects_vec = [fitz.Rect(d["rect"]) for d in drawings
                     if d.get("rect") and fitz.Rect(d["rect"]).get_area() > 0]

        # Simple union-find clustering
        parent = list(range(len(rects_vec)))
        def find(i):
            while parent[i] != i:
                parent[i] = parent[parent[i]]; i = parent[i]
            return i
        expanded = [fitz.Rect(r.x0-8, r.y0-8, r.x1+8, r.y1+8) for r in rects_vec]
        for i in range(len(rects_vec)):
            for j in range(i+1, len(rects_vec)):
                if expanded[i].intersects(expanded[j]):
                    ri, rj = find(i), find(j)
                    if ri != rj:
                        parent[ri] = rj

        groups: dict[int, list[int]] = {}
        for i in range(len(rects_vec)):
            groups.setdefault(find(i), []).append(i)

        page_area = page.rect.width * page.rect.height
        for members in groups.values():
            if len(members) < 4:
                continue
            bbox = fitz.Rect(rects_vec[members[0]])
            for m in members[1:]:
                bbox |= rects_vec[m]
            if bbox.width < 20 or bbox.height < 20:
                continue
            if bbox.width * bbox.height > page_area * 0.5:
                continue

            try:
                pix       = page.get_pixmap(clip=bbox, dpi=150, alpha=False)
                img_bytes = pix.tobytes("png")
            except Exception:
                continue

            label = _nearest_label(page, bbox)
            candidates.append({
                "img_bytes":  img_bytes,
                "ext":        "png",
                "aspect":     bbox.width / bbox.height if bbox.height else 0,
                "page":       idx,
                "page_title": title,
                "label":      label,
                "bbox":       bbox,
                "source":     "vector",
            })

    return candidates


# ── Logo classification & scoring ──────────────────────────────────────────────

def _kind(candidate: dict, doc: fitz.Document) -> str:
    import fitz
    aspect = candidate["aspect"]
    is_wide = aspect >= WIDE_HIGH or 0 < aspect <= WIDE_LOW

    # Detect dark background from surrounding pixels
    bbox = candidate.get("bbox")
    dark = False
    if bbox is not None:
        try:
            page = doc[candidate["page"]]
            pad  = fitz.Rect(bbox.x0-12, bbox.y0-12, bbox.x1+12, bbox.y1+12)
            pix  = page.get_pixmap(clip=pad, dpi=36, alpha=False)
            dark = _luma_dark(pix)
        except Exception:
            pass

    # Label overrides pixel check
    label = (candidate.get("label") or "").lower()
    if _LIGHT_RE.search(label): dark = False
    if _DARK_RE.search(label):  dark = True

    suffix = "_darkback" if dark else ""
    return ("full" if is_wide else "icon") + suffix


def _heuristic_score(candidate: dict) -> float:
    """Score 0–1: how likely this candidate is the primary brand logo."""
    score = 0.5

    title = candidate.get("page_title") or ""
    label = candidate.get("label") or ""

    if _FORBIDDEN_RE.search(title): score -= 0.5
    if _FORBIDDEN_RE.search(label): score -= 0.3
    if _DARK_RE.search(title):      score -= 0.2   # dark-variant pages are secondary
    if candidate["source"] == "raster": score += 0.2
    if candidate["kind"] == "full":     score += 0.1

    bbox = candidate.get("bbox")
    if bbox and (bbox.width < 30 or bbox.height < 30):
        score -= 0.2   # too small — likely a size-demo thumbnail

    return max(0.0, min(1.0, score))


# ── watsonx model ─────────────────────────────────────────────────────

def build_model(
    api_key: str,
    project_id: str,
    url: str,
    model_id: str = "ibm/granite-4-h-small",  # default → text
):
    from ibm_watsonx_ai import Credentials
    from ibm_watsonx_ai.foundation_models import ModelInference

    return ModelInference(
        model_id=model_id,
        credentials=Credentials(api_key=api_key, url=url),
        project_id=project_id,
        params={"max_new_tokens": 200, "temperature": 0},
    )


_VISION_PROMPT = """\
You are reviewing an image extracted from a brand guidelines PDF.

Page title : {title}
Nearby text: {label}

Is this the PRIMARY brand logo intended for use, or is it something else \
(usage example, violation, size demo, decorative element, secondary variant)?

Reply with JSON only:
{{"is_primary": true/false, "confidence": 0.0-1.0, "reason": "one sentence"}}"""


def _ask_vision(model, img_bytes: bytes, ext: str, title: str, label: str) -> dict:
    mime  = "image/png" if ext == "png" else "image/jpeg"
    b64   = base64.b64encode(img_bytes).decode()
    prompt = _VISION_PROMPT.format(title=title or "unknown", label=label or "none")

    try:
        response = model.chat(messages=[{
            "role": "user",
            "content": [
                {"type": "image_url", "image_url": {"url": f"data:{mime};base64,{b64}"}},
                {"type": "text", "text": prompt},
            ],
        }])
        text = response["choices"][0]["message"]["content"]
        m    = re.search(r"\{.*?\}", text, re.DOTALL)
        return json.loads(m.group()) if m else {}
    except Exception as e:
        log.warning("Vision call failed: %s", e)
        return {}


# ── Main entry point ───────────────────────────────────────────────────────────

def extract_guidelines(pdf_bytes: bytes, text_model, vision_model=None) -> dict:
    """
    Extract colors, fonts, and logos from a brand guidelines PDF.

    Pass a vision_model (from build_vision_model) to use watsonx for
    logo ranking, otherwise heuristics are used alone.
    """
    import fitz
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")

    try:
        full_text  = "\n".join(doc[i].get_text() for i in range(doc.page_count))
        logo_pages, font_pages = _classify_pages(doc)

        colors = _extract_colors(full_text)
        fonts  = _extract_fonts(doc, font_pages, text_model)

        # Collect logo candidates from logo pages; fall back to all pages
        candidates = _collect_logos(doc, logo_pages or list(range(doc.page_count)))

        # Deduplicate by IoU > 0.5 on same page (keep larger)
        by_page: dict[int, list] = {}
        for c in candidates:
            by_page.setdefault(c["page"], []).append(c)

        deduped = []
        for group in by_page.values():
            group.sort(key=lambda c: -(c["bbox"].get_area() if c["bbox"] else 0))
            kept = []
            for c in group:
                cb = c["bbox"]
                overlap = any(
                    cb and k["bbox"] and
                    (fitz.Rect(cb) & fitz.Rect(k["bbox"])).get_area() /
                    max((fitz.Rect(cb) | fitz.Rect(k["bbox"])).get_area(), 1) > 0.5
                    for k in kept
                )
                if not overlap:
                    kept.append(c)
            deduped.extend(kept)

        # Classify each candidate (full/icon, light/dark)
        for c in deduped:
            c["kind"] = _kind(c, doc)

    finally:
        doc.close()

    # Score and rank
    for c in deduped:
        c["heuristic_score"] = _heuristic_score(c)

    if vision_model:
        # Only send candidates with a decent heuristic score to save API calls
        for c in deduped:
            if c["heuristic_score"] >= 0.3:
                result = _ask_vision(
                    vision_model, c["img_bytes"], c["ext"],
                    c["page_title"], c["label"],
                )
                vision_conf  = float(result.get("confidence", 0))
                is_primary   = bool(result.get("is_primary", False))
                c["score"]   = (vision_conf if is_primary else -vision_conf) * 0.6 \
                               + c["heuristic_score"] * 0.4
                c["reason"]  = result.get("reason", "")
                c["is_primary"] = is_primary
            else:
                c["score"]      = c["heuristic_score"] * 0.4
                c["reason"]     = "skipped by heuristic"
                c["is_primary"] = False
    else:
        for c in deduped:
            c["score"]      = c["heuristic_score"]
            c["is_primary"] = c["heuristic_score"] >= 0.5
            c["reason"]     = "heuristic only"

    logos = sorted(deduped, key=lambda c: c["score"], reverse=True)

    full = sorted([c for c in logos if c["kind"] == "full"], key=lambda c: c["score"], reverse=True)
    full_dark = sorted([c for c in logos if c["kind"] == "full_darkback"], key=lambda c: c["score"], reverse=True)
    icon = sorted([c for c in logos if c["kind"] == "icon"], key=lambda c: c["score"], reverse=True)
    icon_dark = sorted([c for c in logos if c["kind"] == "icon_darkback"], key=lambda c: c["score"], reverse=True)

    # Clean up internal fields not useful downstream
    for c in logos:
        c.pop("bbox", None)

    return {"colors": colors, "fonts": fonts, "full": full[0] if full else None, "full_darkback": full_dark[1] if len(full_dark) > 1 else full_dark[0] if full_dark else None,
            "icon": icon[0] if icon else None, "icon_darkback": icon_dark[0] if icon_dark else None}


# ── CLI ────────────────────────────────────────────────────────────────────────

if __name__ == "__main__":
    import argparse, os
    from dotenv import load_dotenv

    load_dotenv()

    parser = argparse.ArgumentParser()
    parser.add_argument("pdf")
    parser.add_argument("--api-key",    default=os.getenv("WATSONX_API_KEY"))
    parser.add_argument("--project-id", default=os.getenv("WATSONX_PROJECT_ID"))
    parser.add_argument("--url",        default=os.getenv("WATSONX_URL"))
    parser.add_argument("--save-logos", action="store_true")
    args = parser.parse_args()

    vision = None
    text_model = None
    if args.api_key and args.project_id:
        try:
            vision = build_model(args.api_key, args.project_id, args.url, model_id="meta-llama/llama-3-2-11b-vision-instruct")
        except Exception as e:
            print(f"Error occurred while building vision model: {e}")
        try:
            text_model = build_model(args.api_key, args.project_id, args.url)
        except Exception as e:
            print(f"Error occurred while building text model: {e}")
        print("Models ready.")
    else:
        print("No watsonx credentials — using heuristics only.")

    with open(args.pdf, "rb") as f:
        result = extract_guidelines(f.read(), text_model=text_model, vision_model=vision)

    print("\nColors:", result["colors"])
    print("Fonts: ", result["fonts"])
    print("chosen logos:", {k: (v["score"], v["reason"]) for k, v in result.items() if k in ("full", "full_darkback", "icon", "icon_darkback") and v})

    if args.save_logos:
        os.makedirs("logos", exist_ok=True)
        selected_logos = [
            value
            for key, value in result.items()
            if key in ("full", "full_darkback", "icon", "icon_darkback") and value
        ]
        for i, lg in enumerate(selected_logos):
            path = f"logos/{i}_{lg['kind']}_score{lg['score']:.2f}.{lg['ext']}"
            with open(path, "wb") as f:
                f.write(lg["img_bytes"])
        print("\nLogos saved to ./logos/")