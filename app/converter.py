"""
converter.py  –  Dynamic HTML-to-PDF converter
-----------------------------------------------
Handles any HTML format from Microsoft Graph API email bodies:
  - Any CSS unit in widths/padding/heights (px, pt, cm, mm, in, em)
  - Fixed or fluid table layouts
  - Deeply nested tables
  - cid: embedded images
  - External http/https images (kept as-is)
  - <pre> blocks (word-wrapped)
  - Inline styles, colours, backgrounds, borders
  - Outlook-specific noise attributes
"""

from bs4 import BeautifulSoup
from xhtml2pdf import pisa
import io
import re

# ---------------------------------------------------------------------------
# Page constants  (A4 landscape, margins: top/bottom=1.5cm, left/right=2cm)
# ---------------------------------------------------------------------------
PAGE_WIDTH_PT  = 643.0   # 841.89pt - 2*(56.69) = usable width
PAGE_HEIGHT_PT = 482.0   # 595.28pt - 2*(42.52) = usable height
MAX_IMG_WIDTH  = 160     # px cap for embedded images
MAX_IMG_HEIGHT = 160     # px cap for embedded images

# ---------------------------------------------------------------------------
# Base CSS injected into every document
# ---------------------------------------------------------------------------
BASE_CSS = (
    "@page{size:A4 landscape; margin:1.5cm 2cm}"
    "body{font-family:Calibri,Aptos,Helvetica,Arial,sans-serif;"
    "     font-size:12pt;color:#000;margin:0;padding:0;background:white}"
    "p{margin:0}"
    "pre{white-space:pre-wrap;word-wrap:break-word;overflow-wrap:break-word;"
    "    font-size:10pt;font-family:Courier New,Courier,monospace;}"
    "a{color:#0563C1;text-decoration:none}"
    "img{max-width:100%;height:auto}"
    "table{border-collapse:collapse;width:100%;table-layout:fixed}"
    "td,th{word-wrap:break-word;overflow-wrap:break-word;vertical-align:top}"
    "hr{border:none;border-top:1px solid #ccc;margin:8pt 0}"
    "div{word-wrap:break-word}"
)

# Transparent 1x1 px PNG as base64 – placeholder for cid: images
CID_PLACEHOLDER = (
    "data:image/png;base64,"
    "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJ"
    "AAAADUlEQVR42mNkYPhfDwAChwGA60e6kgAAAABJRU5ErkJggg=="
)


# ---------------------------------------------------------------------------
# CSS utility helpers
# ---------------------------------------------------------------------------

def parse_style(style_str):
    """Parse an inline CSS string into an ordered dict."""
    props = {}
    if not style_str:
        return props
    for part in style_str.split(";"):
        part = part.strip()
        if ":" in part:
            k, v = part.split(":", 1)
            props[k.strip().lower()] = v.strip()
    return props


def to_css(d):
    """Serialise a style dict back to an inline CSS string."""
    return "; ".join(f"{k}:{v}" for k, v in d.items() if v is not None)


def to_pt(value_str):
    """
    Convert any CSS length value to points (float).
    Returns None for percentages or unrecognised formats.
    """
    if not value_str:
        return None
    s = str(value_str).strip()
    try:
        if s.endswith("px"):  return float(s[:-2]) * 0.75
        if s.endswith("pt"):  return float(s[:-2])
        if s.endswith("cm"):  return float(s[:-2]) * 28.3465
        if s.endswith("mm"):  return float(s[:-2]) * 2.83465
        if s.endswith("in"):  return float(s[:-2]) * 72.0
        if s.endswith("em"):  return float(s[:-2]) * 12.0
        if s.endswith("rem"): return float(s[:-3]) * 12.0
        if s.endswith("%"):   return None          # handled separately
        return float(s)                            # bare number = pt
    except (ValueError, AttributeError):
        return None


def safe_padding(padding_str):
    """
    Normalise a CSS padding value (any unit mix) to pt values
    capped at 6pt so cells never consume too much space.
    """
    if not padding_str:
        return "0 4pt"
    parts = str(padding_str).strip().split()
    result = []
    for p in parts:
        pt = to_pt(p)
        if pt is None:
            result.append("0")
        else:
            result.append(f"{min(pt, 6.0):.2f}pt")
    return " ".join(result) if result else "0 4pt"


# ---------------------------------------------------------------------------
# Table layout helpers
# ---------------------------------------------------------------------------

def _raw_cell_width(cell):
    """
    Extract the raw width of a <td>/<th> from its style or width attribute.
    Returns the value in points, or None if not determinable.
    """
    style  = parse_style(cell.attrs.get("style", ""))
    w_str  = style.get("width") or str(cell.attrs.get("width", ""))
    pt     = to_pt(w_str)
    if pt is not None:
        return pt
    if w_str.endswith("%"):
        try:
            return float(w_str[:-1]) / 100.0 * PAGE_WIDTH_PT
        except ValueError:
            pass
    return None


def compute_col_widths(table):
    """
    Compute percentage-based column widths that fit within PAGE_WIDTH_PT.
    Strategy:
      1. Read raw widths from the first row.
      2. If total > 95% of page width, scale all down proportionally.
      3. Distribute remaining width evenly to cells without a declared width.
      4. Return list of percentage strings e.g. ["20.0%","80.0%"]
    """
    rows = table.find_all("tr", recursive=False)
    if not rows:
        rows = table.find_all("tr")
    if not rows:
        return []

    # Use first row to determine columns
    cells = rows[0].find_all(["td", "th"])
    if not cells:
        return []

    raw = [_raw_cell_width(c) for c in cells]

    known_total   = sum(w for w in raw if w is not None)
    unknown_count = sum(1 for w in raw if w is None)
    threshold     = PAGE_WIDTH_PT * 0.95

    # Scale down if overflow
    if known_total > threshold:
        if unknown_count == 0:
            scale = threshold / known_total
        else:
            # Leave 15% for unknown cells
            scale = (threshold * 0.85) / known_total
        raw = [w * scale if w is not None else None for w in raw]
        known_total = sum(w for w in raw if w is not None)

    # Fill unknown
    remaining = max(PAGE_WIDTH_PT - known_total, 0)
    fill = remaining / unknown_count if unknown_count > 0 else (PAGE_WIDTH_PT / max(len(raw), 1))
    final = [w if w is not None else fill for w in raw]

    # Clamp tiny columns to at least 3%
    min_pct = 3.0
    total   = sum(final)
    pcts    = [max(w / PAGE_WIDTH_PT * 100, min_pct) for w in final]

    # Re-normalise to 100%
    pct_sum = sum(pcts)
    pcts    = [p / pct_sum * 100 for p in pcts]

    return [f"{p:.1f}%" for p in pcts]


def process_table(table):
    """
    Rewrite a <table> and all its cells with safe, dynamic styles.
    Preserves background colours, borders, and alignment.
    """
    old = parse_style(table.attrs.get("style", ""))
    new = {
        "border-collapse": "collapse",
        "width":           "100%",
        "table-layout":    "fixed",
    }
    for prop in ("background", "background-color", "border", "border-top",
                 "border-bottom", "border-left", "border-right"):
        if prop in old:
            new[prop] = old[prop]

    table.attrs["style"] = to_css(new)
    table.attrs.pop("width",  None)
    table.attrs.pop("height", None)
    table.attrs.pop("cellpadding", None)
    table.attrs.pop("cellspacing", None)

    col_widths = compute_col_widths(table)

    for row in table.find_all("tr"):
        cells = row.find_all(["td", "th"])
        for i, cell in enumerate(cells):
            old_c = parse_style(cell.attrs.get("style", ""))
            new_c = {
                "vertical-align":  old_c.get("vertical-align", "top"),
                "word-wrap":       "break-word",
                "overflow-wrap":   "break-word",
                "padding":         safe_padding(old_c.get("padding", "0 4pt")),
            }
            # Preserve visual properties
            for prop in ("background", "background-color", "color",
                         "border", "border-top", "border-bottom",
                         "border-left", "border-right", "text-align"):
                if prop in old_c:
                    new_c[prop] = old_c[prop]

            # Apply computed width
            if col_widths and i < len(col_widths):
                new_c["width"] = col_widths[i]

            cell.attrs["style"] = to_css(new_c)
            cell.attrs.pop("width",  None)
            cell.attrs.pop("height", None)


# ---------------------------------------------------------------------------
# Image helper
# ---------------------------------------------------------------------------

def process_img(img):
    """
    - Replace cid: references with transparent placeholder
    - Cap dimensions so image doesn't overflow its cell
    - Keep http/https images untouched (xhtml2pdf fetches them)
    """
    src = img.attrs.get("src", "")
    if src.startswith("cid:"):
        img.attrs["src"] = CID_PLACEHOLDER
        img.attrs["alt"] = "[embedded image]"

    # Determine current dimensions in pt
    style  = parse_style(img.attrs.get("style", ""))
    w_pt   = to_pt(style.get("width"))  or to_pt(str(img.attrs.get("width",  "")))
    h_pt   = to_pt(style.get("height")) or to_pt(str(img.attrs.get("height", "")))

    # Convert MAX to pt
    max_w_pt = MAX_IMG_WIDTH  * 0.75
    max_h_pt = MAX_IMG_HEIGHT * 0.75

    if w_pt and w_pt > max_w_pt:
        ratio = max_w_pt / w_pt
        w_pt  = max_w_pt
        if h_pt:
            h_pt = h_pt * ratio

    if h_pt and h_pt > max_h_pt:
        ratio = max_h_pt / h_pt
        h_pt  = max_h_pt
        if w_pt:
            w_pt = w_pt * ratio

    if w_pt or h_pt:
        new_style = {}
        if w_pt: new_style["width"]  = f"{w_pt:.1f}pt"
        if h_pt: new_style["height"] = f"{h_pt:.1f}pt"
        new_style["max-width"] = "100%"
        img.attrs["style"]  = to_css(new_style)
        if w_pt: img.attrs["width"]  = str(int(w_pt / 0.75))
        if h_pt: img.attrs["height"] = str(int(h_pt / 0.75))
    else:
        img.attrs["style"] = "max-width:100%;height:auto"


# ---------------------------------------------------------------------------
# Main cleaning pipeline
# ---------------------------------------------------------------------------

def _clean_html(raw_html):
    """
    Full cleaning pipeline:
      1. Strip broken Outlook <style> blocks
      2. Remove Outlook-specific noise attributes
      3. Process images (cid → placeholder, cap size)
      4. Dynamically rewrite tables for safe rendering
      5. Wrap in a clean HTML skeleton with our BASE_CSS
    """
    soup = BeautifulSoup(raw_html, "html.parser")

    # 1. Remove <style> blocks (we inject our own)
    for tag in soup.find_all("style"):
        tag.decompose()

    # 2. Strip noisy Outlook attributes
    for tag in soup.find_all(True):
        tag.attrs.pop("data-outlook-trace", None)
        tag.attrs.pop("class", None)
        # Remove x_ / OWA-prefixed IDs
        tid = tag.attrs.get("id", "")
        if isinstance(tid, str) and (tid.startswith("x_") or tid.startswith("OWA")):
            del tag.attrs["id"]

    # 3. Process images
    for img in soup.find_all("img"):
        process_img(img)

    # 4. Process tables (outermost to innermost avoids double-scaling)
    seen = set()
    for table in soup.find_all("table"):
        if id(table) not in seen:
            process_table(table)
            seen.add(id(table))

    # 5. Build final document
    body = soup.find("body")
    body_html = str(body) if body else str(soup)

    return (
        '<!DOCTYPE html>\n'
        '<html lang="en">\n'
        '<head>\n'
        '  <meta charset="utf-8">\n'
        f'  <style>{BASE_CSS}</style>\n'
        '</head>\n'
        + body_html
        + '\n</html>'
    )


# ---------------------------------------------------------------------------
# Public entry point
# ---------------------------------------------------------------------------

def html_to_pdf(raw_html: str) -> bytes:
    """
    Convert raw HTML (Microsoft Graph API email body) to PDF bytes.
    Raises Exception on xhtml2pdf errors.
    """
    clean = _clean_html(raw_html)
    buf   = io.BytesIO()
    result = pisa.CreatePDF(clean, dest=buf)
    if result.err:
        raise Exception(f"PDF rendering error (code {result.err})")
    return buf.getvalue()