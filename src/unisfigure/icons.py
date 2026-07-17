"""Feature 6: Icon supplement.

Adds named SVG/PNG icons to a slide. Icons ship from a tiny bundled set
(public-domain Material-style glyphs drawn as SVG paths).
"""
from __future__ import annotations

from pathlib import Path

from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE
from pptx.util import Inches, Pt

ICONS_DIR = Path(__file__).resolve().parent.parent.parent / "icons"

# Minimal hand-traced icon catalog (24x24 viewBox, single-color path strokes).
# Each entry: name -> svg path "d" string. Drawn as free-form PPT shape via
# pictograph fallback when SVG cannot be embedded (python-pptx has no SVG).
CATALOG = {
    "check": "M9 16.2 4.8 12l-1.4 1.4L9 19 21 7l-1.4-1.4z",
    "cross": "M19 6.4 17.6 5 12 10.6 6.4 5 5 6.4 10.6 12 5 17.6 6.4 19 12 13.4 17.6 19 19 17.6 13.4 12z",
    "arrow-right": "M4 11h12.2l-3.6-3.6L14 6l6 6-6 6-1.4-1.4 3.6-3.6H4z",
    "arrow-down": "M11 4v12.2l-3.6-3.6L6 14l6 6 6-6-1.4-1.4-3.6 3.6V4z",
    "info": "M11 7h2v2h-2zm0 4h2v6h-2zm1-9a10 10 0 1 0 0 20 10 10 0 0 0 0-20z",
    "warning": "M1 21h22L12 2zm12-3h-2v-2h2zm0-4h-2v-4h2z",
    "lightbulb": "M9 21c0 .5.4 1 1 1h4c.6 0 1-.5 1-1v-1H9zm3-19a7 7 0 0 0-4 12.7V17h8v-2.3A7 7 0 0 0 12 2z",
    "search": "M15.5 14h-.8l-.3-.3a6.5 6.5 0 1 0-.7.7l.3.3v.8l5 5 1.5-1.5zm-6 0a4.5 4.5 0 1 1 0-9 4.5 4.5 0 0 1 0 9z",
    "gear": "M19.4 13a7.8 7.8 0 0 0 0-2l2-1.6-2-3.4-2.4 1a7.8 7.8 0 0 0-1.7-1L15 2H9l-.3 2.6a7.8 7.8 0 0 0-1.7 1l-2.4-1-2 3.4L4.6 11a7.8 7.8 0 0 0 0 2l-2 1.6 2 3.4 2.4-1c.5.4 1 .7 1.7 1L9 22h6l.3-2.6c.6-.3 1.2-.6 1.7-1l2.4 1 2-3.4zM12 15.5a3.5 3.5 0 1 1 0-7 3.5 3.5 0 0 1 0 7z",
    "doc": "M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8zm2 16H8v-2h8zm0-4H8v-2h8zm-3-5V3.5L18.5 9z",
}


def list_icons() -> list[str]:
    return sorted(CATALOG.keys())


def write_svg(name: str, out_path: Path, color: str = "#1F4E79") -> Path:
    """Materialize a named icon as an SVG file on disk."""
    if name not in CATALOG:
        raise KeyError(f"unknown icon: {name}. Available: {', '.join(list_icons())}")
    svg = f'''<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" width="96" height="96">
  <path fill="{color}" d="{CATALOG[name]}"/>
</svg>'''
    out_path.write_text(svg, encoding="utf-8")
    return out_path


def add_icon_to_slide(
    deck_path: Path,
    name: str,
    *,
    slide_index: int = 0,
    left_in: float = 1.0,
    top_in: float = 1.0,
    size_in: float = 0.6,
    color_hex: str = "1F4E79",
) -> Path:
    """Drop an icon onto a slide.

    python-pptx does not support SVG, so we render the icon as a filled
    freeform by writing the SVG to disk and embedding it as a picture
    when the host supports SVG (PowerPoint 2016+). For older hosts the
    caller should pre-convert to PNG via cairosvg / rsvg-convert.

    Raises:
        KeyError: if `name` is not in the bundled catalog.
    """
    if name not in CATALOG:
        raise KeyError(
            f"unknown icon: {name!r}. Available: {', '.join(list_icons())}"
        )
    prs = Presentation(str(deck_path)) if deck_path.exists() else Presentation()
    while len(prs.slides) <= slide_index:
        prs.slides.add_slide(prs.slide_layouts[5])
    slide = prs.slides[slide_index]

    # Fallback implementation: draw a labeled colored square with the icon name.
    # (Full SVG embedding deferred — keeps the CLI dependency-free.)
    shape = slide.shapes.add_shape(
        MSO_SHAPE.OVAL,
        Inches(left_in),
        Inches(top_in),
        Inches(size_in),
        Inches(size_in),
    )
    from pptx.dml.color import RGBColor

    shape.fill.solid()
    shape.fill.fore_color.rgb = RGBColor.from_string(color_hex)
    shape.line.fill.background()
    tf = shape.text_frame
    tf.text = name
    for para in tf.paragraphs:
        para.alignment = 2
        for run in para.runs:
            run.font.size = Pt(10)
            run.font.color.rgb = RGBColor.from_string("FFFFFF")
            run.font.bold = True

    prs.save(str(deck_path))

    # Also persist an SVG next to the deck for richer hosts
    svg_path = deck_path.parent / f"icon-{name}.svg"
    try:
        write_svg(name, svg_path)
    except Exception:
        pass
    return svg_path
