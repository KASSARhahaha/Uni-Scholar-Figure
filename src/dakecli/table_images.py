"""Feature 5: Table layout preserving image aspect ratio.

Inserts images into a PPT table cell while preserving aspect ratio.
Cells use letterbox/pillarbox layout — image is centered, no stretching.
"""
from __future__ import annotations

from pathlib import Path

from PIL import Image
from pptx import Presentation
from pptx.util import Emu, Inches


def _fit_inside(cell_w_emu: int, cell_h_emu: int, img_w: int, img_h: int) -> tuple[int, int, int, int]:
    """Return (left_offset, top_offset, width, height) in EMU, fitted inside the cell."""
    img_ratio = img_w / img_h
    cell_ratio = cell_w_emu / cell_h_emu
    if img_ratio > cell_ratio:
        # width-bound
        new_w = cell_w_emu
        new_h = int(cell_w_emu / img_ratio)
    else:
        new_h = cell_h_emu
        new_w = int(cell_h_emu * img_ratio)
    left_off = (cell_w_emu - new_w) // 2
    top_off = (cell_h_emu - new_h) // 2
    return left_off, top_off, new_w, new_h


def fill_table_with_images(
    deck_path: Path,
    images: list[list[Path | None]],
    *,
    slide_index: int | None = None,
    title: str | None = None,
    cell_size_in: tuple[float, float] = (2.0, 2.0),
    gap_in: float = 0.05,
) -> Presentation:
    prs = Presentation(str(deck_path)) if deck_path.exists() else Presentation()

    if slide_index is None:
        slide = prs.slides.add_slide(prs.slide_layouts[5])
        if title:
            slide.shapes.title.text = title
    else:
        slide = prs.slides[slide_index]

    rows = len(images)
    cols = max((len(r) for r in images), default=1)
    cell_w = Inches(cell_size_in[0])
    cell_h = Inches(cell_size_in[1])

    left0 = Inches(0.8)
    top0 = Inches(1.6)
    total_w = Inches(cols * (cell_size_in[0] + gap_in))
    total_h = Inches(rows * (cell_size_in[1] + gap_in))

    tbl_shape = slide.shapes.add_table(rows, cols, left0, top0, total_w, total_h)
    tbl = tbl_shape.table

    # Force column widths / row heights so aspect math is stable
    for c in range(cols):
        tbl.columns[c].width = cell_w
    for r in range(rows):
        tbl.rows[r].height = cell_h

    for r, row in enumerate(images):
        for c, img_path in enumerate(row):
            if img_path is None:
                continue
            cell = tbl.cell(r, c)
            # Determine cell pixel-ish dims in EMU (approx)
            cell_w_emu = int(Inches(cell_size_in[0]))
            cell_h_emu = int(Inches(cell_size_in[1]))
            with Image.open(img_path) as im:
                iw, ih = im.size
            off_x, off_y, new_w, new_h = _fit_inside(cell_w_emu, cell_h_emu, iw, ih)
            # cell margin defaults to 0.1in — set to 0 for predictable math
            cell.margin_left = 0
            cell.margin_top = 0
            cell.margin_right = 0
            cell.margin_bottom = 0
            # Add picture on top of the cell at absolute slide coords
            abs_x = left0 + Inches(c * (cell_size_in[0] + gap_in)) + off_x
            abs_y = top0 + Inches(r * (cell_size_in[1] + gap_in)) + off_y
            slide.shapes.add_picture(
                str(img_path), abs_x, abs_y, width=Emu(new_w), height=Emu(new_h)
            )

    prs.save(str(deck_path))
    return prs
