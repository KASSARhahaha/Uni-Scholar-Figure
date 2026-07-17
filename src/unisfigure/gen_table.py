"""Feature 2: Generate Table.

Reads a CSV / Markdown / pipe-table / plain-list input and inserts a native
PowerPoint table on a new slide (or appends to an existing deck).
"""
from __future__ import annotations

import csv
import logging
from io import StringIO
from pathlib import Path
from typing import Iterable

from pptx import Presentation
from pptx.util import Inches, Pt

logger = logging.getLogger(__name__)

ALIGN_LOOKUP = {"left": None, "center": None, "right": None}  # set below


def _parse_csv(text: str) -> list[list[str]]:
    return [row for row in csv.reader(StringIO(text))]


def _parse_markdown_table(text: str) -> list[list[str]]:
    rows: list[list[str]] = []
    for line in text.strip().splitlines():
        line = line.strip()
        if not line.startswith("|"):
            continue
        # skip alignment row like |---|---|
        if set(line.replace("|", "").replace("-", "").replace(":", "").strip()) <= set():
            continue
        cells = [c.strip() for c in line.strip("|").split("|")]
        rows.append(cells)
    return rows


def parse_input(text: str, fmt: str = "auto") -> list[list[str]]:
    fmt = fmt.lower()
    text = text.strip()
    if fmt == "auto":
        if "|" in text and text.count("\n") >= 1:
            fmt = "markdown"
        elif "," in text and "\n" in text:
            fmt = "csv"
        else:
            fmt = "lines"
    if fmt == "csv":
        return _parse_csv(text)
    if fmt in ("markdown", "md"):
        return _parse_markdown_table(text)
    if fmt == "lines":
        return [[line] for line in text.splitlines() if line.strip()]
    raise ValueError(f"unknown format: {fmt}")


def add_table(
    deck_path: Path,
    rows: list[list[str]],
    *,
    slide_index: int | None = None,
    title: str | None = None,
    style: str = "Medium Style 2 - Accent 1",
    font_size: int = 14,
) -> Presentation:
    """Add a table to the deck. If slide_index is None, append a new blank slide."""
    prs = Presentation(str(deck_path)) if deck_path.exists() else Presentation()

    if slide_index is None:
        slide_layout = prs.slide_layouts[5]  # Title Only
        slide = prs.slides.add_slide(slide_layout)
        if title:
            slide.shapes.title.text = title
    else:
        slide = prs.slides[slide_index]

    n_rows = max(len(rows), 1)
    n_cols = max(len(rows[0]) if rows else 1, 1)
    # Detect ragged rows and warn (still normalize — don't fail the build)
    ragged = [(i, len(r)) for i, r in enumerate(rows) if len(r) != n_cols]
    if ragged:
        logger.warning(
            "ragged CSV: %d row(s) have column count != %d (first 3: %s). "
            "Short rows padded with empty cells, long rows truncated.",
            len(ragged), n_cols, ragged[:3],
        )
    # Normalize ragged rows
    norm = [(r + [""] * n_cols)[:n_cols] for r in rows]

    left = Inches(0.8)
    top = Inches(1.6)
    width = Inches(8.4)
    height = Inches(0.4 * n_rows)

    tbl_shape = slide.shapes.add_table(n_rows, n_cols, left, top, width, height)
    tbl = tbl_shape.table
    try:
        tbl_style = tbl._tbl.find
    except Exception:
        pass
    # Apply style name
    from pptx.oxml.ns import qn

    tbl_pr = tbl._tbl.find(qn("a:tblPr"))
    if tbl_pr is not None:
        tbl_pr.set("firstRow", "1")
        tbl_pr.set("bandRow", "1")

    for r, row in enumerate(norm):
        for c, val in enumerate(row):
            cell = tbl.cell(r, c)
            cell.text = val
            for para in cell.text_frame.paragraphs:
                for run in para.runs:
                    run.font.size = Pt(font_size)

    prs.save(str(deck_path))
    return prs


def gen_table_from_text(
    deck_path: Path, text: str, fmt: str = "auto", title: str | None = None
) -> int:
    rows = parse_input(text, fmt=fmt)
    add_table(deck_path, rows, title=title)
    return len(rows)
