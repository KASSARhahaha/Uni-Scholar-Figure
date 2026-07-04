"""Smoke tests covering all 7 features end-to-end on a real .pptx."""
from __future__ import annotations

import csv
import io
import sys
from pathlib import Path

from PIL import Image, ImageDraw
import pytest

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))

from unisfigure import RELEASE_DATE, __version__  # noqa: E402
from unisfigure.gen_table import gen_table_from_text  # noqa: E402
from unisfigure.icons import add_icon_to_slide, list_icons  # noqa: E402
from unisfigure.layer_stack import render_layer_stack  # noqa: E402
from unisfigure.matrix_grid import render_matrix  # noqa: E402
from unisfigure.table_images import fill_table_with_images  # noqa: E402
from unisfigure.trim_png import trim_one  # noqa: E402

SAMPLES = Path(__file__).resolve().parent / "samples"
SAMPLES.mkdir(exist_ok=True)


def _make_padded_png(path: Path, content_size: int = 80, canvas: int = 200) -> Path:
    """White canvas with a small black square in the middle — lots of blank margin."""
    img = Image.new("RGB", (canvas, canvas), "white")
    d = ImageDraw.Draw(img)
    margin = (canvas - content_size) // 2
    d.rectangle([margin, margin, margin + content_size, margin + content_size], fill="black")
    img.save(path, "PNG")
    return path


def test_release_info():
    assert __version__ == "1.2.0"
    assert RELEASE_DATE == "2026-07-03"


def test_trim_png(tmp_path):
    src = tmp_path / "padded.png"
    _make_padded_png(src)
    bbox = trim_one(src)
    # bbox is (left, top, right, bottom) — should be tightly around the 80px square
    assert abs(bbox[2] - bbox[0] - 80) < 5
    assert abs(bbox[3] - bbox[1] - 80) < 5


def test_gen_table(tmp_path):
    deck = tmp_path / "t.pptx"
    md = "| A | B |\n|---|---|\n| 1 | 2 |\n| 3 | 4 |"
    n = gen_table_from_text(deck, md, title="Test Table")
    assert n == 3  # header + 2 data rows (alignment row filtered)
    assert deck.exists() and deck.stat().st_size > 0


def test_layer_stack(tmp_path):
    deck = tmp_path / "ls.pptx"
    render_layer_stack(deck, ["UI", "Service", "Data", "Storage"], title="Stack")
    assert deck.exists() and deck.stat().st_size > 0


def test_matrix_offset(tmp_path):
    deck = tmp_path / "m.pptx"
    render_matrix(deck, rows=3, cols=4, row_offset_in=0.5, offset_mode="alternate")
    assert deck.exists() and deck.stat().st_size > 0


def test_table_images(tmp_path):
    deck = tmp_path / "ti.pptx"
    grid = []
    for r in range(2):
        row = []
        for c in range(3):
            p = tmp_path / f"img_{r}_{c}.png"
            # Vary aspect ratios to prove the fit logic
            Image.new("RGB", (100 + c * 30, 100 + r * 40), "red").save(p, "PNG")
            row.append(p)
        grid.append(row)
    fill_table_with_images(deck, grid)
    assert deck.exists() and deck.stat().st_size > 0


def test_icons():
    assert "check" in list_icons()
    assert len(list_icons()) >= 8


def test_add_icon(tmp_path):
    deck = tmp_path / "icon.pptx"
    add_icon_to_slide(deck, "warning", slide_index=0)
    assert deck.exists() and deck.stat().st_size > 0
