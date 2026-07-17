"""Smoke + edge-case tests covering all 7 features end-to-end on a real .pptx."""
from __future__ import annotations

import csv
import io
import sys
from pathlib import Path

from PIL import Image, ImageDraw
import pytest
from typer.testing import CliRunner

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))

from unisfigure import RELEASE_DATE, __version__  # noqa: E402
from unisfigure.cli import app  # noqa: E402
from unisfigure.gen_table import gen_table_from_text, parse_input  # noqa: E402
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
    assert __version__ == "1.2.3"
    assert RELEASE_DATE == "2026-07-17"


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


# ---------------------------------------------------------------------------
# Edge-case tests (added in v1.2.3)
# ---------------------------------------------------------------------------

def test_trim_png_all_white(tmp_path):
    """All-white PNG should return full-image bbox, not crash."""
    src = tmp_path / "blank.png"
    Image.new("RGB", (100, 100), "white").save(src, "PNG")
    bbox = trim_one(src)
    assert bbox == (0, 0, 100, 100)


def test_trim_png_colored_content(tmp_path):
    """Bright red content (L≈76) must be detected as non-blank."""
    src = tmp_path / "red.png"
    img = Image.new("RGB", (200, 200), "white")
    ImageDraw.Draw(img).rectangle([60, 60, 140, 140], fill="red")
    img.save(src, "PNG")
    bbox = trim_one(src)
    cropped = Image.open(src)
    # Cropped tightly around the red square (within a couple px padding tolerance)
    assert 70 <= cropped.width <= 90
    assert 70 <= cropped.height <= 90


def test_gen_table_ragged_csv(tmp_path, caplog):
    """Ragged CSV should still build and log a warning."""
    deck = tmp_path / "ragged.pptx"
    text = "A,B,C\n1,2,3\n4,5\n6,7,8,9"  # row 2 short, row 3 long
    with caplog.at_level("WARNING", logger="unisfigure.gen_table"):
        n = gen_table_from_text(deck, text, fmt="csv")
    assert n == 4
    assert deck.exists() and deck.stat().st_size > 0
    assert any("ragged" in rec.message for rec in caplog.records)


def test_gen_table_empty_input(tmp_path):
    """Empty input should still produce a 1x1 table, not crash."""
    deck = tmp_path / "empty.pptx"
    n = gen_table_from_text(deck, "", fmt="lines")
    assert n == 0  # no rows parsed
    assert deck.exists()


def test_gen_table_markdown_with_alignment_row(tmp_path):
    """Markdown alignment row (|:---|) must be filtered out."""
    deck = tmp_path / "md.pptx"
    md = "| A | B |\n|---|---|\n| 1 | 2 |"
    n = gen_table_from_text(deck, md, fmt="markdown")
    assert n == 2  # header + 1 data row, alignment row excluded


def test_parse_input_unknown_format():
    with pytest.raises(ValueError, match="unknown format"):
        parse_input("foo", fmt="xml")


def test_layer_stack_unicode(tmp_path):
    """Unicode layer names (Chinese, emoji) should render correctly."""
    deck = tmp_path / "unicode.pptx"
    render_layer_stack(deck, ["输入层", "隐藏层", "输出层 🧠"], title="架构")
    assert deck.exists() and deck.stat().st_size > 0


def test_layer_stack_empty_layers(tmp_path):
    """Empty layer list should not crash (no-op or minimal deck)."""
    deck = tmp_path / "empty.pptx"
    render_layer_stack(deck, [], title="Empty")
    assert deck.exists()


def test_matrix_non_positive_rows(tmp_path):
    """CLI must reject rows <= 0."""
    runner = CliRunner()
    r = runner.invoke(app, ["matrix", "--rows", "0", "--cols", "3"])
    assert r.exit_code != 0


def test_matrix_bad_mode(tmp_path):
    """CLI must reject unknown mode."""
    runner = CliRunner()
    r = runner.invoke(app, ["matrix", "--rows", "2", "--cols", "2", "--mode", "banana"])
    assert r.exit_code != 0


def test_add_icon_bad_color(tmp_path):
    """CLI must reject malformed hex color."""
    runner = CliRunner()
    r = runner.invoke(app, ["add-icon", "warning", "--color", "XYZ"])
    assert r.exit_code != 0


def test_add_icon_unknown_name(tmp_path):
    """Unknown icon name should raise (no silent fallback)."""
    deck = tmp_path / "x.pptx"
    with pytest.raises((KeyError, ValueError)):
        add_icon_to_slide(deck, "nonexistent_icon_xyz")
