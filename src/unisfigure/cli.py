"""Top-level CLI: unisfig <command> [opts]."""
from __future__ import annotations

import sys
from pathlib import Path

import typer

from . import RELEASE_DATE, __version__
from .gen_table import gen_table_from_text
from .icons import add_icon_to_slide, list_icons
from .layer_stack import render_layer_stack
from .matrix_grid import render_matrix
from .table_images import fill_table_with_images
from .trim_png import trim_many, trim_one

app = typer.Typer(
    name="unisfig",
    help=f"Uni-Scholar Figure (v{__version__}, released {RELEASE_DATE})",
    no_args_is_help=True,
    rich_markup_mode="rich",
)


@app.command()
def version() -> None:
    """Show version + unified release date (feature 7)."""
    typer.echo(f"unisfig v{__version__}  (release: {RELEASE_DATE})")


@app.command("trim-png")
def trim_png_cmd(
    paths: list[Path] = typer.Argument(..., help="PNG file(s) or directory"),
    out_dir: Path | None = typer.Option(None, "--out-dir", "-o", help="Output dir (default: in place)"),
    tol: int = typer.Option(245, "--tol", help="L threshold below which a pixel counts as content (0-255)"),
) -> None:
    """Feature 1: trim PNG blank margins."""
    files: list[Path] = []
    for p in paths:
        if p.is_dir():
            files.extend(sorted(p.glob("*.png")))
        elif p.suffix.lower() == ".png":
            files.append(p)
        else:
            typer.echo(f"skip non-png: {p}", err=True)
    if not files:
        typer.echo("no PNGs found", err=True)
        raise typer.Exit(1)
    results = trim_many(files, out_dir=out_dir, in_place=out_dir is None, tol=tol)
    for f, bbox in results:
        typer.echo(f"trimmed {f.name}: bbox={bbox}")


@app.command("gen-table")
def gen_table_cmd(
    input_file: Path = typer.Argument(..., help="CSV / Markdown / plain text file"),
    deck: Path = typer.Option(Path("out.pptx"), "--deck", "-d"),
    fmt: str = typer.Option("auto", "--fmt", help="auto|csv|markdown|lines"),
    title: str | None = typer.Option(None, "--title"),
) -> None:
    """Feature 2: generate a PPT table from a text file."""
    text = input_file.read_text(encoding="utf-8")
    n = gen_table_from_text(deck, text, fmt=fmt, title=title)
    typer.echo(f"added {n} rows -> {deck}")


@app.command("layer-stack")
def layer_stack_cmd(
    layers_file: Path = typer.Argument(..., help="One layer name per line"),
    deck: Path = typer.Option(Path("out.pptx"), "--deck", "-d"),
    title: str | None = typer.Option(None, "--title"),
    gap: float = typer.Option(0.25, "--gap"),
    box_h: float = typer.Option(0.8, "--box-h"),
) -> None:
    """Feature 3: render a layered-architecture figure."""
    layers = [ln.strip() for ln in layers_file.read_text(encoding="utf-8").splitlines() if ln.strip()]
    render_layer_stack(deck, layers, title=title, box_height_in=box_h, gap_in=gap)
    typer.echo(f"rendered {len(layers)} layers -> {deck}")


@app.command("matrix")
def matrix_cmd(
    rows: int = typer.Option(4, "--rows"),
    cols: int = typer.Option(5, "--cols"),
    offset: float = typer.Option(0.5, "--offset", help="Row offset in inches"),
    mode: str = typer.Option("alternate", "--mode", help="alternate|progressive|none"),
    deck: Path = typer.Option(Path("out.pptx"), "--deck", "-d"),
    title: str | None = typer.Option(None, "--title"),
    label: bool = typer.Option(False, "--label", help="Number cells R,C"),
) -> None:
    """Feature 4: render a matrix with per-row offset."""
    label_fn = (lambda r, c: f"{r},{c}") if label else None
    render_matrix(
        deck, rows=rows, cols=cols, row_offset_in=offset, offset_mode=mode,
        title=title, label_fn=label_fn,
    )
    typer.echo(f"rendered {rows}x{cols} matrix -> {deck}")


@app.command("table-images")
def table_images_cmd(
    images_dir: Path = typer.Argument(..., help="Directory of images (sorted by filename)"),
    rows: int = typer.Option(2, "--rows"),
    cols: int = typer.Option(3, "--cols"),
    deck: Path = typer.Option(Path("out.pptx"), "--deck", "-d"),
    title: str | None = typer.Option(None, "--title"),
    cell_size: float = typer.Option(2.0, "--cell-in"),
) -> None:
    """Feature 5: insert images into a table preserving aspect ratio."""
    exts = {".png", ".jpg", ".jpeg", ".gif", ".bmp"}
    files = sorted(p for p in images_dir.iterdir() if p.suffix.lower() in exts)
    if not files:
        typer.echo("no images found", err=True)
        raise typer.Exit(1)
    grid: list[list[Path | None]] = []
    for r in range(rows):
        row = []
        for c in range(cols):
            idx = r * cols + c
            row.append(files[idx] if idx < len(files) else None)
        grid.append(row)
    fill_table_with_images(
        deck, grid, title=title, cell_size_in=(cell_size, cell_size),
    )
    typer.echo(f"placed {min(len(files), rows*cols)} images -> {deck}")


@app.command("add-icon")
def add_icon_cmd(
    name: str = typer.Argument(...),
    deck: Path = typer.Option(Path("out.pptx"), "--deck", "-d"),
    slide: int = typer.Option(0, "--slide"),
    left: float = typer.Option(1.0, "--left"),
    top: float = typer.Option(1.0, "--top"),
    size: float = typer.Option(0.6, "--size"),
    color: str = typer.Option("1F4E79", "--color"),
) -> None:
    """Feature 6: add a named icon to a slide."""
    if name == "list":
        for n in list_icons():
            typer.echo(n)
        return
    add_icon_to_slide(deck, name, slide_index=slide, left_in=left, top_in=top, size_in=size, color_hex=color)
    typer.echo(f"added icon '{name}' -> {deck}")


if __name__ == "__main__":
    app()
