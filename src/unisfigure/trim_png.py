"""Feature 1: PNG blank-margin auto-trim.

Detects the non-white bounding box of a PNG and crops the whitespace border.
Batch-mode over a directory is supported.
"""
from __future__ import annotations

from collections.abc import Iterable
from pathlib import Path

from PIL import Image

WHITE_TOL = 245  # L values < this count as content (i.e. anything not near-white)


def _blank_bbox(img: Image.Image, tol: int = WHITE_TOL) -> tuple[int, int, int, int]:
    """Return bbox of non-blank pixels. Falls back to full image if all blank.

    "Blank" = near-white. Anything darker than `tol` counts as content,
    so bright colors (red, yellow, etc.) are correctly treated as content.
    """
    gray = img.convert("L")
    # Content mask: 255 where content (darker than tol), 0 where blank
    content_mask = gray.point(lambda p: 255 if p < tol else 0)
    bbox = content_mask.getbbox()
    if bbox is None:
        return (0, 0, img.width, img.height)
    return bbox  # type: ignore[return-value]


def trim_one(path: Path, out: Path | None = None, tol: int = WHITE_TOL) -> tuple[int, int, int, int]:
    """Crop whitespace from a single PNG. Returns the cropped bbox."""
    img = Image.open(path)
    if img.mode not in ("RGB", "RGBA"):
        img = img.convert("RGBA")
    # Composite alpha over white so transparent pixels count as blank
    if img.mode == "RGBA":
        bg = Image.new("RGBA", img.size, (255, 255, 255, 255))
        composed = Image.alpha_composite(bg, img).convert("RGB")
    else:
        composed = img.convert("RGB")
    bbox = _blank_bbox(composed, tol=tol)
    cropped = img.crop(bbox)
    target = out or path
    cropped.save(target, format="PNG", optimize=True)
    return bbox


def trim_many(
    paths: Iterable[Path],
    *,
    out_dir: Path | None = None,
    in_place: bool = False,
    tol: int = WHITE_TOL,
) -> list[tuple[Path, tuple[int, int, int, int]]]:
    """Trim many PNGs. If neither out_dir nor in_place, files are overwritten."""
    results: list[tuple[Path, tuple[int, int, int, int]]] = []
    for p in paths:
        out = None
        if not in_place and out_dir is not None:
            out_dir.mkdir(parents=True, exist_ok=True)
            out = out_dir / p.name
        bbox = trim_one(p, out=out, tol=tol)
        results.append((p, bbox))
    return results
