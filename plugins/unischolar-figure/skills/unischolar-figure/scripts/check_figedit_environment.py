#!/usr/bin/env python3
"""Check the bundled FigEdit snapshot and its required Python imports."""

from __future__ import annotations

import argparse
import importlib.util
from pathlib import Path


DETERMINISTIC_MODULES = {
    "numpy": "numpy",
    "PIL": "Pillow",
    "matplotlib": "matplotlib",
    "fontTools": "fonttools",
    "lxml": "lxml",
    "pptx": "python-pptx",
}

FULL_MODULES = {
    "cv2": "opencv-python",
    "paddleocr": "paddleocr",
    "paddle": "paddlepaddle",
    "scipy": "scipy",
    "latex2mathml": "latex2mathml",
    **DETERMINISTIC_MODULES,
}


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "--mode",
        choices=["deterministic", "full"],
        default="deterministic",
        help="Use full for OCR and arbitrary raster reconstruction",
    )
    parser.add_argument("--require-math", action="store_true")
    args = parser.parse_args()

    skill_dir = Path(__file__).resolve().parents[1]
    figedit_dir = skill_dir.parent / "figedit-v2"
    required = [figedit_dir / "SKILL.md", figedit_dir / "scripts", figedit_dir / "templates"]
    missing_paths = [str(path) for path in required if not path.exists()]
    modules = dict(FULL_MODULES if args.mode == "full" else DETERMINISTIC_MODULES)
    if args.require_math:
        modules["latex2mathml"] = "latex2mathml"
    missing_modules = [package for module, package in modules.items() if importlib.util.find_spec(module) is None]

    print(f"Bundled FigEdit: {figedit_dir}")
    print(f"Mode: {args.mode}; editable math required: {args.require_math}")
    print(f"Snapshot files: {'ok' if not missing_paths else 'missing'}")
    if missing_paths:
        for path in missing_paths:
            print(f"  missing path: {path}")
    print(f"Python imports: {'ok' if not missing_modules else 'missing'}")
    if missing_modules:
        print("  missing packages: " + ", ".join(missing_modules))
        print(f"  requirements file: {figedit_dir / 'requirements.txt'}")
    return 1 if missing_paths or missing_modules else 0


if __name__ == "__main__":
    raise SystemExit(main())
