#!/usr/bin/env python3
"""Run a supported command from the bundled FigEdit snapshot."""

from __future__ import annotations

import argparse
import subprocess
import sys
from pathlib import Path


COMMANDS = {
    "prepare": "prepare_measurements.py",
    "measure": "measure.py",
    "draft-elements": "draft_elements.py",
    "manifest-edit": "manifest_edit.py",
    "compose": "compose_svg_package.py",
    "validate-manifest": "validate_manifest.py",
    "quality-audit": "quality_audit.py",
    "editability-audit": "audit_editability.py",
    "render-pptx": "render_pptx.py",
}


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("command", choices=sorted(COMMANDS))
    parser.add_argument("arguments", nargs=argparse.REMAINDER)
    args = parser.parse_args()

    research_skill = Path(__file__).resolve().parents[1]
    figedit_scripts = research_skill.parent / "figedit-v2" / "scripts"
    script = figedit_scripts / COMMANDS[args.command]
    if not script.is_file():
        parser.error(f"bundled FigEdit command is missing: {script}")

    completed = subprocess.run([sys.executable, str(script), *args.arguments], check=False)
    return int(completed.returncode)


if __name__ == "__main__":
    raise SystemExit(main())
