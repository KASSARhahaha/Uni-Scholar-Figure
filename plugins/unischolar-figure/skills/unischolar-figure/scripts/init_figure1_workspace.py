#!/usr/bin/env python3
"""Create a reproducible workspace for a Figure 1 research and FigEdit task."""

from __future__ import annotations

import argparse
from pathlib import Path


FILES = {
    "research/research_brief.md": """# Research brief\n\nStatus: draft\n\n## Research question\n\nUNSET\n\n## Intended audience and journal\n\nUNSET\n\n## Scientific gap\n\nUNSET\n\n## Supported claims\n\nUNSET\n\n## Conflicts and missing evidence\n\nUNSET\n\n## Defensible Figure 1 conclusion\n\nUNSET\n""",
    "research/evidence_ledger.csv": "claim_id,claim_text,evidence_type,source_title,doi_or_url,publication_year,accessed_date,support_level,figure_element,notes\n",
    "design/figure_contract.md": """# Figure contract\n\nStatus: draft\n\nCore conclusion: UNSET\n\nFigure role: UNSET\n\nFigure archetype: UNSET\n\nTarget journal and article type: UNSET\n\nFinal width and height: UNSET\n\nReading direction: UNSET\n\n## Panel map\n\n- a: UNSET\n\n## Evidence hierarchy\n\nHero evidence: UNSET\n\nSupporting evidence: UNSET\n\nControls or validation: UNSET\n\n## Integrity and review\n\nSource data needed: UNSET\n\nStatistics needed: UNSET\n\nImage-integrity notes: UNSET\n\nMain reviewer risk: UNSET\n""",
    "design/style_reference.md": """# Style reference\n\nStatus: draft\n\nReference citation: UNSET\n\nReference URL or DOI: UNSET\n\nTarget journal guide: UNSET\n\n## Abstract design tokens\n\n- Aspect ratio and final size: UNSET\n- Panel hierarchy and reading direction: UNSET\n- Panel-label convention: UNSET\n- Typography: UNSET\n- Palette roles: UNSET\n- Connector and line-weight grammar: UNSET\n- Direct labels and legend strategy: UNSET\n\n## Originality statement\n\nUNSET\n""",
    "delivery/release_notes.md": """# Figure 1 release notes\n\nStatus: draft\n\nCore conclusion: UNSET\n\nTarget journal: UNSET\n\nResearch cutoff date: UNSET\n\nReference figure: UNSET\n\nSource and conversion path: UNSET\n\nEditability summary: UNSET\n\nNative PowerPoint rendering: deferred\n\nUnresolved risks: UNSET\n""",
    "source/.gitkeep": "",
    "figedit/work/.gitkeep": "",
    "figedit/out/.gitkeep": "",
}


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("directory", type=Path, help="Task directory to create")
    parser.add_argument("--force", action="store_true", help="Replace existing scaffold files")
    args = parser.parse_args()

    root = args.directory.expanduser().resolve()
    created = 0
    skipped = 0
    for relative, content in FILES.items():
        path = root / relative
        path.parent.mkdir(parents=True, exist_ok=True)
        if path.exists() and not args.force:
            skipped += 1
            continue
        path.write_text(content, encoding="utf-8")
        created += 1

    print(f"Figure 1 workspace: {root}")
    print(f"Created: {created}; preserved existing: {skipped}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
