# Uni-Scholar Figure 1.3

A cross-platform toolkit for scientific figure creation in PowerPoint — available as both a Python CLI (Linux/Mac/Windows) and a VBA PowerPoint add-in (Windows + Mac).

## Version
`1.3.0` — release `2026-07-18`

## Install (Python CLI)

```bash
cd /home/yangkai/00-make-money/Uni-Scholar-Figure
source .venv/bin/activate
uv pip install -e .
unisfig --help
```

## Commands (one per feature)

| # | Feature | Command |
|---|---------|---------|
| 1 | PNG 空白边裁除 | `unisfig trim-png *.png [--out-dir trimmed]` |
| 2 | 生成表格 | `unisfig gen-table data.csv -d deck.pptx [--title "Results"]` |
| 3 | 层状结构 | `unisfig layer-stack layers.txt -d deck.pptx` |
| 4 | 矩阵逐行偏移 | `unisfig matrix --rows 4 --cols 5 --offset 0.5 --mode alternate` |
| 5 | 表格排版保持图片比例 | `unisfig table-images ./imgs --rows 2 --cols 3` |
| 6 | 图标补充 | `unisfig add-icon warning --slide 0 --left 5 --top 5` |
| 7 | 统一安装/发布日期 | `unisfig version` |

Input formats for `gen-table`: `auto` (default), `csv`, `markdown` (pipe-table), `lines` (one cell per line).

`matrix --mode` accepts `alternate` (odd rows shifted), `progressive` (each row shifts more), or `none`.

## Install (PowerPoint Add-in)

See `dist/README.md` for one-click install instructions (Windows PowerShell + Mac).

## Codex Plugin

The evidence-backed Figure 1 workflow plugin lives in `plugins/unischolar-figure/`.
It supports research tracing, original journal-style workflow design, and FigEdit
export to editable SVG and PowerPoint.

```bash
codex plugin marketplace add KASSARhahaha/Uni-Scholar-Figure --ref main
codex plugin add unischolar-figure@unischolar-figure
```

Start a new Codex task after installation so the `unischolar-figure` skill is available.

## Public Codex submission materials

The repository includes the public materials needed to prepare a skills-only
Codex plugin submission. They are not an approval or a promise of listing;
publication requires a verified publisher to submit through OpenAI's review
portal.

- [Submission overview](docs/public-listing/README.md)
- [Privacy policy](docs/privacy.md)
- [Terms of use](docs/terms.md)
- [Support](docs/support.md)
- [Reviewer test cases](docs/plugin-review-test-cases.md)

The listing uses the publisher identity `KASSARhahaha`, subject to the identity
verified in the OpenAI Platform submission portal.

## Layout

```
src/unisfigure/         # Python CLI package
  __init__.py           # version + RELEASE_DATE
  release.py            # feature 7: single source of truth
  trim_png.py           # feature 1
  gen_table.py          # feature 2
  layer_stack.py        # feature 3
  matrix_grid.py        # feature 4
  table_images.py       # feature 5
  icons.py              # feature 6
  cli.py                # typer top-level app
dist/                   # PowerPoint add-in distribution
  UniScholarFigure.bas  # VBA module (cross-platform, #If Mac Then branches)
  customUI14.xml        # ribbon definition
  install.ps1           # Windows one-click installer
  install-mac.command   # Mac installer
  uninstall.ps1         # Windows uninstaller
tests/
  test_smoke.py         # 23 passing tests
  samples/              # layers.txt + table.csv for manual trial
task_plan.md            # planning-with-files tracker
notes.md                # behavior notes
```

## Test

```bash
pytest tests/ -v   # 23 passed
```

## Privacy and data boundary

Read the [privacy policy](docs/privacy.md) before using online features or
submitting sensitive material. Local layout and PowerPoint operations run on
the user's device. Network access happens only when the user selects an
online feature: **Check for Updates**, **Research Records**, Codex literature
research, or an approved image-generation route. The repository does not
operate a telemetry or analytics service.

## Out of scope for the Python CLI and PowerPoint Add-in
- AI image generation (not part of the core v1.2 feature set)
- Licensing / 兑换码 / 机器码 flow
- Decompiling or shipping any byte of third-party binaries
