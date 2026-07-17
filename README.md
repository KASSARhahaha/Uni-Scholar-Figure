# Uni-Scholar Figure 1.2

A cross-platform toolkit for scientific figure creation in PowerPoint — available as both a Python CLI (Linux/Mac/Windows) and a VBA PowerPoint add-in (Windows + Mac).

## Version
`1.2.4` — release `2026-07-17`

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
  test_smoke.py         # 8 passing smoke tests
  samples/              # layers.txt + table.csv for manual trial
task_plan.md            # planning-with-files tracker
notes.md                # behavior notes
```

## Test

```bash
pytest tests/ -v   # 20 passed
```

## Privacy

**Uni-Scholar Figure makes zero outbound network connections during normal use.**

- No telemetry, no analytics, no crash reporter beacon.
- No license-server or 兑换码 / 机器码 check — the add-in is fully offline.
- No file is uploaded anywhere. Every transformation runs locally on the
  user's machine (PowerPoint + VBA on Win/Mac, Python + Pillow on CLI).
- The only feature that *can* touch the network is the optional
  **Check for Updates** ribbon button (added in v1.2.4), which issues a
  single GET to `api.github.com/repos/KASSARhahaha/Uni-Scholar-Figure/releases/latest`
  when the user explicitly clicks it. No version, identifier, or usage
  data is sent — the response is parsed locally and shown in a MsgBox.

This is a deliberate design choice, and a feature: scientists and
engineers in air-gapped or restricted-network environments can deploy
Uni-Scholar Figure without firewall exceptions or DPA review.

## Out of scope (deliberate)
- AI image generation (not part of v1.2 feature set)
- Licensing / 兑换码 / 机器码 flow
- Decompiling or shipping any byte of third-party binaries
