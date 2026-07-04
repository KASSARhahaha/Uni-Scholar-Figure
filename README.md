# dake — DakeSCI 1.2 Clean-Room CLI Clone

A Linux-runnable Python re-implementation of the **behaviors** listed in the
DakeSCI 1.2 release notes (2026-07-03). Not a binary copy, not a UI clone, not
a license-server clone — just the seven shape-manipulation features as a
scriptable CLI so you can batch-process `.pptx` files without PowerPoint.

> Built from the public release-note list. The original DakeSCI 1.2 is a
> Windows PowerPoint VSTO add-in sold by `dakekyht.top` for ¥99/device.
> This clone is independent code; no proprietary assets are shipped.

## Version
`1.2.0` — release `2026-07-03` (mirrors the original 1.2 release date, feature 7).

## Install

```bash
cd /home/yangkai/00-make-money/dakesci-clone
source .venv/bin/activate
uv pip install -e .
dake --help
```

## Commands (one per feature)

| # | Feature | Command |
|---|---------|---------|
| 1 | PNG 空白边裁除 | `dake trim-png *.png [--out-dir trimmed]` |
| 2 | 生成表格 | `dake gen-table data.csv -d deck.pptx [--title "Results"]` |
| 3 | 层状结构 | `dake layer-stack layers.txt -d deck.pptx` |
| 4 | 矩阵逐行偏移 | `dake matrix --rows 4 --cols 5 --offset 0.5 --mode alternate` |
| 5 | 表格排版保持图片比例 | `dake table-images ./imgs --rows 2 --cols 3` |
| 6 | 图标补充 | `dake add-icon warning --slide 0 --left 5 --top 5` |
| 7 | 统一安装/发布日期 | `dake version` |

Input formats for `gen-table`: `auto` (default), `csv`, `markdown` (pipe-table),
`lines` (one cell per line).

`matrix --mode` accepts `alternate` (odd rows shifted), `progressive` (each row
shifts more), or `none`.

## Layout

```
src/dakecli/
  __init__.py        # version + RELEASE_DATE
  release.py         # feature 7: single source of truth
  trim_png.py        # feature 1
  gen_table.py       # feature 2
  layer_stack.py     # feature 3
  matrix_grid.py     # feature 4
  table_images.py    # feature 5
  icons.py           # feature 6
  cli.py             # typer top-level app
tests/
  test_smoke.py      # 8 passing smoke tests
  samples/           # layers.txt + table.csv for manual trial
task_plan.md         # planning-with-files tracker
notes.md             # behavior guesses where the spec was ambiguous
```

## Test

```bash
pytest tests/ -v   # 8 passed
```

## Out of scope (deliberate)
- PowerPoint ribbon UI (no PPT on Linux)
- AI image generation (not in 1.2 release notes; site branding only)
- Licensing / 兑换码 / 机器码 flow (separate web concern)
- Decompiling or shipping any byte of the original `.exe`
