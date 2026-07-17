# UniScholarFigure 1.2 Clean-Room Clone — Task Plan  ✅ DONE

## Goal
Reproduce the 7 listed UniScholarFigure 1.2 behaviors as a Linux-runnable Python CLI.
NOT a binary crack, NOT a UI clone. Behavior-only clean-room reimplementation.

## Source of truth (release notes from user)
1. PNG 空白边自动裁除 — ✅ `unisfig trim-png`
2. 生成表格 — ✅ `unisfig gen-table`
3. 层状结构 — ✅ `unisfig layer-stack`
4. 矩阵逐行偏移 — ✅ `unisfig matrix`
5. 表格排版保持图片比例 — ✅ `unisfig table-images`
6. 图标补充 — ✅ `unisfig add-icon` (55 bundled icons)
7. 统一安装日期显示 — ✅ `unisfig version` + `RELEASE_DATE`

## Platform: Option B
- Python 3.11 + python-pptx 1.0.2 + Pillow + typer 0.26
- Runs on Linux, no PowerPoint required
- venv at `/home/yangkai/00-make-money/Uni-Scholar-Figure/.venv`
- Package installed editable as `unisfigure`

## Status
- [x] P0 Scaffold
- [x] P1 PNG trim — fixed WHITE_TOL bug (was 8, needed 245 to detect colored content)
- [x] P2 Generate Table (CSV/Markdown/lines)
- [x] P3 Layer Stack
- [x] P4 Matrix offset (alternate/progressive/none)
- [x] P5 Table images (aspect-ratio preserved, letterbox/pillarbox)
- [x] P6 Icons (10 hand-traced SVG path catalog)
- [x] P7 CLI glue + README
- [x] P8 End-to-end verified: `demo.pptx` builds with 3 slides, 31 KB

## Verification
- `pytest tests/ -v` → 8 passed
- CLI end-to-end: gen-table → layer-stack → matrix → add-icon all write into the same `demo.pptx`

## Risk / boundary
- No decompilation of the original `.exe` was performed
- No proprietary assets (icons / images / manifest) copied
- All shipped icons are hand-traced public-domain-style glyphs
- Behavior guesses (esp. exact ribbon UI) documented in `notes.md`

## Future (out of scope for v1)
- PowerPoint VSTO/.NET add-in shell for the same behaviors (needs Windows)
- AI image generation (not in 1.2 release notes; site branding only)
- License server / 兑换码 / 机器码 flow (separate web concern)
