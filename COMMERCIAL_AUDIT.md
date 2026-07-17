# Uni-Scholar Figure — Commercial Readiness Audit

**Scope**: Full audit against commercial distribution bar (paid product at ¥99/device class).
**Date**: 2026-07-17
**Auditor**: Claude Scholar
**Verdict**: **Not yet commercial-ready.** 4 P0 blockers, 6 P1 issues, 9 P2 improvements.

---

## Conclusion (Conclusion-first)

The project is a credible **clean-room clone** of the 7 listed behaviors, but as a **paid commercial product** it currently fails on four fronts that a paying customer would hit within the first hour:

1. **Legal**: No LICENSE, no clean-room provenance document, no trademark clearance.
2. **Mac UX**: Ribbon tab won't appear for Mac-only users without a Windows-built `.ppam`.
3. **Trust & Safety**: Unsigned `.ps1` requires the user to lower PowerShell policy; `.ppam` triggers the yellow "Macro Security" bar — both feel sketchy to paying users.
4. **Quality assurance**: 8 happy-path tests, zero edge-case / cross-platform / VBA tests; one brand residue bug (`/tmp/dake_trim_`) slipped through.

Fix the 4 P0 items below, then ship. Anything else can ship in v1.3.

---

## Risk-Priority Order

### P0 — Must fix before charging money

| # | Issue | File:Line | Fix |
|---|-------|-----------|-----|
| P0-1 | **Brand residue**: temp file prefix still uses old name `/tmp/dake_trim_…` | `dist/UniScholarFigure.bas:248` | Rename to `/tmp/unisfig_trim_` |
| P0-2 | **No LICENSE file** at repo root | repo root | Add `LICENSE` (Apache-2.0 or MIT) + `NOTICE.md` documenting clean-room reverse-engineering from public release notes (no decompilation, no proprietary assets) |
| P0-3 | **Version mismatch**: `pyproject.toml` says `1.2.0`, release tag is `v1.2.1`, VBA `UNISFIG_VERSION="1.2.0"` | `pyproject.toml`, `src/unisfigure/__init__.py`, `dist/UniScholarFigure.bas:15-16` | Bump all three to `1.2.1` consistently, cut tag `v1.2.1` → `v1.2.2` |
| P0-4 | **Mac ribbon tab does not appear** unless user builds `.ppam` on Windows first | `dist/install-mac.command:69-76` | Either (a) ship a prebuilt `.ppam` inside the ZIP (you can build once on a Windows VM / Parallels and commit the binary), or (b) write a Mac-side `.ppam` builder using AppleScript `do shell script` + a bundled python-pptx-generated `.pptm`→`.ppam` conversion (limited; Office for Mac does not expose `SaveAs ppSaveAsAddIn` via COM). Recommend (a). |

### P1 — Ship blockers for "commercial quality"

| # | Issue | Evidence | Fix |
|---|-------|----------|-----|
| P1-1 | **Test coverage is happy-path only** | `tests/test_smoke.py` — 8 tests, all use well-formed inputs | Add: empty file, malformed CSV (uneven columns), non-PNG argument, missing slide index, color hex without `#`, very large image (OOM?), unicode layer names, 0×0 matrix |
| P1-2 | **Zero VBA tests** | no VBA test harness | At minimum, document a manual Win + Mac smoke test checklist in `dist/QUICK-TEST.md`; ideally add `tests/vba_smoke.md` with screenshots from a real install |
| P1-3 | **No CHANGELOG.md** | repo root | Add `CHANGELOG.md` with v1.2.0 + v1.2.1 entries (Keep a Changelog format) |
| P1-4 | **No SHA256 checksum** for release asset | `gh release view v1.2.1` — 1 asset, no checksums | Run `sha256sum Uni-Scholar-Figure-v1.2.zip > SHA256SUMS`, attach to release, link in README |
| P1-5 | **No input validation** in CLI/VBA hot paths | `cli.py:82-96` — `matrix --rows -1` accepted; `gen-table` accepts ragged CSV silently | Add `typer.Option(..., callback=validator)` for positive ints, hex color regex, ragged-row warning |
| P1-6 | **PowerShell execution policy not documented** | `dist/install.ps1` opens but `ExecutionPolicy -ExecutionPolicy Bypass` usage not in README | Add a one-line "Right-click → Run with PowerShell" instruction + a signed `$SIG` future plan |

### P2 — Polish / roadmap

1. **No auto-update**: user must re-download ZIP. Consider a "Check for updates" ribbon button that fetches `https://api.github.com/repos/.../releases/latest`.
2. **No error reporting channel**: no issues template, no feedback email in ribbon.
3. **Generate Demo button doesn't draw real icons** (`UniScholarFigure.bas:97-104`) — uses `AddShape(9, …)` (rectangle) with icon name as text, not the icon path from `icons.py`. Misleads users evaluating feature 6.
4. **`demo.pptx` is in working tree but gitignored** — confusing for new contributors.
5. **No type-check config**: add `[tool.mypy]` and `[tool.ruff]` sections to `pyproject.toml`, run `ruff check .` clean.
6. **No PR/Issue templates** (`.github/ISSUE_TEMPLATE/`).
7. **No CONTRIBUTING.md** — boundary between Python CLI and VBA module is non-obvious.
8. **Mac installer hand-waves the ribbon**: even if `.ppam` is shipped, Mac PowerPoint's "Tools → Add-ins" path varies by Office version (2016/2019/2021/365); document screenshots.
9. **Privacy statement missing**: state clearly "no telemetry, no network calls" in README — it's a selling point.

---

## What's already good (evidence-backed)

| Area | Evidence |
|------|----------|
| Clean-room provenance | `dist/UniScholarFigure.bas:10-11` documents "No proprietary assets are used" |
| Test signal | `pytest tests/ -v` → **8 passed in 0.18s** (2026-07-17) |
| VBA conditional compile | `#If Mac Then` × 3 balanced with `#If Not Mac Then` × 3 (6/6) |
| Rebrand hygiene | Only **1** leftover `dake` token across 1807 LOC (`dist/UniScholarFigure.bas:248`) |
| Code style | Files all 65–148 LOC, well under the 400-line ceiling |
| Type hints | 100% on public functions in `trim_png.py`, `gen_table.py`, `icons.py`, `table_images.py` |
| Immutability | No mutable default args; constants module-level UPPER_SNAKE |
| Privacy | Zero `requests` / `urllib` / socket calls in source — fully offline |
| Ribbon XML | Well-formed, validated, 8 buttons with screentips |
| Cross-platform strategy | Dual-path VBA (`#If Mac Then` shell-out to Python+Pillow; Windows native GDI+) is the right call given no Office Web Add-in investment |

---

## Feature parity vs. DakeSCI 1.2 release notes

| Listed behavior | Python CLI | VBA | Parity |
|-----------------|:---:|:---:|:---:|
| PNG blank-margin trim | ✅ | ✅ Win+Mac | Full |
| Generate Table | ✅ | ✅ | Full |
| Layer Stack | ✅ | ✅ | Full |
| Matrix row-offset | ✅ | ✅ | Full |
| Table images w/ aspect ratio | ✅ | ✅ | Full |
| Icon supplement | ✅ 55 icons | ⚠ demo draws rectangles, not real icons | **Partial** — see P2-3 |
| Unified release-date display | ✅ | ✅ | Full |

**Net**: 6.5 / 7. The Icon feature in VBA is the only soft spot.

---

## Recommended release sequence

1. **Today (P0)**: Fix brand residue, add LICENSE/NOTICE, sync version to 1.2.2, ship Mac prebuilt `.ppam`.
2. **This week (P1)**: Edge-case tests, CHANGELOG, SHA256SUMS, input validation, README polish on execution policy.
3. **v1.3 roadmap (P2)**: Real icons in Demo, auto-update button, mypy/ruff clean, screenshots.

---

## Files inspected

- `README.md`, `pyproject.toml`, `.gitignore`
- `src/unisfigure/{__init__,cli,trim_png,gen_table,layer_stack,matrix_grid,table_images,icons,release}.py`
- `dist/{UniScholarFigure.bas,customUI14.xml,install.ps1,install-mac.command,uninstall.ps1,QUICK-TEST.md}`
- `tests/test_smoke.py`
- `Uni-Scholar-Figure-v1.2.zip` (release artifact, 6 files, 16 KB)
- `demo.pptx`, `render_preview.py`

## Verification performed

- `pytest tests/ -v` → 8 passed
- `grep -i dake` across all tracked files → 1 hit (the P0-1 bug)
- `#If Mac` / `#If Not Mac` balance check → 3/3
- `unzip -l` on release asset → 6 files at root, no nesting issues
- `gh release view v1.2.1` → published, asset attached
