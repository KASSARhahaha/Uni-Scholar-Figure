# Uni-Scholar Figure — Distribution Package

This folder produces `Uni-Scholar-Figure-v1.2.zip` — a one-click installable
PowerPoint add-in for Windows + Mac.

## Files
- `UniScholarFigure.bas` — VBA module implementing all 7 features (831 lines,
  cross-platform `#If Mac Then` branches).
- `customUI14.xml` — ribbon definition with 8 buttons (7 features + Generate Demo).
- `install.ps1` — Windows PowerShell installer (import VBA, save as .ppam,
  inject ribbon via ZipFile, register for auto-load).
- `install-mac.command` — Mac installer (copies `.bas` + optional `.ppam` to
  `~/Library/Group Containers/UBF8T346G9.Office/UserContent/Add-Ins/`).
- `uninstall.ps1` — Windows uninstaller.
- `README.md` — user-facing instructions (Chinese).

## Build artifact
- `../Uni-Scholar-Figure-v1.2.zip` (6 files, ~16 KB)
- SHA256: see `../SHA256SUMS`

## Automated checks (run before every release)

- `customUI14.xml`: well-formed XML
- All 8 ribbon callbacks exist as `Public Sub` in `UniScholarFigure.bas`
- ZIP structure: 6 files at root, no nesting
- `#If Mac Then` / `#If Not Mac Then` blocks balanced (3/3)
- `pytest tests/ -v` → all passing (20 tests as of v1.2.3)
- `grep -rni "dake\|DakeSCI\|dakekyht" dist/ src/ tests/` → 0 hits
- Version triple check: `pyproject.toml` = `__init__.py` = `UNISFIG_VERSION` in `.bas`

## Manual smoke test — Windows

Run on a clean Windows 10/11 + PowerPoint 365 VM before tagging a release.

| # | Step | Expected |
|---|------|----------|
| W1 | Right-click `install.ps1` → Run with PowerShell | Three green `[OK]` lines, then "Installation complete!" |
| W2 | Open PowerPoint | "Uni-Scholar Figure" tab visible in ribbon |
| W3 | Click "Generate Demo" button | New presentation with 3 slides (title+table, layer-stack, matrix) appears |
| W4 | Click "About" button | MsgBox showing version `1.2.x` and release date `YYYY-MM-DD` |
| W5 | Select a PNG with whitespace → click "Trim PNG" | File overwritten, margins removed; MsgBox "Trimmed 1 file(s)." |
| W6 | Click "Generate Table" → paste CSV | Table inserted on current slide |
| W7 | Click "Layer Stack" → type 3 layer names | 3 stacked rectangles drawn with arrows |
| W8 | Click "Matrix Offset" → 3×4 alternate | Matrix drawn with alternating row offset |
| W9 | Close PowerPoint, run `uninstall.ps1` | Ribbon tab gone on next launch |

## Manual smoke test — Mac

Run on macOS 13+ with PowerPoint for Mac.

### If testing the `.pptm` path (Mac-only, no `.ppam`)

| # | Step | Expected |
|---|------|----------|
| M1 | `bash install-mac.command` | Script detects Python3 + Pillow, copies `.bas` to Add-Ins folder |
| M2 | Open PowerPoint → New Blank → Save As `.pptm` | — |
| M3 | Option+F11 → File → Import File → `UniScholarFigure.bas` | Module appears in project explorer |
| M4 | Close VBA editor, save `.pptm` | — |
| M5 | "Uni-Scholar Figure" tab appears in ribbon | Tab visible while `.pptm` is open |
| M6 | Click "Trim PNG" on a real PNG | If Python3+Pillow installed: trimmed; otherwise "Setup Required" MsgBox |

### If testing the `.ppam` path (pre-built `.ppam` shipped)

| # | Step | Expected |
|---|------|----------|
| MP1 | Put `UniScholarFigure.ppam` next to `install-mac.command` | — |
| MP2 | `bash install-mac.command` | Script prints "Pre-built .ppam detected" branch, copies to Add-Ins |
| MP3 | Open PowerPoint → Tools → PowerPoint Add-ins → + Add → pick `.ppam` | Add-in registered |
| MP4 | Restart PowerPoint (Cmd+Q + relaunch) | "Uni-Scholar Figure" tab permanently in ribbon |

## Versioning checklist before tagging

1. Bump `pyproject.toml`, `src/unisfigure/__init__.py`, `dist/UniScholarFigure.bas`
   (`UNISFIG_VERSION` + `UNISFIG_RELEASE`).
2. Bump `dist/install.ps1` + `dist/install-mac.command` banner strings.
3. Update both `README.md` + `dist/README.md` version line.
4. Update `tests/test_smoke.py::test_release_info` assertions.
5. Add `CHANGELOG.md` entry under `[Unreleased]` → `[x.y.z]`.
6. Run `pytest tests/ -v` → all pass.
7. Rebuild ZIP, regenerate `SHA256SUMS`.
8. Commit, tag `vx.y.z`, push, create GitHub release with both assets.
