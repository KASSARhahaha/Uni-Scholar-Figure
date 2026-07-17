# Changelog

All notable changes to Uni-Scholar Figure are documented here.
The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [1.3.0] — 2026-07-18

### Added — Uni-Scholar main-site integration
- **📚 Research Records ribbon button** (10th button): pulls the user's saved
  literature from `https://uni-scholar.asia/api/literature/papers` and drops
  them as a table on the current slide. Columns: Title / Authors / Journal /
  Year / DOI / Catalyst. Windows uses `MSXML2.XMLHTTP`, Mac uses `curl`.
- Token management: InputBox prompts on first use, cached at
  `%APPDATA%\UniScholarFigure\token.txt` (Win) /
  `~/Library/Application Support/UniScholarFigure/token.txt` (Mac).
  Hold **Shift + click** the button to reset the token.
- Tiny JSON parser in VBA (`ExtractJsonArray`, `CountJsonObjects`,
  `SplitJsonObjects`, `CleanJsonArray`) — handles the API response without
  pulling in a third-party library.
- **CLI command `unisfig research-records`**: same behavior, plus
  `--year-from` / `--year-to` / `--search` / `--limit` filters, token cached
  at `~/.config/unisfigure/token.txt` (mode 600).
- 3 new tests (offline, mocked records): render with data, empty list,
  CLI without token.

### Changed
- Bumped version to `1.3.0` across 8 locations (was 7 — the ribbon button
  count check is now part of the test surface).
- Ribbon has 10 buttons (was 9): the 7 features + Generate Demo +
  Check Updates + Research Records.
- 23 tests passing (was 20). `ruff check .` still clean.

## [1.2.4] — 2026-07-17

### Added
- **Check for Updates** ribbon button — single GET to
  `api.github.com/repos/.../releases/latest`, version-compared locally,
  zero telemetry. Uses `MSXML2.XMLHTTP` on Windows, `curl` via `MacScript`
  on Mac. New VBA helper `OnCheckUpdate` + `VersionGe` + `ExtractJsonField`.
- **Real icons in Generate Demo** — bundled SVG paths (check / info / warning /
  lightbulb / search / gear / arrow / cross / doc) are now rendered as
  PowerPoint Freeform shapes via a minimal M/L/H/V SVG path parser
  (`DrawIconFreeform`). Unknown names still fall back to a labeled circle.
- **Privacy statement** in `README.md` — makes the zero-network posture
  explicit (only the new Check-for-Updates button opts in).
- `dist/screenshots/` placeholder with naming convention for future captures.
- `[tool.ruff]` and `[tool.mypy]` config blocks in `pyproject.toml`.

### Changed
- Ruff linter clean (0 warnings). Removed unused `tbl_style`/`top0`/`gap`/
  `bbox` variables across `gen_table.py`, `layer_stack.py`,
  `matrix_grid.py`, `tests/test_smoke.py`.
- `.gitignore` now excludes `.ruff_cache/` and `.mypy_cache/`.
- Bump version to `1.2.4` across all 7 locations.
- Ribbon now exposes 9 buttons (was 8): the original 7 features + Generate
  Demo + Check for Updates.

## [1.2.3] — 2026-07-17

### Added
- 12 edge-case tests covering empty / malformed inputs, ragged CSV (with
  warning), unicode layer names, all-white PNGs, colored content detection,
  and unknown icon names (`tests/test_smoke.py` 8 → 20 tests).
- CLI input validation: `matrix --rows`/`--cols` must be positive int,
  `--offset` must be ≥ 0, `--mode` must be `alternate|progressive|none`,
  `add-icon --color` must match `^[0-9A-Fa-f]{6}$`. All violations exit 2
  with a clear `typer.BadParameter` message.
- `gen_table` now logs a WARNING when rows have inconsistent column counts
  (still normalizes — never fails the build).
- `add_icon_to_slide` now raises `KeyError` upfront on unknown icon names
  instead of silently drawing a fallback shape.
- Expanded `dist/QUICK-TEST.md` into a full Windows + Mac manual smoke
  checklist (W1-W9, M1-M6, MP1-MP4) plus pre-tag versioning checklist.

### Changed
- Bump version to `1.2.3` across all 7 locations.

## [1.2.2] — 2026-07-17

### Fixed
- **Brand residue**: renamed `/tmp/dake_trim_*.py` temp file to
  `/tmp/unisfig_trim_*.py` in the Mac VBA branch
  (`dist/UniScholarFigure.bas`). Zero `dake*` tokens remain in the codebase.

### Changed
- Version sync: `pyproject.toml`, `src/unisfigure/__init__.py`,
  `dist/UniScholarFigure.bas` (`UNISFIG_VERSION` / `UNISFIG_RELEASE`),
  `dist/install.ps1`, `dist/install-mac.command`, both READMEs, and the
  release-info test now all read `1.2.2` / `2026-07-17`.

### Added
- `LICENSE` — Apache-2.0.
- `NOTICE.md` — clean-room provenance statement (no decompilation, no
  proprietary assets, trademark disclaimer).
- `CHANGELOG.md` — this file.
- `SHA256SUMS` — integrity checksum for the release ZIP.
- Mac installer now prints step-by-step guidance that branches on whether
  a pre-built `.ppam` is present, including the exact Finder path and the
  Cmd+Q restart hint.

### Security / Privacy
- Reaffirmed: zero network calls in any source file. No telemetry,
  no auto-update beacon, no license-server check.

## [1.2.1] — 2026-07-04

### Added
- Mac compatibility via `#If Mac Then` conditional compilation.
- "Generate Demo" ribbon button (3-slide showcase).
- Icon catalog expanded from 10 to 55 glyphs.

## [1.2.0] — 2026-07-03

### Added
- Initial clean-room release reproducing the 7 DakeSCI 1.2 behaviors.
- Python CLI (`unisfig`) + Windows VBA add-in (`UniScholarFigure.ppam`).
- 7 ribbon buttons, 8 smoke tests.
