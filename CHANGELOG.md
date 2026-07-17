# Changelog

All notable changes to Uni-Scholar Figure are documented here.
The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

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
