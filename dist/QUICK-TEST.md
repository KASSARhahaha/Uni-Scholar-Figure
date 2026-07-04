# DakeSCI Clone — Distribution Package

This folder produces `DakeSCI-Clone-v1.2.zip` — a one-click installable
PowerPoint add-in for Windows.

## Files
- `DakeSCI.bas` — VBA module implementing all 7 features (623 lines, pure VBA + GDI+ flat API)
- `customUI14.xml` — ribbon definition with 7 buttons
- `install.ps1` — PowerShell installer (import VBA, save as .ppam, inject ribbon, register)
- `uninstall.ps1` — uninstaller
- `README.md` — user-facing instructions (Chinese)

## Build artifact
- `../DakeSCI-Clone-v1.2.zip` (11 KB)
- SHA: see zip itself

## Validation done
- customUI14.xml: well-formed XML ✅
- All 7 ribbon callbacks defined as VBA Public Sub ✅
- ZIP structure: 5 files at root ✅
