# Capture Guide — Uni-Scholar Figure Screenshots

Run this on a **Windows 10/11 + PowerPoint 365/2021** VM (or real machine).
Estimated time: **30 minutes** for 9 screenshots.

After completion, push the PNGs to this folder — they will be auto-linked
from `../README.md`.

---

## Environment prep (one time)

1. Fresh Windows 10/11 VM (VirtualBox / Parallels / Hyper-V all fine).
2. Install **Microsoft 365** or **Office 2021** (30-day trial is enough).
3. Enable VBA trust access **before** running `install.ps1`:
   `PowerPoint → File → Options → Trust Center → Trust Center Settings
    → Macro Settings → ☑ Trust access to the VBA project object model`
4. Copy the contents of `Uni-Scholar-Figure-v1.2.zip` to `C:\UniScholarFigure\`.
5. Open PowerShell as current user (NOT admin), `cd C:\UniScholarFigure`.

---

## Screenshot list (9 required)

Each screenshot: **1280×720 minimum, PNG, keep under 300 KB each**
(run `pngquant <file>.png` if too big).

| # | File name | What to show | How to capture |
|---|-----------|--------------|----------------|
| 1 | `win-01-install-ps1.png` | `install.ps1` output with three green `[OK]` lines | Run `.\install.ps1` in PowerShell. When "Installation complete!" appears, Win+Shift+S to snip the whole window. |
| 2 | `win-02-ribbon-tab.png` | PowerPoint ribbon with "Uni-Scholar Figure" tab visible + the 9 buttons | Open PowerPoint → new blank deck → click the "Uni-Scholar Figure" tab. Press PrtScn. Crop to the ribbon area only. |
| 3 | `win-03-about-dialog.png` | MsgBox from clicking "About" | Click the **About** button. Screenshot the MsgBox showing version `1.2.4` + release `2026-07-17`. |
| 4 | `win-04-generate-demo.png` | 3-slide demo deck after Generate Demo | Click **Generate Demo**. After the 3 slides are built, switch to Slide Sorter view (View → Slide Sorter). Screenshot showing all 3 slides. |
| 5 | `win-05-gen-table.png` | Slide with a CSV-pasted table | Click **Generate Table** → paste the prefilled `Method,Precision,Recall,F1...` → OK. Screenshot the slide with the new table. |
| 6 | `win-06-layer-stack.png` | Slide with 5 colored stacked rectangles + arrows | Click **Layer Stack** → type `Presentation|Application|Business Logic|Data Access|Storage` → OK. Screenshot. |
| 7 | `win-07-matrix-offset.png` | Slide with 3×4 alternating-offset matrix | Click **Matrix Offset** → accept defaults (3 rows, 4 cols, alternate). Screenshot. |
| 8 | `win-08-trim-png.png` | Before/after of PNG Trim | (a) Create a PNG with whitespace using Snipping Tool or any image. (b) Click **Trim PNG** → select the file. (c) MsgBox "Trimmed 1 file(s)." Screenshot MsgBox OR a before/after side-by-side. |
| 9 | `win-09-check-updates.png` | MsgBox from clicking Check Updates | Click **Check Updates**. Should show "A new version is available" (if v1.2.4 isn't the latest) or "You're up to date." Screenshot the MsgBox. |

**Optional extras** (nice-to-have):

- `win-10-uninstall.png` — output of `uninstall.ps1` (ribbon gone after restart)
- `win-11-vba-editor.png` — Alt+F11 showing the imported `UniScholarFigure` module

---

## Naming convention

Strict: `win-NN-something.png` (lowercase, hyphens, zero-padded number).

The `win-` prefix is required because Mac screenshots will use `mac-pptm-NN-*.png`
and `mac-ppam-NN-*.png` later.

## Helper script (optional)

Run `.\capture.ps1` to open an interactive helper that:

1. Prompts for the next number (1-9)
2. Waits 5 seconds (time to focus the window)
3. Captures the primary screen to `win-NN-<timestamp>.png`
4. Asks for a semantic suffix to rename to the final name

You can also use Win+Shift+S (Snip & Sketch) manually — whichever is faster.

## After capturing

1. Verify all 9 files exist in this folder.
2. Run `pngquant *.png --ext .png --force` to shrink them.
3. Commit + push to GitHub.
4. The `../README.md` will automatically display them (image links are
   already wired with relative paths).

## Troubleshooting

- **"PowerPoint blocks VBA project access"** → enable Trust Access (env prep step 3).
- **Ribbon tab doesn't appear** → fully quit POWERPNT.EXE (Task Manager), relaunch.
- **Trim PNG does nothing on Mac/Linux-transferred PNG** → make sure the PNG
  has actual white margins; some PNGs from Linux have RGBA transparency that
  the Windows GDI+ reader reports as opaque black. Re-save with white bg first.
- **Check Updates fails** → network/firewall. Either allow
  `api.github.com:443` or accept that the screenshot shows a network error.
