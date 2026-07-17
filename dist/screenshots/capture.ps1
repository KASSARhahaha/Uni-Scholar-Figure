# Uni-Scholar Figure — Screenshot Helper for Windows
# Usage: .\capture.ps1
# Prompts for a number 1-11, waits 5 seconds, captures primary screen,
# saves as win-NN-<name>.png in this folder.

#Requires -Version 5

$ErrorActionPreference = 'Stop'
$here = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $here

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

function Take-Shot([int]$num, [string]$name) {
    Write-Host "`nCapturing in 5 seconds... focus the target window now." -ForegroundColor Yellow
    for ($i = 5; $i -ge 1; $i--) {
        Write-Host "`r$i...  " -NoNewline
        Start-Sleep -Milliseconds 900
    }
    Write-Host ""

    # Primary screen bounds
    $bounds = [System.Windows.Forms.Screen]::PrimaryScreen.Bounds
    $bmp = New-Object System.Drawing.Bitmap $bounds.Width, $bounds.Height
    $g = [System.Drawing.Graphics]::FromImage($bmp)
    $g.CopyFromScreen($bounds.Location, [System.Drawing.Point]::Empty, $bounds.Size)
    $g.Dispose()

    # Safe filename
    $safeName = if ($name) { "-" + ($name -replace '[^A-Za-z0-9-]', '-').Trim('-') } else { "" }
    $file = Join-Path $here ("win-{0:D2}{1}.png" -f $num, $safeName)
    $enc = New-Object System.Drawing.Imaging.ImageCodecInfo 0,0
    $enc = [System.Drawing.Imaging.ImageCodecInfo]::GetImageEncoders() | Where-Object { $_.MimeType -eq 'image/png' }
    $params = New-Object System.Drawing.Imaging.EncoderParameters 1
    $params.Param[0] = New-Object System.Drawing.Imaging.EncoderParameter(
        [System.Drawing.Imaging.Encoder]::ColorDepth, [int64]24)
    $bmp.Save($file, $enc, $params)
    $bmp.Dispose()

    Write-Host "[OK] saved: $file" -ForegroundColor Green
    Write-Host "    size: $((Get-Item $file).Length) bytes"
}

Write-Host "===== Uni-Scholar Figure — Screenshot Helper =====" -ForegroundColor Cyan
Write-Host "Recommended: run pngquant on each file after capture."
Write-Host "Type 'q' to quit.`n"

while ($true) {
    $input = Read-Host "Number (1-11) + optional name (e.g. '4 generate-demo')"
    if (-not $input -or $input -eq 'q') { break }

    $parts = $input -split '\s+', 2
    $num = 0
    if (-not [int]::TryParse($parts[0], [ref]$num) -or $num -lt 1 -or $num -gt 20) {
        Write-Host "Invalid number. Try again." -ForegroundColor Red
        continue
    }
    $name = if ($parts.Count -gt 1) { $parts[1] } else { "" }
    Take-Shot $num $name
}

Write-Host "`nDone. Files in: $here" -ForegroundColor Green
