# DakeSCI Clone — Uninstaller
#Requires -Version 5

$ErrorActionPreference = 'SilentlyContinue'

Write-Host ""
Write-Host "DakeSCI Clone uninstaller" -ForegroundColor Cyan
Write-Host "=========================="
Write-Host ""

$addInPath = Join-Path $env:APPDATA 'Microsoft\AddIns\DakeSCI.ppam'

# Try to unload from a running PowerPoint
try {
    $pp = New-Object -ComObject PowerPoint.Application
    foreach ($a in $pp.AddIns) {
        if ($a.Name -like 'DakeSCI*') {
            $a.Registered = $false
            Write-Host "[OK] Unloaded from running PowerPoint" -ForegroundColor Green
        }
    }
    $pp.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($pp) | Out-Null
} catch {
    Write-Host "(PowerPoint not running — skipping live unload)" -ForegroundColor DarkGray
}

# Remove add-in file
if (Test-Path $addInPath) {
    Remove-Item $addInPath -Force
    Write-Host "[OK] Removed: $addInPath" -ForegroundColor Green
} else {
    Write-Host "Add-in file not found (already removed)" -ForegroundColor DarkGray
}

# Clean registry
$officeVer = '16.0'
$regBase = "HKCU:\Software\Microsoft\Office\$officeVer\PowerPoint\AddIns\DakeSCI"
if (Test-Path $regBase) {
    Remove-Item $regBase -Recurse -Force
    Write-Host "[OK] Cleaned registry" -ForegroundColor Green
}

Write-Host ""
Write-Host "Uninstall complete. Restart PowerPoint to clear the ribbon." -ForegroundColor Cyan
Read-Host "Press Enter to close"
