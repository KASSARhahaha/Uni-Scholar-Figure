# Uni-Scholar Figure — Windows Installer (PowerShell)
# Usage: Right-click this file → Run with PowerShell
# Or from PowerShell:  .\install.ps1

#Requires -Version 5

$ErrorActionPreference = 'Stop'
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $scriptDir

Write-Host ""
Write-Host "Uni-Scholar Figure v1.2.0 installer" -ForegroundColor Cyan
Write-Host "================================"
Write-Host ""

# ----- 1. Validate source files -----
$basisPath  = Join-Path $scriptDir 'UniScholarFigure.bas'
$ribbonPath = Join-Path $scriptDir 'customUI14.xml'
if (-not (Test-Path $basisPath))  { throw "Missing UniScholarFigure.bas next to this installer." }
if (-not (Test-Path $ribbonPath)) { throw "Missing customUI14.xml next to this installer." }

# ----- 2. Locate PowerPoint -----
try {
    $pp = New-Object -ComObject PowerPoint.Application
} catch {
    Write-Host "ERROR: Microsoft PowerPoint is not installed or not registered." -ForegroundColor Red
    Write-Host "Install Microsoft Office first, then re-run this script."
    exit 1
}
$pp.Visible = $true  # PowerPoint COM requires Visible=MsoTrue for editing

# ----- 3. Try importing VBA (needs Trust Access enabled) -----
$pres = $pp.Presentations.Add()

# Trust-access check
$trustOk = $true
try {
    $null = $pres.VBProject.Name  # throws if trust access disabled
} catch {
    $trustOk = $false
}

if (-not $trustOk) {
    Write-Host ""
    Write-Host "PowerPoint blocks VBA project access. Enable it once:" -ForegroundColor Yellow
    Write-Host "  File > Options > Trust Center > Trust Center Settings >" -ForegroundColor Yellow
    Write-Host "  Macro Settings > [x] Trust access to the VBA project object model" -ForegroundColor Yellow
    Write-Host ""
    Write-Host "Then re-run this installer." -ForegroundColor Yellow
    $pres.Close()
    $pp.Quit()
    exit 2
}

# ----- 4. Import VBA module -----
try {
    $pres.VBProject.VBComponents.Import($basisPath) | Out-Null
    Write-Host "[OK] Imported VBA module" -ForegroundColor Green
} catch {
    Write-Host "ERROR importing VBA: $_" -ForegroundColor Red
    $pres.Close(); $pp.Quit(); exit 3
}

# ----- 5. Save as .ppam add-in -----
$addInDir = Join-Path $env:APPDATA 'Microsoft\AddIns'
if (-not (Test-Path $addInDir)) { New-Item -ItemType Directory -Path $addInDir -Force | Out-Null }
$addInPath = Join-Path $addInDir 'UniScholarFigure.ppam'

# ppSaveAsAddIn = 8
$pres.SaveAs($addInPath, 8)
Write-Host "[OK] Saved add-in: $addInPath" -ForegroundColor Green

$pres.Close()
$pp.Quit()
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($pp) | Out-Null

# ----- 6. Inject customUI14.xml into the .ppam (ribbon) -----
Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem

$zip = [System.IO.Compression.ZipFile]::Open($addInPath, 'Update')

# Remove existing customUI entries if re-installing
@('customUI/customUI14.xml', 'customUI/customUI12.xml', '_rels/.rels') | ForEach-Object {
    $existing = $zip.Entries | Where-Object { $_.FullName -eq $_ }
    # We intentionally do NOT remove .rels — we patch it below
}

# Add customUI14.xml
$uiEntry = $zip.GetEntry('customUI/customUI14.xml')
if ($uiEntry) { $uiEntry.Delete() }
$uiEntry = $zip.CreateEntry('customUI/customUI14.xml')
$writer = New-Object System.IO.StreamWriter($uiEntry.Open())
$writer.Write([System.IO.File]::ReadAllText($ribbonPath))
$writer.Close()

# Patch .rels to reference customUI14.xml (if not present)
$relsEntry = $zip.GetEntry('_rels/.rels')
$reader = New-Object System.IO.StreamReader($relsEntry.Open())
$relsXml = $reader.ReadToEnd()
$reader.Close()
if ($relsXml -notmatch 'customUI14\.xml') {
    $newRel = '<Relationship Id="rIdUniScholarFigure" Type="http://schemas.microsoft.com/office/2007/relationships/ui/extensibility" Target="customUI/customUI14.xml"/>'
    $relsXml = $relsXml -replace '</Relationships>', "$newRel</Relationships>"
    $relsWriter = New-Object System.IO.StreamWriter($relsEntry.Open())
    $relsWriter.Write($relsXml)
    $relsWriter.Close()
    Write-Host "[OK] Patched .rels for ribbon" -ForegroundColor Green
}

# Patch [Content_Types].xml to recognize the macro-enabled format (just in case)
$ctEntry = $zip.GetEntry('[Content_Types].xml')

$zip.Dispose()
Write-Host "[OK] Injected ribbon XML" -ForegroundColor Green

# ----- 7. Register the add-in for auto-load -----
$officeVer = '16.0'  # Office 2016/2019/2021/365
$regBase = "HKCU:\Software\Microsoft\Office\$officeVer\PowerPoint\AddIns\UniScholarFigure"
if (-not (Test-Path $regBase)) { New-Item -Path $regBase -Force | Out-Null }
Set-ItemProperty -Path $regBase -Name 'AutoLoad' -Value 1 -Type DWord
Set-ItemProperty -Path $regBase -Name 'Loaded'  -Value 1 -Type DWord
Write-Host "[OK] Registered for auto-load" -ForegroundColor Green

# ----- 8. Done -----
Write-Host ""
Write-Host "Installation complete!" -ForegroundColor Green
Write-Host "Restart PowerPoint (or close all presentations and reopen) —" -ForegroundColor Cyan
Write-Host 'the "Uni-Scholar Figure" ribbon tab will appear at the top.'
Write-Host ""
Write-Host "To uninstall: run uninstall.ps1"
Write-Host ""
Read-Host "Press Enter to close"
