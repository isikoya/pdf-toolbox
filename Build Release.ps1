#Requires -Version 5.1
<#
    BUILD A RELEASE

    Run this in the toolbox folder when you want to publish a new version.
    It reads the version from PDFToolbox.ps1, zips the files people need,
    and writes version.json so the in-app update check can see the release.

    HOW TO PUBLISH A NEW VERSION
      1. Change $script:Version near the top of PDFToolbox.ps1 (e.g. to 1.1).
      2. Run this script. It makes:  dist\PDF Toolbox v1.1.zip  and  version.json
      3. Commit and push version.json to the repository.
      4. Make a new Release on GitHub, tag it v1.1, and attach the .zip.

    The update check reads version.json from the repository, so step 3 is what
    tells everyone a new version exists. Attaching the zip in step 4 is what
    they actually download.
#>

param(
    [string]$Notes = '',
    [string]$OutputFolder = 'dist'
)

$ErrorActionPreference = 'Stop'
$here = $PSScriptRoot
if (-not $here) { $here = (Get-Location).Path }

function Write-Step { param([string]$Text) Write-Host "-- $Text" -ForegroundColor Cyan }

# ----------------------------------------------------------------------
# Read the version from the engine, so there is only one place to change it
# ----------------------------------------------------------------------
$engine = Join-Path $here 'PDFToolbox.ps1'
if (-not (Test-Path -LiteralPath $engine)) {
    Write-Host "PDFToolbox.ps1 is not in this folder: $here" -ForegroundColor Red
    return
}
$versionLine = Select-String -LiteralPath $engine -Pattern "^\s*\`$script:Version\s*=\s*'([^']+)'" | Select-Object -First 1
if (-not $versionLine) {
    Write-Host 'Could not find $script:Version in PDFToolbox.ps1.' -ForegroundColor Red
    return
}
$version = $versionLine.Matches[0].Groups[1].Value
Write-Step "Version: $version"

# ----------------------------------------------------------------------
# The files a person needs. Anything missing is reported, not silently skipped.
# ----------------------------------------------------------------------
# Patterns, not exact names, so a renamed launcher is still picked up.
$required = @(
    'PDFToolbox.ps1',
    'PDFToolboxApp.ps1',
    'InstallPDFToolbox.ps1',
    'Install PDFToolbox.bat',
    'Launch PDFToolbox.bat',
    'Launch PDF Toolbox Window.bat',
    'Launch PDF Toolbox*flash*.vbs',
    'PDFToolbox.ico'
)
$optional = @('README.md', 'README.txt', 'version.json')

function Resolve-ReleaseFile {
    param([string]$Pattern)
    return @(Get-ChildItem -LiteralPath $here -Filter $Pattern -File -ErrorAction SilentlyContinue)
}

$include = @()
$missing = @()
foreach ($pattern in $required) {
    $found = Resolve-ReleaseFile $pattern
    if ($found.Count -gt 0) { $include += $found[0].FullName } else { $missing += $pattern }
}
foreach ($pattern in $optional) {
    $found = Resolve-ReleaseFile $pattern
    if ($found.Count -gt 0) { $include += $found[0].FullName }
}
if ($missing.Count -gt 0) {
    Write-Host ''
    Write-Host 'These files are missing and will not be in the zip:' -ForegroundColor Yellow
    foreach ($name in $missing) { Write-Host "   $name" -ForegroundColor Yellow }
    Write-Host ''
    $answer = Read-Host 'Carry on anyway? (Y/N) [N]'
    if ($answer -notmatch '^[Yy]') { return }
}

# ----------------------------------------------------------------------
# Build the zip
# ----------------------------------------------------------------------
$outDir = Join-Path $here $OutputFolder
New-Item -ItemType Directory -Force -Path $outDir | Out-Null
$zipPath = Join-Path $outDir "PDF Toolbox v$version.zip"
if (Test-Path -LiteralPath $zipPath) { Remove-Item -LiteralPath $zipPath -Force }

Write-Step "Zipping $($include.Count) file(s)"
Compress-Archive -LiteralPath $include -DestinationPath $zipPath -CompressionLevel Optimal
$zipSize = [Math]::Round((Get-Item -LiteralPath $zipPath).Length / 1KB)
Write-Host "   $zipPath  ($zipSize KB)" -ForegroundColor Green

# ----------------------------------------------------------------------
# Write version.json, which is what the in-app check reads
# ----------------------------------------------------------------------
$downloadLine = Select-String -LiteralPath $engine -Pattern "UpdateDownloadUrl\s*=\s*'([^']+)'" | Select-Object -First 1
$download = $(if ($downloadLine) { $downloadLine.Matches[0].Groups[1].Value } else { '' })
if (-not $Notes) { $Notes = Read-Host "One line describing this release (optional)" }

$manifest = [ordered]@{
    version  = $version
    released = (Get-Date -Format 'yyyy-MM-dd')
    download = $download
    notes    = $Notes
}
$manifestPath = Join-Path $here 'version.json'
$manifest | ConvertTo-Json | Set-Content -LiteralPath $manifestPath -Encoding UTF8
Write-Step 'Wrote version.json'
Get-Content -LiteralPath $manifestPath | ForEach-Object { Write-Host "   $_" -ForegroundColor DarkGray }

if ($download -like '*GITHUBUSER*') {
    Write-Host ''
    Write-Host 'The download link is still the placeholder. Edit these two lines' -ForegroundColor Yellow
    Write-Host 'near the top of PDFToolbox.ps1 before publishing:' -ForegroundColor Yellow
    Write-Host '   $script:UpdateManifestUrl' -ForegroundColor Yellow
    Write-Host '   $script:UpdateDownloadUrl' -ForegroundColor Yellow
}

Write-Host ''
Write-Host 'NEXT STEPS' -ForegroundColor Cyan
Write-Host '  1. Commit and push version.json to the repository.'
Write-Host "  2. Create a Release on GitHub, tag it v$version, and attach:"
Write-Host "     $zipPath"
Write-Host '  3. Anyone on an older version will see it when they press'
Write-Host '     Check for a new version on the About page.'
Write-Host ''
[void](Read-Host 'Press Enter to close')
