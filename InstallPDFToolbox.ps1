#Requires -Version 5.1
<#
    PDF TOOLBOX INSTALLER

    Checks what is already on the machine and installs only what is missing.
    Nothing is reinstalled. Safe to run again at any time.

    Start it by double-clicking "Install PDFToolbox.bat".

    It installs, only if missing:
      Python 3.12        (winget, user scope where allowed)
      Tesseract OCR      (winget)
      Ghostscript        (winget)
      pypdf, cryptography, ocrmypdf  (pip, --user, no admin rights needed)

    It also fixes the two things that usually make a working install look broken:
      1. PATH inside this window is refreshed after each install, so the checks
         below see the new programs without you reopening anything.
      2. The Python user Scripts folder is added to your PATH permanently,
         which is where pip puts ocrmypdf.exe.
#>

$ErrorActionPreference = 'Stop'
try { $Host.UI.RawUI.WindowTitle = 'PDF Toolbox Installer' } catch { }

$script:PyExe = $null
$script:Notes = New-Object System.Collections.Generic.List[string]

# ======================================================================
# HELPERS
# ======================================================================
function Write-Title {
    param([string]$Text)
    Write-Host ''
    Write-Host '==========================================' -ForegroundColor DarkCyan
    Write-Host ('  ' + $Text) -ForegroundColor Cyan
    Write-Host '==========================================' -ForegroundColor DarkCyan
    Write-Host ''
}

function Write-Step {
    param([string]$Text)
    Write-Host ''
    Write-Host "-- $Text" -ForegroundColor Cyan
}

function Write-Check {
    param([string]$Name, [bool]$Ok, [string]$Detail)
    $tag = if ($Ok) { '  FOUND   ' } else { '  MISSING ' }
    Write-Host $tag -NoNewline -ForegroundColor $(if ($Ok) { 'Green' } else { 'Yellow' })
    Write-Host ('{0,-14} {1}' -f $Name, $Detail)
}

function Read-YesNo {
    param([string]$Prompt, [string]$Default = 'Y')
    while ($true) {
        $answer = ([string](Read-Host "$Prompt (Y/N) [$Default]")).Trim()
        if (-not $answer) { $answer = $Default }
        if ($answer -match '^[Yy]') { return $true }
        if ($answer -match '^[Nn]') { return $false }
    }
}

# Rebuilds this window's PATH from the registry, so programs installed a
# moment ago are visible without closing and reopening PowerShell.
function Update-SessionPath {
    $machine = [Environment]::GetEnvironmentVariable('Path', 'Machine')
    $user = [Environment]::GetEnvironmentVariable('Path', 'User')
    $env:Path = (@($machine, $user) | Where-Object { $_ }) -join ';'
    # Common locations that some installers do not add to PATH at all.
    $extra = @('C:\Program Files\Tesseract-OCR')
    $gs = Get-ChildItem 'C:\Program Files\gs' -Recurse -Filter gswin64c.exe -ErrorAction SilentlyContinue | Select-Object -First 1
    if ($gs) { $extra += $gs.DirectoryName }
    foreach ($e in $extra) {
        if ((Test-Path -LiteralPath $e) -and ($env:Path -notlike "*$e*")) { $env:Path += ";$e" }
    }
}

function Test-Tool {
    param([string]$Name)
    $cmd = Get-Command $Name -ErrorAction SilentlyContinue
    if (-not $cmd) { return $null }
    return $cmd.Source
}

# "python" in the Start menu can be a Microsoft Store stub that installs
# nothing and opens the Store instead. This finds a Python that actually runs.
function Find-Python {
    $previous = $ErrorActionPreference
    $ErrorActionPreference = 'Continue'
    try {
        foreach ($candidate in @('py', 'python')) {
            $cmd = Get-Command $candidate -ErrorAction SilentlyContinue
            if (-not $cmd) { continue }
            if ($cmd.Source -like '*\WindowsApps\*') { continue }
            $out = & $candidate --version 2>&1
            if ($LASTEXITCODE -eq 0 -and "$out" -match '\d+\.\d+') { return $candidate }
        }
    } finally { $ErrorActionPreference = $previous }
    return $null
}

function Test-PyModule {
    param([string]$Module)
    if (-not $script:PyExe) { return $false }
    $previous = $ErrorActionPreference
    $ErrorActionPreference = 'Continue'
    try { & $script:PyExe -c "import $Module" *> $null; return ($LASTEXITCODE -eq 0) }
    finally { $ErrorActionPreference = $previous }
}

function Test-Winget {
    return [bool](Get-Command winget -ErrorAction SilentlyContinue)
}

# ======================================================================
# INSTALLERS
# ======================================================================
function Install-WingetPackage {
    param([string]$Name, [string]$Id, [switch]$TryUserScope)
    Write-Step "Installing $Name"
    $common = @('--exact', '--id', $Id, '--accept-source-agreements', '--accept-package-agreements', '--disable-interactivity')
    $previous = $ErrorActionPreference
    $ErrorActionPreference = 'Continue'
    try {
        if ($TryUserScope) {
            Write-Host '   Trying a user install first (no admin rights needed)...' -ForegroundColor DarkGray
            & winget install @common --scope user
            if ($LASTEXITCODE -eq 0) { Update-SessionPath; return $true }
            Write-Host '   User install did not work. Trying the normal install (Windows may ask for admin).' -ForegroundColor DarkGray
        }
        & winget install @common
        $code = $LASTEXITCODE
    } finally { $ErrorActionPreference = $previous }
    Update-SessionPath
    # -1978335189 and 0x8A150061 both mean "already installed"
    if ($code -eq 0 -or $code -eq -1978335189) { return $true }
    Write-Host "   winget could not install $Name (exit code $code)." -ForegroundColor Red
    $script:Notes.Add("$Name did not install through winget. See the manual steps at the end.")
    return $false
}

function Install-PipPackages {
    param([string[]]$Packages)
    if ($Packages.Count -eq 0) { return }
    Write-Step ("Installing Python packages: " + ($Packages -join ', '))
    $previous = $ErrorActionPreference
    $ErrorActionPreference = 'Continue'
    try {
        & $script:PyExe -m pip install --user --upgrade pip
        & $script:PyExe -m pip install --user @Packages
        $code = $LASTEXITCODE
        if ($code -ne 0) {
            Write-Host ''
            Write-Host '   pip could not reach the package index.' -ForegroundColor Red
            Write-Host '   On a company network this is usually the proxy or SSL inspection blocking pypi.org.' -ForegroundColor Yellow
            Write-Host ''
            Write-Host '   A retry can skip the certificate check for pypi.org only. That is weaker than' -ForegroundColor Yellow
            Write-Host '   a normal install, so use it only if your IT team is fine with it.' -ForegroundColor Yellow
            if (Read-YesNo '   Retry that way?' 'N') {
                & $script:PyExe -m pip install --user --trusted-host pypi.org --trusted-host files.pythonhosted.org @Packages
                $code = $LASTEXITCODE
            }
        }
        if ($code -ne 0) { $script:Notes.Add('The Python packages did not install. Your network is most likely blocking pypi.org. Ask IT to allow pypi.org and files.pythonhosted.org.') }
    } finally { $ErrorActionPreference = $previous }
}

# pip --user puts ocrmypdf.exe in the Python user Scripts folder, which is
# usually not on PATH. This is the most common reason a correct install looks broken.
function Add-UserScriptsToPath {
    if (-not $script:PyExe) { return }
    $previous = $ErrorActionPreference
    $ErrorActionPreference = 'Continue'
    try { $base = "$(& $script:PyExe -c 'import site; print(site.USER_BASE)' 2>$null)".Trim() }
    finally { $ErrorActionPreference = $previous }
    if (-not $base) { return }
    $scripts = Join-Path $base 'Scripts'
    if (-not (Test-Path -LiteralPath $scripts)) { return }
    $userPath = [Environment]::GetEnvironmentVariable('Path', 'User')
    if ($userPath -and ($userPath -split ';' | Where-Object { $_.TrimEnd('\') -ieq $scripts.TrimEnd('\') })) {
        Update-SessionPath
        return
    }
    Write-Step 'Adding the Python Scripts folder to your PATH'
    Write-Host "   $scripts" -ForegroundColor DarkGray
    $newPath = if ($userPath) { "$($userPath.TrimEnd(';'));$scripts" } else { $scripts }
    [Environment]::SetEnvironmentVariable('Path', $newPath, 'User')
    Update-SessionPath
}

# Creates a desktop shortcut that runs the toolbox directly, with the
# custom icon. Pointing at powershell.exe avoids a second console window.
function New-ToolboxShortcut {
    param([string]$ScriptName = 'PDFToolbox.ps1', [string]$LinkName = 'PDF Toolbox.lnk', [switch]$Quiet)
    $folder = $PSScriptRoot
    $scriptPath = Join-Path $folder $ScriptName
    if (-not (Test-Path -LiteralPath $scriptPath)) {
        if (-not $Quiet) { Write-Host "  $ScriptName is not in this folder, so no shortcut was made." -ForegroundColor Yellow }
        return
    }
    $desktop = [Environment]::GetFolderPath('Desktop')
    $linkPath = Join-Path $desktop $LinkName
    try {
        $shell = New-Object -ComObject WScript.Shell
        $link = $shell.CreateShortcut($linkPath)
        $link.TargetPath = "$env:WINDIR\System32\WindowsPowerShell\v1.0\powershell.exe"
        $link.Arguments = '-NoProfile -ExecutionPolicy Bypass -STA -WindowStyle Hidden -File "' + $scriptPath + '"'
        # 7 = start minimised, so the console does not flash on screen before
        # the script hides it. The window version hides its own console too.
        $link.WindowStyle = 7
        $link.WorkingDirectory = $folder
        $link.Description = 'PDF Toolbox - OCR, merge, split, convert'
        $icon = Join-Path $folder 'PDFToolbox.ico'
        if (Test-Path -LiteralPath $icon) { $link.IconLocation = "$icon,0" }
        $link.Save()
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($shell)
        Write-Host "  Shortcut created: $linkPath" -ForegroundColor Green
        Write-Host '  If you move this folder, delete the shortcut and run this installer again.' -ForegroundColor DarkGray
    } catch {
        Write-Host "  Could not create the shortcut: $($_.Exception.Message)" -ForegroundColor Yellow
    }
}

function Add-ShortcutIfWanted {
    $desktop = [Environment]::GetFolderPath('Desktop')
    $hasWindowVersion = Test-Path -LiteralPath (Join-Path $PSScriptRoot 'PDFToolboxApp.ps1')
    Write-Host ''
    if (-not $hasWindowVersion) {
        if (Test-Path -LiteralPath (Join-Path $desktop 'PDF Toolbox.lnk')) {
            Write-Host '  Desktop shortcut: already there.' -ForegroundColor Green
            return
        }
        if (Read-YesNo 'Put a PDF Toolbox shortcut on the desktop?' 'Y') { New-ToolboxShortcut }
        return
    }
    Write-Host '  Desktop shortcut. Which version should it open?'
    Write-Host '    1  The window version (recommended)'
    Write-Host '    2  The console version'
    Write-Host '    3  Both'
    Write-Host '    4  No shortcut'
    $choice = ''
    while ($choice -notin @('1', '2', '3', '4')) {
        $choice = ([string](Read-Host '  Choice [1]')).Trim()
        if (-not $choice) { $choice = '1' }
    }
    if ($choice -eq '4') { return }
    if ($choice -eq '1' -or $choice -eq '3') {
        New-ToolboxShortcut -ScriptName 'PDFToolboxApp.ps1' -LinkName 'PDF Toolbox.lnk'
    }
    if ($choice -eq '2') {
        New-ToolboxShortcut -ScriptName 'PDFToolbox.ps1' -LinkName 'PDF Toolbox.lnk'
    }
    if ($choice -eq '3') {
        New-ToolboxShortcut -ScriptName 'PDFToolbox.ps1' -LinkName 'PDF Toolbox (console).lnk'
    }
}

function Show-ManualSteps {
    Write-Host ''
    Write-Host 'Manual install links (use these if winget is blocked on this machine):' -ForegroundColor Yellow
    Write-Host '  Python       https://www.python.org/downloads/   (tick "Add Python to PATH")'
    Write-Host '  Tesseract    https://github.com/UB-Mannheim/tesseract/wiki'
    Write-Host '  Ghostscript  https://ghostscript.com/releases/gsdnld.html   (64-bit)'
    Write-Host ''
    Write-Host 'After installing by hand, run this installer again. It will skip what is already there.' -ForegroundColor Yellow
}

# ======================================================================
# MAIN
# ======================================================================
Write-Title 'PDF TOOLBOX INSTALLER'
Write-Host 'Checking what is already on this machine. Nothing is installed twice.'
Write-Host ''

Update-SessionPath
$script:PyExe = Find-Python
$hasWinget = Test-Winget

$tesseract = Test-Tool 'tesseract'
$ghost = Test-Tool 'gswin64c'
$pypdf = Test-PyModule 'pypdf'
$crypto = Test-PyModule 'cryptography'
$ocr = Test-PyModule 'ocrmypdf'
$plumber = Test-PyModule 'pdfplumber'
$openpyxl = Test-PyModule 'openpyxl'
$pillow = Test-PyModule 'PIL'

Write-Check 'Python' ([bool]$script:PyExe) $(if ($script:PyExe) { (& $script:PyExe --version 2>&1) } else { 'will install Python 3.12' })
Write-Check 'Tesseract' ([bool]$tesseract) $(if ($tesseract) { $tesseract } else { 'will install' })
Write-Check 'Ghostscript' ([bool]$ghost) $(if ($ghost) { $ghost } else { 'will install' })
Write-Check 'pypdf' $pypdf $(if ($pypdf) { 'merge / split / extract / count' } else { 'will install' })
Write-Check 'cryptography' $crypto $(if ($crypto) { 'password-protected PDFs' } else { 'will install' })
Write-Check 'OCRmyPDF' $ocr $(if ($ocr) { 'OCR' } else { 'will install' })
Write-Check 'pdfplumber' $plumber $(if ($plumber) { 'PDF tables to Excel' } else { 'will install' })
Write-Check 'openpyxl' $openpyxl $(if ($openpyxl) { 'writes .xlsx files' } else { 'will install' })
Write-Check 'Pillow' $pillow $(if ($pillow) { 'images to PDF' } else { 'will install' })

$missingApps = @()
if (-not $script:PyExe) { $missingApps += 'Python' }
if (-not $tesseract) { $missingApps += 'Tesseract' }
if (-not $ghost) { $missingApps += 'Ghostscript' }
$missingPkgs = @()
if (-not $pypdf) { $missingPkgs += 'pypdf' }
if (-not $crypto) { $missingPkgs += 'cryptography' }
if (-not $ocr) { $missingPkgs += 'ocrmypdf' }
if (-not $plumber) { $missingPkgs += 'pdfplumber' }
if (-not $openpyxl) { $missingPkgs += 'openpyxl' }
if (-not $pillow) { $missingPkgs += 'pillow' }

if ($missingApps.Count -eq 0 -and $missingPkgs.Count -eq 0) {
    Add-UserScriptsToPath
    Write-Host ''
    Write-Host 'Everything is already installed. Nothing to do.' -ForegroundColor Green
    Add-ShortcutIfWanted
    Write-Host ''
    Write-Host 'Start the toolbox from the desktop shortcut, or "Launch PDFToolbox.bat".'
    Write-Host ''
    [void](Read-Host 'Press Enter to close')
    return
}

Write-Host ''
Write-Host ('To install: ' + (($missingApps + $missingPkgs) -join ', ')) -ForegroundColor Yellow
Write-Host 'This can take 5 to 10 minutes. You can keep using the machine.' -ForegroundColor DarkGray
Write-Host ''
if (-not (Read-YesNo 'Install these now?' 'Y')) { return }

if ($missingApps.Count -gt 0 -and -not $hasWinget) {
    Write-Host ''
    Write-Host 'winget (Windows Package Manager) is not available on this machine,' -ForegroundColor Red
    Write-Host 'so Python, Tesseract and Ghostscript have to be installed by hand.' -ForegroundColor Red
    Show-ManualSteps
    [void](Read-Host 'Press Enter to close')
    return
}

if (-not $script:PyExe) {
    [void](Install-WingetPackage -Name 'Python 3.12' -Id 'Python.Python.3.12' -TryUserScope)
    $script:PyExe = Find-Python
    if (-not $script:PyExe) {
        Write-Host ''
        Write-Host 'Python still cannot be found after installing. Close this window, reopen the' -ForegroundColor Red
        Write-Host 'installer, and it will pick up the new install.' -ForegroundColor Red
        [void](Read-Host 'Press Enter to close')
        return
    }
}
if (-not $tesseract) { [void](Install-WingetPackage -Name 'Tesseract OCR' -Id 'UB-Mannheim.TesseractOCR') }
if (-not $ghost) { [void](Install-WingetPackage -Name 'Ghostscript' -Id 'Artifex.GhostScript') }

Install-PipPackages -Packages $missingPkgs
Add-UserScriptsToPath

# ----------------------------------------------------------------------
# FINAL CHECK
# ----------------------------------------------------------------------
Write-Title 'RESULT'
Update-SessionPath
$script:PyExe = Find-Python
$final = [ordered]@{
    'Python'       = [bool]$script:PyExe
    'Tesseract'    = [bool](Test-Tool 'tesseract')
    'Ghostscript'  = [bool](Test-Tool 'gswin64c')
    'pypdf'        = Test-PyModule 'pypdf'
    'cryptography' = Test-PyModule 'cryptography'
    'OCRmyPDF'     = Test-PyModule 'ocrmypdf'
    'pdfplumber'   = Test-PyModule 'pdfplumber'
    'openpyxl'     = Test-PyModule 'openpyxl'
    'Pillow'       = Test-PyModule 'PIL'
}
foreach ($k in $final.Keys) { Write-Check $k $final[$k] '' }

$stillMissing = @($final.Keys | Where-Object { -not $final[$_] })
Write-Host ''
if ($stillMissing.Count -eq 0) {
    Write-Host 'All set.' -ForegroundColor Green
    Add-ShortcutIfWanted
    Write-Host ''
    Write-Host 'Start the toolbox from the desktop shortcut, or "Launch PDFToolbox.bat".' -ForegroundColor Green
} else {
    Write-Host ('Still missing: ' + ($stillMissing -join ', ')) -ForegroundColor Red
    Write-Host 'The toolbox will still run, but the parts that need these will be switched off.' -ForegroundColor Yellow
    foreach ($n in $script:Notes) { Write-Host "  - $n" -ForegroundColor Yellow }
    Show-ManualSteps
}
Write-Host ''
[void](Read-Host 'Press Enter to close')
