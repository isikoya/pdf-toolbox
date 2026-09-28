@echo off
REM Installs everything the PDF Toolbox needs. Skips anything already installed.
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0InstallPDFToolbox.ps1"
