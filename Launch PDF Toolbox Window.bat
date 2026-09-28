@echo off
REM Opens the PDF Toolbox window. Keep all files in the same folder.
REM -WindowStyle Hidden keeps the PowerShell console off screen.
start "" powershell.exe -NoProfile -ExecutionPolicy Bypass -STA -WindowStyle Hidden -File "%~dp0PDFToolboxApp.ps1"
