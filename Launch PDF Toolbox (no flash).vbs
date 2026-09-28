' Opens the PDF Toolbox window with no console flash at all.
' Keep this file in the same folder as PDFToolboxApp.ps1.
Set shell = CreateObject("WScript.Shell")
folder = Left(WScript.ScriptFullName, InStrRev(WScript.ScriptFullName, "\"))
shell.Run "powershell.exe -NoProfile -ExecutionPolicy Bypass -STA -WindowStyle Hidden -File """ & folder & "PDFToolboxApp.ps1""", 0, False
