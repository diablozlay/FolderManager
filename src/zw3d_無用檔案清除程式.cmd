@echo off
rem ZW3D junk file cleaner launcher: runs the .ps1 next to this file with Windows PowerShell 5.1.
rem The PowerShell console is hidden; only the program window is shown.
cd /d "%~dp0"
for %%F in ("%~dp0*.ps1") do (
  start "" powershell.exe -NoProfile -ExecutionPolicy Bypass -STA -WindowStyle Hidden -File "%%~fF"
  goto :eof
)
echo Program file (.ps1) not found next to this launcher.
pause
