@echo off
rem ============================================================================
rem Folder Manager build script (v1.5.2, repo layout)
rem Requires: Windows + .NET Framework 4.x (csc.exe), Windows PowerShell 5.1,
rem           Python 3.x with Pillow (pip install pillow)
rem Output  : ..\檔案管理工具.exe  (single-file portable exe, icon embedded)
rem ============================================================================
chcp 65001 >nul
setlocal
cd /d "%~dp0"
set CSC=%WINDIR%\Microsoft.NET\Framework64\v4.0.30319\csc.exe
set SMA=C:\Windows\Microsoft.NET\assembly\GAC_MSIL\System.Management.Automation\v4.0_3.0.0.0__31bf3856ad364e35\System.Management.Automation.dll
set OUT=..\檔案管理工具.exe

echo [1/4] build icon (..\圖示\app.ico)
python make_ico.py
if errorlevel 1 goto fail

echo [2/4] compile %OUT%
"%CSC%" /nologo /target:winexe /platform:anycpu /codepage:65001 /optimize+ ^
    /out:"%OUT%" /r:"%SMA%" /r:System.dll /r:System.Windows.Forms.dll ^
    /resource:FileManager.ps1 Program.cs
if errorlevel 1 goto fail

echo [3/4] inject icon
python w32_icon.py "%OUT%" ..\圖示\app.ico
if errorlevel 1 goto fail

echo [4/4] selftest
start /wait "" "%OUT%" -selftest
if errorlevel 1 goto fail
if exist "..\logs\runtime_errors.log" goto fail
if exist "..\logs\crash.log" goto fail

echo BUILD OK: %OUT%
endlocal
exit /b 0
:fail
echo BUILD FAILED
endlocal
exit /b 1
