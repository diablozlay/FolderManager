@echo off
set "PYDIR=%LOCALAPPDATA%\Programs\Python\Python315"
set "PATH=%PYDIR%;%PYDIR%\Scripts;%PATH%"
set "LOG=%~dp0_build_output.txt"
> "%LOG%" echo === python ===
where python >> "%LOG%" 2>&1
python --version >> "%LOG%" 2>&1
python -c "import PIL; print(PIL.__version__)" >> "%LOG%" 2>&1
if errorlevel 1 (
  echo === pip install pillow === >> "%LOG%"
  python -m pip install --user pillow >> "%LOG%" 2>&1
)
echo === build.cmd === >> "%LOG%"
call "%~dp0build.cmd" >> "%LOG%" 2>&1
echo EXIT=%ERRORLEVEL% >> "%LOG%"
