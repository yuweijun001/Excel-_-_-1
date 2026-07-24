@echo off
setlocal
cd /d "%~dp0"

where py >nul 2>nul
if not errorlevel 1 (
  set "PY=py"
) else (
  where python >nul 2>nul
  if errorlevel 1 (
    echo Python was not found.
    echo Install Python 3.10 or newer and select Add Python to PATH.
    pause
    exit /b 1
  )
  set "PY=python"
)

%PY% -m pip install --upgrade pip
if errorlevel 1 goto :failed

%PY% -m pip install -r requirements.txt
if errorlevel 1 goto :failed

%PY% -m PyInstaller --noconfirm --clean --onefile --windowed ^
  --name "ExcelWorkbookSplitter" ^
  --icon "assets\app.ico" ^
  --add-data "assets\app.ico;." ^
  workbook_splitter_gui.py
if errorlevel 1 goto :failed

echo.
echo Build completed.
echo EXE: %CD%\dist\ExcelWorkbookSplitter.exe
explorer "%CD%\dist"
pause
exit /b 0

:failed
echo.
echo Build failed. Review the messages above.
pause
exit /b 1
