@echo off
chcp 65001 >nul
setlocal
cd /d "%~dp0"
title Build PPT-Presenter

where python >nul 2>nul
if errorlevel 1 (
  echo [!] Python not found. Install Python 3 from https://www.python.org and tick "Add to PATH".
  pause
  exit /b 1
)

if not exist ".venv\Scripts\python.exe" (
  echo [1/4] Creating virtual environment...
  python -m venv .venv
)

rem ---- Version: taken from the first "## [X.Y]" heading in CHANGELOG.md
echo [2/4] Reading version from CHANGELOG.md...
".venv\Scripts\python.exe" version.py > "%TEMP%\ppt-presenter-ver.txt"
if errorlevel 1 (
  echo [!] No version found in CHANGELOG.md. Add a heading like:  ## [1.8] - 2026-10-14
  pause
  exit /b 1
)
set /p VER=<"%TEMP%\ppt-presenter-ver.txt"
del "%TEMP%\ppt-presenter-ver.txt" >nul 2>nul
set "NAME=PPT-Presenter-V%VER%"
echo        Version %VER%  -^>  %NAME%.exe

echo [3/4] Installing packages...
".venv\Scripts\python.exe" -m pip install -q --disable-pip-version-check -r requirements.txt pyinstaller
if errorlevel 1 (
  echo [!] Failed to install packages. Check your internet connection.
  pause
  exit /b 1
)

echo [4/4] Building %NAME%.exe, please wait...
if exist build rmdir /s /q build
if exist "dist\%NAME%.exe" del /q "dist\%NAME%.exe"
if exist "dist\%NAME%.exe" (
  echo [!] dist\%NAME%.exe is still running. Close the control panel first.
  pause
  exit /b 1
)
".venv\Scripts\python.exe" -m PyInstaller --noconfirm --clean --onefile --windowed ^
  --name "%NAME%" ^
  --add-data "static;static" ^
  --add-data "CHANGELOG.md;." ^
  --collect-all webview ^
  --hidden-import win32com.client ^
  --hidden-import win32com.server.util ^
  --hidden-import pythoncom ^
  --hidden-import pywintypes ^
  --hidden-import waitress ^
  --exclude-module tkinter ^
  server.py
if errorlevel 1 (
  echo.
  echo [!] Build failed. See the messages above.
  pause
  exit /b 1
)

if exist build rmdir /s /q build
if exist "%NAME%.spec" del /q "%NAME%.spec"

echo.
echo ============================================================
echo  Done:  %CD%\dist\%NAME%.exe
echo.
echo  Copy the exe into the folder that holds your PPT files and
echo  double-click it. config.json, the thumbnail cache and
echo  ppt-presenter.log will be created next to the exe.
echo ============================================================
explorer "%CD%\dist"
pause
