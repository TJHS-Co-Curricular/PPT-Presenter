@echo off
chcp 65001 >nul
cd /d "%~dp0"
title PPT Control Panel

where python >nul 2>nul
if errorlevel 1 (
  echo [!] Python not found. Install Python 3 from https://www.python.org and tick "Add to PATH".
  pause
  exit /b 1
)

if not exist ".venv\Scripts\python.exe" (
  echo First run: creating virtual environment...
  python -m venv .venv
)
".venv\Scripts\python.exe" -m pip install -q --disable-pip-version-check -r requirements.txt
if errorlevel 1 (
  echo [!] Failed to install packages. Check your internet connection.
  pause
  exit /b 1
)

".venv\Scripts\python.exe" server.py
pause
