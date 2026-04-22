@echo off
echo ============================================
echo  BioSync EXE Builder
echo ============================================
echo.

echo [1] Installing packages...
pip install pyinstaller pyodbc requests tkcalendar babel --quiet

echo.
echo [2] Building BioSync.exe...
pyinstaller ^
  --onefile ^
  --windowed ^
  --name BioSync ^
  --hidden-import tkcalendar ^
  --hidden-import babel ^
  --hidden-import babel.numbers ^
  --hidden-import babel.dates ^
  --hidden-import babel.core ^
  --hidden-import requests ^
  --hidden-import pyodbc ^
  --hidden-import json ^
  --hidden-import threading ^
  --hidden-import subprocess ^
  --collect-data tkcalendar ^
  --collect-data babel ^
  biosync_app.py

echo.
echo ============================================
if exist "dist\BioSync.exe" (
    echo  SUCCESS!
    copy "dist\BioSync.exe" "BioSync.exe" >nul
    echo  BioSync.exe is ready in this folder!
) else (
    echo  Build failed. Check errors above.
)
echo ============================================
pause