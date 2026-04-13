@echo off
echo ============================================
echo  BioSync EXE Builder
echo ============================================
echo.
echo [1] Installing packages...
pip install pyinstaller pyodbc requests tkcalendar --quiet
echo.
echo [2] Building BioSync.exe...
pyinstaller --onefile --windowed --name BioSync --hidden-import tkcalendar --hidden-import babel.numbers biosync_app.py
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
