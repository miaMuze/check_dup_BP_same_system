@echo off
chcp 65001 >nul
echo.
echo ==========================================
echo  BP Duplicate Checker - Install and Build
echo ==========================================
echo.

REM Check Python
python --version >nul 2>&1
if errorlevel 1 (
    echo ERROR: Python is not installed
    echo Please install Python 3.8+ from https://www.python.org/downloads/
    pause
    exit /b 1
)

echo [1/3] Installing dependencies...
pip install pandas openpyxl rapidfuzz pyinstaller --quiet

echo [2/3] Building executable...
python -m PyInstaller --noconfirm --onefile --windowed --name "BP_Duplicate_Checker" --add-data "src;src" --hidden-import=rapidfuzz --hidden-import=openpyxl --hidden-import=pandas --hidden-import=numpy main.py

echo [3/3] Cleaning up...
rmdir /s /q build 2>nul
del /q BP_Duplicate_Checker.spec 2>nul

echo.
if exist "dist\BP_Duplicate_Checker.exe" (
    echo ==========================================
    echo  BUILD SUCCESS!
    echo ==========================================
    echo.
    echo EXE file: dist\BP_Duplicate_Checker.exe
    echo.
    start "" "dist"
) else (
    echo BUILD FAILED! Check errors above.
)

pause
