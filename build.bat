@echo off
REM ============================================================
REM  Excel Splitter -- Windows build script
REM
REM  Builds ExcelSplitter.exe using PyInstaller. Run this from a
REM  Command Prompt / PowerShell window, from the project root
REM  (the folder that contains this file).
REM
REM  First-time setup (only needs to be done once):
REM      python -m venv venv
REM      venv\Scripts\activate
REM      pip install -r requirements.txt
REM
REM  Then, any time you want to (re)build the .exe:
REM      build.bat
REM ============================================================

setlocal

echo.
echo === Excel Splitter build ===
echo.

where python >nul 2>nul
if errorlevel 1 (
    echo ERROR: Python was not found on PATH.
    echo Install Python 3.12+ from https://www.python.org/downloads/ and try again.
    exit /b 1
)

if not exist "venv\Scripts\activate.bat" (
    echo No virtual environment found at .\venv
    echo Run this first:
    echo     python -m venv venv
    echo     venv\Scripts\activate
    echo     pip install -r requirements.txt
    exit /b 1
)

call venv\Scripts\activate.bat

echo.
echo Cleaning previous build output...
if exist "build" rmdir /s /q "build"
if exist "dist" rmdir /s /q "dist"

echo.
echo Running PyInstaller...
pyinstaller ExcelSplitter.spec --noconfirm

if errorlevel 1 (
    echo.
    echo BUILD FAILED. Scroll up for the PyInstaller error output.
    exit /b 1
)

echo.
echo ============================================================
echo  Build succeeded.
echo  Your executable is at:  dist\ExcelSplitter.exe
echo ============================================================

endlocal
