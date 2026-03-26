@echo off
REM ─────────────────────────────────────────────────────────────────────────────
REM  Build script for MCMX Quote Generator  —  Windows
REM  Run from the project folder:  build_windows.bat
REM ─────────────────────────────────────────────────────────────────────────────

echo [1/3] Cleaning previous build...
if exist build   rmdir /s /q build
if exist dist    rmdir /s /q dist

echo [2/3] Running PyInstaller...
pyinstaller quote_generator.spec

echo [3/3] Done.
echo.
echo Output:  dist\QuoteGenerator\QuoteGenerator.exe
echo Zip the dist\QuoteGenerator folder and distribute it.
pause
