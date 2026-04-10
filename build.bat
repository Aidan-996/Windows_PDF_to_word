@echo off
title PDF to Word - Build Tool

:: ── Log setup ────────────────────────────────────────────────────────────────
if not exist "build_log" mkdir build_log
for /f "tokens=1-3 delims=/ " %%a in ("%date%") do set _D=%%a%%b%%c
for /f "tokens=1-3 delims=:. " %%a in ("%time: =0%") do set _T=%%a%%b%%c
set LOG_FILE=build_log\build_%_D%_%_T%.log

call :LOG "================================================"
call :LOG "  PDF to Word - PyInstaller Package Tool"
call :LOG "================================================"
call :LOG ""
call :LOG "[INFO] Log file: %LOG_FILE%"
call :LOG ""

:: ── Check Python ─────────────────────────────────────────────────────────────
python --version >nul 2>&1
if errorlevel 1 (
    call :LOG "[ERROR] Python not found. Please install Python 3.8+"
    pause
    exit /b 1
)
for /f "delims=" %%v in ('python --version 2^>^&1') do call :LOG "[INFO] %%v detected"

:: ── Check poppler ─────────────────────────────────────────────────────────────
if not exist "poppler\Library\bin\pdftoppm.exe" (
    call :LOG "[ERROR] poppler\Library\bin\pdftoppm.exe not found."
    call :LOG "[ERROR] Make sure the poppler folder is in the same directory."
    pause
    exit /b 1
)
call :LOG "[INFO] poppler found OK"

:: ── Check main script ────────────────────────────────────────────────────────
if not exist "pdf_to_word.py" (
    call :LOG "[ERROR] pdf_to_word.py not found in current directory."
    pause
    exit /b 1
)
call :LOG "[INFO] pdf_to_word.py found OK"
call :LOG ""

:: ── Step 1: Install dependencies ─────────────────────────────────────────────
call :LOG "[1/4] Installing dependencies..."
pip install pdf2docx pdf2image Pillow python-docx pypdf easyocr numpy tkinterdnd2 >> "%LOG_FILE%" 2>&1
if errorlevel 1 (
    call :LOG "[WARNING] Some packages may have failed, check log for details."
) else (
    call :LOG "[INFO] Dependencies installed OK"
)

:: ── Step 2: Install PyInstaller ──────────────────────────────────────────────
call :LOG ""
call :LOG "[2/4] Installing PyInstaller..."
pip install pyinstaller >> "%LOG_FILE%" 2>&1
if errorlevel 1 (
    call :LOG "[ERROR] PyInstaller installation failed. See: %LOG_FILE%"
    pause
    exit /b 1
)
call :LOG "[INFO] PyInstaller installed OK"

:: ── Step 3: Package ──────────────────────────────────────────────────────────
call :LOG ""
call :LOG "[3/4] Packaging... (please wait, this may take several minutes)"

pyinstaller ^
    --onedir ^
    --windowed ^
    --name "PDF2Word" ^
    --add-data "poppler;poppler" ^
    --add-data "core;core" ^
    --collect-all easyocr ^
    --collect-all pdf2docx ^
    --collect-all pdf2image ^
    --collect-all docx ^
    --collect-all pypdf ^
    --collect-all tkinterdnd2 ^
    --hidden-import PIL ^
    --hidden-import PIL.Image ^
    --hidden-import PIL.ImageTk ^
    --hidden-import numpy ^
    --hidden-import re ^
    --hidden-import threading ^
    --hidden-import tkinter ^
    --hidden-import tkinter.ttk ^
    --hidden-import tkinter.filedialog ^
    --hidden-import tkinter.messagebox ^
    --hidden-import core ^
    --hidden-import core.deps ^
    --hidden-import core.converter ^
    --hidden-import core.theme ^
    --hidden-import core.app ^
    --hidden-import core.ui_annotation ^
    --hidden-import core.ui_preview ^
    --hidden-import core.ui_file_list ^
    --hidden-import core.ui_output_opts ^
    pdf_to_word.py >> "%LOG_FILE%" 2>&1

if errorlevel 1 (
    call :LOG ""
    call :LOG "[ERROR] Build FAILED. See full log: %LOG_FILE%"
    pause
    exit /b 1
)

:: ── Step 4: Done ─────────────────────────────────────────────────────────────
call :LOG ""
call :LOG "[4/4] Build complete!"
call :LOG ""
call :LOG "================================================"
call :LOG "  Output folder : dist\PDF2Word\"
call :LOG "  Share the entire PDF2Word folder to users."
call :LOG "  Users double-click PDF2Word.exe to launch."
call :LOG "------------------------------------------------"
call :LOG "  Full log saved : %LOG_FILE%"
call :LOG "================================================"

explorer dist
pause
exit /b 0

:: ── Subroutine: print to console AND append to log ───────────────────────────
:LOG
echo %~1
echo %~1 >> "%LOG_FILE%"
goto :eof
