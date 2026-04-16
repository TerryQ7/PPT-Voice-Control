@echo off
REM PPT Voice Control (English) - Windows build
REM
REM Usage: double-click or run build_app_windows_en.bat
REM Output: dist\PPT-Voice-Control\

echo =========================================
echo   PPT Voice Control (English) - Windows build
echo =========================================
echo.

cd /d "%~dp0"

echo Building...
echo.

pyinstaller ^
    --name "PPT-Voice-Control" ^
    --windowed ^
    --onedir ^
    --noconfirm ^
    --clean ^
    --hidden-import sounddevice ^
    --hidden-import numpy ^
    --hidden-import funasr ^
    --hidden-import torch ^
    --hidden-import torchaudio ^
    --hidden-import modelscope ^
    --hidden-import pynput ^
    --hidden-import pynput.keyboard ^
    --hidden-import pynput.keyboard._win32 ^
    --collect-all sounddevice ^
    --collect-all funasr ^
    main_en.py

echo.
echo =========================================
echo   Build complete.
echo   Output folder: dist\PPT-Voice-Control\
echo =========================================
echo.
pause
