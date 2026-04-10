@echo off
REM PPT语音控制助手 - Windows 打包脚本
REM
REM 使用方法: 双击运行或在命令行执行 build_app_windows.bat
REM 输出: dist\PPT语音控制助手\

echo =========================================
echo   PPT语音控制助手 - Windows 打包
echo =========================================
echo.

cd /d "%~dp0"

echo 开始打包...
echo.

pyinstaller ^
    --name "PPT语音控制助手" ^
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
    main.py

echo.
echo =========================================
echo   打包完成！
echo   应用位于: dist\PPT语音控制助手\
echo =========================================
echo.
pause
