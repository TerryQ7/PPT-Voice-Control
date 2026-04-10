@echo off
setlocal enabledelayedexpansion

REM Offline Windows build:
REM - Bundles Python runtime and all dependencies via PyInstaller
REM - Bundles FunASR model files so target machines need no internet
REM - Target machines do NOT need Python installed

cd /d "%~dp0"

set "APP_NAME=PPT-Voice-Control-Offline"
set "MODEL_DIR=%CD%\funasr_model"
set "MODEL_CACHE=%USERPROFILE%\.cache\modelscope\hub\models\iic\speech_paraformer-large_asr_nat-zh-cn-16k-common-vocab8404-pytorch"
set "DIST_DIR=%CD%\dist\%APP_NAME%"
set "RELEASE_DIR=%CD%\release\%APP_NAME%"
set "ZIP_PATH=%CD%\release\%APP_NAME%.zip"

echo =========================================
echo   Building offline Windows package
echo =========================================
echo.

if not exist "%MODEL_DIR%" mkdir "%MODEL_DIR%"

if not exist "%MODEL_DIR%\model.pt" (
    echo [1/5] Local model not found, trying to copy from cache...
    if exist "%MODEL_CACHE%\model.pt" (
        robocopy "%MODEL_CACHE%" "%MODEL_DIR%" /E >nul
    ) else (
        echo ERROR: FunASR model not found.
        echo Please run app once with internet, let model download finish, then retry.
        exit /b 1
    )
)

if not exist "%MODEL_DIR%\configuration.json" (
    echo ERROR: Missing "%MODEL_DIR%\configuration.json"
    exit /b 1
)
if not exist "%MODEL_DIR%\config.yaml" (
    echo ERROR: Missing "%MODEL_DIR%\config.yaml"
    exit /b 1
)
if not exist "%MODEL_DIR%\model.pt" (
    echo ERROR: Missing "%MODEL_DIR%\model.pt"
    exit /b 1
)

echo [2/5] Model folder validated.

if exist "%CD%\venv\Scripts\pyinstaller.exe" (
    set "PYI=%CD%\venv\Scripts\pyinstaller.exe"
) else (
    set "PYI=pyinstaller"
)

echo [3/5] Running PyInstaller...
"%PYI%" ^
    --name "%APP_NAME%" ^
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
    --add-data "funasr_model;funasr_model" ^
    main.py
if errorlevel 1 exit /b 1

if not exist "%DIST_DIR%\_internal\base_library.zip" (
    echo ERROR: Build output is incomplete (missing base_library.zip).
    exit /b 1
)

echo [4/5] Preparing release directory...
if exist "%RELEASE_DIR%" rmdir /s /q "%RELEASE_DIR%"
mkdir "%RELEASE_DIR%"
robocopy "%DIST_DIR%" "%RELEASE_DIR%" /E >nul

echo [5/5] Creating zip package...
if exist "%ZIP_PATH%" del /f /q "%ZIP_PATH%"
powershell -NoProfile -Command "Compress-Archive -Path '%RELEASE_DIR%\*' -DestinationPath '%ZIP_PATH%' -Force"

echo.
echo =========================================
echo   Offline package build complete
echo -----------------------------------------
echo   Folder: %RELEASE_DIR%
echo   Zip   : %ZIP_PATH%
echo =========================================
echo.

endlocal
