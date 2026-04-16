#!/bin/bash
# PPT Voice Control (English) — macOS build script
#
# Usage:
#   chmod +x build_app_en.sh && ./build_app_en.sh
#
# Output: dist/PPT-Voice-Control/

set -e
cd "$(dirname "$0")"

echo "========================================="
echo "  PPT Voice Control (English) — macOS build"
echo "========================================="
echo

pyinstaller \
    --name "PPT-Voice-Control" \
    --windowed \
    --onedir \
    --noconfirm \
    --clean \
    --hidden-import sounddevice \
    --hidden-import numpy \
    --hidden-import funasr \
    --hidden-import torch \
    --hidden-import torchaudio \
    --hidden-import modelscope \
    --collect-all sounddevice \
    --collect-all funasr \
    main_en.py

echo
echo "========================================="
echo "  Build complete."
echo "  App bundle: dist/PPT-Voice-Control/"
echo "========================================="
echo
echo "Notes:"
echo "  1. On first launch, grant Accessibility permission"
echo "     (System Settings -> Privacy & Security -> Accessibility)."
echo "  2. Also allow microphone access when prompted."
