#!/bin/bash
# PPT语音控制助手 - macOS 打包脚本
#
# 使用方法:
#   chmod +x build_app.sh && ./build_app.sh
#
# 输出: dist/PPT语音控制助手/

set -e
cd "$(dirname "$0")"

echo "========================================="
echo "  PPT语音控制助手 - macOS 打包"
echo "========================================="
echo

pyinstaller \
    --name "PPT语音控制助手" \
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
    main.py

echo
echo "========================================="
echo "  ✓ 打包完成！"
echo "  应用位于: dist/PPT语音控制助手/"
echo "========================================="
echo
echo "注意："
echo "1. 首次运行需在「系统设置 → 隐私与安全 → 辅助功能」中授权"
echo "2. 首次运行需允许麦克风访问权限"
