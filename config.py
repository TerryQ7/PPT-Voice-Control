"""PPT语音控制助手 - 全局配置"""

import os

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
MODEL_DIR = os.path.join(BASE_DIR, "models")

# ==================== ASR 引擎选择 ====================
# "funasr" — 阿里 Paraformer（高精度，中英双语，需 PyTorch ~2GB）
# "vosk"   — Vosk（轻量，打包体积小）
ASR_ENGINE = "funasr"

# ==================== Vosk 配置（仅 ASR_ENGINE="vosk" 时使用） ====================
VOSK_MODEL_CN = "vosk-model-cn-0.22"
VOSK_MODEL_CN_SMALL = "vosk-model-small-cn-0.22"

VOSK_MODEL_URLS = {
    "vosk-model-cn-0.22": "https://alphacephei.com/vosk/models/vosk-model-cn-0.22.zip",
    "vosk-model-small-cn-0.22": "https://alphacephei.com/vosk/models/vosk-model-small-cn-0.22.zip",
}

DEFAULT_VOSK_MODEL = VOSK_MODEL_CN
VOSK_MODEL_PATH = os.path.join(MODEL_DIR, DEFAULT_VOSK_MODEL)

# ==================== 音频参数 ====================
SAMPLE_RATE = 16000
BLOCK_SIZE = 4000  # Vosk 每次处理 0.25 秒音频

# ==================== 命令去重 ====================
DEBOUNCE_SECONDS = 2.0

# ==================== 中英文命令关键词 ====================
NEXT_KEYWORDS = [
    "下一页", "下一张", "下一个", "翻到下一页", "往后翻",
    "next page", "next slide", "next one",
]

PREV_KEYWORDS = [
    "上一页", "上一张", "上一个", "翻到上一页", "往前翻", "前一页",
    "previous page", "previous slide", "last page", "go back",
]

FIRST_KEYWORDS = [
    "第一页", "第一张", "回到开头", "回到第一页", "从头开始",
    "first page", "first slide",
]

LAST_KEYWORDS = [
    "最后一页", "最后一张", "跳到最后",
    "last slide", "final page", "final slide",
]
