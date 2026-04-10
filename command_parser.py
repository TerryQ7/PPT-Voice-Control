"""PPT语音控制助手 - 命令解析器

支持中英文语音命令解析，包括：
- 下一页 / next page
- 上一页 / previous page
- 第N页 / go to page N（支持中文数字）
- 第一页 / 最后一页
"""

import re
import time
from dataclasses import dataclass
from enum import Enum
from typing import Optional

from config import (
    DEBOUNCE_SECONDS,
    FIRST_KEYWORDS,
    LAST_KEYWORDS,
    NEXT_KEYWORDS,
    PREV_KEYWORDS,
)


class CommandType(Enum):
    NEXT = "next"
    PREV = "prev"
    GOTO = "goto"
    FIRST = "first"
    LAST = "last"


@dataclass
class Command:
    type: CommandType
    page: Optional[int] = None

    def __str__(self):
        if self.type == CommandType.GOTO:
            return f"跳转到第{self.page}页"
        labels = {
            CommandType.NEXT: "下一页",
            CommandType.PREV: "上一页",
            CommandType.FIRST: "第一页",
            CommandType.LAST: "最后一页",
        }
        return labels.get(self.type, str(self.type))


# ==================== 中文数字转阿拉伯数字 ====================

CN_DIGIT_MAP = {
    "零": 0, "〇": 0,
    "一": 1, "壹": 1,
    "二": 2, "贰": 2, "两": 2,
    "三": 3, "叁": 3,
    "四": 4, "肆": 4,
    "五": 5, "伍": 5,
    "六": 6, "陆": 6,
    "七": 7, "柒": 7,
    "八": 8, "捌": 8,
    "九": 9, "玖": 9,
}

CN_UNIT_MAP = {
    "十": 10, "拾": 10,
    "百": 100, "佰": 100,
    "千": 1000, "仟": 1000,
}


def chinese_to_int(cn_str: str) -> Optional[int]:
    """将中文数字字符串转换为整数。

    支持: 一, 十二, 二十三, 一百二十三, 两百, 三十 等
    """
    if not cn_str:
        return None

    cn_str = cn_str.strip()

    # 处理纯阿拉伯数字
    if cn_str.isdigit():
        return int(cn_str)

    result = 0
    current = 0

    for char in cn_str:
        if char in CN_DIGIT_MAP:
            current = CN_DIGIT_MAP[char]
        elif char in CN_UNIT_MAP:
            unit = CN_UNIT_MAP[char]
            if current == 0:
                # 处理 "十二" 这种省略 "一" 的情况
                current = 1
            result += current * unit
            current = 0
        else:
            continue

    result += current

    return result if result > 0 else None


# ==================== 跳转页码正则模式 ====================

GOTO_PATTERNS_CN = [
    re.compile(r"(?:回到|跳到|翻到|去|到|转到)?第\s*([一二三四五六七八九十百千两零〇壹贰叁肆伍陆柒捌玖拾佰仟\d]+)\s*[页张]"),
    re.compile(r"(?:go\s*to|jump\s*to|slide|page)\s*(\d+)", re.IGNORECASE),
]


class CommandParser:
    """语音命令解析器，带去重（debounce）功能。"""

    def __init__(self, debounce_sec: float = DEBOUNCE_SECONDS):
        self.debounce_sec = debounce_sec
        self._last_command: Optional[Command] = None
        self._last_time: float = 0

    # FunASR 可能输出的标点符号
    _PUNCTUATION = str.maketrans("", "", "。，、？！；：""''（）《》【】…—·,.?!;:\"'()[]")

    def parse(self, text: str) -> Optional[Command]:
        """解析文本，返回命令。如果在去重窗口内返回相同命令，则返回 None。"""
        text = text.strip().translate(self._PUNCTUATION)
        if not text:
            return None

        cmd = self._extract_command(text)
        if cmd is None:
            return None

        return self._debounce(cmd)

    def _extract_command(self, text: str) -> Optional[Command]:
        # FunASR 中文输出带空格（"下 一 页"），去空格用于中文匹配
        text_compact = text.replace(" ", "")
        text_lower = text.lower()

        for pattern in GOTO_PATTERNS_CN:
            match = pattern.search(text_compact)
            if match:
                num_str = match.group(1)
                page = chinese_to_int(num_str)
                if page is not None and page > 0:
                    if page == 1:
                        return Command(type=CommandType.FIRST)
                    return Command(type=CommandType.GOTO, page=page)

        for kw in LAST_KEYWORDS:
            if kw in text_compact or kw in text_lower:
                return Command(type=CommandType.LAST)

        for kw in FIRST_KEYWORDS:
            if kw in text_compact or kw in text_lower:
                return Command(type=CommandType.FIRST)

        for kw in NEXT_KEYWORDS:
            if kw in text_compact or kw in text_lower:
                return Command(type=CommandType.NEXT)

        for kw in PREV_KEYWORDS:
            if kw in text_compact or kw in text_lower:
                return Command(type=CommandType.PREV)

        return None

    def _debounce(self, cmd: Command) -> Optional[Command]:
        now = time.time()
        if (
            self._last_command is not None
            and self._last_command.type == cmd.type
            and self._last_command.page == cmd.page
            and (now - self._last_time) < self.debounce_sec
        ):
            return None

        self._last_command = cmd
        self._last_time = now
        return cmd

    def reset(self):
        self._last_command = None
        self._last_time = 0
