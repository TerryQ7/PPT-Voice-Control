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


def english_words_to_int(number_text: str) -> Optional[int]:
    """将英文数字/序数短语转换成整数。"""
    if not number_text:
        return None

    text = number_text.lower().strip().replace("-", " ")
    if not text:
        return None

    if text.isdigit():
        return int(text)

    tokens = [tok for tok in text.split() if tok]
    if not tokens:
        return None

    total = 0
    current = 0

    for tok in tokens:
        if tok in _EN_NUMBER_CONNECTORS:
            continue
        if tok in ("a", "an"):
            if current == 0:
                current = 1
            continue
        if tok in _EN_NUMBER_SMALL:
            current += _EN_NUMBER_SMALL[tok]
            continue
        if tok in _EN_ORDINAL_SMALL:
            current += _EN_ORDINAL_SMALL[tok]
            continue
        if tok in _EN_NUMBER_TENS:
            current += _EN_NUMBER_TENS[tok]
            continue
        if tok in _EN_ORDINAL_TENS:
            current += _EN_ORDINAL_TENS[tok]
            continue
        if tok == "hundred":
            if current == 0:
                current = 1
            current *= 100
            continue
        if tok == "thousand":
            if current == 0:
                current = 1
            total += current * 1000
            current = 0
            continue
        if tok == "hundredth":
            if current == 0:
                current = 1
            current *= 100
            continue
        if tok == "thousandth":
            if current == 0:
                current = 1
            total += current * 1000
            current = 0
            continue
        return None

    value = total + current
    return value if value > 0 else None


# ==================== 跳转页码正则模式 ====================

# 带显式跳转前缀的命令
_GOTO_EXPLICIT_CN = re.compile(
    r"(?:回到|跳到|翻到|去|到|转到)第\s*([一二三四五六七八九十百千两零〇壹贰叁肆伍陆柒捌玖拾佰仟\d]+)\s*[页张]"
)
# 无前缀的"第N页"
_GOTO_BARE_CN = re.compile(
    r"第\s*([一二三四五六七八九十百千两零〇壹贰叁肆伍陆柒捌玖拾佰仟\d]+)\s*[页张]"
)
# 英文：区分带跳转前缀和不带的
_GOTO_EXPLICIT_EN = re.compile(
    r"(?:go\s*to|jump\s*to)\s*(?:page|slide)?\s*(\d+)", re.IGNORECASE,
)
_GOTO_BARE_EN = re.compile(r"(?:slide|page)\s*(\d+)", re.IGNORECASE)
_GOTO_EXPLICIT_EN_SUFFIX = re.compile(
    r"(?:go\s*to|jump\s*to)\s*(\d+)\s*(?:page|slide)", re.IGNORECASE,
)

_EN_NUMBER_SMALL = {
    "zero": 0,
    "one": 1,
    "two": 2,
    "three": 3,
    "four": 4,
    "five": 5,
    "six": 6,
    "seven": 7,
    "eight": 8,
    "nine": 9,
    "ten": 10,
    "eleven": 11,
    "twelve": 12,
    "thirteen": 13,
    "fourteen": 14,
    "fifteen": 15,
    "sixteen": 16,
    "seventeen": 17,
    "eighteen": 18,
    "nineteen": 19,
}
_EN_NUMBER_TENS = {
    "twenty": 20,
    "thirty": 30,
    "forty": 40,
    "fifty": 50,
    "sixty": 60,
    "seventy": 70,
    "eighty": 80,
    "ninety": 90,
}
_EN_ORDINAL_SMALL = {
    "first": 1,
    "second": 2,
    "third": 3,
    "fourth": 4,
    "fifth": 5,
    "sixth": 6,
    "seventh": 7,
    "eighth": 8,
    "ninth": 9,
    "tenth": 10,
    "eleventh": 11,
    "twelfth": 12,
    "thirteenth": 13,
    "fourteenth": 14,
    "fifteenth": 15,
    "sixteenth": 16,
    "seventeenth": 17,
    "eighteenth": 18,
    "nineteenth": 19,
}
_EN_ORDINAL_TENS = {
    "twentieth": 20,
    "thirtieth": 30,
    "fortieth": 40,
    "fiftieth": 50,
    "sixtieth": 60,
    "seventieth": 70,
    "eightieth": 80,
    "ninetieth": 90,
}
_EN_NUMBER_CONNECTORS = {"and"}
_EN_NUMBER_WORDS = (
    set(_EN_NUMBER_SMALL)
    | set(_EN_NUMBER_TENS)
    | set(_EN_ORDINAL_SMALL)
    | set(_EN_ORDINAL_TENS)
    | {"hundred", "thousand", "hundredth", "thousandth", "a", "an", "and"}
)
_EN_NUMBER_WORD_RE = "(?:" + "|".join(
    sorted((re.escape(w) for w in _EN_NUMBER_WORDS), key=len, reverse=True)
) + ")"
_EN_NUMBER_PHRASE_RE = rf"{_EN_NUMBER_WORD_RE}(?:[\s-]+{_EN_NUMBER_WORD_RE})*"
_GOTO_EXPLICIT_EN_WORDS = re.compile(
    rf"(?:go\s*to|jump\s*to)\s*(?:page|slide)?\s*({_EN_NUMBER_PHRASE_RE})\s*(?:page|slide)?",
    re.IGNORECASE,
)
_GOTO_BARE_EN_WORDS_PREFIX = re.compile(
    rf"(?:slide|page)\s*({_EN_NUMBER_PHRASE_RE})",
    re.IGNORECASE,
)
_GOTO_BARE_EN_WORDS_SUFFIX = re.compile(
    rf"({_EN_NUMBER_PHRASE_RE})\s*(?:slide|page)",
    re.IGNORECASE,
)

# 长度预算按语言分别设置：
#   - 中文一字一意，命令本身就 3-4 字（"下一页"），多余信息超过 4 字
#     基本就不是命令了（"我们下一页吧" 这种边界情况也在 4 字以内）；
#   - 英文则常带"载体短语"（"continue with next page"、"please go to
#     next page"、"could you switch to slide three" 等），需要 20+ 字
#     的余量；过窄会让英文命令直接被丢弃。
_MAX_EXTRA_CHARS_CN = 4
_MAX_EXTRA_CHARS_EN = 24


def _is_chinese_token(token: str) -> bool:
    """判断给定关键词是否为中文（包含至少一个 CJK 字符）。"""
    return any("\u4e00" <= ch <= "\u9fff" for ch in token)


# 第二道防线：描述性上下文正则。中英文各一套，分别只对各自语言的命令生效，
# 避免英文叙述句被中文规则错误放行（反之亦然）。
_DESCRIPTIVE_CN = re.compile(
    r"像.*第.+[页张]"
    r"|如同.*第.+[页张]"
    r"|就像.*第.+[页张]"
    r"|正如.*第.+[页张]"
    r"|第.+[页张].*(?:那样|一样|似的|所述|所示|描述|说的|说过|提到|展示)"
    r"|在第.+[页张][中里上内]"
    r"|和第.+[页张].*(?:一样|相同|类似)"
    r"|跟第.+[页张].*(?:一样|相同|类似)"
    r"|见第.+[页张]"
    r"|参[考见照]第.+[页张]"
)
_DESCRIPTIVE_EN = re.compile(
    # ① 围绕带数字页码的叙述：as shown on page 3 / like the page 5
    r"(?:like|as|described|mentioned|shown|see|refer|similar\s+to)\s.*?(?:page|slide)\s*\d+"
    # ② 数字页码后接叙述动词：page 3 was discussed / slide 5 contains
    r"|(?:page|slide)\s*\d+\s.*?(?:described|mentioned|shown|said|stated|contains?|shows?|appears?|talked|discussed|presented|displays?)"
    # ③ 不带数字的叙述句："the next page is interesting" / "previous slide shows ..."
    r"|\b(?:the|this|that|previous|next|last|first|any|every|each)\s+(?:page|slide)\s+(?:is|was|were|will|would|has|have|had|contains?|shows?|describes?|displays?|presents?|appears?|talks?|discusses?|illustrates?)\b"
    # ④ 比喻："like the next page" / "similar to the previous slide"
    r"|\b(?:like|as|similar\s+to)\s+(?:the\s+)?(?:previous|next|last|first|same|other)\s+(?:page|slide)\b",
    re.IGNORECASE,
)


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
        text_lower = text.lower().strip()
        text_normalized = " ".join(text_lower.split())

        is_desc_cn = bool(_DESCRIPTIVE_CN.search(text_compact))
        is_desc_en = bool(_DESCRIPTIVE_EN.search(text_normalized))

        # --- 跳转类命令 ---
        # 中文路径：用中文长度预算 + 中文描述性过滤
        # 英文路径：用英文长度预算 + 英文描述性过滤
        # 两条路径完全独立，互不干扰。

        # 1) 中文显式前缀（"跳到第三页"）
        if not is_desc_cn:
            match = _GOTO_EXPLICIT_CN.search(text_compact)
            if match and len(text_compact) - len(match.group(0)) <= _MAX_EXTRA_CHARS_CN:
                page = chinese_to_int(match.group(1))
                if page is not None and page > 0:
                    return self._page_cmd(page)

        # 2) 中文无前缀（"第三页"）
        if not is_desc_cn:
            match = _GOTO_BARE_CN.search(text_compact)
            if match and len(text_compact) - len(match.group(0)) <= _MAX_EXTRA_CHARS_CN:
                page = chinese_to_int(match.group(1))
                if page is not None and page > 0:
                    return self._page_cmd(page)

        # 3) 英文显式前缀（"go to page 3" / "jump to 5"）
        if not is_desc_en:
            for pattern in (_GOTO_EXPLICIT_EN, _GOTO_EXPLICIT_EN_SUFFIX):
                match = pattern.search(text_normalized)
                if match and len(text_normalized) - len(match.group(0)) <= _MAX_EXTRA_CHARS_EN:
                    page = int(match.group(1))
                    if page > 0:
                        return self._page_cmd(page)

            match = _GOTO_EXPLICIT_EN_WORDS.search(text_normalized)
            if match and len(text_normalized) - len(match.group(0)) <= _MAX_EXTRA_CHARS_EN:
                page = english_words_to_int(match.group(1))
                if page is not None and page > 0:
                    return self._page_cmd(page)

        # 4) 英文无前缀（"page 3" / "slide 5"）
        if not is_desc_en:
            match = _GOTO_BARE_EN.search(text_normalized)
            if match and len(text_normalized) - len(match.group(0)) <= _MAX_EXTRA_CHARS_EN:
                page = int(match.group(1))
                if page > 0:
                    return self._page_cmd(page)
            for pattern in (_GOTO_BARE_EN_WORDS_PREFIX, _GOTO_BARE_EN_WORDS_SUFFIX):
                match = pattern.search(text_normalized)
                if match and len(text_normalized) - len(match.group(0)) <= _MAX_EXTRA_CHARS_EN:
                    page = english_words_to_int(match.group(1))
                    if page is not None and page > 0:
                        return self._page_cmd(page)

        # --- 关键词类命令 ---
        # _match_keyword 已经按关键词的语言分别检查描述性过滤和长度预算，
        # 不能再用 ``is_desc_cn or is_desc_en`` 做笼统拦截，否则任意一种语
        # 言的描述句会顺带屏蔽另一种语言的命令。

        for keywords, cmd_type in (
            (LAST_KEYWORDS, CommandType.LAST),
            (FIRST_KEYWORDS, CommandType.FIRST),
            (NEXT_KEYWORDS, CommandType.NEXT),
            (PREV_KEYWORDS, CommandType.PREV),
        ):
            if self._match_keyword(text_compact, text_normalized,
                                   keywords, is_desc_cn, is_desc_en):
                return Command(type=cmd_type)

        return None

    @staticmethod
    def _page_cmd(page: int) -> Command:
        if page == 1:
            return Command(type=CommandType.FIRST)
        return Command(type=CommandType.GOTO, page=page)

    @staticmethod
    def _match_keyword(text_compact: str, text_normalized: str,
                       keywords: list[str],
                       is_desc_cn: bool, is_desc_en: bool) -> bool:
        """关键词匹配，按关键词语言分别使用对应的预算和描述性过滤。

        - 中文关键词在 ``text_compact`` 中匹配，受 ``is_desc_cn`` 拦截，
          长度溢出阈值为 :data:`_MAX_EXTRA_CHARS_CN`；
        - 英文关键词在 ``text_normalized``（小写、单空格归一）中匹配，
          受 ``is_desc_en`` 拦截，长度溢出阈值为
          :data:`_MAX_EXTRA_CHARS_EN`。
        """
        for kw in keywords:
            if _is_chinese_token(kw):
                if is_desc_cn:
                    continue
                if kw in text_compact and \
                        len(text_compact) - len(kw) <= _MAX_EXTRA_CHARS_CN:
                    return True
            else:
                if is_desc_en:
                    continue
                kw_lower = kw.lower()
                if kw_lower in text_normalized and \
                        len(text_normalized) - len(kw_lower) <= _MAX_EXTRA_CHARS_EN:
                    return True
        return False

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
