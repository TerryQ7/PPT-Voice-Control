"""PPT语音控制助手 - PPT幻灯片控制

macOS 下使用 Quartz CGEvent 发送底层键盘事件（最可靠）。
Windows 下使用 pynput。
"""

import platform
import time
from typing import Optional

from command_parser import Command, CommandType

IS_MACOS = platform.system() == "Darwin"

# macOS virtual key codes
_MAC_KEY_CODES = {
    "right": 124, "left": 123, "up": 126, "down": 125,
    "return": 36, "escape": 53, "home": 115, "end": 119,
    "0": 29, "1": 18, "2": 19, "3": 20, "4": 21,
    "5": 23, "6": 22, "7": 26, "8": 28, "9": 25,
}


def check_accessibility_permission() -> bool:
    if not IS_MACOS:
        return True
    try:
        import subprocess
        result = subprocess.run(
            ["osascript", "-e",
             'tell application "System Events" to keystroke ""'],
            capture_output=True, timeout=3,
        )
        return result.returncode == 0
    except Exception:
        return True


class PPTController:
    """通过键盘事件控制PPT幻灯片放映。"""

    def __init__(self):
        if not IS_MACOS:
            from pynput.keyboard import Controller, Key
            self._keyboard = Controller()
            self._Key = Key

    def execute(self, cmd: Command):
        handler = {
            CommandType.NEXT: self._next_slide,
            CommandType.PREV: self._prev_slide,
            CommandType.FIRST: self._first_slide,
            CommandType.LAST: self._last_slide,
            CommandType.GOTO: self._goto_slide,
        }.get(cmd.type)

        if handler:
            if cmd.type == CommandType.GOTO:
                handler(cmd.page)
            else:
                handler()

    def _next_slide(self):
        self._send_key("right")

    def _prev_slide(self):
        self._send_key("left")

    def _first_slide(self):
        self._goto_slide(1)

    def _last_slide(self):
        if IS_MACOS:
            self._goto_slide(9999)
        else:
            self._send_key("end")

    def _goto_slide(self, page: Optional[int]):
        """在放映模式下输入页码 + Return 跳转到指定页。"""
        if page is None:
            return
        page_str = str(page)

        if IS_MACOS:
            for char in page_str:
                self._send_key(char)
                time.sleep(0.05)
            time.sleep(0.15)
            self._send_key("return")
        else:
            for char in page_str:
                self._keyboard.press(char)
                self._keyboard.release(char)
                time.sleep(0.05)
            time.sleep(0.1)
            self._keyboard.press(self._Key.enter)
            self._keyboard.release(self._Key.enter)

    def _send_key(self, key_name: str):
        if IS_MACOS:
            self._cgevent_key(key_name)
        else:
            key_map = {
                "right": self._Key.right, "left": self._Key.left,
                "up": self._Key.up, "down": self._Key.down,
                "return": self._Key.enter, "end": self._Key.end,
            }
            key = key_map.get(key_name)
            if key:
                self._keyboard.press(key)
                self._keyboard.release(key)

    @staticmethod
    def _cgevent_key(key_name: str):
        """通过 Quartz CGEvent 发送底层键盘事件。"""
        code = _MAC_KEY_CODES.get(key_name)
        if code is None:
            return
        from Quartz import CGEventCreateKeyboardEvent, CGEventPost, kCGHIDEventTap

        # Key down
        event_down = CGEventCreateKeyboardEvent(None, code, True)
        CGEventPost(kCGHIDEventTap, event_down)
        time.sleep(0.02)
        # Key up
        event_up = CGEventCreateKeyboardEvent(None, code, False)
        CGEventPost(kCGHIDEventTap, event_up)
