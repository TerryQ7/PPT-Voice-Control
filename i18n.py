"""i18n string tables for the PPT Voice Control assistant.

Only two locales are shipped: Simplified Chinese (``zh``) and English (``en``).
The active locale is selected at startup via the ``--lang`` CLI flag or the
``PPT_LANG`` environment variable; see ``main.py``.
"""

from __future__ import annotations

from typing import Dict

LOCALE_ZH = "zh"
LOCALE_EN = "en"

_STRINGS: Dict[str, Dict[str, str]] = {
    LOCALE_ZH: {
        "app.title": "PPT 语音控制助手",
        "app.subtitle": "说出“下一页 / next slide / 第N页”即可控制 PPT（支持中英文）",
        "tip.title": "使用提示",
        "tip.line1": "① PPT 放映窗口必须置于最前台（焦点窗口），否则按键无法送达。",
        "tip.line2": "② 远程会议时请选择「系统音频(Loopback)」作为音频源。",
        "section.audio": "音频输入",
        "label.device": "输入设备",
        "btn.refresh": "刷新",
        "label.default_device": "系统默认",
        "label.default_device_named": "系统默认 · {name}",
        "status.idle": "未启动",
        "status.ready_funasr": "就绪 · FunASR 引擎",
        "status.ready_vosk": "就绪 · Vosk 引擎",
        "status.loading_model": "正在加载模型…",
        "status.listening": "正在监听 · {engine}",
        "status.stopped": "已停止",
        "status.download_model": "模型未下载",
        "status.downloading": "下载中 {pct:.0f}%",
        "status.download_done": "模型下载完成",
        "status.download_failed": "下载失败",
        "status.start_failed": "启动失败",
        "status.no_audio": "未检测到音频输入",
        "label.partial": "实时识别",
        "label.log": "命令执行记录",
        "placeholder.partial": "等待语音…",
        "btn.start": "开始监听",
        "btn.stop": "停止",
        "msg.engine_funasr_hint": "引擎：FunASR Paraformer（首次启动将自动下载模型，约 1 GB）",
        "msg.start_listening": "— 开始监听（{engine}） —",
        "msg.stop_listening": "— 停止监听 —",
        "msg.audio_device": "音频设备：{name}",
        "msg.download_start": "正在下载语音识别模型，请稍候…",
        "msg.download_done": "✓ 模型下载完成",
        "msg.download_failed": "✗ 下载失败：{error}",
        "msg.download_manual": "⚠ 请手动运行 python download_model.py 下载模型",
        "msg.exec_failed": "  ✗ 执行失败：{error}",
        "msg.accessibility_missing": "⚠ 未授予辅助功能权限，键盘模拟将无法工作。",
        "msg.accessibility_path": "  请前往：系统设置 → 隐私与安全性 → 辅助功能",
        "msg.no_audio_log": "⚠ 连续多秒未检测到有效音频信号，当前设备：{name}",
        "partial.listening": "（正在聆听…）",
        "partial.ready": "[就绪] 噪声基线 = {value:.4f}",
        "partial.no_result": "（无识别结果，{duration:.1f}s）",
        "partial.recog_error": "[识别错误] {error}",
        "dialog.accessibility_title": "需要辅助功能权限",
        "dialog.accessibility_body": (
            "未检测到辅助功能权限，PPT 翻页将无法工作。\n\n"
            "请前往：系统设置 → 隐私与安全性 → 辅助功能\n\n"
            "添加并启用当前终端或打包的 .app。"
        ),
        "dialog.model_missing_title": "模型未找到",
        "dialog.model_missing_body": "Vosk 模型不存在，是否立即下载？\n（{model}）",
        "dialog.model_vosk_error_title": "错误",
        "dialog.model_vosk_error_body": "Vosk 模型未下载，无法启动。",
        "dialog.start_failed_title": "启动失败",
        "dialog.no_audio_title": "未检测到音频输入",
        "dialog.no_audio_body": (
            "已持续多秒未检测到有效音频信号，当前麦克风可能未正常工作。\n\n"
            "当前设备：{name}\n\n"
            "请检查：\n"
            "1. 麦克风是否已连接并开启\n"
            "2. 系统音频输入设置是否正确\n"
            "3. 是否选择了正确的输入设备\n\n"
            "其他可用设备：\n{devices}\n\n"
            "点击「停止」后可在列表中切换，然后重新开始监听。"
        ),
        "dialog.no_audio_none": "  （无其他可用设备）",
        "log.source_partial": "部分",
        "log.source_final": "最终",
        "cmd.next": "下一页",
        "cmd.prev": "上一页",
        "cmd.first": "第一页",
        "cmd.last": "最后一页",
        "cmd.goto": "跳转到第 {page} 页",
        "commands.title": "语音指令示例",
        "commands.line1": "• 下一页 · 上一页 · 第一页 · 最后一页",
        "commands.line2": "• 第 N 页（支持中文数字：第二十三页）",
        "commands.line3": "• next slide · previous slide · go to page 10",
    },
    LOCALE_EN: {
        "app.title": "PPT Voice Control",
        "app.subtitle": 'Say "next slide", "previous slide" or "go to page N" to control PowerPoint hands-free.',
        "tip.title": "Tips",
        "tip.line1": "① PowerPoint must be the foreground (focused) window for keystrokes to reach it.",
        "tip.line2": "② For remote meetings, select a system audio (Loopback) device as the audio source.",
        "section.audio": "Audio Input",
        "label.device": "Input device",
        "btn.refresh": "Refresh",
        "label.default_device": "System default",
        "label.default_device_named": "System default · {name}",
        "status.idle": "Not started",
        "status.ready_funasr": "Ready · FunASR engine",
        "status.ready_vosk": "Ready · Vosk engine",
        "status.loading_model": "Loading model…",
        "status.listening": "Listening · {engine}",
        "status.stopped": "Stopped",
        "status.download_model": "Model not downloaded",
        "status.downloading": "Downloading {pct:.0f}%",
        "status.download_done": "Model download complete",
        "status.download_failed": "Download failed",
        "status.start_failed": "Start failed",
        "status.no_audio": "No audio input detected",
        "label.partial": "Live transcript",
        "label.log": "Command history",
        "placeholder.partial": "Waiting for speech…",
        "btn.start": "Start listening",
        "btn.stop": "Stop",
        "msg.engine_funasr_hint": "Engine: FunASR Paraformer (the model is downloaded automatically on first launch, about 1 GB).",
        "msg.start_listening": "— Listening started ({engine}) —",
        "msg.stop_listening": "— Listening stopped —",
        "msg.audio_device": "Audio device: {name}",
        "msg.download_start": "Downloading the speech model, please wait…",
        "msg.download_done": "✓ Model download complete",
        "msg.download_failed": "✗ Download failed: {error}",
        "msg.download_manual": "⚠ Run `python download_model.py` manually to download the model.",
        "msg.exec_failed": "  ✗ Execution failed: {error}",
        "msg.accessibility_missing": "⚠ Accessibility permission is missing; keystroke simulation will not work.",
        "msg.accessibility_path": "  Open System Settings → Privacy & Security → Accessibility",
        "msg.no_audio_log": "⚠ No audio detected for several seconds on device: {name}",
        "partial.listening": "(listening…)",
        "partial.ready": "[Ready] noise baseline = {value:.4f}",
        "partial.no_result": "(no recognition result, {duration:.1f}s)",
        "partial.recog_error": "[Recognition error] {error}",
        "dialog.accessibility_title": "Accessibility permission required",
        "dialog.accessibility_body": (
            "Accessibility permission is not granted. Slide navigation will not work.\n\n"
            "Open: System Settings → Privacy & Security → Accessibility\n\n"
            "Then add and enable your terminal application or the packaged .app."
        ),
        "dialog.model_missing_title": "Model not found",
        "dialog.model_missing_body": "The Vosk model is not available. Download it now?\n({model})",
        "dialog.model_vosk_error_title": "Error",
        "dialog.model_vosk_error_body": "The Vosk model is not downloaded; cannot start.",
        "dialog.start_failed_title": "Start failed",
        "dialog.no_audio_title": "No audio input detected",
        "dialog.no_audio_body": (
            "No audio signal has been detected for several seconds. The selected microphone may not be working.\n\n"
            "Current device: {name}\n\n"
            "Please check:\n"
            "1. The microphone is connected and enabled\n"
            "2. The system input settings are correct\n"
            "3. The right input device is selected\n\n"
            "Other available devices:\n{devices}\n\n"
            "Click Stop, pick another device, then start listening again."
        ),
        "dialog.no_audio_none": "  (no other input devices)",
        "log.source_partial": "partial",
        "log.source_final": "final",
        "cmd.next": "Next slide",
        "cmd.prev": "Previous slide",
        "cmd.first": "First slide",
        "cmd.last": "Last slide",
        "cmd.goto": "Go to slide {page}",
        "commands.title": "Voice command examples",
        "commands.line1": "• Next slide · Previous slide · First slide · Last slide",
        "commands.line2": "• Go to page N (English or Chinese numerals)",
        "commands.line3": "• 下一页 · 上一页 · 第 N 页（支持中文数字）",
    },
}


class Translator:
    """Lightweight lookup helper with graceful fallbacks."""

    def __init__(self, locale: str = LOCALE_ZH):
        self.locale = locale if locale in _STRINGS else LOCALE_ZH

    def t(self, key: str, **kwargs) -> str:
        value = _STRINGS.get(self.locale, {}).get(key)
        if value is None:
            value = _STRINGS[LOCALE_ZH].get(key, key)
        if kwargs:
            try:
                return value.format(**kwargs)
            except (KeyError, IndexError):
                return value
        return value

    def __call__(self, key: str, **kwargs) -> str:
        return self.t(key, **kwargs)


def normalize_locale(raw: str | None) -> str:
    if not raw:
        return LOCALE_ZH
    raw = raw.strip().lower()
    if raw in ("en", "en-us", "en_us", "english"):
        return LOCALE_EN
    if raw in ("zh", "zh-cn", "zh_cn", "chinese", "zh-hans"):
        return LOCALE_ZH
    return LOCALE_ZH
