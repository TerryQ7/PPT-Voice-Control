"""PPT语音控制助手 - 语音识别引擎

支持两种离线引擎：
  1. VoskEngine   — 轻量流式识别（Kaldi 底层）
  2. FunASREngine — 阿里 Paraformer 非流式 + 实时能量 VAD（中英文高精度）
"""

import json
import os
import queue
import threading
from abc import ABC, abstractmethod
from typing import Callable, Optional

import numpy as np
import sounddevice as sd

from config import SAMPLE_RATE, BLOCK_SIZE


class ASREngineBase(ABC):
    """语音识别引擎抽象基类。"""

    @abstractmethod
    def start(self, on_partial: Callable[[str], None], on_final: Callable[[str], None]):
        ...

    @abstractmethod
    def stop(self):
        ...

    @abstractmethod
    def is_running(self) -> bool:
        ...


# =====================================================================
#  Vosk 引擎
# =====================================================================

class VoskEngine(ASREngineBase):
    """基于 Vosk 的离线流式语音识别引擎。"""

    def __init__(self, model_path: str, sample_rate: int = SAMPLE_RATE):
        self.model_path = model_path
        self.sample_rate = sample_rate
        self._running = False
        self._audio_queue: queue.Queue = queue.Queue()
        self._stream: Optional[sd.RawInputStream] = None
        self._thread: Optional[threading.Thread] = None
        self._on_partial: Optional[Callable] = None
        self._on_final: Optional[Callable] = None
        self._model = None
        self._recognizer = None

    def _load_model(self):
        if self._model is not None:
            return
        import vosk
        vosk.SetLogLevel(-1)
        if not os.path.isdir(self.model_path):
            raise FileNotFoundError(
                f"Vosk 模型目录不存在: {self.model_path}\n"
                "请先运行 download_model.py 下载模型。"
            )
        self._model = vosk.Model(self.model_path)

    def start(self, on_partial: Callable[[str], None], on_final: Callable[[str], None]):
        if self._running:
            return
        self._load_model()
        import vosk
        self._recognizer = vosk.KaldiRecognizer(self._model, self.sample_rate)
        self._recognizer.SetWords(False)
        self._on_partial = on_partial
        self._on_final = on_final
        self._running = True
        self._flush_queue()
        self._stream = sd.RawInputStream(
            samplerate=self.sample_rate, blocksize=BLOCK_SIZE,
            dtype="int16", channels=1, callback=self._audio_callback,
        )
        self._stream.start()
        self._thread = threading.Thread(target=self._recognition_loop, daemon=True)
        self._thread.start()

    def stop(self):
        self._running = False
        if self._stream is not None:
            self._stream.stop()
            self._stream.close()
            self._stream = None
        if self._thread is not None:
            self._thread.join(timeout=2)
            self._thread = None
        self._recognizer = None

    def is_running(self) -> bool:
        return self._running

    def _flush_queue(self):
        while not self._audio_queue.empty():
            try:
                self._audio_queue.get_nowait()
            except queue.Empty:
                break

    def _audio_callback(self, indata, frames, time_info, status):
        if self._running:
            self._audio_queue.put(bytes(indata))

    def _recognition_loop(self):
        while self._running:
            try:
                data = self._audio_queue.get(timeout=0.5)
            except queue.Empty:
                continue
            if self._recognizer is None:
                break
            if self._recognizer.AcceptWaveform(data):
                result = json.loads(self._recognizer.Result())
                text = result.get("text", "").strip()
                if text and self._on_final:
                    self._on_final(text)
            else:
                partial = json.loads(self._recognizer.PartialResult())
                text = partial.get("partial", "").strip()
                if text and self._on_partial:
                    self._on_partial(text)


# =====================================================================
#  FunASR 引擎（非流式 Paraformer + 实时能量 VAD）
# =====================================================================

class FunASREngine(ASREngineBase):
    """高精度离线引擎：非流式 Paraformer-zh + 实时能量 VAD。

    架构: 麦克风 → 100ms 小块采集 → 能量 VAD 检测说话起止 → 语音段送入 Paraformer → 文本

    模型首次加载时自动从 ModelScope 下载（约 1GB），之后完全离线运行。
    """

    CAPTURE_CHUNK = 1600        # 100ms @ 16kHz
    SILENCE_DURATION = 0.45     # 说完后静默 450ms 视为语句结束
    MIN_SPEECH_DURATION = 0.2   # 过滤极短噪音
    MAX_SPEECH_DURATION = 8.0   # 防止无限录制
    NOISE_SMOOTH = 0.92         # 噪声底估计的平滑系数
    SPEECH_THRESHOLD_RATIO = 4.0

    _FUNASR_MODEL_ID = (
        "iic/speech_seaco_paraformer_large_asr_nat-zh-cn-16k-common-vocab8404-pytorch"
    )
    _LOCAL_MODEL_DIR = os.path.expanduser(
        f"~/.cache/modelscope/hub/models/{_FUNASR_MODEL_ID}"
    )

    def __init__(self, sample_rate: int = SAMPLE_RATE):
        self.sample_rate = sample_rate
        self._running = False
        self._audio_queue: queue.Queue = queue.Queue()
        self._stream: Optional[sd.InputStream] = None
        self._thread: Optional[threading.Thread] = None
        self._on_partial: Optional[Callable] = None
        self._on_final: Optional[Callable] = None
        self._model = None

    def _load_model(self):
        if self._model is not None:
            return
        from funasr import AutoModel

        model_id = (
            self._LOCAL_MODEL_DIR
            if os.path.isdir(self._LOCAL_MODEL_DIR)
            else self._FUNASR_MODEL_ID
        )
        self._model = AutoModel(model=model_id, device="cpu", disable_update=True)

    def start(self, on_partial: Callable[[str], None], on_final: Callable[[str], None]):
        if self._running:
            return
        self._load_model()
        self._on_partial = on_partial
        self._on_final = on_final
        self._running = True

        while not self._audio_queue.empty():
            try:
                self._audio_queue.get_nowait()
            except queue.Empty:
                break

        self._stream = sd.InputStream(
            samplerate=self.sample_rate,
            blocksize=self.CAPTURE_CHUNK,
            dtype="float32",
            channels=1,
            callback=self._audio_callback,
        )
        self._stream.start()
        self._thread = threading.Thread(target=self._recognition_loop, daemon=True)
        self._thread.start()

    def stop(self):
        self._running = False
        if self._stream is not None:
            self._stream.stop()
            self._stream.close()
            self._stream = None
        if self._thread is not None:
            self._thread.join(timeout=3)
            self._thread = None

    def is_running(self) -> bool:
        return self._running

    def _audio_callback(self, indata, frames, time_info, status):
        if self._running:
            self._audio_queue.put(indata[:, 0].copy())

    def _recognition_loop(self):
        buffer: list[np.ndarray] = []
        is_speaking = False
        silence_chunks = 0
        chunk_sec = self.CAPTURE_CHUNK / self.sample_rate
        max_silence = int(self.SILENCE_DURATION / chunk_sec)
        max_speech_chunks = int(self.MAX_SPEECH_DURATION / chunk_sec)
        noise_energy = 0.01

        while self._running:
            try:
                chunk = self._audio_queue.get(timeout=0.5)
            except queue.Empty:
                if is_speaking and buffer:
                    self._recognize_segment(buffer)
                    buffer, is_speaking = [], False
                continue

            energy = float(np.sqrt(np.mean(chunk ** 2)))
            threshold = max(noise_energy * self.SPEECH_THRESHOLD_RATIO, 0.008)

            if energy > threshold:
                if not is_speaking:
                    is_speaking = True
                    if self._on_partial:
                        self._on_partial("(listening...)")
                silence_chunks = 0
                buffer.append(chunk)
            elif is_speaking:
                buffer.append(chunk)
                silence_chunks += 1
                if silence_chunks >= max_silence:
                    self._recognize_segment(buffer)
                    buffer, is_speaking, silence_chunks = [], False, 0
            else:
                noise_energy = self.NOISE_SMOOTH * noise_energy + (1 - self.NOISE_SMOOTH) * energy

            if is_speaking and len(buffer) > max_speech_chunks:
                self._recognize_segment(buffer)
                buffer, is_speaking, silence_chunks = [], False, 0

    def _recognize_segment(self, buffer: list[np.ndarray]):
        audio = np.concatenate(buffer)
        if len(audio) / self.sample_rate < self.MIN_SPEECH_DURATION:
            return
        try:
            res = self._model.generate(input=audio, batch_size_s=300)
        except Exception:
            return
        text = self._extract_text(res)
        if text and self._on_final:
            self._on_final(text)

    @staticmethod
    def _extract_text(res) -> str:
        if not res or not isinstance(res, list):
            return ""
        for item in res:
            text = item.get("text", "")
            if text:
                return text
        return ""


def get_available_devices():
    """返回可用麦克风设备列表。"""
    devices = sd.query_devices()
    return [(i, d["name"]) for i, d in enumerate(devices) if d["max_input_channels"] > 0]
