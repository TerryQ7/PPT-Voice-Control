"""PPT语音控制助手 - 语音识别引擎

支持两种离线引擎：
  1. VoskEngine   — 轻量流式识别（Kaldi 底层）
  2. FunASREngine — 阿里 Paraformer 非流式 + 实时能量 VAD（中英文高精度）
"""

import json
import logging
import os
import platform
import sys
import queue
import threading
import time
from abc import ABC, abstractmethod
from typing import Callable, Optional

import numpy as np
import sounddevice as sd

from config import SAMPLE_RATE, BLOCK_SIZE

logger = logging.getLogger("asr_engine")
IS_WINDOWS = platform.system() == "Windows"


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
    MIN_SPEECH_DURATION = 0.3   # 过滤极短噪音
    FORCED_TIMEOUT = 2.5        # 兜底：说话后 2.5s 必定送识别
    MAX_SPEECH_DURATION = 6.0   # 绝对上限
    NOISE_SMOOTH = 0.92
    CALIBRATION_SECONDS = 0.6   # 启动后先采集环境噪音

    # --- 平台自适应参数 ---
    if IS_WINDOWS:
        SPEECH_START_RATIO = 3.0        # Windows 麦克风增益高，降低起始阈值
        START_THRESHOLD_CAP = 0.04      # 起始阈值上限，防止永远进不了说话态
        SPEECH_END_RATIO = 0.40         # 放宽结束阈值，让段落更容易闭合
        SILENCE_DURATION = 0.7          # Windows 上稍长的静默窗口
    else:
        SPEECH_START_RATIO = 5.0
        START_THRESHOLD_CAP = float("inf")
        SPEECH_END_RATIO = 0.20
        SILENCE_DURATION = 0.5

    # "speech_seaco_..." 在新版本 ModelScope 上可能不可用，使用稳定可下载的 ID。
    _FUNASR_MODEL_ID = (
        "iic/speech_paraformer-large_asr_nat-zh-cn-16k-common-vocab8404-pytorch"
    )
    _BUNDLED_MODEL_NAME = "funasr_model"

    def __init__(self, sample_rate: int = SAMPLE_RATE):
        self.sample_rate = sample_rate
        self._running = False
        self._audio_queue: queue.Queue = queue.Queue()
        self._stream: Optional[sd.InputStream] = None
        self._thread: Optional[threading.Thread] = None
        self._on_partial: Optional[Callable] = None
        self._on_final: Optional[Callable] = None
        self._model = None

    @staticmethod
    def _is_valid_model_dir(path: str) -> bool:
        """判断目录是否为可用 FunASR 模型目录（避免半下载缓存导致启动失败）。"""
        if not os.path.isdir(path):
            return False
        required = ("configuration.json", "config.yaml", "model.pt")
        return all(os.path.isfile(os.path.join(path, name)) for name in required)

    def _find_model(self) -> str:
        """按优先级查找模型路径: 打包内置 → ModelScope 缓存 → 模型ID(自动下载)。"""
        if getattr(sys, "frozen", False):
            bundled = os.path.join(sys._MEIPASS, self._BUNDLED_MODEL_NAME)
            if self._is_valid_model_dir(bundled):
                return bundled

        local = os.path.join(os.path.dirname(os.path.abspath(__file__)), self._BUNDLED_MODEL_NAME)
        if self._is_valid_model_dir(local):
            return local

        cache_root = os.path.expanduser("~/.cache/modelscope/hub/models")
        cache_candidates = [
            os.path.join(
                cache_root,
                "iic",
                "speech_paraformer-large_asr_nat-zh-cn-16k-common-vocab8404-pytorch",
            ),
            os.path.join(
                cache_root,
                "damo",
                "speech_paraformer-large_asr_nat-zh-cn-16k-common-vocab8404-pytorch",
            ),
        ]
        for cache in cache_candidates:
            if self._is_valid_model_dir(cache):
                return cache

        return self._FUNASR_MODEL_ID

    def _load_model(self):
        if self._model is not None:
            return
        import torch
        torch.set_num_threads(2)
        torch.set_grad_enabled(False)
        from funasr import AutoModel

        model_path = self._find_model()
        logger.info("Loading model from: %s", model_path)
        self._model = AutoModel(model=model_path, device="cpu", disable_update=True)

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
        speech_peak = 0.0
        chunk_sec = self.CAPTURE_CHUNK / self.sample_rate
        max_silence = int(self.SILENCE_DURATION / chunk_sec)
        forced_chunks = int(self.FORCED_TIMEOUT / chunk_sec)
        max_speech_chunks = int(self.MAX_SPEECH_DURATION / chunk_sec)
        calibration_chunks = int(self.CALIBRATION_SECONDS / chunk_sec)
        debug_counter = 0

        noise_energy = self._calibrate_noise(calibration_chunks)
        logger.info("[VAD] calibrated noise=%.4f, platform=%s", noise_energy,
                     "Windows" if IS_WINDOWS else "macOS/Linux")
        if self._on_partial:
            self._on_partial(f"[就绪] 噪声基线={noise_energy:.4f}")

        while self._running:
            try:
                chunk = self._audio_queue.get(timeout=0.5)
            except queue.Empty:
                if is_speaking and buffer:
                    logger.info("[VAD] queue timeout while speaking, forcing recognize (%d chunks)", len(buffer))
                    self._recognize_segment(buffer)
                    buffer, is_speaking, speech_peak = [], False, 0.0
                continue

            energy = float(np.sqrt(np.mean(chunk ** 2)))
            start_threshold = min(
                max(noise_energy * self.SPEECH_START_RATIO, 0.008),
                self.START_THRESHOLD_CAP,
            )

            # 每秒打印一次调试信息
            debug_counter += 1
            if debug_counter % 10 == 0:
                logger.debug(
                    "[VAD] energy=%.4f noise=%.4f start_th=%.4f speaking=%s buf=%d silence=%d",
                    energy, noise_energy, start_threshold, is_speaking, len(buffer), silence_chunks,
                )

            if not is_speaking:
                if energy > start_threshold:
                    is_speaking = True
                    speech_peak = energy
                    silence_chunks = 0
                    buffer.append(chunk)
                    logger.info("[VAD] speech START energy=%.4f threshold=%.4f", energy, start_threshold)
                    if self._on_partial:
                        self._on_partial("(listening...)")
                else:
                    noise_energy = (self.NOISE_SMOOTH * noise_energy
                                    + (1 - self.NOISE_SMOOTH) * energy)
            else:
                buffer.append(chunk)
                speech_peak = max(speech_peak, energy)

                end_threshold = max(speech_peak * self.SPEECH_END_RATIO,
                                    noise_energy * 1.5)

                if energy < end_threshold:
                    silence_chunks += 1
                    if silence_chunks >= max_silence:
                        dur = len(buffer) * chunk_sec
                        logger.info("[VAD] speech END (silence) dur=%.1fs peak=%.4f", dur, speech_peak)
                        self._recognize_segment(buffer)
                        buffer, is_speaking, silence_chunks, speech_peak = [], False, 0, 0.0
                        continue
                else:
                    silence_chunks = 0

                if len(buffer) >= forced_chunks:
                    dur = len(buffer) * chunk_sec
                    logger.info("[VAD] speech END (forced timeout) dur=%.1fs", dur)
                    self._recognize_segment(buffer)
                    buffer, is_speaking, silence_chunks, speech_peak = [], False, 0, 0.0
                    continue

                if len(buffer) > max_speech_chunks:
                    logger.info("[VAD] speech END (max duration)")
                    self._recognize_segment(buffer)
                    buffer, is_speaking, silence_chunks, speech_peak = [], False, 0, 0.0

    def _calibrate_noise(self, num_chunks: int) -> float:
        """启动时采集环境噪音，返回噪声能量基线。"""
        energies = []
        for _ in range(num_chunks):
            if not self._running:
                break
            try:
                chunk = self._audio_queue.get(timeout=1.0)
                energies.append(float(np.sqrt(np.mean(chunk ** 2))))
            except queue.Empty:
                continue
        if energies:
            return sum(energies) / len(energies)
        return 0.01

    def _recognize_segment(self, buffer: list[np.ndarray]):
        import torch
        audio = np.concatenate(buffer)
        duration = len(audio) / self.sample_rate

        if duration < self.MIN_SPEECH_DURATION:
            logger.debug("[ASR] skipped short segment: %.2fs", duration)
            return

        # 确保音频数据格式正确（Windows 驱动兼容性）
        audio = np.ascontiguousarray(audio, dtype=np.float32)
        audio = np.clip(audio, -1.0, 1.0)

        try:
            with torch.no_grad():
                res = self._model.generate(input=audio, batch_size_s=300)
        except Exception as e:
            logger.error("[ASR] generate() failed: %s", e)
            if self._on_partial:
                self._on_partial(f"[识别错误] {e}")
            return

        text = self._extract_text(res)
        logger.info("[ASR] recognized (%.1fs): '%s'", duration, text)

        if text and self._on_final:
            self._on_final(text)
        elif not text and self._on_partial:
            self._on_partial(f"(无识别结果, {duration:.1f}s)")

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
