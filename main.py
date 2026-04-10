"""PPT语音控制助手 - 主程序入口

实时监听麦克风，识别中英文翻页命令，自动控制PPT翻页。
"""

import datetime
import logging
import os
import platform
import sys
import threading
import tkinter as tk
from tkinter import ttk, scrolledtext, messagebox

class _NullStream:
    """兼容无控制台窗口场景的标准输出/错误占位流。"""

    def write(self, _data):
        return 0

    def flush(self):
        return None

    def isatty(self):
        return False


def _ensure_std_streams():
    """在 windowed 打包下，避免第三方库向 None 的 stderr/stdout 写入。"""
    for name in ("stdout", "stderr", "__stdout__", "__stderr__"):
        if getattr(sys, name, None) is None:
            setattr(sys, name, _NullStream())


_ensure_std_streams()

from config import ASR_ENGINE, VOSK_MODEL_PATH, MODEL_DIR, DEFAULT_VOSK_MODEL, VOSK_MODEL_URLS
from command_parser import CommandParser, Command
from ppt_controller import PPTController, check_accessibility_permission
from asr_engine import VoskEngine, FunASREngine, ASREngineBase

logging.basicConfig(
    level=logging.DEBUG,
    format="%(asctime)s [%(name)s] %(levelname)s: %(message)s",
    datefmt="%H:%M:%S",
)


class PPTVoiceApp:
    """PPT 语音控制助手 GUI 应用。"""

    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("PPT 语音控制助手")
        self.root.geometry("580x520")
        self.root.resizable(False, False)

        self.parser = CommandParser()
        self.controller = PPTController()
        self.engine: ASREngineBase | None = None

        self._build_ui()
        self._check_accessibility()
        self._check_model()

    # ==================== UI ====================

    def _build_ui(self):
        self.root.configure(bg="#f0f0f0")

        style = ttk.Style()
        style.theme_use("clam")
        style.configure("Title.TLabel", font=("Helvetica", 18, "bold"), background="#f0f0f0")
        style.configure("Status.TLabel", font=("Helvetica", 13), background="#f0f0f0")
        style.configure("Start.TButton", font=("Helvetica", 13), padding=10)
        style.configure("Stop.TButton", font=("Helvetica", 13), padding=10)
        style.configure("Info.TLabel", font=("Helvetica", 11), background="#f0f0f0", foreground="#666")

        main_frame = ttk.Frame(self.root, padding=20)
        main_frame.pack(fill=tk.BOTH, expand=True)

        ttk.Label(main_frame, text="PPT 语音控制助手", style="Title.TLabel").pack(pady=(0, 5))
        ttk.Label(
            main_frame,
            text='说出 "下一页 / 上一页 / 第N页" 等指令即可自动控制PPT',
            style="Info.TLabel",
        ).pack(pady=(0, 15))

        status_frame = ttk.Frame(main_frame)
        status_frame.pack(fill=tk.X, pady=(0, 10))
        self.status_canvas = tk.Canvas(
            status_frame, width=16, height=16, bg="#f0f0f0", highlightthickness=0
        )
        self.status_canvas.pack(side=tk.LEFT, padx=(0, 8))
        self.status_dot = self.status_canvas.create_oval(2, 2, 14, 14, fill="#999", outline="")
        self.status_label = ttk.Label(status_frame, text="未启动", style="Status.TLabel")
        self.status_label.pack(side=tk.LEFT)

        ttk.Label(main_frame, text="实时识别:", style="Info.TLabel").pack(anchor=tk.W, pady=(5, 2))
        self.partial_var = tk.StringVar(value="...")
        self.partial_label = ttk.Label(
            main_frame, textvariable=self.partial_var, font=("Helvetica", 12), foreground="#aaa",
            background="#fff", relief="sunken", padding=6, wraplength=520
        )
        self.partial_label.pack(fill=tk.X, pady=(0, 10))

        ttk.Label(main_frame, text="命令执行记录:", style="Info.TLabel").pack(anchor=tk.W, pady=(0, 2))
        self.log_text = scrolledtext.ScrolledText(
            main_frame, height=10, font=("Courier", 11), state=tk.DISABLED, wrap=tk.WORD
        )
        self.log_text.pack(fill=tk.BOTH, expand=True, pady=(0, 15))

        btn_frame = ttk.Frame(main_frame)
        btn_frame.pack(fill=tk.X)
        self.start_btn = ttk.Button(
            btn_frame, text="▶  开始监听", style="Start.TButton", command=self._on_start
        )
        self.start_btn.pack(side=tk.LEFT, expand=True, fill=tk.X, padx=(0, 5))
        self.stop_btn = ttk.Button(
            btn_frame, text="■  停止", style="Stop.TButton", command=self._on_stop, state=tk.DISABLED
        )
        self.stop_btn.pack(side=tk.LEFT, expand=True, fill=tk.X, padx=(5, 0))

        self.root.protocol("WM_DELETE_WINDOW", self._on_close)

    # ==================== 权限检查 ====================

    def _check_accessibility(self):
        if platform.system() != "Darwin":
            return
        if not check_accessibility_permission():
            self._log("⚠ 未授予辅助功能权限，键盘模拟将无法工作。")
            self._log("  请前往: 系统设置 → 隐私与安全性 → 辅助功能")
            messagebox.showwarning(
                "需要辅助功能权限",
                "未检测到辅助功能权限，PPT翻页将无法工作。\n\n"
                "请前往:\n系统设置 → 隐私与安全性 → 辅助功能\n\n"
                "添加并启用当前终端应用。",
            )

    # ==================== 模型检查 ====================

    def _check_model(self):
        if ASR_ENGINE == "funasr":
            self._set_status("就绪 (FunASR)，点击「开始监听」", "#4CAF50")
            self._log("引擎: FunASR Paraformer (首次启动需下载模型，约1GB)")
            return

        if os.path.isdir(VOSK_MODEL_PATH):
            self._set_status("就绪 (Vosk)，点击「开始监听」", "#4CAF50")
            return

        self._set_status("模型未下载", "#FF9800")
        if messagebox.askyesno("模型未找到", f"Vosk 模型不存在，是否立即下载？\n({DEFAULT_VOSK_MODEL})"):
            self._download_model_async()
        else:
            self._log("⚠ 请手动运行 python download_model.py 下载模型")

    def _download_model_async(self):
        self.start_btn.config(state=tk.DISABLED)
        self._set_status("正在下载模型...", "#2196F3")
        self._log("开始下载语音识别模型，请稍候...")

        def _do_download():
            try:
                from download_model import download_and_extract
                download_and_extract(
                    VOSK_MODEL_URLS[DEFAULT_VOSK_MODEL], MODEL_DIR,
                    progress_cb=lambda pct: self.root.after(
                        0, lambda: self._set_status(f"下载中 {pct:.0f}%...", "#2196F3")
                    ),
                )
                self.root.after(0, self._on_download_done, True)
            except Exception as e:
                self.root.after(0, self._on_download_done, False, str(e))

        threading.Thread(target=_do_download, daemon=True).start()

    def _on_download_done(self, success: bool, error: str = ""):
        if success:
            self._set_status("模型下载完成，可以开始", "#4CAF50")
            self._log("✓ 模型下载完成！")
        else:
            self._set_status("下载失败", "#F44336")
            self._log(f"✗ 下载失败: {error}")
        self.start_btn.config(state=tk.NORMAL)

    # ==================== 开始 / 停止 ====================

    def _on_start(self):
        if ASR_ENGINE == "vosk" and not os.path.isdir(VOSK_MODEL_PATH):
            messagebox.showerror("错误", "Vosk 模型未下载，无法启动。")
            return

        self.start_btn.config(state=tk.DISABLED)
        self._set_status("正在加载模型...", "#2196F3")
        self._log("正在加载语音识别模型，请稍候...")
        self.root.update()

        def _do_start():
            try:
                engine: ASREngineBase
                if ASR_ENGINE == "funasr":
                    engine = FunASREngine()
                else:
                    engine = VoskEngine(model_path=VOSK_MODEL_PATH)
                engine.start(on_partial=self._on_partial, on_final=self._on_final)
                self.root.after(0, self._on_engine_started, engine)
            except Exception as e:
                self.root.after(0, self._on_engine_failed, str(e))

        threading.Thread(target=_do_start, daemon=True).start()

    def _on_engine_started(self, engine: ASREngineBase):
        self.engine = engine
        self.parser.reset()
        name = "FunASR" if ASR_ENGINE == "funasr" else "Vosk"
        self._set_status(f"正在监听... ({name})", "#4CAF50")
        self._log(f"--- 开始监听 ({name}) ---")
        self.stop_btn.config(state=tk.NORMAL)

    def _on_engine_failed(self, error: str):
        self._set_status("启动失败", "#F44336")
        self._log(f"✗ 启动失败: {error}")
        self.start_btn.config(state=tk.NORMAL)
        messagebox.showerror("启动失败", error)

    def _on_stop(self):
        if self.engine and self.engine.is_running():
            self.engine.stop()
        self._set_status("已停止", "#999")
        self._log("--- 停止监听 ---")
        self.start_btn.config(state=tk.NORMAL)
        self.stop_btn.config(state=tk.DISABLED)
        self.partial_var.set("...")

    def _on_close(self):
        if self.engine and self.engine.is_running():
            self.engine.stop()
        self.root.destroy()
        os._exit(0)

    # ==================== 识别回调 ====================

    def _on_partial(self, text: str):
        self.root.after(0, self._handle_partial, text)

    def _on_final(self, text: str):
        self.root.after(0, self._handle_final, text)

    def _handle_partial(self, text: str):
        self.partial_var.set(text)
        self.partial_label.configure(foreground="#aaa")
        cmd = self.parser.parse(text)
        if cmd:
            self._execute_command(cmd, text, is_partial=True)

    def _handle_final(self, text: str):
        self.partial_var.set(text)
        self.partial_label.configure(foreground="#333")
        cmd = self.parser.parse(text)
        if cmd:
            self._execute_command(cmd, text, is_partial=False)

    def _execute_command(self, cmd: Command, raw_text: str, is_partial: bool):
        now = datetime.datetime.now().strftime("%H:%M:%S")
        source = "部分" if is_partial else "最终"
        self._log(f'[{now}] ({source}) "{raw_text}" → {cmd}')
        try:
            self.controller.execute(cmd)
        except Exception as e:
            self._log(f"  ✗ 执行失败: {e}")

    # ==================== 工具方法 ====================

    def _set_status(self, text: str, color: str):
        self.status_label.config(text=text)
        self.status_canvas.itemconfig(self.status_dot, fill=color)

    def _log(self, msg: str):
        self.log_text.config(state=tk.NORMAL)
        self.log_text.insert(tk.END, msg + "\n")
        self.log_text.see(tk.END)
        self.log_text.config(state=tk.DISABLED)


def main():
    root = tk.Tk()
    PPTVoiceApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
