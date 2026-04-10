"""PPT语音控制助手 - Vosk 模型下载工具

独立运行: python download_model.py
也可作为模块被 main.py 调用。
"""

import os
import sys
import urllib.request
import zipfile
from typing import Callable, Optional

from config import MODEL_DIR, DEFAULT_VOSK_MODEL, VOSK_MODEL_URLS


def download_and_extract(
    url: str,
    target_dir: str,
    progress_cb: Optional[Callable[[float], None]] = None,
):
    """下载 zip 文件并解压到目标目录。"""
    os.makedirs(target_dir, exist_ok=True)
    filename = url.split("/")[-1]
    tmp_path = os.path.join(target_dir, filename)

    def _reporthook(block_num, block_size, total_size):
        if total_size > 0 and progress_cb:
            progress_cb(min(block_num * block_size / total_size * 100, 100))

    try:
        urllib.request.urlretrieve(url, tmp_path, reporthook=_reporthook)
    except Exception as e:
        if os.path.exists(tmp_path):
            os.remove(tmp_path)
        raise RuntimeError(f"下载失败: {e}") from e

    try:
        with zipfile.ZipFile(tmp_path, "r") as zf:
            zf.extractall(target_dir)
    except Exception as e:
        raise RuntimeError(f"解压失败: {e}") from e
    finally:
        if os.path.exists(tmp_path):
            os.remove(tmp_path)


def main():
    model_name = DEFAULT_VOSK_MODEL
    url = VOSK_MODEL_URLS[model_name]
    dest = os.path.join(MODEL_DIR, model_name)

    if os.path.isdir(dest):
        print(f"模型已存在: {dest}")
        ans = input("是否重新下载? (y/N): ").strip().lower()
        if ans != "y":
            return

    print(f"模型: {model_name}")
    print(f"URL : {url}")
    print(f"目标: {MODEL_DIR}\n")

    bar_len = 40

    def cli_progress(pct: float):
        filled = int(bar_len * pct / 100)
        bar = "█" * filled + "░" * (bar_len - filled)
        sys.stdout.write(f"\r  [{bar}] {pct:5.1f}%")
        sys.stdout.flush()
        if pct >= 100:
            print()

    download_and_extract(url, MODEL_DIR, progress_cb=cli_progress)
    print(f"\n✓ 模型已下载至: {dest}")


if __name__ == "__main__":
    main()
