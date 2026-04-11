# PPT Voice Control Assistant / PPT 语音控制助手

Real-time voice-controlled slide navigation for PowerPoint and Keynote.
Listens to your microphone, recognizes Chinese & English commands, and automatically switches slides — no clicking needed.

**Default engine: [FunASR Paraformer](https://github.com/modelscope/FunASR)** — high-accuracy offline speech recognition from Alibaba, supporting Chinese–English bilingual input.

## Features

- **Fully offline** — after initial model download, no internet required
- **Chinese & English** voice commands with Chinese numeral support (e.g. "第二十三页")
- **Platform-adaptive VAD** — energy-based voice activity detection with auto-calibrated noise baseline, tuned separately for macOS and Windows
- **Dual ASR engine** — FunASR Paraformer (high accuracy) or Vosk (lightweight)
- **Cross-platform** — macOS (Quartz CGEvent) and Windows (pynput) keyboard simulation
- **One-click packaging** — PyInstaller scripts for macOS & Windows, plus Inno Setup installer for offline Windows distribution

## Supported Commands

| Voice Command | Action |
|---|---|
| "下一页" / "next page" / "next slide" | Next slide |
| "上一页" / "previous page" / "go back" | Previous slide |
| "第N页" / "回到第N页" / "go to page N" | Jump to slide N |
| "第一页" / "first page" | First slide |
| "最后一页" / "last slide" | Last slide |

> Chinese numerals are fully supported — e.g. "第二十三页" (page 23), "翻到第一百页" (page 100).

## Requirements

- Python 3.10+
- macOS or Windows
- Microphone

## Quick Start

### 1. Create Virtual Environment

```bash
# macOS
python3 -m venv venv
source venv/bin/activate

# Windows
python -m venv venv
venv\Scripts\activate
```

```bash
pip install -r requirements.txt
```

> **macOS note:** If you see `No module named '_tkinter'`:
> ```bash
> brew install python-tk@3.13   # match your Python version
> ```

### 2. Run

```bash
python main.py
```

Open your PPT in **slideshow mode**, click **Start Listening**, and speak your commands.

The FunASR model (~1 GB) will be downloaded automatically on first launch via ModelScope. After that, the system runs fully offline.

### 3. macOS Permissions

Two permissions are required:

1. **Microphone** — system will prompt automatically on first launch.
2. **Accessibility** — required for keyboard simulation. Go to:
   `System Settings → Privacy & Security → Accessibility`
   and add your terminal app (Terminal / iTerm / Cursor) or the packaged `.app`.

## Architecture

```
Microphone
    ↓
Audio Capture (sounddevice, 100ms chunks)
    ↓
Voice Activity Detection (energy-based, platform-adaptive)
    ├── Auto noise calibration on startup (~0.6s)
    ├── Relative-peak energy for end-of-speech detection
    └── Forced 2.5s timeout as safety net
    ↓
Speech Recognition (FunASR Paraformer / Vosk)
    ↓
Command Parser (regex + keyword matching, Chinese numeral support, debounce)
    ↓
PPT Controller
    ├── macOS: Quartz CGEvent (low-level hardware key events)
    └── Windows: pynput keyboard simulation
```

## Project Structure

```
PPT_Project/
├── main.py                        # GUI entry point (tkinter)
├── asr_engine.py                  # ASR engines: FunASREngine + VoskEngine
├── command_parser.py              # Voice command parser (Chinese/English, debounce)
├── ppt_controller.py              # Keyboard simulation (Quartz CGEvent / pynput)
├── config.py                      # Global configuration (engine, keywords, params)
├── download_model.py              # Vosk model download utility
├── requirements.txt               # Python dependencies
├── build_app.sh                   # macOS packaging (PyInstaller)
├── build_app_windows.bat          # Windows packaging (PyInstaller)
├── build_app_windows_offline.bat  # Windows offline packaging (bundles model)
├── build_installer_offline.iss    # Inno Setup installer script
└── .github/workflows/
    └── build-windows.yml          # GitHub Actions CI for Windows build
```

## Configuration

Edit `config.py` to customize:

| Option | Description | Default |
|---|---|---|
| `ASR_ENGINE` | Recognition engine (`"funasr"` or `"vosk"`) | `"funasr"` |
| `DEBOUNCE_SECONDS` | Min interval between duplicate commands | `2.0` s |
| `SAMPLE_RATE` | Audio sample rate | `16000` Hz |
| `NEXT_KEYWORDS`, etc. | Trigger keywords for each command | See file |

### Switching to Vosk

For a lighter setup (no PyTorch needed, ~42 MB model):

1. Set `ASR_ENGINE = "vosk"` in `config.py`
2. Run `python download_model.py` to download the Vosk model
3. Run `python main.py`

## Packaging

### macOS

```bash
chmod +x build_app.sh && ./build_app.sh
```

Output: `dist/PPT语音控制助手/`

### Windows (online)

```bash
build_app_windows.bat
```

Output: `dist/PPT语音控制助手/`

### Windows (offline, bundles model)

Builds a standalone package that includes the FunASR model — target machines need no Python or internet.

```bash
build_app_windows_offline.bat
```

Output: `release/PPT-Voice-Control-Offline/` + `.zip`

To create a Windows installer (requires [Inno Setup](https://jrsoftware.org/isinfo.php)):

1. Copy the build output into `installer_staging/`
2. Run `build_installer_offline.iss` with Inno Setup Compiler

### GitHub Actions

A CI workflow (`.github/workflows/build-windows.yml`) automatically builds the Windows executable on push to `main`. The artifact is uploaded and available for 30 days.

## License

Internal use — United Imaging.
