# PPT Voice Control Assistant / PPT 语音控制助手

Real-time voice-controlled slide navigation for PowerPoint and Keynote.
Listens to your microphone, recognizes Chinese and English commands, and automatically switches slides — no clicking needed.

**Default engine: [FunASR Paraformer](https://github.com/modelscope/FunASR)** — high-accuracy offline speech recognition from Alibaba, supporting Chinese–English bilingual input.

Both a **Simplified Chinese** and an **English** edition are shipped. The two share the same engine and code path; only the UI strings differ.

## Downloads

Pre-built Windows binaries for both languages are produced by GitHub Actions and attached to each tagged release. See the [Releases page](../../releases).

| Edition | Asset (Windows, offline, model included) |
|---|---|
| English (en-US) | `PPT-Voice-Control-en-US-Windows.zip` |
| Simplified Chinese (zh-CN) | `PPT-Voice-Control-zh-CN-Windows.zip` |

Each zip is a self-contained bundle — the target machine needs neither Python nor an internet connection.

## Features

- **Fully offline** — after the initial model download, no internet connection is required
- **Bilingual UI** — Simplified Chinese (`--lang zh`, default) and English (`--lang en`)
- **Bilingual speech** — recognizes Chinese & English commands, including Chinese numerals ("第二十三页")
- **Microphone selection** — choose from available audio input devices in the UI; supports system default or a specific mic
- **System audio capture (Loopback)** — for remote-meeting scenarios, capture speaker output instead of mic input (Windows WASAPI Loopback; on macOS install a virtual audio device such as BlackHole)
- **Smart context filtering** — distinguishes commands from descriptive speech (e.g. "像第三页那样" will NOT trigger a jump, but "跳到第三页" will)
- **Platform-adaptive VAD** — energy-based voice activity detection with auto-calibrated noise baseline, tuned separately for macOS and Windows
- **Dual ASR engine** — FunASR Paraformer (high accuracy) or Vosk (lightweight)
- **Cross-platform** — macOS (Quartz CGEvent) and Windows (pynput) keyboard simulation
- **Tesla-inspired UI** — flat white canvas, Carbon Dark text, Electric Blue accent, zero shadows, generous whitespace (see [`DESIGN.md`](./DESIGN.md))
- **One-click packaging** — PyInstaller scripts for macOS & Windows, plus Inno Setup installer for offline Windows distribution

## Supported Commands

| Voice Command | Action |
|---|---|
| "下一页" / "next page" / "next slide" / "forward" | Next slide |
| "上一页" / "previous page" / "previous one" / "go back" | Previous slide |
| "第N页" / "回到第N页" / "go to page N" / "first page" / "twenty-third page" / "one hundred and first page" | Jump to slide N |
| "第一页" / "first page" / "start over" | First slide |
| "最后一页" / "last slide" / "go to end" | Last slide |

> Chinese numerals are fully supported — e.g. "第二十三页" (page 23), "翻到第一百页" (page 100).
>
> English ordinal/cardinal page phrases are also supported — e.g. "first page", "twenty third page", "go to one hundred and first page", "page fifty", "slide ninety ninth".
>
> **Context-aware filtering**: The system distinguishes commands from descriptive speech. Saying "像第三页描述的那样" or "as shown on page 3" will NOT trigger a page jump. Only explicit commands like "第三页", "跳到第三页", or "go to page 3" will navigate.

### Important Notes

- **PowerPoint must be the foreground window** — the application sends keyboard events to the active window. If PPT is not in focus, commands will not reach it.
- **Remote meetings**: When presenting remotely (Zoom/Teams/etc.), the remote audience's voice comes through speakers, not the mic. Select a **Loopback** audio device to capture system audio output. On Windows, WASAPI loopback devices appear automatically. On macOS, install a virtual audio device like [BlackHole](https://github.com/ExistentialAudio/BlackHole).

## Requirements

- Python 3.10+
- macOS or Windows
- Microphone

## Quick Start

### 1. Create a virtual environment

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
# Simplified Chinese (default)
python main.py

# English
python main.py --lang en
# equivalent: python main_en.py
# equivalent: PPT_LANG=en python main.py
```

Open your PPT in **slideshow mode**, click **Start Listening** / **开始监听**, and speak your commands.

The FunASR model (~1 GB) is downloaded automatically on first launch via ModelScope. After that, the system runs fully offline.

### 3. macOS Permissions

Two permissions are required:

1. **Microphone** — the system will prompt automatically on first launch.
2. **Accessibility** — required for keyboard simulation. Go to:
   `System Settings → Privacy & Security → Accessibility`
   and add your terminal app (Terminal / iTerm / Cursor) or the packaged `.app`.

## Design Language

The UI follows [`DESIGN.md`](./DESIGN.md), a Tesla-inspired minimal system:

- Single Electric Blue (`#3E6AE1`) accent used only for the primary CTA
- Pure white canvas with Light Ash (`#F4F4F4`) as the sole alternate surface
- Carbon Dark / Graphite / Pewter three-tier type hierarchy
- No shadows, no gradients, no borders — depth is spacing and contrast
- SF Pro (macOS) / Segoe UI Variable (Windows) as Universal-Sans stand-ins

Both editions render an identical layout; only the text is localised.

## Architecture

```
Audio Source (user-selectable)
    ├── Microphone (local presentation)
    └── System Audio / Loopback (remote meeting)
    ↓
Audio Capture (sounddevice, 100ms chunks, configurable device)
    ↓
Voice Activity Detection (energy-based, platform-adaptive)
    ├── Auto noise calibration on startup (~0.6s)
    ├── Relative-peak energy for end-of-speech detection
    └── Forced 2.5s timeout as safety net
    ↓
Speech Recognition (FunASR Paraformer / Vosk)
    ↓
Command Parser
    ├── Keyword matching (Chinese + English, expanded vocabulary)
    ├── Regex-based page jump ("第N页" / "go to page N")
    ├── Descriptive context filter (suppresses "像第三页那样" etc.)
    └── Debounce (2s dedup)
    ↓
PPT Controller
    ├── macOS: Quartz CGEvent (low-level hardware key events)
    └── Windows: pynput keyboard simulation
```

## Project Structure

```
PPT_Project/
├── main.py                         # Bilingual GUI entry point (tkinter, Tesla theme)
├── main_en.py                      # English-only launcher (used by PyInstaller)
├── i18n.py                         # Chinese/English string tables
├── asr_engine.py                   # ASR engines: FunASREngine + VoskEngine
├── command_parser.py               # Voice command parser (Chinese/English, debounce)
├── ppt_controller.py               # Keyboard simulation (Quartz CGEvent / pynput)
├── config.py                       # Global configuration (engine, keywords, params)
├── download_model.py               # Vosk model download utility
├── requirements.txt                # Python dependencies
├── DESIGN.md                       # Tesla-inspired design system
├── build_app.sh                    # macOS packaging — Chinese edition
├── build_app_en.sh                 # macOS packaging — English edition
├── build_app_windows.bat           # Windows packaging — Chinese edition
├── build_app_windows_en.bat        # Windows packaging — English edition
├── build_app_windows_offline.bat   # Windows offline packaging (bundles model)
├── build_installer_offline.iss     # Inno Setup installer script
└── .github/workflows/
    ├── build-windows.yml           # CI — builds both zh and en Windows artefacts
    └── release.yml                 # Release — on v* tag, uploads zh & en zips
```

## Configuration

Edit `config.py` to customize:

| Option | Description | Default |
|---|---|---|
| `ASR_ENGINE` | Recognition engine (`"funasr"` or `"vosk"`) | `"funasr"` |
| `DEBOUNCE_SECONDS` | Min interval between duplicate commands | `2.0` s |
| `SAMPLE_RATE` | Audio sample rate | `16000` Hz |
| `NEXT_KEYWORDS`, etc. | Trigger keywords for each command | see file |

### Switching to Vosk

For a lighter setup (no PyTorch needed, ~42 MB model):

1. Set `ASR_ENGINE = "vosk"` in `config.py`
2. Run `python download_model.py` to download the Vosk model
3. Run `python main.py`

## Packaging

### macOS — Chinese edition

```bash
chmod +x build_app.sh && ./build_app.sh
```

Output: `dist/PPT语音控制助手/`

### macOS — English edition

```bash
chmod +x build_app_en.sh && ./build_app_en.sh
```

Output: `dist/PPT-Voice-Control/`

### Windows — Chinese edition

```bash
build_app_windows.bat
```

Output: `dist/PPT语音控制助手/`

### Windows — English edition

```bash
build_app_windows_en.bat
```

Output: `dist/PPT-Voice-Control/`

### Windows (offline, bundles model)

Builds a standalone package that includes the FunASR model — target machines need no Python and no internet.

```bash
build_app_windows_offline.bat
```

Output: `release/PPT-Voice-Control-Offline/` + `.zip`

To create a Windows installer (requires [Inno Setup](https://jrsoftware.org/isinfo.php)):

1. Copy the build output into `installer_staging/`
2. Run `build_installer_offline.iss` with Inno Setup Compiler

### GitHub Actions

- `.github/workflows/build-windows.yml` — builds **both** the `zh` and `en` Windows packages on every push to `main` (matrix build). Artifacts are retained for 30 days.
- `.github/workflows/release.yml` — on any `v*` tag push (e.g. `git tag v1.2.0 && git push --tags`), builds both editions, zips them as `PPT-Voice-Control-zh-CN-Windows.zip` / `PPT-Voice-Control-en-US-Windows.zip`, and attaches them to the matching GitHub Release.

## License

Internal use — United Imaging.
