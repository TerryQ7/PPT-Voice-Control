# PPT Voice Control

Hands-free slide navigation for PowerPoint and Keynote. The app listens to your microphone (or system audio for remote meetings), recognizes Chinese and English commands, and switches slides for you — no clicking, no remote.

> 简体中文版说明：[**README.zh-CN.md**](./README.zh-CN.md)

The default speech engine is [FunASR Paraformer](https://github.com/modelscope/FunASR), Alibaba's high-accuracy offline ASR with native Chinese–English bilingual support. Vosk is available as a lightweight alternative.

A **Simplified Chinese edition** and an **English edition** are shipped from the same codebase; only the UI strings differ.

---

## Downloads

Each tagged release on GitHub Actions ships **two Windows installers** — one per UI language. The setup wizard, install path, shortcuts, and window title all follow the chosen language.

| Edition  | UI language | GitHub Release asset                              |
|----------|-------------|---------------------------------------------------|
| zh-CN    | 简体中文     | `PPT-Voice-Control-zh-CN-Setup-v<version>.exe`    |
| en-US    | English     | `PPT-Voice-Control-en-US-Setup-v<version>.exe`    |

Latest stable: **[v1.0.1 release page](../../releases/latest)** · all versions: **[Releases](../../releases)**.

Each installer is a fully self-contained Inno Setup wizard (~950 MB) that bundles the FunASR model — the target machine needs neither Python nor internet access. Both editions use distinct App IDs and install paths, so they can coexist on the same machine.

> ℹ️ GitHub Releases sanitises non-ASCII characters in asset URLs, so the Chinese installer is uploaded with an ASCII-safe `zh-CN` marker. Once installed, the program itself shows up as `PPT语音控制助手` everywhere (Start menu, install directory, window title). When you build locally with `installer.iss`, the file keeps its native Chinese name `PPT语音控制助手-安装程序-v<version>.exe`.

---

## Highlights

- **Fully offline.** After the first-launch model download (~1 GB), no internet is required.
- **Bilingual UI.** Simplified Chinese (`--lang zh`, default) and English (`--lang en`).
- **Bilingual speech recognition.** Mixed Chinese / English commands work, including Chinese numerals like *第二十三页* and English ordinals like *twenty-third page*.
- **Carrier-phrase tolerance.** Natural English presenter speech such as *"continue with next page"* or *"we can move on to the next page"* works out of the box. (Fixed in v1.0.1.)
- **Microphone or system audio.** Pick any input device in the UI. For remote meetings, use a Loopback device (Windows WASAPI) or a virtual driver such as [BlackHole](https://github.com/ExistentialAudio/BlackHole) on macOS to capture speaker output.
- **Smart context filtering.** *"像第三页那样"* / *"as shown on page 3"* will **not** trigger a jump; only explicit commands do.
- **Platform-adaptive VAD.** Energy-based voice-activity detection with auto-calibrated noise baseline, tuned separately for macOS and Windows.
- **Dual ASR engine.** FunASR Paraformer (high accuracy, default) or Vosk (lightweight, ~42 MB).
- **Cross-platform.** macOS (Quartz CGEvent) and Windows (pynput) keyboard simulation.
- **Tesla-inspired UI.** Pure white canvas, Carbon Dark text, Electric Blue accent, zero shadows, generous whitespace. See [`DESIGN.md`](./DESIGN.md).
- **One-click packaging.** PyInstaller scripts for macOS and Windows, plus a parameterized Inno Setup template for both languages.

---

## Voice commands

| Spoken phrase | Action |
|---|---|
| *下一页* / *next page* / *next slide* / *forward* / *continue with next page* | Next slide |
| *上一页* / *previous page* / *previous slide* / *go back* | Previous slide |
| *第N页* / *跳到第N页* / *go to page N* / *jump to slide N* / *page fifty* | Jump to slide N |
| *第一页* / *first page* / *start over* | First slide |
| *最后一页* / *last slide* / *go to end* | Last slide |

- Chinese numerals are fully supported, e.g. *第二十三页* (page 23), *翻到第一百页* (page 100).
- English cardinals and ordinals are supported, e.g. *go to page fifty*, *twenty-third page*, *one hundred and first page*, *slide ninety-ninth*.
- **Context-aware filter.** Narrative speech is rejected: *像第三页描述的那样*, *the next page is interesting*, *as shown on page 3*, *similar to the previous slide* will not navigate.

### Important

- **PowerPoint must be the foreground window** — the app sends keyboard events to the active window. If PPT is not in focus, commands will not reach it.
- **Remote meetings.** Voice from the audience comes through speakers, not your microphone. Pick a Loopback / virtual audio device in the input selector.

---

## Quick start (development)

### Requirements

- Python 3.10+
- macOS or Windows
- Microphone

### Install

```bash
# macOS
python3 -m venv venv && source venv/bin/activate

# Windows
python -m venv venv
venv\Scripts\activate
```

```bash
pip install -r requirements.txt
```

> **macOS note.** If you see `No module named '_tkinter'`:
>
> ```bash
> brew install python-tk@3.13   # match your Python version
> ```

### Run

```bash
# Simplified Chinese (default)
python main.py

# English
python main.py --lang en
# equivalent: python main_en.py
# equivalent: PPT_LANG=en python main.py
```

Open your slides in **slideshow mode**, click **Start Listening** / **开始监听**, and start speaking.

The FunASR model (~1 GB) is downloaded on first launch via ModelScope. After that, the system runs fully offline.

### macOS permissions

1. **Microphone** — granted via the standard system prompt on first launch.
2. **Accessibility** — required for keyboard simulation. Open `System Settings → Privacy & Security → Accessibility` and add your terminal app (Terminal / iTerm / Cursor) or the packaged `.app`.

---

## Configuration

Edit `config.py`:

| Option | Description | Default |
|---|---|---|
| `ASR_ENGINE` | `"funasr"` (default) or `"vosk"` | `"funasr"` |
| `DEBOUNCE_SECONDS` | Min interval between duplicate commands | `2.0` |
| `SAMPLE_RATE` | Audio sample rate | `16000` Hz |
| `NEXT_KEYWORDS`, `PREV_KEYWORDS`, `FIRST_KEYWORDS`, `LAST_KEYWORDS` | Per-command trigger keywords | see file |

### Switching to Vosk

For a lighter setup (no PyTorch, ~42 MB model):

1. Set `ASR_ENGINE = "vosk"` in `config.py`.
2. Run `python download_model.py` to fetch the Vosk model.
3. Run `python main.py`.

---

## Architecture

```
Audio Source (user-selectable)
   ├── Microphone (local presentation)
   └── System Audio / Loopback (remote meeting)
        ↓
Audio Capture (sounddevice, 100ms chunks)
        ↓
Voice Activity Detection (energy-based, platform-adaptive)
   ├── Auto noise calibration on startup (~0.6 s)
   ├── Relative-peak energy for end-of-speech detection
   └── Forced 2.5 s timeout safety net
        ↓
Speech Recognition (FunASR Paraformer / Vosk)
        ↓
Command Parser
   ├── Keyword matching (per-language length budget)
   ├── Regex page-jump (Chinese numerals + English cardinals/ordinals)
   ├── Per-language descriptive-context filter
   └── 2 s debounce
        ↓
PPT Controller
   ├── macOS: Quartz CGEvent (low-level hardware key events)
   └── Windows: pynput keyboard simulation
```

### Project layout

```
PPT_Project/
├── main.py                         # Bilingual GUI entry point (tkinter, Tesla theme)
├── main_en.py                      # English-only launcher (used by PyInstaller)
├── i18n.py                         # Chinese / English string tables
├── asr_engine.py                   # ASR engines: FunASREngine + VoskEngine
├── command_parser.py               # Voice command parser
├── ppt_controller.py               # Keyboard simulation (Quartz CGEvent / pynput)
├── config.py                       # Global configuration
├── download_model.py               # Vosk model download utility
├── requirements.txt                # Python dependencies
├── DESIGN.md                       # Tesla-inspired design system
├── installer.iss                   # Bilingual Inno Setup template (zh / en)
├── build_app.sh / build_app_en.sh  # macOS PyInstaller scripts
├── build_app_windows.bat /
│   build_app_windows_en.bat        # Windows PyInstaller scripts
├── build_app_windows_offline.bat   # Windows offline packaging
└── .github/workflows/
    ├── build-windows.yml           # CI — builds zh and en Windows artifacts on each push
    └── release.yml                 # Release — on v* tag, builds both installers and uploads them
```

---

## Packaging

### macOS

```bash
chmod +x build_app.sh    && ./build_app.sh        # Chinese  → dist/PPT语音控制助手/
chmod +x build_app_en.sh && ./build_app_en.sh     # English  → dist/PPT-Voice-Control/
```

### Windows (PyInstaller folder)

```bat
build_app_windows.bat        :: Chinese  -> dist\PPT语音控制助手\
build_app_windows_en.bat     :: English  -> dist\PPT-Voice-Control\
```

### Windows (offline bundle with model)

```bat
build_app_windows_offline.bat   :: -> release\PPT-Voice-Control-Offline\ + .zip
```

### Windows installer (.exe) — local build

Requires [Inno Setup 6](https://jrsoftware.org/isinfo.php). The Chinese build also needs `ChineseSimplified.isl` (download from [jrsoftware/issrc](https://github.com/jrsoftware/issrc/blob/main/Files/Languages/Unofficial/ChineseSimplified.isl) into the Inno Setup `Languages` folder).

```bat
:: 1. Build the PyInstaller bundle for the language you want.
build_app_windows.bat            :: Chinese
:: or
build_app_windows_en.bat         :: English

:: 2. Stage the build output where the installer expects it.
xcopy /E /I /Y "dist\PPT语音控制助手"      installer_staging
:: or for English:
xcopy /E /I /Y "dist\PPT-Voice-Control"   installer_staging

:: 3. Compile the bilingual installer template.
"C:\Program Files (x86)\Inno Setup 6\ISCC.exe" /DAppLang=zh /DMyAppVersion=1.0.1 installer.iss
:: or:
"C:\Program Files (x86)\Inno Setup 6\ISCC.exe" /DAppLang=en /DMyAppVersion=1.0.1 installer.iss
```

Outputs land in `release/`:

- `PPT语音控制助手-安装程序-v1.0.1.exe` (Chinese, native filename)
- `PPT-Voice-Control-Setup-v1.0.1.exe` (English)

### GitHub Actions

- **`build-windows.yml`** — builds the `zh` and `en` Windows PyInstaller bundles on every push to `main` (matrix build). Artifacts retained for 30 days.
- **`release.yml`** — on any `v*` tag push (e.g. `git tag v1.0.2 && git push origin v1.0.2`), builds both editions, compiles `installer.iss` for each language, and attaches the resulting `.exe` files to the matching GitHub Release. The Chinese installer is renamed to the ASCII-safe `PPT-Voice-Control-zh-CN-Setup-v<version>.exe` for upload (the wizard and installed product remain Chinese).

> Per repository policy, every release ships **exactly two `.exe` installers** (zh + en). No zip, no loose binaries. See `.cursor/rules/release-packaging.mdc`.

---

## License

Internal use — United Imaging.
