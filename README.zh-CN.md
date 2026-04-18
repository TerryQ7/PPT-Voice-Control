# PPT 语音控制助手

为 PowerPoint 与 Keynote 提供"免点击"的语音翻页。程序持续监听麦克风（或在远程会议中监听系统输出），识别中英文命令并自动切换幻灯片，主讲人不再需要鼠标和翻页器。

> English version: [**README.md**](./README.md)

默认识别引擎为 [FunASR Paraformer](https://github.com/modelscope/FunASR)（阿里达摩院的高精度离线 ASR，原生支持中英文混合识别）。轻量场景可切换到 Vosk。

仓库同时发布**简体中文版**与**英文版**两个发行版本，使用同一份代码、同一套引擎，只在 UI 字串上做本地化。

---

## 下载

每个 GitHub tag 都会触发 GitHub Actions 自动产出**两个 Windows 安装包**——每种 UI 语言一个。安装向导、安装路径、开始菜单/桌面快捷方式以及程序窗口标题均与所选语言一致。

| 版本     | 界面语言   | GitHub Release 中的资源文件                       |
|----------|------------|---------------------------------------------------|
| zh-CN    | 简体中文   | `PPT-Voice-Control-zh-CN-Setup-v<version>.exe`    |
| en-US    | English    | `PPT-Voice-Control-en-US-Setup-v<version>.exe`    |

最新正式版：**[v1.0.1 发布页](../../releases/latest)** · 历史版本：**[Releases](../../releases)**。

每个安装包都是完整的 Inno Setup 向导（约 950 MB），内置 FunASR 模型——目标机器**不需要安装 Python，也不需要联网**。两种版本使用不同的 AppId 和安装路径，可在同一台机器上共存。

> ℹ️ GitHub Releases 服务端会把资源 URL 中的非 ASCII 字符替换为 `.`，所以中文版安装包以 ASCII 形式（带 `zh-CN` 标识）上传。**安装完成后**，无论是开始菜单、安装目录还是窗口标题，都仍然是 `PPT语音控制助手`。如果用 `installer.iss` 在本地编译，文件名会保留原生的 `PPT语音控制助手-安装程序-v<version>.exe`。

---

## 主要特性

- **完全离线**：首次启动会自动下载约 1 GB 的模型，之后全程不联网。
- **双语界面**：简体中文（`--lang zh`，默认）与英文（`--lang en`）。
- **中英文语音识别**：支持中英文混合命令、中文数字（如 *第二十三页*）和英文基数/序数（如 *twenty-third page*、*page fifty*）。
- **载体短语容忍**（v1.0.1）：英文里像 *"continue with next page"*、*"we can move on to the next page"* 这种自然口语也能识别。
- **可选麦克风或系统音频**：UI 中可选择任意输入设备。远程会议时改用 Loopback（Windows WASAPI）或 macOS 上的虚拟声卡（如 [BlackHole](https://github.com/ExistentialAudio/BlackHole)）来捕获扬声器输出。
- **智能上下文过滤**：*"像第三页那样"*、*"as shown on page 3"* 这种叙述句**不会**触发跳转；只有明确的命令才会切片。
- **平台自适应 VAD**：基于能量的语音活动检测，启动时自动校准噪声基线，macOS 与 Windows 各自单独调参。
- **双 ASR 引擎**：默认 FunASR Paraformer（高精度），也可切到 Vosk（轻量，约 42 MB）。
- **跨平台键盘模拟**：macOS 走 Quartz CGEvent，Windows 走 pynput。
- **特斯拉风极简 UI**：纯白画布、Carbon Dark 文字、Electric Blue 强调色，无阴影、无渐变、留白克制。详见 [`DESIGN.md`](./DESIGN.md)。
- **一键打包**：macOS 与 Windows 都有 PyInstaller 脚本；同一份 `installer.iss` 模板支持中英文 Inno Setup 安装包。

---

## 支持的语音命令

| 语音内容 | 动作 |
|---|---|
| *下一页* / *next page* / *next slide* / *forward* / *continue with next page* | 下一页 |
| *上一页* / *previous page* / *previous slide* / *go back* | 上一页 |
| *第N页* / *跳到第N页* / *go to page N* / *jump to slide N* / *page fifty* | 跳到第 N 页 |
| *第一页* / *first page* / *start over* | 第一页 |
| *最后一页* / *last slide* / *go to end* | 最后一页 |

- 中文数字完整支持：*第二十三页* (page 23)、*翻到第一百页* (page 100)。
- 英文基数和序数都支持：*go to page fifty*、*twenty-third page*、*one hundred and first page*、*slide ninety-ninth*。
- **上下文感知过滤**：以下叙述句不会触发跳转——*像第三页描述的那样*、*the next page is interesting*、*as shown on page 3*、*similar to the previous slide*。

### 注意事项

- **PowerPoint 必须处于前台窗口**：程序通过键盘事件控制激活窗口；如果 PPT 不在前台，命令不会作用到它。
- **远程会议**：观众的声音从扬声器出，而不是麦克风进。请在输入选择器中切换到 Loopback 或虚拟声卡设备。

---

## 快速开始（开发者）

### 环境要求

- Python 3.10+
- macOS 或 Windows
- 一个可用的麦克风

### 安装依赖

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

> **macOS 提示**：如果出现 `No module named '_tkinter'`：
>
> ```bash
> brew install python-tk@3.13   # 版本号要和你的 Python 对应
> ```

### 运行

```bash
# 简体中文（默认）
python main.py

# 英文
python main.py --lang en
# 等价：python main_en.py
# 等价：PPT_LANG=en python main.py
```

打开 PPT 进入**幻灯片放映模式**，点击 **开始监听** / **Start Listening**，然后开始说话。

FunASR 模型（约 1 GB）会在首次启动时通过 ModelScope 下载，之后程序完全离线运行。

### macOS 权限

需要授予两个权限：

1. **麦克风**：首次启动时系统会自动弹窗。
2. **辅助功能**：用于模拟键盘事件。打开 `系统设置 → 隐私与安全性 → 辅助功能`，把你的终端（Terminal / iTerm / Cursor）或打包后的 `.app` 加入名单。

---

## 配置项

编辑 `config.py`：

| 选项 | 说明 | 默认值 |
|---|---|---|
| `ASR_ENGINE` | 识别引擎，`"funasr"`（默认）或 `"vosk"` | `"funasr"` |
| `DEBOUNCE_SECONDS` | 同一命令的最小重复间隔 | `2.0` 秒 |
| `SAMPLE_RATE` | 音频采样率 | `16000` Hz |
| `NEXT_KEYWORDS` 等 | 各命令的触发关键词列表 | 见文件 |

### 切换到 Vosk

如果想要更轻量的部署（无需 PyTorch、模型仅约 42 MB）：

1. 把 `config.py` 中的 `ASR_ENGINE` 改成 `"vosk"`。
2. 运行 `python download_model.py` 下载 Vosk 模型。
3. 运行 `python main.py`。

---

## 架构

```
音频源（用户可选）
   ├── 麦克风（本地演讲）
   └── 系统音频 / Loopback（远程会议）
        ↓
音频采集（sounddevice，100 ms 分片）
        ↓
语音活动检测（基于能量，平台自适应）
   ├── 启动时自动校准噪声（约 0.6 秒）
   ├── 相对峰值能量判定语音结束
   └── 强制 2.5 秒超时兜底
        ↓
语音识别（FunASR Paraformer / Vosk）
        ↓
命令解析
   ├── 关键词匹配（按语言分别使用长度预算）
   ├── 正则页码跳转（中文数字 + 英文基数/序数）
   ├── 按语言独立的描述性上下文过滤
   └── 2 秒去重
        ↓
PPT 控制
   ├── macOS：Quartz CGEvent（底层硬件级按键事件）
   └── Windows：pynput 键盘模拟
```

### 项目结构

```
PPT_Project/
├── main.py                         # 双语 GUI 入口（tkinter，Tesla 风）
├── main_en.py                      # 英文专用启动器（PyInstaller 用）
├── i18n.py                         # 中英文字串表
├── asr_engine.py                   # ASR 引擎：FunASREngine + VoskEngine
├── command_parser.py               # 语音命令解析
├── ppt_controller.py               # 键盘模拟（Quartz CGEvent / pynput）
├── config.py                       # 全局配置
├── download_model.py               # Vosk 模型下载脚本
├── requirements.txt                # Python 依赖
├── DESIGN.md                       # Tesla 风设计系统
├── installer.iss                   # 双语 Inno Setup 模板（zh / en）
├── build_app.sh / build_app_en.sh  # macOS PyInstaller 脚本
├── build_app_windows.bat /
│   build_app_windows_en.bat        # Windows PyInstaller 脚本
├── build_app_windows_offline.bat   # Windows 离线打包（含模型）
└── .github/workflows/
    ├── build-windows.yml           # CI：每次推送 main 都构建 zh + en 两个产物
    └── release.yml                 # Release：v* tag 触发，构建并上传两个安装包
```

---

## 打包

### macOS

```bash
chmod +x build_app.sh    && ./build_app.sh        # 中文版  → dist/PPT语音控制助手/
chmod +x build_app_en.sh && ./build_app_en.sh     # 英文版  → dist/PPT-Voice-Control/
```

### Windows（PyInstaller 文件夹）

```bat
build_app_windows.bat        :: 中文版  -> dist\PPT语音控制助手\
build_app_windows_en.bat     :: 英文版  -> dist\PPT-Voice-Control\
```

### Windows（含模型的离线包）

```bat
build_app_windows_offline.bat   :: -> release\PPT-Voice-Control-Offline\ + .zip
```

### Windows 安装包（.exe）—— 本地编译

需要安装 [Inno Setup 6](https://jrsoftware.org/isinfo.php)。中文版还需要 `ChineseSimplified.isl`（从 [jrsoftware/issrc](https://github.com/jrsoftware/issrc/blob/main/Files/Languages/Unofficial/ChineseSimplified.isl) 下载，放到 Inno Setup 的 `Languages` 目录）。

```bat
:: 1. 先用 PyInstaller 编出对应语言的目录
build_app_windows.bat            :: 中文
:: 或：
build_app_windows_en.bat         :: 英文

:: 2. 把产物放到 installer 期望的位置
xcopy /E /I /Y "dist\PPT语音控制助手"      installer_staging
:: 英文版：
xcopy /E /I /Y "dist\PPT-Voice-Control"   installer_staging

:: 3. 用同一个模板编译两种语言
"C:\Program Files (x86)\Inno Setup 6\ISCC.exe" /DAppLang=zh /DMyAppVersion=1.0.1 installer.iss
:: 或：
"C:\Program Files (x86)\Inno Setup 6\ISCC.exe" /DAppLang=en /DMyAppVersion=1.0.1 installer.iss
```

输出会落在 `release/`：

- `PPT语音控制助手-安装程序-v1.0.1.exe`（中文版，原生中文文件名）
- `PPT-Voice-Control-Setup-v1.0.1.exe`（英文版）

### GitHub Actions

- **`build-windows.yml`**：每次推送 `main` 时矩阵构建 `zh` 和 `en` 两个 Windows PyInstaller 产物，工件保留 30 天。
- **`release.yml`**：推送任意 `v*` tag（如 `git tag v1.0.2 && git push origin v1.0.2`）时，矩阵构建两个版本，分别用 `installer.iss` 编译出 `.exe`，再附加到对应的 GitHub Release。中文安装包在上传时会重命名为 ASCII 安全的 `PPT-Voice-Control-zh-CN-Setup-v<version>.exe`（向导和安装出来的程序仍然是中文）。

> 仓库约定：每次 release 必须严格上传**两个 `.exe` 安装包**（zh + en），不传压缩包、不传裸文件。详见 `.cursor/rules/release-packaging.mdc`。

---

## 许可

内部使用 —— 联影集团（United Imaging）。
