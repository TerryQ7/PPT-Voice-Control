"""PPT Voice Control Assistant — bilingual main entry.

The same app runs in Simplified Chinese or English; the locale is selected at
startup via ``--lang {zh|en}`` or the ``PPT_LANG`` environment variable.

The UI is inspired by Tesla's marketing site as described in ``DESIGN.md``:
pure white canvas, Carbon Dark text, Electric Blue as the sole accent, flat
surfaces with no shadows, Universal-Sans-style typography (approximated with
system fonts), and a single blue CTA for the primary action.
"""

from __future__ import annotations

import argparse
import datetime
import logging
import os
import platform
import sys
import threading
import tkinter as tk
from tkinter import ttk, scrolledtext, messagebox
from typing import Optional


class _NullStream:
    """Placeholder stream for ``--windowed`` builds without a console."""

    def write(self, _data):
        return 0

    def flush(self):
        return None

    def isatty(self):
        return False


def _ensure_std_streams():
    for name in ("stdout", "stderr", "__stdout__", "__stderr__"):
        if getattr(sys, name, None) is None:
            setattr(sys, name, _NullStream())


_ensure_std_streams()

import sounddevice as sd

from config import ASR_ENGINE, VOSK_MODEL_PATH, MODEL_DIR, DEFAULT_VOSK_MODEL, VOSK_MODEL_URLS
from command_parser import CommandParser, Command, CommandType
from ppt_controller import PPTController, check_accessibility_permission
from asr_engine import VoskEngine, FunASREngine, ASREngineBase, get_available_devices
from i18n import Translator, normalize_locale

logging.basicConfig(
    level=logging.DEBUG,
    format="%(asctime)s [%(name)s] %(levelname)s: %(message)s",
    datefmt="%H:%M:%S",
)


# =====================================================================
# Tesla-inspired design tokens (mirrors DESIGN.md)
# =====================================================================

class Theme:
    # Colors
    ELECTRIC_BLUE = "#3E6AE1"
    ELECTRIC_BLUE_HOVER = "#3457B2"
    WHITE = "#FFFFFF"
    LIGHT_ASH = "#F4F4F4"
    CARBON_DARK = "#171A20"
    GRAPHITE = "#393C41"
    PEWTER = "#5C5E62"
    SILVER_FOG = "#8E8E8E"
    CLOUD_GRAY = "#EEEEEE"
    PALE_SILVER = "#D0D1D2"
    SUCCESS = "#3E6AE1"   # status dot uses the same blue
    WARNING = "#B4741A"   # muted amber for warnings
    DANGER = "#C2361A"    # muted red for errors

    # Type scale (Universal Sans is not freely redistributable; fall back to
    # system fonts that carry the same geometric feel: SF Pro on macOS,
    # Segoe UI Variable on Windows, and plain sans elsewhere.)
    @staticmethod
    def font_family() -> str:
        system = platform.system()
        if system == "Darwin":
            return "SF Pro Text"
        if system == "Windows":
            return "Segoe UI Variable"
        return "Helvetica"

    @staticmethod
    def font_display() -> str:
        system = platform.system()
        if system == "Darwin":
            return "SF Pro Display"
        if system == "Windows":
            return "Segoe UI Variable"
        return "Helvetica"


# =====================================================================
# Minimal flat "CTA" button (no shadows, 4px-ish radius via a solid fill)
# =====================================================================

class FlatButton(tk.Frame):
    """A flat, Tesla-style CTA button for Tk.

    Tk doesn't support rounded corners on widgets directly, so we draw the
    rectangle (and any optional icon glyph) on a Canvas and let hover/active
    states change the fill colour — Tesla's colour-only interaction pattern
    from ``DESIGN.md``. Three visual variants are supported:

    - ``primary``: solid Electric Blue rectangle, white label.
    - ``secondary``: white rectangle with a subtle Pale Silver border.
    - ``ghost``: borderless icon button that only paints a background on
      hover. Used for the refresh control next to the device picker.
    """

    def __init__(
        self,
        master,
        text: str = "",
        command=None,
        *,
        variant: str = "primary",
        icon: str | None = None,
        width: int = 200,
        height: int = 40,
        tooltip: str | None = None,
    ):
        super().__init__(master, bg=master.cget("bg") if isinstance(master, tk.Widget) else Theme.WHITE)
        self._command = command
        self._variant = variant
        self._icon = icon
        self._state = "normal"
        self._width = width
        self._height = height

        self.canvas = tk.Canvas(
            self,
            width=width,
            height=height,
            highlightthickness=0,
            bd=0,
            bg=self.cget("bg"),
        )
        self.canvas.pack()

        self._rect = self.canvas.create_rectangle(
            0, 0, width, height,
            fill=self._bg_for("normal"),
            outline=self._border_for("normal"),
        )

        self._icon_items: list[int] = []
        self._label = None
        if icon:
            self._draw_icon()
        if text:
            text_x = width // 2
            if icon:
                text_x = width // 2 + 10
            self._label = self.canvas.create_text(
                text_x,
                height // 2,
                text=text,
                fill=self._fg_for("normal"),
                font=(Theme.font_family(), 13, "normal"),
            )

        self.canvas.bind("<Enter>", self._on_enter)
        self.canvas.bind("<Leave>", self._on_leave)
        self.canvas.bind("<Button-1>", self._on_click)
        self.canvas.bind("<ButtonRelease-1>", self._on_release)

        if tooltip:
            _Tooltip(self.canvas, tooltip)

    # --- icon rendering ---

    def _draw_icon(self):
        """Draw the named icon centered in the canvas."""
        cx = self._width // 2 if not self._label else 22
        cy = self._height // 2
        color = self._fg_for("normal")

        if self._icon == "refresh":
            r = 7
            # Open ring (a circular-arrow). Two arc segments leave a gap so we
            # can attach an arrowhead at the end — that small triangle is what
            # makes it read as "refresh" rather than a generic spinner.
            self._icon_items.append(self.canvas.create_arc(
                cx - r, cy - r, cx + r, cy + r,
                start=40, extent=270, style="arc",
                outline=color, width=2,
            ))
            # Arrowhead at the end of the arc (top-right of the circle).
            tip_x, tip_y = cx + r - 1, cy - r + 2
            self._icon_items.append(self.canvas.create_polygon(
                tip_x - 4, tip_y - 1,
                tip_x + 4, tip_y - 1,
                tip_x, tip_y + 4,
                fill=color, outline=color,
            ))

    # --- state helpers ---

    def _bg_for(self, state: str) -> str:
        if self._state == "disabled":
            return Theme.CLOUD_GRAY if self._variant != "ghost" else self.cget("bg")
        if self._variant == "primary":
            return Theme.ELECTRIC_BLUE_HOVER if state == "hover" else Theme.ELECTRIC_BLUE
        if self._variant == "ghost":
            # Transparent at rest, faint blue tint on hover.
            return "#E7EDFB" if state == "hover" else self.cget("bg")
        # secondary
        return Theme.LIGHT_ASH if state == "hover" else Theme.WHITE

    def _fg_for(self, state: str) -> str:
        if self._state == "disabled":
            return Theme.SILVER_FOG
        if self._variant == "primary":
            return Theme.WHITE
        if self._variant == "ghost":
            return Theme.ELECTRIC_BLUE if state == "hover" else Theme.GRAPHITE
        return Theme.GRAPHITE

    def _border_for(self, state: str) -> str:
        if self._variant == "primary":
            return self._bg_for(state)
        if self._variant == "ghost":
            return self._bg_for(state)
        return Theme.PALE_SILVER

    def _paint(self, state: str = "normal"):
        self.canvas.itemconfig(self._rect,
                               fill=self._bg_for(state),
                               outline=self._border_for(state))
        if self._label is not None:
            self.canvas.itemconfig(self._label, fill=self._fg_for(state))
        for item in self._icon_items:
            kind = self.canvas.type(item)
            color = self._fg_for(state)
            if kind == "arc":
                self.canvas.itemconfig(item, outline=color)
            else:
                self.canvas.itemconfig(item, fill=color, outline=color)

    def set_text(self, text: str):
        if self._label is not None:
            self.canvas.itemconfig(self._label, text=text)

    def configure_state(self, state: str):
        self._state = state
        self._paint("normal")

    # --- events ---

    def _on_enter(self, _e):
        if self._state == "disabled":
            return
        self._paint("hover")

    def _on_leave(self, _e):
        if self._state == "disabled":
            return
        self._paint("normal")

    def _on_click(self, _e):
        if self._state == "disabled":
            return
        self._paint("hover")

    def _on_release(self, _e):
        if self._state == "disabled":
            return
        self._paint("hover")
        if callable(self._command):
            self._command()


class _Tooltip:
    """Minimal Tesla-style tooltip: small Carbon Dark pill on hover."""

    def __init__(self, widget: tk.Widget, text: str, delay_ms: int = 350):
        self._widget = widget
        self._text = text
        self._delay = delay_ms
        self._after_id: str | None = None
        self._tip: tk.Toplevel | None = None
        widget.bind("<Enter>", self._schedule, add="+")
        widget.bind("<Leave>", self._hide, add="+")
        widget.bind("<ButtonPress>", self._hide, add="+")

    def _schedule(self, _e=None):
        self._cancel()
        self._after_id = self._widget.after(self._delay, self._show)

    def _cancel(self):
        if self._after_id is not None:
            self._widget.after_cancel(self._after_id)
            self._after_id = None

    def _show(self):
        if self._tip is not None:
            return
        x = self._widget.winfo_rootx() + self._widget.winfo_width() // 2
        y = self._widget.winfo_rooty() + self._widget.winfo_height() + 6
        tw = tk.Toplevel(self._widget)
        tw.wm_overrideredirect(True)
        tw.configure(bg=Theme.CARBON_DARK)
        try:
            tw.attributes("-topmost", True)
        except tk.TclError:
            pass
        label = tk.Label(
            tw,
            text=self._text,
            bg=Theme.CARBON_DARK,
            fg=Theme.WHITE,
            font=(Theme.font_family(), 10),
            padx=10,
            pady=4,
            bd=0,
        )
        label.pack()
        tw.update_idletasks()
        tw.wm_geometry(f"+{x - tw.winfo_width() // 2}+{y}")
        self._tip = tw

    def _hide(self, _e=None):
        self._cancel()
        if self._tip is not None:
            self._tip.destroy()
            self._tip = None


# =====================================================================
# Main window
# =====================================================================

class PPTVoiceApp:
    """Tesla-inspired PPT voice controller."""

    def __init__(self, root: tk.Tk, locale: str = "zh"):
        self.t = Translator(locale).t
        self.locale = locale

        self.root = root
        self.root.title(self.t("app.title"))
        self.root.geometry("720x760")
        self.root.minsize(440, 560)
        self.root.configure(bg=Theme.WHITE)

        self.parser = CommandParser()
        self.controller = PPTController()
        self.engine: ASREngineBase | None = None

        self._audio_devices: list[tuple[int, str]] = []
        self._selected_device_idx: Optional[int] = None

        # Labels whose ``wraplength`` should follow the container width.
        # Each entry is ``(label, padding_compensation_px)``.
        self._responsive_labels: list[tuple[tk.Label, int]] = []
        self._last_wrap_width = 0

        self._configure_styles()
        self._build_ui()
        self._refresh_devices()
        self._check_accessibility()
        self._check_model()

        self.root.bind("<Configure>", self._on_root_configure)

    # --- styles ---

    def _configure_styles(self):
        style = ttk.Style()
        try:
            style.theme_use("clam")
        except tk.TclError:
            pass

        family = Theme.font_family()
        style.configure(
            "Tesla.TCombobox",
            fieldbackground=Theme.WHITE,
            background=Theme.WHITE,
            foreground=Theme.CARBON_DARK,
            arrowcolor=Theme.CARBON_DARK,
            bordercolor=Theme.PALE_SILVER,
            lightcolor=Theme.PALE_SILVER,
            darkcolor=Theme.PALE_SILVER,
            borderwidth=1,
            padding=6,
        )
        style.map(
            "Tesla.TCombobox",
            fieldbackground=[("readonly", Theme.WHITE)],
            foreground=[("readonly", Theme.CARBON_DARK)],
        )
        self.root.option_add("*TCombobox*Listbox.background", Theme.WHITE)
        self.root.option_add("*TCombobox*Listbox.foreground", Theme.CARBON_DARK)
        self.root.option_add("*TCombobox*Listbox.selectBackground", Theme.ELECTRIC_BLUE)
        self.root.option_add("*TCombobox*Listbox.selectForeground", Theme.WHITE)
        self.root.option_add("*TCombobox*Listbox.font", (family, 12))

    # --- layout ---

    def _build_ui(self):
        family = Theme.font_family()
        display = Theme.font_display()

        # Outer side margins shrink slightly on narrow screens via the
        # ``<Configure>`` handler below; start with the full Tesla-style 40px.
        self._h_padding = 40

        self._container = tk.Frame(self.root, bg=Theme.WHITE)
        self._container.pack(fill=tk.BOTH, expand=True,
                             padx=self._h_padding, pady=(32, 24))
        container = self._container

        # Header — hero-like title & subtitle, no chrome.
        title_label = tk.Label(
            container,
            text=self.t("app.title"),
            font=(display, 28, "normal"),
            fg=Theme.CARBON_DARK,
            bg=Theme.WHITE,
            anchor="w",
            justify="left",
        )
        title_label.pack(anchor="w", fill=tk.X)
        self._responsive_labels.append((title_label, 0))

        subtitle = tk.Label(
            container,
            text=self.t("app.subtitle"),
            font=(family, 13, "normal"),
            fg=Theme.PEWTER,
            bg=Theme.WHITE,
            wraplength=640,
            justify="left",
            anchor="w",
        )
        subtitle.pack(anchor="w", fill=tk.X, pady=(6, 22))
        self._responsive_labels.append((subtitle, 0))

        # Tip card — Light Ash surface, no border, flat.
        tip_card = tk.Frame(container, bg=Theme.LIGHT_ASH)
        tip_card.pack(fill=tk.X, pady=(0, 20))
        tip_title = tk.Label(
            tip_card,
            text=self.t("tip.title"),
            font=(family, 12, "normal"),
            fg=Theme.CARBON_DARK,
            bg=Theme.LIGHT_ASH,
            anchor="w",
            justify="left",
        )
        tip_title.pack(anchor="w", fill=tk.X, padx=16, pady=(12, 2))
        self._responsive_labels.append((tip_title, 32))

        tip_l1 = tk.Label(
            tip_card,
            text=self.t("tip.line1"),
            font=(family, 12),
            fg=Theme.GRAPHITE,
            bg=Theme.LIGHT_ASH,
            wraplength=620,
            justify="left",
            anchor="w",
        )
        tip_l1.pack(anchor="w", fill=tk.X, padx=16)
        self._responsive_labels.append((tip_l1, 32))

        tip_l2 = tk.Label(
            tip_card,
            text=self.t("tip.line2"),
            font=(family, 12),
            fg=Theme.GRAPHITE,
            bg=Theme.LIGHT_ASH,
            wraplength=620,
            justify="left",
            anchor="w",
        )
        tip_l2.pack(anchor="w", fill=tk.X, padx=16, pady=(0, 12))
        self._responsive_labels.append((tip_l2, 32))

        # Audio input panel
        tk.Label(
            container,
            text=self.t("section.audio"),
            font=(family, 13, "normal"),
            fg=Theme.CARBON_DARK,
            bg=Theme.WHITE,
        ).pack(anchor="w", pady=(2, 8))

        mic_row = tk.Frame(container, bg=Theme.WHITE)
        mic_row.pack(fill=tk.X, pady=(0, 18))

        tk.Label(
            mic_row,
            text=self.t("label.device"),
            font=(family, 12),
            fg=Theme.PEWTER,
            bg=Theme.WHITE,
        ).pack(side=tk.LEFT, padx=(0, 10))

        # Refresh button anchors to the right so the combobox flexes with width.
        self.refresh_btn = FlatButton(
            mic_row,
            command=self._refresh_devices,
            variant="ghost",
            icon="refresh",
            width=36,
            height=36,
            tooltip=self.t("btn.refresh"),
        )
        self.refresh_btn.pack(side=tk.RIGHT, padx=(8, 0))

        self.device_var = tk.StringVar()
        self.device_combo = ttk.Combobox(
            mic_row,
            textvariable=self.device_var,
            state="readonly",
            style="Tesla.TCombobox",
            font=(family, 12),
        )
        self.device_combo.pack(side=tk.LEFT, fill=tk.X, expand=True)
        self.device_combo.bind("<<ComboboxSelected>>", self._on_device_selected)

        # Status row — dot + label (label wraps so long status messages
        # stay visible at narrow widths instead of being clipped on the right).
        status_row = tk.Frame(container, bg=Theme.WHITE)
        status_row.pack(fill=tk.X, pady=(0, 6))
        self.status_canvas = tk.Canvas(
            status_row, width=14, height=14,
            bg=Theme.WHITE, highlightthickness=0, bd=0,
        )
        self.status_canvas.pack(side=tk.LEFT, padx=(0, 10), pady=(2, 0))
        self.status_dot = self.status_canvas.create_oval(
            1, 1, 13, 13, fill=Theme.SILVER_FOG, outline=""
        )
        self.status_label = tk.Label(
            status_row,
            text=self.t("status.idle"),
            font=(family, 13, "normal"),
            fg=Theme.CARBON_DARK,
            bg=Theme.WHITE,
            wraplength=500,
            justify="left",
            anchor="w",
        )
        self.status_label.pack(side=tk.LEFT, fill=tk.X, expand=True)
        self._responsive_labels.append((self.status_label, 24))  # 14 dot + 10 pad

        # Live partial transcript
        tk.Label(
            container,
            text=self.t("label.partial"),
            font=(family, 12, "normal"),
            fg=Theme.PEWTER,
            bg=Theme.WHITE,
        ).pack(anchor="w", pady=(12, 6))

        partial_surface = tk.Frame(container, bg=Theme.LIGHT_ASH)
        partial_surface.pack(fill=tk.X)
        self.partial_var = tk.StringVar(value=self.t("placeholder.partial"))
        self.partial_label = tk.Label(
            partial_surface,
            textvariable=self.partial_var,
            font=(family, 13),
            fg=Theme.SILVER_FOG,
            bg=Theme.LIGHT_ASH,
            anchor="w",
            justify="left",
            wraplength=600,
        )
        self.partial_label.pack(fill=tk.X, padx=14, pady=12)
        self._responsive_labels.append((self.partial_label, 28))

        # Command log
        tk.Label(
            container,
            text=self.t("label.log"),
            font=(family, 12, "normal"),
            fg=Theme.PEWTER,
            bg=Theme.WHITE,
        ).pack(anchor="w", pady=(18, 6))

        log_surface = tk.Frame(container, bg=Theme.LIGHT_ASH)
        log_surface.pack(fill=tk.BOTH, expand=True)
        mono_family = "Menlo" if platform.system() == "Darwin" else ("Consolas" if platform.system() == "Windows" else "Courier")
        self.log_text = scrolledtext.ScrolledText(
            log_surface,
            height=8,
            font=(mono_family, 11),
            state=tk.DISABLED,
            wrap=tk.WORD,
            bg=Theme.LIGHT_ASH,
            fg=Theme.GRAPHITE,
            borderwidth=0,
            highlightthickness=0,
            relief="flat",
            padx=14,
            pady=10,
        )
        self.log_text.pack(fill=tk.BOTH, expand=True)

        # CTA buttons (primary + secondary). Side-by-side on desktop; the
        # stop button drops below the start button on narrow widths via the
        # ``<Configure>`` handler.
        self._btn_row = tk.Frame(container, bg=Theme.WHITE)
        self._btn_row.pack(fill=tk.X, pady=(22, 0))

        self.start_btn = FlatButton(
            self._btn_row,
            text=self.t("btn.start"),
            command=self._on_start,
            variant="primary",
            width=220,
            height=44,
        )
        self.start_btn.pack(side=tk.LEFT, padx=(0, 12))

        self.stop_btn = FlatButton(
            self._btn_row,
            text=self.t("btn.stop"),
            command=self._on_stop,
            variant="secondary",
            width=160,
            height=44,
        )
        self.stop_btn.pack(side=tk.LEFT)
        self.stop_btn.configure_state("disabled")
        self._cta_stacked = False

        # Footer helper text — Tesla-style tertiary links
        footer = tk.Frame(self.root, bg=Theme.WHITE)
        footer.pack(fill=tk.X, side=tk.BOTTOM, padx=40, pady=(0, 18))

        footer_title = tk.Label(
            footer,
            text=self.t("commands.title"),
            font=(family, 11, "normal"),
            fg=Theme.CARBON_DARK,
            bg=Theme.WHITE,
            anchor="w",
            justify="left",
        )
        footer_title.pack(anchor="w", fill=tk.X)
        self._responsive_labels.append((footer_title, 0))

        for key in ("commands.line1", "commands.line2", "commands.line3"):
            line = tk.Label(
                footer,
                text=self.t(key),
                font=(family, 11),
                fg=Theme.PEWTER,
                bg=Theme.WHITE,
                wraplength=640,
                justify="left",
                anchor="w",
            )
            line.pack(anchor="w", fill=tk.X)
            self._responsive_labels.append((line, 0))

        self.root.protocol("WM_DELETE_WINDOW", self._on_close)

    # ==================== responsive layout ====================

    def _on_root_configure(self, event):
        """Reflow the UI when the window is resized."""
        # We only care about resize events on the root itself; child widgets
        # also fire <Configure> as they're laid out.
        if event.widget is not self.root:
            return

        width = self.root.winfo_width()
        if width <= 1:
            return

        # Tighten the outer side padding on narrow windows so the content has
        # room to breathe. Tesla shrinks the chrome on mobile too.
        new_pad = 40 if width >= 600 else (24 if width >= 480 else 16)
        if new_pad != self._h_padding:
            self._h_padding = new_pad
            self._container.pack_configure(padx=new_pad)

        # Recompute wraplength for every text-flow label. We use the actual
        # container width minus each label's local padding. The threshold
        # avoids a feedback loop where wrap → height change → another resize.
        wrap_base = max(width - 2 * self._h_padding, 200)
        if abs(wrap_base - self._last_wrap_width) >= 4:
            self._last_wrap_width = wrap_base
            for label, compensation in self._responsive_labels:
                target = max(wrap_base - compensation, 160)
                try:
                    label.configure(wraplength=target)
                except tk.TclError:
                    pass

        # Stack the CTA pair vertically when the row would overflow.
        # 220 (start) + 12 (gap) + 160 (stop) = 392px. Anything narrower
        # than ~440 internal width should stack.
        usable = width - 2 * self._h_padding
        should_stack = usable < 410
        if should_stack and not self._cta_stacked:
            self.stop_btn.pack_forget()
            self.start_btn.pack_forget()
            self.start_btn.pack(anchor="w", pady=(0, 10))
            self.stop_btn.pack(anchor="w")
            self._cta_stacked = True
        elif not should_stack and self._cta_stacked:
            self.stop_btn.pack_forget()
            self.start_btn.pack_forget()
            self.start_btn.pack(side=tk.LEFT, padx=(0, 12))
            self.stop_btn.pack(side=tk.LEFT)
            self._cta_stacked = False

    # ==================== audio devices ====================

    def _refresh_devices(self):
        self._audio_devices = get_available_devices()
        display_names = [f"[{idx}] {name}" for idx, name in self._audio_devices]

        default_label = self.t("label.default_device")
        try:
            default_dev = sd.query_devices(kind="input")
            if default_dev:
                default_label = self.t("label.default_device_named", name=default_dev["name"])
        except Exception:
            pass

        choices = [default_label] + display_names
        self.device_combo["values"] = choices
        if not self.device_var.get() or self.device_var.get() not in choices:
            self.device_combo.current(0)
            self._selected_device_idx = None

    def _on_device_selected(self, _event=None):
        sel = self.device_combo.current()
        if sel <= 0:
            self._selected_device_idx = None
        else:
            self._selected_device_idx = self._audio_devices[sel - 1][0]

    # ==================== permissions ====================

    def _check_accessibility(self):
        if platform.system() != "Darwin":
            return
        if not check_accessibility_permission():
            self._log(self.t("msg.accessibility_missing"))
            self._log(self.t("msg.accessibility_path"))
            messagebox.showwarning(
                self.t("dialog.accessibility_title"),
                self.t("dialog.accessibility_body"),
            )

    # ==================== model check ====================

    def _check_model(self):
        if ASR_ENGINE == "funasr":
            self._set_status(self.t("status.ready_funasr"), Theme.ELECTRIC_BLUE)
            self._log(self.t("msg.engine_funasr_hint"))
            return

        if os.path.isdir(VOSK_MODEL_PATH):
            self._set_status(self.t("status.ready_vosk"), Theme.ELECTRIC_BLUE)
            return

        self._set_status(self.t("status.download_model"), Theme.WARNING)
        if messagebox.askyesno(
            self.t("dialog.model_missing_title"),
            self.t("dialog.model_missing_body", model=DEFAULT_VOSK_MODEL),
        ):
            self._download_model_async()
        else:
            self._log(self.t("msg.download_manual"))

    def _download_model_async(self):
        self.start_btn.configure_state("disabled")
        self._set_status(self.t("status.downloading", pct=0), Theme.ELECTRIC_BLUE)
        self._log(self.t("msg.download_start"))

        def _do_download():
            try:
                from download_model import download_and_extract
                download_and_extract(
                    VOSK_MODEL_URLS[DEFAULT_VOSK_MODEL], MODEL_DIR,
                    progress_cb=lambda pct: self.root.after(
                        0, lambda: self._set_status(
                            self.t("status.downloading", pct=pct), Theme.ELECTRIC_BLUE
                        )
                    ),
                )
                self.root.after(0, self._on_download_done, True)
            except Exception as e:
                self.root.after(0, self._on_download_done, False, str(e))

        threading.Thread(target=_do_download, daemon=True).start()

    def _on_download_done(self, success: bool, error: str = ""):
        if success:
            self._set_status(self.t("status.download_done"), Theme.ELECTRIC_BLUE)
            self._log(self.t("msg.download_done"))
        else:
            self._set_status(self.t("status.download_failed"), Theme.DANGER)
            self._log(self.t("msg.download_failed", error=error))
        self.start_btn.configure_state("normal")

    # ==================== start / stop ====================

    def _on_start(self):
        if ASR_ENGINE == "vosk" and not os.path.isdir(VOSK_MODEL_PATH):
            messagebox.showerror(
                self.t("dialog.model_vosk_error_title"),
                self.t("dialog.model_vosk_error_body"),
            )
            return

        self.start_btn.configure_state("disabled")
        self._set_status(self.t("status.loading_model"), Theme.ELECTRIC_BLUE)
        self._log(self.t("status.loading_model"))
        self.root.update()

        device_idx = self._selected_device_idx

        def _do_start():
            try:
                engine: ASREngineBase
                if ASR_ENGINE == "funasr":
                    engine = FunASREngine(device_index=device_idx)
                else:
                    engine = VoskEngine(model_path=VOSK_MODEL_PATH, device_index=device_idx)
                engine.start(
                    on_partial=self._on_partial,
                    on_final=self._on_final,
                    on_no_audio_warning=self._on_no_audio_warning,
                )
                self.root.after(0, self._on_engine_started, engine)
            except Exception as e:
                self.root.after(0, self._on_engine_failed, str(e))

        threading.Thread(target=_do_start, daemon=True).start()

    def _on_engine_started(self, engine: ASREngineBase):
        self.engine = engine
        self.parser.reset()
        name = "FunASR" if ASR_ENGINE == "funasr" else "Vosk"
        device_name = self.device_var.get()
        self._set_status(self.t("status.listening", engine=name), Theme.ELECTRIC_BLUE)
        self._log(self.t("msg.start_listening", engine=name))
        self._log(self.t("msg.audio_device", name=device_name))
        self.stop_btn.configure_state("normal")
        self.device_combo.config(state=tk.DISABLED)

    def _on_engine_failed(self, error: str):
        self._set_status(self.t("status.start_failed"), Theme.DANGER)
        self._log(f"✗ {error}")
        self.start_btn.configure_state("normal")
        messagebox.showerror(self.t("dialog.start_failed_title"), error)

    def _on_stop(self):
        if self.engine and self.engine.is_running():
            self.engine.stop()
        self._set_status(self.t("status.stopped"), Theme.SILVER_FOG)
        self._log(self.t("msg.stop_listening"))
        self.start_btn.configure_state("normal")
        self.stop_btn.configure_state("disabled")
        self.device_combo.config(state="readonly")
        self.partial_var.set(self.t("placeholder.partial"))

    def _on_close(self):
        if self.engine and self.engine.is_running():
            self.engine.stop()
        self.root.destroy()
        _force_exit = threading.Timer(3.0, os._exit, args=(0,))
        _force_exit.daemon = True
        _force_exit.start()

    # ==================== recognition callbacks ====================

    def _on_partial(self, text: str):
        self.root.after(0, self._handle_partial, text)

    def _on_final(self, text: str):
        self.root.after(0, self._handle_final, text)

    def _on_no_audio_warning(self, current_device: str):
        self.root.after(0, self._handle_no_audio_warning, current_device)

    def _handle_no_audio_warning(self, current_device: str):
        self._set_status(self.t("status.no_audio"), Theme.WARNING)
        self._log(self.t("msg.no_audio_log", name=current_device))

        other_devices = []
        for idx, name in self._audio_devices:
            if self._selected_device_idx is not None:
                if idx != self._selected_device_idx:
                    other_devices.append(name)
            else:
                other_devices.append(name)

        device_list = "\n".join(f"  - {name}" for name in other_devices[:8])
        if not device_list:
            device_list = self.t("dialog.no_audio_none")

        messagebox.showwarning(
            self.t("dialog.no_audio_title"),
            self.t("dialog.no_audio_body", name=current_device, devices=device_list),
        )

    def _handle_partial(self, text: str):
        self.partial_var.set(text)
        self.partial_label.configure(foreground=Theme.SILVER_FOG)
        cmd = self.parser.parse(text)
        if cmd:
            self._execute_command(cmd, text, is_partial=True)

    def _handle_final(self, text: str):
        self.partial_var.set(text)
        self.partial_label.configure(foreground=Theme.CARBON_DARK)
        cmd = self.parser.parse(text)
        if cmd:
            self._execute_command(cmd, text, is_partial=False)

    def _execute_command(self, cmd: Command, raw_text: str, is_partial: bool):
        now = datetime.datetime.now().strftime("%H:%M:%S")
        source = self.t("log.source_partial" if is_partial else "log.source_final")
        label = self._format_command(cmd)
        self._log(f'[{now}] ({source}) "{raw_text}" → {label}')
        try:
            self.controller.execute(cmd)
        except Exception as e:
            self._log(self.t("msg.exec_failed", error=str(e)))

    def _format_command(self, cmd: Command) -> str:
        mapping = {
            CommandType.NEXT: self.t("cmd.next"),
            CommandType.PREV: self.t("cmd.prev"),
            CommandType.FIRST: self.t("cmd.first"),
            CommandType.LAST: self.t("cmd.last"),
        }
        if cmd.type == CommandType.GOTO:
            return self.t("cmd.goto", page=cmd.page)
        return mapping.get(cmd.type, str(cmd))

    # ==================== helpers ====================

    def _set_status(self, text: str, color: str):
        self.status_label.config(text=text)
        self.status_canvas.itemconfig(self.status_dot, fill=color)

    def _log(self, msg: str):
        self.log_text.config(state=tk.NORMAL)
        self.log_text.insert(tk.END, msg + "\n")
        self.log_text.see(tk.END)
        self.log_text.config(state=tk.DISABLED)


def _parse_locale_from_args() -> str:
    parser = argparse.ArgumentParser(add_help=False)
    parser.add_argument("--lang", dest="lang", default=None)
    parser.add_argument("-h", "--help", action="store_true")
    args, _ = parser.parse_known_args()
    if args.lang:
        return normalize_locale(args.lang)
    return normalize_locale(os.environ.get("PPT_LANG"))


def main():
    locale = _parse_locale_from_args()
    root = tk.Tk()
    PPTVoiceApp(root, locale=locale)
    root.mainloop()


if __name__ == "__main__":
    main()
