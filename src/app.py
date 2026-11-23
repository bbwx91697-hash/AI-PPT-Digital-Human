"""Modern Windows Desktop GUI for AI PPT Digital Human Teaching Assistant."""

from __future__ import annotations

import ctypes
import os
import subprocess
import sys
import threading
import time
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk
from PIL import Image, ImageTk

SRC_DIR = Path(__file__).resolve().parent
ROOT_DIR = SRC_DIR.parent
if str(SRC_DIR) not in sys.path:
    sys.path.insert(0, str(SRC_DIR))

import audio_manager
import layout_geometry
import ppt_inspector
import project_store
import tts_service

PALETTE = {
    "bg": "#F1F5F9",
    "card": "#FFFFFF",
    "border": "#CBD5E1",
    "border_soft": "#E2E8F0",
    "text": "#0F172A",
    "muted": "#64748B",
    "primary": "#2563EB",
    "primary_hover": "#1D4ED8",
    "primary_soft": "#DBEAFE",
    "success": "#059669",
    "success_bg": "#D1FAE5",
    "warn": "#B45309",
    "warn_bg": "#FEF3C7",
    "canvas_bg": "#0F172A",
    "sel_row": "#EFF6FF",
}

FONT_BODY = ("Microsoft YaHei UI", 9)
FONT_TITLE = ("Microsoft YaHei UI", 12, "bold")
FONT_SECTION = ("Microsoft YaHei UI", 10, "bold")
FONT_MONO = ("Consolas", 9)


def _enable_high_dpi() -> None:
    try:
        ctypes.windll.shcore.SetProcessDpiAwareness(1)
    except Exception:
        pass


class DigitalHumanApp(tk.Tk):
    def __init__(self) -> None:
        _enable_high_dpi()
        super().__init__()
        self.title("AI PPT 数字人教学助手")
        self.geometry("1360x840")
        self.minsize(1140, 720)
        self.configure(bg=PALETTE["bg"])

        self.proj_dir: Path | None = None
        self.proj_data: dict | None = None
        self.deck_info: ppt_inspector.DeckInfo | None = None
        self.audio_cache: audio_manager.PageAudioCache | None = None
        self.thumb_paths: list[Path] = []
        self.current_slide: int | None = None
        self.last_exported_pptx: Path | None = None

        self._preview_photo: ImageTk.PhotoImage | None = None
        self._avatar_photo: ImageTk.PhotoImage | None = None
        self._avatar_img_size: tuple[int, int] = (480, 720)
        self._slide_rect: tuple[int, int, int, int] = (0, 0, 1, 1)  # ox, oy, rw, rh
        self._drag_state: dict | None = None
        self._busy = False

        self._configure_styles()
        self._build_ui()
        self._refresh_key_badge()
        self._refresh_avatar_label()

    # ------------------------------------------------------------------ Theme & UI
    def _configure_styles(self) -> None:
        style = ttk.Style(self)
        style.theme_use("clam")

        style.configure(".", background=PALETTE["bg"], foreground=PALETTE["text"], font=FONT_BODY)
        style.configure("Card.TFrame", background=PALETTE["card"])
        style.configure("CardTitle.TLabel", background=PALETTE["card"], foreground=PALETTE["text"], font=FONT_SECTION)
        style.configure("Muted.TLabel", background=PALETTE["card"], foreground=PALETTE["muted"], font=FONT_BODY)
        style.configure("Card.TCheckbutton", background=PALETTE["card"], foreground=PALETTE["text"], font=FONT_BODY)
        style.configure("Card.TRadiobutton", background=PALETTE["card"], foreground=PALETTE["text"], font=FONT_BODY)

        style.configure(
            "Primary.TButton",
            background=PALETTE["primary"],
            foreground="#FFFFFF",
            borderwidth=0,
            focusthickness=0,
            padding=(12, 7),
            font=("Microsoft YaHei UI", 9, "bold"),
        )
        style.map(
            "Primary.TButton",
            background=[("active", PALETTE["primary_hover"]), ("pressed", PALETTE["primary_hover"])],
        )

        style.configure(
            "Secondary.TButton",
            background="#F8FAFC",
            foreground=PALETTE["text"],
            bordercolor=PALETTE["border"],
            lightcolor=PALETTE["border"],
            darkcolor=PALETTE["border"],
            borderwidth=1,
            padding=(10, 6),
            font=FONT_BODY,
        )
        style.map("Secondary.TButton", background=[("active", "#E2E8F0")])

        style.configure(
            "Pill.TButton",
            background="#EFF6FF",
            foreground=PALETTE["primary"],
            bordercolor="#BFDBFE",
            borderwidth=1,
            padding=(8, 4),
            font=("Microsoft YaHei UI", 8, "bold"),
        )
        style.map("Pill.TButton", background=[("active", "#DBEAFE")])

        style.configure(
            "Treeview",
            background=PALETTE["card"],
            fieldbackground=PALETTE["card"],
            foreground=PALETTE["text"],
            rowheight=32,
            borderwidth=0,
            font=FONT_BODY,
        )
        style.configure(
            "Treeview.Heading",
            background="#F8FAFC",
            foreground=PALETTE["muted"],
            relief="flat",
            borderwidth=0,
            padding=(6, 6),
            font=("Microsoft YaHei UI", 9, "bold"),
        )
        style.map(
            "Treeview",
            background=[("selected", PALETTE["primary_soft"])],
            foreground=[("selected", PALETTE["primary"])],
        )

    def _make_card(self, parent, **pack_opts) -> tk.Frame:
        outer = tk.Frame(
            parent,
            bg=PALETTE["card"],
            highlightbackground=PALETTE["border_soft"],
            highlightthickness=1,
            bd=0,
        )
        outer.pack(**pack_opts)
        return outer

    def _build_ui(self) -> None:
        # Top Header Bar
        header = tk.Frame(
            self, bg=PALETTE["card"], highlightbackground=PALETTE["border_soft"], highlightthickness=1,
        )
        header.pack(fill="x", side="top")
        header_inner = tk.Frame(header, bg=PALETTE["card"], padx=16, pady=10)
        header_inner.pack(fill="x")

        brand = tk.Frame(header_inner, bg=PALETTE["card"])
        brand.pack(side="left")
        badge = tk.Label(
            brand, text="AI", bg=PALETTE["primary"], fg="#FFFFFF",
            font=("Microsoft YaHei UI", 10, "bold"), padx=8, pady=3,
        )
        badge.pack(side="left", padx=(0, 10))
        title_box = tk.Frame(brand, bg=PALETTE["card"])
        title_box.pack(side="left")
        tk.Label(title_box, text="AI PPT 数字人教学助手", bg=PALETTE["card"], fg=PALETTE["text"], font=FONT_TITLE).pack(anchor="w")
        self.deck_subtitle_var = tk.StringVar(value="未打开课件 · 保留原 PPT 可编辑性与原始画面")
        tk.Label(title_box, textvariable=self.deck_subtitle_var, bg=PALETTE["card"], fg=PALETTE["muted"], font=("Microsoft YaHei UI", 8)).pack(anchor="w")

        actions = tk.Frame(header_inner, bg=PALETTE["card"])
        actions.pack(side="right")
        self.key_label = tk.Label(actions, text="", font=("Microsoft YaHei UI", 8, "bold"), padx=8, pady=4)
        self.key_label.pack(side="left", padx=(0, 8))
        self.avatar_var = tk.StringVar()
        tk.Label(actions, textvariable=self.avatar_var, bg="#F1F5F9", fg=PALETTE["muted"], font=("Microsoft YaHei UI", 8), padx=8, pady=4).pack(side="left", padx=(0, 10))

        ttk.Button(actions, text="打开 / 导入 PPTX", style="Primary.TButton", command=self.on_open_pptx).pack(side="left", padx=4)
        ttk.Button(actions, text="保存配置", style="Secondary.TButton", command=self.on_save_project).pack(side="left", padx=4)
        ttk.Button(actions, text="更换数字人 GIF", style="Secondary.TButton", command=self.on_change_avatar).pack(side="left", padx=4)
        ttk.Button(actions, text="API Key 设置", style="Secondary.TButton", command=self.on_config_key).pack(side="left", padx=(4, 0))

        # Main 3-column Workspace
        workspace = tk.Frame(self, bg=PALETTE["bg"], padx=12, pady=10)
        workspace.pack(fill="both", expand=True)
        workspace.columnconfigure(0, weight=2, minsize=290)
        workspace.columnconfigure(1, weight=5, minsize=520)
        workspace.columnconfigure(2, weight=3, minsize=380)
        workspace.rowconfigure(0, weight=1)

        # ---- Left Column: Slide List ----
        left_card = tk.Frame(workspace, bg=PALETTE["card"], highlightbackground=PALETTE["border_soft"], highlightthickness=1)
        left_card.grid(row=0, column=0, sticky="nsew", padx=(0, 8))
        left_head = tk.Frame(left_card, bg=PALETTE["card"], padx=12, pady=10)
        left_head.pack(fill="x")
        tk.Label(left_head, text="幻灯片页面", bg=PALETTE["card"], fg=PALETTE["text"], font=FONT_SECTION).pack(side="left")
        self.enabled_count_var = tk.StringVar(value="0 页已启用")
        tk.Label(left_head, textvariable=self.enabled_count_var, bg=PALETTE["primary_soft"], fg=PALETTE["primary"], font=("Microsoft YaHei UI", 8, "bold"), padx=7, pady=2).pack(side="right")

        tree_wrap = tk.Frame(left_card, bg=PALETTE["card"])
        tree_wrap.pack(fill="both", expand=True, padx=8, pady=(0, 8))
        cols = ("page", "title", "enabled", "audio", "anim")
        self.tree = ttk.Treeview(tree_wrap, columns=cols, show="headings", selectmode="browse")
        for cid, heading, width, anchor in [
            ("page", "页", 40, "center"),
            ("title", "标题", 135, "w"),
            ("enabled", "启用", 44, "center"),
            ("audio", "语音", 60, "center"),
            ("anim", "动画", 44, "center"),
        ]:
            self.tree.heading(cid, text=heading)
            self.tree.column(cid, width=width, anchor=anchor)
        self.tree.tag_configure("odd", background="#FFFFFF")
        self.tree.tag_configure("even", background="#F8FAFC")
        scroll = ttk.Scrollbar(tree_wrap, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=scroll.set)
        self.tree.pack(side="left", fill="both", expand=True)
        scroll.pack(side="right", fill="y")
        self.tree.bind("<<TreeviewSelect>>", self.on_select_slide)

        # ---- Center Column: Interactive Preview & Position Controls ----
        center_card = tk.Frame(workspace, bg=PALETTE["card"], highlightbackground=PALETTE["border_soft"], highlightthickness=1)
        center_card.grid(row=0, column=1, sticky="nsew", padx=8)

        center_head = tk.Frame(center_card, bg=PALETTE["card"], padx=14, pady=10)
        center_head.pack(fill="x")
        self.preview_title_var = tk.StringVar(value="页面预览与数字人自由定位")
        tk.Label(center_head, textvariable=self.preview_title_var, bg=PALETTE["card"], fg=PALETTE["text"], font=FONT_SECTION).pack(side="left")
        tk.Label(
            center_head,
            text="直接拖动人物移位 · 拖动四角手柄缩放",
            bg="#EFF6FF", fg=PALETTE["primary"], font=("Microsoft YaHei UI", 8), padx=8, pady=2,
        ).pack(side="right")

        canvas_wrap = tk.Frame(center_card, bg=PALETTE["canvas_bg"], padx=6, pady=6)
        canvas_wrap.pack(fill="both", expand=True, padx=12)
        self.canvas = tk.Canvas(canvas_wrap, bg=PALETTE["canvas_bg"], highlightthickness=0, cursor="arrow")
        self.canvas.pack(fill="both", expand=True)
        self.canvas.bind("<Configure>", lambda _e: self._redraw_preview())
        self.canvas.bind("<ButtonPress-1>", self._on_canvas_press)
        self.canvas.bind("<B1-Motion>", self._on_canvas_drag)
        self.canvas.bind("<ButtonRelease-1>", self._on_canvas_release)
        self.canvas.bind("<Motion>", self._on_canvas_hover)

        # Position toolbar below canvas
        pos_panel = tk.Frame(center_card, bg=PALETTE["card"], padx=14, pady=10)
        pos_panel.pack(fill="x")

        preset_row = tk.Frame(pos_panel, bg=PALETTE["card"])
        preset_row.pack(fill="x", pady=(0, 8))
        tk.Label(preset_row, text="快捷对齐:", bg=PALETTE["card"], fg=PALETTE["muted"], font=FONT_BODY).pack(side="left", padx=(0, 6))
        for label, box in [
            ("左上角", (0.03, 0.06, 0.16, 0.35)),
            ("右上角", (0.81, 0.06, 0.16, 0.35)),
            ("左下角", (0.03, 0.60, 0.16, 0.35)),
            ("右下角", (0.81, 0.60, 0.16, 0.35)),
            ("底部居中", (0.42, 0.60, 0.16, 0.35)),
        ]:
            ttk.Button(preset_row, text=label, style="Pill.TButton", command=lambda b=box: self._apply_preset(*b)).pack(side="left", padx=3)

        self.lock_aspect_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(
            preset_row, text="锁定人物比例", variable=self.lock_aspect_var, style="Card.TCheckbutton",
        ).pack(side="right")

        num_row = tk.Frame(pos_panel, bg=PALETTE["card"])
        num_row.pack(fill="x")
        self.left_pct_var = tk.DoubleVar(value=81.0)
        self.top_pct_var = tk.DoubleVar(value=60.0)
        self.width_pct_var = tk.DoubleVar(value=16.0)
        self.height_pct_var = tk.DoubleVar(value=35.0)

        for idx, (lbl, var, lo, hi) in enumerate([
            ("X 位置(%)", self.left_pct_var, 0.0, 95.0),
            ("Y 位置(%)", self.top_pct_var, 0.0, 92.0),
            ("宽度(%)", self.width_pct_var, 5.0, 80.0),
            ("高度(%)", self.height_pct_var, 8.0, 90.0),
        ]):
            tk.Label(num_row, text=lbl, bg=PALETTE["card"], fg=PALETTE["muted"], font=("Microsoft YaHei UI", 8)).pack(side="left", padx=(0, 3))
            spin = ttk.Spinbox(
                num_row, from_=lo, to=hi, increment=1.0, textvariable=var, width=6, format="%.1f",
                command=self.on_numeric_layout_changed,
            )
            spin.pack(side="left", padx=(0, 10))
            spin.bind("<Return>", lambda _e: self.on_numeric_layout_changed())
            spin.bind("<FocusOut>", lambda _e: self.on_numeric_layout_changed())

        ttk.Button(
            num_row, text="应用位置到所有启用页", style="Secondary.TButton", command=self.on_apply_layout_all,
        ).pack(side="right")

        # ---- Right Column: Script, Voice, Audio & Export ----
        right_card = tk.Frame(workspace, bg=PALETTE["card"], highlightbackground=PALETTE["border_soft"], highlightthickness=1)
        right_card.grid(row=0, column=2, sticky="nsew", padx=(8, 0))
        right_inner = tk.Frame(right_card, bg=PALETTE["card"], padx=14, pady=12)
        right_inner.pack(fill="both", expand=True)

        self.enabled_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(
            right_inner, text="在本页启用数字人提问", variable=self.enabled_var,
            style="Card.TCheckbutton", command=self.on_enabled_toggled,
        ).pack(anchor="w", pady=(0, 8))

        script_head = tk.Frame(right_inner, bg=PALETTE["card"])
        script_head.pack(fill="x")
        tk.Label(script_head, text="数字人台词", bg=PALETTE["card"], fg=PALETTE["text"], font=FONT_SECTION).pack(side="left")
        self.char_count_var = tk.StringVar(value="0 字")
        tk.Label(script_head, textvariable=self.char_count_var, bg=PALETTE["card"], fg=PALETTE["muted"], font=("Microsoft YaHei UI", 8)).pack(side="right")

        text_box_wrap = tk.Frame(right_inner, bg=PALETTE["border"], padx=1, pady=1)
        text_box_wrap.pack(fill="x", pady=(4, 10))
        self.script_text = tk.Text(
            text_box_wrap, height=5, wrap="word", font=("Microsoft YaHei UI", 10),
            relief="flat", bg="#FFFFFF", fg=PALETTE["text"], padx=8, pady=6,
        )
        self.script_text.pack(fill="x")
        self.script_text.bind("<KeyRelease>", self._on_script_edited)

        tk.Label(right_inner, text="全自动语音配置", bg=PALETTE["card"], fg=PALETTE["text"], font=FONT_SECTION).pack(anchor="w", pady=(0, 4))
        tts_grid = tk.Frame(right_inner, bg=PALETTE["card"])
        tts_grid.pack(fill="x", pady=(0, 8))
        init_engine = tts_service.default_engine()
        self.engine_label_var = tk.StringVar(value=tts_service.ENGINE_LABELS[init_engine])
        self.model_var = tk.StringVar(value=tts_service.MODELS[0])
        self.voice_var = tk.StringVar(value="")
        self.style_var = tk.StringVar(value="温和、亲切、清晰")
        self.pace_var = tk.StringVar(value="适中")

        tk.Label(tts_grid, text="引擎", bg=PALETTE["card"], fg=PALETTE["muted"], width=5, anchor="w").grid(row=0, column=0, pady=3)
        self.engine_combo = ttk.Combobox(
            tts_grid, textvariable=self.engine_label_var,
            values=tuple(tts_service.ENGINE_LABELS.values()), state="readonly",
        )
        self.engine_combo.grid(row=0, column=1, sticky="ew", padx=(6, 0), pady=3)
        self.engine_combo.bind("<<ComboboxSelected>>", self._on_engine_changed)

        tk.Label(tts_grid, text="音色", bg=PALETTE["card"], fg=PALETTE["muted"], width=5, anchor="w").grid(row=1, column=0, pady=3)
        self.voice_combo = ttk.Combobox(tts_grid, textvariable=self.voice_var, state="readonly")
        self.voice_combo.grid(row=1, column=1, sticky="ew", padx=(6, 0), pady=3)

        tk.Label(tts_grid, text="语速", bg=PALETTE["card"], fg=PALETTE["muted"], width=5, anchor="w").grid(row=2, column=0, pady=3)
        ttk.Combobox(tts_grid, textvariable=self.pace_var, values=("偏慢", "适中", "偏快")).grid(row=2, column=1, sticky="ew", padx=(6, 0), pady=3)

        tk.Label(tts_grid, text="模型", bg=PALETTE["card"], fg=PALETTE["muted"], width=5, anchor="w").grid(row=3, column=0, pady=3)
        self.model_combo = ttk.Combobox(tts_grid, textvariable=self.model_var, values=tts_service.MODELS, state="readonly")
        self.model_combo.grid(row=3, column=1, sticky="ew", padx=(6, 0), pady=3)

        tk.Label(tts_grid, text="风格", bg=PALETTE["card"], fg=PALETTE["muted"], width=5, anchor="w").grid(row=4, column=0, pady=3)
        self.style_entry = ttk.Entry(tts_grid, textvariable=self.style_var)
        self.style_entry.grid(row=4, column=1, sticky="ew", padx=(6, 0), pady=3)
        tts_grid.columnconfigure(1, weight=1)
        self._apply_engine_ui_state(init_engine)

        btn_row1 = tk.Frame(right_inner, bg=PALETTE["card"])
        btn_row1.pack(fill="x", pady=(2, 4))
        ttk.Button(btn_row1, text="一键生成本页语音", style="Primary.TButton", command=self.on_generate_current).pack(side="left", fill="x", expand=True, padx=(0, 4))
        ttk.Button(btn_row1, text="导入音频 / 视频音轨", style="Secondary.TButton", command=self.on_import_local_audio).pack(side="left", fill="x", expand=True, padx=(4, 0))

        btn_row2 = tk.Frame(right_inner, bg=PALETTE["card"])
        btn_row2.pack(fill="x", pady=4)
        ttk.Button(btn_row2, text="试听本页", style="Secondary.TButton", command=self.on_play_audio).pack(side="left", fill="x", expand=True, padx=(0, 3))
        ttk.Button(btn_row2, text="停止试听", style="Secondary.TButton", command=audio_manager.stop_playback).pack(side="left", fill="x", expand=True, padx=3)
        ttk.Button(btn_row2, text="批量生成启用页", style="Secondary.TButton", command=self.on_generate_all).pack(side="left", fill="x", expand=True, padx=(3, 0))

        self.audio_info_var = tk.StringVar(value="当前页音频：未配置")
        self.audio_badge = tk.Label(
            right_inner, textvariable=self.audio_info_var, bg="#F8FAFC", fg=PALETTE["muted"],
            font=("Microsoft YaHei UI", 8), anchor="w", padx=10, pady=6,
            highlightbackground=PALETTE["border_soft"], highlightthickness=1,
        )
        self.audio_badge.pack(fill="x", pady=(6, 12))

        tk.Frame(right_inner, bg=PALETTE["border_soft"], height=1).pack(fill="x", pady=(0, 10))

        tk.Label(right_inner, text="课件导出与课堂放映", bg=PALETTE["card"], fg=PALETTE["text"], font=FONT_SECTION).pack(anchor="w", pady=(0, 6))
        mode_frame = tk.Frame(right_inner, bg=PALETTE["card"])
        mode_frame.pack(fill="x", pady=(0, 8))
        self.mode_var = tk.StringVar(value="auto")
        ttk.Radiobutton(mode_frame, text="自动播放(含点击重播)", value="auto", variable=self.mode_var, style="Card.TRadiobutton").pack(side="left", padx=(0, 10))
        ttk.Radiobutton(mode_frame, text="点击人物播放", value="manual", variable=self.mode_var, style="Card.TRadiobutton").pack(side="left")

        ttk.Button(right_inner, text="一键生成并导出新 PPTX", style="Primary.TButton", command=self.on_export_pptx).pack(fill="x", pady=(2, 6))

        show_row = tk.Frame(right_inner, bg=PALETTE["card"])
        show_row.pack(fill="x")
        ttk.Button(show_row, text="启动智能放映(前后翻页重播)", style="Secondary.TButton", command=self.on_start_smart_show).pack(side="left", fill="x", expand=True, padx=(0, 4))
        ttk.Button(show_row, text="在 PowerPoint 中打开", style="Secondary.TButton", command=self.on_open_exported).pack(side="left", fill="x", expand=True, padx=(4, 0))

        # Bottom Status Bar
        status_bar = tk.Frame(self, bg="#E2E8F0", padx=14, pady=5)
        status_bar.pack(fill="x", side="bottom")
        init_status = (
            "就绪：点击右上角“打开 / 导入 PPTX”开始配置课件（默认 Qwen3-TTS 本地高音质）。"
            if tts_service.is_qwen3_available()
            else "就绪：暂未安装本地 Qwen3 模型（可双击目录内「一键安装本地Qwen3高音质模型.bat」升级，当前暂用微软神经语音）。"
        )
        self.status_var = tk.StringVar(value=init_status)
        tk.Label(status_bar, textvariable=self.status_var, bg="#E2E8F0", fg=PALETTE["text"], font=("Microsoft YaHei UI", 8)).pack(side="left")
        self.progress = ttk.Progressbar(status_bar, mode="indeterminate", length=140)
        self.progress.pack(side="right")

    # ------------------------------------------------------------------ Helpers
    @staticmethod
    def _safe_initialdir(sub: str, create: bool = False) -> str:
        target = ROOT_DIR / sub
        if create:
            target.mkdir(parents=True, exist_ok=True)
        return str(target if target.is_dir() else ROOT_DIR)

    def _set_status(self, msg: str) -> None:
        self.status_var.set(msg)

    def _refresh_key_badge(self) -> None:
        if tts_service.has_api_key():
            self.key_label.configure(text="Gemini Key: 已加密", bg=PALETTE["success_bg"], fg=PALETTE["success"])
        else:
            self.key_label.configure(text="Gemini Key: 未配置", bg=PALETTE["warn_bg"], fg=PALETTE["warn"])

    def _refresh_avatar_label(self) -> None:
        gif = Path(self.proj_data["avatar_gif"]) if self.proj_data else project_store.DEFAULT_GIF
        idle = Path(self.proj_data["avatar_idle"]) if self.proj_data else project_store.DEFAULT_IDLE
        if idle.is_file():
            try:
                with Image.open(idle) as im:
                    self._avatar_img_size = im.size
            except Exception:
                pass
        self.avatar_var.set(f"全局数字人: {gif.name}")

    def _refresh_enabled_count(self) -> None:
        if not self.proj_data:
            self.enabled_count_var.set("0 页已启用")
            return
        total = len(self.proj_data.get("pages", {}))
        enabled = sum(1 for c in self.proj_data.get("pages", {}).values() if c.get("enabled"))
        self.enabled_count_var.set(f"已启用 {enabled} / {total} 页")

    def _run_bg(self, label: str, fn, on_ok) -> None:
        if self._busy:
            messagebox.showinfo("提示", "当前有任务正在执行，请稍候。")
            return
        self._busy = True
        self._busy_label = label
        self._busy_start = time.monotonic()
        self._set_status(label)
        self.progress.start(12)
        self.after(1000, self._tick_busy_status)

        def worker():
            try:
                res = fn()
            except Exception as exc:
                self.after(0, lambda err=exc: self._on_bg_error(err))
                return
            self.after(0, lambda out=res: self._on_bg_success(out, on_ok))

        threading.Thread(target=worker, daemon=True).start()

    def _tick_busy_status(self) -> None:
        if not self._busy:
            return
        elapsed = int(time.monotonic() - self._busy_start)
        if elapsed >= 2:
            self._set_status(f"{self._busy_label}（已耗时 {elapsed}s）")
        self.after(1000, self._tick_busy_status)

    def _on_bg_error(self, exc: Exception) -> None:
        self._busy = False
        self.progress.stop()
        self._set_status(f"操作失败: {exc}")
        messagebox.showerror("操作失败", str(exc))

    def _on_bg_success(self, res, on_ok) -> None:
        self._busy = False
        self.progress.stop()
        on_ok(res)

    def _current_box(self) -> tuple[float, float, float, float]:
        return layout_geometry.clamp_position(
            self.left_pct_var.get() / 100.0,
            self.top_pct_var.get() / 100.0,
            self.width_pct_var.get() / 100.0,
            self.height_pct_var.get() / 100.0,
        )

    def _set_box_vars(self, left: float, top: float, width: float, height: float) -> None:
        l, t, w, h = layout_geometry.clamp_position(left, top, width, height)
        self.left_pct_var.set(round(l * 100.0, 1))
        self.top_pct_var.set(round(t * 100.0, 1))
        self.width_pct_var.set(round(w * 100.0, 1))
        self.height_pct_var.set(round(h * 100.0, 1))

    def _avatar_aspect_hw(self) -> float | None:
        if not self.deck_info:
            return None
        iw, ih = self._avatar_img_size
        if iw <= 0 or ih <= 0 or self.deck_info.height_pt <= 0:
            return None
        return (self.deck_info.width_pt / self.deck_info.height_pt) * (ih / float(iw))

    # ------------------------------------------------------------------ Interactive Canvas Drag & Resize
    def _box_canvas_coords(self) -> tuple[float, float, float, float]:
        ox, oy, rw, rh = self._slide_rect
        l, t, w, h = self._current_box()
        return ox + l * rw, oy + t * rh, ox + (l + w) * rw, oy + (t + h) * rh

    def _hit_test(self, x: float, y: float) -> str | None:
        if self.current_slide is None or not self.thumb_paths:
            return None
        x0, y0, x1, y1 = self._box_canvas_coords()
        pad = 9
        corners = {
            "nw": (x0, y0),
            "ne": (x1, y0),
            "sw": (x0, y1),
            "se": (x1, y1),
        }
        for name, (cx, cy) in corners.items():
            if abs(x - cx) <= pad and abs(y - cy) <= pad:
                return name
        if x0 <= x <= x1 and y0 <= y <= y1:
            return "move"
        return None

    def _on_canvas_hover(self, event) -> None:
        if self._drag_state is not None:
            return
        hit = self._hit_test(event.x, event.y)
        cursor_map = {
            "move": "fleur",
            "nw": "size_nw_se",
            "se": "size_nw_se",
            "ne": "size_ne_sw",
            "sw": "size_ne_sw",
        }
        self.canvas.configure(cursor=cursor_map.get(hit, "arrow"))

    def _on_canvas_press(self, event) -> None:
        hit = self._hit_test(event.x, event.y)
        if not hit:
            return
        self._drag_state = {
            "mode": hit,
            "start_x": event.x,
            "start_y": event.y,
            "orig_box": self._current_box(),
        }

    def _on_canvas_drag(self, event) -> None:
        if not self._drag_state:
            return
        _ox, _oy, rw, rh = self._slide_rect
        if rw <= 1 or rh <= 1:
            return
        dx_pct = (event.x - self._drag_state["start_x"]) / rw
        dy_pct = (event.y - self._drag_state["start_y"]) / rh
        orig = self._drag_state["orig_box"]

        if self._drag_state["mode"] == "move":
            nl, nt, nw, nh = layout_geometry.clamp_position(orig[0] + dx_pct, orig[1] + dy_pct, orig[2], orig[3])
            # Gentle snap when near horizontal center
            if abs((nl + nw / 2.0) - 0.5) < 0.015:
                nl = round(0.5 - nw / 2.0, 4)
        else:
            nl, nt, nw, nh = layout_geometry.resize_from_handle(
                self._drag_state["mode"],
                orig,
                dx_pct,
                dy_pct,
                lock_aspect=bool(self.lock_aspect_var.get()),
                aspect_ratio_hw=self._avatar_aspect_hw(),
            )
        self._set_box_vars(nl, nt, nw, nh)
        self._sync_current_form_to_data()
        self._redraw_preview(show_guides=(self._drag_state["mode"] == "move"))

    def _on_canvas_release(self, _event) -> None:
        if self._drag_state:
            self._drag_state = None
            self._redraw_preview(show_guides=False)
            if self.proj_dir and self.proj_data:
                project_store.save_project(self.proj_dir, self.proj_data)

    # ------------------------------------------------------------------ Project & Slides
    def on_open_pptx(self) -> None:
        file_path = filedialog.askopenfilename(
            title="选择已有 PowerPoint 课件",
            initialdir=self._safe_initialdir("input"),
            filetypes=[("PowerPoint 文件", "*.pptx")],
        )
        if not file_path:
            return

        def task():
            proj_dir, data, info = project_store.create_or_load_project(Path(file_path))
            thumbs = ppt_inspector.export_thumbnails(Path(data["source_copy"]), proj_dir / "thumbs")
            return proj_dir, data, info, thumbs

        def done(res):
            self.proj_dir, self.proj_data, self.deck_info, self.thumb_paths = res
            self.audio_cache = audio_manager.PageAudioCache(self.proj_dir)
            self.mode_var.set(self.proj_data.get("mode", "auto"))
            self.deck_subtitle_var.set(
                f"当前课件：《{self.proj_data['name']}》 · 共 {self.deck_info.slide_count} 页 · 原文件保持只读"
            )
            self._refresh_avatar_label()
            self._populate_tree()
            self._refresh_enabled_count()
            if self.deck_info.slides:
                first_id = str(self.deck_info.slides[0].slide)
                self.tree.selection_set(first_id)
                self.tree.focus(first_id)
                self.on_select_slide()
            self._set_status(f"已导入课件《{self.proj_data['name']}》（共 {self.deck_info.slide_count} 页）。")

        self._run_bg("正在导入课件并生成高清缩略图...", task, done)

    def _populate_tree(self) -> None:
        for item in self.tree.get_children():
            self.tree.delete(item)
        if not self.deck_info or not self.proj_data:
            return
        for idx, s in enumerate(self.deck_info.slides):
            cfg = self.proj_data["pages"][str(s.slide)]
            en = "√" if cfg.get("enabled") else ""
            aud_path = cfg.get("audio_path", "")
            aud = "就绪" if aud_path and Path(aud_path).is_file() else "—"
            anim = "保留" if s.has_existing_animations else ""
            tag = "even" if idx % 2 == 0 else "odd"
            self.tree.insert("", "end", iid=str(s.slide), values=(s.slide, s.title, en, aud, anim), tags=(tag,))

    def _update_tree_row(self, slide: int) -> None:
        if not self.proj_data or not self.deck_info:
            return
        cfg = self.proj_data["pages"][str(slide)]
        s = self.deck_info.slides[slide - 1]
        en = "√" if cfg.get("enabled") else ""
        aud_path = cfg.get("audio_path", "")
        aud = "就绪" if aud_path and Path(aud_path).is_file() else "—"
        anim = "保留" if s.has_existing_animations else ""
        self.tree.item(str(slide), values=(s.slide, s.title, en, aud, anim))
        self._refresh_enabled_count()

    def _current_engine(self) -> str:
        return tts_service.LABEL_TO_ENGINE.get(self.engine_label_var.get(), tts_service.default_engine())

    @staticmethod
    def _voice_display_for(voices: tuple[tuple[str, str], ...], code: str | None) -> str:
        for c, desc in voices:
            if c == code:
                return f"{c} · {desc}"
        c, desc = voices[0]
        return f"{c} · {desc}"

    def _voice_code_from_ui(self) -> str:
        val = self.voice_var.get().strip()
        return val.split(" · ")[0].strip() if " · " in val else val

    def _apply_engine_ui_state(self, engine: str, selected_voice: str | None = None) -> None:
        if engine == "qwen3":
            choices = tuple(f"{c} · {desc}" for c, desc in tts_service.QWEN3_VOICES)
            self.voice_combo.configure(values=choices)
            code = selected_voice if selected_voice in tts_service.QWEN3_VOICE_CODES else tts_service.QWEN3_VOICE_CODES[0]
            self.voice_var.set(self._voice_display_for(tts_service.QWEN3_VOICES, code))
            self.model_combo.configure(state="disabled")
            self.style_entry.configure(state="disabled")
        elif engine == "edge":
            choices = tuple(f"{c} · {desc}" for c, desc in tts_service.EDGE_VOICES)
            self.voice_combo.configure(values=choices)
            code = selected_voice if selected_voice in tts_service.EDGE_VOICE_CODES else tts_service.EDGE_VOICE_CODES[0]
            self.voice_var.set(self._voice_display_for(tts_service.EDGE_VOICES, code))
            self.model_combo.configure(state="disabled")
            self.style_entry.configure(state="disabled")
        elif engine == "windows":
            self.voice_combo.configure(values=tts_service.WINDOWS_VOICES)
            v = selected_voice if selected_voice in tts_service.WINDOWS_VOICES else tts_service.WINDOWS_VOICES[0]
            self.voice_var.set(v)
            self.model_combo.configure(state="disabled")
            self.style_entry.configure(state="disabled")
        else:
            self.voice_combo.configure(values=tts_service.AVAILABLE_VOICES)
            v = selected_voice if selected_voice in tts_service.AVAILABLE_VOICES else tts_service.AVAILABLE_VOICES[0]
            self.voice_var.set(v)
            self.model_combo.configure(state="readonly")
            self.style_entry.configure(state="normal")

    def _on_engine_changed(self, _event=None) -> None:
        engine = self._current_engine()
        self._apply_engine_ui_state(engine)
        self._sync_current_form_to_data()

    def _sync_current_form_to_data(self) -> None:
        if self.current_slide is None or not self.proj_data:
            return
        cfg = self.proj_data["pages"][str(self.current_slide)]
        cfg["enabled"] = bool(self.enabled_var.get())
        cfg["script"] = self.script_text.get("1.0", "end-1c")
        cfg["engine"] = self._current_engine()
        cfg["model"] = self.model_var.get()
        cfg["voice"] = self._voice_code_from_ui()
        cfg["style"] = self.style_var.get()
        cfg["pace"] = self.pace_var.get()
        l, t, w, h = self._current_box()
        cfg["left_pct"], cfg["top_pct"], cfg["width_pct"], cfg["height_pct"] = l, t, w, h
        self.proj_data["mode"] = self.mode_var.get()

    def _on_script_edited(self, _event=None) -> None:
        self._sync_current_form_to_data()
        txt = self.script_text.get("1.0", "end-1c").strip()
        self.char_count_var.set(f"{len(txt)} 字")

    def on_select_slide(self, _event=None) -> None:
        sel = self.tree.selection()
        if not sel or not self.proj_data:
            return
        self._sync_current_form_to_data()
        slide = int(sel[0])
        self.current_slide = slide
        cfg = self.proj_data["pages"][str(slide)]
        s = self.deck_info.slides[slide - 1]
        self.preview_title_var.set(f"第 {slide} 页 · {s.title[:28]}")
        self.enabled_var.set(bool(cfg.get("enabled")))
        self.script_text.delete("1.0", "end")
        self.script_text.insert("1.0", cfg.get("script", ""))
        self.char_count_var.set(f"{len(cfg.get('script', '').strip())} 字")
        engine = cfg.get("engine", tts_service.default_engine())
        if engine not in tts_service.ENGINE_LABELS:
            engine = tts_service.default_engine()
        self.engine_label_var.set(tts_service.ENGINE_LABELS[engine])
        self.model_var.set(cfg.get("model", tts_service.MODELS[0]))
        self.style_var.set(cfg.get("style", "温和、亲切、清晰"))
        self.pace_var.set(cfg.get("pace", "适中"))
        self._apply_engine_ui_state(engine, cfg.get("voice"))
        self._set_box_vars(
            cfg.get("left_pct", 0.81),
            cfg.get("top_pct", 0.60),
            cfg.get("width_pct", 0.16),
            cfg.get("height_pct", 0.35),
        )
        self._refresh_audio_label()
        self._redraw_preview()

    def _refresh_audio_label(self) -> None:
        if self.current_slide is None or not self.proj_data:
            self.audio_info_var.set("当前页音频：未配置")
            self.audio_badge.configure(bg="#F8FAFC", fg=PALETTE["muted"])
            return
        aud = self.proj_data["pages"][str(self.current_slide)].get("audio_path", "")
        if aud and Path(aud).is_file():
            try:
                dur = audio_manager.validate_wav_bytes(Path(aud).read_bytes())
                self.audio_info_var.set(f"已就绪：{Path(aud).name}（{dur:.2f} 秒）")
            except Exception:
                self.audio_info_var.set(f"已就绪：{Path(aud).name}")
            self.audio_badge.configure(bg=PALETTE["success_bg"], fg=PALETTE["success"])
        else:
            self.audio_info_var.set("当前页音频：未配置（可点击生成或导入本地音频）")
            self.audio_badge.configure(bg="#F8FAFC", fg=PALETTE["muted"])

    def on_enabled_toggled(self) -> None:
        self._sync_current_form_to_data()
        if self.current_slide is not None:
            self._update_tree_row(self.current_slide)

    def _apply_preset(self, l: float, t: float, w: float, h: float) -> None:
        if self.lock_aspect_var.get() and self.deck_info:
            hw = self._avatar_aspect_hw()
            if hw:
                h = min(1.0 - t, w * hw)
        self._set_box_vars(l, t, w, h)
        self._sync_current_form_to_data()
        self._redraw_preview()

    def on_numeric_layout_changed(self) -> None:
        try:
            l = self.left_pct_var.get() / 100.0
            t = self.top_pct_var.get() / 100.0
            w = self.width_pct_var.get() / 100.0
            h = self.height_pct_var.get() / 100.0
        except tk.TclError:
            return
        self._set_box_vars(l, t, w, h)
        self._sync_current_form_to_data()
        self._redraw_preview()

    def on_apply_layout_all(self) -> None:
        if not self.proj_data or not self.proj_dir:
            return
        l, t, w, h = self._current_box()
        applied = 0
        for cfg in self.proj_data["pages"].values():
            if cfg.get("enabled"):
                cfg["left_pct"], cfg["top_pct"], cfg["width_pct"], cfg["height_pct"] = l, t, w, h
                applied += 1
        project_store.save_project(self.proj_dir, self.proj_data)
        self._set_status(f"已将当前位置 ({l*100:.1f}%, {t*100:.1f}%) 应用到 {applied} 个已启用页面。")

    def _redraw_preview(self, *, show_guides: bool = False) -> None:
        self.canvas.delete("all")
        if self.current_slide is None or not self.thumb_paths:
            cw = max(200, self.canvas.winfo_width())
            ch = max(120, self.canvas.winfo_height())
            self.canvas.create_text(
                cw // 2, ch // 2,
                text="请先从左上角导入 PPTX 课件\n导入后可在此处直接用鼠标拖动或缩放数字人位置",
                fill="#94A3B8", font=("Microsoft YaHei UI", 10), justify="center",
            )
            return
        idx = self.current_slide - 1
        if not (0 <= idx < len(self.thumb_paths)):
            return
        cw = max(200, self.canvas.winfo_width())
        ch = max(120, self.canvas.winfo_height())
        thumb = Image.open(self.thumb_paths[idx]).convert("RGBA")
        scale = min(cw / thumb.width, ch / thumb.height)
        rw, rh = max(1, int(thumb.width * scale)), max(1, int(thumb.height * scale))
        ox, oy = (cw - rw) // 2, (ch - rh) // 2
        self._slide_rect = (ox, oy, rw, rh)

        resized = thumb.resize((rw, rh), Image.Resampling.LANCZOS)
        self._preview_photo = ImageTk.PhotoImage(resized)
        self.canvas.create_image(ox, oy, anchor="nw", image=self._preview_photo)

        if show_guides:
            self.canvas.create_line(ox + rw / 2, oy, ox + rw / 2, oy + rh, fill="#38BDF8", dash=(4, 4), width=1)
            self.canvas.create_line(ox, oy + rh / 2, ox + rw, oy + rh / 2, fill="#38BDF8", dash=(4, 4), width=1)

        l, t, w, h = self._current_box()
        x0, y0 = ox + l * rw, oy + t * rh
        bw, bh = max(16, int(w * rw)), max(16, int(h * rh))
        x1, y1 = x0 + bw, y0 + bh

        idle_path = Path(self.proj_data["avatar_idle"]) if self.proj_data else project_store.DEFAULT_IDLE
        if idle_path.is_file():
            av_img = Image.open(idle_path).convert("RGBA").resize((bw, bh), Image.Resampling.LANCZOS)
            self._avatar_photo = ImageTk.PhotoImage(av_img)
            self.canvas.create_image(x0, y0, anchor="nw", image=self._avatar_photo)

        self.canvas.create_rectangle(x0, y0, x1, y1, outline="#2563EB", width=2, dash=(5, 3))
        for hx, hy in ((x0, y0), (x1, y0), (x0, y1), (x1, y1)):
            self.canvas.create_rectangle(hx - 5, hy - 5, hx + 5, hy + 5, fill="#FFFFFF", outline="#2563EB", width=2)

    # ------------------------------------------------------------------ Actions
    def on_save_project(self) -> None:
        if not self.proj_dir or not self.proj_data:
            return
        self._sync_current_form_to_data()
        project_store.save_project(self.proj_dir, self.proj_data)
        self._set_status("项目配置已保存。")

    def on_config_key(self) -> None:
        dlg = tk.Toplevel(self)
        dlg.title("管理员设置 - Gemini API Key")
        dlg.geometry("480x175")
        dlg.configure(bg=PALETTE["card"])
        dlg.transient(self)
        dlg.grab_set()
        tk.Label(
            dlg, text="请输入 Gemini API Key（通过本机 Windows DPAPI 加密保存，不写入源码或日志）：",
            bg=PALETTE["card"], fg=PALETTE["text"], font=FONT_BODY, wraplength=440, justify="left",
        ).pack(anchor="w", padx=16, pady=(16, 8))
        entry = ttk.Entry(dlg, show="*", width=56)
        entry.pack(fill="x", padx=16, pady=4)

        def save_it():
            val = entry.get().strip()
            if not val:
                messagebox.showwarning("提示", "Key 不能为空。", parent=dlg)
                return
            try:
                tts_service.save_api_key(val)
                self._refresh_key_badge()
                self._set_status("Gemini API Key 已安全加密保存。")
                dlg.destroy()
            except Exception as exc:
                messagebox.showerror("保存失败", str(exc), parent=dlg)

        ttk.Button(dlg, text="加密保存", style="Primary.TButton", command=save_it).pack(pady=14)

    def on_change_avatar(self) -> None:
        if not self.proj_dir or not self.proj_data:
            messagebox.showinfo("提示", "请先打开或导入一份 PPTX 课件。")
            return
        gif_file = filedialog.askopenfilename(
            title="选择全局数字人 GIF（带透明背景）",
            initialdir=str(ROOT_DIR / "assets"),
            filetypes=[("GIF 动画", "*.gif")],
        )
        if not gif_file:
            return
        idle_out = self.proj_dir / "custom_avatar_idle.png"
        project_store.extract_idle_frame(Path(gif_file), idle_out)
        self.proj_data["avatar_gif"] = str(Path(gif_file).resolve())
        self.proj_data["avatar_idle"] = str(idle_out.resolve())
        project_store.save_project(self.proj_dir, self.proj_data)
        self._refresh_avatar_label()
        self._redraw_preview()
        self._set_status("已更换全局数字人 GIF，并自动提取第 0 帧静态待机图。")

    def on_generate_current(self) -> None:
        if self.current_slide is None or not self.proj_data or not self.audio_cache:
            return
        self._sync_current_form_to_data()
        slide = self.current_slide
        cfg = self.proj_data["pages"][str(slide)]
        script = cfg["script"].strip()
        if not script:
            messagebox.showwarning("提示", "请先输入当前页的数字人台词。")
            return
        project_store.save_project(self.proj_dir, self.proj_data)
        engine = cfg.get("engine", tts_service.default_engine())
        cached = self.audio_cache.get_cached_tts(slide, script, cfg["voice"], cfg["style"], cfg["pace"], cfg["model"], engine=engine)
        if cached:
            cfg["audio_path"] = str(cached)
            cfg["enabled"] = True
            self.enabled_var.set(True)
            project_store.save_project(self.proj_dir, self.proj_data)
            self._update_tree_row(slide)
            self._refresh_audio_label()
            self._set_status(f"第 {slide} 页命中本地语音缓存，无需重复生成。")
            return

        def task():
            wav = tts_service.synthesize_speech(
                script, engine=engine, voice=cfg["voice"], style=cfg["style"], pace=cfg["pace"], model=cfg["model"],
            )
            return self.audio_cache.put_tts(slide, script, cfg["voice"], cfg["style"], cfg["pace"], cfg["model"], wav, engine=engine)

        def done(wav_path: Path):
            cfg["audio_path"] = str(wav_path)
            cfg["enabled"] = True
            self.enabled_var.set(True)
            project_store.save_project(self.proj_dir, self.proj_data)
            self._update_tree_row(slide)
            self._refresh_audio_label()
            self._set_status(f"第 {slide} 页语音已生成并缓存。")

        eng_name = tts_service.ENGINE_LABELS.get(engine, engine)
        self._run_bg(f"正在使用 {eng_name} 生成第 {slide} 页语音...", task, done)

    def on_generate_all(self) -> None:
        if not self.proj_data or not self.audio_cache:
            return
        self._sync_current_form_to_data()
        targets = [
            int(k) for k, cfg in self.proj_data["pages"].items()
            if cfg.get("enabled") and cfg.get("script", "").strip()
        ]
        if not targets:
            messagebox.showinfo("提示", "没有已启用且填写了台词的页面。")
            return
        project_store.save_project(self.proj_dir, self.proj_data)

        def task():
            generated, cached_hits = 0, 0
            for s in targets:
                cfg = self.proj_data["pages"][str(s)]
                eng = cfg.get("engine", tts_service.default_engine())
                hit = self.audio_cache.get_cached_tts(s, cfg["script"], cfg["voice"], cfg["style"], cfg["pace"], cfg["model"], engine=eng)
                if hit:
                    cfg["audio_path"] = str(hit)
                    cached_hits += 1
                    continue
                wav = tts_service.synthesize_speech(
                    cfg["script"], engine=eng, voice=cfg["voice"], style=cfg["style"], pace=cfg["pace"], model=cfg["model"],
                )
                wav_path = self.audio_cache.put_tts(s, cfg["script"], cfg["voice"], cfg["style"], cfg["pace"], cfg["model"], wav, engine=eng)
                cfg["audio_path"] = str(wav_path)
                generated += 1
            return generated, cached_hits

        def done(counts):
            generated, cached_hits = counts
            project_store.save_project(self.proj_dir, self.proj_data)
            for s in targets:
                self._update_tree_row(s)
            self._refresh_audio_label()
            self._set_status(f"批量语音完成：新生成 {generated} 页，命中缓存 {cached_hits} 页。")

        self._run_bg(f"正在批量检查/生成 {len(targets)} 页语音...", task, done)

    def on_import_local_audio(self) -> None:
        if self.current_slide is None or not self.proj_data or not self.audio_cache:
            return
        path = filedialog.askopenfilename(
            title="导入本地音频或提取视频音轨",
            filetypes=[("音频 / 视频文件", "*.wav *.mp3 *.m4a *.ogg *.flac *.aac *.wma *.mp4 *.mov *.mkv"), ("所有文件", "*.*")],
        )
        if not path:
            return
        try:
            dest = self.audio_cache.import_local_audio(self.current_slide, Path(path))
            cfg = self.proj_data["pages"][str(self.current_slide)]
            cfg["audio_path"] = str(dest)
            cfg["enabled"] = True
            self.enabled_var.set(True)
            project_store.save_project(self.proj_dir, self.proj_data)
            self._update_tree_row(self.current_slide)
            self._refresh_audio_label()
            self._set_status(f"第 {self.current_slide} 页已导入本地音频：{dest.name}")
        except Exception as exc:
            messagebox.showerror("导入失败", str(exc))

    def on_play_audio(self) -> None:
        if self.current_slide is None or not self.proj_data:
            return
        aud = self.proj_data["pages"][str(self.current_slide)].get("audio_path", "")
        if not aud or not Path(aud).is_file():
            messagebox.showinfo("提示", "当前页尚未生成或导入音频。")
            return
        try:
            audio_manager.play_wav_async(Path(aud))
            self._set_status(f"正在试听第 {self.current_slide} 页音频...")
        except Exception as exc:
            messagebox.showerror("试听失败", str(exc))

    def on_export_pptx(self) -> None:
        if not self.proj_dir or not self.proj_data or not self.deck_info:
            messagebox.showinfo("提示", "请先打开并配置一份 PPTX 课件。")
            return
        self._sync_current_form_to_data()
        project_store.save_project(self.proj_dir, self.proj_data)
        default_name = f"{self.proj_data['name']}_数字人版.pptx"
        dest = filedialog.asksaveasfilename(
            title="导出新 PPTX 文件",
            initialdir=self._safe_initialdir("output", create=True),
            initialfile=default_name,
            defaultextension=".pptx",
            filetypes=[("PowerPoint 文件", "*.pptx")],
        )
        if not dest:
            return

        def task():
            return project_store.export_presentation(self.proj_dir, self.proj_data, self.deck_info, Path(dest))

        def done(out_path: Path):
            self.last_exported_pptx = out_path
            self._set_status(f"导出成功：{out_path}")
            messagebox.showinfo("导出完成", f"已生成可放映 PPTX：\n{out_path}")

        self._run_bg("正在将数字人 GIF 与语音安全嵌入 PPTX...", task, done)

    def on_start_smart_show(self) -> None:
        target = self.last_exported_pptx
        if not target or not target.is_file():
            chosen = filedialog.askopenfilename(
                title="选择要智能放映的 PPTX",
                initialdir=self._safe_initialdir("output"),
                filetypes=[("PowerPoint 文件", "*.pptx")],
            )
            if not chosen:
                return
            target = Path(chosen)
        smart_ps1 = SRC_DIR / "smart_slideshow.ps1"
        subprocess.Popen([
            "powershell.exe", "-NoProfile", "-ExecutionPolicy", "Bypass",
            "-File", str(smart_ps1), "-Presentation", str(target.resolve()),
        ])
        self._set_status(f"已启动智能放映：{target.name}（按 ESC 结束放映）。")

    def on_open_exported(self) -> None:
        target = self.last_exported_pptx
        if not target or not target.is_file():
            chosen = filedialog.askopenfilename(
                title="选择要打开的 PPTX",
                initialdir=self._safe_initialdir("output"),
                filetypes=[("PowerPoint 文件", "*.pptx")],
            )
            if not chosen:
                return
            target = Path(chosen)
        os.startfile(str(target.resolve()))


def main() -> None:
    app = DigitalHumanApp()

    def _on_close() -> None:
        tts_service.shutdown_qwen3_worker()
        app.destroy()

    app.protocol("WM_DELETE_WINDOW", _on_close)
    app.mainloop()


if __name__ == "__main__":
    main()
