"""主应用窗口：组装所有 UI 面板，驱动转换流程 —— 央视风格"""

import os
import threading
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

from . import DND_AVAILABLE, DND_FILES, TkinterDnD
from .theme import (PRIMARY, PRIMARY_DK, PRIMARY_LT, ACCENT_GOLD, ACCENT_GOLD_DK,
                    HEADER_BG, LIGHT_BG, WHITE, CARD_BG, BORDER, TXT, TXT_LIGHT,
                    TXT_WHITE, LIST_HDR, OCR_COLOR, OCR_BG, SEL_BG,
                    STATUS_OK, STATUS_ERR, STATUS_RUN, styled_btn)
from .converter import convert_to_word, convert_to_word_ocr, parse_page_range
from .deps import ensure_ocr_deps
from .ui_annotation import AnnotationMixin
from .ui_preview import PreviewMixin
from .ui_file_list import FileListMixin
from .ui_output_opts import OutputOptsMixin

_BaseClass = TkinterDnD.Tk if DND_AVAILABLE else tk.Tk


class App(AnnotationMixin, PreviewMixin, FileListMixin, OutputOptsMixin,
          _BaseClass):

    def __init__(self):
        super().__init__()
        self.title("PDF转Word —— 智能文档转换工具 | 作者VX：XNHSDJ")
        self.geometry("1100x750")
        self.minsize(900, 620)
        self.configure(bg=LIGHT_BG)

        self._files       = []
        self._sel_set     = set()
        self._last_sel    = None
        self._row_widgets = []
        self._out_dir     = tk.StringVar()
        self._page_sel    = tk.StringVar(value="all")
        self._page_txt    = tk.StringVar()
        self._use_ocr     = tk.BooleanVar(value=False)
        self._ocr_lang    = tk.StringVar(value="chi_sim+eng")
        self._busy        = False

        # 预览状态
        self._prev_path   = None
        self._prev_pages  = []
        self._prev_page   = 0
        self._prev_zoom   = 1.0
        self._prev_token  = 0
        self._img_render  = None

        # 输出选项
        self._out_fmt           = tk.StringVar(value="docx")
        self._opt_toc           = tk.BooleanVar(value=True)
        self._opt_header_footer = tk.BooleanVar(value=True)
        self._opt_list_para     = tk.BooleanVar(value=True)

        # 标注工具状态
        self._anno_tool       = tk.StringVar(value="none")
        self._anno_regions    = []
        self._anno_drag_start = None
        self._anno_rubber     = None
        self._anno_sel_idx    = None
        self._anno_show       = tk.BooleanVar(value=True)

        default_out = os.path.join(
            os.path.expanduser("~"), "Documents", "PDFPai")
        self._out_dir.set(default_out)

        self._build_ui()

    # ══════════════════════════════════════════════════════════════════════════
    # UI 骨架
    # ══════════════════════════════════════════════════════════════════════════
    def _build_ui(self):
        # ── 深色标题栏 ──
        title_bar = tk.Frame(self, bg=HEADER_BG, height=52)
        title_bar.pack(fill="x")
        title_bar.pack_propagate(False)

        # Canvas 自绘 App 图标（PDF→W 转换图标）
        self._draw_app_icon(title_bar)

        tk.Label(title_bar, text="PDF转Word",
                 font=("微软雅黑", 14, "bold"),
                 bg=HEADER_BG, fg=TXT_WHITE).pack(side="left")
        tk.Label(title_bar, text="智能文档转换",
                 font=("微软雅黑", 9),
                 bg=HEADER_BG, fg="#8899AA").pack(side="left", padx=(8, 0))

        # 金色点缀线
        gold_line = tk.Frame(self, bg=ACCENT_GOLD, height=2)
        gold_line.pack(fill="x")

        # ── 底部版权栏（先 pack，保证在最底部）──
        footer = tk.Frame(self, bg=HEADER_BG, height=26)
        footer.pack(side="bottom", fill="x")
        footer.pack_propagate(False)
        tk.Label(footer,
                 text="Copyright \u00a9 2025 Hum0ro. All rights reserved.  |  作者VX：XNHSDJ  |  v1.0",
                 font=("微软雅黑", 8), bg=HEADER_BG, fg="#667788",
                 ).pack(expand=True)

        pane = tk.PanedWindow(self, orient="horizontal",
                              bg=LIGHT_BG, sashwidth=4, sashrelief="flat")
        pane.pack(fill="both", expand=True)

        left  = tk.Frame(pane, bg=WHITE)
        right = tk.Frame(pane, bg=WHITE)
        pane.add(left,  minsize=400, width=440)
        pane.add(right, minsize=400)

        self._build_left(left)
        self._build_right(right)

        # ── 标注工具栏 ──
        def _place_anno_bar(event=None):
            self.update_idletasks()
            rx = right.winfo_x()
            rw = right.winfo_width()
            bh = title_bar.winfo_height()
            self._anno_bar.place(x=rx, y=0, width=rw, height=bh)
            self._anno_bar.lift()

        self._anno_bar = tk.Frame(self, bg=HEADER_BG)
        self._build_anno_bar(self._anno_bar)

        self.bind("<Configure>", lambda e: _place_anno_bar())
        self.after(100, _place_anno_bar)

    # ── Canvas 自绘 App 图标 ─────────────────────────────────────────────────
    def _draw_app_icon(self, parent):
        S = 36
        cv = tk.Canvas(parent, width=S, height=S, bg=HEADER_BG,
                       highlightthickness=0)
        cv.pack(side="left", padx=(16, 10), pady=8)

        # 红色圆角背景
        r = 6
        cv.create_arc(0, 0, 2 * r, 2 * r, start=90, extent=90,
                       fill=PRIMARY, outline="")
        cv.create_arc(S - 2 * r, 0, S, 2 * r, start=0, extent=90,
                       fill=PRIMARY, outline="")
        cv.create_arc(0, S - 2 * r, 2 * r, S, start=180, extent=90,
                       fill=PRIMARY, outline="")
        cv.create_arc(S - 2 * r, S - 2 * r, S, S, start=270, extent=90,
                       fill=PRIMARY, outline="")
        cv.create_rectangle(r, 0, S - r, S, fill=PRIMARY, outline="")
        cv.create_rectangle(0, r, S, S - r, fill=PRIMARY, outline="")

        # 左侧：PDF 文档轮廓（白色小文档）
        dx, dy = 4, 7
        dw, dh = 12, 22
        cv.create_rectangle(dx, dy, dx + dw, dy + dh,
                            fill=WHITE, outline="", width=0)
        # 折角
        cv.create_polygon(dx + dw - 4, dy, dx + dw, dy + 4,
                          dx + dw, dy, fill="#E8C4C0", outline="")
        # "P" 字母
        cv.create_text(dx + dw // 2, dy + dh // 2 + 1,
                       text="P", font=("Arial", 7, "bold"),
                       fill=PRIMARY, anchor="center")

        # 中间金色箭头 →
        ax = 19; ay = S // 2
        cv.create_line(ax, ay, ax + 6, ay, fill=ACCENT_GOLD, width=2)
        cv.create_line(ax + 4, ay - 3, ax + 7, ay, ax + 4, ay + 3,
                       fill=ACCENT_GOLD, width=1.5)

        # 右侧：Word 文档（白色小文档）
        wx = 26; wy = 7; ww = 12; wh = 22
        # 简化绘制 - 不超出背景
        rw = min(wx + ww, S - 1)
        cv.create_rectangle(wx - 2, wy, rw - 2, wy + wh,
                            fill=WHITE, outline="", width=0)
        cv.create_text(wx + ww // 2 - 2, wy + wh // 2 + 1,
                       text="W", font=("Arial", 7, "bold"),
                       fill="#2B579A", anchor="center")

    # ══════════════════════════════════════════════════════════════════════════
    # 左面板
    # ══════════════════════════════════════════════════════════════════════════
    def _build_left(self, parent):
        toolbar = tk.Frame(parent, bg=WHITE, pady=10, padx=14)
        toolbar.pack(fill="x")
        self._btn_add   = styled_btn(toolbar, "添加",     self._add_files)
        self._btn_del   = styled_btn(toolbar, "删除",     self._delete_selected)
        self._btn_start = styled_btn(toolbar, "开始转换",
                                     self._start_convert, width=8,
                                     color=PRIMARY)
        self._btn_opts  = styled_btn(toolbar, "输出选项",
                                     self._show_output_opts, width=8,
                                     color=ACCENT_GOLD)
        # "输出选项"按钮用金色 hover
        self._btn_opts.bind("<Enter>",
                            lambda e: self._btn_opts.config(bg=ACCENT_GOLD_DK))
        self._btn_opts.bind("<Leave>",
                            lambda e: self._btn_opts.config(bg=ACCENT_GOLD))
        for btn in (self._btn_add, self._btn_del,
                    self._btn_start, self._btn_opts):
            btn.pack(side="left", padx=(0, 8))

        tk.Frame(parent, bg=BORDER, height=1).pack(fill="x")

        # ── 文件列表区 ──
        list_area = tk.Frame(parent, bg=WHITE)
        list_area.pack(fill="both", expand=True, padx=14)

        hdr = tk.Frame(list_area, bg=LIST_HDR, height=32)
        hdr.pack(fill="x"); hdr.pack_propagate(False)
        tk.Label(hdr, text="No.",  font=("微软雅黑", 9), bg=LIST_HDR,
                 fg=TXT_LIGHT, width=5).pack(side="left", padx=(10, 0))
        tk.Label(hdr, text="名称", font=("微软雅黑", 9), bg=LIST_HDR,
                 fg=TXT_LIGHT).pack(side="left", padx=(4, 0),
                                    expand=True, anchor="w")
        tk.Label(hdr, text="状态", font=("微软雅黑", 9), bg=LIST_HDR,
                 fg=TXT_LIGHT, width=8).pack(side="right", padx=(0, 10))

        list_body = tk.Frame(list_area, bg=WHITE, bd=1, relief="solid",
                             highlightthickness=1, highlightbackground=BORDER)
        list_body.pack(fill="both", expand=True)

        canvas = tk.Canvas(list_body, bg=WHITE, highlightthickness=0)
        vsb    = ttk.Scrollbar(list_body, orient="vertical",
                               command=canvas.yview)
        canvas.configure(yscrollcommand=vsb.set)
        vsb.pack(side="right", fill="y")
        canvas.pack(side="left", fill="both", expand=True)

        self._inner = tk.Frame(canvas, bg=WHITE)
        self._cwin  = canvas.create_window((0, 0), window=self._inner,
                                           anchor="nw")
        self._inner.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.bind(
            "<Configure>",
            lambda e: canvas.itemconfig(self._cwin, width=e.width))
        self._canvas = canvas

        # ── 拖拽放置区 ──
        self._drop_zone = tk.Label(
            list_area,
            text="📂  拖拽 PDF 文件到此处  （支持批量拖入）",
            font=("微软雅黑", 10), bg=PRIMARY_LT, fg=PRIMARY,
            relief="flat", bd=1, pady=10, cursor="hand2",
        )
        if DND_AVAILABLE:
            self._drop_zone.pack(fill="x", pady=(4, 0))
            for widget in (list_body, canvas, self._inner, self._drop_zone):
                widget.drop_target_register(DND_FILES)
                widget.dnd_bind("<<Drop>>", self._on_drop)

        # ── OCR 选项 ──
        tk.Frame(parent, bg=BORDER, height=1).pack(fill="x", padx=14,
                                                    pady=(8, 0))
        ocr_card = tk.Frame(parent, bg=OCR_BG, bd=1, relief="solid",
                            highlightthickness=1,
                            highlightbackground=PRIMARY_LT)
        ocr_card.pack(fill="x", padx=14, pady=(4, 0))

        ocr_hdr = tk.Frame(ocr_card, bg=OCR_BG, padx=10, pady=6)
        ocr_hdr.pack(fill="x")
        self._ocr_cb = tk.Checkbutton(
            ocr_hdr, text="启用 OCR 识别（扫描版PDF）",
            variable=self._use_ocr,
            font=("微软雅黑", 9, "bold"), bg=OCR_BG, fg=OCR_COLOR,
            activebackground=OCR_BG, selectcolor=OCR_BG,
            command=self._on_ocr_toggle, cursor="hand2",
        )
        self._ocr_cb.pack(side="left")
        tk.Label(ocr_hdr, text="ℹ 适用于图片型扫描版PDF",
                 font=("微软雅黑", 8), bg=OCR_BG,
                 fg="#B07070").pack(side="left", padx=(8, 0))

        self._ocr_lang_frame = tk.Frame(ocr_card, bg=OCR_BG,
                                        padx=10, pady=8)
        tk.Label(self._ocr_lang_frame, text="识别语言：",
                 font=("微软雅黑", 9), bg=OCR_BG,
                 fg=TXT).pack(side="left")
        lang_opts = [
            ("中文 + 英文", "chi_sim+eng"),
            ("仅英文",      "eng"),
            ("繁体中文",    "chi_tra+eng"),
            ("日文",        "jpn+eng"),
        ]
        self._lang_name2code = {k: v for k, v in lang_opts}
        self._lang_display   = {v: k for k, v in lang_opts}
        self._lang_combo = ttk.Combobox(
            self._ocr_lang_frame, textvariable=self._ocr_lang,
            values=[k for k, _ in lang_opts],
            state="readonly", width=12, font=("微软雅黑", 9),
        )
        self._lang_combo.pack(side="left", padx=(0, 4))
        self._ocr_lang.set("中文 + 英文")

        # ── 页面设置 ──
        tk.Frame(parent, bg=BORDER, height=1).pack(fill="x", padx=14)
        pg_frame = tk.Frame(parent, bg=WHITE, padx=14, pady=8)
        pg_frame.pack(fill="x")
        tk.Label(pg_frame, text="页面设置",
                 font=("微软雅黑", 9, "bold"), bg=WHITE,
                 fg=TXT).pack(anchor="w", pady=(0, 6))

        r1 = tk.Frame(pg_frame, bg=WHITE); r1.pack(fill="x", pady=2)
        tk.Radiobutton(
            r1, text="所有页面", variable=self._page_sel, value="all",
            font=("微软雅黑", 9), bg=WHITE, fg=TXT,
            activebackground=WHITE, selectcolor=WHITE,
            command=self._on_page_sel,
        ).pack(side="left")

        r2 = tk.Frame(pg_frame, bg=WHITE); r2.pack(fill="x", pady=2)
        tk.Radiobutton(
            r2, text="所选页面", variable=self._page_sel, value="custom",
            font=("微软雅黑", 9), bg=WHITE, fg=TXT,
            activebackground=WHITE, selectcolor=WHITE,
            command=self._on_page_sel,
        ).pack(side="left")
        self._page_entry = tk.Entry(
            r2, textvariable=self._page_txt, font=("微软雅黑", 9),
            width=18, bg="#F5F6F8", relief="flat", state="disabled",
        )
        self._page_entry.pack(side="left", padx=(8, 0), ipady=3)
        tk.Label(r2, text="如：1,3-5,8", font=("微软雅黑", 8),
                 bg=WHITE, fg=TXT_LIGHT).pack(side="left", padx=6)

        # ── 输出文件夹 ──
        tk.Frame(parent, bg=BORDER, height=1).pack(fill="x", padx=14)
        out_frame = tk.Frame(parent, bg=WHITE, padx=14, pady=10)
        out_frame.pack(fill="x")
        tk.Label(out_frame, text="输出文件夹",
                 font=("微软雅黑", 9, "bold"), bg=WHITE,
                 fg=TXT).pack(anchor="w", pady=(0, 6))

        dir_row = tk.Frame(out_frame, bg=WHITE); dir_row.pack(fill="x")
        tk.Entry(dir_row, textvariable=self._out_dir, font=("微软雅黑", 9),
                 bg="#F5F6F8", relief="flat", state="readonly",
                 readonlybackground="#F5F6F8", fg=TXT,
                 ).pack(side="left", fill="x", expand=True,
                        ipady=4, padx=(0, 6))

        btn_browse = self._make_folder_icon_btn(dir_row, self._browse_dir,
                                                style="closed")
        btn_browse.pack(side="left", padx=(0, 4))
        self._bind_tooltip(btn_browse, "选择文件夹",
                           "点击选择输出文件夹路径")

        btn_open = self._make_folder_icon_btn(dir_row, self._open_dir,
                                              style="open")
        btn_open.pack(side="left")
        self._bind_tooltip(btn_open, "打开文件夹",
                           "在文件管理器中打开输出文件夹")

    # ══════════════════════════════════════════════════════════════════════════
    # 右面板
    # ══════════════════════════════════════════════════════════════════════════
    def _build_right(self, parent):
        parent.configure(bg=CARD_BG)
        tk.Frame(parent, bg=BORDER, width=1).pack(side="left", fill="y")

        # 底部导航栏
        tk.Frame(parent, bg=BORDER, height=1).pack(side="bottom", fill="x")
        nav = tk.Frame(parent, bg=WHITE, height=46)
        nav.pack(fill="x", side="bottom")
        nav.pack_propagate(False)

        nav_center = tk.Frame(nav, bg=WHITE)
        nav_center.place(relx=0.5, rely=0.5, anchor="center")

        def _circle_btn(parent, symbol, command, size=28):
            cv = tk.Canvas(parent, width=size, height=size,
                           bg=WHITE, highlightthickness=0, cursor="hand2")

            def _draw(outline=BORDER, fg_color=TXT, fill=WHITE):
                cv.delete("all")
                cv.create_oval(1, 1, size - 1, size - 1,
                               outline=outline, fill=fill, width=1)
                cv.create_text(size // 2, size // 2, text=symbol,
                               font=("Arial", 11), fill=fg_color,
                               anchor="center")
            _draw()

            def _enter(e): _draw(PRIMARY, PRIMARY, PRIMARY_LT)
            def _leave(e): _draw(BORDER,  TXT,     WHITE)
            def _press(e): _draw(PRIMARY_DK, WHITE, PRIMARY)
            def _release(e):
                _draw(PRIMARY, PRIMARY, PRIMARY_LT)
                command()

            cv.bind("<Enter>",           _enter)
            cv.bind("<Leave>",           _leave)
            cv.bind("<ButtonPress-1>",   _press)
            cv.bind("<ButtonRelease-1>", _release)
            return cv

        self._btn_prev = _circle_btn(nav_center, "\u2039",
                                     self._prev_page_cmd, size=30)
        self._btn_prev.pack(side="left", padx=(0, 4))
        self._btn_next = _circle_btn(nav_center, "\u203a",
                                     self._next_page_cmd, size=30)
        self._btn_next.pack(side="left", padx=(0, 8))

        self._page_var = tk.StringVar(value="1")
        page_entry_frame = tk.Frame(nav_center, bg=BORDER,
                                    highlightthickness=1,
                                    highlightbackground=BORDER)
        page_entry_frame.pack(side="left", padx=(0, 4))
        self._page_entry_nav = tk.Entry(
            page_entry_frame, textvariable=self._page_var,
            font=("微软雅黑", 9), width=3, justify="center",
            bg=WHITE, fg=TXT, relief="flat", bd=4,
        )
        self._page_entry_nav.pack()
        self._page_entry_nav.bind("<Return>",   self._on_page_entry_jump)
        self._page_entry_nav.bind("<FocusOut>", self._on_page_entry_jump)

        self._lbl_total = tk.Label(nav_center, text="/ —",
                                   font=("微软雅黑", 9), bg=WHITE,
                                   fg=TXT_LIGHT)
        self._lbl_total.pack(side="left", padx=(2, 12))

        self._btn_zoom_in = _circle_btn(nav_center, "+",
                                        self._zoom_in, size=30)
        self._btn_zoom_in.pack(side="left", padx=(0, 4))
        self._btn_zoom_out = _circle_btn(nav_center, "\u2212",
                                         self._zoom_out, size=30)
        self._btn_zoom_out.pack(side="left")

        self._lbl_zoom = tk.Label(nav_center, text="",
                                  font=("微软雅黑", 8), bg=WHITE,
                                  fg=TXT_LIGHT, width=5)
        self._lbl_zoom.pack(side="left", padx=(6, 0))

        # ── 预览区 ──
        self._preview_frame = tk.Frame(parent, bg=CARD_BG)
        self._preview_frame.pack(fill="both", expand=True)

        self._preview_placeholder = tk.Label(
            self._preview_frame,
            text="📄\n\n请添加 PDF 文件\n选中后此处显示预览",
            font=("微软雅黑", 12), bg=CARD_BG, fg="#BCC5D0",
            justify="center")
        self._preview_placeholder.place(relx=.5, rely=.45, anchor="center")

        self._prev_canvas = tk.Canvas(self._preview_frame, bg="#E8E8E8",
                                      highlightthickness=0)
        prev_vsb = ttk.Scrollbar(self._preview_frame, orient="vertical",
                                 command=self._prev_canvas.yview)
        prev_hsb = ttk.Scrollbar(self._preview_frame, orient="horizontal",
                                 command=self._prev_canvas.xview)
        self._prev_canvas.configure(yscrollcommand=prev_vsb.set,
                                    xscrollcommand=prev_hsb.set)
        prev_vsb.pack(side="right", fill="y")
        prev_hsb.pack(side="bottom", fill="x")
        self._prev_canvas.pack(side="left", fill="both", expand=True)

        self._prev_canvas.bind("<MouseWheel>",
                               self._on_preview_scroll)
        self._prev_canvas.bind("<Button-4>",
                               self._on_preview_scroll)
        self._prev_canvas.bind("<Button-5>",
                               self._on_preview_scroll)
        self._prev_canvas.bind("<Control-MouseWheel>",
                               self._on_preview_ctrl_scroll)
        self._prev_canvas.bind("<Control-Button-4>",
                               self._on_preview_ctrl_scroll)
        self._prev_canvas.bind("<Control-Button-5>",
                               self._on_preview_ctrl_scroll)
        self._prev_canvas.bind("<Configure>",
                               lambda e: self._redraw_preview())

        # 标注鼠标绑定
        self._prev_canvas.bind("<ButtonPress-1>",   self._anno_on_press)
        self._prev_canvas.bind("<B1-Motion>",       self._anno_on_drag)
        self._prev_canvas.bind("<ButtonRelease-1>", self._anno_on_release)

    # ── 文件夹图标按钮（红金配色）──────────────────────────────────────────────
    def _make_folder_icon_btn(self, parent, command, style="closed"):
        SIZE = 32
        cv = tk.Canvas(parent, width=SIZE, height=SIZE,
                       bg=LIST_HDR, highlightthickness=1,
                       highlightbackground=BORDER, cursor="hand2")

        def _draw(bg_c=LIST_HDR, border="#B0BAC5"):
            cv.delete("all")
            cv.configure(bg=bg_c)
            if style == "closed":
                cv.create_rectangle(4, 13, SIZE - 4, SIZE - 5,
                                    fill="#F5D6D0", outline=border, width=1)
                cv.create_rectangle(4, 10, 14, 14,
                                    fill="#F5D6D0", outline=border, width=1)
                cv.create_line(5, 13, 13, 13, fill="#F5D6D0", width=1)
                for yi in range(17, SIZE - 6, 4):
                    cv.create_line(7, yi, SIZE - 7, yi,
                                   fill=border, dash=(2, 2), width=1)
            else:
                cv.create_rectangle(5, 12, SIZE - 4, SIZE - 6,
                                    fill="#F5D6D0", outline=border, width=1)
                cv.create_rectangle(5, 9, 14, 13,
                                    fill="#F5D6D0", outline=border, width=1)
                cv.create_line(6, 12, 13, 12, fill="#F5D6D0", width=1)
                pts = [4, SIZE - 6, 4, 16, SIZE - 4, 13, SIZE - 4, SIZE - 6]
                cv.create_polygon(pts, fill="#FAE5E0", outline=border,
                                  width=1)
                for yi in range(18, SIZE - 6, 4):
                    cv.create_line(7, yi, SIZE - 7, yi,
                                   fill=border, dash=(2, 2), width=1)
        _draw()

        def _enter(e):   _draw(PRIMARY_LT, PRIMARY)
        def _leave(e):   _draw(LIST_HDR,   "#B0BAC5")
        def _press(e):   _draw("#F0C5BE",  PRIMARY_DK)
        def _release(e):
            _draw(PRIMARY_LT, PRIMARY)
            command()

        cv.bind("<Enter>",           _enter)
        cv.bind("<Leave>",           _leave)
        cv.bind("<ButtonPress-1>",   _press)
        cv.bind("<ButtonRelease-1>", _release)
        return cv

    # ══════════════════════════════════════════════════════════════════════════
    # 页面选择 / 文件夹操作
    # ══════════════════════════════════════════════════════════════════════════
    def _on_page_sel(self):
        self._page_entry.config(
            state="normal" if self._page_sel.get() == "custom"
            else "disabled")

    def _browse_dir(self):
        d = filedialog.askdirectory(title="选择输出文件夹")
        if d:
            self._out_dir.set(d)

    def _open_dir(self):
        d = self._out_dir.get()
        if os.path.isdir(d):
            if os.name == "nt":
                os.startfile(d)
            else:
                os.system(f'xdg-open "{d}"')
        else:
            messagebox.showinfo("提示", "输出文件夹不存在，请先选择。")

    # ══════════════════════════════════════════════════════════════════════════
    # 转换
    # ══════════════════════════════════════════════════════════════════════════
    def _start_convert(self):
        if self._busy:
            return
        if not self._files:
            messagebox.showwarning("提示", "请先添加 PDF 文件！")
            return

        out_dir = self._out_dir.get()
        os.makedirs(out_dir, exist_ok=True)

        page_range = None
        if self._page_sel.get() == "custom":
            txt = self._page_txt.get().strip()
            if not txt:
                messagebox.showwarning("提示",
                                       "请输入页码范围，如：1,3-5,8")
                return
            page_range = txt

        use_ocr  = self._use_ocr.get()
        if use_ocr and not ensure_ocr_deps():
            return
        ocr_lang = self._lang_name2code.get(
            self._ocr_lang.get(), "chi_sim+eng")
        out_fmt  = self._out_fmt.get()
        opt_toc  = self._opt_toc.get()
        opt_hf   = self._opt_header_footer.get()
        opt_lp   = self._opt_list_para.get()

        self._busy = True
        self._btn_start.config(
            state="disabled",
            text="OCR转换中…" if use_ocr else "转换中…")

        def run():
            for path, svar, out_var, err_var, type_var in self._files:
                if svar.get() == "完成":
                    continue
                status = "OCR中…" if use_ocr else "转换中…"
                self.after(0, lambda s=svar, st=status: s.set(st))
                try:
                    ext      = ".doc" if out_fmt == "doc" else ".docx"
                    out_name = (os.path.splitext(
                        os.path.basename(path))[0] + ext)
                    out_path = os.path.join(out_dir, out_name)

                    pr = None
                    if page_range:
                        from pypdf import PdfReader
                        total = len(PdfReader(path).pages)
                        pr = parse_page_range(page_range, total)

                    out_opts = {
                        "fmt":           out_fmt,
                        "detect_toc":    opt_toc,
                        "header_footer": opt_hf,
                        "list_para":     opt_lp,
                    }
                    if use_ocr:
                        convert_to_word_ocr(
                            path, out_path,
                            page_range=pr, lang=ocr_lang,
                            progress_cb=None)
                    else:
                        convert_to_word(path, out_path, pr,
                                        out_opts=out_opts)

                    self.after(0, lambda s=svar, op=out_path, ov=out_var: (
                        s.set("完成"), ov.set(op)))
                except Exception:
                    import traceback
                    err_msg = traceback.format_exc()
                    self.after(0, lambda s=svar, ev=err_var, em=err_msg: (
                        s.set("失败"), ev.set(em)))

            self.after(0, self._on_done)

        threading.Thread(target=run, daemon=True).start()

    def _on_done(self):
        self._busy = False
        self._btn_start.config(state="normal", text="开始转换")
        ok   = sum(1 for f in self._files if f[1].get() == "完成")
        fail = sum(1 for f in self._files if f[1].get() == "失败")
        fail_list = [
            (os.path.basename(f[0]), f[3].get())
            for f in self._files if f[1].get() == "失败"
        ]
        if fail == 0:
            messagebox.showinfo(
                "转换完成",
                f"全部 {ok} 个文件转换成功！\n\n"
                f"输出目录：\n{self._out_dir.get()}")
        else:
            detail = "\n".join(
                f"• {name}\n  {err or '未知错误'}"
                for name, err in fail_list)
            messagebox.showwarning(
                "完成（有失败）",
                f"成功：{ok} 个　失败：{fail} 个\n\n"
                f"失败详情：\n{detail}\n\n"
                f"输出目录：\n{self._out_dir.get()}")
