"""输出选项弹窗 Mixin —— 央视风格"""

import tkinter as tk

from .theme import (PRIMARY, PRIMARY_DK, PRIMARY_LT, ACCENT_GOLD,
                    HEADER_BG, WHITE, TXT, TXT_WHITE, BORDER, STATUS_OK)


class OutputOptsMixin:

    def _show_output_opts(self):
        dlg = tk.Toplevel(self)
        dlg.title("输出选项")
        dlg.resizable(False, False)
        dlg.configure(bg=WHITE)
        dlg.grab_set()

        # 深色标题
        hdr = tk.Frame(dlg, bg=HEADER_BG, height=40)
        hdr.pack(fill="x"); hdr.pack_propagate(False)
        tk.Label(hdr, text="输出选项", font=("微软雅黑", 10, "bold"),
                 bg=HEADER_BG, fg=TXT_WHITE).pack(side="left", padx=14)
        # 金色点缀线
        tk.Frame(dlg, bg=ACCENT_GOLD, height=2).pack(fill="x")

        body = tk.Frame(dlg, bg=WHITE, padx=20, pady=14)
        body.pack(fill="both", expand=True)

        tk.Label(body, text="请选择目标格式", font=("微软雅黑", 9, "bold"),
                 bg=WHITE, fg=TXT).pack(anchor="w", pady=(0, 6))

        _fmt = tk.StringVar(value=self._out_fmt.get())

        fmt_frame = tk.Frame(body, bg=WHITE)
        fmt_frame.pack(anchor="w", padx=8)
        for label, val in [("DOC", "doc"), ("DOCX", "docx")]:
            tk.Radiobutton(
                fmt_frame, text=label, variable=_fmt, value=val,
                font=("微软雅黑", 9), bg=WHITE, fg=TXT,
                activebackground=WHITE, selectcolor=PRIMARY,
            ).pack(anchor="w", pady=3)

        tk.Frame(body, bg=BORDER, height=1).pack(fill="x", pady=(10, 8))

        tk.Label(body, text="识别选项", font=("微软雅黑", 9, "bold"),
                 bg=WHITE, fg=TXT).pack(anchor="w", pady=(0, 4))

        _toc = tk.BooleanVar(value=self._opt_toc.get())
        _hf  = tk.BooleanVar(value=self._opt_header_footer.get())
        _lp  = tk.BooleanVar(value=self._opt_list_para.get())

        opts_frame = tk.Frame(body, bg=WHITE, padx=8)
        opts_frame.pack(anchor="w", fill="x")

        cb_toc = tk.Checkbutton(
            opts_frame, text="识别目录", variable=_toc,
            font=("微软雅黑", 9), bg=WHITE, fg=TXT,
            activebackground=WHITE, selectcolor=STATUS_OK, cursor="hand2")
        cb_toc.pack(anchor="w", pady=3)

        cb_hf = tk.Checkbutton(
            opts_frame, text="识别页眉页脚", variable=_hf,
            font=("微软雅黑", 9), bg=WHITE, fg=TXT,
            activebackground=WHITE, selectcolor=STATUS_OK, cursor="hand2")
        cb_hf.pack(anchor="w", pady=3)

        cb_lp = tk.Checkbutton(
            opts_frame, text="识别列表段落", variable=_lp,
            font=("微软雅黑", 9), bg=WHITE, fg=TXT,
            activebackground=WHITE, selectcolor=STATUS_OK, cursor="hand2")
        cb_lp.pack(anchor="w", pady=3)

        def _on_fmt_change(*_):
            state = "normal" if _fmt.get() == "docx" else "disabled"
            for cb in (cb_toc, cb_hf, cb_lp):
                cb.config(state=state)
        _fmt.trace_add("write", _on_fmt_change)
        _on_fmt_change()

        tk.Frame(dlg, bg=BORDER, height=1).pack(fill="x")
        btn_row = tk.Frame(dlg, bg=WHITE, pady=10)
        btn_row.pack(fill="x")

        def _confirm():
            self._out_fmt.set(_fmt.get())
            if _fmt.get() == "docx":
                self._opt_toc.set(_toc.get())
                self._opt_header_footer.set(_hf.get())
                self._opt_list_para.set(_lp.get())
            dlg.destroy()

        tk.Button(btn_row, text="确定", command=_confirm,
                  bg=PRIMARY, fg=WHITE, font=("微软雅黑", 9, "bold"),
                  relief="flat", padx=20, pady=5, cursor="hand2",
                  activebackground=PRIMARY_DK, activeforeground=WHITE,
                  ).pack(side="right", padx=(0, 14))
        tk.Button(btn_row, text="取消", command=dlg.destroy,
                  bg="#F0F2F5", fg=TXT, font=("微软雅黑", 9),
                  relief="flat", padx=20, pady=5, cursor="hand2",
                  activebackground=BORDER,
                  ).pack(side="right", padx=(0, 6))

        dlg.update_idletasks()
        w = 320
        h = dlg.winfo_reqheight()
        x = self.winfo_x() + (self.winfo_width()  - w) // 2
        y = self.winfo_y() + (self.winfo_height() - h) // 2
        dlg.geometry(f"{w}x{h}+{x}+{y}")
