"""文件列表面板 Mixin：行选择、拖拽添加、状态显示 —— 央视风格"""

import os
import re
import threading
import tkinter as tk
from tkinter import filedialog, messagebox

from . import DND_AVAILABLE, DND_FILES
from .theme import (WHITE, TXT, TXT_LIGHT, PRIMARY, PRIMARY_LT, BORDER, SEL_BG,
                    CARD_BG, ACCENT_GOLD, STATUS_OK, STATUS_ERR, STATUS_RUN)
from .converter import detect_pdf_type


class FileListMixin:

    # ── 列表渲染 ──────────────────────────────────────────────────────────────
    def _row_bg(self, idx):
        if idx in self._sel_set:
            return SEL_BG
        return WHITE if idx % 2 == 0 else CARD_BG

    def _refresh_list(self):
        for w in self._inner.winfo_children():
            w.destroy()
        self._row_widgets = []
        for idx, (path, svar, out_var, err_var, type_var) in enumerate(self._files):
            bg = self._row_bg(idx)
            row = tk.Frame(self._inner, bg=bg)
            row.pack(fill="x")
            self._row_widgets.append(row)

            info_row = tk.Frame(row, bg=bg, height=32)
            info_row.pack(fill="x"); info_row.pack_propagate(False)

            for w in (row, info_row):
                w.bind("<Button-1>",
                       lambda e, i=idx: self._on_row_click(e, i))
                w.bind("<Control-Button-1>",
                       lambda e, i=idx: self._on_row_ctrl(e, i))
                w.bind("<Shift-Button-1>",
                       lambda e, i=idx: self._on_row_shift(e, i))

            if DND_AVAILABLE:
                for w in (row, info_row):
                    w.drop_target_register(DND_FILES)
                    w.dnd_bind("<<Drop>>", self._on_drop)

            num_lbl = tk.Label(info_row, text=str(idx + 1),
                               font=("微软雅黑", 9), bg=bg, fg=TXT_LIGHT,
                               width=5)
            num_lbl.pack(side="left", padx=(10, 0))
            name_lbl = tk.Label(info_row, text=os.path.basename(path),
                                font=("微软雅黑", 9), bg=bg, fg=TXT,
                                anchor="w")
            name_lbl.pack(side="left", fill="x", expand=True, padx=(4, 0))

            type_lbl = tk.Label(info_row, textvariable=type_var,
                                font=("微软雅黑", 8), bg=bg, fg=TXT_LIGHT,
                                width=7)
            type_lbl.pack(side="right", padx=(0, 4))

            def _upd_type(lbl=type_lbl, tv=type_var):
                v = tv.get()
                if "扫描" in v:   lbl.config(fg=STATUS_RUN)
                elif "文字" in v: lbl.config(fg=STATUS_OK)
                else:             lbl.config(fg=TXT_LIGHT)
            type_var.trace_add("write", lambda *a, fn=_upd_type: fn())

            slbl = tk.Label(info_row, textvariable=svar,
                            font=("微软雅黑", 9), bg=bg, width=7)
            slbl.pack(side="right", padx=(0, 4))

            for child in (num_lbl, name_lbl, type_lbl, slbl):
                child.bind("<Button-1>",
                           lambda e, i=idx: self._on_row_click(e, i))
                child.bind("<Control-Button-1>",
                           lambda e, i=idx: self._on_row_ctrl(e, i))
                child.bind("<Shift-Button-1>",
                           lambda e, i=idx: self._on_row_shift(e, i))

            # ── 快捷链接行 ──
            link_row = tk.Frame(row, bg=bg)
            tk.Label(link_row, bg=bg, width=5).pack(side="left", padx=(10, 0))
            link_lbl = tk.Label(link_row, text="",
                                font=("微软雅黑", 8, "underline"),
                                bg=bg, fg=PRIMARY, anchor="w",
                                cursor="hand2")
            link_lbl.pack(side="left", fill="x", expand=True,
                          padx=(4, 0), pady=(0, 4))

            def _upd(lbl=slbl, sv=svar, ev=err_var, p=path):
                v = sv.get()
                if "完成" in v:
                    lbl.config(fg=STATUS_OK)
                elif "失败" in v:
                    lbl.config(fg=STATUS_ERR)
                    def _show_err(e, ev=ev, p=p):
                        msg = ev.get() or "未知错误"
                        messagebox.showerror(
                            "转换失败",
                            f"文件：{os.path.basename(p)}\n\n错误信息：\n{msg}")
                    lbl.bind("<Button-1>", _show_err)
                    lbl.config(cursor="hand2")
                elif "转换中" in v or "OCR中" in v:
                    lbl.config(fg=STATUS_RUN)
                else:
                    lbl.config(fg=TXT_LIGHT)
            svar.trace_add("write", lambda *a, fn=_upd: fn())

            def _upd_link(lrow=link_row, llbl=link_lbl, ov=out_var):
                v = ov.get()
                if v:
                    fname = os.path.basename(v)
                    llbl.config(text=f"📄 {fname}")
                    def _open(e, p=v):
                        if os.name == "nt":
                            os.startfile(p)
                        else:
                            os.system(f'xdg-open "{p}"')
                    llbl.bind("<Button-1>", _open)
                    lrow.pack(fill="x")
                else:
                    lrow.pack_forget()
            out_var.trace_add("write", lambda *a, fn=_upd_link: fn())
            _upd_link()

        self._canvas.update_idletasks()
        self._canvas.configure(scrollregion=self._canvas.bbox("all"))

    def _repaint_rows(self):
        for idx, row in enumerate(self._row_widgets):
            bg = self._row_bg(idx)
            row.config(bg=bg)
            for child in row.winfo_children():
                try: child.config(bg=bg)
                except Exception: pass
                for grandchild in child.winfo_children():
                    try: grandchild.config(bg=bg)
                    except Exception: pass

    # ── 行点击 ────────────────────────────────────────────────────────────────
    def _on_row_click(self, event, idx):
        self._sel_set  = {idx}
        self._last_sel = idx
        self._repaint_rows()
        self._trigger_preview(idx)

    def _on_row_ctrl(self, event, idx):
        if idx in self._sel_set:
            self._sel_set.discard(idx)
        else:
            self._sel_set.add(idx)
        self._last_sel = idx
        self._repaint_rows()
        self._trigger_preview(idx)

    def _on_row_shift(self, event, idx):
        anchor = self._last_sel if self._last_sel is not None else idx
        lo, hi = min(anchor, idx), max(anchor, idx)
        self._sel_set = set(range(lo, hi + 1))
        self._repaint_rows()
        self._trigger_preview(idx)

    def _trigger_preview(self, idx):
        if 0 <= idx < len(self._files):
            path = self._files[idx][0]
            threading.Thread(target=self._show_preview, args=(path,),
                             daemon=True).start()

    # ── 文件操作 ──────────────────────────────────────────────────────────────
    def _add_files_from_paths(self, paths):
        added = 0
        for p in paths:
            p = p.strip()
            if (p.lower().endswith(".pdf")
                    and not any(f[0] == p for f in self._files)):
                type_var = tk.StringVar(value="检测中…")
                self._files.append((
                    p, tk.StringVar(value="等待开始"),
                    tk.StringVar(value=""), tk.StringVar(value=""),
                    type_var,
                ))

                def _do_detect(path=p, tv=type_var):
                    t = detect_pdf_type(path)
                    label = "🔍扫描版" if t == "scanned" else "📝文字版"
                    self.after(0, lambda: tv.set(label))
                threading.Thread(target=_do_detect, daemon=True).start()
                added += 1
        if added:
            self._refresh_list()
        return added

    def _on_drop(self, event):
        raw = event.data
        paths = re.findall(r'\{([^}]+)\}|(\S+)', raw)
        parsed = [a or b for a, b in paths]
        all_pdfs = []
        for p in parsed:
            if os.path.isdir(p):
                for fn in sorted(os.listdir(p)):
                    if fn.lower().endswith(".pdf"):
                        all_pdfs.append(os.path.join(p, fn))
            else:
                all_pdfs.append(p)
        n = self._add_files_from_paths(all_pdfs)
        if n:
            self._drop_zone.config(
                bg="#D5F5E3", fg=STATUS_OK,
                text=f"✅  已添加 {n} 个文件")
            self.after(2000, lambda: self._drop_zone.config(
                bg=PRIMARY_LT, fg=PRIMARY,
                text="📂  拖拽 PDF 文件到此处  （支持批量拖入）"))

    def _add_files(self):
        paths = filedialog.askopenfilenames(
            title="选择 PDF 文件",
            filetypes=[("PDF 文件", "*.pdf"), ("所有文件", "*.*")])
        self._add_files_from_paths(list(paths))

    def _delete_selected(self):
        if not self._files:
            return
        if self._sel_set:
            for idx in sorted(self._sel_set, reverse=True):
                if 0 <= idx < len(self._files):
                    self._files.pop(idx)
            self._sel_set  = set()
            self._last_sel = None
        else:
            self._files.pop()
        self._refresh_list()
        self._prev_pages = []
        self._preview_placeholder.place(relx=.5, rely=.45, anchor="center")

    # ── OCR toggle ────────────────────────────────────────────────────────────
    def _on_ocr_toggle(self):
        if self._use_ocr.get():
            self._ocr_lang_frame.pack(fill="x")
        else:
            self._ocr_lang_frame.pack_forget()

    def _on_lang_change(self, *_):
        pass
