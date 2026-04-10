"""标注工具栏 Mixin：在预览画布上绘制/管理表格与图片区域 —— 央视风格"""

import tkinter as tk

from .theme import (WHITE, TXT_LIGHT, PRIMARY, PRIMARY_DK, PRIMARY_LT,
                    HEADER_BG, ACCENT_GOLD, STATUS_OK, STATUS_ERR, BORDER,
                    ANNO_TABLE, ANNO_TABLE_SEL, ANNO_TABLE_BG, ANNO_TABLE_BG_S,
                    ANNO_IMAGE, ANNO_IMAGE_SEL, ANNO_IMAGE_BG, ANNO_IMAGE_BG_S)


class AnnotationMixin:

    # ── 标注工具栏构建 ────────────────────────────────────────────────────────
    def _build_anno_bar(self, parent):
        """9个图标均匀排列在深色标题栏右侧"""
        self._anno_btns        = {}
        self._anno_tooltip_win = None
        S = 28

        # ── 图标绘制函数 ──────────────────────────────────────────────────
        def _draw_show_table(cv, S, lc, hl):
            m = S // 5
            cv.create_rectangle(m, m, S - m, S - m, outline=lc, width=1)
            mx = S // 2; my = S // 2
            cv.create_line(mx, m, mx, S - m, fill=lc, width=1)
            cv.create_line(m, my, S - m, my, fill=lc, width=1)
            cr = S // 6; cx = S - m - 1; cy = S - m - 1
            cv.create_oval(cx - cr, cy - cr, cx + cr, cy + cr,
                           outline=hl, fill=HEADER_BG, width=1)
            cv.create_line(cx + cr - 1, cy + cr - 1, S - 2, S - 2,
                           fill=hl, width=2)

        def _draw_draw_table(cv, S, lc, hl):
            m = S // 7; r = S - S // 4
            cv.create_rectangle(m, m, r, r, outline=lc, width=1, dash=(3, 2))
            mx = (m + r) // 2; my = (m + r) // 2
            cv.create_line(mx, m, mx, r, fill=lc, width=1, dash=(2, 2))
            cv.create_line(m, my, r, my, fill=lc, width=1, dash=(2, 2))
            px = S - S // 4; py = S - S // 4; ps = S // 5
            pts = [px + ps, py, px + ps + ps // 2, py + ps // 2,
                   px + ps // 2, py + ps + ps // 2, px, py + ps]
            cv.create_polygon(pts, fill=hl, outline=hl)
            cv.create_line(px, py + ps, px - 2, py + ps + 4, fill=hl, width=1)
            cv.create_line(px + ps // 2, py + ps + ps // 2,
                           px - 2, py + ps + 4, fill=hl, width=1)

        def _draw_draw_image(cv, S, lc, hl):
            m = S // 7; r = S - S // 4
            cv.create_rectangle(m, m, r, r, outline=lc, width=1, dash=(3, 2))
            mid = (m + r) // 2
            cv.create_line(m + 1, r - 2, mid - 1, m + S // 6,
                           mid + 1, m + S // 5 + 1, r - 1, m + 2,
                           fill=lc, width=1)
            sc = S // 8
            cv.create_oval(m + 1, m + 1, m + sc + 1, m + sc + 1,
                           outline=lc, width=1)
            px = S - S // 4; py = S - S // 4; ps = S // 5
            pts = [px + ps, py, px + ps + ps // 2, py + ps // 2,
                   px + ps // 2, py + ps + ps // 2, px, py + ps]
            cv.create_polygon(pts, fill=hl, outline=hl)
            cv.create_line(px, py + ps, px - 2, py + ps + 4, fill=hl, width=1)
            cv.create_line(px + ps // 2, py + ps + ps // 2,
                           px - 2, py + ps + 4, fill=hl, width=1)

        def _draw_delete_region(cv, S, lc, hl):
            m = S // 7; r = S - S // 4
            cv.create_rectangle(m, m, r, r, outline=lc, width=1, dash=(3, 2))
            mx = (m + r) // 2; my = (m + r) // 2
            cv.create_line(mx, m, mx, r, fill=lc, width=1, dash=(2, 2))
            cv.create_line(m, my, r, my, fill=lc, width=1, dash=(2, 2))
            cr2 = S // 6; cx2 = S - cr2 - 1; cy2 = S - cr2 - 1
            cv.create_oval(cx2 - cr2, cy2 - cr2, cx2 + cr2, cy2 + cr2,
                           fill=STATUS_ERR, outline=HEADER_BG, width=1)
            d = cr2 // 2 + 1
            cv.create_line(cx2 - d, cy2 - d, cx2 + d, cy2 + d,
                           fill=WHITE, width=1)
            cv.create_line(cx2 + d, cy2 - d, cx2 - d, cy2 + d,
                           fill=WHITE, width=1)

        def _draw_add_row(cv, S, lc, hl):
            m = S // 5; bm = S - m; mid = (m + bm) // 2
            cv.create_rectangle(m, m, bm, bm, outline=lc, width=1)
            cv.create_line(m, mid, bm, mid, fill=lc, width=1)
            cr = S // 7; px = bm; py = bm
            cv.create_oval(px - cr, py - cr, px + cr, py + cr,
                           fill=hl, outline="")
            d = cr - 1
            cv.create_line(px - d, py, px + d, py, fill=HEADER_BG, width=1)
            cv.create_line(px, py - d, px, py + d, fill=HEADER_BG, width=1)

        def _draw_add_col(cv, S, lc, hl):
            m = S // 5; bm = S - m; mid = (m + bm) // 2
            cv.create_rectangle(m, m, bm, bm, outline=lc, width=1)
            cv.create_line(mid, m, mid, bm, fill=lc, width=1)
            cr = S // 7; px = bm; py = bm
            cv.create_oval(px - cr, py - cr, px + cr, py + cr,
                           fill=hl, outline="")
            d = cr - 1
            cv.create_line(px - d, py, px + d, py, fill=HEADER_BG, width=1)
            cv.create_line(px, py - d, px, py + d, fill=HEADER_BG, width=1)

        def _draw_delete_line(cv, S, lc, hl):
            m = S // 5; bm = S - m; mid = (m + bm) // 2
            cv.create_rectangle(m, m, bm, bm, outline=lc, width=1)
            cv.create_line(m, mid, bm, mid, fill=lc, width=1, dash=(3, 2))
            cr = S // 7; cx2 = bm; cy2 = bm
            cv.create_oval(cx2 - cr, cy2 - cr, cx2 + cr, cy2 + cr,
                           fill=STATUS_ERR, outline=HEADER_BG, width=1)
            d = cr - 1
            cv.create_line(cx2 - d, cy2 - d, cx2 + d, cy2 + d,
                           fill=WHITE, width=1)
            cv.create_line(cx2 + d, cy2 - d, cx2 - d, cy2 + d,
                           fill=WHITE, width=1)

        def _draw_split_cell(cv, S, lc, hl):
            m = S // 6; bm = S - m; mid = (m + bm) // 2; ay = S // 2
            cv.create_rectangle(m, m, bm, bm, outline=lc, width=1)
            cv.create_line(mid, m + 1, mid, bm - 1, fill=hl, width=1,
                           dash=(2, 2))
            a = S // 7
            cv.create_line(mid - 2, ay, m + 2, ay, fill=hl, width=1)
            cv.create_line(m + 2, ay, m + 2 + a, ay - a // 2,
                           fill=hl, width=1)
            cv.create_line(m + 2, ay, m + 2 + a, ay + a // 2,
                           fill=hl, width=1)
            cv.create_line(mid + 2, ay, bm - 2, ay, fill=hl, width=1)
            cv.create_line(bm - 2, ay, bm - 2 - a, ay - a // 2,
                           fill=hl, width=1)
            cv.create_line(bm - 2, ay, bm - 2 - a, ay + a // 2,
                           fill=hl, width=1)

        def _draw_merge_cell(cv, S, lc, hl):
            m = S // 6; bm = S - m; mid = (m + bm) // 2; ay = S // 2
            cv.create_rectangle(m, m, mid - 1, bm, outline=lc, width=1)
            cv.create_rectangle(mid + 1, m, bm, bm, outline=lc, width=1)
            a = S // 7
            cv.create_line(m + a, ay, mid - 2, ay, fill=hl, width=1)
            cv.create_line(mid - 2, ay, mid - 2 - a, ay - a // 2,
                           fill=hl, width=1)
            cv.create_line(mid - 2, ay, mid - 2 - a, ay + a // 2,
                           fill=hl, width=1)
            cv.create_line(bm - a, ay, mid + 2, ay, fill=hl, width=1)
            cv.create_line(mid + 2, ay, mid + 2 + a, ay - a // 2,
                           fill=hl, width=1)
            cv.create_line(mid + 2, ay, mid + 2 + a, ay + a // 2,
                           fill=hl, width=1)

        _ICONS = [
            ("show_table",    _draw_show_table,    "显示表格",
             "切换预览中标注区域叠层显示/隐藏"),
            ("draw_table",    _draw_draw_table,    "画表格区域",
             "在预览图上拖拽绘制表格识别区域（蓝色框）"),
            ("draw_image",    _draw_draw_image,    "画图片区域",
             "在预览图上拖拽绘制图片提取区域（绿色框）"),
            ("delete_region", _draw_delete_region, "删除区域",
             "点击选中标注区域后将其删除"),
            ("add_row",       _draw_add_row,       "添加行",
             "在选中表格区域末尾追加一行"),
            ("add_col",       _draw_add_col,       "添加列",
             "在选中表格区域末尾追加一列"),
            ("delete_line",   _draw_delete_line,   "删除线",
             "删除选中表格区域中最后一条分割线"),
            ("split_cell",    _draw_split_cell,    "拆分表格单元",
             "将选中单元格水平或垂直一分为二"),
            ("merge_cell",    _draw_merge_cell,    "合并表格单元",
             "将相邻选中单元格合并为一个"),
        ]

        # ── 创建所有图标 Canvas（深色底）──
        canvases = []
        for tool_id, draw_fn, tip_s, tip_l in _ICONS:
            cv = tk.Canvas(parent, width=S, height=S,
                           bg=HEADER_BG, highlightthickness=0, cursor="hand2")
            self._anno_btns[tool_id] = cv
            self._bind_tooltip(cv, tip_s, tip_l)

            def _make_draw(fn, c):
                def _redraw(bg=HEADER_BG, lc="#7F8C9A", hl=ACCENT_GOLD):
                    c.delete("all")
                    c.configure(bg=bg)
                    fn(c, S, lc, hl)
                return _redraw

            _redraw = _make_draw(draw_fn, cv)
            _redraw()

            def _make_events(rd, tid):
                def _enter(e):   rd(HEADER_BG, "#C0C8D0", ACCENT_GOLD)
                def _leave(e):   rd(HEADER_BG, "#7F8C9A", ACCENT_GOLD)
                def _press(e):   rd("#0F1526", "#E0E0E0", "#FFD700")
                def _release(e):
                    rd(HEADER_BG, "#C0C8D0", ACCENT_GOLD)
                    self._anno_btn_click(tid)
                return _enter, _leave, _press, _release

            en, le, pr, re = _make_events(_redraw, tool_id)
            cv.bind("<Enter>",           en)
            cv.bind("<Leave>",           le)
            cv.bind("<ButtonPress-1>",   pr)
            cv.bind("<ButtonRelease-1>", re)
            canvases.append(cv)

        self._anno_status = tk.Label(
            parent, text="", font=("微软雅黑", 7),
            bg=HEADER_BG, fg="#8899AA", anchor="center",
        )

        def _relayout(event=None):
            w = parent.winfo_width()
            h = parent.winfo_height() or 52
            if w < 10:
                return
            n    = len(canvases)
            slot = w / n
            y0   = (h - S) // 2
            for i, cv in enumerate(canvases):
                cx = int(slot * i + slot / 2)
                cv.place(x=cx - S // 2, y=max(0, y0))
            self._anno_status.place(x=0, y=h - 14, width=w)

        parent.bind("<Configure>", lambda e: _relayout(e))
        parent.after(80, _relayout)

    # ── Tooltip（深色风格）────────────────────────────────────────────────────
    def _bind_tooltip(self, widget, title, detail=""):
        widget.bind("<Enter>", lambda e: self._show_tooltip(widget, title, detail))
        widget.bind("<Leave>", lambda e: self._hide_tooltip())

    def _show_tooltip(self, widget, title, detail=""):
        self._hide_tooltip()
        x = widget.winfo_rootx() + widget.winfo_width() // 2
        y = widget.winfo_rooty() + widget.winfo_height() + 4
        self._anno_tooltip_win = tw = tk.Toplevel(self)
        tw.wm_overrideredirect(True)
        tw.wm_geometry(f"+{x}+{y}")
        tw.configure(bg=HEADER_BG)
        outer = tk.Frame(tw, bg=HEADER_BG, bd=0); outer.pack(padx=1, pady=1)
        inner = tk.Frame(outer, bg="#FAFBFC", bd=0); inner.pack()
        tk.Label(inner, text=title, font=("微软雅黑", 9, "bold"),
                 bg="#FAFBFC", fg=HEADER_BG, padx=10, pady=4).pack(anchor="w")
        if detail:
            tk.Label(inner, text=detail, font=("微软雅黑", 8),
                     bg="#FAFBFC", fg="#7F8C8D", padx=10, pady=(0, 6),
                     wraplength=240, justify="left").pack(anchor="w")

    def _hide_tooltip(self):
        if self._anno_tooltip_win:
            try:
                self._anno_tooltip_win.destroy()
            except Exception:
                pass
            self._anno_tooltip_win = None

    # ── 按键分发 ──────────────────────────────────────────────────────────────
    def _anno_btn_click(self, tool_id):
        toggle_tools = {"draw_table", "draw_image"}
        if tool_id == "show_table":
            self._anno_toggle_show()
        elif tool_id in toggle_tools:
            self._anno_activate(tool_id)
        elif tool_id == "delete_region":
            self._anno_delete_region()
        elif tool_id == "add_row":
            self._anno_add_row()
        elif tool_id == "add_col":
            self._anno_add_col()
        elif tool_id == "delete_line":
            self._anno_delete_line()
        elif tool_id == "split_cell":
            self._anno_split_cell()
        elif tool_id == "merge_cell":
            self._anno_merge_cell()
        self._anno_refresh_btn_states()

    def _anno_refresh_btn_states(self):
        active = self._anno_tool.get()
        for tid, cv in self._anno_btns.items():
            if tid == active:
                cv.configure(bg="#2A3A5E", highlightthickness=1,
                             highlightbackground=ACCENT_GOLD)
            elif tid == "show_table" and self._anno_show.get():
                cv.configure(bg="#1A3A2E", highlightthickness=1,
                             highlightbackground=STATUS_OK)
            else:
                cv.configure(bg=HEADER_BG, highlightthickness=0)

    # ── 显示/隐藏 ────────────────────────────────────────────────────────────
    def _anno_toggle_show(self):
        self._anno_show.set(not self._anno_show.get())
        self._anno_redraw_regions()
        state = "显示" if self._anno_show.get() else "隐藏"
        self._anno_status.config(text=f"标注叠层已{state}")
        self._anno_refresh_btn_states()

    # ── 激活绘制工具 ──────────────────────────────────────────────────────────
    def _anno_activate(self, tool_id):
        if self._anno_tool.get() == tool_id:
            self._anno_tool.set("none")
            self._anno_status.config(text="")
            self._prev_canvas.config(cursor="")
        else:
            self._anno_tool.set(tool_id)
            if tool_id == "draw_table":
                self._anno_status.config(text="在预览图上拖拽绘制表格区域（蓝框）")
                self._prev_canvas.config(cursor="crosshair")
            elif tool_id == "draw_image":
                self._anno_status.config(text="在预览图上拖拽绘制图片区域（绿框）")
                self._prev_canvas.config(cursor="crosshair")

    # ── 鼠标拖拽绘制 ──────────────────────────────────────────────────────────
    def _anno_on_press(self, event):
        tool = self._anno_tool.get()
        if tool in ("draw_table", "draw_image"):
            cx = self._prev_canvas.canvasx(event.x)
            cy = self._prev_canvas.canvasy(event.y)
            self._anno_drag_start = (cx, cy)
            color = ANNO_TABLE if tool == "draw_table" else ANNO_IMAGE
            self._anno_rubber = self._prev_canvas.create_rectangle(
                cx, cy, cx, cy,
                outline=color, width=2, dash=(4, 3), tags="rubber",
            )
        elif tool == "none":
            self._anno_pick_region(event)

    def _anno_on_drag(self, event):
        if self._anno_drag_start and self._anno_rubber:
            cx = self._prev_canvas.canvasx(event.x)
            cy = self._prev_canvas.canvasy(event.y)
            x1, y1 = self._anno_drag_start
            self._prev_canvas.coords(self._anno_rubber, x1, y1, cx, cy)

    def _anno_on_release(self, event):
        tool = self._anno_tool.get()
        if tool in ("draw_table", "draw_image") and self._anno_drag_start:
            cx = self._prev_canvas.canvasx(event.x)
            cy = self._prev_canvas.canvasy(event.y)
            x1, y1 = self._anno_drag_start
            if self._anno_rubber:
                self._prev_canvas.delete(self._anno_rubber)
                self._anno_rubber = None
            self._anno_drag_start = None
            if abs(cx - x1) < 10 or abs(cy - y1) < 10:
                return
            rx1, rx2 = (x1, cx) if x1 < cx else (cx, x1)
            ry1, ry2 = (y1, cy) if y1 < cy else (cy, y1)
            nx1, ny1, nx2, ny2 = self._canvas_to_norm(rx1, ry1, rx2, ry2)
            if nx1 is None:
                return
            region = {
                "type":  "table" if tool == "draw_table" else "image",
                "page":  self._prev_page,
                "nx1": nx1, "ny1": ny1, "nx2": nx2, "ny2": ny2,
                "rows": 1, "cols": 1,
            }
            self._anno_regions.append(region)
            self._anno_sel_idx = len(self._anno_regions) - 1
            self._anno_redraw_regions()
            rtype = "表格" if region["type"] == "table" else "图片"
            pw = int((nx2 - nx1) * 100); ph = int((ny2 - ny1) * 100)
            self._anno_status.config(
                text=f"已绘制{rtype}区域 ({pw}%x{ph}%)，"
                     f"共 {len(self._anno_regions)} 个标注")

    # ── 坐标转换 ──────────────────────────────────────────────────────────────
    def _canvas_to_norm(self, cx1, cy1, cx2, cy2):
        r = self._img_render
        if not r or r["w"] == 0 or r["h"] == 0:
            return None, None, None, None
        nx1 = max(0.0, min(1.0, (cx1 - r["x_off"]) / r["w"]))
        ny1 = max(0.0, min(1.0, (cy1 - r["y_off"]) / r["h"]))
        nx2 = max(0.0, min(1.0, (cx2 - r["x_off"]) / r["w"]))
        ny2 = max(0.0, min(1.0, (cy2 - r["y_off"]) / r["h"]))
        return nx1, ny1, nx2, ny2

    def _norm_to_canvas(self, nx1, ny1, nx2, ny2):
        r = self._img_render
        if not r:
            return None, None, None, None
        cx1 = r["x_off"] + nx1 * r["w"]
        cy1 = r["y_off"] + ny1 * r["h"]
        cx2 = r["x_off"] + nx2 * r["w"]
        cy2 = r["y_off"] + ny2 * r["h"]
        return cx1, cy1, cx2, cy2

    # ── 命中测试 ──────────────────────────────────────────────────────────────
    def _anno_pick_region(self, event):
        cx = self._prev_canvas.canvasx(event.x)
        cy = self._prev_canvas.canvasy(event.y)
        nx, ny, _, _ = self._canvas_to_norm(cx, cy, cx, cy)
        if nx is None:
            return
        for i, r in enumerate(self._anno_regions):
            if (r["page"] == self._prev_page
                    and r["nx1"] <= nx <= r["nx2"]
                    and r["ny1"] <= ny <= r["ny2"]):
                self._anno_sel_idx = i
                self._anno_redraw_regions()
                rtype = "表格" if r["type"] == "table" else "图片"
                self._anno_status.config(
                    text=f"已选中第 {i+1} 个{rtype}区域 "
                         f"(行:{r['rows']} 列:{r['cols']})")
                return
        self._anno_sel_idx = None
        self._anno_redraw_regions()
        self._anno_status.config(text="")

    # ── 绘制所有标注叠层 ──────────────────────────────────────────────────────
    def _anno_redraw_regions(self):
        self._prev_canvas.delete("anno")
        if not self._anno_show.get() or not self._img_render:
            return
        for i, r in enumerate(self._anno_regions):
            if r["page"] != self._prev_page:
                continue
            x1, y1, x2, y2 = self._norm_to_canvas(
                r["nx1"], r["ny1"], r["nx2"], r["ny2"])
            if x1 is None:
                continue
            is_sel = (i == self._anno_sel_idx)
            if r["type"] == "table":
                fill_c    = ANNO_TABLE_BG if not is_sel else ANNO_TABLE_BG_S
                outline_c = ANNO_TABLE    if not is_sel else ANNO_TABLE_SEL
            else:
                fill_c    = ANNO_IMAGE_BG if not is_sel else ANNO_IMAGE_BG_S
                outline_c = ANNO_IMAGE    if not is_sel else ANNO_IMAGE_SEL

            self._prev_canvas.create_rectangle(
                x1, y1, x2, y2,
                fill=fill_c, outline="", stipple="gray25", tags="anno")
            self._prev_canvas.create_rectangle(
                x1, y1, x2, y2,
                fill="", outline=outline_c,
                width=2 if is_sel else 1, tags="anno")

            if r["type"] == "table" and (r["rows"] > 1 or r["cols"] > 1):
                w = x2 - x1; h = y2 - y1
                for row_i in range(1, r["rows"]):
                    yy = y1 + h * row_i / r["rows"]
                    self._prev_canvas.create_line(
                        x1, yy, x2, yy,
                        fill=outline_c, width=1, dash=(4, 2), tags="anno")
                for col_i in range(1, r["cols"]):
                    xx = x1 + w * col_i / r["cols"]
                    self._prev_canvas.create_line(
                        xx, y1, xx, y2,
                        fill=outline_c, width=1, dash=(4, 2), tags="anno")

            label = ("表格" if r["type"] == "table" else "图片") + f" {i+1}"
            if r["type"] == "table" and (r["rows"] > 1 or r["cols"] > 1):
                label += f"  {r['rows']}行x{r['cols']}列"
            self._prev_canvas.create_text(
                x1 + 4, y1 + 3, anchor="nw", text=label,
                font=("微软雅黑", 8), fill=outline_c, tags="anno")

    # ── 删除区域 ──────────────────────────────────────────────────────────────
    def _anno_delete_region(self):
        if self._anno_sel_idx is None:
            self._anno_status.config(text="请先点击选中一个区域")
            return
        idx = self._anno_sel_idx
        if 0 <= idx < len(self._anno_regions):
            rtype = ("表格" if self._anno_regions[idx]["type"] == "table"
                     else "图片")
            self._anno_regions.pop(idx)
            self._anno_sel_idx = None
            self._anno_redraw_regions()
            self._anno_status.config(
                text=f"已删除第 {idx+1} 个{rtype}区域，"
                     f"剩余 {len(self._anno_regions)} 个")

    # ── 添加行 / 列 ──────────────────────────────────────────────────────────
    def _anno_add_row(self):
        r = self._anno_get_sel_table()
        if r is None:
            return
        r["rows"] += 1
        self._anno_redraw_regions()
        self._anno_status.config(
            text=f"已添加行 -> {r['rows']} 行 x {r['cols']} 列")

    def _anno_add_col(self):
        r = self._anno_get_sel_table()
        if r is None:
            return
        r["cols"] += 1
        self._anno_redraw_regions()
        self._anno_status.config(
            text=f"已添加列 -> {r['rows']} 行 x {r['cols']} 列")

    # ── 删除线 ────────────────────────────────────────────────────────────────
    def _anno_delete_line(self):
        r = self._anno_get_sel_table()
        if r is None:
            return
        if r["rows"] > 1:
            r["rows"] -= 1
            self._anno_redraw_regions()
            self._anno_status.config(
                text=f"已删除最后一行线 -> {r['rows']} 行 x {r['cols']} 列")
        elif r["cols"] > 1:
            r["cols"] -= 1
            self._anno_redraw_regions()
            self._anno_status.config(
                text=f"已删除最后一列线 -> {r['rows']} 行 x {r['cols']} 列")
        else:
            self._anno_status.config(text="该区域已无分割线可删除")

    # ── 拆分 / 合并 ──────────────────────────────────────────────────────────
    def _anno_split_cell(self):
        r = self._anno_get_sel_table()
        if r is None:
            return
        r["rows"] += 1
        self._anno_redraw_regions()
        self._anno_status.config(
            text=f"已拆分单元格（水平）-> {r['rows']} 行 x {r['cols']} 列")

    def _anno_merge_cell(self):
        r = self._anno_get_sel_table()
        if r is None:
            return
        if r["rows"] > 1:
            r["rows"] -= 1
            self._anno_redraw_regions()
            self._anno_status.config(
                text=f"已合并单元格（末尾两行）-> {r['rows']} 行 x {r['cols']} 列")
        elif r["cols"] > 1:
            r["cols"] -= 1
            self._anno_redraw_regions()
            self._anno_status.config(
                text=f"已合并单元格（末尾两列）-> {r['rows']} 行 x {r['cols']} 列")
        else:
            self._anno_status.config(text="该区域只有一个单元格，无法合并")

    # ── 辅助 ──────────────────────────────────────────────────────────────────
    def _anno_get_sel_table(self):
        if self._anno_sel_idx is None:
            self._anno_status.config(text="请先点击选中一个表格区域")
            return None
        idx = self._anno_sel_idx
        if not (0 <= idx < len(self._anno_regions)):
            self._anno_status.config(text="所选区域无效")
            return None
        r = self._anno_regions[idx]
        if r["type"] != "table":
            self._anno_status.config(text="此功能仅适用于表格区域，请选中蓝色框")
            return None
        return r
