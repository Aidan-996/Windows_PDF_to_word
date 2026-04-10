"""预览面板 Mixin：PDF 页面渲染、翻页、缩放"""

import threading


class PreviewMixin:

    def _show_preview(self, pdf_path):
        """加载 PDF 所有页到内存，然后显示第一页"""
        self._prev_token += 1
        token = self._prev_token

        def _load():
            try:
                from pdf2image import convert_from_path
                pages = convert_from_path(pdf_path, dpi=130)
                if token != self._prev_token:
                    return

                def _done():
                    if token != self._prev_token:
                        return
                    self._prev_path  = pdf_path
                    self._prev_pages = pages
                    self._prev_page  = 0
                    self._prev_zoom  = 1.0
                    self._update_nav()
                    self._redraw_preview()
                    self._preview_placeholder.place_forget()

                self.after(0, _done)
            except Exception:
                def _err():
                    self._preview_placeholder.config(
                        text="⚠ 预览失败\n请确认已安装 pdf2image 和 Pillow")
                    self._preview_placeholder.place(
                        relx=.5, rely=.45, anchor="center")
                self.after(0, _err)

        threading.Thread(target=_load, daemon=True).start()

    def _redraw_preview(self):
        """根据当前页、缩放倍率，在画布上绘制页面图像"""
        if not self._prev_pages:
            return
        try:
            from PIL import ImageTk, Image
            img = self._prev_pages[self._prev_page].copy()
            cw = self._prev_canvas.winfo_width()  or 500
            ch = self._prev_canvas.winfo_height() or 600

            base_ratio = cw / img.width
            ratio      = base_ratio * self._prev_zoom
            new_w = max(1, int(img.width  * ratio))
            new_h = max(1, int(img.height * ratio))

            img = img.resize((new_w, new_h), Image.LANCZOS)
            tk_img = ImageTk.PhotoImage(img)

            x_off = max(0, (cw - new_w) // 2)
            y_off = max(0, (ch - new_h) // 2) if new_h < ch else 0

            self._img_render = {
                "x_off": x_off, "y_off": y_off,
                "w": new_w, "h": new_h,
            }

            self._prev_canvas.delete("all")
            self._prev_canvas.config(
                scrollregion=(0, 0,
                              max(new_w + x_off * 2, cw),
                              max(new_h + y_off * 2, ch)))
            self._prev_canvas.create_image(x_off, y_off, anchor="nw",
                                           image=tk_img)
            self._tk_img = tk_img   # 防止 GC
            self._anno_redraw_regions()
        except Exception:
            pass

    def _update_nav(self):
        total = len(self._prev_pages)
        cur   = self._prev_page
        self._page_var.set(str(cur + 1) if total else "—")
        self._lbl_total.config(text=f"/ {total}" if total else "/ —")
        self._lbl_zoom.config(text=f"{int(self._prev_zoom * 100)}%")

    def _on_page_entry_jump(self, event=None):
        total = len(self._prev_pages)
        if not total:
            return
        try:
            page = int(self._page_var.get()) - 1
            page = max(0, min(total - 1, page))
        except ValueError:
            page = self._prev_page
        self._prev_page = page
        self._page_var.set(str(page + 1))
        self._update_nav()
        self._redraw_preview()

    def _prev_page_cmd(self):
        if self._prev_page > 0:
            self._prev_page -= 1
            self._update_nav()
            self._redraw_preview()

    def _next_page_cmd(self):
        if self._prev_page < len(self._prev_pages) - 1:
            self._prev_page += 1
            self._update_nav()
            self._redraw_preview()

    def _zoom_in(self):
        if self._prev_zoom < 3.0:
            self._prev_zoom = round(min(3.0, self._prev_zoom + 0.2), 1)
            self._update_nav()
            self._redraw_preview()

    def _zoom_out(self):
        if self._prev_zoom > 0.4:
            self._prev_zoom = round(max(0.4, self._prev_zoom - 0.2), 1)
            self._update_nav()
            self._redraw_preview()

    def _on_preview_scroll(self, event):
        delta = 0
        if hasattr(event, "delta") and event.delta:
            delta = -1 if event.delta > 0 else 1
        elif event.num == 4:
            delta = -1
        elif event.num == 5:
            delta = 1
        self._prev_canvas.yview_scroll(delta * 3, "units")

    def _on_preview_ctrl_scroll(self, event):
        if hasattr(event, "delta") and event.delta:
            if event.delta > 0:
                self._zoom_in()
            else:
                self._zoom_out()
        elif event.num == 4:
            self._zoom_in()
        elif event.num == 5:
            self._zoom_out()
