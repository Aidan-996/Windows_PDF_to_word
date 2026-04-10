"""UI 主题色 & 公用小组件 —— 央视风格"""

import tkinter as tk

# ── 央视风格调色板 ────────────────────────────────────────────────────────────
PRIMARY      = "#C0392B"     # 中国红（主色）
PRIMARY_DK   = "#922B21"     # 深红（按下/hover）
PRIMARY_LT   = "#FADBD8"     # 浅红（背景点缀）
ACCENT_GOLD  = "#C9A84C"     # 金色点缀
ACCENT_GOLD_DK = "#A38735"   # 深金
HEADER_BG    = "#1A1A2E"     # 深色标题栏
HEADER_BG2   = "#16213E"     # 标题栏渐变（备用）
LIGHT_BG     = "#F2F3F5"     # 主背景
WHITE        = "#FFFFFF"
CARD_BG      = "#FAFBFC"     # 卡片/面板背景
BORDER       = "#E0E3E8"     # 边框线
TXT          = "#1A1A2E"     # 主文字（近黑）
TXT_LIGHT    = "#8C939A"     # 辅助文字
TXT_WHITE    = "#FDFEFE"     # 白色文字
LIST_HDR     = "#F5F6F8"     # 列表表头
STATUS_OK    = "#27AE60"     # 成功绿
STATUS_ERR   = "#E74C3C"     # 失败红
STATUS_RUN   = "#E67E22"     # 进行中橙
OCR_COLOR    = "#C0392B"     # OCR 标识（统一用红）
OCR_BG       = "#FDF2F2"     # OCR 卡片背景
SEL_BG       = "#FDE8E8"     # 选中行背景

# ── 标注区域功能色（保持蓝/绿区分表格/图片）──────────────────────────────────
ANNO_TABLE      = "#2980B9"
ANNO_TABLE_SEL  = "#1A5276"
ANNO_TABLE_BG   = "#EBF5FB"
ANNO_TABLE_BG_S = "#D6EAF8"
ANNO_IMAGE      = "#27AE60"
ANNO_IMAGE_SEL  = "#1E8449"
ANNO_IMAGE_BG   = "#EAFAF1"
ANNO_IMAGE_BG_S = "#D5F5E3"


def styled_btn(parent, text, command, width=7, color=PRIMARY):
    """央视风格扁平按钮"""
    hover = PRIMARY_DK if color == PRIMARY else "#7B241C"
    btn = tk.Button(
        parent, text=text, command=command,
        bg=color, fg=WHITE,
        font=("微软雅黑", 10, "bold"),
        relief="flat", bd=0, padx=12, pady=6,
        cursor="hand2", width=width,
        activebackground=hover, activeforeground=WHITE,
    )
    btn.bind("<Enter>", lambda e: btn.config(bg=hover))
    btn.bind("<Leave>", lambda e: btn.config(bg=color))
    return btn
