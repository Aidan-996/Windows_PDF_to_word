# PDF转换Word智能工具 v1.0

一键把 PDF 变成可编辑的 Word 文档，支持**文字版**和**扫描版（OCR识别）**两种模式。

---

## 功能一览

- **普通转换** — 文字版 PDF 直接提取内容转 Word，保留表格、目录、分页
- **OCR 识别** — 扫描版 PDF（图片型）通过 AI 文字识别转 Word，支持中/英/繁体/日文
- **批量处理** — 同时添加多个文件，一键全部转换
- **实时预览** — 右侧面板预览 PDF 内容，支持翻���、缩放
- **拖拽添加** — 直接拖拽 PDF 文件到��口即可添加
- **标注工具** — 在预览中框选表格/图片区域

## 快速开始

### 方法一：直接运行（推荐）

下载 Release 中的压缩包，解压后双击 `PDF转换Word智能工具v1.0.exe` 即可使用。

### 方法二：从源码运行

```bash
# 1. 安装依赖
pip install pdf2docx pdf2image Pillow python-docx pypdf numpy tkinterdnd2

# 2. 运行
python pdf_to_word.py
```

> 首次启用 OCR 功能时会自动安装 easyocr（含 PyTorch，约 200MB）。

### 方法三：自行打包 exe

```bash
# 运行打包脚本（需要 poppler 放在同级目录）
build.bat
```

## 项目结构

```
PDF-to-Word/
├── pdf_to_word.py        # 入口文件
├── build.bat             # PyInstaller 打包脚本
├── version_info.py       # exe 版本信息
├── app_icon.ico          # 应用图标
└── core/                 # 核心代码
    ├── __init__.py       # DnD 拖拽支持检测
    ├── deps.py           # 依赖自动检测与安装
    ├── converter.py      # PDF→Word 转换引擎
    ├── theme.py          # 界面主题配色
    ├── app.py            # 主窗口
    ├── ui_file_list.py   # 文件列表面板
    ├── ui_preview.py     # PDF 预览面板
    ├── ui_annotation.py  # 标注工具栏
    └── ui_output_opts.py # 输出选项弹窗
```

## 系统要求

- Windows 10 / 11
- 普通转换���无额外要求
- OCR 模式：需要联网安装 easyocr（首次使用时自动安装）

## 截图

启动后添加 PDF 文件，点击「开始转换」即可。输出默认保存在 `我的文档/PDFPai/` 目录。

---

## 开源协议

本项目采用 **MIT License** 开源，欢迎二次开发。

**但请保留原作者版权信息：**

```
Copyright © 2025 Hum0ro. All rights reserved.
作者VX：XNHSDJ
Hum0ro & Jack
```

二次开发时请在程序界面或文档中保留上述信息，感谢��重原创。

---

> **Copyright © 2025 Hum0ro & Jack. All rights reserved.**
>
> 作者VX：XNHSDJ
