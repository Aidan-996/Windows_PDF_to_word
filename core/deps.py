"""依赖检测与自动安装（Python 包 + poppler）"""

import sys
import subprocess
import os
import shutil

_REQUIRED = {
    "pdf2docx":   "pdf2docx",
    "pdf2image":  "pdf2image",
    "PIL":        "Pillow",
    "docx":       "python-docx",
    "pypdf":      "pypdf",
    "numpy":      "numpy",
}

# OCR 依赖（仅启用 OCR 时按需安装，不打包进 exe）
_OCR_REQUIRED = {
    "easyocr":    "easyocr",
}

# poppler 安装目录（与项目根目录同级的 poppler\ 文件夹）
_PROJECT_DIR  = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
_POPPLER_DIR  = os.path.join(_PROJECT_DIR, "poppler", "Library", "bin")
_POPPLER_URL  = (
    "https://github.com/oschwartz10612/poppler-windows/releases/download/"
    "v24.02.0-0/Release-24.02.0-0.zip"
)


def _check_poppler():
    """返回 True 表示 poppler 可用（pdftoppm 在 PATH 或本地目录中）"""
    if os.path.isfile(os.path.join(_POPPLER_DIR, "pdftoppm.exe")):
        _inject_poppler_path()
        return True
    return shutil.which("pdftoppm") is not None


def _inject_poppler_path():
    """把本地 poppler/bin 加入当前进程 PATH"""
    cur = os.environ.get("PATH", "")
    if _POPPLER_DIR not in cur:
        os.environ["PATH"] = _POPPLER_DIR + os.pathsep + cur


def _install_poppler_windows(lbl, prog_root):
    """下载并解压 poppler-windows 到项目目录"""
    import urllib.request
    import zipfile
    import tempfile

    zip_path = os.path.join(tempfile.gettempdir(), "poppler_win.zip")
    dest_dir = os.path.join(_PROJECT_DIR, "poppler")

    lbl.config(text="正在下载 poppler（约 30MB）…")
    prog_root.update()
    try:
        urllib.request.urlretrieve(_POPPLER_URL, zip_path)
    except Exception as e:
        return False, str(e)

    lbl.config(text="正在解压 poppler…")
    prog_root.update()
    try:
        with zipfile.ZipFile(zip_path, "r") as zf:
            members = zf.namelist()
            prefix  = members[0].split("/")[0] + "/"
            for member in members:
                target = os.path.join(
                    dest_dir,
                    member[len(prefix):] if member.startswith(prefix) else member,
                )
                if member.endswith("/"):
                    os.makedirs(target, exist_ok=True)
                else:
                    os.makedirs(os.path.dirname(target), exist_ok=True)
                    with zf.open(member) as src, open(target, "wb") as dst:
                        dst.write(src.read())
        os.remove(zip_path)
    except Exception as e:
        return False, str(e)

    _inject_poppler_path()
    return True, ""


def ensure_deps():
    """检查并安装缺失的依赖，阻塞直到完成或用户取消。"""
    import tkinter as _tk
    from tkinter import messagebox as _mb

    # ── 检查 Python 包 ──
    missing_pkgs = []
    for mod, pkg in _REQUIRED.items():
        try:
            __import__(mod)
        except ImportError:
            missing_pkgs.append(pkg)

    # ── 检查 poppler（仅 Windows） ──
    need_poppler = (sys.platform == "win32") and (not _check_poppler())

    if not missing_pkgs and not need_poppler:
        _inject_poppler_path()
        return

    # ── 询问用户 ──
    lines = []
    if missing_pkgs:
        lines.append("缺少 Python 库：")
        lines += [f"  • {p}" for p in missing_pkgs]
    if need_poppler:
        if lines:
            lines.append("")
        lines.append("缺少系统工具：")
        lines.append("  • poppler（PDF渲染引擎，约 30MB）")
    msg = "\n".join(lines) + "\n\n点击「是」自动安装，完成后程序继续启动。"

    root = _tk.Tk(); root.withdraw()
    ok = _mb.askyesno("缺少依赖", msg, icon="warning")
    root.destroy()
    if not ok:
        sys.exit(0)

    # ── 进度窗口 ──
    prog_root = _tk.Tk()
    prog_root.title("正在安装依赖…")
    prog_root.geometry("460x130")
    prog_root.resizable(False, False)
    prog_root.attributes("-topmost", True)
    lbl = _tk.Label(prog_root, text="准备中…", font=("微软雅黑", 10),
                    wraplength=440, justify="left", padx=16, pady=10)
    lbl.pack(fill="x")
    bar_frame = _tk.Frame(prog_root, height=6, bg="#E0E0E0")
    bar_frame.pack(fill="x", padx=16)
    bar_fill = _tk.Frame(bar_frame, bg="#4A9FD4", height=6)
    bar_fill.place(x=0, y=0, relheight=1, relwidth=0)
    prog_root.update()

    total  = len(missing_pkgs) + (1 if need_poppler else 0)
    done   = 0
    failed = []

    for pkg in missing_pkgs:
        done += 1
        lbl.config(text=f"正在安装 {pkg}  （{done}/{total}）…")
        bar_fill.place(relwidth=done / total)
        prog_root.update()
        ret = subprocess.run(
            [sys.executable, "-m", "pip", "install", pkg, "--quiet"],
            capture_output=True,
        )
        if ret.returncode != 0:
            failed.append(pkg)

    if need_poppler:
        done += 1
        bar_fill.place(relwidth=done / total)
        ok2, err2 = _install_poppler_windows(lbl, prog_root)
        if not ok2:
            failed.append(f"poppler（{err2}）")

    prog_root.destroy()

    if failed:
        r2 = _tk.Tk(); r2.withdraw()
        _mb.showerror(
            "安装失败",
            "以下组件安装失败，请手动处理后重试：\n\n"
            + "\n".join(f"  • {p}" for p in failed)
            + "\n\npoppler 手动下载：\n"
            "  https://github.com/oschwartz10612/poppler-windows/releases",
        )
        r2.destroy()
        sys.exit(1)


def ensure_ocr_deps():
    """检查 OCR 依赖是否可用，缺失时弹窗询问安装。返回 True 表示就绪。"""
    import tkinter as _tk
    from tkinter import messagebox as _mb

    missing = []
    for mod, pkg in _OCR_REQUIRED.items():
        try:
            __import__(mod)
        except ImportError:
            missing.append(pkg)

    if not missing:
        return True

    root = _tk.Tk(); root.withdraw()
    ok = _mb.askyesno(
        "OCR 依赖",
        "首次使用 OCR 功能需要安装以下组件：\n\n"
        + "\n".join(f"  • {p}" for p in missing)
        + "\n\n（包含 PyTorch，约 200MB，安装可能需要几分钟）\n\n"
        "点击「是」自动安装。",
        icon="info",
    )
    root.destroy()
    if not ok:
        return False

    prog_root = _tk.Tk()
    prog_root.title("正在安装 OCR 组件…")
    prog_root.geometry("460x100")
    prog_root.resizable(False, False)
    prog_root.attributes("-topmost", True)
    lbl = _tk.Label(prog_root, text="", font=("微软雅黑", 10),
                    wraplength=440, justify="left", padx=16, pady=16)
    lbl.pack(fill="x")
    prog_root.update()

    failed = []
    for i, pkg in enumerate(missing):
        lbl.config(text=f"正在安装 {pkg}（{i+1}/{len(missing)}）…")
        prog_root.update()
        ret = subprocess.run(
            [sys.executable, "-m", "pip", "install", pkg, "--quiet"],
            capture_output=True,
        )
        if ret.returncode != 0:
            failed.append(pkg)

    prog_root.destroy()

    if failed:
        r2 = _tk.Tk(); r2.withdraw()
        _mb.showerror("安装失败",
                      "以下 OCR 组件安装失败：\n\n"
                      + "\n".join(f"  • {p}" for p in failed))
        r2.destroy()
        return False
    return True
