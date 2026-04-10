# ── 隐藏控制台黑窗口，压制第三方库警告输出 ──────────────────────────────────
import sys, os

if sys.platform == "win32":
    try:
        import ctypes
        ctypes.windll.kernel32.FreeConsole()
    except Exception:
        pass

class _NullWriter:
    def write(self, *a): pass
    def flush(self): pass

if not sys.flags.debug and not os.environ.get("PDF2WORD_DEBUG"):
    sys.stdout = _NullWriter()
    sys.stderr = _NullWriter()

import warnings
warnings.filterwarnings("ignore")

# ── 检查并安装依赖 ────────────────────────────────────────────────────────────
from core.deps import ensure_deps
ensure_deps()

# ── 启动应用 ──────────────────────────────────────────────────────────────────
from core.app import App

if __name__ == "__main__":
    app = App()
    app.mainloop()
