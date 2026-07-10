import os
import sys
from streamlit.web import cli as stcli

def resource_path(relative_path):
    base_path = getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base_path, relative_path)

def _print_startup_diagnostics():
    """Console-only diagnostic to help debug packaged export failures.

    Never raises - a broken diagnostic must not prevent Streamlit from
    starting.
    """
    try:
        print("[DataBridge] sys.executable =", sys.executable)
        print("[DataBridge] frozen =", bool(getattr(sys, 'frozen', False)))
        print("[DataBridge] sys._MEIPASS =", getattr(sys, '_MEIPASS', '<not frozen>'))
        for mod in ("docx", "reportlab", "openpyxl", "matplotlib", "plotly", "kaleido"):
            try:
                __import__(mod)
                print(f"[DataBridge] import {mod}: OK")
            except Exception as exc:
                print(f"[DataBridge] import {mod}: FAILED ({exc})")
    except Exception as exc:
        print("[DataBridge] startup diagnostics failed:", exc)

if __name__ == "__main__":
    _print_startup_diagnostics()
    sys.argv = [
        "streamlit", "run", resource_path("app.py"),
        "--global.developmentMode=false",
        "--server.headless=true",
        "--server.port=8504",
    ]
    sys.exit(stcli.main())
