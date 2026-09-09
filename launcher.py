import socket
import sys
import threading
import webbrowser
from pathlib import Path

from streamlit.web import bootstrap


def find_free_port(start: int = 8501, end: int = 8599) -> int:
    for port in range(start, end + 1):
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as sock:
            sock.settimeout(0.2)
            if sock.connect_ex(("127.0.0.1", port)) != 0:
                return port
    raise RuntimeError("没有可用端口，请关闭占用 8501-8599 端口的程序后重试。")


def run():
    base_dir = Path(__file__).resolve().parent
    candidates = [
        base_dir / "app.py",
        base_dir / "_internal" / "app.py",
    ]
    if getattr(sys, "frozen", False):
        meipass = Path(getattr(sys, "_MEIPASS", ""))
        if str(meipass):
            candidates.insert(0, meipass / "app.py")

    app_path = next((p for p in candidates if p.exists()), None)
    if app_path is None:
        raise FileNotFoundError("未找到 app.py，请检查打包产物完整性。")

    port = find_free_port()
    url = f"http://127.0.0.1:{port}"
    flags = {
        "server.headless": True,
        "server.port": port,
        "browser.gatherUsageStats": False,
    }
    threading.Timer(2.0, lambda: webbrowser.open(url)).start()
    bootstrap.run(str(app_path), False, [], flags)


if __name__ == "__main__":
    run()
