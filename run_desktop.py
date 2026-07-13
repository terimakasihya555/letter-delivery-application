import socket
import threading
import webbrowser
from waitress import serve

from app import app


def find_free_port(start_port=5000):
    for port in range(start_port, start_port + 50):
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
            try:
                s.bind(("127.0.0.1", port))
                return port
            except OSError:
                continue

    return start_port


def open_browser(port):
    url = f"http://127.0.0.1:{port}"
    webbrowser.open(url)


if __name__ == "__main__":
    port = find_free_port(5000)

    threading.Timer(1.5, open_browser, args=(port,)).start()

    serve(
        app,
        host="127.0.0.1",
        port=port
    )