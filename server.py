"""
Здрасте KPI — веб-сервер (локально и на Railway).
"""

import os
import sys
import threading
import time
import webbrowser
from http.server import BaseHTTPRequestHandler, HTTPServer
from pathlib import Path
from datetime import datetime

sys.path.insert(0, str(Path(__file__).parent))

PORT = int(os.environ.get("PORT", 8765))
IS_RAILWAY = "RAILWAY_ENVIRONMENT" in os.environ

loading_html = """<!DOCTYPE html>
<html lang="ru">
<head>
<meta charset="UTF-8">
<meta http-equiv="refresh" content="3">
<title>Здрасте — загрузка...</title>
<style>
  body { background: #FBF4EC; display: flex; align-items: center;
         justify-content: center; height: 100vh; margin: 0;
         font-family: Arial, sans-serif; }
  .box { text-align: center; }
  .logo { font-size: 32px; font-weight: 900; color: #B85C38;
          letter-spacing: 2px; margin-bottom: 16px; }
  .spinner { width: 48px; height: 48px; border: 4px solid #F2E0CC;
             border-top-color: #B85C38; border-radius: 50%;
             animation: spin 0.8s linear infinite; margin: 0 auto 16px; }
  @keyframes spin { to { transform: rotate(360deg); } }
  .msg { color: #9C7B6A; font-size: 14px; }
</style>
</head>
<body>
<div class="box">
  <div class="logo">ЗДРАСТЕ</div>
  <div class="spinner"></div>
  <div class="msg">Загружаем данные из iiko...<br>Страница обновится автоматически</div>
</div>
</body>
</html>"""


class ReportHandler(BaseHTTPRequestHandler):
    _cache = None
    _cache_time = 0
    _cache_lock = threading.Lock()
    _loading = False

    def log_message(self, format, *args):
        t = datetime.now().strftime("%H:%M:%S")
        print(f"[{t}] {format % args}")

    def do_GET(self):
        if self.path not in ("/", "/report"):
            self.send_response(404)
            self.end_headers()
            return

        with ReportHandler._cache_lock:
            loading = ReportHandler._loading

        if loading:
            self._serve_html(loading_html)
            return

        now = time.time()
        with ReportHandler._cache_lock:
            stale = (now - ReportHandler._cache_time) > 60
            has_cache = ReportHandler._cache is not None

        if stale and not loading:
            with ReportHandler._cache_lock:
                ReportHandler._loading = True
            thread = threading.Thread(target=self._load_report, daemon=True)
            thread.start()
            if not has_cache:
                self._serve_html(loading_html)
                return

        with ReportHandler._cache_lock:
            html = ReportHandler._cache

        self._serve_html(html if html else loading_html)

    def _serve_html(self, html: str):
        data = html.encode("utf-8")
        self.send_response(200)
        self.send_header("Content-Type", "text/html; charset=utf-8")
        self.send_header("Content-Length", str(len(data)))
        self.end_headers()
        self.wfile.write(data)

    @classmethod
    def _load_report(cls):
        try:
            print("[INFO] Загрузка данных...")
            from iiko_report import generate_report_html
            html = generate_report_html()
            with cls._cache_lock:
                cls._cache = html
                cls._cache_time = time.time()
                cls._loading = False
            print(f"[OK] Отчёт обновлён в {datetime.now().strftime('%H:%M:%S')}")
        except Exception as e:
            print(f"[ERROR] {e}")
            with cls._cache_lock:
                cls._loading = False


def main():
    print(f"\n{'='*54}")
    print(f"  Здрасте KPI — сервер запущен")
    print(f"  Порт: {PORT}")
    print(f"{'='*54}\n")

    ReportHandler._loading = True
    thread = threading.Thread(target=ReportHandler._load_report, daemon=True)
    thread.start()

    if not IS_RAILWAY:
        def open_browser():
            time.sleep(2)
            webbrowser.open(f"http://localhost:{PORT}")
        threading.Thread(target=open_browser, daemon=True).start()

    server = HTTPServer(("0.0.0.0", PORT), ReportHandler)
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("\n[INFO] Сервер остановлен")


if __name__ == "__main__":
    main()
