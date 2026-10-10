"""
Здрасте KPI — небольшой HTTP-сервер для локального запуска и Render.
"""

import base64
import binascii
import hmac
import json
import os
import threading
import time
import webbrowser
from datetime import datetime, timedelta
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import urlsplit
from zoneinfo import ZoneInfo

import sys

sys.path.insert(0, str(Path(__file__).parent))

APP_TIMEZONE = ZoneInfo(os.environ.get("APP_TIMEZONE", "Europe/Moscow"))
PORT = int(os.environ.get("PORT", "8765"))
CACHE_TTL_SECONDS = max(60, int(os.environ.get("REPORT_CACHE_TTL_SECONDS", "900")))
DAILY_REFRESH_HOUR = int(os.environ.get("REPORT_DAILY_REFRESH_HOUR", "23"))
OPEN_BROWSER = os.environ.get("OPEN_BROWSER", "false").strip().lower() in {
    "1", "true", "yes", "on",
}
REPORT_USERNAME = os.environ.get("REPORT_USERNAME", "")
REPORT_PASSWORD = os.environ.get("REPORT_PASSWORD", "")

if not 0 <= DAILY_REFRESH_HOUR <= 23:
    raise RuntimeError("REPORT_DAILY_REFRESH_HOUR должен быть от 0 до 23")
if bool(REPORT_USERNAME) != bool(REPORT_PASSWORD):
    raise RuntimeError(
        "REPORT_USERNAME и REPORT_PASSWORD должны быть заданы вместе"
    )


def _validate_render_configuration():
    """Не заменять рабочий deploy неполностью настроенной версией."""
    if os.environ.get("RENDER", "").strip().lower() not in {
        "1", "true", "yes", "on",
    }:
        return

    required = ("IIKO_SERVER", "IIKO_LOGIN", "IIKO_PASSWORD")
    missing = [name for name in required if not os.environ.get(name)]
    database_url = os.environ.get("DATABASE_URL")
    pg_complete = all(
        os.environ.get(name)
        for name in ("PG_HOST", "PG_DB", "PG_USER", "PG_PASS")
    )
    if not database_url and not pg_complete:
        missing.append("DATABASE_URL или полный набор PG_*")
    if not os.environ.get("BARISTA_NAMES") and not os.environ.get("WAITER_NAMES"):
        missing.append("BARISTA_NAMES или WAITER_NAMES")
    if missing:
        raise RuntimeError(
            "Render не настроен: отсутствуют обязательные переменные окружения: "
            + ", ".join(missing)
        )

LOADING_HTML = """<!doctype html>
<html lang="ru">
<head>
<meta charset="utf-8">
<meta http-equiv="refresh" content="5">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Здрасте — загрузка</title>
<style>
body { background:#fbf4ec; display:flex; align-items:center; justify-content:center;
height:100vh; margin:0; font-family:Arial,sans-serif; color:#6b3e26; }
.box { text-align:center; padding:24px; }
.logo { font-size:32px; font-weight:900; color:#b85c38; letter-spacing:2px; }
.spinner { width:48px; height:48px; border:4px solid #f2e0cc;
border-top-color:#b85c38; border-radius:50%; animation:spin .8s linear infinite;
margin:20px auto; }
@keyframes spin { to { transform:rotate(360deg); } }
</style>
</head>
<body><main class="box"><div class="logo">ЗДРАСТЕ</div>
<div class="spinner"></div><p>Загружаем данные…</p></main></body>
</html>"""

ERROR_HTML = """<!doctype html>
<html lang="ru">
<head><meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1">
<title>Здрасте — данные недоступны</title>
<style>body{font-family:Arial,sans-serif;background:#fbf4ec;color:#6b3e26;
max-width:680px;margin:15vh auto;padding:24px}h1{color:#b85c38}</style></head>
<body><h1>Данные временно недоступны</h1>
<p>Предыдущая версия отчёта ещё не создана. Повторите попытку позже.</p></body>
</html>"""


def _now() -> datetime:
    return datetime.now(APP_TIMEZONE)


class ReportHandler(BaseHTTPRequestHandler):
    server_version = "ZdrasteKPI"
    sys_version = ""

    _cache = None
    _cache_time = 0.0
    _cache_lock = threading.Lock()
    _loading = False
    _last_error_at = None

    def log_message(self, format_string, *args):
        timestamp = _now().strftime("%Y-%m-%d %H:%M:%S %z")
        print("[{}] {}".format(timestamp, format_string % args))

    @classmethod
    def start_refresh(cls) -> bool:
        """Атомарно запустить одно фоновое обновление."""
        with cls._cache_lock:
            if cls._loading:
                return False
            cls._loading = True
        threading.Thread(target=cls._load_report, daemon=True).start()
        return True

    def do_GET(self):
        path = urlsplit(self.path).path

        if path == "/healthz":
            with ReportHandler._cache_lock:
                payload = {
                    "status": "ok",
                    "ready": ReportHandler._cache is not None,
                    "refreshing": ReportHandler._loading,
                    "last_error_at": ReportHandler._last_error_at,
                }
            self._serve_json(payload)
            return

        if path == "/favicon.ico":
            self.send_response(204)
            self._send_security_headers()
            self.end_headers()
            return

        if not self._is_authorized():
            self.send_response(401)
            self.send_header("WWW-Authenticate", 'Basic realm="Zdraste KPI", charset="UTF-8"')
            self._send_security_headers()
            self.end_headers()
            return

        if path not in ("/", "/report"):
            self.send_response(404)
            self._send_security_headers()
            self.end_headers()
            return

        now = time.time()
        with ReportHandler._cache_lock:
            cached_html = ReportHandler._cache
            stale = now - ReportHandler._cache_time > CACHE_TTL_SECONDS
            loading = ReportHandler._loading

        if stale and not loading:
            ReportHandler.start_refresh()

        if cached_html is not None:
            self._serve_html(cached_html)
            return

        with ReportHandler._cache_lock:
            loading = ReportHandler._loading
        self._serve_html(LOADING_HTML if loading else ERROR_HTML, status=200 if loading else 503)

    def _is_authorized(self) -> bool:
        if not REPORT_USERNAME:
            return True
        value = self.headers.get("Authorization", "")
        if not value.startswith("Basic "):
            return False
        try:
            decoded = base64.b64decode(value[6:], validate=True).decode("utf-8")
        except (binascii.Error, UnicodeDecodeError):
            return False
        username, separator, password = decoded.partition(":")
        if not separator:
            return False
        return hmac.compare_digest(username, REPORT_USERNAME) and hmac.compare_digest(
            password, REPORT_PASSWORD
        )

    def _send_security_headers(self):
        self.send_header("Cache-Control", "private, no-store, max-age=0")
        self.send_header("Pragma", "no-cache")
        self.send_header("X-Content-Type-Options", "nosniff")
        self.send_header("Referrer-Policy", "no-referrer")
        self.send_header("Permissions-Policy", "camera=(), microphone=(), geolocation=()")
        self.send_header(
            "Content-Security-Policy",
            "default-src 'none'; style-src 'unsafe-inline' https://fonts.googleapis.com; "
            "font-src https://fonts.gstatic.com; img-src data:; frame-ancestors 'none'; "
            "base-uri 'none'; form-action 'none'",
        )

    def _serve_html(self, html_text: str, status: int = 200):
        data = html_text.encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "text/html; charset=utf-8")
        self.send_header("Content-Length", str(len(data)))
        self._send_security_headers()
        self.end_headers()
        self.wfile.write(data)

    def _serve_json(self, payload: dict):
        data = json.dumps(payload, ensure_ascii=False).encode("utf-8")
        self.send_response(200)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(data)))
        self._send_security_headers()
        self.end_headers()
        self.wfile.write(data)

    @classmethod
    def _load_report(cls):
        try:
            print("[INFO] Загрузка данных…")
            from iiko_report import generate_report_html

            html_text = generate_report_html()
            if not html_text:
                raise RuntimeError("Генератор вернул пустой отчёт")
            with cls._cache_lock:
                cls._cache = html_text
                cls._cache_time = time.time()
                cls._last_error_at = None
            print("[OK] Отчёт обновлён в {}".format(_now().strftime("%H:%M:%S")))
        except Exception as exc:
            # Не печатаем URL/параметры исключения: там могут быть токены.
            with cls._cache_lock:
                cls._last_error_at = _now().isoformat()
            print("[ERROR] Обновление отчёта: {}".format(type(exc).__name__))
        finally:
            with cls._cache_lock:
                cls._loading = False


class ReportServer(ThreadingHTTPServer):
    daemon_threads = True
    allow_reuse_address = True


def _daily_refresh_loop():
    while True:
        now = _now()
        target = now.replace(
            hour=DAILY_REFRESH_HOUR,
            minute=0,
            second=0,
            microsecond=0,
        )
        if target <= now:
            target += timedelta(days=1)
        delay = max(1.0, (target - now).total_seconds())
        if threading.Event().wait(delay):
            return
        ReportHandler.start_refresh()


def main():
    _validate_render_configuration()
    print("\n======================================================")
    print("  Здрасте KPI — сервер запущен")
    print("  Порт: {}".format(PORT))
    print("  Часовой пояс: {}".format(APP_TIMEZONE.key))
    print("======================================================\n")

    ReportHandler.start_refresh()
    threading.Thread(target=_daily_refresh_loop, daemon=True).start()

    if OPEN_BROWSER:
        def open_browser():
            time.sleep(2)
            webbrowser.open("http://localhost:{}".format(PORT))

        threading.Thread(target=open_browser, daemon=True).start()

    server = ReportServer(("0.0.0.0", PORT), ReportHandler)
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("\n[INFO] Сервер остановлен")
    finally:
        server.server_close()


if __name__ == "__main__":
    main()
