import base64
import os
import threading
import time
import unittest
import urllib.request
from email.message import Message
from unittest import mock

import server


class AuthorizationTests(unittest.TestCase):
    def make_handler(self, authorization=None):
        handler = server.ReportHandler.__new__(server.ReportHandler)
        headers = Message()
        if authorization:
            headers["Authorization"] = authorization
        handler.headers = headers
        return handler

    def test_auth_is_optional_when_not_configured(self):
        handler = self.make_handler()
        with mock.patch.object(server, "REPORT_USERNAME", ""), mock.patch.object(
            server, "REPORT_PASSWORD", ""
        ):
            self.assertTrue(handler._is_authorized())

    def test_valid_basic_auth(self):
        token = base64.b64encode("user:secret".encode()).decode()
        handler = self.make_handler("Basic " + token)
        with mock.patch.object(server, "REPORT_USERNAME", "user"), mock.patch.object(
            server, "REPORT_PASSWORD", "secret"
        ):
            self.assertTrue(handler._is_authorized())

    def test_invalid_basic_auth(self):
        token = base64.b64encode("user:wrong".encode()).decode()
        handler = self.make_handler("Basic " + token)
        with mock.patch.object(server, "REPORT_USERNAME", "user"), mock.patch.object(
            server, "REPORT_PASSWORD", "secret"
        ):
            self.assertFalse(handler._is_authorized())


class ProductionConfigurationTests(unittest.TestCase):
    def test_local_run_does_not_require_render_environment(self):
        with mock.patch.dict(os.environ, {"RENDER": ""}, clear=False):
            server._validate_render_configuration()

    def test_incomplete_render_environment_fails_before_server_start(self):
        names = (
            "IIKO_SERVER", "IIKO_LOGIN", "IIKO_PASSWORD", "DATABASE_URL",
            "PG_HOST", "PG_DB", "PG_USER", "PG_PASS", "BARISTA_NAMES",
            "WAITER_NAMES",
        )
        values = {name: "" for name in names}
        values["RENDER"] = "true"
        with mock.patch.dict(os.environ, values, clear=False):
            with self.assertRaisesRegex(RuntimeError, "Render не настроен"):
                server._validate_render_configuration()

    def test_complete_render_environment_passes(self):
        values = {
            "RENDER": "true",
            "IIKO_SERVER": "https://iiko.example/resto",
            "IIKO_LOGIN": "user",
            "IIKO_PASSWORD": "secret",
            "DATABASE_URL": "postgresql://configured",
            "BARISTA_NAMES": "Тестовый Сотрудник",
        }
        with mock.patch.dict(os.environ, values, clear=False):
            server._validate_render_configuration()


class HttpSmokeTests(unittest.TestCase):
    def test_cached_report_and_health_endpoint(self):
        original_state = (
            server.ReportHandler._cache,
            server.ReportHandler._cache_time,
            server.ReportHandler._loading,
            server.ReportHandler._last_error_at,
        )
        httpd = None
        try:
            server.ReportHandler._cache = "<h1>ok</h1>"
            server.ReportHandler._cache_time = time.time()
            server.ReportHandler._loading = False
            server.ReportHandler._last_error_at = None
            httpd = server.ReportServer(("127.0.0.1", 0), server.ReportHandler)
            thread = threading.Thread(target=httpd.serve_forever, daemon=True)
            thread.start()
            base_url = "http://127.0.0.1:{}".format(httpd.server_port)

            with urllib.request.urlopen(base_url + "/", timeout=3) as response:
                self.assertEqual(response.read(), b"<h1>ok</h1>")
                self.assertEqual(response.headers["X-Content-Type-Options"], "nosniff")
                self.assertIn("no-store", response.headers["Cache-Control"])

            with urllib.request.urlopen(base_url + "/healthz", timeout=3) as response:
                body = response.read().decode("utf-8")
                self.assertIn('"ready": true', body)
        finally:
            if httpd is not None:
                httpd.shutdown()
                httpd.server_close()
            (
                server.ReportHandler._cache,
                server.ReportHandler._cache_time,
                server.ReportHandler._loading,
                server.ReportHandler._last_error_at,
            ) = original_state


if __name__ == "__main__":
    unittest.main()

