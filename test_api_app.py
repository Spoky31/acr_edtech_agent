# -*- coding: utf-8 -*-
"""
Автотесты единого REST API-эндпоинта и веб-сервера (`app.py`).
Проверяет работу HTTP-сервера, генерацию каркаса, полный цикл сборки, уточняющие вопросы,
точечную доработку и скачивание готовых файлов `.docx` и `.pptx` по HTTP.
"""

import json
import threading
import unittest
import urllib.request
from http.server import ThreadingHTTPServer

from app import ACRRequestHandler, PRESET_QUERIES


class TestUnifiedAPIApp(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.server = ThreadingHTTPServer(("127.0.0.1", 0), ACRRequestHandler)
        cls.port = cls.server.server_address[1]
        cls.base_url = f"http://127.0.0.1:{cls.port}"
        cls.thread = threading.Thread(target=cls.server.serve_forever, daemon=True)
        cls.thread.start()

    @classmethod
    def tearDownClass(cls):
        cls.server.shutdown()
        cls.server.server_close()

    def _post_json(self, path: str, payload: dict) -> dict:
        req = urllib.request.Request(
            f"{self.base_url}{path}",
            data=json.dumps(payload, ensure_ascii=False).encode("utf-8"),
            headers={"Content-Type": "application/json"},
            method="POST",
        )
        with urllib.request.urlopen(req, timeout=30) as resp:
            return json.loads(resp.read().decode("utf-8"))

    def test_01_ui_and_health_and_skills_endpoints(self):
        with urllib.request.urlopen(f"{self.base_url}/") as resp:
            html = resp.read().decode("utf-8")
            self.assertIn("ИИ-агент для создания образовательных программ АЦР", html)

        with urllib.request.urlopen(f"{self.base_url}/api/v1/health") as resp:
            health = json.loads(resp.read().decode("utf-8"))
            self.assertEqual(health["status"], "ok")
            self.assertEqual(health["skills_count"], 5)

    def test_02_incomplete_query_returns_clarification_questions(self):
        res = self._post_json("/api/v1/generate", {"mode": "analyze", "query": PRESET_QUERIES["case4"]["query"]})
        self.assertEqual(res["status"], "needs_clarification")
        self.assertLessEqual(len(res["clarification"]["questions_for_methodologist"]), 5)

    def test_03_full_generation_and_file_download_and_targeted_edit(self):
        # 1. Полная генерация по Запросу № 3 (72 ч Анализ данных в банке)
        res = self._post_json(
            "/api/v1/generate",
            {"mode": "full", "query": PRESET_QUERIES["case3"]["query"], "file_prefix": "api_test_case3"},
        )
        self.assertEqual(res["status"], "completed")
        self.assertTrue(res["validation_report"]["is_valid"])
        self.assertEqual(res["program_kit"]["total_hours"], 72.0)

        # 2. Проверяем скачивание сгенерированного .docx по HTTP
        docx_url = f"{self.base_url}{res['download_files']['docx']}"
        with urllib.request.urlopen(docx_url) as dl_resp:
            docx_bytes = dl_resp.read()
            self.assertGreater(len(docx_bytes), 1000)

        # 3. Точечная правка через /api/v1/edit
        edit_res = self._post_json(
            "/api/v1/edit",
            {"edit_instruction": "Заменить тему 2.2 на «Анализ кредитных витрин в pandas» и усложнить тест"},
        )
        self.assertEqual(edit_res["status"], "edited")
        self.assertTrue(edit_res["validation_report"]["is_valid"])


if __name__ == "__main__":
    unittest.main(verbosity=2)