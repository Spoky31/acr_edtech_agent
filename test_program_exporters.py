# -*- coding: utf-8 -*-
"""
Тесты и демонстрация генераторов выгрузки комплекта (`program_exporters.py`).

При запуске создает в папке `generated_output/` полные комплекты файлов для:
1. Запроса № 1 (ДООП «Первые шаги в Python», 16 ч) — `.docx`, `.pptx`, `.gift.txt`, `.moodle.xml`, `.csv`, `.json`, `.md`.
2. Запроса № 5 (ДПП ПК Стажировка «DevOps-инженер от Linux до CI/CD», 256 ч) — `.docx`, `.pptx`, `.json`, `.md`.

Затем прогоняет созданные `.docx` файлы через наш инспектор `inspect_reference_docx()`,
доказывая, что в сгенерированных Word-документах 0 расхождений часов между УП, УТП и РПД.
"""

import json
import os
import unittest
import xml.etree.ElementTree as ET

from program_exporters import export_full_bundle
from program_validator import inspect_reference_docx
from test_program_validator import (
    build_valid_query1_doop_16h,
    build_valid_query5_internship_256h,
    ensure_sample_248_docx,
)


class TestProgramExporters(unittest.TestCase):
    def setUp(self):
        self.out_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "generated_output")
        os.makedirs(self.out_dir, exist_ok=True)

    def test_export_query1_doop_bundle_and_verify_zero_docx_defects(self):
        """Генерирует полный комплект для Запроса № 1 (ДООП 16 ч) и проверяет все выходные файлы."""
        q1, kit1 = build_valid_query1_doop_16h()
        paths = export_full_bundle(
            kit=kit1,
            output_dir=self.out_dir,
            file_prefix="query1_python_16h_doop",
            query=q1,
        )

        for key, path in paths.items():
            self.assertTrue(os.path.exists(path), f"Файл {key} не создан: {path}")
            self.assertGreater(os.path.getsize(path), 50, f"Файл {key} пуст: {path}")

        # Проверяем, что инспектор таблиц .docx находит ровно 0 ошибок в сгенерированном документе Word!
        docx_issues = inspect_reference_docx(paths["docx"])
        self.assertEqual(
            len(docx_issues),
            0,
            msg=f"В сгенерированном .docx обнаружены расхождения: {docx_issues}",
        )

        # Проверяем валидность Moodle XML
        tree = ET.parse(paths["moodle_xml"])
        root = tree.getroot()
        self.assertEqual(root.tag, "quiz")
        multichoice_qs = root.findall("./question[@type='multichoice']")
        self.assertEqual(len(multichoice_qs), 10)

        # Проверяем наличие всех 10 вопросов в GIFT
        with open(paths["gift"], "r", encoding="utf-8") as f:
            gift_text = f.read()
        self.assertEqual(gift_text.count("::Q"), 10)

        # Проверяем валидность JSON
        with open(paths["json"], "r", encoding="utf-8") as f:
            payload = json.load(f)
        self.assertTrue(payload["validation_report"]["is_valid"])
        self.assertEqual(payload["program_kit"]["total_hours"], 16.0)

    def test_export_query5_dpp_pk_internship_bundle_and_verify_zero_docx_defects(self):
        """Генерирует полный комплект для Запроса № 5 (Стажировка DevOps 256 ч) и проверяет .docx на отсутствие ошибки образца."""
        q5, kit5 = build_valid_query5_internship_256h()
        ref_248 = ensure_sample_248_docx()
        paths = export_full_bundle(
            kit=kit5,
            output_dir=self.out_dir,
            file_prefix="query5_devops_256h_internship",
            query=q5,
            reference_docx_paths=[ref_248],
        )

        for key, path in paths.items():
            self.assertTrue(os.path.exists(path), f"Файл {key} не создан: {path}")

        # Главная проверка: в сгенерированном нами .docx для стажировки 256 ч 0 ошибок УП/УТП/РПД!
        docx_issues = inspect_reference_docx(paths["docx"])
        self.assertEqual(
            len(docx_issues),
            0,
            msg=f"В сгенерированном .docx стажировки найдены ошибки: {docx_issues}",
        )


def run_export_demo() -> None:
    out_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "generated_output")
    q1, kit1 = build_valid_query1_doop_16h()
    p1 = export_full_bundle(kit1, out_dir, "query1_python_16h_doop", query=q1)

    q5, kit5 = build_valid_query5_internship_256h()
    p5 = export_full_bundle(
        kit5,
        out_dir,
        "query5_devops_256h_internship",
        query=q5,
        reference_docx_paths=[ensure_sample_248_docx()],
    )

    print("=" * 80)
    print("ВЫГРУЗКА КОМПЛЕКТОВ ПРОГРАММ ЗАВЕРШЕНА УСПЕШНО")
    print(f"Папка с готовыми файлами: {out_dir}")
    print("-" * 80)
    print("Комплект № 1 (ДООП 16 ч · Первые шаги в Python):")
    for k, v in p1.items():
        print(f"  - {k:10s}: {os.path.basename(v)}")
    print("Комплект № 5 (ДПП ПК Стажировка 256 ч · DevOps-инженер):")
    for k, v in p5.items():
        print(f"  - {k:10s}: {os.path.basename(v)}")
    print("=" * 80 + "\n")


if __name__ == "__main__":
    run_export_demo()
    unittest.main(verbosity=2)