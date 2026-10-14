# -*- coding: utf-8 -*-
"""
Комплексные тесты агентных навыков (`agent_skills.py`) по всем 5 эталонным запросам АЦР,
а также проверка точечного редактирования, актуализации старой программы, масштабирования часов
и расчета ROI / токенов Cloud.ru.

При запуске автоматически генерирует и выгружает в папку `generated_output/` полные комплекты
(.docx, .pptx, .gift.txt, .moodle.xml, .csv, .json, .report.md) для ВСЕХ 5 открытых запросов кейса!
"""

import os
import unittest
from agent_skills import (
    CloudRuLLMClient,
    EducationalProgramAgent,
    GIGACOWORK_SKILLS_CATALOG,
)
from program_exporters import export_full_bundle
from test_program_validator import ensure_sample_248_docx


class TestAgentSkillsEndToEnd(unittest.TestCase):
    def setUp(self):
        self.agent = EducationalProgramAgent(CloudRuLLMClient())
        self.out_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "generated_output")
        os.makedirs(self.out_dir, exist_ok=True)

    def test_all_5_open_acr_queries_end_to_end(self):
        """Прогоняет все 5 открытых запросов методиста АЦР через полный пайплайн навыков и выгружает комплекты."""
        # --- ЗАПРОС № 1: Python для школьников (16 ч, ДООП) ---
        q1_text = (
            "Нужна дополнительная общеразвивающая программа «Первые шаги в Python» для школьников 8 и 9 классов "
            "без опыта программирования. Объем 16 часов, занятия очно по 2 часа два раза в неделю. "
            "Практики должно быть не меньше половины всех часов. В конце каждый ученик защищает небольшой проект, "
            "например простую игру или чат-бота. Нужны итоговый тест на 10 вопросов и слайды для первых двух занятий."
        )
        q1_spec, clar1, _ = self.agent.step1_analyze_request(q1_text)
        self.assertFalse(clar1.needs_clarification)
        self.assertEqual(q1_spec.program_type, "doop")
        skel1 = self.agent.step2_build_skeleton(q1_spec)
        kit1, rep1, roi1 = self.agent.step3_generate_full_content(skel1, q1_spec)
        self.assertTrue(rep1.is_valid, msg=rep1.to_markdown())
        self.assertEqual(kit1.total_hours, 16.0)
        self.assertEqual(len(kit1.test_questions), 10)
        self.assertEqual(len(kit1.lecture_decks), 2)
        self.assertGreater(roi1.manual_hours_saved, 50.0)
        export_full_bundle(kit1, self.out_dir, "case1_python_school_16h", query=q1_spec)

        # --- ЗАПРОС № 2: ИИ в работе учителя (36 ч, ДПП ПК, сервисы РФ, профстандарт Педагог) ---
        q2_text = (
            "Программа повышения квалификации «Искусственный интеллект в работе учителя» для учителей школ Ингушетии, "
            "36 часов, очно-заочно с применением дистанционных технологий. Учителя должны научиться готовить с помощью "
            "нейросетей планы уроков, задания и проверочные работы, а также понимать риски и правила безопасного "
            "использования ИИ в работе с детьми. Итоговая аттестация в форме зачёта по портфолио из трёх материалов, "
            "созданных с помощью ИИ. Использовать только сервисы, которые доступны в России."
        )
        q2_spec, clar2, _ = self.agent.step1_analyze_request(q2_text)
        self.assertFalse(clar2.needs_clarification)
        skel2 = self.agent.step2_build_skeleton(q2_spec)
        kit2, rep2, _ = self.agent.step3_generate_full_content(skel2, q2_spec)
        self.assertTrue(rep2.is_valid, msg=rep2.to_markdown())
        self.assertEqual(kit2.total_hours, 36.0)
        self.assertEqual(kit2.competencies[0].prof_standard_link.standard_code, "01.001")
        export_full_bundle(kit2, self.out_dir, "case2_ai_teacher_36h", query=q2_spec)

        # --- ЗАПРОС № 3: Анализ данных на Python для банков (72 ч, ДПП ПК, слайды для модуля про pandas) ---
        q3_text = (
            "Нужна программа повышения квалификации на 72 часа «Анализ данных на Python» для сотрудников региональных "
            "отделений банков, которые сейчас работают в Excel. Формат очно-заочный, срок 4 недели. Основные темы это "
            "pandas, визуализация, SQL для выгрузки данных и основы статистики. Итоговая аттестация в виде защиты проекта "
            "на обезличенных банковских данных. Слайды нужны для модуля про pandas."
        )
        q3_spec, clar3, _ = self.agent.step1_analyze_request(q3_text)
        self.assertFalse(clar3.needs_clarification)
        skel3 = self.agent.step2_build_skeleton(q3_spec)
        kit3, rep3, _ = self.agent.step3_generate_full_content(skel3, q3_spec)
        self.assertTrue(rep3.is_valid, msg=rep3.to_markdown())
        self.assertEqual(kit3.total_hours, 72.0)
        self.assertEqual([d.topic_id for d in kit3.lecture_decks], ["2.1", "2.2"])
        export_full_bundle(kit3, self.out_dir, "case3_bank_data_analysis_72h", query=q3_spec)

        # --- ЗАПРОС № 4: Курс по кибербезопасности для колледжа (сначала неполный -> вопросы -> ответы -> 72 ч) ---
        q4_incomplete = (
            "Сделайте курс по кибербезопасности для студентов колледжа, примерно на месяц, "
            "чтобы после него ребята могли пойти работать в эту сферу."
        )
        _, clar4_before, _ = self.agent.step1_analyze_request(q4_incomplete)
        self.assertTrue(clar4_before.needs_clarification)
        self.assertLessEqual(len(clar4_before.questions_for_methodologist), 5)
        self.assertIsNotNone(clar4_before.expectation_reality_gap_note)

        # Передаем ответы методиста из таблицы Запроса № 4
        answers4 = {
            "Тип программы": "Повышение квалификации",
            "Объём": "72 часа",
            "Слушатели": "Студенты 3 курса специальности «Информационные системы и программирование», знают основы сетей и Linux",
            "Форма обучения": "Очная, занятия на базе колледжа",
            "Итоговая аттестация": "Экзамен с практическим заданием на стенде",
        }
        q4_spec, clar4_after, _ = self.agent.step1_analyze_request(q4_incomplete, methodologist_answers=answers4)
        self.assertFalse(clar4_after.needs_clarification)
        skel4 = self.agent.step2_build_skeleton(q4_spec)
        kit4, rep4, _ = self.agent.step3_generate_full_content(skel4, q4_spec)
        self.assertTrue(rep4.is_valid, msg=rep4.to_markdown())
        self.assertEqual(kit4.total_hours, 72.0)
        self.assertEqual(kit4.competencies[0].prof_standard_link.standard_code, "06.033")
        self.assertTrue(any("начальный (базовый)" in n for n in kit4.methodologist_notes))
        export_full_bundle(kit4, self.out_dir, "case4_college_cybersec_72h", query=q4_spec)

        # --- ЗАПРОС № 5: Стажировка DevOps 256 ч + обнаружение ошибки в образце 248 ч ---
        q5_text = (
            "Нужна программа повышения квалификации в форме стажировки «DevOps-инженер от Linux до CI/CD» на 256 часов. "
            "Слушатели проходят стажировку в ИТ-компаниях региона, вся работа самостоятельная под руководством наставника "
            "от компании, аудиторных занятий нет. Итоговая аттестация занимает 2 часа и проходит в форме защиты отчёта о "
            "стажировке. Срок около двух месяцев, не больше 4 часов в день. Возьмите за образец структуры нашу стажировку по бизнес-анализу."
        )
        ref_248 = ensure_sample_248_docx()
        q5_spec, clar5, ref_defects = self.agent.step1_analyze_request(q5_text, reference_docx_paths=[ref_248])
        self.assertFalse(clar5.needs_clarification)
        self.assertGreater(len(ref_defects), 0)
        skel5 = self.agent.step2_build_skeleton(q5_spec)
        kit5, rep5, _ = self.agent.step3_generate_full_content(skel5, q5_spec, reference_docx_paths=[ref_248])
        self.assertTrue(rep5.is_valid, msg=rep5.to_markdown())
        self.assertEqual(kit5.total_hours, 256.0)
        self.assertEqual(kit5.declared_totals["self_study"], 254.0)
        self.assertEqual(kit5.declared_totals["practice"], 0.0)
        export_full_bundle(kit5, self.out_dir, "case5_devops_internship_256h", query=q5_spec, reference_docx_paths=[ref_248])

    def test_targeted_edit_and_legacy_modernization_and_adaptation(self):
        """Проверяет точечную правку (Шаг 7), актуализацию старой программы и адаптацию объема часов (Сверх плана)."""
        q3_text = "Программа повышения квалификации на 72 часа «Анализ данных на Python» для сотрудников региональных отделений банков..."
        q3_spec, _, _ = self.agent.step1_analyze_request(q3_text)
        skel3 = self.agent.step2_build_skeleton(q3_spec)
        kit3, _, _ = self.agent.step3_generate_full_content(skel3, q3_spec)

        # 1. Точечная замена темы 2.2 и усложнение теста
        edited_kit, edit_rep, edit_msg = self.agent.step4_targeted_edit(
            kit3,
            "Прошу заменить тему 2.2 на «Продвинутая агрегация временных рядов и скоринговых витрин в pandas» и усложнить тест",
            query=q3_spec,
        )
        self.assertTrue(edit_rep.is_valid, msg=edit_rep.to_markdown())
        self.assertIn("2.2", edit_msg)
        self.assertEqual(edited_kit.modules[1].topics[1].title, "Продвинутая агрегация временных рядов и скоринговых витрин в pandas")
        self.assertEqual(edited_kit.modules[1].topics[0].title, kit3.modules[1].topics[0].title)

        # 2. Актуализация старой программы (Сверх плана)
        edited_kit.software_and_tools.append("Python 3.7, ChatGPT и метод DataFrame.append()")
        mod_kit, mod_logs = self.agent.step5_modernize_legacy_program(edited_kit)
        self.assertTrue(any("Python 3.12" in l for l in mod_logs))
        self.assertTrue(any("GigaChat" in l for l in mod_logs))

        # 3. Адаптация программы на новый объем 144 часа (Сверх плана)
        adapted_kit, adapt_rep, _ = self.agent.step6_adapt_program_volume_or_audience(
            mod_kit,
            new_total_hours=144.0,
            new_target_audience="Ведущие аналитики департамента рисков банка",
        )
        self.assertTrue(adapt_rep.is_valid, msg=adapt_rep.to_markdown())
        self.assertEqual(adapted_kit.total_hours, 144.0)


def run_skills_demo() -> None:
    agent = EducationalProgramAgent(CloudRuLLMClient())
    print("=" * 80)
    print(f"КАТАЛОГ НАВЫКОВ АГЕНТА (СОВМЕСТИМЫЙ С GIGACOWORK): {len(GIGACOWORK_SKILLS_CATALOG)} НАВЫКОВ")
    for s in GIGACOWORK_SKILLS_CATALOG:
        hitl = " [Human-in-the-Loop]" if s.human_in_the_loop_gate else ""
        print(f"  * {s.skill_id}: {s.title}{hitl}")
    print("-" * 80)

    q3_text = (
        "Нужна программа повышения квалификации на 72 часа «Анализ данных на Python» для сотрудников региональных "
        "отделений банков, которые сейчас работают в Excel. Формат очно-заочный, срок 4 недели. Основные темы это "
        "pandas, визуализация, SQL для выгрузки данных и основы статистики. Итоговая аттестация в виде защиты проекта "
        "на обезличенных банковских данных. Слайды нужны для модуля про pandas."
    )
    q3_spec, _, _ = agent.step1_analyze_request(q3_text)
    skel3 = agent.step2_build_skeleton(q3_spec)
    kit3, rep3, roi3 = agent.step3_generate_full_content(skel3, q3_spec)
    print(f"Демонстрация сборки программы № 3: «{kit3.program_title}» ({kit3.total_hours:g} ак. ч.)")
    print(roi3.to_markdown())
    print("=" * 80 + "\n")


if __name__ == "__main__":
    run_skills_demo()
    unittest.main(verbosity=2)