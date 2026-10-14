# -*- coding: utf-8 -*-
"""
Набор модульных и сценарных тестов для проверки детерминированного валидатора (`program_validator.py`)
на всех 5 эталонных запросах АЦР и реальных файлах-образцах (.docx).
"""

import os
import unittest
from typing import Tuple

from program_validator import (
    CalendarScheduleSpec,
    CompetencyItem,
    CurriculumModule,
    CurriculumTopic,
    FinalAttestationSpec,
    InterviewQuestion,
    LearningOutcome,
    LiteratureSource,
    ProfStandardLink,
    ProgramKit,
    ProgramValidator,
    RPDActivity,
    RequisitesPlaceholders,
    SeminarTask,
    SlideDeckSpec,
    TestQuestion,
    UserQuerySpec,
    analyze_methodologist_query,
    auto_fix_hours,
    inspect_reference_docx,
)


def build_valid_query1_doop_16h() -> Tuple[UserQuerySpec, ProgramKit]:
    """Запрос № 1: «Первые шаги в Python» для школьников (ДООП, 16 ч, практика >= 8 ч, 10 вопросов, слайды 1 и 2)."""
    query = UserQuerySpec(
        raw_text="Нужна дополнительная общеразвивающая программа «Первые шаги в Python» для школьников 8 и 9 классов без опыта программирования. Объем 16 часов...",
        program_type="doop",
        total_hours=16.0,
        target_audience="Школьники 8–9 классов без опыта программирования",
        entry_level="zero",
        study_form="ochnaya",
        final_attestation_form="Защита индивидуального мини-проекта",
        min_practice_share=0.5,
        planned_weeks=4,
        sessions_per_week=2,
        hours_per_session=2.0,
        required_test_questions_count=10,
    )

    outcomes = [
        LearningOutcome("LO_K1", "knowledge", "Знание базовых понятий алгоритмизации, переменных и типов данных в Python"),
        LearningOutcome("LO_S1", "skill", "Умение применять условные операторы и циклы для решения простых задач"),
        LearningOutcome("LO_H1", "habit", "Навык создания и отладки простой консольной игры или чат-бота на Python"),
    ]

    topics = [
        CurriculumTopic(
            id="1.1",
            title="Знакомство с Python: переменные, ввод-вывод и типы данных",
            total_hours=4.0,
            lecture_hours=2.0,
            practice_hours=2.0,
            learning_outcome_ids=["LO_K1"],
            rpd_activities=[
                RPDActivity("ЛК", 2.0, "Понятие алгоритма и переменной (простыми словами: «коробка с наклейкой для хранения данных»)."),
                RPDActivity("ПР", 2.0, "Написание первой программы-анкеты с input() и print()."),
            ],
        ),
        CurriculumTopic(
            id="1.2",
            title="Логика программы: условия if-else и циклы while/for",
            total_hours=6.0,
            lecture_hours=2.0,
            practice_hours=4.0,
            learning_outcome_ids=["LO_S1"],
            rpd_activities=[
                RPDActivity("ЛК", 2.0, "Ветвления и повторения в программах."),
                RPDActivity("ПР", 4.0, "Создание мини-викторины и игры «Угадай число»."),
            ],
        ),
        CurriculumTopic(
            id="2.1",
            title="Сборка собственного проекта: простая текстовая игра или чат-бот",
            total_hours=4.0,
            lecture_hours=1.0,
            practice_hours=3.0,
            learning_outcome_ids=["LO_H1"],
            rpd_activities=[
                RPDActivity("ЛК", 1.0, "Как спроектировать сценарий игры или помощника."),
                RPDActivity("ПР", 3.0, "Разработка и тестирование проекта школьника."),
            ],
        ),
    ]

    modules = [
        CurriculumModule(
            id="1",
            title="Основы синтаксиса Python",
            total_hours=10.0,
            lecture_hours=4.0,
            practice_hours=6.0,
            topics=topics[:2],
        ),
        CurriculumModule(
            id="2",
            title="Проектная мастерская",
            total_hours=4.0,
            lecture_hours=1.0,
            practice_hours=3.0,
            topics=topics[2:],
        ),
    ]

    questions = [
        TestQuestion(
            id=f"Q{i}",
            question_text=f"Вопрос № {i} по основам Python для 8–9 класса",
            options=["Вариант А", "Вариант Б", "Вариант В", "Вариант Г"],
            correct_option_index=0,
            explanation="Пояснение к верному ответу",
            learning_outcome_ids=[outcomes[(i - 1) % 3].id],
            topic_id=topics[(i - 1) % 3].id,
        )
        for i in range(1, 11)
    ]

    tasks = [
        SeminarTask(
            id="T1",
            topic_id="1.1",
            title="Практикум: Диалоговая открытка",
            scenario_description="Напишите программу, которая спрашивает имя друга и выводит персональное поздравление.",
            deliverable="Файл greeting.py",
            evaluation_criteria=["Программа запускается без ошибок", "Корректно используется функция input()"],
            learning_outcome_ids=["LO_K1"],
        ),
        SeminarTask(
            id="T2",
            topic_id="1.2",
            title="Практикум: Игра «Угадай число»",
            scenario_description="Реализуйте игру с циклом и подсказками «больше/меньше».",
            deliverable="Файл guess_game.py",
            evaluation_criteria=["Работает условие проверки", "Цикл завершается при угадывании"],
            learning_outcome_ids=["LO_S1", "LO_H1"],
        ),
    ]

    slides = [
        SlideDeckSpec(
            topic_id="1.1",
            topic_title="Знакомство с Python: переменные, ввод-вывод и типы данных",
            duration_hours=2.0,
            lecture_plan_points=["Что такое язык программирования", "Переменные и типы данных"],
            slides=[{"slide_num": 1, "title": "Первые шаги в Python", "bullets": ["Переменная как коробка для данных"]}],
        ),
        SlideDeckSpec(
            topic_id="1.2",
            topic_title="Логика программы: условия if-else и циклы while/for",
            duration_hours=2.0,
            lecture_plan_points=["Условный оператор if-else", "Циклы while и for"],
            slides=[{"slide_num": 1, "title": "Условия и циклы", "bullets": ["Как научить программу принимать решения"]}],
        ),
    ]

    kit = ProgramKit(
        program_title="Первые шаги в Python",
        template_type="doop",
        is_internship=False,
        is_self_study_only=False,
        study_form="ochnaya",
        issued_document="Сертификат об освоении программы",
        total_hours=16.0,
        target_audience="Школьники 8 и 9 классов без опыта программирования",
        entry_level="zero",
        explanatory_note="Дополнительная общеразвивающая программа технической направленности для школьников 8–9 классов.",
        relevance_text="Программирование на Python развивает алгоритмическое мышление подростков.",
        goal_text="Формирование базовых навыков алгоритмизации и создания простых программ на языке Python.",
        tasks_list=["Познакомить с базовым синтаксисом Python", "Научить создавать простую игру или чат-бота"],
        regulatory_documents=[
            "Федеральный закон от 29.12.2012 № 273-ФЗ «Об образовании в Российской Федерации»",
            "Приказ Минпросвещения России от 27.07.2022 № 629 «Об утверждении Порядка организации и осуществления образовательной деятельности по дополнительным общеобразовательным программам»",
        ],
        learning_outcomes=outcomes,
        competencies=[
            CompetencyItem(
                id="C1",
                name="Основы алгоритмизации и программирования на Python",
                prof_standard_link=None,
                learning_outcome_ids=["LO_K1", "LO_S1", "LO_H1"],
                topic_ids=["1.1", "1.2", "2.1"],
                assessment_ids=["test:Q1", "task:T1", "task:T2"],
            )
        ],
        calendar_schedule=CalendarScheduleSpec(
            total_weeks=4,
            total_days=8,
            max_hours_per_day=2.0,
            sessions_count=8,
            hours_per_session=2.0,
            description="4 недели, 8 очных занятий по 2 часа (2 раза в неделю).",
        ),
        modules=modules,
        final_attestation=FinalAttestationSpec(
            form="Защита индивидуального мини-проекта (игра или чат-бот)",
            hours=2.0,
            requirements_description="Ученик демонстрирует работающую программу и объясняет логику ее работы.",
            evaluation_criteria=[
                "Программа запускается и выполняет основную задачу (играет или отвечает на сообщения)",
                "Использованы переменные, условия и циклы",
                "Ученик может своими словами объяснить, как работает его код",
            ],
        ),
        declared_totals={"total": 16.0, "lecture": 5.0, "practice": 9.0, "lab": 0.0, "self_study": 0.0, "final_attestation": 2.0},
        test_questions=questions,
        seminar_tasks=tasks,
        lecture_decks=slides,
        literature=[
            LiteratureSource(
                category="main",
                citation="Бриггс Дж. Python для детей. Самоучитель по программированию. — М.: МИФ, 2022.",
                year=2022,
                is_verified=True,
            ),
            LiteratureSource(
                category="internet",
                citation="Официальная документация Python 3. — URL: https://docs.python.org/ru/3/",
                url="https://docs.python.org/ru/3/",
                is_verified=True,
            ),
        ],
        software_and_tools=["Python 3.12", "VS Code / Thonny IDE"],
    )
    return query, kit


def build_valid_query5_internship_256h() -> Tuple[UserQuerySpec, ProgramKit]:
    """Запрос № 5: Стажировка «DevOps-инженер от Linux до CI/CD» (256 ч: 254 ч СРС + 2 ч ИА, ЛК=ПР=ЛБ=0 везде)."""
    query = UserQuerySpec(
        raw_text="Нужна программа повышения квалификации в форме стажировки «DevOps-инженер от Linux до CI/CD» на 256 часов...",
        program_type="dpp_pk",
        is_internship=True,
        is_self_study_only=True,
        total_hours=256.0,
        target_audience="Начинающие системные администраторы и разработчики",
        entry_level="basic",
        study_form="internship",
        final_attestation_form="Защита отчета о стажировке",
        max_hours_per_day=4.0,
        required_topics=["Linux", "сети", "контейнеризац", "CI/CD", "мониторинг", "инфраструктур"],
    )

    outcomes = [
        LearningOutcome("LO_K1", "knowledge", "Знание архитектуры ОС Linux, сетевых протоколов и принципов контейнеризации"),
        LearningOutcome("LO_S1", "skill", "Умение настраивать конвейеры CI/CD, мониторинг и описывать инфраструктуру как код (IaC)"),
        LearningOutcome("LO_P1", "practical_experience", "Практический опыт автоматизации развертывания и сопровождения сервисов в условиях стажировки"),
    ]

    topics = [
        CurriculumTopic("1.1", "Администрирование ОС Linux и сетевые сервисы", 44.0, self_study_hours=44.0, learning_outcome_ids=["LO_K1"]),
        CurriculumTopic("1.2", "Контейнеризация приложений: Docker и оркестрация", 50.0, self_study_hours=50.0, learning_outcome_ids=["LO_K1", "LO_S1"]),
        CurriculumTopic("1.3", "Построение пайплайнов CI/CD (GitLab CI / GitVerse)", 56.0, self_study_hours=56.0, learning_outcome_ids=["LO_S1", "LO_P1"]),
        CurriculumTopic("1.4", "Управление инфраструктурой как кодом (Terraform, Ansible)", 54.0, self_study_hours=54.0, learning_outcome_ids=["LO_S1", "LO_P1"]),
        CurriculumTopic("1.5", "Мониторинг и логирование (Prometheus, Grafana, ELK)", 50.0, self_study_hours=50.0, learning_outcome_ids=["LO_S1", "LO_P1"]),
    ]

    module = CurriculumModule(
        id="1",
        title="DevOps-инженер: от администрирования Linux до промышленного CI/CD",
        total_hours=254.0,
        lecture_hours=0.0,
        practice_hours=0.0,
        lab_hours=0.0,
        self_study_hours=254.0,
        intermediate_attestation_form="Собеседование с руководителем стажировки",
        topics=topics,
    )

    interviews = [
        InterviewQuestion(
            id=f"IQ{idx}",
            module_id="1",
            topic_id=t.id,
            question_text=f"Контрольный вопрос собеседования по теме {t.title}",
            expected_answer_points=["Архитектурное обоснование", "Демонстрация конфигурационных файлов"],
            learning_outcome_ids=t.learning_outcome_ids,
        )
        for idx, t in enumerate(topics, 1)
    ] + [
        InterviewQuestion(
            id="IQ6",
            module_id="1",
            topic_id="1.1",
            question_text="Диагностика сетевой связности и прав доступа в Linux",
            expected_answer_points=["Использование ss/netstat, iptables/nftables, systemd"],
            learning_outcome_ids=["LO_K1"],
        )
    ]

    kit = ProgramKit(
        program_title="DevOps-инженер от Linux до CI/CD",
        template_type="dpp_pk",
        is_internship=True,
        is_self_study_only=True,
        study_form="internship",
        issued_document="Удостоверение о повышении квалификации установленного образца",
        total_hours=256.0,
        target_audience="Лица со средним профессиональным или высшим образованием",
        entry_level="basic",
        explanatory_note="Стажировка проводится в ИТ-компаниях региона под руководством наставника.",
        relevance_text="Востребованность инженеров DevOps обусловлена автоматизацией цикла поставки ПО.",
        goal_text="Совершенствование профессиональных компетенций в области автоматизации развертывания и сопровождения ИТ-инфраструктуры.",
        tasks_list=["Освоить администрирование Linux и сетей", "Внедрить CI/CD, IaC и мониторинг"],
        regulatory_documents=[
            "Федеральный закон от 29.12.2012 № 273-ФЗ «Об образовании в Российской Федерации»",
            "Приказ Минобрнауки России от 01.07.2013 № 499 «Об утверждении Порядка организации и осуществления образовательной деятельности по дополнительным профессиональным программам»",
            "Приказ Минтруда России от 30.08.2021 № 606н «Об утверждении профессионального стандарта «Системный администратор информационно-коммуникационных систем»»",
        ],
        learning_outcomes=outcomes,
        competencies=[
            CompetencyItem(
                id="PC1",
                name="Администрирование систем управления конфигурациями, контейнеризации и автоматизации развертывания",
                prof_standard_link=ProfStandardLink(
                    standard_code="06.026",
                    standard_title="Системный администратор информационно-коммуникационных систем",
                    mintrud_order="Приказ Минтруда России от 30.08.2021 № 606н",
                    labor_function_code="B/02.5",
                    labor_function_title="Администрирование систем управления конфигурациями, контейнеризации и автоматизации развертывания",
                ),
                learning_outcome_ids=["LO_K1", "LO_S1", "LO_P1"],
                topic_ids=[t.id for t in topics],
            )
        ],
        calendar_schedule=CalendarScheduleSpec(
            total_weeks=9,
            total_days=64,
            max_hours_per_day=4.0,
            description="64 учебных дня (около 2 месяцев), не более 4 часов самостоятельной работы в день.",
        ),
        modules=[module],
        final_attestation=FinalAttestationSpec(
            form="Защита отчёта о стажировке (зачёт)",
            hours=2.0,
            requirements_description="Отчёт о выполненных инфраструктурных задачах в компании, репозиторий с конфигурациями IaC и CI/CD.",
            evaluation_criteria=[
                "Полнота и работоспособность конфигураций (Docker, CI/CD, Terraform/Ansible, мониторинг) — до 40%",
                "Обоснованность архитектурных решений и соблюдение практик безопасности — до 30%",
                "Качество оформления отчёта о стажировке и ответы на вопросы комиссии — до 30%",
            ],
        ),
        declared_totals={"total": 256.0, "lecture": 0.0, "practice": 0.0, "lab": 0.0, "self_study": 254.0, "final_attestation": 2.0},
        interview_questions=interviews,
        literature=[
            LiteratureSource(
                category="main",
                citation="Моррис К. Инфраструктура как код. — СПб.: Питер, 2022.",
                year=2022,
                is_verified=True,
            ),
            LiteratureSource(
                category="internet",
                citation="Официальная документация Docker. — URL: https://docs.docker.com/",
                url="https://docs.docker.com/",
                is_verified=True,
            ),
        ],
        software_and_tools=["Linux (Ubuntu / Astra Linux)", "Docker", "GitLab CI / GitVerse", "Ansible", "Prometheus & Grafana"],
    )
    return query, kit


def ensure_sample_248_docx() -> str:
    """
    Находит реальный файл образца стажировки 248 ч (35f2d459...docx) в папке проекта,
    а если его нет — создает полноценный тестовый .docx с ошибками образца для автономной проверки.
    """
    candidate_paths = [
        "35f2d459-054d-487a-af30-62fbc241f7d5.docx",
        os.path.join(os.path.dirname(__file__) or ".", "35f2d459-054d-487a-af30-62fbc241f7d5.docx"),
        
    ]
    found = next((p for p in candidate_paths if os.path.exists(p)), None)
    if found is not None:
        return found

    sample_248_path = os.path.join(os.path.dirname(__file__) or ".", "_temp_sample_248.docx")
    tables_spec = [
        [
            ["№ п/п", "Наименование разделов программы", "Трудоёмкость, всего (час.)", "ЛК", "ПР", "ЛБ", "СРС (час.)", "Форма аттестации"],
            ["1.", "Бизнес и системный анализ", "246", "0", "0", "0", "246", "Собеседование"],
            ["Итоговая аттестация", "Итоговая аттестация", "2", "", "", "", "", "зачет"],
            ["Итого (час.)", "Итого (час.)", "248", "0", "66", "6", "174", ""],
        ],
        [
            ["№ п/п", "Наименование разделов, тем программы", "Трудоёмкость, всего (час.)", "ЛК", "ПР", "ЛБ", "СРС (час.)", "Форма контроля"],
            ["1", "Бизнес- и системный анализ", "246", "0", "0", "0", "246", ""],
            ["1.11.", "Текущий контроль", "6", "0", "0", "0", "0", "Собеседование"],
            ["1.2.", "Работа с требованиями", "24", "0", "9", "0", "15", "Собеседование"],
            ["Итоговая аттестация", "Итоговая аттестация", "2", "", "", "", "", "зачет"],
            ["Итого (час.)", "Итого (час.)", "248", "0", "66", "6", "174", ""],
        ],
        [
            ["Наименование компонентов программы", "Содержание учебного материала и формы организации деятельности слушателей", "Всего (час.)"],
            ["Тема 1.6.", "ПР (9 часов): Создание wireframes в Figma", "12"],
            ["Тема 1.7.", "ПР (6 часов): Выделение сущностей и построение модели", "9"],
            ["Итого", "", "248"],
        ],
    ]
    try:
        import docx as _docx
        d = _docx.Document()
        for tbl in tables_spec:
            t = d.add_table(rows=len(tbl), cols=len(tbl[0]))
            for r_i, row in enumerate(tbl):
                for c_i, val in enumerate(row):
                    t.rows[r_i].cells[c_i].text = val
        d.save(sample_248_path)
    except ImportError:
        import zipfile
        xml_parts = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>']
        for tbl in tables_spec:
            xml_parts.append("<w:tbl>")
            for row in tbl:
                xml_parts.append("<w:tr>")
                for cell in row:
                    xml_parts.append(f"<w:tc><w:p><w:r><w:t>{cell}</w:t></w:r></w:p></w:tc>")
                xml_parts.append("</w:tr>")
            xml_parts.append("</w:tbl>")
        xml_parts.append("</w:body></w:document>")
        with zipfile.ZipFile(sample_248_path, "w") as zf:
            zf.writestr("word/document.xml", "".join(xml_parts))
    return sample_248_path


class TestProgramValidator(unittest.TestCase):
    def setUp(self):
        self.validator = ProgramValidator()

    def test_query1_doop_valid_passes_all_checks(self):
        """Эталонный комплект для Запроса № 1 (16 ч ДООП Python для школьников) проходит все проверки без ошибок."""
        query, kit = build_valid_query1_doop_16h()
        report = self.validator.validate(kit, query)
        self.assertTrue(report.is_valid, msg=report.to_markdown())
        self.assertEqual(report.error_count, 0)
        self.assertTrue(report.hours_check_passed)
        self.assertTrue(report.coverage_check_passed)
        self.assertTrue(report.normative_check_passed)

    def test_query1_catches_pedagogical_order_and_profstandard_violations(self):
        """Проверяет, что валидатор ловит нарушение порядка тем (классы раньше переменных) и лишний профстандарт в ДООП."""
        query, kit = build_valid_query1_doop_16h()
        kit.modules[0].topics[0].title = "Классы, ООП и внешние API в Python"
        kit.competencies[0].prof_standard_link = ProfStandardLink(
            standard_code="06.001",
            standard_title="Программист",
            mintrud_order="Приказ № 424н",
            labor_function_code="A/01.3",
            labor_function_title="Разработка кода",
        )
        report = self.validator.validate(kit, query)
        self.assertFalse(report.is_valid)
        codes = {i.code for i in report.issues}
        self.assertIn("PEDAGOGICAL_ORDER_VIOLATION", codes)
        self.assertIn("DOOP_FORBIDDEN_PROFSTANDARD_LINK", codes)

    def test_query2_restricted_foreign_services_detected(self):
        """Проверяет, что для Запроса № 2 (учителя Ингушетии, сервисы только в РФ) блокируются ChatGPT, Midjourney и Kahoot."""
        _, kit = build_valid_query1_doop_16h()
        kit.template_type = "dpp_pk"
        kit.issued_document = "Удостоверение о повышении квалификации"
        kit.regulatory_documents.append("Приказ Минобрнауки России от 01.07.2013 № 499")
        kit.competencies[0].prof_standard_link = ProfStandardLink(
            standard_code="01.001",
            standard_title="Педагог",
            mintrud_order="Приказ Минтруда № 544н",
            labor_function_code="A/01.6",
            labor_function_title="Общепедагогическая функция. Обучение",
        )
        kit.software_and_tools = ["ChatGPT", "Midjourney", "Kahoot"]

        query2 = UserQuerySpec(
            raw_text="Программа повышения квалификации «Искусственный интеллект в работе учителя»... Использовать только сервисы, которые доступны в России.",
            program_type="dpp_pk",
            total_hours=16.0,
            russia_available_services_only=True,
        )
        report = self.validator.validate(kit, query2)
        self.assertFalse(report.is_valid)
        codes = {i.code for i in report.issues}
        self.assertIn("RESTRICTED_FOREIGN_SERVICE_USED", codes)

    def test_query3_deprecated_2026_tech_and_verbatim_copying_detected(self):
        """Проверяет отлов устаревших библиотек на 2026 год (Python 3.7, df.append, load_boston) и дословного копирования."""
        query, kit = build_valid_query1_doop_16h()
        kit.software_and_tools = ["Python 3.7", "pandas DataFrame.append()", "sklearn load_boston"]
        ref_paragraph = (
            "Программа дает слушателям четкое представление о современных подходах к разработке продуктов "
            "и формирует навык выбора оптимальных архитектурных решений и работы с открытыми интерфейсами."
        )
        kit.explanatory_note = ref_paragraph * 5
        query.reference_texts = [ref_paragraph]

        report = self.validator.validate(kit, query)
        self.assertFalse(report.is_valid)
        codes = {i.code for i in report.issues}
        self.assertIn("DEPRECATED_TOOL_OR_LIBRARY_2026", codes)
        self.assertIn("EXCESSIVE_VERBATIM_COPYING", codes)

    def test_query4_incomplete_request_generates_at_most_5_questions_and_gap_note(self):
        """Проверяет Запрос № 4 (неполный запрос по кибербезопасности): ровно <= 5 вопросов и аналитическая заметка о реалистичности."""
        q4 = UserQuerySpec(
            raw_text="Сделайте курс по кибербезопасности для студентов колледжа, примерно на месяц, чтобы после него ребята могли пойти работать в эту сферу.",
            promises_job_placement=True,
        )
        analysis = analyze_methodologist_query(q4)
        self.assertTrue(analysis.needs_clarification)
        self.assertLessEqual(len(analysis.questions_for_methodologist), 5)
        self.assertGreaterEqual(len(analysis.questions_for_methodologist), 4)
        self.assertIsNotNone(analysis.expectation_reality_gap_note)
        self.assertIn("начальный (базовый)", analysis.expectation_reality_gap_note)

    def test_query5_internship_valid_and_detects_reference_docx_bug(self):
        """
        Проверяет Запрос № 5 (Стажировка DevOps 256 ч) и инспекцию реального файла образца АЦР на 248 ч,
        в котором допущена ошибка в итоговой строке (66 ПР, 6 ЛБ, 174 СРС при 246 СРС в разделе).
        """
        query5, kit5 = build_valid_query5_internship_256h()
        sample_248_path = ensure_sample_248_docx()

        # 1. Инспекция самого образца должна найти ошибку итоговой строки и расхождения УТП/РПД
        ref_issues = inspect_reference_docx(sample_248_path)
        ref_codes = {i.code for i in ref_issues}
        self.assertIn("REF_TOTAL_ROW_MISMATCH", ref_codes)
        self.assertIn("REF_UTP_TOPICS_VS_MODULE_MISMATCH", ref_codes)
        self.assertIn("REF_RPD_TEXT_VS_COLUMN_MISMATCH", ref_codes)

        # 2. Наш сгенерированный комплект на 256 ч проходит валидацию без единой ошибки ERROR,
        #    но сохраняет в отчете предупреждения (WARNING) об ошибках в приложенном образце для методиста!
        report = self.validator.validate(kit5, query5, reference_docx_paths=[sample_248_path])
        self.assertTrue(report.is_valid, msg=report.to_markdown())
        self.assertEqual(report.error_count, 0)
        self.assertGreater(report.warning_count, 0)

        # 3. Если бы агент перенес ошибку образца в строку «Итого» (например, 66 ПР и 6 ЛБ), валидатор выдает жесткий ERROR!
        kit5.declared_totals["practice"] = 66.0
        kit5.declared_totals["lab"] = 6.0
        kit5.declared_totals["self_study"] = 182.0
        bad_report = self.validator.validate(kit5, query5)
        self.assertFalse(bad_report.is_valid)
        bad_codes = {i.code for i in bad_report.issues}
        self.assertIn("DECLARED_TOTAL_ROW_MISMATCH", bad_codes)
        self.assertIn("SELF_STUDY_NONZERO_CONTACT_HOURS", bad_codes)

        # 4. Проверяем, что auto_fix_hours() автоматически чинит сломанную итоговую строку
        fixed_kit = auto_fix_hours(kit5)
        fixed_report = self.validator.validate(fixed_kit, query5)
        self.assertTrue(fixed_report.is_valid)
        self.assertEqual(fixed_kit.declared_totals["practice"], 0.0)
        self.assertEqual(fixed_kit.declared_totals["self_study"], 254.0)


def print_demo_summary() -> None:
    """Печатает наглядную сводку работы валидатора в консоль при запуске из Visual Studio / VS Code."""
    validator = ProgramValidator()
    print("=" * 80)
    print("ДЕМОНСТРАЦИЯ РАБОТЫ ДЕТЕРМИНИРОВАННОГО ВАЛИДАТОРА АЦР")
    print("=" * 80)

    # 1. Запрос № 1 (16 ч ДООП)
    q1, kit1 = build_valid_query1_doop_16h()
    rep1 = validator.validate(kit1, q1)
    print("\n[Запрос № 1 · Python для школьников, 16 ч ДООП]")
    print(rep1.to_markdown())

    # 2. Запрос № 4 (Неполный запрос по кибербезопасности)
    q4 = UserQuerySpec(
        raw_text="Сделайте курс по кибербезопасности для студентов колледжа, примерно на месяц, чтобы после него ребята могли пойти работать в эту сферу.",
        promises_job_placement=True,
    )
    clar4 = analyze_methodologist_query(q4)
    print("\n" + "-" * 80)
    print("[Запрос № 4 · Неполный запрос по кибербезопасности -> Вопросы методисту]")
    for idx, q_text in enumerate(clar4.questions_for_methodologist, 1):
        print(f"  {idx}. {q_text}")
    if clar4.expectation_reality_gap_note:
        print(f"  * {clar4.expectation_reality_gap_note}")

    # 3. Запрос № 5 + проверка дефектного образца стажировки 248 ч
    q5, kit5 = build_valid_query5_internship_256h()
    ref_path = ensure_sample_248_docx()
    rep5 = validator.validate(kit5, q5, reference_docx_paths=[ref_path])
    print("\n" + "-" * 80)
    print("[Запрос № 5 · Стажировка DevOps 256 ч + Инспекция образца 248 ч]")
    print(rep5.to_markdown())
    print("=" * 80 + "\n")


if __name__ == "__main__":
    print_demo_summary()
    unittest.main(verbosity=2)