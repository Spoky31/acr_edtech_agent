# -*- coding: utf-8 -*-
"""
Ядро валидации и контроля целостности образовательных программ АЦР (GigaCowork / Cloud.ru).
Модуль реализует жесткие детерминированные проверки для обоих шаблонов Академии цифрового развития:
1. ДООП (Дополнительная общеобразовательная общеразвивающая программа) — Приказ Минпросвещения № 629.
2. ДПП ПК (Дополнительная профессиональная программа повышения квалификации, включая стажировку) — 273-ФЗ, Приказ Минобрнауки № 499.
Ключевые возможности:
- Предварительный анализ запроса методиста (выявление недостающих параметров, генерация до 5 понятных уточняющих вопросов, проверка реалистичности обещаний относительно объема часов).
- Инспекция приложенных референсных документов .docx (автоматический поиск арифметических ошибок в образцах УП, УТП и РПД, включая коллизию в образце стажировки на 248 ч).
- Многоуровневая проверка арифметики часов (тема -> модуль -> УП -> УТП -> РПД -> календарный график, контроль доли практики, контроль нулевых аудиторных часов для стажировок).
- Проверка матрицы трассировки и карты компетенций (связь каждого результата обучения минимум с 1 темой и минимум с 2 оценочными средствами, проверка ключей тестов и критериев семинаров, проверка дидактического порядка тем для начинающих).
- Проверка нормативной базы, реестра профстандартов Минтруда РФ, региональных ограничений по ИИ-сервисам в РФ, актуальности инструментов на 2026 год, антиплагиата по референсам и пустых полей реквизитов.
- Детерминированная автокоррекция сумм часов (auto_fix_hours).
"""
from __future__ import annotations
import dataclasses
from dataclasses import dataclass, field
import json
import math
import re
from typing import Any, Dict, List, Literal, Optional, Set, Tuple
try:
    import docx  # python-docx для инспекции референсных .docx
except ImportError:
    docx = None

# =====================================================================
# 1. СПРАВОЧНИКИ И НОРМАТИВНЫЕ БАЗЫ (2026 г.)

# =====================================================================
TemplateType = Literal["doop", "dpp_pk"]
EntryLevel = Literal["zero", "basic", "intermediate", "advanced"]
StudyForm = Literal["ochnaya", "ochno_zaochnaya", "zaochnaya_dot", "internship"]

@dataclass(frozen=True)

class ProfStandardInfo:
    """Эталонная запись из реестра профессиональных стандартов Минтруда РФ."""
    code: str
    title: str
    mintrud_order: str
    labor_functions: Dict[str, str]
# Верифицированный реестр профильных профстандартов для ИТ и образования
VERIFIED_PROF_STANDARDS: Dict[str, ProfStandardInfo] = {
    "01.001": ProfStandardInfo(
        code="01.001",
        title="Педагог (педагогическая деятельность в сфере дошкольного, начального общего, основного общего, среднего общего образования) (воспитатель, учитель)",
        mintrud_order="Приказ Минтруда России от 18.10.2013 № 544н",
        labor_functions={
            "A/01.6": "Общепедагогическая функция. Обучение",
            "A/02.6": "Воспитательная деятельность",
            "A/03.6": "Развивающая деятельность",
            "B/02.6": "Педагогическая деятельность по реализации программ основного и среднего общего образования",
        },
    ),
    "06.015": ProfStandardInfo(
        code="06.015",
        title="Специалист по информационным системам",
        mintrud_order="Приказ Минтруда России от 13.07.2023 № 586н",
        labor_functions={
            "C/14.6": "Разработка архитектуры ИС в рамках выполнения работ и управления работами по созданию (модификации) и сопровождению ИС",
            "B/01.5": "Техническая поддержка процессов создания (модификации) и сопровождения ИС",
            "C/01.6": "Выявление и анализ требований к ИС",
        },
    ),
    "06.001": ProfStandardInfo(
        code="06.001",
        title="Программист",
        mintrud_order="Приказ Минтруда России от 20.07.2022 № 424н",
        labor_functions={
            "A/01.3": "Разработка и отладка программного кода",
            "A/02.3": "Проверка работоспособности и рефакторинг кода программного обеспечения",
            "B/01.4": "Интеграция программных модулей и компонентов и верификация выпусков программного продукта",
            "D/02.6": "Разработка требований и проектирование программного обеспечения",
        },
    ),
    "06.022": ProfStandardInfo(
        code="06.022",
        title="Системный аналитик",
        mintrud_order="Приказ Минтруда России от 27.04.2023 № 367н",
        labor_functions={
            "A/01.5": "Сбор, обработка и документирование требований к системе",
            "B/01.6": "Анализ, обобщение и систематизация требований к системе и подсистемам",
            "B/02.6": "Разработка концепции и проектирование логической архитектуры системы",
        },
    ),
    "06.026": ProfStandardInfo(
        code="06.026",
        title="Системный администратор информационно-коммуникационных систем",
        mintrud_order="Приказ Минтруда России от 30.08.2021 № 606н",
        labor_functions={
            "A/01.5": "Установка, настройка и обслуживание системного и прикладного программного обеспечения",
            "B/02.5": "Администрирование систем управления конфигурациями, контейнеризации и автоматизации развертывания",
            "C/01.6": "Обеспечение надежности, мониторинга и отказоустойчивости инфраструктуры ИКС",
        },
    ),
    "06.033": ProfStandardInfo(
        code="06.033",
        title="Специалист по защите информации в автоматизированных системах",
        mintrud_order="Приказ Минтруда России от 14.09.2022 № 525н",
        labor_functions={
            "A/01.5": "Контроль защищенности информации в автоматизированных системах",
            "A/02.5": "Эксплуатация средств защиты информации в автоматизированных системах",
            "B/01.6": "Анализ уязвимостей и реагирование на инциденты информационной безопасности",
        },
    ),
    "06.035": ProfStandardInfo(
        code="06.035",
        title="Специалист по большим данным",
        mintrud_order="Приказ Минтруда России от 06.07.2020 № 405н",
        labor_functions={
            "A/01.6": "Сбор, первичная обработка и подготовка данных для анализа",
            "A/02.6": "Проведение исследовательского анализа данных и визуализация результатов",
            "B/01.6": "Разработка и проверка статистических и аналитических моделей",
        },
    ),
    "06.042": ProfStandardInfo(
        code="06.042",
        title="Специалист по искусственному интеллекту",
        mintrud_order="Приказ Минтруда России от 04.03.2024 № 93н",
        labor_functions={
            "A/01.6": "Подготовка данных и применение моделей искусственного интеллекта для решения прикладных задач",
            "B/01.6": "Интеграция и валидация систем искусственного интеллекта в прикладных процессах",
        },
    ),
}
# Сервисы, недоступные в РФ напрямую (блокируются или закрыли доступ для РФ)
RESTRICTED_FOREIGN_SERVICES_IN_RF: Dict[str, str] = {
    "chatgpt": "GigaChat / YandexGPT",
    "openai": "GigaChat API / Cloud.ru Evolution Foundation Models",
    "midjourney": "Kandinsky (Сбер) / Шедеврум (Яндекс)",
    "dall-e": "Kandinsky (Сбер) / Шедеврум (Яндекс)",
    "claude": "GigaChat / Qwen / DeepSeek",
    "anthropic": "Cloud.ru Evolution Foundation Models",
    "canva": "Supa / Flyvi / Яндекс Презентации",
    "kahoot": "MyQuiz / Российская ЭИОС / Сферум",
    "mentimeter": "Яндекс Формы / MyQuiz",
    "coursera": "Stepik / Открытое образование",
}
# Устаревшие технологии и библиотеки (неактуальные на 2026 год)
DEPRECATED_TECH_PATTERNS_2026: List[Tuple[str, str]] = [
    (r"\bpython\s*2(\.\d+)?\b", "Python 2 снят с поддержки; используйте Python 3.11+ / 3.12+"),
    (r"\bpython\s*3\.[0-8]\b", "Версии Python 3.0–3.8 устарели (EOL); используйте Python 3.10+"),
    (r"\b\.append\s*\(", "В pandas 2.0+ метод DataFrame.append() удален; используйте pd.concat()"),
    (r"\bboston\s+housing\b|\bload_boston\b", "Датасет load_boston удален из scikit-learn; используйте fetch_california_housing или банковские датасеты"),
    (r"\bsklearn\.externals\.joblib\b", "sklearn.externals.joblib удален; используйте прямой импорт import joblib"),
    (r"\btensorflow\s*1(\.\d+)?\b", "TensorFlow 1.x устарел; используйте PyTorch 2.x или TensorFlow 2.15+"),
    (r"\bdisutils\b|\bdistutils\b", "Модуль distutils удален в Python 3.12+; используйте setuptools / pyproject.toml"),
]

# =====================================================================
# 2. СТРУКТУРЫ ДАННЫХ КОМПЛЕКТА ПРОГРАММЫ (JSON SCHEMA / DATACLASSES)

# =====================================================================

@dataclass

class UserQuerySpec:
    """Входной запрос методиста и распознанные ограничения."""
    raw_text: str
    program_type: Optional[TemplateType] = None  # "doop" или "dpp_pk"
    is_internship: bool = False
    is_self_study_only: bool = False
    total_hours: Optional[float] = None
    topic_direction: Optional[str] = None
    target_audience: Optional[str] = None
    entry_level: Optional[EntryLevel] = None
    study_form: Optional[StudyForm] = None
    final_attestation_form: Optional[str] = None
    min_practice_share: Optional[float] = None  # Например, 0.5 для Запроса 1
    max_hours_per_day: Optional[float] = None   # Например, 4.0 для Запроса 5
    planned_weeks: Optional[int] = None
    sessions_per_week: Optional[int] = None
    hours_per_session: Optional[float] = None
    required_topics: List[str] = field(default_factory=list)
    slides_target_scope: Optional[str] = None   # Например, "занятия 1 и 2" или "модуль про pandas"
    required_test_questions_count: Optional[int] = None
    russia_available_services_only: bool = False
    promises_job_placement: bool = False
    reference_texts: List[str] = field(default_factory=list)

@dataclass

class ProfStandardLink:
    """Привязка к профессиональному стандарту (только для ДПП ПК)."""
    standard_code: str
    standard_title: str
    mintrud_order: str
    labor_function_code: str
    labor_function_title: str

@dataclass

class LearningOutcome:
    """Планируемый результат обучения (Знание / Умение / Навык / Практический опыт)."""
    id: str
    category: Literal["knowledge", "skill", "habit", "practical_experience"]
    description: str
    competency_code: Optional[str] = None

@dataclass

class CompetencyItem:
    """Компетенция в карте компетенций со связями с ЗУН, темами и ФОС."""
    id: str
    name: str
    prof_standard_link: Optional[ProfStandardLink] = None
    learning_outcome_ids: List[str] = field(default_factory=list)
    topic_ids: List[str] = field(default_factory=list)
    assessment_ids: List[str] = field(default_factory=list)

@dataclass

class RPDActivity:
    """Элемент рабочей программы дисциплин (РПД) внутри темы."""
    activity_type: Literal["ЛК", "ПР", "ЛБ", "СРС", "ТК"]
    hours: float
    title_and_content: str

@dataclass

class CurriculumTopic:
    """Тема в учебно-тематическом плане (УТП) и рабочей программе (РПД)."""
    id: str                     # Например, "1.1"
    title: str
    total_hours: float
    lecture_hours: float = 0.0
    practice_hours: float = 0.0
    lab_hours: float = 0.0
    self_study_hours: float = 0.0
    control_hours: float = 0.0
    control_form: str = ""
    content_summary: str = ""
    rpd_activities: List[RPDActivity] = field(default_factory=list)
    learning_outcome_ids: List[str] = field(default_factory=list)

@dataclass

class CurriculumModule:
    """Модуль / раздел в учебном плане (УП)."""
    id: str                     # Например, "1" или "0" (Вводное занятие)
    title: str
    total_hours: float
    lecture_hours: float = 0.0
    practice_hours: float = 0.0
    lab_hours: float = 0.0
    self_study_hours: float = 0.0
    control_hours: float = 0.0
    intermediate_attestation_form: str = "—"
    topics: List[CurriculumTopic] = field(default_factory=list)

@dataclass

class CalendarScheduleSpec:
    """Параметры календарного учебного графика."""
    total_weeks: int
    total_days: int
    max_hours_per_day: float
    sessions_count: Optional[int] = None
    hours_per_session: Optional[float] = None
    description: str = ""

@dataclass

class TestQuestion:
    """Тестовый вопрос с ключом и привязкой к результату обучения."""
    id: str
    question_text: str
    options: List[str]
    correct_option_index: int
    explanation: str
    learning_outcome_ids: List[str]
    topic_id: str

@dataclass

class SeminarTask:
    """Практическое задание для семинара/практикума с критериями оценки."""
    id: str
    topic_id: str
    title: str
    scenario_description: str
    deliverable: str
    evaluation_criteria: List[str]
    learning_outcome_ids: List[str]

@dataclass

class InterviewQuestion:
    """Вопрос для текущего контроля (например, собеседование с руководителем стажировки)."""
    id: str
    module_id: str
    topic_id: str
    question_text: str
    expected_answer_points: List[str]
    learning_outcome_ids: List[str]

@dataclass

class FinalAttestationSpec:
    """Описание итоговой аттестации и критериев оценивания."""
    form: str                   # Защита проекта / Зачет по портфолио / Экзамен на стенде / Защита отчета о стажировке
    hours: float
    requirements_description: str
    evaluation_criteria: List[str]
    passing_threshold: str = "Зачтено (не менее 70% выполненных критериев)"

@dataclass

class SlideDeckSpec:
    """План лекции и структура слайдов (.pptx) для темы/занятия."""
    topic_id: str
    topic_title: str
    duration_hours: float
    lecture_plan_points: List[str]
    slides: List[Dict[str, Any]]  # [{"slide_num": 1, "title": "...", "bullets": [...], "visual_note": "..."}]

@dataclass

class LiteratureSource:
    """Источник литературы или нормативно-правовой акт."""
    category: Literal["normative", "main", "additional", "internet"]
    citation: str
    url: Optional[str] = None
    year: Optional[int] = None
    is_verified: bool = True
    verification_tag: str = ""  # Например, "[Требует подтверждения методистом]"

@dataclass

class RequisitesPlaceholders:
    """Поля реквизитов, которые обязаны оставаться пустыми для заполнения человеком."""
    order_number: str = "____-ОБР"
    order_date: str = "«___» __________ 2026 г."
    approver_name: str = ""
    approver_signature: str = ""
    staff_table_Order_names: List[str] = field(default_factory=list)  # Должны быть пустыми строками

@dataclass

class ProgramKit:
    """Полный комплект образовательной программы АЦР (единая JSON-модель)."""
    program_title: str
    template_type: TemplateType
    is_internship: bool
    is_self_study_only: bool
    study_form: StudyForm
    issued_document: str
    total_hours: float
    target_audience: str
    entry_level: EntryLevel
    explanatory_note: str
    relevance_text: str
    goal_text: str
    tasks_list: List[str]
    regulatory_documents: List[str]
    learning_outcomes: List[LearningOutcome]
    competencies: List[CompetencyItem]
    calendar_schedule: CalendarScheduleSpec
    modules: List[CurriculumModule]
    final_attestation: FinalAttestationSpec
    declared_totals: Dict[str, float]  # {"total": ..., "lecture": ..., "practice": ..., "lab": ..., "self_study": ..., "final_attestation": ...}
    test_questions: List[TestQuestion] = field(default_factory=list)
    seminar_tasks: List[SeminarTask] = field(default_factory=list)
    interview_questions: List[InterviewQuestion] = field(default_factory=list)
    lecture_decks: List[SlideDeckSpec] = field(default_factory=list)
    literature: List[LiteratureSource] = field(default_factory=list)
    software_and_tools: List[str] = field(default_factory=list)
    methodologist_notes: List[str] = field(default_factory=list)
    requisites: RequisitesPlaceholders = field(default_factory=RequisitesPlaceholders)
    def to_dict(self) -> Dict[str, Any]:
        return dataclasses.asdict(self)
    def to_json(self, indent: int = 2) -> str:
        return json.dumps(self.to_dict(), ensure_ascii=False, indent=indent)

# =====================================================================
# 3. РЕЗУЛЬТАТЫ ВАЛИДАЦИИ И ОТЧЕТ АВТОПРОВЕРКИ

# =====================================================================

@dataclass

class ValidationIssue:
    """Отдельное замечание валидатора с точной локацией и правилом исправления."""
    code: str
    severity: Literal["ERROR", "WARNING", "INFO"]
    category: Literal[
        "HOURS_ARITHMETIC",
        "COMPETENCY_COVERAGE",
        "TEMPLATE_NORMATIVE",
        "TOOLS_AND_SOURCES",
        "PEDAGOGY_AND_AUDIENCE",
        "REFERENCE_SAMPLE_DEFECT",
    ]
    location: str
    message: str
    suggested_fix: str = ""

@dataclass

class ClarificationAnalysis:
    """Результат проверки полноты входного запроса методиста (Шаг 2 сценария)."""
    needs_clarification: bool
    missing_fields: List[str]
    questions_for_methodologist: List[str]  # Строго не более 5 понятных вопросов без жаргона
    expectation_reality_gap_note: Optional[str] = None
    recommended_template: Optional[TemplateType] = None

@dataclass

class ValidationReport:
    """Итоговый отчет самопроверки агента для методиста."""
    is_valid: bool
    error_count: int
    warning_count: int
    hours_check_passed: bool
    coverage_check_passed: bool
    normative_check_passed: bool
    sources_and_tools_passed: bool
    issues: List[ValidationIssue] = field(default_factory=list)
    coverage_summary: Dict[str, Any] = field(default_factory=dict)
    hours_summary: Dict[str, Any] = field(default_factory=dict)
    def to_dict(self) -> Dict[str, Any]:
        return dataclasses.asdict(self)
    def to_markdown(self) -> str:
        status = "ПРОЙДЕНА УСПЕШНО" if self.is_valid else "ТРЕБУЕТ ИСПРАВЛЕНИЙ"
        lines = [
            f"## Отчет автопроверки комплекта программы: {status}",
            f"- Ошибок (ERROR): **{self.error_count}** | Предупреждений (WARNING): **{self.warning_count}**",
            f"- Сходимость часов (УП ↔ УТП ↔ РПД ↔ График): **{'ОК' if self.hours_check_passed else 'ОШИБКА'}**",
            f"- Покрытие компетенций и результатов обучения: **{'ОК' if self.coverage_check_passed else 'ОШИБКА'}**",
            f"- Соответствие шаблону АЦР и нормативной базе: **{'ОК' if self.normative_check_passed else 'ОШИБКА'}**",
            f"- Достоверность источников и ограничения по сервисам: **{'ОК' if self.sources_and_tools_passed else 'ОШИБКА'}**",
        ]
        if self.issues:
            lines.append("\n### Детализация проверок")
            for idx, issue in enumerate(self.issues, 1):
                fix_str = f" *(Рекомендация: {issue.suggested_fix})*" if issue.suggested_fix else ""
                lines.append(
                    f"{idx}. **[{issue.severity} · {issue.code}]** `{issue.location}` — {issue.message}{fix_str}"
                )
        return "\n".join(lines)

# =====================================================================
# 4. АНАЛИЗАТОР ВХОДНОГО ЗАПРОСА И ИНСПЕКТОР РЕФЕРЕНСОВ (.DOCX)

# =====================================================================

def analyze_methodologist_query(query: UserQuerySpec) -> ClarificationAnalysis:
    """
    Проверяет запрос методиста до генерации каркаса.
    Если критичных сведений не хватает (как в Запросе № 4), формирует до 5 понятных вопросов.
    Также фиксирует расхождение между завышенными ожиданиями («пойти работать за месяц») и реальным объемом.
    """
    missing: List[str] = []
    questions: List[str] = []
    # 1. Тип программы
    if not query.program_type:
        missing.append("program_type")
        questions.append(
            "Какой тип программы требуется разработать: программу повышения квалификации (с выдачей удостоверения и привязкой к профстандарту) или дополнительную общеразвивающую программу (с выдачей сертификата)?"
        )
    # 2. Точный объем в академических часах
    if query.total_hours is None or query.total_hours <= 0:
        missing.append("total_hours")
        questions.append(
            "Какой точный объем программы в академических часах нужно заложить в учебный план (например, 36, 72 или 144 часа)?"
        )
    # 3. Входной уровень и профиль слушателей
    if not query.entry_level or not query.target_audience:
        missing.append("entry_level_and_audience")
        questions.append(
            "Каков начальный уровень подготовки слушателей (на каком курсе или в какой должности они обучаются, знакомы ли с базовыми инструментами, сетями или программированием)?"
        )
    # 4. Форма обучения
    if not query.study_form:
        missing.append("study_form")
        questions.append(
            "В какой форме планируется проводить обучение: очной, очно-заочной с применением дистанционных технологий или в формате стажировки?"
        )
    # 5. Форма итоговой аттестации
    if not query.final_attestation_form:
        missing.append("final_attestation_form")
        questions.append(
            "В какой форме должна проходить итоговая аттестация (например, экзамен с практическим заданием, защита итогового проекта или зачет по портфолио)?"
        )
    # Ограничиваем список строго максимум 5 вопросами (по требованию чек-листа Запроса № 4)
    questions = questions[:5]
    # Проверка реалистичности обещаний («пойти работать в сферу» за короткий курс <= 108 ч)
    gap_note: Optional[str] = None
    effective_hours = query.total_hours or 72.0
    if query.promises_job_placement or re.search(
        r"пойти\s+работать|трудоустр|стать\s+готовым\s+специалистом|с\s+нуля\s+до\s+(middle|работы)",
        query.raw_text.lower(),
    ):
        if effective_hours <= 144:
            gap_note = (
                f"Аналитическая заметка для методиста: В запросе указана цель «пойти работать в сферу», "
                f"однако объем курса ({int(effective_hours)} ак. ч.) позволяет сформировать только начальный (базовый) "
                f"практический уровень для выполнения отдельных типовых задач стажера/младшего специалиста. "
                f"В программе результаты обучения сформулированы реалистично для {int(effective_hours)} часов без обещания "
                f"полной профессиональной переподготовки."
            )
    return ClarificationAnalysis(
        needs_clarification=len(missing) > 0,
        missing_fields=missing,
        questions_for_methodologist=questions,
        expectation_reality_gap_note=gap_note,
        recommended_template=query.program_type,
    )

def _parse_float_ru(val: str) -> Optional[float]:
    """Парсит число из ячейки таблицы Word (поддерживает запятую и точку)."""
    cleaned = val.strip().replace(",", ".")
    if not cleaned or cleaned in ("—", "-", "–"):
        return 0.0
    match = re.search(r"^\d+(\.\d+)?$", cleaned)
    if match:
        return float(cleaned)
    return None

def _extract_tables_from_docx(docx_path: str) -> List[List[List[str]]]:
    """
    Извлекает таблицы из .docx файла. Работает через python-docx, а если пакет не установлен —
    через встроенные модули Python (zipfile + xml.etree.ElementTree) без внешних зависимостей!
    """
    if docx is not None:
        try:
            d = docx.Document(docx_path)
            result: List[List[List[str]]] = []
            for t in d.tables:
                rows_data: List[List[str]] = []
                for r in t.rows:
                    rows_data.append([c.text.strip().replace("\n", " ") for c in r.cells])
                result.append(rows_data)
            return result
        except Exception:
            pass
    import zipfile
    import xml.etree.ElementTree as ET
    ns = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
    result = []
    with zipfile.ZipFile(docx_path, "r") as zf:
        xml_bytes = zf.read("word/document.xml")
    root = ET.fromstring(xml_bytes)
    for tbl in root.findall(".//w:tbl", ns):
        rows_data = []
        for tr in tbl.findall("./w:tr", ns):
            row_cells: List[str] = []
            for tc in tr.findall("./w:tc", ns):
                texts = [t.text for t in tc.findall(".//w:t", ns) if t.text]
                cell_str = "".join(texts).strip().replace("\n", " ")
                grid_span_el = tc.find("./w:tcPr/w:gridSpan", ns)
                span = 1
                if grid_span_el is not None:
                    val_attr = grid_span_el.attrib.get(f"{{{ns['w']}}}val")
                    if val_attr and val_attr.isdigit():
                        span = int(val_attr)
                for _ in range(span):
                    row_cells.append(cell_str)
            if row_cells:
                rows_data.append(row_cells)
        if rows_data:
            result.append(rows_data)
    return result

def inspect_reference_docx(docx_path: str) -> List[ValidationIssue]:
    """
    Инспектирует приложенный референсный файл .docx (например, образец стажировки 248 ч)
    и находит в нем арифметические и структурные ошибки, чтобы агент предупредил методиста
    и не перенес ошибки образца в новую программу.
    """
    issues: List[ValidationIssue] = []
    tables = _extract_tables_from_docx(docx_path)
    for t_idx, table_rows in enumerate(tables):
        if len(table_rows) < 2:
            continue
        headers = table_rows[0]
        header_joined = " | ".join(headers).lower()
        # Проверка таблиц Учебного плана (УП) и Учебно-тематического плана (УТП)
        if "трудоёмкость" in header_joined or "общая трудоемкость" in header_joined:
            module_rows: List[Tuple[str, float, float, float, float, float]] = []
            topic_rows: List[Tuple[str, float, float, float, float, float]] = []
            ias_hours = 0.0
            total_row_vals: Optional[Tuple[float, float, float, float, float]] = None
            for r_idx in range(1, len(table_rows)):
                cells = table_rows[r_idx]
                if len(cells) < 7:
                    continue
                row_num = cells[0].strip()
                row_title = cells[1].strip()
                if "наименование" in row_title.lower() or "лекции" in cells[4].lower():
                    continue
                total_h = _parse_float_ru(cells[2])
                if total_h is None:
                    continue
                # Определяем формат колонок: ДПП ПК (Всего, ЛК, ПР, ЛБ, СРС) vs ДООП (Всего, Аудиторные, ЛК, ПР, СРС)
                is_doop_layout = "аудиторных" in header_joined
                if is_doop_layout:
                    aud_h = _parse_float_ru(cells[3]) or 0.0
                    lk_h = _parse_float_ru(cells[4]) or 0.0
                    pr_h = _parse_float_ru(cells[5]) or 0.0
                    lb_h = 0.0
                    srs_h = _parse_float_ru(cells[6]) or 0.0
                    if "итого" not in row_title.lower() and "аттестация" not in row_title.lower():
                        if not math.isclose(aud_h, lk_h + pr_h, abs_tol=1e-2):
                            issues.append(
                                ValidationIssue(
                                    code="REF_DOOP_AUD_MISMATCH",
                                    severity="WARNING",
                                    category="REFERENCE_SAMPLE_DEFECT",
                                    location=f"Референс Таблица #{t_idx + 1}, строка '{row_title}'",
                                    message=f"В образце аудиторные часы ({aud_h}) не равны сумме ЛК ({lk_h}) + ПР ({pr_h}).",
                                    suggested_fix="Использовать корректную сумму при генерации новой программы.",
                                )
                            )
                else:
                    lk_h = _parse_float_ru(cells[3]) or 0.0
                    pr_h = _parse_float_ru(cells[4]) or 0.0
                    lb_h = _parse_float_ru(cells[5]) or 0.0
                    srs_h = _parse_float_ru(cells[6]) or 0.0
                title_lower = row_title.lower()
                if "итоговая аттестация" in title_lower or row_num.lower() == "итоговая аттестация":
                    ias_hours += total_h
                elif title_lower.startswith("итого") or row_num.lower().startswith("итого"):
                    total_row_vals = (total_h, lk_h, pr_h, lb_h, srs_h)
                else:
                    # Проверяем горизонтальную сумму строки
                    row_sum = lk_h + pr_h + lb_h + srs_h
                    if not math.isclose(total_h, row_sum, abs_tol=1e-2):
                        issues.append(
                            ValidationIssue(
                                code="REF_ROW_HORIZONTAL_MISMATCH",
                                severity="WARNING",
                                category="REFERENCE_SAMPLE_DEFECT",
                                location=f"Референс Таблица #{t_idx + 1}, строка '{row_num} {row_title}'",
                                message=(
                                    f"В образце трудоёмкость строки ({total_h} ч.) не сходится с суммой "
                                    f"видов работ: ЛК({lk_h}) + ПР({pr_h}) + ЛБ({lb_h}) + СРС({srs_h}) = {row_sum} ч."
                                ),
                                suggested_fix="Не переносить эту ошибку из образца; выровнять сумму по видам работ.",
                            )
                        )
                    # Отличаем строку модуля (например, '1' или '1.') от подтемы ('1.1.', '1.2.')
                    if re.match(r"^\d+\.\d+", row_num):
                        topic_rows.append((f"{row_num} {row_title}", total_h, lk_h, pr_h, lb_h, srs_h))
                    else:
                        module_rows.append((f"{row_num} {row_title}", total_h, lk_h, pr_h, lb_h, srs_h))
            # Сверяем сумму модулей со строкой «Итого»
            if total_row_vals and module_rows:
                tot_all, tot_lk, tot_pr, tot_lb, tot_srs = total_row_vals
                mod_all = sum(m[1] for m in module_rows) + ias_hours
                mod_lk = sum(m[2] for m in module_rows)
                mod_pr = sum(m[3] for m in module_rows)
                mod_lb = sum(m[4] for m in module_rows)
                mod_srs = sum(m[5] for m in module_rows)
                if not (
                    math.isclose(tot_all, mod_all, abs_tol=1e-2)
                    and math.isclose(tot_lk, mod_lk, abs_tol=1e-2)
                    and math.isclose(tot_pr, mod_pr, abs_tol=1e-2)
                    and math.isclose(tot_lb, mod_lb, abs_tol=1e-2)
                    and math.isclose(tot_srs, mod_srs, abs_tol=1e-2)
                ):
                    issues.append(
                        ValidationIssue(
                            code="REF_TOTAL_ROW_MISMATCH",
                            severity="WARNING",
                            category="REFERENCE_SAMPLE_DEFECT",
                            location=f"Референс Таблица #{t_idx + 1} (Итоговая строка)",
                            message=(
                                f"Обнаружена ошибка в референсном образце! По строкам модулей заявлено: "
                                f"Всего={mod_all:g} ч. (ЛК={mod_lk:g}, ПР={mod_pr:g}, ЛБ={mod_lb:g}, СРС={mod_srs:g}, ИА={ias_hours:g}), "
                                f"но в строке «Итого» указано: Всего={tot_all:g} ч. (ЛК={tot_lk:g}, ПР={tot_pr:g}, ЛБ={tot_lb:g}, СРС={tot_srs:g})."
                            ),
                            suggested_fix="При генерации новой программы рассчитывать строку «Итого» строго как сумму строк разделов без копирования ошибки образца.",
                        )
                    )
            # Сверяем сумму подтем с модулем (если это УТП)
            if topic_rows and module_rows:
                top_all = sum(t[1] for t in topic_rows)
                top_lk = sum(t[2] for t in topic_rows)
                top_pr = sum(t[3] for t in topic_rows)
                top_lb = sum(t[4] for t in topic_rows)
                top_srs = sum(t[5] for t in topic_rows)
                mod_all = sum(m[1] for m in module_rows)
                mod_lk = sum(m[2] for m in module_rows)
                mod_pr = sum(m[3] for m in module_rows)
                mod_lb = sum(m[4] for m in module_rows)
                mod_srs = sum(m[5] for m in module_rows)
                if not (
                    math.isclose(top_all, mod_all, abs_tol=1e-2)
                    and math.isclose(top_lk, mod_lk, abs_tol=1e-2)
                    and math.isclose(top_pr, mod_pr, abs_tol=1e-2)
                    and math.isclose(top_lb, mod_lb, abs_tol=1e-2)
                    and math.isclose(top_srs, mod_srs, abs_tol=1e-2)
                ):
                    issues.append(
                        ValidationIssue(
                            code="REF_UTP_TOPICS_VS_MODULE_MISMATCH",
                            severity="WARNING",
                            category="REFERENCE_SAMPLE_DEFECT",
                            location=f"Референс Таблица #{t_idx + 1} (УТП: темы vs раздел)",
                            message=(
                                f"В образце УТП сумма тем (Всего={top_all:g}, ЛК={top_lk:g}, ПР={top_pr:g}, ЛБ={top_lb:g}, СРС={top_srs:g}) "
                                f"противоречит строке раздела (Всего={mod_all:g}, ЛК={mod_lk:g}, ПР={mod_pr:g}, ЛБ={mod_lb:g}, СРС={mod_srs:g})."
                            ),
                            suggested_fix="Синхронизировать распределение часов между разделом и темами.",
                        )
                    )
        # Проверка таблицы РПД (Рабочая программа учебных дисциплин), где в тексте написано «ПР (9 часов)», а в колонке «12»
        if "содержание учебного материала" in header_joined and len(table_rows[0]) == 3:
            rpd_running_sum = 0.0
            declared_rpd_total: Optional[float] = None
            for r_idx in range(1, len(table_rows)):
                cells = table_rows[r_idx]
                comp_name, content_text, col_hours_str = cells[0], cells[1], cells[2]
                col_hours = _parse_float_ru(col_hours_str)
                if col_hours is None:
                    continue
                if comp_name.lower().startswith("раздел"):
                    continue
                if comp_name.lower().startswith("итого") and "аттестация" not in comp_name.lower():
                    declared_rpd_total = col_hours
                    continue
                rpd_running_sum += col_hours
                # Проверяем совпадение часов в тексте «ПР (9 часов)» и в правом столбце «12»
                match_in_text = re.search(r"(?:СРС|ПР|ЛР|ЛБ|ЛК|Собеседование)\s*\((\d+(?:[.,]\d+)?)\s*час", content_text)
                if match_in_text:
                    text_hours = float(match_in_text.group(1).replace(",", "."))
                    if not math.isclose(text_hours, col_hours, abs_tol=1e-2):
                        issues.append(
                            ValidationIssue(
                                code="REF_RPD_TEXT_VS_COLUMN_MISMATCH",
                                severity="WARNING",
                                category="REFERENCE_SAMPLE_DEFECT",
                                location=f"Референс Таблица #{t_idx + 1} (РПД), строка #{r_idx} '{comp_name}'",
                                message=(
                                    f"В тексте ячейки РПД указано {text_hours:g} ч., а в столбце «Всего (час.)» — {col_hours:g} ч."
                                ),
                                suggested_fix="Генерировать текст РПД из единой модели часов темы.",
                            )
                        )
            if declared_rpd_total is not None and not math.isclose(rpd_running_sum, declared_rpd_total, abs_tol=1e-2):
                issues.append(
                    ValidationIssue(
                        code="REF_RPD_TOTAL_SUM_MISMATCH",
                        severity="WARNING",
                        category="REFERENCE_SAMPLE_DEFECT",
                        location=f"Референс Таблица #{t_idx + 1} (РПД Итого)",
                        message=(
                            f"В образце РПД фактическая сумма строк активностей составляет {rpd_running_sum:g} ч., "
                            f"тогда как в строке «Итого» указано {declared_rpd_total:g} ч."
                        ),
                        suggested_fix="Исключить дублирование строк текущего контроля и пересчитать сумму РПД.",
                    )
                )
    return issues

# =====================================================================
# 5. ГЛАВНЫЙ ДЕТЕРМИНИРОВАННЫЙ ВАЛИДАТОР ПРОГРАММЫ (ProgramValidator)

# =====================================================================

class ProgramValidator:
    """
    Выполняет полную детерминированную проверку сгенерированного комплекта программы (ProgramKit)
    относительно общих правил АЦР и специфических требований запроса методиста (UserQuerySpec).
    """
    def validate(
        self,
        kit: ProgramKit,
        query: Optional[UserQuerySpec] = None,
        reference_docx_paths: Optional[List[str]] = None,
    ) -> ValidationReport:
        issues: List[ValidationIssue] = []
        # 0. Если приложены референсные .docx, проверяем их и добавляем предупреждения о найденных багах образца
        if reference_docx_paths:
            for path in reference_docx_paths:
                issues.extend(inspect_reference_docx(path))
        # 1. Проверка арифметики часов на всех уровнях
        hours_issues, hours_summary = self._validate_hours(kit, query)
        issues.extend(hours_issues)
        # 2. Проверка карты компетенций, результатов обучения, тестов, семинаров и порядка тем
        cov_issues, cov_summary = self._validate_competencies_and_coverage(kit, query)
        issues.extend(cov_issues)
        # 3. Проверка соответствия шаблону АЦР, нормативной базе, профстандартам и пустым полям
        norm_issues = self._validate_template_and_regulations(kit, query)
        issues.extend(norm_issues)
        # 4. Проверка источников, актуальности инструментов 2026 г., сервисов РФ и антиплагиата
        tools_issues = self._validate_sources_and_tools(kit, query)
        issues.extend(tools_issues)
        error_count = sum(1 for i in issues if i.severity == "ERROR")
        warning_count = sum(1 for i in issues if i.severity == "WARNING")
        hours_passed = not any(i.severity == "ERROR" and i.category == "HOURS_ARITHMETIC" for i in issues)
        cov_passed = not any(
            i.severity == "ERROR" and i.category in ("COMPETENCY_COVERAGE", "PEDAGOGY_AND_AUDIENCE")
            for i in issues
        )
        norm_passed = not any(i.severity == "ERROR" and i.category == "TEMPLATE_NORMATIVE" for i in issues)
        tools_passed = not any(i.severity == "ERROR" and i.category == "TOOLS_AND_SOURCES" for i in issues)
        return ValidationReport(
            is_valid=(error_count == 0),
            error_count=error_count,
            warning_count=warning_count,
            hours_check_passed=hours_passed,
            coverage_check_passed=cov_passed,
            normative_check_passed=norm_passed,
            sources_and_tools_passed=tools_passed,
            issues=issues,
            coverage_summary=cov_summary,
            hours_summary=hours_summary,
        )
    # -----------------------------------------------------------------
    # БЛОК 1: АРИФМЕТИКА ЧАСОВ (УП ↔ УТП ↔ РПД ↔ КАЛЕНДАРНЫЙ ГРАФИК)
    # -----------------------------------------------------------------
    def _validate_hours(
        self, kit: ProgramKit, query: Optional[UserQuerySpec]
    ) -> Tuple[List[ValidationIssue], Dict[str, Any]]:
        issues: List[ValidationIssue] = []
        # 1.1. Соответствие запрошенному объему часов
        if query and query.total_hours is not None:
            if not math.isclose(kit.total_hours, query.total_hours, abs_tol=1e-2):
                issues.append(
                    ValidationIssue(
                        code="HOURS_QUERY_MISMATCH",
                        severity="ERROR",
                        category="HOURS_ARITHMETIC",
                        location="ProgramKit.total_hours",
                        message=f"Общий объем программы ({kit.total_hours:g} ч.) не равен запрошенному ({query.total_hours:g} ч.).",
                        suggested_fix=f"Установить общий объем программы ровно {query.total_hours:g} ч.",
                    )
                )
        calc_mod_total = 0.0
        calc_mod_lk = 0.0
        calc_mod_pr = 0.0
        calc_mod_lb = 0.0
        calc_mod_srs = 0.0
        calc_mod_ctrl = 0.0
        for mod in kit.modules:
            # 1.2. Горизонтальная сумма строки модуля
            mod_row_sum = (
                mod.lecture_hours
                + mod.practice_hours
                + mod.lab_hours
                + mod.self_study_hours
                + mod.control_hours
            )
            if not math.isclose(mod.total_hours, mod_row_sum, abs_tol=1e-2):
                issues.append(
                    ValidationIssue(
                        code="MODULE_HORIZONTAL_SUM_ERROR",
                        severity="ERROR",
                        category="HOURS_ARITHMETIC",
                        location=f"Модуль '{mod.id}. {mod.title}'",
                        message=(
                            f"Трудоемкость модуля ({mod.total_hours:g} ч.) не равна сумме видов работ: "
                            f"ЛК({mod.lecture_hours:g}) + ПР({mod.practice_hours:g}) + ЛБ({mod.lab_hours:g}) + "
                            f"СРС({mod.self_study_hours:g}) + ТК({mod.control_hours:g}) = {mod_row_sum:g} ч."
                        ),
                        suggested_fix="Выровнять виды работ модуля так, чтобы их сумма давала ровно total_hours.",
                    )
                )
            # 1.3. Проверка каждой темы внутри модуля и сверка суммы тем с модулем
            if not mod.topics:
                issues.append(
                    ValidationIssue(
                        code="MODULE_EMPTY_TOPICS",
                        severity="ERROR",
                        category="HOURS_ARITHMETIC",
                        location=f"Модуль '{mod.id}. {mod.title}'",
                        message="В модуле отсутствуют темы учебно-тематического плана.",
                        suggested_fix="Декомпозировать модуль минимум на 2 темы.",
                    )
                )
            else:
                t_sum_total = sum(t.total_hours for t in mod.topics)
                t_sum_lk = sum(t.lecture_hours for t in mod.topics)
                t_sum_pr = sum(t.practice_hours for t in mod.topics)
                t_sum_lb = sum(t.lab_hours for t in mod.topics)
                t_sum_srs = sum(t.self_study_hours for t in mod.topics)
                t_sum_ctrl = sum(t.control_hours for t in mod.topics)
                for topic in mod.topics:
                    t_row_sum = (
                        topic.lecture_hours
                        + topic.practice_hours
                        + topic.lab_hours
                        + topic.self_study_hours
                        + topic.control_hours
                    )
                    if not math.isclose(topic.total_hours, t_row_sum, abs_tol=1e-2):
                        issues.append(
                            ValidationIssue(
                                code="TOPIC_HORIZONTAL_SUM_ERROR",
                                severity="ERROR",
                                category="HOURS_ARITHMETIC",
                                location=f"Тема '{topic.id}. {topic.title}'",
                                message=(
                                    f"Объем темы ({topic.total_hours:g} ч.) не равен сумме видов работ: "
                                    f"ЛК({topic.lecture_hours:g}) + ПР({topic.practice_hours:g}) + "
                                    f"ЛБ({topic.lab_hours:g}) + СРС({topic.self_study_hours:g}) + ТК({topic.control_hours:g}) = {t_row_sum:g} ч."
                                ),
                                suggested_fix="Скорректировать часы темы по видам работ.",
                            )
                        )
                    # Проверка активностей РПД внутри темы (если заполнены)
                    if topic.rpd_activities:
                        rpd_lk = sum(a.hours for a in topic.rpd_activities if a.activity_type == "ЛК")
                        rpd_pr = sum(a.hours for a in topic.rpd_activities if a.activity_type == "ПР")
                        rpd_lb = sum(a.hours for a in topic.rpd_activities if a.activity_type == "ЛБ")
                        rpd_srs = sum(a.hours for a in topic.rpd_activities if a.activity_type == "СРС")
                        rpd_tk = sum(a.hours for a in topic.rpd_activities if a.activity_type == "ТК")
                        rpd_tot = rpd_lk + rpd_pr + rpd_lb + rpd_srs + rpd_tk
                        if not (
                            math.isclose(rpd_tot, topic.total_hours, abs_tol=1e-2)
                            and math.isclose(rpd_lk, topic.lecture_hours, abs_tol=1e-2)
                            and math.isclose(rpd_pr, topic.practice_hours, abs_tol=1e-2)
                            and math.isclose(rpd_lb, topic.lab_hours, abs_tol=1e-2)
                            and math.isclose(rpd_srs, topic.self_study_hours, abs_tol=1e-2)
                            and math.isclose(rpd_tk, topic.control_hours, abs_tol=1e-2)
                        ):
                            issues.append(
                                ValidationIssue(
                                    code="TOPIC_RPD_MISMATCH",
                                    severity="ERROR",
                                    category="HOURS_ARITHMETIC",
                                    location=f"РПД Темы '{topic.id}. {topic.title}'",
                                    message=(
                                        f"Часы в РПД темы (Всего={rpd_tot:g}: ЛК={rpd_lk:g}, ПР={rpd_pr:g}, ЛБ={rpd_lb:g}, СРС={rpd_srs:g}, ТК={rpd_tk:g}) "
                                        f"не совпадают с УТП темы (Всего={topic.total_hours:g}: ЛК={topic.lecture_hours:g}, ПР={topic.practice_hours:g}, "
                                        f"ЛБ={topic.lab_hours:g}, СРС={topic.self_study_hours:g}, ТК={topic.control_hours:g})."
                                    ),
                                    suggested_fix="Синхронизировать элементы РПД с часами темы в УТП.",
                                )
                            )
                # Сверка суммы тем с модулем по каждому столбцу
                if not (
                    math.isclose(mod.total_hours, t_sum_total, abs_tol=1e-2)
                    and math.isclose(mod.lecture_hours, t_sum_lk, abs_tol=1e-2)
                    and math.isclose(mod.practice_hours, t_sum_pr, abs_tol=1e-2)
                    and math.isclose(mod.lab_hours, t_sum_lb, abs_tol=1e-2)
                    and math.isclose(mod.self_study_hours, t_sum_srs, abs_tol=1e-2)
                    and math.isclose(mod.control_hours, t_sum_ctrl, abs_tol=1e-2)
                ):
                    issues.append(
                        ValidationIssue(
                            code="MODULE_TOPICS_VERTICAL_SUM_ERROR",
                            severity="ERROR",
                            category="HOURS_ARITHMETIC",
                            location=f"Модуль '{mod.id}. {mod.title}' (УП vs УТП)",
                            message=(
                                f"Сумма тем внутри модуля (Всего={t_sum_total:g}, ЛК={t_sum_lk:g}, ПР={t_sum_pr:g}, "
                                f"ЛБ={t_sum_lb:g}, СРС={t_sum_srs:g}, ТК={t_sum_ctrl:g}) не равна строке модуля "
                                f"(Всего={mod.total_hours:g}, ЛК={mod.lecture_hours:g}, ПР={mod.practice_hours:g}, "
                                f"ЛБ={mod.lab_hours:g}, СРС={mod.self_study_hours:g}, ТК={mod.control_hours:g})."
                            ),
                            suggested_fix="Пересчитать итоги модуля как точную сумму входящих в него тем.",
                        )
                    )
            calc_mod_total += mod.total_hours
            calc_mod_lk += mod.lecture_hours
            calc_mod_pr += mod.practice_hours
            calc_mod_lb += mod.lab_hours
            calc_mod_srs += mod.self_study_hours
            calc_mod_ctrl += mod.control_hours
        # 1.4. Проверка общего итога курса (Модули + Итоговая аттестация == Общий объем и строка Итого)
        ias_h = kit.final_attestation.hours
        expected_course_total = calc_mod_total + ias_h
        if not math.isclose(kit.total_hours, expected_course_total, abs_tol=1e-2):
            issues.append(
                ValidationIssue(
                    code="COURSE_TOTAL_SUM_ERROR",
                    severity="ERROR",
                    category="HOURS_ARITHMETIC",
                    location="Учебный план (Итого по курсу)",
                    message=(
                        f"Сумма модулей ({calc_mod_total:g} ч.) + Итоговая аттестация ({ias_h:g} ч.) = {expected_course_total:g} ч., "
                        f"что не совпадает с заявленным объемом программы ({kit.total_hours:g} ч.)."
                    ),
                    suggested_fix="Скорректировать часы модулей или итоговой аттестации.",
                )
            )
        # Сверка со строкой declared_totals («Итого (час.)» в таблице УП/УТП — защита от ошибки образца 248 ч!)
        decl = kit.declared_totals or {}
        decl_total = decl.get("total", kit.total_hours)
        decl_lk = decl.get("lecture", calc_mod_lk)
        decl_pr = decl.get("practice", calc_mod_pr)
        decl_lb = decl.get("lab", calc_mod_lb)
        decl_srs = decl.get("self_study", calc_mod_srs)
        if not (
            math.isclose(decl_total, expected_course_total, abs_tol=1e-2)
            and math.isclose(decl_lk, calc_mod_lk, abs_tol=1e-2)
            and math.isclose(decl_pr, calc_mod_pr, abs_tol=1e-2)
            and math.isclose(decl_lb, calc_mod_lb, abs_tol=1e-2)
            and math.isclose(decl_srs, calc_mod_srs, abs_tol=1e-2)
        ):
            issues.append(
                ValidationIssue(
                    code="DECLARED_TOTAL_ROW_MISMATCH",
                    severity="ERROR",
                    category="HOURS_ARITHMETIC",
                    location="Строка «Итого (час.)» учебного плана",
                    message=(
                        f"Строка «Итого» (Всего={decl_total:g}, ЛК={decl_lk:g}, ПР={decl_pr:g}, ЛБ={decl_lb:g}, СРС={decl_srs:g}) "
                        f"противоречит сумме разделов (Всего={expected_course_total:g}, ЛК={calc_mod_lk:g}, ПР={calc_mod_pr:g}, "
                        f"ЛБ={calc_mod_lb:g}, СРС={calc_mod_srs:g})."
                    ),
                    suggested_fix="Вызвать auto_fix_hours() для пересчета строки «Итого» строго по сумме разделов.",
                )
            )
        # 1.5. Проверка режима стажировки без аудиторных занятий (Запрос № 5: ЛК=0, ПР=0, ЛБ=0 везде)
        is_self_study_mode = kit.is_self_study_only or (query and query.is_self_study_only)
        if is_self_study_mode:
            if not (
                math.isclose(calc_mod_lk, 0.0, abs_tol=1e-2)
                and math.isclose(calc_mod_pr, 0.0, abs_tol=1e-2)
                and math.isclose(calc_mod_lb, 0.0, abs_tol=1e-2)
                and math.isclose(decl_lk, 0.0, abs_tol=1e-2)
                and math.isclose(decl_pr, 0.0, abs_tol=1e-2)
                and math.isclose(decl_lb, 0.0, abs_tol=1e-2)
            ):
                issues.append(
                    ValidationIssue(
                        code="SELF_STUDY_NONZERO_CONTACT_HOURS",
                        severity="ERROR",
                        category="HOURS_ARITHMETIC",
                        location="Учебный план (Стажировка / СРС)",
                        message=(
                            f"Для программы без аудиторных занятий лекции, практики и лабораторные должны быть равны 0 "
                            f"во всех строках, включая итоговую (сейчас: ЛК={decl_lk:g}, ПР={decl_pr:g}, ЛБ={decl_lb:g})."
                        ),
                        suggested_fix="Обнулить ЛК, ПР и ЛБ во всех темах, модулях и итоговой строке, перенеся часы в СРС.",
                    )
                )
        # 1.6. Проверка минимальной доли практики (Запрос № 1: практики не меньше половины всех часов)
        min_prac_share = query.min_practice_share if query and query.min_practice_share is not None else None
        total_prac_and_lab = calc_mod_pr + calc_mod_lb
        if min_prac_share is not None:
            required_prac_hours = kit.total_hours * min_prac_share
            if total_prac_and_lab + 1e-2 < required_prac_hours:
                issues.append(
                    ValidationIssue(
                        code="INSUFFICIENT_PRACTICE_SHARE",
                        severity="ERROR",
                        category="HOURS_ARITHMETIC",
                        location="Учебный план (Доля практических занятий)",
                        message=(
                            f"Объем практических и лабораторных работ ({total_prac_and_lab:g} ч.) меньше требуемого порога "
                            f"({required_prac_hours:g} ч., т.е. {min_prac_share * 100:g}% от {kit.total_hours:g} ч.)."
                        ),
                        suggested_fix=f"Увеличить часы практических занятий минимум до {required_prac_hours:g} ч.",
                    )
                )
        # 1.7. Проверка календарного учебного графика (Запрос № 1: 8 занятий по 2 ч = 16 ч за 4 недели; Запрос № 5: <= 4 ч/день)
        sched = kit.calendar_schedule
        if query and query.max_hours_per_day is not None:
            if sched.max_hours_per_day > query.max_hours_per_day + 1e-2:
                issues.append(
                    ValidationIssue(
                        code="SCHEDULE_DAILY_LIMIT_EXCEEDED",
                        severity="ERROR",
                        category="HOURS_ARITHMETIC",
                        location="Календарный учебный график",
                        message=(
                            f"Дневная нагрузка в графике ({sched.max_hours_per_day:g} ч./день) превышает ограничение запроса "
                            f"({query.max_hours_per_day:g} ч./день)."
                        ),
                        suggested_fix=f"Ограничить режим занятий до {query.max_hours_per_day:g} ч. в день.",
                    )
                )
        if sched.sessions_count and sched.hours_per_session:
            sched_product = sched.sessions_count * sched.hours_per_session
            if not math.isclose(sched_product, kit.total_hours, abs_tol=1e-2):
                issues.append(
                    ValidationIssue(
                        code="SCHEDULE_SESSIONS_TOTAL_MISMATCH",
                        severity="ERROR",
                        category="HOURS_ARITHMETIC",
                        location="Календарный учебный график",
                        message=(
                            f"Произведение числа занятий ({sched.sessions_count}) на длительность занятия "
                            f"({sched.hours_per_session:g} ч.) равно {sched_product:g} ч., что не совпадает с объемом программы ({kit.total_hours:g} ч.)."
                        ),
                        suggested_fix="Синхронизировать количество занятий в календарном графике с общим объемом часов.",
                    )
                )
        if sched.total_days > 0 and sched.max_hours_per_day > 0:
            max_capacity = sched.total_days * sched.max_hours_per_day
            if max_capacity + 1e-2 < kit.total_hours:
                issues.append(
                    ValidationIssue(
                        code="SCHEDULE_CAPACITY_TOO_SMALL",
                        severity="ERROR",
                        category="HOURS_ARITHMETIC",
                        location="Календарный учебный график",
                        message=(
                            f"За {sched.total_days} учебных дней при нагрузке до {sched.max_hours_per_day:g} ч./день "
                            f"невозможно освоить {kit.total_hours:g} ч. (максимум {max_capacity:g} ч.)."
                        ),
                        suggested_fix="Увеличить количество учебных дней или скорректировать график.",
                    )
                )
        summary = {
            "total_hours": kit.total_hours,
            "calculated_total": expected_course_total,
            "lecture_hours": calc_mod_lk,
            "practice_hours": calc_mod_pr,
            "lab_hours": calc_mod_lb,
            "self_study_hours": calc_mod_srs,
            "control_hours": calc_mod_ctrl,
            "final_attestation_hours": ias_h,
            "practice_share": round(total_prac_and_lab / kit.total_hours, 3) if kit.total_hours > 0 else 0.0,
        }
        return issues, summary
    # -----------------------------------------------------------------
    # БЛОК 2: КАРТА КОМПЕТЕНЦИЙ, ПОКРЫТИЕ РО, ТЕСТЫ И ДИДАКТИКА
    # -----------------------------------------------------------------
    def _validate_competencies_and_coverage(
        self, kit: ProgramKit, query: Optional[UserQuerySpec]
    ) -> Tuple[List[ValidationIssue], Dict[str, Any]]:
        issues: List[ValidationIssue] = []
        all_topics: Dict[str, CurriculumTopic] = {}
        topic_order_list: List[CurriculumTopic] = []
        for mod in kit.modules:
            for t in mod.topics:
                all_topics[t.id] = t
                topic_order_list.append(t)
        lo_ids: Set[str] = {lo.id for lo in kit.learning_outcomes}
        if not lo_ids:
            issues.append(
                ValidationIssue(
                    code="EMPTY_LEARNING_OUTCOMES",
                    severity="ERROR",
                    category="COMPETENCY_COVERAGE",
                    location="Планируемые результаты обучения",
                    message="Список результатов обучения (Знания / Умения / Навыки) пуст.",
                    suggested_fix="Сформировать минимум по 2–3 результата обучения в каждой категории.",
                )
            )
        # Проверяем наличие всех обязательных категорий РО (Знания, Умения, Навыки/Практический опыт)
        lo_categories = {lo.category for lo in kit.learning_outcomes}
        if "knowledge" not in lo_categories or "skill" not in lo_categories:
            issues.append(
                ValidationIssue(
                    code="MISSING_ZUN_CATEGORIES",
                    severity="ERROR",
                    category="COMPETENCY_COVERAGE",
                    location="Планируемые результаты обучения",
                    message="В результатах обучения должны быть выделены как минимум Знания (knowledge), Умения (skill) и Навыки/Практический опыт.",
                    suggested_fix="Добавить недостающие категории результатов обучения.",
                )
            )
        # Считаем покрытие каждого результата обучения темами и оценочными средствами
        lo_to_topics: Dict[str, Set[str]] = {lo_id: set() for lo_id in lo_ids}
        lo_to_assessments: Dict[str, Set[str]] = {lo_id: set() for lo_id in lo_ids}
        # Привязка из тем
        for t in topic_order_list:
            for lid in t.learning_outcome_ids:
                if lid in lo_to_topics:
                    lo_to_topics[lid].add(t.id)
                else:
                    issues.append(
                        ValidationIssue(
                            code="UNKNOWN_LO_IN_TOPIC",
                            severity="ERROR",
                            category="COMPETENCY_COVERAGE",
                            location=f"Тема '{t.id}'",
                            message=f"Тема ссылается на несуществующий ID результата обучения '{lid}'.",
                            suggested_fix="Указать корректный ID из списка learning_outcomes.",
                        )
                    )
        # Привязка из карты компетенций
        for comp in kit.competencies:
            for lid in comp.learning_outcome_ids:
                if lid in lo_to_topics:
                    lo_to_topics[lid].update(comp.topic_ids)
                    lo_to_assessments[lid].update(comp.assessment_ids)
        # Проверка тестовых вопросов (ключи, варианты, привязка к РО)
        for q in kit.test_questions:
            if len(q.options) < 2:
                issues.append(
                    ValidationIssue(
                        code="TEST_QUESTION_FEW_OPTIONS",
                        severity="ERROR",
                        category="COMPETENCY_COVERAGE",
                        location=f"Тестовый вопрос '{q.id}'",
                        message="В тестовом вопросе меньше 2 вариантов ответа.",
                        suggested_fix="Добавить 4 варианта ответа.",
                    )
                )
            if not (0 <= q.correct_option_index < len(q.options)):
                issues.append(
                    ValidationIssue(
                        code="TEST_QUESTION_INVALID_KEY",
                        severity="ERROR",
                        category="COMPETENCY_COVERAGE",
                        location=f"Тестовый вопрос '{q.id}'",
                        message=f"Индекс правильного ответа ({q.correct_option_index}) выходит за границы списка вариантов ({len(q.options)}).",
                        suggested_fix="Указать валидный ключ ответа (0..N-1).",
                    )
                )
            if not q.learning_outcome_ids:
                issues.append(
                    ValidationIssue(
                        code="TEST_QUESTION_NO_LO_LINK",
                        severity="ERROR",
                        category="COMPETENCY_COVERAGE",
                        location=f"Тестовый вопрос '{q.id}'",
                        message="Тестовый вопрос не привязан ни к одному результату обучения.",
                        suggested_fix="Привязать вопрос к соответствующему ID результата обучения.",
                    )
                )
            for lid in q.learning_outcome_ids:
                if lid in lo_to_assessments:
                    lo_to_assessments[lid].add(f"test:{q.id}")
        # Проверка точного количества вопросов теста, если задано в запросе (Запрос № 1: 10 вопросов)
        if query and query.required_test_questions_count is not None:
            if len(kit.test_questions) != query.required_test_questions_count:
                issues.append(
                    ValidationIssue(
                        code="TEST_QUESTIONS_COUNT_MISMATCH",
                        severity="ERROR",
                        category="COMPETENCY_COVERAGE",
                        location="Оценочные материалы (Тест)",
                        message=(
                            f"В тесте содержится {len(kit.test_questions)} вопросов, "
                            f"хотя в запросе требовалось ровно {query.required_test_questions_count}."
                        ),
                        suggested_fix=f"Сформировать тест ровно из {query.required_test_questions_count} вопросов с ключами.",
                    )
                )
        # Проверка заданий семинаров и вопросов собеседования
        for task in kit.seminar_tasks:
            if not task.evaluation_criteria:
                issues.append(
                    ValidationIssue(
                        code="SEMINAR_TASK_NO_CRITERIA",
                        severity="ERROR",
                        category="COMPETENCY_COVERAGE",
                        location=f"Задание семинара '{task.id}' ({task.title})",
                        message="У практического задания отсутствуют критерии оценки.",
                        suggested_fix="Добавить четкие критерии оценивания выполнения задания.",
                    )
                )
            for lid in task.learning_outcome_ids:
                if lid in lo_to_assessments:
                    lo_to_assessments[lid].add(f"task:{task.id}")
        for iq in kit.interview_questions:
            for lid in iq.learning_outcome_ids:
                if lid in lo_to_assessments:
                    lo_to_assessments[lid].add(f"interview:{iq.id}")
        # Главная проверка: «Каждый результат обучения раскрыт хотя бы в одной теме и проверяется хотя бы двумя оценочными средствами»
        uncovered_by_topics: List[str] = []
        under_assessed_los: List[str] = []
        for lo in kit.learning_outcomes:
            t_count = len(lo_to_topics.get(lo.id, set()))
            a_count = len(lo_to_assessments.get(lo.id, set()))
            if t_count < 1:
                uncovered_by_topics.append(lo.id)
                issues.append(
                    ValidationIssue(
                        code="LO_NOT_COVERED_BY_TOPIC",
                        severity="ERROR",
                        category="COMPETENCY_COVERAGE",
                        location=f"Результат обучения '{lo.id}'",
                        message=f"Результат обучения '{lo.id}' ({lo.description[:60]}...) не раскрыт ни в одной теме учебного плана.",
                        suggested_fix="Привязать результат обучения минимум к одной теме УТП.",
                    )
                )
            if a_count < 2:
                under_assessed_los.append(lo.id)
                issues.append(
                    ValidationIssue(
                        code="LO_FEWER_THAN_TWO_ASSESSMENTS",
                        severity="ERROR",
                        category="COMPETENCY_COVERAGE",
                        location=f"Результат обучения '{lo.id}'",
                        message=(
                            f"Результат обучения '{lo.id}' проверяется только {a_count} оценочным(и) средством(ами), "
                            f"а по требованиям АЦР нужно минимум 2."
                        ),
                        suggested_fix="Добавить тестовый вопрос, практическое задание или вопрос собеседования для данного РО.",
                    )
                )
        # Проверка дидактического порядка тем для начинающих (Запрос № 1: переменные, условия и циклы ДО классов или API)
        if kit.entry_level == "zero" or (query and query.entry_level == "zero"):
            basic_idx: Optional[int] = None
            advanced_idx: Optional[int] = None
            advanced_title = ""
            for idx, t in enumerate(topic_order_list):
                text_combined = f"{t.title} {t.content_summary}".lower()
                if basic_idx is None and re.search(r"переменн|услов|цикл|базов\w+\s+синтаксис|типы\s+данных", text_combined):
                    basic_idx = idx
                if advanced_idx is None and re.search(r"\bкласс\w*\b|\bооп\b|\bapi\b|объектно-ориентир", text_combined):
                    advanced_idx = idx
                    advanced_title = f"{t.id}. {t.title}"
            if advanced_idx is not None and (basic_idx is None or advanced_idx < basic_idx):
                issues.append(
                    ValidationIssue(
                        code="PEDAGOGICAL_ORDER_VIOLATION",
                        severity="ERROR",
                        category="PEDAGOGY_AND_AUDIENCE",
                        location=f"Тема '{advanced_title}'",
                        message=(
                            "Нарушена дидактическая последовательность для нулевого уровня: сложные темы (классы/ООП или API) "
                            "появляются раньше базовых тем (переменные, условия, циклы)."
                        ),
                        suggested_fix="Переместить изучение переменных, условий и циклов в первые темы программы.",
                    )
                )
        # Проверка обязательных тем из запроса (например, Запрос № 3: pandas, визуализация, SQL, основы статистики + вводный блок Python)
        if query and query.required_topics:
            all_topics_text = " ".join(f"{t.title} {t.content_summary}" for t in topic_order_list).lower()
            for req_topic in query.required_topics:
                req_low = req_topic.lower().strip()
                # Учитываем морфологическую основу для русских терминов (например, "сети" -> "сет")
                stem = req_low[:-1] if len(req_low) >= 4 and req_low[-1] in "иыаяеоу" else req_low
                if req_low not in all_topics_text and stem not in all_topics_text:
                    issues.append(
                        ValidationIssue(
                            code="MISSING_REQUIRED_TOPIC",
                            severity="ERROR",
                            category="PEDAGOGY_AND_AUDIENCE",
                            location="Учебно-тематический план",
                            message=f"Обязательная тема из запроса '{req_topic}' не найдена в учебно-тематическом плане.",
                            suggested_fix=f"Выделить '{req_topic}' в отдельную тему или модуль с часами.",
                        )
                    )
        # Проверка наличия и совпадения слайдов лекций (.pptx) с темами УТП
        if not kit.is_self_study_only and not kit.lecture_decks:
            issues.append(
                ValidationIssue(
                    code="MISSING_LECTURE_SLIDES",
                    severity="ERROR",
                    category="COMPETENCY_COVERAGE",
                    location="Планы и слайды лекций",
                    message="Отсутствуют планы и слайды лекций (.pptx).",
                    suggested_fix="Сгенерировать планы и слайды лекций для целевых тем/модулей.",
                )
            )
        for deck in kit.lecture_decks:
            if deck.topic_id not in all_topics:
                issues.append(
                    ValidationIssue(
                        code="SLIDE_DECK_UNKNOWN_TOPIC",
                        severity="ERROR",
                        category="COMPETENCY_COVERAGE",
                        location=f"Слайды к теме '{deck.topic_id}'",
                        message=f"Презентация привязана к теме '{deck.topic_id}', которой нет в учебном плане.",
                        suggested_fix="Синхронизировать ID и названия тем между УТП и презентациями.",
                    )
                )
            else:
                utp_topic = all_topics[deck.topic_id]
                if utp_topic.title.strip().lower() != deck.topic_title.strip().lower():
                    issues.append(
                        ValidationIssue(
                            code="SLIDE_DECK_TITLE_MISMATCH",
                            severity="ERROR",
                            category="COMPETENCY_COVERAGE",
                            location=f"Слайды к теме '{deck.topic_id}'",
                            message=(
                                f"Название темы в слайдах ('{deck.topic_title}') не совпадает с названием темы в УТП ('{utp_topic.title}')."
                            ),
                            suggested_fix="Установить точное совпадение заголовка презентации и темы в УТП.",
                        )
                    )
        summary = {
            "total_learning_outcomes": len(lo_ids),
            "uncovered_by_topics": uncovered_by_topics,
            "under_assessed_learning_outcomes": under_assessed_los,
            "total_test_questions": len(kit.test_questions),
            "total_seminar_tasks": len(kit.seminar_tasks),
            "total_interview_questions": len(kit.interview_questions),
            "total_slide_decks": len(kit.lecture_decks),
        }
        return issues, summary
    # -----------------------------------------------------------------
    # БЛОК 3: ШАБЛОН АЦР, НОРМАТИВНАЯ БАЗА, ПРОФСТАНДАРТЫ И РЕКВИЗИТЫ
    # -----------------------------------------------------------------
    def _validate_template_and_regulations(
        self, kit: ProgramKit, query: Optional[UserQuerySpec]
    ) -> List[ValidationIssue]:
        issues: List[ValidationIssue] = []
        if query and query.program_type and kit.template_type != query.program_type:
            issues.append(
                ValidationIssue(
                    code="WRONG_TEMPLATE_SELECTED",
                    severity="ERROR",
                    category="TEMPLATE_NORMATIVE",
                    location="ProgramKit.template_type",
                    message=f"Выбран шаблон '{kit.template_type}', хотя по запросу требуется '{query.program_type}'.",
                    suggested_fix=f"Переключить шаблон на '{query.program_type}'.",
                )
            )
        reg_joined = " ".join(kit.regulatory_documents).lower()
        # Правила для ДООП (Общеразвивающая программа — Приказ Минпросвещения № 629)
        if kit.template_type == "doop":
            if "сертификат" not in kit.issued_document.lower():
                issues.append(
                    ValidationIssue(
                        code="DOOP_WRONG_DOCUMENT",
                        severity="ERROR",
                        category="TEMPLATE_NORMATIVE",
                        location="Выдаваемый документ",
                        message=f"Для общеразвивающей программы (ДООП) должен выдаваться сертификат, а указано: '{kit.issued_document}'.",
                        suggested_fix="Указать: «Сертификат об освоении программы».",
                    )
                )
            if "629" not in reg_joined:
                issues.append(
                    ValidationIssue(
                        code="DOOP_MISSING_ORDER_629",
                        severity="ERROR",
                        category="TEMPLATE_NORMATIVE",
                        location="Нормативная база (ДООП)",
                        message="В нормативной базе общеразвивающей программы отсутствует ссылка на Приказ Минпросвещения России № 629.",
                        suggested_fix="Добавить Приказ Минпросвещения России от 27.07.2022 № 629.",
                    )
                )
            # Чек-лист Запроса № 1: «В программе нет привязки к профстандартам и трудовым функциям, потому что для общеразвивающей программы она не требуется»
            has_prof_links = any(c.prof_standard_link is not None for c in kit.competencies)
            if has_prof_links:
                issues.append(
                    ValidationIssue(
                        code="DOOP_FORBIDDEN_PROFSTANDARD_LINK",
                        severity="ERROR",
                        category="TEMPLATE_NORMATIVE",
                        location="Карта компетенций (ДООП)",
                        message="В общеразвивающей программе (ДООП) обнаружена привязка к профстандартам и трудовым функциям, которая не требуется и запрещена чек-листом.",
                        suggested_fix="Удалить ссылки на профстандарты и коды трудовых функций для шаблона ДООП.",
                    )
                )
        # Правила для ДПП ПК (Повышение квалификации — 273-ФЗ, Приказ Минобрнауки № 499)
        elif kit.template_type == "dpp_pk":
            if "удостоверение" not in kit.issued_document.lower():
                issues.append(
                    ValidationIssue(
                        code="DPPPK_WRONG_DOCUMENT",
                        severity="ERROR",
                        category="TEMPLATE_NORMATIVE",
                        location="Выдаваемый документ",
                        message=f"Для программы повышения квалификации должно выдаваться удостоверение о повышении квалификации, а указано: '{kit.issued_document}'.",
                        suggested_fix="Указать: «Удостоверение о повышении квалификации установленного образца».",
                    )
                )
            if "273-фз" not in reg_joined or "499" not in reg_joined:
                issues.append(
                    ValidationIssue(
                        code="DPPPK_MISSING_REGULATIONS",
                        severity="ERROR",
                        category="TEMPLATE_NORMATIVE",
                        location="Нормативная база (ДПП ПК)",
                        message="Программа повышения квалификации обязана опираться на 273-ФЗ «Об образовании в РФ» и Приказ Минобрнауки России № 499.",
                        suggested_fix="Включить 273-ФЗ и Приказ Минобрнауки России от 01.07.2013 № 499 в нормативно-правовую базу.",
                    )
                )
            # Проверка привязки к реальным профстандартам и кодам трудовых функций
            prof_links = [c.prof_standard_link for c in kit.competencies if c.prof_standard_link is not None]
            if not prof_links:
                issues.append(
                    ValidationIssue(
                        code="DPPPK_MISSING_PROFSTANDARD",
                        severity="ERROR",
                        category="TEMPLATE_NORMATIVE",
                        location="Карта компетенций (ДПП ПК)",
                        message="В программе повышения квалификации компетенции не привязаны ни к одному профессиональному стандарту.",
                        suggested_fix="Привязать компетенции к действующему профстандарту из реестра Минтруда РФ с указанием кодов трудовых функций.",
                    )
                )
            else:
                for link in prof_links:
                    std = VERIFIED_PROF_STANDARDS.get(link.standard_code)
                    if std is None:
                        issues.append(
                            ValidationIssue(
                                code="UNVERIFIED_PROFSTANDARD_CODE",
                                severity="ERROR",
                                category="TEMPLATE_NORMATIVE",
                                location=f"Профстандарт '{link.standard_code}'",
                                message=f"Код профстандарта '{link.standard_code}' ({link.standard_title}) не найден в верифицированном реестре Минтруда РФ.",
                                suggested_fix=f"Использовать подтвержденный профстандарт из реестра ({', '.join(sorted(VERIFIED_PROF_STANDARDS.keys()))}).",
                            )
                        )
                    else:
                        if link.labor_function_code not in std.labor_functions:
                            issues.append(
                                ValidationIssue(
                                    code="INVALID_LABOR_FUNCTION_CODE",
                                    severity="ERROR",
                                    category="TEMPLATE_NORMATIVE",
                                    location=f"Профстандарт {std.code}, ТФ '{link.labor_function_code}'",
                                    message=(
                                        f"Код трудовой функции '{link.labor_function_code}' не соответствует профстандарту {std.code} «{std.title}». "
                                        f"Доступные коды: {', '.join(std.labor_functions.keys())}."
                                    ),
                                    suggested_fix=f"Указать верный код трудовой функции (например, {next(iter(std.labor_functions.keys()))}).",
                                )
                            )
        # Проверка пустых полей для реквизитов приказа, ФИО преподавателей и подписей
        req = kit.requisites
        if re.search(r"\d", req.order_number):
            issues.append(
                ValidationIssue(
                    code="REQUISITES_FILLED_ORDER_NUMBER",
                    severity="ERROR",
                    category="TEMPLATE_NORMATIVE",
                    location="Реквизиты приказа",
                    message=f"Номер приказа заполнен конкретным числом ('{req.order_number}'), хотя должен оставаться пустым полем для заполнения человеком.",
                    suggested_fix="Оставить шаблонное поле: 'Приказом № ____-ОБР от «___» __________ 2026 г.'",
                )
            )
        if req.approver_name.strip() or req.approver_signature.strip() or any(n.strip() for n in req.staff_table_Order_names):
            issues.append(
                ValidationIssue(
                    code="REQUISITES_FILLED_PERSONAL_NAMES",
                    severity="ERROR",
                    category="TEMPLATE_NORMATIVE",
                    location="Кадровое обеспечение / Подписи",
                    message="Поля ФИО преподавателей или подписи руководителя заполнены, хотя должны оставаться пустыми полями.",
                    suggested_fix="Очистить ФИО преподавателей и поля подписей.",
                )
            )
        return issues
    # -----------------------------------------------------------------
    # БЛОК 4: ИСТОЧНИКИ, СЕРВИСЫ В РФ, АКТУАЛЬНОСТЬ 2026 И АНТИПЛАГИАТ
    # -----------------------------------------------------------------
    def _validate_sources_and_tools(
        self, kit: ProgramKit, query: Optional[UserQuerySpec]
    ) -> List[ValidationIssue]:
        issues: List[ValidationIssue] = []
        # Собираем весь текст программы для сквозного сканирования
        full_text_parts = [
            kit.program_title,
            kit.explanatory_note,
            kit.relevance_text,
            kit.goal_text,
            " ".join(kit.tasks_list),
            " ".join(kit.software_and_tools),
            " ".join(t.title + " " + t.content_summary for m in kit.modules for t in m.topics),
            " ".join(s.title + " " + s.scenario_description for s in kit.seminar_tasks),
            " ".join(l.citation for l in kit.literature),
        ]
        full_text_lower = " ".join(full_text_parts).lower()
        # 4.1. Проверка ограничений по доступности ИИ-сервисов в России (Запрос № 2)
        enforce_rf_only = query.russia_available_services_only if query else False
        if enforce_rf_only:
            for foreign_srv, ru_alt in RESTRICTED_FOREIGN_SERVICES_IN_RF.items():
                if re.search(rf"\b{re.escape(foreign_srv)}\b", full_text_lower):
                    issues.append(
                        ValidationIssue(
                            code="RESTRICTED_FOREIGN_SERVICE_USED",
                            severity="ERROR",
                            category="TOOLS_AND_SOURCES",
                            location=f"Упоминание сервиса '{foreign_srv}'",
                            message=(
                                f"В программе упомянут зарубежный сервис '{foreign_srv}', который недоступен в России напрямую "
                                f"(нарушение требования запроса)."
                            ),
                            suggested_fix=f"Заменить '{foreign_srv}' на доступный в РФ аналог: {ru_alt}.",
                        )
                    )
        # 4.2. Проверка актуальности инструментов и библиотек на 2026 год (Запрос № 3)
        for pattern, reason in DEPRECATED_TECH_PATTERNS_2026:
            if re.search(pattern, full_text_lower):
                issues.append(
                    ValidationIssue(
                        code="DEPRECATED_TOOL_OR_LIBRARY_2026",
                        severity="ERROR",
                        category="TOOLS_AND_SOURCES",
                        location="Инструменты и содержание тем",
                        message=f"Обнаружено использование устаревшей технологии на 2026 год: {reason}.",
                        suggested_fix=reason,
                    )
                )
        # 4.3. Проверка литературы и маркировки неподтвержденных источников
        if not kit.literature:
            issues.append(
                ValidationIssue(
                    code="EMPTY_LITERATURE_LIST",
                    severity="ERROR",
                    category="TOOLS_AND_SOURCES",
                    location="Список литературы",
                    message="Список литературы и интернет-ресурсов пуст.",
                    suggested_fix="Заполнить основную, дополнительную литературу и интернет-ресурсы.",
                )
            )
        for src in kit.literature:
            if not src.is_verified and not src.verification_tag.strip():
                issues.append(
                    ValidationIssue(
                        code="UNMARKED_UNVERIFIED_SOURCE",
                        severity="ERROR",
                        category="TOOLS_AND_SOURCES",
                        location=f"Источник '{src.citation[:50]}...'",
                        message="Неподтвержденный источник не имеет явной пометки для методиста.",
                        suggested_fix="Пометить источник тегом '[Требует подтверждения методистом]'.",
                    )
                )
        # 4.4. Проверка на дословное копирование из референсов («Нельзя дословно копировать чужие курсы и референсы»)
        if query and query.reference_texts:
            for idx, ref_text in enumerate(query.reference_texts, 1):
                overlap_ratio = _compute_5gram_overlap(full_text_lower, ref_text.lower())
                if overlap_ratio > 0.35:
                    issues.append(
                        ValidationIssue(
                            code="EXCESSIVE_VERBATIM_COPYING",
                            severity="ERROR",
                            category="TOOLS_AND_SOURCES",
                            location=f"Сравнение с референсом #{idx}",
                            message=(
                                f"Доля дословного совпадения текста программы с приложенным референсом составляет "
                                f"{overlap_ratio * 100:.1f}% (превышен порог оригинальности 35%)."
                            ),
                            suggested_fix="Сохранить структуру разделов референса, но полностью переписать текстовое наполнение под новую задачу.",
                        )
                    )
        return issues

def _compute_5gram_overlap(candidate_text: str, reference_text: str) -> float:
    """Вычисляет долю дословного заимствования 5-грамм слов между текстом программы и референсом."""
    cand_words = re.findall(r"[а-яёa-z0-9]+", candidate_text.lower())
    ref_words = re.findall(r"[а-яёa-z0-9]+", reference_text.lower())
    if len(cand_words) < 10 or len(ref_words) < 10:
        return 0.0
    cand_ngrams = {tuple(cand_words[i : i + 5]) for i in range(len(cand_words) - 4)}
    ref_ngrams = {tuple(ref_words[i : i + 5]) for i in range(len(ref_words) - 4)}
    denom = min(len(cand_ngrams), len(ref_ngrams))
    if denom == 0:
        return 0.0
    return len(cand_ngrams.intersection(ref_ngrams)) / denom

# =====================================================================
# 6. ДЕТЕРМИНИРОВАННЫЙ АВТОКОРРЕКТОР ЧАСОВ (auto_fix_hours)

# =====================================================================

def auto_fix_hours(kit: ProgramKit) -> ProgramKit:
    """
    Детерминированно пересчитывает и синхронизирует все суммы часов снизу вверх:
    1) Если в теме есть rpd_activities — выставляет часы темы строго по суммам активностей РПД;
       иначе выставляет topic.total_hours = ЛК + ПР + ЛБ + СРС + ТК.
    2) Если режим is_self_study_only=True — переносит все часы ЛК, ПР и ЛБ в СРС.
    3) Пересчитывает каждый модуль как точную сумму его тем.
    4) Пересчитывает общий объем и строку declared_totals («Итого») как сумму модулей + ИА.
    """
    total_lk = 0.0
    total_pr = 0.0
    total_lb = 0.0
    total_srs = 0.0
    total_ctrl = 0.0
    total_modules_hours = 0.0
    for mod in kit.modules:
        mod_lk = 0.0
        mod_pr = 0.0
        mod_lb = 0.0
        mod_srs = 0.0
        mod_ctrl = 0.0
        for topic in mod.topics:
            if kit.is_self_study_only:
                moved = topic.lecture_hours + topic.practice_hours + topic.lab_hours
                topic.self_study_hours += moved
                topic.lecture_hours = 0.0
                topic.practice_hours = 0.0
                topic.lab_hours = 0.0
                for act in topic.rpd_activities:
                    if act.activity_type in ("ЛК", "ПР", "ЛБ"):
                        act.activity_type = "СРС"
            if topic.rpd_activities:
                topic.lecture_hours = sum(a.hours for a in topic.rpd_activities if a.activity_type == "ЛК")
                topic.practice_hours = sum(a.hours for a in topic.rpd_activities if a.activity_type == "ПР")
                topic.lab_hours = sum(a.hours for a in topic.rpd_activities if a.activity_type == "ЛБ")
                topic.self_study_hours = sum(a.hours for a in topic.rpd_activities if a.activity_type == "СРС")
                topic.control_hours = sum(a.hours for a in topic.rpd_activities if a.activity_type == "ТК")
            topic.total_hours = (
                topic.lecture_hours
                + topic.practice_hours
                + topic.lab_hours
                + topic.self_study_hours
                + topic.control_hours
            )
            mod_lk += topic.lecture_hours
            mod_pr += topic.practice_hours
            mod_lb += topic.lab_hours
            mod_srs += topic.self_study_hours
            mod_ctrl += topic.control_hours
        mod.lecture_hours = mod_lk
        mod.practice_hours = mod_pr
        mod.lab_hours = mod_lb
        mod.self_study_hours = mod_srs
        mod.control_hours = mod_ctrl
        mod.total_hours = mod_lk + mod_pr + mod_lb + mod_srs + mod_ctrl
        total_lk += mod_lk
        total_pr += mod_pr
        total_lb += mod_lb
        total_srs += mod_srs
        total_ctrl += mod_ctrl
        total_modules_hours += mod.total_hours
    kit.total_hours = total_modules_hours + kit.final_attestation.hours
    kit.declared_totals = {
        "total": kit.total_hours,
        "lecture": total_lk,
        "practice": total_pr,
        "lab": total_lb,
        "self_study": total_srs,
        "control": total_ctrl,
        "final_attestation": kit.final_attestation.hours,
    }
    return kit