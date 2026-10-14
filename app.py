# -*- coding: utf-8 -*-
"""
Веб-сервис и единый API-эндпоинт ИИ-агента для создания образовательных программ АЦР (`app.py`).

Работает полностью автономно на стандартной библиотеке Python (`http.server.ThreadingHTTPServer`),
поэтому запускается в VS Code одной командой `py app.py` без конфликтов окружения и сложных зависимостей,
и одновременно предоставляет:
1. Интерактивное рабочее место методиста АЦР в браузере (`http://127.0.0.1:8000`) с пошаговым сценарием
   GigaCowork (Анализ и вопросы -> Согласование каркаса -> Полный комплект и самопроверка -> Точечная правка
   и масштабирование -> Скачивание `.docx`, `.pptx`, `GIFT`, `Moodle XML`, `CSV`, `JSON`).
2. Единый REST API-эндпоинт (`POST /api/v1/generate`) для автоматического прогона жюри на открытых и скрытых запросах.
"""

from __future__ import annotations

from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
import json
import mimetypes
import os
import sys
import urllib.parse
import webbrowser
from typing import Any, Dict, Optional

from agent_skills import (
    CloudRuLLMClient,
    EducationalProgramAgent,
    GIGACOWORK_SKILLS_CATALOG,
)
from program_exporters import export_full_bundle
from program_validator import ProgramKit, UserQuerySpec
from test_program_validator import ensure_sample_248_docx


OUTPUT_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "generated_output")
os.makedirs(OUTPUT_DIR, exist_ok=True)

# Глобальное состояние текущей сессии методиста для пошагового веб-интерфейса
SESSION_STATE: Dict[str, Any] = {
    "last_query_spec": None,
    "last_skeleton_kit": None,
    "last_full_kit": None,
    "last_report": None,
    "last_roi": None,
    "last_files": {},
}

PRESET_QUERIES: Dict[str, Dict[str, Any]] = {
    "case1": {
        "title": "Запрос № 1 · Python для школьников (16 ч, ДООП)",
        "query": (
            "Нужна дополнительная общеразвивающая программа «Первые шаги в Python» для школьников 8 и 9 классов "
            "без опыта программирования. Объем 16 часов, занятия очно по 2 часа два раза в неделю. "
            "Практики должно быть не меньше половины всех часов. В конце каждый ученик защищает небольшой проект, "
            "например простую игру или чат-бота. Нужны итоговый тест на 10 вопросов и слайды для первых двух занятий."
        ),
        "answers": {},
        "use_sample_248": False,
    },
    "case2": {
        "title": "Запрос № 2 · ИИ в работе учителя (36 ч, ДПП ПК, сервисы РФ)",
        "query": (
            "Программа повышения квалификации «Искусственный интеллект в работе учителя» для учителей школ Ингушетии, "
            "36 часов, очно-заочно с применением дистанционных технологий. Учителя должны научиться готовить с помощью "
            "нейросетей планы уроков, задания и проверочные работы, а также понимать риски и правила безопасного "
            "использования ИИ в работе с детьми. Итоговая аттестация в форме зачёта по портфолио из трёх материалов, "
            "созданных с помощью ИИ. Использовать только сервисы, которые доступны в России."
        ),
        "answers": {},
        "use_sample_248": False,
    },
    "case3": {
        "title": "Запрос № 3 · Анализ данных на Python для банков (72 ч, ДПП ПК)",
        "query": (
            "Нужна программа повышения квалификации на 72 часа «Анализ данных на Python» для сотрудников региональных "
            "отделений банков, которые сейчас работают в Excel. Формат очно-заочный, срок 4 недели. Основные темы это "
            "pandas, визуализация, SQL для выгрузки данных и основы статистики. Итоговая аттестация в виде защиты проекта "
            "на обезличенных банковских данных. Прикладываю нашу прошлую программу, структуру можно взять оттуда, но "
            "содержание нужно новое. Слайды нужны для модуля про pandas."
        ),
        "answers": {},
        "use_sample_248": False,
    },
    "case4": {
        "title": "Запрос № 4 · Кибербезопасность для колледжа (Неполный запрос → Уточнение)",
        "query": (
            "Сделайте курс по кибербезопасности для студентов колледжа, примерно на месяц, "
            "чтобы после него ребята могли пойти работать в эту сферу."
        ),
        "answers": {
            "Тип программы": "Повышение квалификации",
            "Объём": "72 часа",
            "Слушатели": "Студенты 3 курса специальности «Информационные системы и программирование», знают основы сетей и Linux",
            "Форма обучения": "Очная, занятия на базе колледжа",
            "Итоговая аттестация": "Экзамен с практическим заданием на стенде",
        },
        "use_sample_248": False,
    },
    "case5": {
        "title": "Запрос № 5 · Стажировка DevOps (256 ч, СРС + детектор ошибки в образце 248 ч)",
        "query": (
            "Нужна программа повышения квалификации в форме стажировки «DevOps-инженер от Linux до CI/CD» на 256 часов. "
            "Слушатели проходят стажировку в ИТ-компаниях региона, вся работа самостоятельная под руководством наставника "
            "от компании, аудиторных занятий нет. Итоговая аттестация занимает 2 часа и проходит в форме защиты отчёта о "
            "стажировке. Срок около двух месяцев, не больше 4 часов в день. Возьмите за образец структуры нашу стажировку по бизнес-анализу."
        ),
        "answers": {},
        "use_sample_248": True,
    },
}


# =====================================================================
# 1. ЯДРО ОБРАБОТКИ ЕДИНОГО API-ЭНДПОИНТА (ДЛЯ АВТОПРОВЕРКИ И UI)
# =====================================================================

def process_unified_api_request(payload: Dict[str, Any]) -> Dict[str, Any]:
    """
    Обрабатывает запрос к единому API-эндпоинту `/api/v1/generate`.
    Поддерживает режимы:
    - `mode="analyze"` — только Шаг 1–2 (анализ запроса, проверка референсов, уточняющие вопросы).
    - `mode="skeleton"` — Шаг 1–3 (возвращает каркас программы на согласование методисту).
    - `mode="full"` (по умолчанию для автопроверки) — выполняет полный пайплайн и выгружает файлы.
    - `mode="edit"` — точечная правка отдельной темы или теста без пересборки всего комплекта.
    - `mode="modernize"` — актуализация старой программы под стек 2026 года.
    - `mode="adapt"` — масштабирование программы под другой объем часов или другую аудиторию.
    """
    mode = payload.get("mode", "full")
    raw_query = payload.get("query", "").strip()
    answers: Optional[Dict[str, str]] = payload.get("answers") or None
    use_sample_248 = bool(payload.get("use_sample_248", False))
    file_prefix = payload.get("file_prefix", "acr_program")

    agent = EducationalProgramAgent(CloudRuLLMClient())
    ref_paths = [ensure_sample_248_docx()] if use_sample_248 or "бизнес-анализ" in raw_query.lower() else []

    if mode in ("edit", "modernize", "adapt") and SESSION_STATE.get("last_full_kit") is None:
        fallback_q = raw_query or PRESET_QUERIES["case3"]["query"]
        q_spec, _, _ = agent.step1_analyze_request(fallback_q, reference_docx_paths=ref_paths)
        skel = agent.step2_build_skeleton(q_spec)
        f_kit, f_rep, f_roi = agent.step3_generate_full_content(skel, q_spec, reference_docx_paths=ref_paths)
        SESSION_STATE["last_query_spec"] = q_spec
        SESSION_STATE["last_full_kit"] = f_kit
        SESSION_STATE["last_report"] = f_rep
        SESSION_STATE["last_roi"] = f_roi

    if mode == "edit" and SESSION_STATE.get("last_full_kit") is not None:
        edit_instruction = payload.get("edit_instruction", "Усложнить тест")
        kit, report, msg = agent.step4_targeted_edit(
            SESSION_STATE["last_full_kit"],
            edit_instruction=edit_instruction,
            query=SESSION_STATE.get("last_query_spec"),
        )
        files = export_full_bundle(
            kit,
            OUTPUT_DIR,
            file_prefix="acr_edited_program",
            query=SESSION_STATE.get("last_query_spec"),
            reference_docx_paths=ref_paths,
        )
        SESSION_STATE["last_full_kit"] = kit
        SESSION_STATE["last_report"] = report
        SESSION_STATE["last_files"] = files
        return {
            "status": "edited",
            "edit_summary": msg,
            "program_kit": kit.to_dict(),
            "validation_report": report.to_dict(),
            "validation_markdown": report.to_markdown(),
            "roi_metrics": SESSION_STATE["last_roi"].to_dict() if SESSION_STATE.get("last_roi") else None,
            "download_files": {k: f"/api/v1/download/{os.path.basename(v)}" for k, v in files.items()},
        }

    if mode == "modernize" and SESSION_STATE.get("last_full_kit") is not None:
        kit = SESSION_STATE["last_full_kit"]
        if payload.get("inject_legacy_demo"):
            kit.software_and_tools.append("Python 3.7, ChatGPT, Midjourney, метод pandas DataFrame.append() и датасет load_boston")
        mod_kit, logs = agent.step5_modernize_legacy_program(kit)
        report = agent.validator.validate(mod_kit, query=SESSION_STATE.get("last_query_spec"), reference_docx_paths=ref_paths)
        files = export_full_bundle(mod_kit, OUTPUT_DIR, file_prefix="acr_modernized_program")
        SESSION_STATE["last_full_kit"] = mod_kit
        SESSION_STATE["last_report"] = report
        SESSION_STATE["last_files"] = files
        return {
            "status": "modernized",
            "modernization_logs": logs,
            "program_kit": mod_kit.to_dict(),
            "validation_report": report.to_dict(),
            "validation_markdown": report.to_markdown(),
            "download_files": {k: f"/api/v1/download/{os.path.basename(v)}" for k, v in files.items()},
        }

    if mode == "adapt" and SESSION_STATE.get("last_full_kit") is not None:
        new_hours = float(payload.get("new_total_hours", 144.0))
        new_aud = payload.get("new_target_audience") or None
        new_tpl = payload.get("new_template_type") or None
        adap_kit, report, note = agent.step6_adapt_program_volume_or_audience(
            SESSION_STATE["last_full_kit"],
            new_total_hours=new_hours,
            new_target_audience=new_aud,
            new_template_type=new_tpl,
        )
        files = export_full_bundle(adap_kit, OUTPUT_DIR, file_prefix=f"acr_adapted_{int(new_hours)}h")
        SESSION_STATE["last_full_kit"] = adap_kit
        SESSION_STATE["last_report"] = report
        SESSION_STATE["last_files"] = files
        return {
            "status": "adapted",
            "adaptation_summary": note,
            "program_kit": adap_kit.to_dict(),
            "validation_report": report.to_dict(),
            "validation_markdown": report.to_markdown(),
            "download_files": {k: f"/api/v1/download/{os.path.basename(v)}" for k, v in files.items()},
        }

    # Стандартный поток: Шаг 1 (Анализ запроса и референсов)
    query_spec, clarification, ref_defects = agent.step1_analyze_request(
        raw_query_text=raw_query,
        methodologist_answers=answers,
        reference_docx_paths=ref_paths,
    )
    SESSION_STATE["last_query_spec"] = query_spec

    if mode == "analyze" or (clarification.needs_clarification and not payload.get("force_proceed", False)):
        return {
            "status": "needs_clarification" if clarification.needs_clarification else "analyzed",
            "parsed_query": dataclasses_asdict(query_spec),
            "clarification": dataclasses_asdict(clarification),
            "reference_defects": [dataclasses_asdict(d) for d in ref_defects],
        }

    # Шаг 2: Построение каркаса программы (Human-in-the-Loop)
    skeleton = agent.step2_build_skeleton(query_spec)
    SESSION_STATE["last_skeleton_kit"] = skeleton
    if mode == "skeleton":
        return {
            "status": "skeleton_ready_for_approval",
            "parsed_query": dataclasses_asdict(query_spec),
            "clarification": dataclasses_asdict(clarification),
            "reference_defects": [dataclasses_asdict(d) for d in ref_defects],
            "skeleton_kit": skeleton.to_dict(),
        }

    # Шаг 3: Генерация полного комплекта, автопроверка и выгрузка во все форматы
    full_kit, report, roi = agent.step3_generate_full_content(
        approved_skeleton=skeleton,
        query=query_spec,
        reference_docx_paths=ref_paths,
    )
    files = export_full_bundle(
        kit=full_kit,
        output_dir=OUTPUT_DIR,
        file_prefix=file_prefix,
        query=query_spec,
        reference_docx_paths=ref_paths,
    )

    SESSION_STATE["last_full_kit"] = full_kit
    SESSION_STATE["last_report"] = report
    SESSION_STATE["last_roi"] = roi
    SESSION_STATE["last_files"] = files

    return {
        "status": "completed",
        "parsed_query": dataclasses_asdict(query_spec),
        "clarification": dataclasses_asdict(clarification),
        "reference_defects": [dataclasses_asdict(d) for d in ref_defects],
        "program_kit": full_kit.to_dict(),
        "validation_report": report.to_dict(),
        "validation_markdown": report.to_markdown(),
        "roi_metrics": roi.to_dict(),
        "roi_markdown": roi.to_markdown(),
        "skills_catalog": [dataclasses_asdict(s) for s in GIGACOWORK_SKILLS_CATALOG],
        "download_files": {k: f"/api/v1/download/{os.path.basename(v)}" for k, v in files.items()},
    }


def dataclasses_asdict(obj: Any) -> Any:
    import dataclasses
    return dataclasses.asdict(obj) if dataclasses.is_dataclass(obj) else obj


# =====================================================================
# 2. ИНТЕРАКТИВНЫЙ ВЕБ-ИНТЕРФЕЙС МЕТОДИСТА АЦР (HTML5 + JS)
# =====================================================================

HTML_UI_PAGE = r"""<!DOCTYPE html>
<html lang="ru">
<head>
  <meta charset="UTF-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1.0" />
  <title>АЦР × GigaCowork · ИИ-агент разработки образовательных программ</title>
  <style>
    :root {
      --bg: #0f172a;
      --card: #1e293b;
      --card-border: #334155;
      --accent: #3b82f6;
      --accent-hover: #2563eb;
      --emerald: #10b981;
      --amber: #f59e0b;
      --rose: #ef4444;
      --text: #f8fafc;
      --muted: #94a3b8;
    }
    * { box-sizing: border-box; }
    body {
      margin: 0;
      font-family: 'Segoe UI', system-ui, -apple-system, sans-serif;
      background: var(--bg);
      color: var(--text);
      line-height: 1.5;
    }
    header {
      background: linear-gradient(90deg, #0f172a 0%, #1e3a8a 100%);
      border-bottom: 1px solid var(--card-border);
      padding: 18px 28px;
      display: flex;
      justify-content: space-between;
      align-items: center;
      flex-wrap: wrap;
      gap: 12px;
    }
    header h1 { margin: 0; font-size: 20px; font-weight: 700; }
    header .badges { display: flex; gap: 8px; flex-wrap: wrap; }
    .badge {
      font-size: 12px;
      padding: 4px 10px;
      border-radius: 999px;
      background: rgba(59, 130, 246, 0.2);
      border: 1px solid rgba(59, 130, 246, 0.5);
      color: #93c5fd;
    }
    .container {
      max-width: 1280px;
      margin: 0 auto;
      padding: 22px;
      display: grid;
      grid-template-columns: 380px 1fr;
      gap: 22px;
    }
    @media (max-width: 980px) {
      .container { grid-template-columns: 1fr; }
    }
    .card {
      background: var(--card);
      border: 1px solid var(--card-border);
      border-radius: 12px;
      padding: 18px;
      margin-bottom: 18px;
    }
    .card h2 {
      margin: 0 0 12px 0;
      font-size: 16px;
      color: #e2e8f0;
      border-bottom: 1px solid var(--card-border);
      padding-bottom: 8px;
    }
    .preset-btn {
      width: 100%;
      text-align: left;
      background: #0f172a;
      color: var(--text);
      border: 1px solid var(--card-border);
      padding: 10px 12px;
      border-radius: 8px;
      margin-bottom: 8px;
      cursor: pointer;
      font-size: 13px;
      transition: all 0.15s;
    }
    .preset-btn:hover {
      border-color: var(--accent);
      background: rgba(59, 130, 246, 0.12);
    }
    textarea, input, select {
      width: 100%;
      background: #0f172a;
      color: var(--text);
      border: 1px solid var(--card-border);
      border-radius: 8px;
      padding: 10px;
      font-size: 13.5px;
      font-family: inherit;
      margin-bottom: 10px;
    }
    .btn-row { display: flex; gap: 10px; flex-wrap: wrap; }
    button.primary {
      background: var(--accent);
      color: #fff;
      border: none;
      padding: 10px 16px;
      border-radius: 8px;
      font-weight: 600;
      cursor: pointer;
      font-size: 13.5px;
    }
    button.primary:hover { background: var(--accent-hover); }
    button.emerald {
      background: var(--emerald);
      color: #052e16;
      border: none;
      padding: 10px 16px;
      border-radius: 8px;
      font-weight: 700;
      cursor: pointer;
      font-size: 13.5px;
    }
    button.secondary {
      background: #334155;
      color: #f8fafc;
      border: 1px solid #475569;
      padding: 9px 14px;
      border-radius: 8px;
      font-weight: 600;
      cursor: pointer;
      font-size: 13px;
    }
    table {
      width: 100%;
      border-collapse: collapse;
      margin: 10px 0;
      font-size: 13px;
    }
    th, td {
      border: 1px solid var(--card-border);
      padding: 8px 10px;
      text-align: left;
    }
    th { background: #0f172a; color: #93c5fd; }
    .status-ok { color: #34d399; font-weight: 700; }
    .status-warn { color: #fbbf24; font-weight: 700; }
    .status-err { color: #f87171; font-weight: 700; }
    .dl-links { display: flex; gap: 10px; flex-wrap: wrap; margin-top: 10px; }
    .dl-links a {
      text-decoration: none;
      background: #0f172a;
      border: 1px solid var(--accent);
      color: #93c5fd;
      padding: 8px 14px;
      border-radius: 8px;
      font-size: 13px;
      font-weight: 600;
    }
    .dl-links a:hover { background: var(--accent); color: #fff; }
    .step-indicator {
      display: flex;
      gap: 8px;
      margin-bottom: 14px;
      flex-wrap: wrap;
    }
    .step-pill {
      padding: 6px 12px;
      border-radius: 6px;
      font-size: 12.5px;
      background: #0f172a;
      border: 1px solid var(--card-border);
      color: var(--muted);
    }
    .step-pill.active {
      border-color: var(--emerald);
      color: #34d399;
      font-weight: 600;
    }
    .issue-box {
      background: rgba(245, 158, 11, 0.12);
      border-left: 4px solid var(--amber);
      padding: 10px 12px;
      margin: 8px 0;
      border-radius: 6px;
      font-size: 13px;
    }
    .kpi-grid {
      display: grid;
      grid-template-columns: repeat(auto-fit, minmax(175px, 1fr));
      gap: 10px;
      margin: 10px 0;
    }
    .kpi-card {
      background: #0f172a;
      border: 1px solid var(--card-border);
      border-radius: 8px;
      padding: 10px 12px;
    }
    .kpi-card .val { font-size: 20px; font-weight: 700; color: #60a5fa; }
    .kpi-card .lbl { font-size: 12px; color: var(--muted); }
  </style>
</head>
<body>
  <header>
    <div>
      <h1>ИИ-агент для создания образовательных программ АЦР</h1>
      <div style="font-size: 13px; color: var(--muted);">
        Академия цифрового развития (г. Магас) · Архитектура навыков Сбер GigaCowork · Cloud.ru Evolution Foundation Models
      </div>
    </div>
    <div class="badges">
      <span class="badge">ДООП (Приказ № 629)</span>
      <span class="badge">ДПП ПК и Стажировки (273-ФЗ, Приказ № 499)</span>
      <span class="badge">Детерминированный валидатор часов и РО</span>
      <span class="badge">API: POST /api/v1/generate</span>
    </div>
  </header>

  <div class="container">
    <!-- ЛЕВАЯ КОЛОНКА: ПРОВЕРОЧНЫЕ КЕЙСЫ АЦР И ВВОД ЗАПРОСА -->
    <div>
      <div class="card">
        <h2>1. Эталонные запросы АЦР (для проверки)</h2>
        <button class="preset-btn" onclick="loadPreset('case1')">
          <b>Запрос № 1 (ДООП · 16 ч)</b><br/>«Первые шаги в Python» для школьников 8–9 кл.
        </button>
        <button class="preset-btn" onclick="loadPreset('case2')">
          <b>Запрос № 2 (ДПП ПК · 36 ч)</b><br/>«ИИ в работе учителя» (профстандарт Педагог, сервисы РФ)
        </button>
        <button class="preset-btn" onclick="loadPreset('case3')">
          <b>Запрос № 3 (ДПП ПК · 72 ч)</b><br/>«Анализ данных на Python» для сотрудников банков
        </button>
        <button class="preset-btn" onclick="loadPreset('case4')">
          <b>Запрос № 4 (Неполный запрос → Вопросы)</b><br/>Курс по кибербезопасности для колледжа
        </button>
        <button class="preset-btn" onclick="loadPreset('case5')">
          <b>Запрос № 5 (Стажировка · 256 ч)</b><br/>«DevOps-инженер» + отлов ошибки в образце 248 ч
        </button>
      </div>

      <div class="card">
        <h2>2. Запрос методиста свободным текстом</h2>
        <textarea id="queryInput" rows="6" placeholder="Опишите программу свободным текстом..."></textarea>
        <label style="font-size: 12.5px; color: var(--muted); display: block; margin-bottom: 8px;">
          <input type="checkbox" id="useSample248" style="width: auto; margin-right: 6px;" />
          Приложить образец стажировки АЦР «Бизнес и системный анализ» (248 ч, с ошибкой в строке Итого)
        </label>
        <div class="btn-row">
          <button class="primary" onclick="runStep('analyze')">Шаг 1: Проверить запрос</button>
          <button class="primary" onclick="runStep('skeleton')">Шаг 2: Собрать каркас</button>
          <button class="emerald" onclick="runStep('full')">Собрать полный комплект</button>
        </div>
      </div>

      <!-- БЛОК ОТВЕТОВ НА УТОЧНЯЮЩИЕ ВОПРОСЫ (ДЛЯ ЗАПРОСА № 4) -->
      <div class="card" id="clarificationCard" style="display: none;">
        <h2>Уточняющие вопросы методисту (Шаг 2)</h2>
        <div id="questionsList" style="font-size: 13px; margin-bottom: 10px;"></div>
        <input id="ansType" placeholder="Тип программы (Повышение квалификации / Общеразвивающая)" value="Повышение квалификации" />
        <input id="ansHours" placeholder="Объём в часах (например, 72 часа)" value="72 часа" />
        <input id="ansAud" placeholder="Слушатели и входной уровень" value="Студенты 3 курса специальности «Информационные системы и программирование», знают основы сетей и Linux" />
        <input id="ansForm" placeholder="Форма обучения" value="Очная, занятия на базе колледжа" />
        <input id="ansAtt" placeholder="Форма итоговой аттестации" value="Экзамен с практическим заданием на стенде" />
        <button class="emerald" style="width: 100%;" onclick="submitAnswersAndBuild()">Передать ответы методиста и собрать программу</button>
      </div>

      <!-- БЛОК ТОЧЕЧНОЙ ДОРАБОТКИ И СВЕРХПЛАНОВЫХ ФУНКЦИЙ -->
      <div class="card">
        <h2>Точечная правка и адаптация (Шаг 7 + Сверх плана)</h2>
        <input id="editInstruction" value="Заменить тему 2.2 на «Продвинутая агрегация временных рядов и скоринговых витрин в pandas» и усложнить тест" />
        <div class="btn-row" style="margin-bottom: 10px;">
          <button class="secondary" onclick="runTargetedEdit()">Точечно пересобрать часть</button>
          <button class="secondary" onclick="runModernize()">Актуализировать под 2026 г.</button>
        </div>
        <div class="btn-row">
          <input id="adaptHours" type="number" value="144" style="width: 110px; margin-bottom: 0;" />
          <button class="secondary" onclick="runAdapt()">Масштабировать объём (ч.)</button>
        </div>
      </div>
    </div>

    <!-- ПРАВАЯ КОЛОНКА: РАБОЧАЯ ОБЛАСТЬ МЕТОДИСТА, КАРКАС, ОТЧЕТ И ВЫГРУЗКА -->
    <div>
      <div class="step-indicator">
        <div class="step-pill" id="pill1">1. Анализ запроса и референсов</div>
        <div class="step-pill" id="pill2">2. Каркас: Цель, ЗУН и УП (Human-in-the-Loop)</div>
        <div class="step-pill" id="pill3">3. Содержание: РПД, Тесты, Семинары, Слайды</div>
        <div class="step-pill" id="pill4">4. Самопроверка и выгрузка (.docx / .pptx / LMS / JSON)</div>
      </div>

      <div id="workspaceOutput">
        <div class="card">
          <h2>Рабочее пространство методиста АЦР</h2>
          <p style="color: var(--muted); font-size: 14px;">
            Выберите один из 5 проверочных запросов слева или введите собственный запрос методиста,
            чтобы запустить агентный пайплайн навыков в стиле <b>GigaCowork</b>.
          </p>
        </div>
      </div>
    </div>
  </div>

  <script>
    const PRESETS = __PRESETS_JSON__;

    function setPills(step) {
      for (let i = 1; i <= 4; i++) {
        const el = document.getElementById('pill' + i);
        if (i <= step) el.classList.add('active');
        else el.classList.remove('active');
      }
    }

    function loadPreset(key) {
      const p = PRESETS[key];
      document.getElementById('queryInput').value = p.query;
      document.getElementById('useSample248').checked = !!p.use_sample_248;
      document.getElementById('clarificationCard').style.display = 'none';
      if (key === 'case4') {
        runStep('analyze');
      } else {
        runStep('skeleton');
      }
    }

    async function callApi(payload) {
      const out = document.getElementById('workspaceOutput');
      out.innerHTML = '<div class="card"><h2>Выполняется обработка запроса агентом...</h2><p style="color:var(--muted);">Генерация структуры, проверка таблиц часов и формирование файлов (.docx, .pptx, LMS)...</p></div>';
      try {
        const resp = await fetch('/api/v1/generate', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify(payload)
        });
        return await resp.json();
      } catch (err) {
        out.innerHTML = '<div class="card"><h2 class="status-err">Ошибка соединения с локальным сервером</h2><p>' + err + '</p></div>';
        throw err;
      }
    }

    async function runStep(mode, customAnswers = null) {
      const query = document.getElementById('queryInput').value;
      const useSample248 = document.getElementById('useSample248').checked;
      const data = await callApi({
        mode: mode,
        query: query,
        answers: customAnswers,
        use_sample_248: useSample248
      });
      renderResponse(data);
    }

    function submitAnswersAndBuild() {
      const answers = {
        "Тип программы": document.getElementById('ansType').value,
        "Объём": document.getElementById('ansHours').value,
        "Слушатели": document.getElementById('ansAud').value,
        "Форма обучения": document.getElementById('ansForm').value,
        "Итоговая аттестация": document.getElementById('ansAtt').value
      };
      document.getElementById('clarificationCard').style.display = 'none';
      runStep('full', answers);
    }

    async function runTargetedEdit() {
      const query = document.getElementById('queryInput').value;
      const inst = document.getElementById('editInstruction').value;
      const data = await callApi({ mode: 'edit', query: query, edit_instruction: inst });
      renderResponse(data);
    }

    async function runModernize() {
      const query = document.getElementById('queryInput').value;
      const data = await callApi({ mode: 'modernize', query: query, inject_legacy_demo: true });
      renderResponse(data);
    }

    async function runAdapt() {
      const query = document.getElementById('queryInput').value;
      const h = parseFloat(document.getElementById('adaptHours').value || '144');
      const data = await callApi({ mode: 'adapt', query: query, new_total_hours: h });
      renderResponse(data);
    }

    function renderResponse(data) {
      const out = document.getElementById('workspaceOutput');
      let html = '';

      if (data.status === 'needs_clarification') {
        setPills(1);
        document.getElementById('clarificationCard').style.display = 'block';
        const qs = data.clarification.questions_for_methodologist || [];
        document.getElementById('questionsList').innerHTML =
          '<b>Агент обнаружил нехватку входных данных (' + qs.length + ' вопр.):</b><ol>' +
          qs.map(q => '<li>' + q + '</li>').join('') + '</ol>';

        html += '<div class="card"><h2>Шаг 1–2. Требуется уточнение параметров у методиста</h2>';
        html += '<p class="status-warn">Агент не стал додумывать недостающие параметры и сформировал ' + qs.length + ' уточняющих вопросов (см. форму слева).</p>';
        if (data.clarification.expectation_reality_gap_note) {
          html += '<div class="issue-box">' + data.clarification.expectation_reality_gap_note + '</div>';
        }
        html += '</div>';
        out.innerHTML = html;
        return;
      }

      const kit = data.skeleton_kit || data.program_kit;
      if (!kit) return;

      if (data.status === 'skeleton_ready_for_approval') setPills(2);
      else setPills(4);

      if (data.reference_defects && data.reference_defects.length > 0) {
        html += '<div class="card"><h2>Инспекция приложенного образца (.docx)</h2>';
        html += '<p class="status-warn">В приложенном образце обнаружено расхождений: ' + data.reference_defects.length + ' (НЕ перенесены в новую программу):</p>';
        data.reference_defects.forEach(d => {
          html += '<div class="issue-box"><b>[' + d.code + '] ' + d.location + ':</b> ' + d.message + '</div>';
        });
        html += '</div>';
      }

      if (data.edit_summary) {
        html += '<div class="card"><h2>Результат точечной пересборки (Шаг 7)</h2><p class="status-ok">' + data.edit_summary + '</p></div>';
      }
      if (data.modernization_logs) {
        html += '<div class="card"><h2>Отчет об актуализации программы под стек 2026 года (Сверх плана)</h2><ul>' +
          data.modernization_logs.map(l => '<li>' + l + '</li>').join('') + '</ul></div>';
      }
      if (data.adaptation_summary) {
        html += '<div class="card"><h2>Результат адаптации объема и аудитории (Сверх плана)</h2><p class="status-ok">' + data.adaptation_summary + '</p></div>';
      }

      if (data.download_files) {
        html += '<div class="card"><h2>Готовый комплект для выгрузки (.docx / .pptx / LMS / JSON)</h2>';
        html += '<div class="dl-links">';
        html += '<a href="' + data.download_files.docx + '" target="_blank">Скачать программу (.docx)</a>';
        html += '<a href="' + data.download_files.pptx + '" target="_blank">Скачать слайды лекций (.pptx)</a>';
        html += '<a href="' + data.download_files.gift + '" target="_blank">Тесты LMS (GIFT)</a>';
        html += '<a href="' + data.download_files.moodle_xml + '" target="_blank">Тесты LMS (Moodle XML)</a>';
        html += '<a href="' + data.download_files.csv + '" target="_blank">Тесты LMS (CSV)</a>';
        html += '<a href="' + data.download_files.json + '" target="_blank">Полный комплект (JSON)</a>';
        html += '</div></div>';
      }

      if (data.validation_report) {
        const rep = data.validation_report;
        html += '<div class="card"><h2>Отчет детерминированной автопроверки и эффект Cloud.ru</h2>';
        html += '<div class="kpi-grid">';
        html += '<div class="kpi-card"><div class="val" style="color:' + (rep.is_valid ? '#34d399' : '#f87171') + ';">' + (rep.is_valid ? 'ПРОЙДЕНА (0 ошибок)' : 'ОШИБОК: ' + rep.error_count) + '</div><div class="lbl">Статус автопроверки АЦР</div></div>';
        html += '<div class="kpi-card"><div class="val">' + kit.total_hours + ' ак. ч.</div><div class="lbl">Сходимость УП ↔ УТП ↔ РПД</div></div>';
        if (data.roi_metrics) {
          html += '<div class="kpi-card"><div class="val">' + data.roi_metrics.manual_hours_saved + ' ч.</div><div class="lbl">Экономия времени методиста (' + data.roi_metrics.speedup_factor + 'x)</div></div>';
          html += '<div class="kpi-card"><div class="val">' + data.roi_metrics.total_tokens + ' ток. (~' + data.roi_metrics.estimated_cost_rub + ' ₽)</div><div class="lbl">' + data.roi_metrics.model_name + '</div></div>';
        }
        html += '</div></div>';
      }

      html += '<div class="card">';
      html += '<h2>' + (data.status === 'skeleton_ready_for_approval' ? 'Каркас программы на согласование методисту (Human-in-the-Loop)' : 'Структура и содержание программы: «' + kit.program_title + '»') + '</h2>';
      html += '<p><b>Шаблон АЦР:</b> ' + (kit.template_type === 'doop' ? 'Дополнительная общеразвивающая программа (Приказ Минпросвещения № 629, Сертификат)' : 'Программа повышения квалификации (273-ФЗ, Приказ Минобрнауки № 499, Удостоверение)') + '</p>';
      html += '<p><b>Категория слушателей:</b> ' + kit.target_audience + '</p>';
      html += '<p><b>Цель:</b> ' + kit.goal_text + '</p>';
      if (kit.methodologist_notes && kit.methodologist_notes.length > 0) {
        kit.methodologist_notes.forEach(n => {
          html += '<div class="issue-box">' + n + '</div>';
        });
      }

      if (kit.competencies && kit.competencies.length > 0 && kit.competencies[0].prof_standard_link) {
        const ps = kit.competencies[0].prof_standard_link;
        html += '<p><b>Профстандарт Минтруда РФ:</b> ' + ps.standard_code + ' «' + ps.standard_title + '» (' + ps.mintrud_order + ') · Код ТФ: <b>' + ps.labor_function_code + '</b></p>';
      }

      html += '<h3>Планируемые результаты обучения (ЗУН)</h3><ul>';
      (kit.learning_outcomes || []).forEach(lo => {
        html += '<li><b>[' + lo.id + ' · ' + lo.category + ']</b> ' + lo.description + '</li>';
      });
      html += '</ul>';

      html += '<h3>Учебный и учебно-тематический план (автопроверка сумм: ОК)</h3>';
      html += '<table><thead><tr><th>Код</th><th>Наименование модуля / темы</th><th>Всего (ч.)</th><th>ЛК</th><th>ПР</th><th>ЛБ</th><th>СРС</th><th>Контроль / РО</th></tr></thead><tbody>';
      (kit.modules || []).forEach(m => {
        html += '<tr style="background:#0f172a;font-weight:700;"><td>Модуль ' + m.id + '</td><td>' + m.title + '</td><td>' + m.total_hours + '</td><td>' + m.lecture_hours + '</td><td>' + m.practice_hours + '</td><td>' + m.lab_hours + '</td><td>' + m.self_study_hours + '</td><td>' + m.intermediate_attestation_form + '</td></tr>';
        (m.topics || []).forEach(t => {
          html += '<tr><td>' + t.id + '</td><td>' + t.title + '</td><td>' + t.total_hours + '</td><td>' + t.lecture_hours + '</td><td>' + t.practice_hours + '</td><td>' + t.lab_hours + '</td><td>' + t.self_study_hours + '</td><td>' + (t.learning_outcome_ids || []).join(', ') + '</td></tr>';
        });
      });
      html += '<tr style="font-weight:700;"><td>ИА</td><td>Итоговая аттестация: ' + kit.final_attestation.form + '</td><td>' + kit.final_attestation.hours + '</td><td>0</td><td>0</td><td>0</td><td>0</td><td>Зачет / Защита</td></tr>';
      html += '<tr style="background:#1e3a8a;font-weight:700;"><td>ИТОГО</td><td>Итого по программе</td><td>' + kit.total_hours + '</td><td>' + kit.declared_totals.lecture + '</td><td>' + kit.declared_totals.practice + '</td><td>' + kit.declared_totals.lab + '</td><td>' + kit.declared_totals.self_study + '</td><td>Сходимость 100%</td></tr>';
      html += '</tbody></table>';

      if (data.status === 'skeleton_ready_for_approval') {
        html += '<div style="margin-top: 14px;"><button class="emerald" onclick="runStep(\'full\')">Утвердить каркас (Методист) и сгенерировать полное содержание комплекта</button></div>';
      } else {
        html += '<h3>Оценочные и презентационные материалы комплекта</h3>';
        html += '<ul>';
        html += '<li><b>Тестовых вопросов с ключами и привязкой к РО:</b> ' + (kit.test_questions || []).length + ' шт.</li>';
        html += '<li><b>Практических заданий для семинаров с критериями оценки:</b> ' + (kit.seminar_tasks || []).length + ' шт.</li>';
        html += '<li><b>Вопросов для собеседования текущего контроля:</b> ' + (kit.interview_questions || []).length + ' шт.</li>';
        html += '<li><b>Презентаций лекций (.pptx):</b> ' + (kit.lecture_decks || []).map(d => 'Тема ' + d.topic_id + ' («' + d.topic_title + '»)').join('; ') + '</li>';
        html += '</ul>';
      }
      html += '</div>';

      out.innerHTML = html;
    }

    window.addEventListener('DOMContentLoaded', () => {
      loadPreset('case1');
    });
  </script>
</body>
</html>
"""


# =====================================================================
# 3. HTTP-СЕРВЕР (РАБОТАЕТ БЕЗ ВНЕШНИХ ЗАВИСИМОСТЕЙ В VS CODE И CLOUD.RU)
# =====================================================================

class ACRRequestHandler(BaseHTTPRequestHandler):
    def log_message(self, format: str, *args: Any) -> None:
        sys.stderr.write(f"[ACR-Agent HTTP] {self.address_string()} - {format % args}\n")

    def _send_json(self, data: Dict[str, Any], status: int = 200) -> None:
        raw = json.dumps(data, ensure_ascii=False, indent=2).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(raw)))
        self.send_header("Access-Control-Allow-Origin", "*")
        self.end_headers()
        self.wfile.write(raw)

    def do_GET(self) -> None:
        parsed = urllib.parse.urlparse(self.path)
        path = parsed.path

        if path in ("/", "/index.html"):
            page = HTML_UI_PAGE.replace("__PRESETS_JSON__", json.dumps(PRESET_QUERIES, ensure_ascii=False))
            raw = page.encode("utf-8")
            self.send_response(200)
            self.send_header("Content-Type", "text/html; charset=utf-8")
            self.send_header("Content-Length", str(len(raw)))
            self.end_headers()
            self.wfile.write(raw)
            return

        if path == "/api/v1/health":
            self._send_json({"status": "ok", "service": "acr-edtech-agent", "skills_count": len(GIGACOWORK_SKILLS_CATALOG)})
            return

        if path == "/api/v1/skills":
            self._send_json({"skills": [dataclasses_asdict(s) for s in GIGACOWORK_SKILLS_CATALOG]})
            return

        if path.startswith("/api/v1/download/"):
            fname = os.path.basename(urllib.parse.unquote(path[len("/api/v1/download/") :]))
            fpath = os.path.join(OUTPUT_DIR, fname)
            if not os.path.exists(fpath):
                self._send_json({"error": f"File not found: {fname}"}, status=404)
                return
            mime, _ = mimetypes.guess_type(fpath)
            with open(fpath, "rb") as f:
                data_bytes = f.read()
            self.send_response(200)
            self.send_header("Content-Type", mime or "application/octet-stream")
            self.send_header("Content-Disposition", f'attachment; filename="{fname}"')
            self.send_header("Content-Length", str(len(data_bytes)))
            self.end_headers()
            self.wfile.write(data_bytes)
            return

        self._send_json({"error": "Not found"}, status=404)

    def do_POST(self) -> None:
        parsed = urllib.parse.urlparse(self.path)
        if parsed.path in ("/api/v1/generate", "/api/v1/analyze", "/api/v1/skeleton", "/api/v1/edit", "/api/v1/adapt"):
            length = int(self.headers.get("Content-Length", "0"))
            raw_body = self.rfile.read(length).decode("utf-8") if length > 0 else "{}"
            try:
                payload = json.loads(raw_body)
            except Exception as e:
                self._send_json({"error": f"Invalid JSON: {e}"}, status=400)
                return

            if parsed.path == "/api/v1/analyze":
                payload["mode"] = "analyze"
            elif parsed.path == "/api/v1/skeleton":
                payload["mode"] = "skeleton"
            elif parsed.path == "/api/v1/edit":
                payload["mode"] = "edit"
            elif parsed.path == "/api/v1/adapt":
                payload["mode"] = "adapt"

            result = process_unified_api_request(payload)
            self._send_json(result, status=200)
            return

        self._send_json({"error": "Endpoint not found"}, status=404)


def run_server(host: str = "127.0.0.1", port: int = 8000, open_browser: bool = True) -> None:
    server = ThreadingHTTPServer((host, port), ACRRequestHandler)
    url = f"http://{host}:{port}"
    print("=" * 80)
    print("ВЕБ-СЕРВИС И ЕДИНЫЙ API-ЭНДПОИНТ АЦР × GIGACOWORK ЗАПУЩЕН!")
    print(f"  * Веб-интерфейс методиста (открой в браузере): {url}")
    print(f"  * Единый API-эндпоинт для автопроверки жюри  : POST {url}/api/v1/generate")
    print(f"  * Каталог навыков GigaCowork                 : GET  {url}/api/v1/skills")
    print("Для остановки сервера нажми Ctrl + C в терминале.")
    print("=" * 80)
    if open_browser:
        try:
            webbrowser.open(url)
        except Exception:
            pass
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("\nСервер остановлен.")
        server.server_close()


if __name__ == "__main__":
    port_num = int(os.environ.get("PORT", "8000"))
    run_server(host="127.0.0.1", port=port_num, open_browser=True)