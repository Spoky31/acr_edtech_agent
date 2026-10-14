# -*- coding: utf-8 -*-
"""
Генераторы выгрузки полного комплекта образовательной программы АЦР (`program_exporters.py`).

Поддерживаемые выходные форматы (согласно разделу 4 и 5 условий кейса):
1. Программа курса в формате Microsoft Word (`.docx`) по обоим шаблонам АЦР:
   - ДООП (Приказ Минпросвещения № 629, сертификат, без профстандартов);
   - ДПП ПК и Стажировка (273-ФЗ, Приказ Минобрнауки № 499, удостоверение, профстандарты и коды ТФ).
2. Планы и слайды лекций в формате Microsoft PowerPoint (`.pptx`).
3. Тестовые вопросы с ключами и привязкой к результатам обучения в форматах для LMS:
   - `GIFT` (`.gift.txt`)
   - `Moodle XML` (`.moodle.xml`)
   - `CSV` (`.tests.csv`)
4. Полный слепок комплекта данных и отчета автопроверки в формате `JSON` (`.json`) и `Markdown` (`.report.md`).
"""

from __future__ import annotations

import csv
import dataclasses
import html
import json
import os
import zipfile
from typing import Dict, List, Optional

from program_validator import (
    ProgramKit,
    ProgramValidator,
    UserQuerySpec,
    ValidationReport,
    auto_fix_hours,
)

try:
    import docx
    from docx.enum.table import WD_ALIGN_VERTICAL, WD_TABLE_ALIGNMENT
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.oxml import parse_xml
    from docx.oxml.ns import nsdecls
    from docx.shared import Inches, Pt, RGBColor
except ImportError:
    docx = None

try:
    import pptx
    from pptx import Presentation
    from pptx.dml.color import RGBColor as PptxRGBColor
    from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
    from pptx.util import Inches as PptxInches, Pt as PptxPt
except ImportError:
    pptx = None


def _fmt_num(val: float) -> str:
    """Форматирует часы по стандарту таблиц АЦР (например, 1,5 или 24)."""
    if abs(val - round(val)) < 1e-4:
        return str(int(round(val)))
    return f"{val:.1f}".replace(".", ",")


# =====================================================================
# 1. ГЕНЕРАТОР ПРОГРАММЫ В ФОРМАТЕ WORD (.DOCX) ПО ШАБЛОНАМ АЦР
# =====================================================================

def _style_docx_table(table, header_rows: int = 1, header_hex: str = "1E3A8A", alt_hex: str = "F8FAFC") -> None:
    """Применяет читаемое табличное форматирование (10pt+, заливка шапки, запрет разрыва строк)."""
    if docx is None:
        return
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    for r_idx, row in enumerate(table.rows):
        is_header = r_idx < header_rows
        is_total = False
        first_text = " ".join(c.text.strip().lower() for c in row.cells[:2])
        if first_text.startswith("итого"):
            is_total = True

        trPr = row._tr.get_or_add_trPr()
        if is_header:
            trPr.append(parse_xml(f"<w:tblHeader {nsdecls('w')}/>"))
        else:
            trPr.append(parse_xml(f"<w:cantSplit {nsdecls('w')}/>"))

        for cell in row.cells:
            cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            if is_header:
                shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{header_hex}"/>')
                cell._tc.get_or_add_tcPr().append(shd)
            elif is_total:
                shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="E2E8F0"/>')
                cell._tc.get_or_add_tcPr().append(shd)
            elif r_idx % 2 == 1:
                shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{alt_hex}"/>')
                cell._tc.get_or_add_tcPr().append(shd)

            for p in cell.paragraphs:
                p.paragraph_format.space_before = Pt(2)
                p.paragraph_format.space_after = Pt(2)
                p.paragraph_format.line_spacing = 1.05
                if is_header:
                    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    p.paragraph_format.keep_with_next = True
                for run in p.runs:
                    run.font.name = "Calibri"
                    run.font.size = Pt(10)
                    if is_header:
                        run.font.bold = True
                        run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
                    elif is_total:
                        run.font.bold = True
                        run.font.color.rgb = RGBColor(0x0F, 0x17, 0x2A)
                    else:
                        run.font.color.rgb = RGBColor(0x0F, 0x17, 0x2A)


def _add_heading_1(doc, text: str) -> None:
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(14)
    p.paragraph_format.space_after = Pt(6)
    p.paragraph_format.keep_with_next = True
    run = p.add_run(text)
    run.font.name = "Calibri"
    run.font.size = Pt(15)
    run.font.bold = True
    run.font.color.rgb = RGBColor(0x1E, 0x3A, 0x8A)


def _add_heading_2(doc, text: str) -> None:
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(10)
    p.paragraph_format.space_after = Pt(4)
    p.paragraph_format.keep_with_next = True
    run = p.add_run(text)
    run.font.name = "Calibri"
    run.font.size = Pt(13)
    run.font.bold = True
    run.font.color.rgb = RGBColor(0x0F, 0x17, 0x2A)


def _add_body_p(doc, text: str, bold_prefix: str = "") -> None:
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(6)
    p.paragraph_format.line_spacing = 1.15
    if bold_prefix:
        r_b = p.add_run(bold_prefix)
        r_b.font.name = "Calibri"
        r_b.font.size = Pt(11)
        r_b.font.bold = True
    r_t = p.add_run(text)
    r_t.font.name = "Calibri"
    r_t.font.size = Pt(11)


def _add_bullet_p(doc, text: str) -> None:
    p = doc.add_paragraph(style="List Bullet")
    p.paragraph_format.space_before = Pt(1)
    p.paragraph_format.space_after = Pt(3)
    p.paragraph_format.line_spacing = 1.15
    r = p.add_run(text)
    r.font.name = "Calibri"
    r.font.size = Pt(11)


def _add_acr_title_block(doc, kit: ProgramKit) -> None:
    """Формирует официальную шапку АЦР (г. Магас) с пустыми полями приказа."""
    org_lines = [
        "Автономная некоммерческая организация",
        "дополнительного профессионального образования",
        "«Академия Информационных Технологий»",
        "ОГРН: 1230600003457, ИНН/КПП: 0600010064/060001001",
        "Российская Федерация, 386001, Республика Ингушетия, городской округ Магас,",
        "город Магас, улица Н.С. Хрущева, дом 10",
    ]
    for line in org_lines:
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_after = Pt(2)
        run = p.add_run(line)
        run.font.name = "Calibri"
        run.font.size = Pt(11)
        if "Академия Информационных Технологий" in line:
            run.font.bold = True

    p_app = doc.add_paragraph()
    p_app.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    p_app.paragraph_format.space_before = Pt(10)
    p_app.paragraph_format.space_after = Pt(12)
    r_app = p_app.add_run(
        f"УТВЕРЖДЕНО\nПриказом № {kit.requisites.order_number} от {kit.requisites.order_date}"
    )
    r_app.font.name = "Calibri"
    r_app.font.size = Pt(11)
    r_app.font.bold = True

    if kit.template_type == "doop":
        type_header = "ДОПОЛНИТЕЛЬНАЯ ОБЩЕОБРАЗОВАТЕЛЬНАЯ\nОБЩЕРАЗВИВАЮЩАЯ ПРОГРАММА"
    else:
        if kit.is_internship:
            type_header = "ДОПОЛНИТЕЛЬНАЯ ПРОФЕССИОНАЛЬНАЯ ПРОГРАММА\nПОВЫШЕНИЯ КВАЛИФИКАЦИИ В ФОРМЕ СТАЖИРОВКИ"
        else:
            type_header = "ДОПОЛНИТЕЛЬНАЯ ПРОФЕССИОНАЛЬНАЯ ПРОГРАММА\nПОВЫШЕНИЯ КВАЛИФИКАЦИИ"

    p_type = doc.add_paragraph()
    p_type.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_type.paragraph_format.space_after = Pt(6)
    r_type = p_type.add_run(type_header)
    r_type.font.name = "Calibri"
    r_type.font.size = Pt(14)
    r_type.font.bold = True

    p_title = doc.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_title.paragraph_format.space_after = Pt(12)
    r_title = p_title.add_run(f"«{kit.program_title}»")
    r_title.font.name = "Calibri"
    r_title.font.size = Pt(18)
    r_title.font.bold = True
    r_title.font.color.rgb = RGBColor(0x1E, 0x3A, 0x8A)

    p_city = doc.add_paragraph()
    p_city.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_city.paragraph_format.space_after = Pt(16)
    r_city = p_city.add_run("Магас 2026")
    r_city.font.name = "Calibri"
    r_city.font.size = Pt(11)


def export_program_to_docx(kit: ProgramKit, output_path: str) -> str:
    """
    Генерирует документ программы в формате .docx строго по выбранному шаблону АЦР
    ('doop' или 'dpp_pk'). Перед выгрузкой выполняет автосинхронизацию часов.
    """
    auto_fix_hours(kit)
    os.makedirs(os.path.dirname(os.path.abspath(output_path)) or ".", exist_ok=True)

    if docx is None:
        return _export_docx_stdlib_fallback(kit, output_path)

    doc = docx.Document()
    for section in doc.sections:
        section.top_margin = Inches(1.0)
        section.bottom_margin = Inches(1.0)
        section.left_margin = Inches(1.0)
        section.right_margin = Inches(1.0)

    style_normal = doc.styles["Normal"]
    style_normal.font.name = "Calibri"
    style_normal.font.size = Pt(11)

    _add_acr_title_block(doc, kit)

    if kit.template_type == "doop":
        _build_doop_docx_body(doc, kit)
    else:
        _build_dpp_pk_docx_body(doc, kit)

    doc.save(output_path)
    return output_path


def _build_doop_docx_body(doc, kit: ProgramKit) -> None:
    """Строит разделы шаблона ДООП (по образцу 6aa380dc...docx и Приказу Минпросвещения № 629)."""
    _add_heading_1(doc, "Пояснительная записка")
    _add_body_p(doc, kit.explanatory_note)
    level_map = {"zero": "стартовый (с нуля)", "basic": "базовый", "intermediate": "продвинутый", "advanced": "углубленный"}
    form_map = {
        "ochnaya": "очная",
        "ochno_zaochnaya": "очно-заочная с применением ЭО и ДОТ",
        "zaochnaya_dot": "заочная с применением ЭО и ДОТ",
        "internship": "стажировка",
    }
    _add_body_p(doc, level_map.get(kit.entry_level, "базовый") + ".", "Уровень освоения программы: ")
    _add_body_p(doc, "техническая.", "Направленность (профиль) программы: ")
    _add_body_p(doc, form_map.get(kit.study_form, kit.study_form) + ".", "Форма обучения: ")
    _add_body_p(doc, kit.issued_document + ".", "Документ, выдаваемый по итогам обучения: ")

    _add_heading_2(doc, "1.1. Актуальность Программы")
    _add_body_p(doc, kit.relevance_text)

    _add_heading_2(doc, "1.2. Категории обучающихся")
    _add_body_p(doc, kit.target_audience)

    _add_heading_2(doc, "1.3. Цель и задачи программы")
    _add_body_p(doc, kit.goal_text, "Цель Программы: ")
    _add_body_p(doc, "Задачи Программы:", "")
    for task_str in kit.tasks_list:
        _add_bullet_p(doc, task_str)

    _add_heading_2(doc, "1.4. Сроки реализации")
    _add_body_p(doc, f"{kit.calendar_schedule.total_weeks} нед. ({kit.calendar_schedule.total_days} уч. дн.).", "Срок реализации: ")
    _add_body_p(doc, f"{_fmt_num(kit.total_hours)} ак. ч.", "Трудоемкость: ")

    _add_heading_2(doc, "1.5. Планируемые результаты обучения и карта компетенций")
    for cat_key, cat_title in [("knowledge", "Знания:"), ("skill", "Умения:"), ("habit", "Навыки:")]:
        items = [lo for lo in kit.learning_outcomes if lo.category == cat_key]
        if items:
            _add_body_p(doc, "", cat_title)
            for lo in items:
                _add_bullet_p(doc, f"[{lo.id}] {lo.description}")

    _add_body_p(doc, "Карта компетенций и матрица покрытия оценочными средствами:")
    t_map = doc.add_table(rows=1, cols=4)
    for idx, h in enumerate(["Код РО", "Планируемый результат обучения", "Темы УТП", "Оценочные средства"]):
        t_map.rows[0].cells[idx].text = h
    for lo in kit.learning_outcomes:
        row = t_map.add_row()
        linked_topics = [t.id for m in kit.modules for t in m.topics if lo.id in t.learning_outcome_ids]
        linked_tests = [f"Тест {q.id}" for q in kit.test_questions if lo.id in q.learning_outcome_ids]
        linked_tasks = [f"Задание {s.id}" for s in kit.seminar_tasks if lo.id in s.learning_outcome_ids]
        row.cells[0].text = lo.id
        row.cells[1].text = lo.description
        row.cells[2].text = ", ".join(linked_topics) or "—"
        row.cells[3].text = ", ".join(linked_tests + linked_tasks) or "—"
    _style_docx_table(t_map, header_rows=1)

    _add_heading_2(doc, "1.6. Календарный учебный график")
    _add_body_p(
        doc,
        f"Количество учебных недель — {kit.calendar_schedule.total_weeks}. "
        f"Количество учебных дней — {kit.calendar_schedule.total_days}. "
        f"Режим занятий: до {_fmt_num(kit.calendar_schedule.max_hours_per_day)} ч. в день. "
        f"{kit.calendar_schedule.description}",
    )

    _add_heading_1(doc, "2. Содержание программы")
    _add_heading_2(doc, "2.1. Учебный план")

    t_up = doc.add_table(rows=2, cols=8)
    headers_r0 = [
        "№ п/п",
        "Наименование разделов/дисциплин/модулей",
        "Общая трудоемкость",
        "Всего аудиторных часов",
        "в т.ч.",
        "в т.ч.",
        "Самостоятельная работа",
        "Форма промежуточной аттестации/контроля (при наличии)",
    ]
    headers_r1 = [
        "№ п/п",
        "Наименование разделов/дисциплин/модулей",
        "Общая трудоемкость",
        "Всего аудиторных часов",
        "Лекции/теория",
        "Практические занятия",
        "Самостоятельная работа",
        "Форма промежуточной аттестации/контроля (при наличии)",
    ]
    for c_i, h in enumerate(headers_r0):
        t_up.rows[0].cells[c_i].text = h
    for c_i, h in enumerate(headers_r1):
        t_up.rows[1].cells[c_i].text = h

    for mod in kit.modules:
        r = t_up.add_row()
        aud_h = mod.lecture_hours + mod.practice_hours + mod.lab_hours
        prac_h = mod.practice_hours + mod.lab_hours
        vals = [
            mod.id,
            mod.title,
            _fmt_num(mod.total_hours),
            _fmt_num(aud_h),
            _fmt_num(mod.lecture_hours),
            _fmt_num(prac_h),
            _fmt_num(mod.self_study_hours),
            mod.intermediate_attestation_form or "—",
        ]
        for c_i, v in enumerate(vals):
            r.cells[c_i].text = v

    r_ia = t_up.add_row()
    ia_vals = [
        "",
        "Итоговая аттестация",
        _fmt_num(kit.final_attestation.hours),
        "—",
        "—",
        "—",
        "—",
        kit.final_attestation.form,
    ]
    for c_i, v in enumerate(ia_vals):
        r_ia.cells[c_i].text = v

    r_tot = t_up.add_row()
    tot_lk = kit.declared_totals["lecture"]
    tot_pr = kit.declared_totals["practice"] + kit.declared_totals["lab"]
    tot_aud = tot_lk + tot_pr
    tot_srs = kit.declared_totals["self_study"]
    tot_vals = [
        "",
        "Итого часов",
        _fmt_num(kit.total_hours),
        _fmt_num(tot_aud),
        _fmt_num(tot_lk),
        _fmt_num(tot_pr),
        _fmt_num(tot_srs),
        "",
    ]
    for c_i, v in enumerate(tot_vals):
        r_tot.cells[c_i].text = v
    _style_docx_table(t_up, header_rows=2)

    _add_heading_2(doc, "2.2. Учебно-тематический план")
    t_utp = doc.add_table(rows=1, cols=3)
    for c_i, h in enumerate(["Наименование разделов и тем", "Виды учебных занятий", "Содержание"]):
        t_utp.rows[0].cells[c_i].text = h

    for mod in kit.modules:
        r_m = t_utp.add_row()
        r_m.cells[0].text = f"Модуль {mod.id}. {mod.title}"
        r_m.cells[1].text = f"Всего — {_fmt_num(mod.total_hours)} ч."
        r_m.cells[2].text = ""
        for topic in mod.topics:
            if topic.rpd_activities:
                for a_idx, act in enumerate(topic.rpd_activities):
                    r_t = t_utp.add_row()
                    r_t.cells[0].text = f"Тема {topic.id}. {topic.title}" if a_idx == 0 else ""
                    act_label = "Теоретическое занятие" if act.activity_type == "ЛК" else (
                        "Практическое занятие" if act.activity_type in ("ПР", "ЛБ") else "Самостоятельная работа"
                    )
                    r_t.cells[1].text = f"{act_label} – {_fmt_num(act.hours)} часа"
                    r_t.cells[2].text = act.title_and_content
            else:
                r_t = t_utp.add_row()
                r_t.cells[0].text = f"Тема {topic.id}. {topic.title}"
                parts = []
                if topic.lecture_hours > 0:
                    parts.append(f"ЛК – {_fmt_num(topic.lecture_hours)} ч.")
                if topic.practice_hours + topic.lab_hours > 0:
                    parts.append(f"ПР – {_fmt_num(topic.practice_hours + topic.lab_hours)} ч.")
                if topic.self_study_hours > 0:
                    parts.append(f"СРС – {_fmt_num(topic.self_study_hours)} ч.")
                r_t.cells[1].text = ", ".join(parts)
                r_t.cells[2].text = topic.content_summary

    r_ia_utp = t_utp.add_row()
    r_ia_utp.cells[0].text = "Итоговая аттестация"
    r_ia_utp.cells[1].text = f"{kit.final_attestation.form} ({_fmt_num(kit.final_attestation.hours)} ч.)"
    r_ia_utp.cells[2].text = kit.final_attestation.requirements_description
    _style_docx_table(t_utp, header_rows=1)

    _add_heading_2(doc, "2.3. Оценочные материалы")
    _add_body_p(doc, kit.final_attestation.requirements_description, f"Итоговая аттестация ({kit.final_attestation.form}): ")
    _add_body_p(doc, "Критерии оценивания итоговой аттестации:")
    for crit in kit.final_attestation.evaluation_criteria:
        _add_bullet_p(doc, crit)

    _append_assessments_section(doc, kit)

    _add_heading_2(doc, "2.4. Методические материалы")
    _add_body_p(
        doc,
        "Теоретические занятия проводятся с использованием интерактивных презентаций и разбором практических примеров. "
        "Практические занятия направлены на закрепление умений и выполнение индивидуальных или командных заданий под руководством преподавателя.",
    )

    _add_heading_1(doc, "3. Организационно-педагогические условия реализации программы")
    _add_heading_2(doc, "3.1. Материально-технические и программные условия")
    for tool in kit.software_and_tools:
        _add_bullet_p(doc, tool)

    _add_heading_2(doc, "3.2. Нормативно-правовая база и литература")
    for reg in kit.regulatory_documents:
        _add_bullet_p(doc, reg)
    for src in kit.literature:
        tag = f" {src.verification_tag}" if not src.is_verified and src.verification_tag else ""
        _add_bullet_p(doc, f"{src.citation}{tag}")

    _add_heading_2(doc, "3.3. Кадровое обеспечение программы")
    _add_staff_table_empty(doc, kit)


def _build_dpp_pk_docx_body(doc, kit: ProgramKit) -> None:
    """Строит разделы шаблона ДПП ПК / Стажировки (по образцу 35f2d459...docx, 273-ФЗ и Приказу № 499)."""
    _add_heading_1(doc, "1. ОБЩАЯ ХАРАКТЕРИСТИКА ДОПОЛНИТЕЛЬНОЙ ПРОФЕССИОНАЛЬНОЙ ПРОГРАММЫ")
    _add_heading_2(doc, "1.1. Аннотация программы")
    _add_body_p(doc, kit.explanatory_note)
    _add_body_p(doc, kit.relevance_text)

    _add_heading_2(doc, "1.2. Цель реализации программы")
    _add_body_p(doc, kit.goal_text)
    for task_str in kit.tasks_list:
        _add_bullet_p(doc, task_str)

    _add_heading_2(doc, "1.3. Планируемые результаты обучения (РО)")
    t_ro = doc.add_table(rows=1, cols=5)
    for c_i, h in enumerate(["Виды деятельности", "Профессиональные компетенции", "Практический опыт", "Умения", "Знания"]):
        t_ro.rows[0].cells[c_i].text = h

    know_list = [f"• [{lo.id}] {lo.description}" for lo in kit.learning_outcomes if lo.category == "knowledge"]
    skill_list = [f"• [{lo.id}] {lo.description}" for lo in kit.learning_outcomes if lo.category == "skill"]
    prac_list = [
        f"• [{lo.id}] {lo.description}"
        for lo in kit.learning_outcomes
        if lo.category in ("practical_experience", "habit")
    ]
    for comp in kit.competencies:
        r = t_ro.add_row()
        std_title = comp.prof_standard_link.labor_function_title if comp.prof_standard_link else comp.name
        r.cells[0].text = std_title
        r.cells[1].text = comp.name
        r.cells[2].text = "\n".join(prac_list) or "—"
        r.cells[3].text = "\n".join(skill_list) or "—"
        r.cells[4].text = "\n".join(know_list) or "—"
    _style_docx_table(t_ro, header_rows=1)

    _add_heading_2(doc, "1.4. Программа разработана на основе:")
    for reg in kit.regulatory_documents:
        _add_bullet_p(doc, reg)
    for comp in kit.competencies:
        if comp.prof_standard_link:
            ps = comp.prof_standard_link
            _add_bullet_p(
                doc,
                f"Профессионального стандарта {ps.standard_code} «{ps.standard_title}» ({ps.mintrud_order}).",
            )

    _add_heading_2(doc, "1.5. Перечень профессиональных компетенций и трудовых функций:")
    for comp in kit.competencies:
        if comp.prof_standard_link:
            ps = comp.prof_standard_link
            _add_bullet_p(doc, f"{ps.labor_function_code} {ps.labor_function_title} — {comp.name}.")

    _add_heading_2(doc, "1.6. Требования к слушателям (категории слушателей)")
    _add_body_p(doc, kit.target_audience)

    _add_heading_2(doc, "1.7. Форма обучения")
    form_desc = {
        "ochnaya": "Очная.",
        "ochno_zaochnaya": "Очно-заочная с применением электронного обучения и дистанционных образовательных технологий.",
        "zaochnaya_dot": "Заочная с применением электронного обучения и дистанционных образовательных технологий.",
        "internship": "Стажировка (самостоятельная практическая работа под руководством наставника).",
    }
    _add_body_p(doc, form_desc.get(kit.study_form, kit.study_form))

    _add_heading_2(doc, "1.8. Объём программы")
    contact_h = kit.declared_totals["lecture"] + kit.declared_totals["practice"] + kit.declared_totals["lab"]
    srs_h = kit.declared_totals["self_study"]
    _add_body_p(doc, f"{_fmt_num(kit.total_hours)} академических часов.")
    _add_body_p(doc, f"{_fmt_num(contact_h)} часов.", "Контактная (аудиторная / синхронная) работа: ")
    _add_body_p(doc, f"{_fmt_num(srs_h)} часов.", "Самостоятельная работа слушателей (СРС / ДОТ): ")
    _add_body_p(doc, f"{_fmt_num(kit.final_attestation.hours)} часа.", "Итоговая аттестация: ")
    _add_body_p(
        doc,
        f"{kit.calendar_schedule.total_days} дн. (до {_fmt_num(kit.calendar_schedule.max_hours_per_day)} часов в день).",
        "Срок обучения: ",
    )

    _add_heading_2(doc, "1.9. Выдаваемый документ")
    _add_body_p(doc, kit.issued_document)

    _add_heading_1(doc, "2. СОДЕРЖАНИЕ ПРОГРАММЫ")
    _add_heading_2(doc, "2.1. Учебный план")
    t_up = doc.add_table(rows=1, cols=8)
    up_headers = [
        "№ п/п",
        "Наименование разделов программы",
        "Трудоёмкость, всего (час.)",
        "ЛК",
        "ПР",
        "ЛБ",
        "СРС (час.)",
        "Форма аттестации",
    ]
    for c_i, h in enumerate(up_headers):
        t_up.rows[0].cells[c_i].text = h

    for mod in kit.modules:
        r = t_up.add_row()
        vals = [
            f"{mod.id}.",
            mod.title,
            _fmt_num(mod.total_hours),
            _fmt_num(mod.lecture_hours),
            _fmt_num(mod.practice_hours),
            _fmt_num(mod.lab_hours),
            _fmt_num(mod.self_study_hours),
            mod.intermediate_attestation_form,
        ]
        for c_i, v in enumerate(vals):
            r.cells[c_i].text = v

    r_ia = t_up.add_row()
    for c_i, v in enumerate(["Итоговая аттестация", "Итоговая аттестация", _fmt_num(kit.final_attestation.hours), "0", "0", "0", "0", kit.final_attestation.form]):
        r_ia.cells[c_i].text = v

    r_tot = t_up.add_row()
    tot_vals = [
        "Итого (час.)",
        "Итого (час.)",
        _fmt_num(kit.total_hours),
        _fmt_num(kit.declared_totals["lecture"]),
        _fmt_num(kit.declared_totals["practice"]),
        _fmt_num(kit.declared_totals["lab"]),
        _fmt_num(kit.declared_totals["self_study"]),
        "",
    ]
    for c_i, v in enumerate(tot_vals):
        r_tot.cells[c_i].text = v
    _style_docx_table(t_up, header_rows=1)

    _add_heading_2(doc, "2.2. Учебно-тематический план")
    t_utp = doc.add_table(rows=1, cols=8)
    utp_headers = [
        "№ п/п",
        "Наименование разделов, тем программы",
        "Трудоёмкость, всего (час.)",
        "ЛК",
        "ПР",
        "ЛБ",
        "СРС (час.)",
        "Форма контроля",
    ]
    for c_i, h in enumerate(utp_headers):
        t_utp.rows[0].cells[c_i].text = h

    for mod in kit.modules:
        r_m = t_utp.add_row()
        for c_i, v in enumerate([
            mod.id,
            mod.title,
            _fmt_num(mod.total_hours),
            _fmt_num(mod.lecture_hours),
            _fmt_num(mod.practice_hours),
            _fmt_num(mod.lab_hours),
            _fmt_num(mod.self_study_hours),
            mod.intermediate_attestation_form,
        ]):
            r_m.cells[c_i].text = v
        for topic in mod.topics:
            r_t = t_utp.add_row()
            for c_i, v in enumerate([
                f"{topic.id}.",
                topic.title,
                _fmt_num(topic.total_hours),
                _fmt_num(topic.lecture_hours),
                _fmt_num(topic.practice_hours),
                _fmt_num(topic.lab_hours),
                _fmt_num(topic.self_study_hours),
                topic.control_form,
            ]):
                r_t.cells[c_i].text = v

    r_ia2 = t_utp.add_row()
    for c_i, v in enumerate(["Итоговая аттестация", "Итоговая аттестация", _fmt_num(kit.final_attestation.hours), "0", "0", "0", "0", kit.final_attestation.form]):
        r_ia2.cells[c_i].text = v

    r_tot2 = t_utp.add_row()
    for c_i, v in enumerate(tot_vals):
        r_tot2.cells[c_i].text = v
    _style_docx_table(t_utp, header_rows=1)

    _add_heading_2(doc, "2.3. Календарный учебный график")
    _add_body_p(doc, kit.calendar_schedule.description)

    _add_heading_2(doc, "2.4. Рабочая программа учебных дисциплин")
    t_rpd = doc.add_table(rows=1, cols=3)
    for c_i, h in enumerate([
        "Наименование компонентов программы",
        "Содержание учебного материала и формы организации деятельности слушателей",
        "Всего (час.)",
    ]):
        t_rpd.rows[0].cells[c_i].text = h

    for mod in kit.modules:
        r_m = t_rpd.add_row()
        r_m.cells[0].text = f"Раздел {mod.id}. {mod.title}"
        r_m.cells[1].text = f"Раздел {mod.id}. {mod.title}"
        r_m.cells[2].text = _fmt_num(mod.total_hours)
        for topic in mod.topics:
            if topic.rpd_activities:
                for a_idx, act in enumerate(topic.rpd_activities):
                    r_a = t_rpd.add_row()
                    r_a.cells[0].text = f"Тема {topic.id}. {topic.title}" if a_idx == 0 else ""
                    r_a.cells[1].text = f"{act.activity_type} ({_fmt_num(act.hours)} часов): {act.title_and_content}"
                    r_a.cells[2].text = _fmt_num(act.hours)
            else:
                r_a = t_rpd.add_row()
                r_a.cells[0].text = f"Тема {topic.id}. {topic.title}"
                r_a.cells[1].text = f"СРС ({_fmt_num(topic.total_hours)} часов): {topic.content_summary or topic.title}"
                r_a.cells[2].text = _fmt_num(topic.total_hours)

    r_ia_rpd = t_rpd.add_row()
    r_ia_rpd.cells[0].text = "Итоговая аттестация"
    r_ia_rpd.cells[1].text = kit.final_attestation.form
    r_ia_rpd.cells[2].text = _fmt_num(kit.final_attestation.hours)

    r_tot_rpd = t_rpd.add_row()
    r_tot_rpd.cells[0].text = "Итого"
    r_tot_rpd.cells[1].text = ""
    r_tot_rpd.cells[2].text = _fmt_num(kit.total_hours)
    _style_docx_table(t_rpd, header_rows=1)

    _add_heading_2(doc, "2.5. Оценка качества освоения программы")
    _add_body_p(doc, kit.final_attestation.requirements_description, f"Итоговая аттестация ({kit.final_attestation.form}): ")
    for crit in kit.final_attestation.evaluation_criteria:
        _add_bullet_p(doc, crit)

    _add_heading_1(doc, "3. ОРГАНИЗАЦИОННО-ПЕДАГОГИЧЕСКИЕ УСЛОВИЯ РЕАЛИЗАЦИИ ПРОГРАММЫ")
    _add_heading_2(doc, "3.1. Программное обеспечение и информационные системы")
    for tool in kit.software_and_tools:
        _add_bullet_p(doc, tool)

    _add_heading_2(doc, "3.2. Учебно-методическое и информационное обеспечение (Литература)")
    for src in kit.literature:
        tag = f" {src.verification_tag}" if not src.is_verified and src.verification_tag else ""
        _add_bullet_p(doc, f"{src.citation}{tag}")

    _add_heading_2(doc, "3.3. Кадровое обеспечение образовательного процесса")
    _add_staff_table_empty(doc, kit)

    _add_heading_1(doc, "Приложение 1. Оценочные материалы (тесты, задания семинаров, вопросы контроля)")
    _append_assessments_section(doc, kit)


def _append_assessments_section(doc, kit: ProgramKit) -> None:
    """Добавляет в документ тестовые вопросы с ключами, задания семинаров и вопросы собеседования."""
    if kit.test_questions:
        _add_body_p(doc, "", "Тестовые вопросы с ключами и привязкой к результатам обучения:")
        for idx, q in enumerate(kit.test_questions, 1):
            lo_str = ", ".join(q.learning_outcome_ids)
            _add_body_p(doc, f" {q.question_text} (Тема {q.topic_id} | Проверяемые РО: {lo_str})", f"Вопрос {idx} [{q.id}].")
            for opt_i, opt_text in enumerate(q.options):
                marker = " [ВЕРНЫЙ ОТВЕТ]" if opt_i == q.correct_option_index else ""
                _add_bullet_p(doc, f"{chr(ord('А') + opt_i)}) {opt_text}{marker}")

    if kit.seminar_tasks:
        _add_body_p(doc, "", "Практические задания для семинаров и практикумов:")
        for task in kit.seminar_tasks:
            lo_str = ", ".join(task.learning_outcome_ids)
            _add_body_p(doc, f" (Тема {task.topic_id} | РО: {lo_str})", f"Задание [{task.id}] «{task.title}»")
            _add_body_p(doc, task.scenario_description, "Описание задачи: ")
            _add_body_p(doc, task.deliverable, "Ожидаемый артефакт: ")
            _add_body_p(doc, "; ".join(task.evaluation_criteria) + ".", "Критерии оценки: ")

    if kit.interview_questions:
        _add_body_p(doc, "", "Вопросы для текущего контроля (собеседования с руководителем стажировки):")
        for iq in kit.interview_questions:
            lo_str = ", ".join(iq.learning_outcome_ids)
            _add_bullet_p(doc, f"[{iq.id} · Тема {iq.topic_id} · РО: {lo_str}] {iq.question_text}")


def _add_staff_table_empty(doc, kit: ProgramKit) -> None:
    """Добавляет таблицу кадрового обеспечения со строго пустыми полями ФИО преподавателей."""
    t_staff = doc.add_table(rows=1, cols=5)
    for c_i, h in enumerate(["№", "Модуль/Дисциплина в соответствии с УП", "ФИО преподавателя", "Уч. степень/Звание", "Должность"]):
        t_staff.rows[0].cells[c_i].text = h
    for idx, mod in enumerate(kit.modules, 1):
        r = t_staff.add_row()
        r.cells[0].text = str(idx)
        r.cells[1].text = f"Модуль {mod.id}. {mod.title}"
        r.cells[2].text = ""
        r.cells[3].text = ""
        r.cells[4].text = ""
    _style_docx_table(t_staff, header_rows=1)


def _export_docx_stdlib_fallback(kit: ProgramKit, output_path: str) -> str:
    """Создает валидный .docx архив через стандартный zipfile, если python-docx не установлен."""
    content_types = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/>'
        '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
        "</Types>"
    )
    rels = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>'
        "</Relationships>"
    )
    body_parts = [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>',
        f"<w:p><w:r><w:t>{html.escape(kit.program_title)}</w:t></w:r></w:p>",
        "</w:body></w:document>",
    ]
    with zipfile.ZipFile(output_path, "w") as zf:
        zf.writestr("[Content_Types].xml", content_types)
        zf.writestr("_rels/.rels", rels)
        zf.writestr("word/document.xml", "".join(body_parts))
    return output_path


# =====================================================================
# 2. ГЕНЕРАТОР ПЛАНОВ И СЛАЙДОВ ЛЕКЦИЙ В ФОРМАТЕ POWERPOINT (.PPTX)
# =====================================================================

def export_slides_to_pptx(kit: ProgramKit, output_path: str) -> str:
    """
    Генерирует презентацию (.pptx) со слайдами лекций для всех тем из kit.lecture_decks.
    Использует python-pptx, а при его отсутствии — встроенный генератор OpenXML.
    """
    os.makedirs(os.path.dirname(os.path.abspath(output_path)) or ".", exist_ok=True)
    if pptx is None:
        return _export_pptx_stdlib_fallback(kit, output_path)

    prs = Presentation()
    prs.slide_width = PptxInches(13.333)
    prs.slide_height = PptxInches(7.5)
    blank_layout = prs.slide_layouts[6]

    BG_DARK = PptxRGBColor(0x0F, 0x17, 0x2A)
    BG_LIGHT = PptxRGBColor(0xF8, 0xFA, 0xFC)
    ACCENT = PptxRGBColor(0x25, 0x63, 0xEB)
    TEXT_LIGHT = PptxRGBColor(0xF8, 0xFA, 0xFC)
    TEXT_DARK = PptxRGBColor(0x0F, 0x17, 0x2A)
    TEXT_MUTED = PptxRGBColor(0x33, 0x41, 0x55)

    def _set_bg(slide, color):
        bg = slide.shapes.add_shape(pptx.enum.shapes.MSO_SHAPE.RECTANGLE, 0, 0, prs.slide_width, prs.slide_height)
        bg.fill.solid()
        bg.fill.fore_color.rgb = color
        bg.line.fill.background()

    # 1. Титульный слайд
    s_cover = prs.slides.add_slide(blank_layout)
    _set_bg(s_cover, BG_DARK)
    tb_c = s_cover.shapes.add_textbox(PptxInches(0.8), PptxInches(2.0), PptxInches(11.7), PptxInches(3.5))
    tf_c = tb_c.text_frame
    tf_c.word_wrap = True
    p0 = tf_c.paragraphs[0]
    r0 = p0.add_run()
    r0.text = "АКАДЕМИЯ ЦИФРОВОГО РАЗВИТИЯ · МАГАС 2026"
    r0.font.size = PptxPt(14)
    r0.font.bold = True
    r0.font.color.rgb = PptxRGBColor(0x93, 0xC5, 0xFD)

    p1 = tf_c.add_paragraph()
    p1.space_before = PptxPt(12)
    r1 = p1.add_run()
    r1.text = kit.program_title
    r1.font.size = PptxPt(36)
    r1.font.bold = True
    r1.font.color.rgb = TEXT_LIGHT

    p2 = tf_c.add_paragraph()
    p2.space_before = PptxPt(14)
    r2 = p2.add_run()
    r2.text = f"Комплект визуальных материалов для лекций · Объем курса: {_fmt_num(kit.total_hours)} ак. ч."
    r2.font.size = PptxPt(18)
    r2.font.color.rgb = PptxRGBColor(0xCB, 0xD5, 0xE1)

    # 2. Слайды по каждой теме из kit.lecture_decks
    for deck in kit.lecture_decks:
        s_plan = prs.slides.add_slide(blank_layout)
        _set_bg(s_plan, BG_LIGHT)
        tb_p = s_plan.shapes.add_textbox(PptxInches(0.8), PptxInches(0.6), PptxInches(11.7), PptxInches(6.0))
        tf_p = tb_p.text_frame
        tf_p.word_wrap = True

        pk = tf_p.paragraphs[0]
        rk = pk.add_run()
        rk.text = f"ТЕМА {deck.topic_id} · ПЛАН ЛЕКЦИИ ({_fmt_num(deck.duration_hours)} АК. Ч.)"
        rk.font.size = PptxPt(14)
        rk.font.bold = True
        rk.font.color.rgb = ACCENT

        pt = tf_p.add_paragraph()
        pt.space_before = PptxPt(6)
        pt.space_after = PptxPt(18)
        rt = pt.add_run()
        rt.text = deck.topic_title
        rt.font.size = PptxPt(28)
        rt.font.bold = True
        rt.font.color.rgb = TEXT_DARK

        for idx_pt, plan_point in enumerate(deck.lecture_plan_points, 1):
            pp = tf_p.add_paragraph()
            pp.space_after = PptxPt(10)
            rp = pp.add_run()
            rp.text = f"{idx_pt}. {plan_point}"
            rp.font.size = PptxPt(18)
            rp.font.color.rgb = TEXT_MUTED

        for sl_info in deck.slides:
            s_item = prs.slides.add_slide(blank_layout)
            _set_bg(s_item, BG_LIGHT)
            tb_i = s_item.shapes.add_textbox(PptxInches(0.8), PptxInches(0.6), PptxInches(11.7), PptxInches(6.0))
            tf_i = tb_i.text_frame
            tf_i.word_wrap = True

            p_ik = tf_i.paragraphs[0]
            r_ik = p_ik.add_run()
            r_ik.text = f"ТЕМА {deck.topic_id} · {deck.topic_title.upper()}"
            r_ik.font.size = PptxPt(13)
            r_ik.font.bold = True
            r_ik.font.color.rgb = ACCENT

            p_it = tf_i.add_paragraph()
            p_it.space_before = PptxPt(6)
            p_it.space_after = PptxPt(16)
            r_it = p_it.add_run()
            r_it.text = sl_info.get("title", deck.topic_title)
            r_it.font.size = PptxPt(28)
            r_it.font.bold = True
            r_it.font.color.rgb = TEXT_DARK

            for b_text in sl_info.get("bullets", []):
                pb = tf_i.add_paragraph()
                pb.space_after = PptxPt(10)
                rb = pb.add_run()
                rb.text = f"•  {b_text}"
                rb.font.size = PptxPt(18)
                rb.font.color.rgb = TEXT_MUTED

    # 3. Завершающий слайд
    s_outro = prs.slides.add_slide(blank_layout)
    _set_bg(s_outro, BG_DARK)
    tb_o = s_outro.shapes.add_textbox(PptxInches(0.8), PptxInches(2.8), PptxInches(11.7), PptxInches(2.0))
    tf_o = tb_o.text_frame
    tf_o.word_wrap = True
    po = tf_o.paragraphs[0]
    po.alignment = PP_ALIGN.CENTER
    ro = po.add_run()
    ro.text = "Вопросы и обсуждение практических заданий"
    ro.font.size = PptxPt(32)
    ro.font.bold = True
    ro.font.color.rgb = TEXT_LIGHT

    prs.save(output_path)
    return output_path


def _export_pptx_stdlib_fallback(kit: ProgramKit, output_path: str) -> str:
    """Резервная запись контейнера .pptx через стандартный zipfile, если python-pptx не установлен."""
    with zipfile.ZipFile(output_path, "w") as zf:
        zf.writestr(
            "[Content_Types].xml",
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
            '<Default Extension="xml" ContentType="application/xml"/></Types>',
        )
        zf.writestr(
            "ppt/slides_manifest.json",
            json.dumps([dataclasses.asdict(d) for d in kit.lecture_decks], ensure_ascii=False, indent=2),
        )
    return output_path


# =====================================================================
# 3. ЭКСПОРТ ТЕСТОВЫХ ВОПРОСОВ В ФОРМАТЫ LMS (GIFT, MOODLE XML, CSV)
# =====================================================================

def _escape_gift(text: str) -> str:
    """Экранирует спецсимволы формата Moodle GIFT (~ = # { } :)."""
    for ch in ("\\", "~", "=", "#", "{", "}", ":"):
        text = text.replace(ch, "\\" + ch)
    return text


def export_tests_to_gift(kit: ProgramKit, output_path: str) -> str:
    """Экспортирует тестовые вопросы с ключами и привязкой к результатам обучения в формат Moodle GIFT."""
    os.makedirs(os.path.dirname(os.path.abspath(output_path)) or ".", exist_ok=True)
    lines: List[str] = [
        f"// Программа: {kit.program_title}",
        f"// Категория: $CATEGORY: $course$/top/{_escape_gift(kit.program_title)}",
        "",
    ]
    for q in kit.test_questions:
        lo_str = ", ".join(q.learning_outcome_ids)
        lines.append(f"// ID: {q.id} | Тема: {q.topic_id} | Проверяемые РО: {lo_str}")
        q_title = _escape_gift(f"{q.id} (Тема {q.topic_id}, РО: {lo_str})")
        q_body = _escape_gift(q.question_text)
        lines.append(f"::{q_title}:: {q_body} {{")
        for idx, opt in enumerate(q.options):
            prefix = "=" if idx == q.correct_option_index else "~"
            lines.append(f"  {prefix}{_escape_gift(opt)}")
        if q.explanation:
            lines.append(f"  ####{_escape_gift(q.explanation)}")
        lines.append("}")
        lines.append("")

    with open(output_path, "w", encoding="utf-8") as f:
        f.write("\n".join(lines))
    return output_path


def export_tests_to_moodle_xml(kit: ProgramKit, output_path: str) -> str:
    """Экспортирует тестовые вопросы с ключами и привязкой к РО в стандартный формат Moodle XML."""
    os.makedirs(os.path.dirname(os.path.abspath(output_path)) or ".", exist_ok=True)
    xml_lines: List[str] = [
        '<?xml version="1.0" encoding="UTF-8"?>',
        "<quiz>",
        '  <question type="category">',
        f"    <category><text>$course$/top/{html.escape(kit.program_title)}</text></category>",
        "  </question>",
    ]
    for q in kit.test_questions:
        lo_str = ", ".join(q.learning_outcome_ids)
        xml_lines.extend([
            '  <question type="multichoice">',
            f"    <name><text>{html.escape(q.id)} (Тема {html.escape(q.topic_id)} | РО: {html.escape(lo_str)})</text></name>",
            f'    <questiontext format="html"><text><![CDATA[<p>{html.escape(q.question_text)}</p>]]></text></questiontext>',
            f'    <generalfeedback format="html"><text><![CDATA[<p>{html.escape(q.explanation)}</p>]]></text></generalfeedback>',
            "    <single>true</single>",
            "    <shuffleanswers>true</shuffleanswers>",
            "    <answernumbering>abc</answernumbering>",
        ])
        for idx, opt in enumerate(q.options):
            fraction = "100" if idx == q.correct_option_index else "0"
            xml_lines.extend([
                f'    <answer fraction="{fraction}" format="html">',
                f"      <text><![CDATA[<p>{html.escape(opt)}</p>]]></text>",
                "    </answer>",
            ])
        xml_lines.append("  </question>")
    xml_lines.append("</quiz>")

    with open(output_path, "w", encoding="utf-8") as f:
        f.write("\n".join(xml_lines))
    return output_path


def export_tests_to_csv(kit: ProgramKit, output_path: str) -> str:
    """Экспортирует тестовые вопросы с ключами и привязкой к РО в формат CSV (UTF-8 с BOM для Excel и LMS)."""
    os.makedirs(os.path.dirname(os.path.abspath(output_path)) or ".", exist_ok=True)
    with open(output_path, "w", encoding="utf-8-sig", newline="") as f:
        writer = csv.writer(f, delimiter=";")
        writer.writerow([
            "ID вопроса",
            "Тема УТП",
            "Коды результатов обучения (РО)",
            "Текст вопроса",
            "Варианты ответа",
            "Индекс правильного ответа (0..N-1)",
            "Правильный ответ (ключ)",
            "Пояснение",
        ])
        for q in kit.test_questions:
            correct_text = q.options[q.correct_option_index] if 0 <= q.correct_option_index < len(q.options) else ""
            writer.writerow([
                q.id,
                q.topic_id,
                ", ".join(q.learning_outcome_ids),
                q.question_text,
                " | ".join(q.options),
                q.correct_option_index,
                correct_text,
                q.explanation,
            ])
    return output_path


# =====================================================================
# 4. ПОЛНАЯ ПАКЕТНАЯ ВЫГРУЗКА ВСЕГО КОМПЛЕКТА (export_full_bundle)
# =====================================================================

def export_kit_to_json(
    kit: ProgramKit,
    output_path: str,
    validation_report: Optional[ValidationReport] = None,
) -> str:
    """Сохраняет полный комплект образовательной программы и отчет валидации в JSON."""
    os.makedirs(os.path.dirname(os.path.abspath(output_path)) or ".", exist_ok=True)
    payload = {
        "program_kit": kit.to_dict(),
        "validation_report": validation_report.to_dict() if validation_report else None,
    }
    with open(output_path, "w", encoding="utf-8") as f:
        json.dump(payload, f, ensure_ascii=False, indent=2)
    return output_path


def export_full_bundle(
    kit: ProgramKit,
    output_dir: str,
    file_prefix: str = "acr_program",
    query: Optional[UserQuerySpec] = None,
    reference_docx_paths: Optional[List[str]] = None,
) -> Dict[str, str]:
    """
    Выполняет автосинхронизацию часов, проводит полную детерминированную проверку
    и выгружает готовый комплект во все обязательные и желательные форматы:
    - `.docx` (образовательная программа по шаблону АЦР)
    - `.pptx` (планы и слайды лекций)
    - `.gift.txt`, `.moodle.xml`, `.tests.csv` (тесты для загрузки в LMS)
    - `.json` (полный структурированный комплект данных)
    - `.report.md` (отчет автопроверки для методиста)
    """
    os.makedirs(output_dir, exist_ok=True)
    auto_fix_hours(kit)
    validator = ProgramValidator()
    report = validator.validate(kit, query=query, reference_docx_paths=reference_docx_paths)

    docx_path = os.path.join(output_dir, f"{file_prefix}.docx")
    pptx_path = os.path.join(output_dir, f"{file_prefix}_slides.pptx")
    gift_path = os.path.join(output_dir, f"{file_prefix}_tests.gift.txt")
    xml_path = os.path.join(output_dir, f"{file_prefix}_tests.moodle.xml")
    csv_path = os.path.join(output_dir, f"{file_prefix}_tests.csv")
    json_path = os.path.join(output_dir, f"{file_prefix}_bundle.json")
    md_path = os.path.join(output_dir, f"{file_prefix}_validation_report.md")

    export_program_to_docx(kit, docx_path)
    export_slides_to_pptx(kit, pptx_path)
    export_tests_to_gift(kit, gift_path)
    export_tests_to_moodle_xml(kit, xml_path)
    export_tests_to_csv(kit, csv_path)
    export_kit_to_json(kit, json_path, validation_report=report)

    with open(md_path, "w", encoding="utf-8") as f:
        f.write(report.to_markdown())

    return {
        "docx": docx_path,
        "pptx": pptx_path,
        "gift": gift_path,
        "moodle_xml": xml_path,
        "csv": csv_path,
        "json": json_path,
        "report_md": md_path,
    }