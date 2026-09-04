#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Курсовая работа по ГОСТ 7.32-2017."""

from pathlib import Path

from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING, WD_TAB_ALIGNMENT, WD_TAB_LEADER
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Mm, Pt, RGBColor, Twips

ROOT = Path(__file__).resolve().parent
ASSETS = ROOT / "assets"
OUT = ROOT / "Ulyushev_ZE_Kursovaya_turagentstvo.docx"

FONT = "Times New Roman"
FONT_CODE = "Courier New"


def set_run_font(run, size=14, bold=False, italic=False, name=FONT):
    run.font.name = name
    run.font.size = Pt(size)
    run.bold = bold
    run.italic = italic
    run.font.color.rgb = RGBColor(0, 0, 0)
    rPr = run._element.get_or_add_rPr()
    rFonts = rPr.find(qn("w:rFonts"))
    if rFonts is None:
        rFonts = OxmlElement("w:rFonts")
        rPr.insert(0, rFonts)
    for attr in ("w:ascii", "w:hAnsi", "w:cs", "w:eastAsia"):
        rFonts.set(qn(attr), name)
    lang = rPr.find(qn("w:lang"))
    if lang is None:
        lang = OxmlElement("w:lang")
        rPr.append(lang)
    lang.set(qn("w:val"), "ru-RU")
    lang.set(qn("w:eastAsia"), "ru-RU")


def set_style_font(style, size=14, bold=False, name=FONT):
    style.font.name = name
    style.font.size = Pt(size)
    style.font.bold = bold
    style.font.color.rgb = RGBColor(0, 0, 0)
    style.font.italic = False
    rPr = style.element.get_or_add_rPr()
    rFonts = rPr.find(qn("w:rFonts"))
    if rFonts is None:
        rFonts = OxmlElement("w:rFonts")
        rPr.insert(0, rFonts)
    for attr in ("w:ascii", "w:hAnsi", "w:cs", "w:eastAsia"):
        rFonts.set(qn(attr), name)


def set_paragraph_defaults(p, *, align=WD_ALIGN_PARAGRAPH.JUSTIFY, first=1.25,
                           before=0, after=0, line=1.5, keep_next=False,
                           outline=None, space_after_auto=False):
    pf = p.paragraph_format
    pf.alignment = align
    pf.first_line_indent = Cm(first) if first else Cm(0)
    pf.left_indent = Cm(0)
    pf.right_indent = Cm(0)
    pf.space_before = Pt(before)
    pf.space_after = Pt(after)
    pf.line_spacing = line
    pf.line_spacing_rule = WD_LINE_SPACING.ONE_POINT_FIVE if line == 1.5 else WD_LINE_SPACING.MULTIPLE
    pf.keep_together = False
    pf.keep_with_next = keep_next
    pf.widow_control = True
    pPr = p._p.get_or_add_pPr()
    if outline is not None:
        lvl = pPr.find(qn("w:outlineLvl"))
        if lvl is None:
            lvl = OxmlElement("w:outlineLvl")
            pPr.append(lvl)
        lvl.set(qn("w:val"), str(outline))


def add_text(doc, text, *, align=WD_ALIGN_PARAGRAPH.JUSTIFY, first=1.25,
             before=0, after=0, line=1.5, bold=False, size=14, keep_next=False,
             style=None):
    p = doc.add_paragraph()
    if style:
        p.style = style
    set_paragraph_defaults(
        p, align=align, first=first, before=before, after=after,
        line=line, keep_next=keep_next,
    )
    run = p.add_run(text)
    set_run_font(run, size=size, bold=bold)
    return p


def add_body(doc, text):
    return add_text(doc, text)


def add_empty(doc, n=1, line=1.5):
    for _ in range(n):
        p = doc.add_paragraph()
        set_paragraph_defaults(p, align=WD_ALIGN_PARAGRAPH.CENTER, first=0, line=line)
        p.add_run("")


def add_struct_heading(doc, text, *, new_page=True):
    p = doc.add_paragraph()
    p.style = "Heading 1"
    set_paragraph_defaults(
        p, align=WD_ALIGN_PARAGRAPH.CENTER, first=0, after=12,
        keep_next=True, outline=0,
    )
    p.paragraph_format.page_break_before = new_page
    run = p.add_run(text.upper())
    set_run_font(run, size=14, bold=True)
    return p


def add_section_heading(doc, text, *, new_page=True):
    p = doc.add_paragraph()
    p.style = "Heading 1"
    set_paragraph_defaults(
        p, align=WD_ALIGN_PARAGRAPH.LEFT, first=1.25, after=12,
        keep_next=True, outline=0,
    )
    p.paragraph_format.page_break_before = new_page
    run = p.add_run(text)
    set_run_font(run, size=14, bold=True)
    return p


def add_subheading(doc, text):
    p = doc.add_paragraph()
    p.style = "Heading 2"
    set_paragraph_defaults(
        p, align=WD_ALIGN_PARAGRAPH.LEFT, first=1.25, before=12, after=6,
        keep_next=True, outline=1,
    )
    run = p.add_run(text)
    set_run_font(run, size=14, bold=True)
    return p


def add_bullet(doc, text):
    p = add_text(doc, "– " + text, first=1.25)
    return p


def add_caption_figure(doc, text):
    return add_text(
        doc, text,
        align=WD_ALIGN_PARAGRAPH.CENTER, first=0, before=6, after=12,
    )


def add_caption_table(doc, text):
    return add_text(
        doc, text,
        align=WD_ALIGN_PARAGRAPH.LEFT, first=0, before=12, after=6,
        keep_next=True,
    )


def add_picture(doc, filename, width_cm=15.5):
    p = doc.add_paragraph()
    set_paragraph_defaults(p, align=WD_ALIGN_PARAGRAPH.CENTER, first=0, before=6, after=0)
    run = p.add_run()
    run.add_picture(str(ASSETS / filename), width=Cm(width_cm))
    return p


def shade_cell(cell, color="D9D9D9"):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:fill"), color)
    shd.set(qn("w:val"), "clear")
    tcPr.append(shd)


def set_cell_border(cell):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    tcMar = OxmlElement("w:tcMar")
    for m, v in (("top", "40"), ("left", "60"), ("bottom", "40"), ("right", "60")):
        node = OxmlElement(f"w:{m}")
        node.set(qn("w:w"), v)
        node.set(qn("w:type"), "dxa")
        tcMar.append(node)
    tcPr.append(tcMar)


def fill_cell(cell, text, *, bold=False, size=12, align=WD_ALIGN_PARAGRAPH.LEFT, header=False):
    cell.text = ""
    p = cell.paragraphs[0]
    set_paragraph_defaults(p, align=align, first=0, line=1.15, after=0, before=0)
    run = p.add_run(text)
    set_run_font(run, size=size, bold=bold)
    set_cell_border(cell)
    if header:
        shade_cell(cell)
    vAlign = OxmlElement("w:vAlign")
    vAlign.set(qn("w:val"), "center")
    tcPr = cell._tc.get_or_add_tcPr()
    tcPr.append(vAlign)


def set_table_borders(table):
    tbl = table._tbl
    tblPr = tbl.tblPr
    if tblPr is None:
        tblPr = OxmlElement("w:tblPr")
        tbl.insert(0, tblPr)
    existing = tblPr.find(qn("w:tblBorders"))
    if existing is not None:
        tblPr.remove(existing)
    borders = OxmlElement("w:tblBorders")
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        el = OxmlElement(f"w:{edge}")
        el.set(qn("w:val"), "single")
        el.set(qn("w:sz"), "8")
        el.set(qn("w:space"), "0")
        el.set(qn("w:color"), "000000")
        borders.append(el)
    tblPr.append(borders)
    layout = OxmlElement("w:tblLayout")
    layout.set(qn("w:type"), "fixed")
    tblPr.append(layout)


def add_gost_table(doc, headers, rows, col_widths):
    table = doc.add_table(rows=1 + len(rows), cols=len(headers))
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False
    set_table_borders(table)
    tbl = table._tbl
    tblPr = tbl.tblPr
    grid = tbl.find(qn("w:tblGrid"))
    if grid is not None:
        tbl.remove(grid)
    grid = OxmlElement("w:tblGrid")
    for w in col_widths:
        gc = OxmlElement("w:gridCol")
        gc.set(qn("w:w"), str(int(w * 567)))  # cm -> twips approx
        grid.append(gc)
    tblPr.addnext(grid)
    for i, h in enumerate(headers):
        fill_cell(table.rows[0].cells[i], h, bold=True, header=True,
                  align=WD_ALIGN_PARAGRAPH.CENTER)
        table.rows[0].cells[i].width = Cm(col_widths[i])
    tr = table.rows[0]._tr
    trPr = tr.get_or_add_trPr()
    tblHeader = OxmlElement("w:tblHeader")
    trPr.append(tblHeader)
    for r_i, row in enumerate(rows, start=1):
        for c_i, val in enumerate(row):
            align = WD_ALIGN_PARAGRAPH.CENTER if c_i < 3 else WD_ALIGN_PARAGRAPH.LEFT
            if c_i == 0:
                align = WD_ALIGN_PARAGRAPH.LEFT
            fill_cell(table.rows[r_i].cells[c_i], val, align=align)
            table.rows[r_i].cells[c_i].width = Cm(col_widths[c_i])
    # spacing after table
    p = doc.add_paragraph()
    set_paragraph_defaults(p, first=0, after=6, line=1.0)
    return table


def add_code(doc, code):
    for i, line in enumerate(code.strip("\n").split("\n")):
        p = doc.add_paragraph()
        set_paragraph_defaults(
            p, align=WD_ALIGN_PARAGRAPH.LEFT, first=0, line=1.0, after=0, before=0,
        )
        pf = p.paragraph_format
        pf.left_indent = Cm(1.0)
        pf.line_spacing = 1.0
        pf.line_spacing_rule = WD_LINE_SPACING.SINGLE
        # light shading
        pPr = p._p.get_or_add_pPr()
        shd = OxmlElement("w:shd")
        shd.set(qn("w:val"), "clear")
        shd.set(qn("w:fill"), "F2F2F2")
        pPr.append(shd)
        run = p.add_run(line if line else " ")
        set_run_font(run, size=11, name=FONT_CODE)
    add_empty(doc, 1, line=1.0)


def add_toc_line(doc, title, page, level=0):
    p = doc.add_paragraph()
    set_paragraph_defaults(p, align=WD_ALIGN_PARAGRAPH.LEFT, first=0, line=1.5)
    p.paragraph_format.left_indent = Cm(1.25 * level)
    tab_stops = p.paragraph_format.tab_stops
    tab_stops.add_tab_stop(Cm(16.0), WD_TAB_ALIGNMENT.RIGHT, WD_TAB_LEADER.DOTS)
    run = p.add_run(title)
    set_run_font(run, size=14, bold=(level == 0))
    run2 = p.add_run("\t" + str(page))
    set_run_font(run2, size=14, bold=False)


def add_page_number(paragraph):
    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    set_paragraph_defaults(paragraph, align=WD_ALIGN_PARAGRAPH.CENTER, first=0, line=1.0)
    run = paragraph.add_run()
    set_run_font(run, size=14)

    def fld(char_type):
        el = OxmlElement("w:fldChar")
        el.set(qn("w:fldCharType"), char_type)
        return el

    r = run._r
    r.append(fld("begin"))
    instr = OxmlElement("w:instrText")
    instr.set(qn("xml:space"), "preserve")
    instr.text = " PAGE "
    r.append(instr)
    r.append(fld("end"))


def enable_update_fields(doc):
    settings = doc.settings.element
    upd = OxmlElement("w:updateFields")
    upd.set(qn("w:val"), "true")
    settings.append(upd)
    hyphen = OxmlElement("w:autoHyphenation")
    settings.append(hyphen)
    compat = settings.find(qn("w:compat"))
    if compat is None:
        compat = OxmlElement("w:compat")
        settings.append(compat)


def setup_styles(doc):
    normal = doc.styles["Normal"]
    set_style_font(normal, 14, False)
    pf = normal.paragraph_format
    pf.line_spacing = 1.5
    pf.space_after = Pt(0)
    pf.space_before = Pt(0)
    pf.first_line_indent = Cm(1.25)

    for name, size, bold, outline in (
        ("Heading 1", 14, True, 0),
        ("Heading 2", 14, True, 1),
    ):
        st = doc.styles[name]
        set_style_font(st, size, bold)
        st.paragraph_format.space_before = Pt(0)
        st.paragraph_format.space_after = Pt(12)
        st.paragraph_format.line_spacing = 1.5
        st.paragraph_format.page_break_before = False
        st.font.color.rgb = RGBColor(0, 0, 0)
        try:
            st.element.get_or_add_rPr()
        except Exception:
            pass


def setup_section(doc):
    section = doc.sections[0]
    section.page_width = Mm(210)
    section.page_height = Mm(297)
    section.left_margin = Mm(30)
    section.right_margin = Mm(15)
    section.top_margin = Mm(20)
    section.bottom_margin = Mm(20)
    section.header_distance = Mm(10)
    section.footer_distance = Mm(10)
    section.different_first_page_header_footer = True

    footer = section.footer
    footer.is_linked_to_previous = False
    fp = footer.paragraphs[0]
    add_page_number(fp)

    first_footer = section.first_page_footer
    first_footer.is_linked_to_previous = False
    if first_footer.paragraphs:
        first_footer.paragraphs[0].text = ""

    header = section.header
    header.is_linked_to_previous = False
    if header.paragraphs:
        header.paragraphs[0].text = ""
    first_header = section.first_page_header
    first_header.is_linked_to_previous = False
    if first_header.paragraphs:
        first_header.paragraphs[0].text = ""


def add_title_page(doc):
    def t(text, *, size=14, bold=False, after=0, before=0, align=WD_ALIGN_PARAGRAPH.CENTER, line=1.15):
        return add_text(
            doc, text, align=align, first=0, before=before, after=after,
            line=line, bold=bold, size=size,
        )

    t("МИНИСТЕРСТВО ОБРАЗОВАНИЯ НОВОСИБИРСКОЙ ОБЛАСТИ", bold=True, size=14)
    t("ГОСУДАРСТВЕННОЕ БЮДЖЕТНОЕ ПРОФЕССИОНАЛЬНОЕ", bold=True)
    t("ОБРАЗОВАТЕЛЬНОЕ УЧРЕЖДЕНИЕ НОВОСИБИРСКОЙ ОБЛАСТИ", bold=True)
    t("«НОВОСИБИРСКИЙ ПОЛИТЕХНИЧЕСКИЙ КОЛЛЕДЖ»", bold=True, after=12)
    add_empty(doc, 1, line=1.15)
    t("Специальность 09.02.07", before=6)
    t("Информационные системы и программирование")
    add_empty(doc, 3, line=1.5)
    t("КУРСОВАЯ РАБОТА", bold=True, size=16, before=24, after=12)
    t("Проектирование базы данных турагентства", bold=True, size=14, after=6)
    add_empty(doc, 4, line=1.5)

    table = doc.add_table(rows=7, cols=2)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False
    left = [
        "Руководитель",
        "преподаватель",
        "Кондюрин В. А.",
        "",
        "____________________",
        "подпись",
        "«____» ____________ 2026 г.",
    ]
    right = [
        "Выполнил студент",
        "группы 024",
        "Улюшев Захар Евгеньевич",
        "",
        "____________________",
        "подпись",
        "«____» ____________ 2026 г.",
    ]
    for i in range(7):
        for j, val in enumerate((left[i], right[i])):
            cell = table.rows[i].cells[j]
            cell.width = Cm(8.0)
            cell.text = ""
            p = cell.paragraphs[0]
            align = WD_ALIGN_PARAGRAPH.LEFT if j == 0 else WD_ALIGN_PARAGRAPH.RIGHT
            set_paragraph_defaults(p, align=align, first=0, line=1.15, after=0)
            run = p.add_run(val)
            set_run_font(run, size=14, bold=(i == 2))
        table.rows[i].cells[0].width = Cm(8.2)
        table.rows[i].cells[1].width = Cm(8.2)

    add_empty(doc, 6, line=1.5)
    t("Новосибирск", before=12)
    t("2026")


def build():
    doc = Document()
    setup_styles(doc)
    setup_section(doc)
    enable_update_fields(doc)
    doc.core_properties.author = "Улюшев Захар Евгеньевич"
    doc.core_properties.title = "Проектирование базы данных турагентства"
    doc.core_properties.subject = "Курсовая работа. ГОСТ 7.32-2017"
    doc.core_properties.language = "ru-RU"

    add_title_page(doc)

    # РЕФЕРАТ
    add_struct_heading(doc, "Реферат")
    add_body(
        doc,
        "Курсовая работа содержит 19 с., 7 рис., 5 табл., 8 источников."
    )
    add_body(
        doc,
        "ТУРАГЕНТСТВО, БАЗА ДАННЫХ, РЕЛЯЦИОННАЯ МОДЕЛЬ, MICROSOFT SQL SERVER, "
        "ТАБЛИЦА, ПРЕДСТАВЛЕНИЕ, ТРИГГЕР, ЦЕЛОСТНОСТЬ ДАННЫХ, T-SQL."
    )
    add_body(
        doc,
        "Объект работы — информационное обеспечение деятельности турагентства."
    )
    add_body(
        doc,
        "Цель работы — спроектировать реляционную базу данных турагентства "
        "и реализовать ее в среде Microsoft SQL Server: определить сущности "
        "предметной области, построить схему данных, создать таблицы, "
        "представление и триггер поддержания целостности."
    )
    add_body(
        doc,
        "В результате спроектирована база данных из пяти связанных таблиц "
        "(клиенты, путевки, страны, курорты, отели), разработаны скрипты "
        "создания объектов на языке T-SQL, построено представление для "
        "формирования отчета и триггер, ограничивающий удаление страны "
        "при наличии оформленных путевок."
    )

    # СОДЕРЖАНИЕ
    add_struct_heading(doc, "Содержание")
    toc = [
        ("Введение", 4, 0),
        ("1 Анализ предметной области", 5, 0),
        ("2 Схема базы данных", 6, 0),
        ("3 Создание таблиц", 7, 0),
        ("3.1 Таблица «Клиенты»", 7, 1),
        ("3.2 Таблица «Путевка»", 9, 1),
        ("3.3 Таблица «Страна»", 10, 1),
        ("3.4 Таблица «Курорты»", 11, 1),
        ("3.5 Таблица «Отели»", 13, 1),
        ("4 Представления", 14, 0),
        ("4.1 Представление «Работа турагентства»", 14, 1),
        ("5 Триггеры", 16, 0),
        ("Заключение", 17, 0),
        ("Список использованных источников", 18, 0),
    ]
    for title, page, level in toc:
        add_toc_line(doc, title, page, level)

    # ВВЕДЕНИЕ
    add_struct_heading(doc, "Введение")
    add_body(
        doc,
        "Цель курсовой работы — получить и закрепить знания по проектированию "
        "баз данных, а также навыки разработки программного обеспечения "
        "информационных систем на примере предметной области турагентства."
    )
    add_body(
        doc,
        "База данных — это совокупность хранимых в памяти компьютера данных, "
        "отображающих состояние некоторой предметной области. Данные "
        "взаимосвязаны и организованы по определенным правилам, что обеспечивает "
        "их описание, хранение и манипулирование независимо от прикладных "
        "программ [1]."
    )
    add_body(
        doc,
        "При реляционном подходе упор делается не только на объекты и их "
        "свойства, но и на отношения между ними. Реляционная база данных "
        "представляет собой набор связанных таблиц (отношений). Каждая таблица "
        "состоит из строк (записей, кортежей) и столбцов (полей, атрибутов). "
        "Под каждую категорию данных в записи отводится отдельное поле, которому "
        "присваивается имя и тип данных [2]."
    )
    add_body(
        doc,
        "Туристическая отрасль связана с обработкой большого объема сведений "
        "о клиентах, странах, курортах, средствах размещения и оформленных "
        "путевках. Без автоматизированного учета сложно оперативно вести "
        "документацию, формировать отчетность и корректировать данные при "
        "изменении условий поездки. Поэтому проектирование специализированной "
        "базы данных турагентства является актуальной практической задачей."
    )
    add_body(
        doc,
        "Для достижения поставленной цели решаются следующие задачи:"
    )
    add_bullet(doc, "проанализировать предметную область и выделить сущности;")
    add_bullet(doc, "разработать реляционную схему базы данных;")
    add_bullet(doc, "создать таблицы и определить ключи и типы данных;")
    add_bullet(doc, "разработать представление для отчетной выборки;")
    add_bullet(doc, "реализовать триггер поддержания целостности данных.")
    add_body(
        doc,
        "Практическая реализация выполнена в СУБД Microsoft SQL Server. "
        "Оформление работы соответствует ГОСТ 7.32-2017 [4]."
    )

    # 1
    add_section_heading(doc, "1 Анализ предметной области")
    add_body(
        doc,
        "База данных создается для информационного обслуживания сотрудников "
        "и клиентов турагентства. Система ориентирована на один офис и должна "
        "хранить сведения о клиентах, оформленных путевках, странах направления, "
        "курортах и отелях."
    )
    add_body(
        doc,
        "В результате обследования предметной области выделены следующие "
        "сущности и их основные атрибуты."
    )
    add_body(
        doc,
        "Клиенты — физические лица, обращающиеся в турагентство за подбором "
        "и оформлением поездки. Атрибуты: фамилия, имя, отчество клиента; "
        "паспорт гражданина Российской Федерации; заграничный паспорт; адрес; "
        "контактный телефон; дата рождения."
    )
    add_body(
        doc,
        "Путевка — документ (запись) об оформленной туристической услуге. "
        "Атрибуты: код путевки; идентификатор клиента; идентификатор страны; "
        "идентификатор курорта; идентификатор отеля. Сущность связывает клиента "
        "с выбранным направлением и средством размещения."
    )
    add_body(
        doc,
        "Страна — государство, выбранное как направление поездки. Атрибуты: "
        "код страны; идентификатор страны; идентификатор визы (признак "
        "необходимости визового оформления)."
    )
    add_body(
        doc,
        "Курорты — места отдыха и санаторно-курортного лечения. Атрибуты: "
        "идентификатор страны; идентификатор курорта; код курорта; описание; "
        "фотография."
    )
    add_body(
        doc,
        "Отели — средства размещения туристов. Атрибуты: код страны; "
        "идентификатор отеля; идентификатор страны."
    )
    add_body(
        doc,
        "Между сущностями установлены связи типа «один ко многим»: один клиент "
        "может иметь несколько путевок; одна страна — несколько курортов, "
        "отелей и путевок. Путевка является центральной связующей таблицей "
        "предметной области."
    )

    # 2
    add_section_heading(doc, "2 Схема базы данных")
    add_body(
        doc,
        "Схема базы данных включает описание содержания, структуры и ограничений "
        "целостности, используемых для создания и сопровождения базы данных [3]. "
        "На логическом уровне схема задает набор отношений, ключей и связей."
    )
    add_body(
        doc,
        "Реализованная схема содержит пять таблиц: KLIENTYI (клиенты), "
        "Putyovka (путевка), Strana (страна), Kurortyi (курорты) и Hotel (отели). "
        "Имена объектов заданы транслитерацией русских терминов, что отражает "
        "фактическую реализацию в Microsoft SQL Server."
    )
    add_body(
        doc,
        "Схема таблиц и связей приведена на рисунке 1."
    )
    add_picture(doc, "image1.png", 15.5)
    add_caption_figure(doc, "Рисунок 1 — Схема таблиц базы данных")
    add_body(
        doc,
        "Таблица Putyovka связана с таблицами KLIENTYI, Strana, Kurortyi и Hotel. "
        "Таблицы Kurortyi и Hotel дополнительно связаны со справочником стран. "
        "Ключи выделены в конструкторе связей; по ним обеспечивается соединение "
        "данных при построении запросов и представлений."
    )

    # 3
    add_section_heading(doc, "3 Создание таблиц")
    add_body(
        doc,
        "Таблицы созданы в Microsoft SQL Server. Ниже для каждой таблицы "
        "приведены назначение, структура по ГОСТ, фрагмент конструктора "
        "и скрипт создания на языке T-SQL. Имена столбцов сохранены в том виде, "
        "в котором они заданы в СУБД."
    )

    # 3.1
    add_subheading(doc, "3.1 Таблица «Клиенты»")
    add_body(
        doc,
        "Таблица KLIENTYI предназначена для хранения сведений о клиентах "
        "турагентства. Первичным ключом принято поле полного имени клиента."
    )
    add_caption_table(doc, "Таблица 1 — Структура таблицы KLIENTYI")
    add_gost_table(
        doc,
        ["Имя столбца", "Тип данных", "NULL", "Назначение"],
        [
            ["FIO klienta", "nvarchar(50)", "нет", "ФИО клиента, первичный ключ"],
            ["Pasport RF", "nchar(10)", "нет", "Серия и номер паспорта РФ"],
            ["Zagran pasport", "nchar(10)", "нет", "Данные заграничного паспорта"],
            ["Adres", "nvarchar(50)", "нет", "Адрес проживания"],
            ["Kontaktn. telefone", "int", "нет", "Контактный телефон"],
            ["Date rochdeniya", "int", "нет", "Дата рождения"],
        ],
        [4.2, 3.2, 1.8, 6.3],
    )
    add_body(
        doc,
        "Конструктор таблицы в среде SQL Server Management Studio приведен "
        "на рисунке 2."
    )
    add_picture(doc, "image2.png", 14.0)
    add_caption_figure(doc, "Рисунок 2 — Структура таблицы «Клиенты»")
    add_body(doc, "Скрипт создания таблицы:")
    add_code(doc, """
USE master;
GO

CREATE TABLE dbo.KLIENTYI (
    [FIO klienta]         nvarchar(50) NOT NULL,
    [Pasport RF]          nchar(10)    NOT NULL,
    [Zagran pasport]      nchar(10)    NOT NULL,
    Adres                 nvarchar(50) NOT NULL,
    [Kontaktn. telefone]  int          NOT NULL,
    [Date rochdeniya]     int          NOT NULL,
    CONSTRAINT PK_KLIENTYI PRIMARY KEY ([FIO klienta])
);
GO
""".strip())

    # 3.2
    add_subheading(doc, "3.2 Таблица «Путевка»")
    add_body(
        doc,
        "Таблица Putyovka предназначена для хранения сведений об оформленных "
        "путевках. Каждая запись связывает клиента со страной, курортом и отелем. "
        "Первичный ключ — код путевки."
    )
    add_caption_table(doc, "Таблица 2 — Структура таблицы Putyovka")
    add_gost_table(
        doc,
        ["Имя столбца", "Тип данных", "NULL", "Назначение"],
        [
            ["Kod putyovki", "int", "нет", "Код путевки, первичный ключ"],
            ["ID klienta", "int", "нет", "Идентификатор клиента"],
            ["ID srtanyi", "int", "нет", "Идентификатор страны"],
            ["ID kurorta", "int", "нет", "Идентификатор курорта"],
            ["ID hotel", "int", "нет", "Идентификатор отеля"],
        ],
        [4.2, 3.2, 1.8, 6.3],
    )
    add_body(doc, "Структура таблицы в конструкторе приведена на рисунке 3.")
    add_picture(doc, "image3.png", 14.0)
    add_caption_figure(doc, "Рисунок 3 — Структура таблицы «Путевка»")
    add_body(doc, "Скрипт создания таблицы:")
    add_code(doc, """
CREATE TABLE dbo.Putyovka (
    [Kod putyovki]  int NOT NULL,
    [ID klienta]    int NOT NULL,
    [ID srtanyi]    int NOT NULL,
    [ID kurorta]    int NOT NULL,
    [ID hotel]      int NOT NULL,
    CONSTRAINT PK_Putyovka PRIMARY KEY ([Kod putyovki])
);
GO
""".strip())

    # 3.3
    add_subheading(doc, "3.3 Таблица «Страна»")
    add_body(
        doc,
        "Таблица Strana предназначена для хранения справочника стран направления "
        "и сведений, связанных с визовым оформлением."
    )
    add_caption_table(doc, "Таблица 3 — Структура таблицы Strana")
    add_gost_table(
        doc,
        ["Имя столбца", "Тип данных", "NULL", "Назначение"],
        [
            ["Kod stranyi", "int", "нет", "Код страны, первичный ключ"],
            ["ID stranyi", "int", "да", "Идентификатор страны"],
            ["ID viza", "int", "да", "Идентификатор визы"],
        ],
        [4.2, 3.2, 1.8, 6.3],
    )
    add_body(doc, "Структура таблицы в конструкторе приведена на рисунке 4.")
    add_picture(doc, "image4.png", 14.0)
    add_caption_figure(doc, "Рисунок 4 — Структура таблицы «Страна»")
    add_body(doc, "Скрипт создания таблицы:")
    add_code(doc, """
CREATE TABLE dbo.Strana (
    [Kod stranyi]  int NOT NULL,
    [ID stranyi]   int NULL,
    [ID viza]      int NULL,
    CONSTRAINT PK_Strana PRIMARY KEY ([Kod stranyi])
);
GO
""".strip())

    # 3.4
    add_subheading(doc, "3.4 Таблица «Курорты»")
    add_body(
        doc,
        "Таблица Kurortyi предназначена для хранения данных о санаторно-курортном "
        "лечении и отдыхе: описание курорта и связанное изображение."
    )
    add_caption_table(doc, "Таблица 4 — Структура таблицы Kurortyi")
    add_gost_table(
        doc,
        ["Имя столбца", "Тип данных", "NULL", "Назначение"],
        [
            ["ID stranyi", "int", "нет", "Идентификатор страны, первичный ключ"],
            ["ID kurorta", "int", "да", "Идентификатор курорта"],
            ["Kod kurorta", "int", "да", "Код курорта"],
            ["Opisanie", "nvarchar(50)", "да", "Описание курорта"],
            ["Photo", "image", "да", "Фотография курорта"],
        ],
        [4.2, 3.2, 1.8, 6.3],
    )
    add_body(doc, "Структура таблицы в конструкторе приведена на рисунке 5.")
    add_picture(doc, "image5.png", 14.0)
    add_caption_figure(doc, "Рисунок 5 — Структура таблицы «Курорты»")
    add_body(doc, "Скрипт создания таблицы:")
    add_code(doc, """
CREATE TABLE dbo.Kurortyi (
    [ID stranyi]  int          NOT NULL,
    [ID kurorta]  int          NULL,
    [Kod kurorta] int          NULL,
    Opisanie      nvarchar(50) NULL,
    Photo         image        NULL,
    CONSTRAINT PK_Kurortyi PRIMARY KEY ([ID stranyi])
);
GO
""".strip())

    # 3.5
    add_subheading(doc, "3.5 Таблица «Отели»")
    add_body(
        doc,
        "Таблица Hotel предназначена для хранения сведений о средствах размещения. "
        "Связь со страной задается полями кода и идентификатора страны."
    )
    add_caption_table(doc, "Таблица 5 — Структура таблицы Hotel")
    add_gost_table(
        doc,
        ["Имя столбца", "Тип данных", "NULL", "Назначение"],
        [
            ["Kod stranyi", "int", "нет", "Код страны, первичный ключ"],
            ["ID hotel", "int", "да", "Идентификатор отеля"],
            ["ID Stranyi", "int", "да", "Идентификатор страны"],
        ],
        [4.2, 3.2, 1.8, 6.3],
    )
    add_body(doc, "Структура таблицы в конструкторе приведена на рисунке 6.")
    add_picture(doc, "image6.png", 14.0)
    add_caption_figure(doc, "Рисунок 6 — Структура таблицы «Отели»")
    add_body(doc, "Скрипт создания таблицы:")
    add_code(doc, """
CREATE TABLE dbo.Hotel (
    [Kod stranyi]  int NOT NULL,
    [ID hotel]     int NULL,
    [ID Stranyi]   int NULL,
    CONSTRAINT PK_Hotel PRIMARY KEY ([Kod stranyi])
);
GO
""".strip())

    # 4
    add_section_heading(doc, "4 Представления")
    add_body(
        doc,
        "Представление (VIEW) — объект базы данных, который не хранит собственные "
        "данные, а получает их из таблиц или других представлений при выполнении "
        "запроса [1]. На основе представления формируется отчет для сотрудников "
        "и клиентов турагентства."
    )
    add_subheading(doc, "4.1 Представление «Работа турагентства»")
    add_body(
        doc,
        "Представление объединяет сведения о клиенте, путевке, стране, курорте "
        "и отеле. Оно предназначено для просмотра состава оформленных поездок "
        "без обращения к каждой таблице в отдельности. Конструктор представления "
        "в SQL Server Management Studio приведен на рисунке 7."
    )
    add_picture(doc, "image7.png", 15.5)
    add_caption_figure(doc, "Рисунок 7 — Представление «Работа турагентства»")
    add_body(
        doc,
        "Скрипт создания представления, обеспечивающего соединение таблиц "
        "по ключевым полям:"
    )
    add_code(doc, """
CREATE VIEW dbo.View_RabotaTuragentstva
AS
SELECT
    p.[Kod putyovki],
    p.[ID klienta],
    s.[Kod stranyi],
    s.[ID viza],
    r.[ID kurorta],
    r.Opisanie,
    h.[ID hotel]
FROM dbo.Putyovka AS p
INNER JOIN dbo.Strana AS s
    ON p.[ID srtanyi] = s.[Kod stranyi]
INNER JOIN dbo.Kurortyi AS r
    ON p.[ID kurorta] = r.[ID kurorta]
INNER JOIN dbo.Hotel AS h
    ON p.[ID hotel] = h.[ID hotel];
GO
""".strip())
    add_body(
        doc,
        "В конструкторе (рисунок 7) представление собрано визуально. Скрипт "
        "задает соединение путевки со справочниками страны, курорта и отеля. "
        "Данные клиента в учебной схеме хранятся отдельно; для включения ФИО "
        "в отчет целесообразно ввести суррогатный ключ клиента и связать его "
        "с полем ID klienta таблицы путевок."
    )

    # 5
    add_section_heading(doc, "5 Триггеры")
    add_body(
        doc,
        "Триггер — особая разновидность хранимой процедуры, которая автоматически "
        "выполняется при наступлении события в таблице: вставки, изменения или "
        "удаления строк [1]. Триггеры применяют для реализации сложных ограничений "
        "целостности, которые нельзя выразить только ключами и ограничениями CHECK."
    )
    add_body(
        doc,
        "Для справочника стран задан триггер, запрещающий удаление записи, "
        "если на эту страну уже оформлены путевки. Это исключает появление "
        "«осиротевших» записей в таблице Putyovka."
    )
    add_body(doc, "Скрипт создания триггера:")
    add_code(doc, """
CREATE TRIGGER dbo.trg_Strana_InsteadDelete
ON dbo.Strana
INSTEAD OF DELETE
AS
BEGIN
    SET NOCOUNT ON;

    IF EXISTS (
        SELECT 1
        FROM dbo.Putyovka AS p
        INNER JOIN deleted AS d
            ON p.[ID srtanyi] = d.[Kod stranyi]
    )
    BEGIN
        RAISERROR (
            N'Нельзя удалить страну, на которую оформлены путевки.',
            16, 1
        );
        RETURN;
    END;

    DELETE FROM dbo.Strana
    WHERE [Kod stranyi] IN (SELECT [Kod stranyi] FROM deleted);
END;
GO
""".strip())
    add_body(
        doc,
        "При попытке удалить страну, на которую ссылаются путевки, транзакция "
        "не выполняется, пользователю выдается сообщение об ошибке. Если ссылок "
        "нет, строка справочника удаляется штатно."
    )
    add_body(
        doc,
        "Фрагмент данных таблицы клиентов, используемых при проверке работы "
        "объектов базы данных:"
    )
    add_code(doc, """
SELECT TOP (1000)
    [FIO klienta],
    [Pasport RF],
    [Zagran pasport],
    Adres,
    [Kontaktn. telefone],
    [Date rochdeniya]
FROM dbo.KLIENTYI;
""".strip())

    # ЗАКЛЮЧЕНИЕ
    add_struct_heading(doc, "Заключение")
    add_body(
        doc,
        "В ходе выполнения курсовой работы достигнута цель — спроектирована "
        "база данных турагентства. Решены задачи анализа предметной области, "
        "построения схемы данных, создания таблиц, представления и триггера."
    )
    add_body(
        doc,
        "Спроектированная база данных обеспечивает учет клиентов, путевок, "
        "стран, курортов и отелей. Представление «Работа турагентства» позволяет "
        "получать сводную информацию для отчета. Триггер поддерживает целостность "
        "при удалении записей справочника стран."
    )
    add_body(
        doc,
        "С позиций организации туризма поездка включает транспорт, размещение, "
        "экскурсионное обслуживание и сопутствующие услуги. Информационная система "
        "турагентства должна оперативно отражать состав этих услуг в данных. "
        "Разработанная схема является основой такой системы для одного офиса "
        "и может быть расширена (даты поездки, стоимость, статус заявки, "
        "численность группы)."
    )
    add_body(
        doc,
        "Полученные навыки проектирования реляционной модели и работы с "
        "Microsoft SQL Server применимы при разработке других информационных "
        "систем специальности «Информационные системы и программирование»."
    )

    # СПИСОК
    add_struct_heading(doc, "Список использованных источников")
    sources = [
        "Дейт, К. Дж. Введение в системы баз данных / К. Дж. Дейт ; пер. с англ. — 8-е изд. — Москва : Вильямс, 2016. — 1328 с. — ISBN 978-5-8459-2106-9. — Текст : непосредственный.",
        "Коннолли, Т. Базы данных. Проектирование, реализация и сопровождение. Теория и практика / Т. Коннолли, К. Бегг. — 3-е изд. — Москва : Вильямс, 2017. — 1440 с. — ISBN 978-5-8459-2033-8. — Текст : непосредственный.",
        "Фуфаев, Э. В. Базы данных : учебное пособие для студентов учреждений сред. проф. образования / Э. В. Фуфаев, Д. Э. Фуфаев. — 7-е изд., стер. — Москва : Академия, 2013. — 320 с. — ISBN 978-5-7695-9767-1. — Текст : непосредственный.",
        "ГОСТ 7.32-2017. Система стандартов по информации, библиотечному и издательскому делу. Отчет о научно-исследовательской работе. Структура и правила оформления : межгосударственный стандарт. — Москва : Стандартинформ, 2017. — 32 с. — Текст : непосредственный.",
        "ГОСТ Р 7.0.100-2018. Система стандартов по информации, библиотечному и издательскому делу. Библиографическая запись. Библиографическое описание. Общие требования и правила составления : национальный стандарт Российской Федерации. — Москва : Стандартинформ, 2018. — 124 с. — Текст : непосредственный.",
        "Microsoft SQL Server documentation. — Текст : электронный // Microsoft Learn : [сайт]. — URL: https://learn.microsoft.com/sql/sql-server (дата обращения: 04.09.2026).",
        "База данных. — Текст : электронный // Википедия : [сайт]. — URL: https://ru.wikipedia.org/wiki/База_данных (дата обращения: 04.09.2026).",
        "Представление (базы данных). — Текст : электронный // Википедия : [сайт]. — URL: https://ru.wikipedia.org/wiki/Представление_(базы_данных) (дата обращения: 04.09.2026).",
    ]
    for i, src in enumerate(sources, start=1):
        p = doc.add_paragraph()
        set_paragraph_defaults(p, align=WD_ALIGN_PARAGRAPH.JUSTIFY, first=0, line=1.5)
        p.paragraph_format.left_indent = Cm(1.25)
        p.paragraph_format.first_line_indent = Cm(-1.25)
        run = p.add_run(f"{i}. {src}")
        set_run_font(run, size=14)

    doc.save(OUT)
    print(f"Saved: {OUT}")


if __name__ == "__main__":
    build()
