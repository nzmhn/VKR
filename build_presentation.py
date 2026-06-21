#!/usr/bin/env python3
"""
Сборка презентации ВКР по теме «Разработка стратегии продвижения
инновационного продукта Face2 ПАО АК БАРС БАНК».

Скрипт строит .pptx-файл напрямую через OOXML (без python-pptx),
поэтому работает в любой среде с Python 3.7+.
"""

from __future__ import annotations

import os
import zipfile
from html import escape
from typing import List, Tuple

OUT = "VKR/Презентация_ВКР_Face2.pptx"

# =====================================================================
# Цветовая палитра (HEX без #)
# =====================================================================
NAVY = "003366"   # тёмно-синий — заголовки, фон карточек
GOLD = "BFA46F"   # золотой АК Барс — акцентные цифры
GREEN = "00B050"  # положительное
RED = "C00000"    # предупреждение, угроза
GREY = "7F7F7F"   # вспомогательный текст
LIGHT_BLUE = "BDD7EE"
LIGHT_RED = "F4B084"
LIGHT_GREEN = "C6E0B4"
LIGHT_YELLOW = "FFE699"
WHITE = "FFFFFF"
BG_LIGHT = "F2F2F2"

# =====================================================================
# Геометрия слайда (16:9), единицы — EMU (914400 EMU = 1 inch)
# =====================================================================
SLIDE_W = 12192000  # 13.33"
SLIDE_H = 6858000   # 7.5"


# =====================================================================
# Хелперы для построения XML
# =====================================================================
def emu(inches: float) -> int:
    return int(inches * 914400)


def text_run(text: str, *, bold=False, size=18, color=NAVY, align="l") -> str:
    """Один <a:r> или строка <a:p> с одним run-ом."""
    return (
        f'<a:r><a:rPr lang="ru-RU" sz="{int(size*100)}" '
        f'b="{1 if bold else 0}" dirty="0">'
        f'<a:solidFill><a:srgbClr val="{color}"/></a:solidFill>'
        f'<a:latin typeface="Calibri"/></a:rPr>'
        f'<a:t>{escape(text)}</a:t></a:r>'
    )


def paragraph(runs: List[str], *, align="l", bullet=False, indent=0) -> str:
    bullet_xml = ""
    if bullet:
        bullet_xml = (
            '<a:pPr marL="285750" indent="-285750"><a:buChar char="•"/></a:pPr>'
        )
    if indent and not bullet:
        bullet_xml = f'<a:pPr marL="{indent}"><a:buNone/></a:pPr>'
    if align != "l":
        # переписываем pPr с алинментом
        algn = {"c": "ctr", "r": "r"}.get(align, "l")
        if bullet:
            bullet_xml = (
                f'<a:pPr marL="285750" indent="-285750" algn="{algn}">'
                f'<a:buChar char="•"/></a:pPr>'
            )
        else:
            bullet_xml = f'<a:pPr algn="{algn}"><a:buNone/></a:pPr>'
    return f"<a:p>{bullet_xml}{''.join(runs)}</a:p>"


def textbox(
    sp_id: int,
    name: str,
    x: int,
    y: int,
    w: int,
    h: int,
    paragraphs: List[str],
    *,
    fill_color: str | None = None,
    line_color: str | None = None,
    line_width: int = 12700,
) -> str:
    """Свободный текстовый блок (rectangle) с произвольным набором абзацев."""
    fill_xml = (
        f'<a:solidFill><a:srgbClr val="{fill_color}"/></a:solidFill>'
        if fill_color
        else "<a:noFill/>"
    )
    line_xml = (
        f'<a:ln w="{line_width}"><a:solidFill><a:srgbClr val="{line_color}"/></a:solidFill></a:ln>'
        if line_color
        else "<a:ln><a:noFill/></a:ln>"
    )
    body = "".join(paragraphs)
    return f"""<p:sp>
<p:nvSpPr><p:cNvPr id="{sp_id}" name="{escape(name)}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
<p:spPr>
<a:xfrm><a:off x="{x}" y="{y}"/><a:ext cx="{w}" cy="{h}"/></a:xfrm>
<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
{fill_xml}
{line_xml}
</p:spPr>
<p:txBody><a:bodyPr wrap="square" anchor="t" anchorCtr="0" lIns="91440" tIns="91440" rIns="91440" bIns="91440"/><a:lstStyle/>
{body}
</p:txBody>
</p:sp>"""


def table_shape(
    sp_id: int,
    x: int,
    y: int,
    w: int,
    h: int,
    rows: List[List[str]],
    *,
    col_widths: List[int] | None = None,
    header_fill: str = NAVY,
    header_color: str = WHITE,
    body_fill: str = WHITE,
    body_color: str = NAVY,
    header_size: int = 16,
    body_size: int = 14,
    bold_first_col: bool = False,
    aligns: List[str] | None = None,
) -> str:
    """Простая таблица. rows[0] — заголовок."""
    n_cols = len(rows[0])
    if col_widths is None:
        col_widths = [w // n_cols] * n_cols
    if aligns is None:
        aligns = ["l"] * n_cols

    grid_cols = "".join(f'<a:gridCol w="{cw}"/>' for cw in col_widths)
    row_h = h // len(rows)

    rows_xml = []
    for r_idx, row in enumerate(rows):
        is_header = r_idx == 0
        cells_xml = []
        for c_idx, cell in enumerate(row):
            fill = header_fill if is_header else body_fill
            color = header_color if is_header else body_color
            sz = header_size if is_header else body_size
            bold = is_header or (bold_first_col and c_idx == 0)
            algn = {"c": "ctr", "r": "r"}.get(aligns[c_idx], "l")
            cells_xml.append(f"""<a:tc>
<a:txBody><a:bodyPr anchor="ctr" lIns="36576" tIns="36576" rIns="36576" bIns="36576"/><a:lstStyle/>
<a:p><a:pPr algn="{algn}"/>
<a:r><a:rPr lang="ru-RU" sz="{sz*100}" b="{1 if bold else 0}" dirty="0">
<a:solidFill><a:srgbClr val="{color}"/></a:solidFill>
<a:latin typeface="Calibri"/></a:rPr>
<a:t>{escape(cell)}</a:t></a:r></a:p>
</a:txBody>
<a:tcPr>
<a:lnL w="6350"><a:solidFill><a:srgbClr val="A6A6A6"/></a:solidFill></a:lnL>
<a:lnR w="6350"><a:solidFill><a:srgbClr val="A6A6A6"/></a:solidFill></a:lnR>
<a:lnT w="6350"><a:solidFill><a:srgbClr val="A6A6A6"/></a:solidFill></a:lnT>
<a:lnB w="6350"><a:solidFill><a:srgbClr val="A6A6A6"/></a:solidFill></a:lnB>
<a:solidFill><a:srgbClr val="{fill}"/></a:solidFill>
</a:tcPr>
</a:tc>""")
        rows_xml.append(f'<a:tr h="{row_h}">{"".join(cells_xml)}</a:tr>')

    return f"""<p:graphicFrame>
<p:nvGraphicFramePr><p:cNvPr id="{sp_id}" name="Table {sp_id}"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr>
<p:xfrm><a:off x="{x}" y="{y}"/><a:ext cx="{w}" cy="{h}"/></p:xfrm>
<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table">
<a:tbl>
<a:tblPr firstRow="1" bandRow="1"/>
<a:tblGrid>{grid_cols}</a:tblGrid>
{"".join(rows_xml)}
</a:tbl>
</a:graphicData></a:graphic>
</p:graphicFrame>"""


# =====================================================================
# Заголовок страницы и подвал
# =====================================================================
def page_header(title: str, page_num: int, total: int = 10) -> str:
    """Тёмно-синяя плашка заголовка + номер страницы внизу справа."""
    title_box = textbox(
        100 + page_num * 10,
        f"Header_{page_num}",
        emu(0.4),
        emu(0.3),
        emu(12.5),
        emu(0.7),
        [paragraph([text_run(title, bold=True, size=28, color=WHITE)])],
        fill_color=NAVY,
    )
    page_box = textbox(
        200 + page_num * 10,
        f"PageNum_{page_num}",
        emu(11.5),
        emu(7.0),
        emu(1.4),
        emu(0.4),
        [paragraph([text_run(f"{page_num} / {total}", size=11, color=GREY)], align="r")],
    )
    return title_box + page_box


def make_slide_xml(shapes_xml: str, *, bg_color: str = WHITE) -> str:
    return f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
       xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
       xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
<p:cSld>
<p:bg><p:bgPr><a:solidFill><a:srgbClr val="{bg_color}"/></a:solidFill><a:effectLst/></p:bgPr></p:bg>
<p:spTree>
<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>
{shapes_xml}
</p:spTree>
</p:cSld>
<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>
</p:sld>"""


# =====================================================================
# СЛАЙД 1 — Титульный
# =====================================================================
def slide_1() -> str:
    shapes = []
    # Узкая золотая полоса сверху
    shapes.append(textbox(
        10, "TopBar", 0, 0, SLIDE_W, emu(0.3), [], fill_color=GOLD,
    ))
    # Шапка вуза
    shapes.append(textbox(
        11, "VuzHeader", emu(0.5), emu(0.5), emu(12.3), emu(1.2),
        [
            paragraph([text_run("МИНОБРНАУКИ РОССИЙСКОЙ ФЕДЕРАЦИИ", size=14, color=GREY)], align="c"),
            paragraph([text_run("[ПОЛНОЕ НАЗВАНИЕ ВУЗА]", bold=True, size=18, color=NAVY)], align="c"),
            paragraph([text_run("Институт экономики, управления и финансов", size=14, color=NAVY)], align="c"),
            paragraph([text_run("Кафедра [название кафедры]", size=14, color=NAVY)], align="c"),
        ],
    ))
    # Разделитель
    shapes.append(textbox(
        12, "Sep1", emu(3), emu(2.0), emu(7.3), emu(0.05),
        [paragraph([text_run("", size=1)])],
        fill_color=GOLD,
    ))
    # Тип работы
    shapes.append(textbox(
        13, "DocType", emu(0.5), emu(2.2), emu(12.3), emu(0.5),
        [paragraph([text_run("ВЫПУСКНАЯ КВАЛИФИКАЦИОННАЯ РАБОТА", bold=True, size=20, color=NAVY)], align="c")],
    ))
    # Тема (главная плашка)
    shapes.append(textbox(
        14, "Theme", emu(0.8), emu(3.0), emu(11.7), emu(2.0),
        [
            paragraph([text_run("Тема:", size=14, color=GREY)], align="c"),
            paragraph([text_run("«Разработка стратегии продвижения", bold=True, size=24, color=WHITE)], align="c"),
            paragraph([text_run("инновационного продукта Face2", bold=True, size=24, color=WHITE)], align="c"),
            paragraph([text_run("ПАО „АК БАРС“ БАНК»", bold=True, size=24, color=GOLD)], align="c"),
        ],
        fill_color=NAVY,
    ))
    # Автор и руководитель — два столбца
    shapes.append(textbox(
        15, "Author", emu(0.8), emu(5.4), emu(5.7), emu(1.3),
        [
            paragraph([text_run("Выполнил(а):", bold=True, size=14, color=GREY)]),
            paragraph([text_run("студент(ка) группы [______]", size=14, color=NAVY)]),
            paragraph([text_run("[Фамилия Имя Отчество]", size=14, color=NAVY)]),
        ],
    ))
    shapes.append(textbox(
        16, "Supervisor", emu(6.8), emu(5.4), emu(5.7), emu(1.3),
        [
            paragraph([text_run("Научный руководитель:", bold=True, size=14, color=GREY)]),
            paragraph([text_run("[уч. степень, должность]", size=14, color=NAVY)]),
            paragraph([text_run("[Фамилия Имя Отчество]", size=14, color=NAVY)]),
        ],
    ))
    # Низ
    shapes.append(textbox(
        17, "City", emu(0.5), emu(6.9), emu(12.3), emu(0.5),
        [paragraph([text_run("Казань, 2026", bold=True, size=14, color=NAVY)], align="c")],
    ))
    return make_slide_xml("".join(shapes))


# =====================================================================
# СЛАЙД 2 — Актуальность (3 карточки)
# =====================================================================
def slide_2() -> str:
    shapes = [page_header("Актуальность исследования", 2)]

    cards = [
        (NAVY, "1. РОСТ РЫНКА", "145,9 МЛН", "биометрических платежей", "× 9 за год", "112,5 млрд ₽ оборот", "10+ млн пользователей в ЕБС"),
        (RED, "2. УЖЕСТОЧЕНИЕ ЗАКОНА", "20 МЛН ₽", "штраф по КоАП", "до 5 лет тюрьмы", "572-ФЗ + 420-ФЗ + 421-ФЗ", "только аккредитованные"),
        (GOLD, "3. КОНКУРЕНЦИЯ", "700 ТЫС.", "терминалов у Сбера", "30 млн его клиентов", "VisionLabs · NtechLab · OVISION", "Face2: точность 99,5 %"),
    ]
    card_w = emu(3.95)
    gap = emu(0.2)
    start_x = emu(0.4)
    for i, (color, num_label, big_num, small_label, accent1, accent2, accent3) in enumerate(cards):
        x = start_x + i * (card_w + gap)
        shapes.append(textbox(
            300 + i, f"Card{i}", x, emu(1.3), card_w, emu(5.4),
            [
                paragraph([text_run(num_label, bold=True, size=16, color=WHITE)], align="c"),
                paragraph([text_run("", size=8)]),
                paragraph([text_run(big_num, bold=True, size=44, color=GOLD if color != GOLD else WHITE)], align="c"),
                paragraph([text_run(small_label, size=14, color=WHITE)], align="c"),
                paragraph([text_run("", size=10)]),
                paragraph([text_run("─────────", size=10, color=WHITE)], align="c"),
                paragraph([text_run("", size=8)]),
                paragraph([text_run(accent1, bold=True, size=14, color=WHITE)], align="c"),
                paragraph([text_run(accent2, size=12, color=WHITE)], align="c"),
                paragraph([text_run(accent3, size=12, color=WHITE)], align="c"),
            ],
            fill_color=color,
        ))

    # Источники
    shapes.append(textbox(
        390, "Sources", emu(0.4), emu(6.85), emu(12.3), emu(0.4),
        [paragraph([text_run("Источники: Банк России (2025); ФЗ № 572-ФЗ; ПАО Сбербанк (2024); face2.ru", size=10, color=GREY)])],
    ))
    return make_slide_xml("".join(shapes))


# =====================================================================
# СЛАЙД 3 — Цель, задачи, методология
# =====================================================================
def slide_3() -> str:
    shapes = [page_header("Цель, задачи и методология", 3)]

    # Цель — главная плашка
    shapes.append(textbox(
        310, "Goal", emu(1.0), emu(1.3), emu(11.3), emu(0.9),
        [
            paragraph([text_run("ЦЕЛЬ:", bold=True, size=14, color=GOLD)]),
            paragraph([text_run("разработка стратегии продвижения продукта Face2 ПАО «АК БАРС» БАНК", bold=True, size=18, color=WHITE)]),
        ],
        fill_color=NAVY,
    ))

    # 5 задач — горизонтальный ряд
    tasks = [
        ("1", "Теория", "маркетинга\nинноваций"),
        ("2", "Биометрия", "специфика\nв банках"),
        ("3", "Банк и продукт", "АК Барс\nи Face2"),
        ("4", "Аудит", "SWOT\nи gap-анализ"),
        ("5", "Стратегия", "и оценка\nэффекта"),
    ]
    task_w = emu(2.3)
    task_gap = emu(0.1)
    start_x = emu(0.5)
    for i, (num, name, sub) in enumerate(tasks):
        x = start_x + i * (task_w + task_gap)
        shapes.append(textbox(
            320 + i, f"Task{i+1}", x, emu(2.4), task_w, emu(1.7),
            [
                paragraph([text_run(num, bold=True, size=36, color=GOLD)], align="c"),
                paragraph([text_run(name, bold=True, size=14, color=NAVY)], align="c"),
                paragraph([text_run(sub.split("\n")[0], size=11, color=GREY)], align="c"),
                paragraph([text_run(sub.split("\n")[1] if "\n" in sub else "", size=11, color=GREY)], align="c"),
            ],
            fill_color=BG_LIGHT,
            line_color=NAVY,
            line_width=12700,
        ))

    # Объект и предмет
    shapes.append(textbox(
        330, "Object", emu(0.5), emu(4.4), emu(6.0), emu(1.0),
        [
            paragraph([text_run("ОБЪЕКТ:", bold=True, size=12, color=GOLD)]),
            paragraph([text_run("маркетинговая деятельность банка по продвижению Face2", size=13, color=NAVY)]),
        ],
        fill_color=BG_LIGHT,
    ))
    shapes.append(textbox(
        331, "Subject", emu(6.7), emu(4.4), emu(6.1), emu(1.0),
        [
            paragraph([text_run("ПРЕДМЕТ:", bold=True, size=12, color=GOLD)]),
            paragraph([text_run("инструменты и управленческие отношения стратегии", size=13, color=NAVY)]),
        ],
        fill_color=BG_LIGHT,
    ))

    # Методология
    shapes.append(textbox(
        332, "Methods", emu(0.5), emu(5.6), emu(12.3), emu(1.4),
        [
            paragraph([text_run("МЕТОДОЛОГИЯ:", bold=True, size=14, color=GOLD)]),
            paragraph([text_run("Общенаучные — анализ, синтез, обобщение, сравнение.", size=13, color=NAVY)]),
            paragraph([text_run("Специальные — диффузия инноваций Э. Роджерса · STP · 7P · маркетинговая воронка · PESTEL · SWOT · 5 сил М. Портера · финансово-экономический анализ", size=13, color=NAVY)]),
        ],
        fill_color=NAVY if False else BG_LIGHT,
    ))
    return make_slide_xml("".join(shapes))


# =====================================================================
# СЛАЙД 4 — Задача 1. Теория (воронка с подсветкой)
# =====================================================================
def slide_4() -> str:
    shapes = [page_header("Задача 1. Теоретические основы продвижения инноваций", 4)]

    # Левая часть — список моделей
    shapes.append(textbox(
        410, "ModelsTitle", emu(0.5), emu(1.4), emu(5.5), emu(0.5),
        [paragraph([text_run("ПРИМЕНЁННЫЕ МОДЕЛИ", bold=True, size=16, color=GOLD)])],
    ))
    models = [
        "✓ Диффузия инноваций Э. Роджерса",
        "✓ STP (сегментация → таргетинг → позиционирование)",
        "✓ Расширенный комплекс маркетинга 7P",
        "✓ Жизненный цикл инновационного продукта",
        "✓ Маркетинговая воронка",
    ]
    shapes.append(textbox(
        411, "ModelsList", emu(0.5), emu(2.0), emu(5.5), emu(3.5),
        [paragraph([text_run(m, size=15, color=NAVY)]) for m in models],
        fill_color=BG_LIGHT,
    ))

    # Правая часть — воронка как 6 прямоугольников разной ширины
    funnel_y = emu(1.4)
    funnel_x_center = emu(9.5)
    funnel_widths = [emu(5.5), emu(4.5), emu(3.7), emu(2.9), emu(2.1), emu(1.3)]
    funnel_labels = [
        ("Осведомлённость", BG_LIGHT, NAVY),
        ("Интерес", BG_LIGHT, NAVY),
        ("ДОВЕРИЕ ★", GOLD, WHITE),
        ("Пробное использование", BG_LIGHT, NAVY),
        ("Принятие", BG_LIGHT, NAVY),
        ("Лояльность", BG_LIGHT, NAVY),
    ]
    block_h = emu(0.55)
    for i, ((label, color, text_color), width) in enumerate(zip(funnel_labels, funnel_widths)):
        x = funnel_x_center - width // 2
        y = funnel_y + i * (block_h + emu(0.05))
        shapes.append(textbox(
            420 + i, f"Funnel{i}", x, y, width, block_h,
            [paragraph([text_run(label, bold=(i == 2), size=14, color=text_color)], align="c")],
            fill_color=color,
            line_color=NAVY if i == 2 else None,
            line_width=25400 if i == 2 else 0,
        ))

    # Вывод
    shapes.append(textbox(
        430, "Result", emu(0.5), emu(5.7), emu(12.3), emu(1.3),
        [
            paragraph([text_run("РЕЗУЛЬТАТ", bold=True, size=14, color=GOLD)]),
            paragraph([text_run("В воронке инновационного продукта обособляется самостоятельный этап ", size=15, color=NAVY),
                       text_run("ДОВЕРИЯ", bold=True, size=15, color=GOLD),
                       text_run(" — без него потребитель не переходит от интереса к пробному использованию", size=15, color=NAVY)]),
        ],
        fill_color=NAVY if False else BG_LIGHT,
    ))
    return make_slide_xml("".join(shapes))


# =====================================================================
# СЛАЙД 5 — Задача 2. Биометрия (5 vs 4)
# =====================================================================
def slide_5() -> str:
    shapes = [page_header("Задача 2. Особенности продвижения биометрии в банках", 5)]

    # Левый блок — 5 барьеров
    shapes.append(textbox(
        510, "BarriersTitle", emu(0.5), emu(1.4), emu(5.8), emu(0.6),
        [paragraph([text_run("5 ГРУПП БАРЬЕРОВ", bold=True, size=18, color=WHITE)], align="c")],
        fill_color=RED,
    ))
    barriers = [
        ("Психологические", "страх утечки биометрии"),
        ("Когнитивные", "непонимание принципа работы"),
        ("Поведенческие", "привычка к картам и паролям"),
        ("Организационные (B2B)", "цена внедрения, длинный цикл"),
        ("Регуляторно-репутационные", "риск штрафов и инцидентов"),
    ]
    for i, (name, desc) in enumerate(barriers):
        shapes.append(textbox(
            511 + i, f"Barrier{i}", emu(0.5), emu(2.0) + i * emu(0.7), emu(5.8), emu(0.65),
            [
                paragraph([text_run(f"{i+1}. {name}", bold=True, size=13, color=NAVY)]),
                paragraph([text_run(desc, size=11, color=GREY)]),
            ],
            fill_color=BG_LIGHT,
        ))

    # Правый блок — 4 опоры
    shapes.append(textbox(
        530, "PillarsTitle", emu(6.9), emu(1.4), emu(5.8), emu(0.6),
        [paragraph([text_run("4 ОПОРЫ УСПЕХА", bold=True, size=18, color=WHITE)], align="c")],
        fill_color=GREEN,
    ))
    pillars = [
        ("Правовая", "572-ФЗ, аккредитация Минцифры"),
        ("Социально-психологическая", "доверие аудитории"),
        ("Технологическая", "точность, скорость, надёжность"),
        ("Экономическая", "понятный ROI для B2B"),
    ]
    for i, (name, desc) in enumerate(pillars):
        shapes.append(textbox(
            531 + i, f"Pillar{i}", emu(6.9), emu(2.0) + i * emu(0.85), emu(5.8), emu(0.78),
            [
                paragraph([text_run(f"{i+1}. {name}", bold=True, size=13, color=NAVY)]),
                paragraph([text_run(desc, size=11, color=GREY)]),
            ],
            fill_color=BG_LIGHT,
        ))

    # Вывод
    shapes.append(textbox(
        540, "Result", emu(0.5), emu(5.8), emu(12.3), emu(1.1),
        [
            paragraph([text_run("РЕЗУЛЬТАТ", bold=True, size=14, color=GOLD)]),
            paragraph([text_run("Модель продвижения биометрических банковских продуктов с учётом 572-ФЗ и архитектуры ЕБС", size=15, color=NAVY)]),
        ],
        fill_color=BG_LIGHT,
    ))
    return make_slide_xml("".join(shapes))


# =====================================================================
# СЛАЙД 6 — АК Барс Банк (финансовый дашборд)
# =====================================================================
def slide_6() -> str:
    shapes = [page_header("Задача 3. ПАО «АК БАРС» БАНК — финансовый масштаб", 6)]

    # 4 KPI-плитки
    kpis = [
        ("АКТИВЫ", "1 157", "млрд ₽", "↑ 22 % за год"),
        ("ПРИБЫЛЬ", "15,7", "млрд ₽", "↑ × 2 к 2024"),
        ("КАПИТАЛ", "126", "млрд ₽", "↑ 13 %"),
        ("ESG", "A", "НКР", "+ ESG-III(b) Эксперт РА"),
    ]
    plate_w = emu(2.95)
    plate_gap = emu(0.13)
    start_x = emu(0.45)
    for i, (label, num, unit, change) in enumerate(kpis):
        x = start_x + i * (plate_w + plate_gap)
        shapes.append(textbox(
            610 + i, f"KPI{i}", x, emu(1.3), plate_w, emu(2.0),
            [
                paragraph([text_run(label, bold=True, size=14, color=WHITE)], align="c"),
                paragraph([text_run("", size=4)]),
                paragraph([text_run(num, bold=True, size=44, color=GOLD)], align="c"),
                paragraph([text_run(unit, size=13, color=WHITE)], align="c"),
                paragraph([text_run("─────", size=8, color=GOLD)], align="c"),
                paragraph([text_run(change, size=11, color=WHITE)], align="c"),
            ],
            fill_color=NAVY,
        ))

    # Гистограмма роста активов (горизонтальные столбики)
    shapes.append(textbox(
        620, "ChartTitle", emu(0.5), emu(3.6), emu(8.0), emu(0.5),
        [paragraph([text_run("ДИНАМИКА АКТИВОВ ГРУППЫ, млрд ₽", bold=True, size=14, color=NAVY)])],
    ))
    # 2024
    shapes.append(textbox(
        621, "Bar2024Lbl", emu(0.5), emu(4.2), emu(0.7), emu(0.4),
        [paragraph([text_run("2024", bold=True, size=12, color=NAVY)])],
    ))
    shapes.append(textbox(
        622, "Bar2024", emu(1.3), emu(4.2), emu(5.5), emu(0.4),
        [paragraph([text_run("  947", bold=True, size=14, color=WHITE)])],
        fill_color=NAVY,
    ))
    # 2025
    shapes.append(textbox(
        623, "Bar2025Lbl", emu(0.5), emu(4.7), emu(0.7), emu(0.4),
        [paragraph([text_run("2025", bold=True, size=12, color=NAVY)])],
    ))
    shapes.append(textbox(
        624, "Bar2025", emu(1.3), emu(4.7), emu(6.7), emu(0.4),
        [paragraph([text_run("  1 157", bold=True, size=14, color=WHITE)])],
        fill_color=GOLD,
    ))

    # Рейтинги (правая колонка)
    shapes.append(textbox(
        625, "Ratings", emu(8.7), emu(3.6), emu(4.0), emu(2.0),
        [
            paragraph([text_run("КРЕДИТНЫЕ РЕЙТИНГИ", bold=True, size=12, color=GOLD)]),
            paragraph([text_run("• A(RU) — АКРА", size=12, color=NAVY)]),
            paragraph([text_run("• ruA — Эксперт РА", size=12, color=NAVY)]),
            paragraph([text_run("• A.ru — НКР", size=12, color=NAVY)]),
            paragraph([text_run("", size=6)]),
            paragraph([text_run("ESG-РЕЙТИНГИ", bold=True, size=12, color=GOLD)]),
            paragraph([text_run("• ESG-A — НКР", size=12, color=NAVY)]),
            paragraph([text_run("• ESG-III(b) — Эксперт РА", size=12, color=NAVY)]),
        ],
        fill_color=BG_LIGHT,
    ))

    # Вывод
    shapes.append(textbox(
        630, "Result", emu(0.5), emu(5.85), emu(12.3), emu(1.1),
        [
            paragraph([text_run("ВЫВОД", bold=True, size=14, color=GOLD)]),
            paragraph([text_run("Банк располагает финансовыми ресурсами и ESG-рейтингом «А» для масштабирования инноваций;", size=14, color=NAVY)]),
            paragraph([text_run("дивиденды не выплачиваются — ресурсы остаются в группе и могут быть направлены на продвижение Face2", size=14, color=NAVY)]),
        ],
        fill_color=BG_LIGHT,
    ))
    return make_slide_xml("".join(shapes))


# =====================================================================
# СЛАЙД 7 — Face2 (4 продукта + KPI)
# =====================================================================
def slide_7() -> str:
    shapes = [page_header("Задача 3. Продукт Face2 — линейка и параметры", 7)]

    # 4 продуктовые карточки
    products = [
        ("Face2Pay", "Оплата по лицу", "Банки, ритейл,\nмаркетплейсы"),
        ("Face2Pass", "СКУД", "Офисы, ВУЗы,\nзаводы"),
        ("Face2Check-in", "Возраст и личность", "Аэропорты, спорт,\nрежимные объекты"),
        ("КБС под ключ", "Полная биосистема", "Банки и\nгосорганизации"),
    ]
    card_w = emu(2.95)
    card_gap = emu(0.13)
    start_x = emu(0.45)
    for i, (name, what, segment) in enumerate(products):
        x = start_x + i * (card_w + card_gap)
        shapes.append(textbox(
            710 + i, f"Prod{i}", x, emu(1.3), card_w, emu(2.0),
            [
                paragraph([text_run(name, bold=True, size=18, color=GOLD)], align="c"),
                paragraph([text_run("─────────", size=8, color=WHITE)], align="c"),
                paragraph([text_run(what, bold=True, size=14, color=WHITE)], align="c"),
                paragraph([text_run("", size=4)]),
                paragraph([text_run(segment.split("\n")[0], size=11, color=WHITE)], align="c"),
                paragraph([text_run(segment.split("\n")[1] if "\n" in segment else "", size=11, color=WHITE)], align="c"),
            ],
            fill_color=NAVY,
        ))

    # KPI продукта — 6 цифр в ряд
    shapes.append(textbox(
        720, "KPITitle", emu(0.45), emu(3.55), emu(12.3), emu(0.4),
        [paragraph([text_run("ПАРАМЕТРЫ ПРОДУКТА", bold=True, size=14, color=GOLD)])],
    ))
    kpis = [
        ("99,5 %", "точность"),
        ("2 сек", "сканирование"),
        ("№ 4-А", "аккредитация"),
        ("5 000+", "сотрудников"),
        ("200", "биотерминалов"),
        ("15 000", "макс. векторов"),
    ]
    kpi_w = emu(2.0)
    kpi_gap = emu(0.06)
    start_x = emu(0.45)
    for i, (num, lbl) in enumerate(kpis):
        x = start_x + i * (kpi_w + kpi_gap)
        shapes.append(textbox(
            721 + i, f"KP{i}", x, emu(4.0), kpi_w, emu(1.5),
            [
                paragraph([text_run(num, bold=True, size=28, color=GOLD)], align="c"),
                paragraph([text_run(lbl, size=11, color=NAVY)], align="c"),
            ],
            fill_color=BG_LIGHT,
            line_color=GOLD,
            line_width=12700,
        ))

    # Партнёры
    shapes.append(textbox(
        730, "Partners", emu(0.45), emu(5.7), emu(12.3), emu(1.2),
        [
            paragraph([text_run("ПАРТНЁРЫ ФЕДЕРАЛЬНОГО УРОВНЯ", bold=True, size=14, color=GOLD)]),
            paragraph([text_run("Минцифры России · Минтранс России · НСПК · Центр биометрических технологий ·", size=13, color=NAVY)]),
            paragraph([text_run("Казанский метрополитен · Национальная библиотека Республики Татарстан", size=13, color=NAVY)]),
        ],
        fill_color=BG_LIGHT,
    ))
    return make_slide_xml("".join(shapes))


# =====================================================================
# СЛАЙД 8 — SWOT (квадрант 2×2)
# =====================================================================
def slide_8() -> str:
    shapes = [page_header("Задача 4. SWOT-анализ продвижения Face2", 8)]

    quadrants = [
        # (label, items, fill_color, name_color, x, y)
        ("STRENGTHS — СИЛЬНЫЕ СТОРОНЫ",
         ["Аккредитация КБС № 4-А", "Сертификация ФСБ / ФСТЭК", "Точность 99,5 %", "Рейтинг «A» и ESG-A", "Партнёры федерального уровня"],
         LIGHT_GREEN, "006100", emu(0.45), emu(1.3)),
        ("WEAKNESSES — СЛАБЫЕ СТОРОНЫ",
         ["Низкая узнаваемость вне РТ", "Слабое присутствие в деловых СМИ", "Нет публичной NIST-валидации", "Узкая партнёрская сеть"],
         LIGHT_YELLOW, "806000", emu(6.65), emu(1.3)),
        ("OPPORTUNITIES — ВОЗМОЖНОСТИ",
         ["Рынок × 9 за год", "Регуляторная зачистка локальных БС", "Запрос на отечественное ПО в B2G", "Новые сценарии (отель, возраст)"],
         LIGHT_BLUE, "1F3864", emu(0.45), emu(4.25)),
        ("THREATS — УГРОЗЫ",
         ["Сбер — 700 тыс. терминалов", "Конкуренция VL · NtechLab · OVISION", "B2B-цикл сделки 6–12 мес.", "Репутационные риски от инцидентов"],
         LIGHT_RED, "8B1A1A", emu(6.65), emu(4.25)),
    ]
    for i, (label, items, fill, color, x, y) in enumerate(quadrants):
        paras = [paragraph([text_run(label, bold=True, size=14, color=color)])]
        for it in items:
            paras.append(paragraph([text_run("• ", size=13, color=color), text_run(it, size=13, color=NAVY)]))
        shapes.append(textbox(
            810 + i, f"SWOT{i}", x, y, emu(6.1), emu(2.85), paras,
            fill_color=fill,
        ))

    # Вывод
    shapes.append(textbox(
        820, "SwotResult", emu(0.45), emu(7.18), emu(12.3), emu(0.0),
        [],
    ))
    # Используем штатное место под source
    shapes.append(textbox(
        821, "SwotConclusion", emu(0.45), emu(7.0), emu(11.0), emu(0.4),
        [paragraph([text_run("ВЫВОД: технологический каркас укомплектован — коммуникационная составляющая отстаёт (управляемая проблема)", bold=True, size=12, color=NAVY)])],
    ))
    return make_slide_xml("".join(shapes))


# =====================================================================
# СЛАЙД 9 — Стратегия и эффект
# =====================================================================
def slide_9() -> str:
    shapes = [page_header("Задача 5. Стратегия продвижения Face2 на 2026–2028 гг.", 9)]

    # Левая колонка — каркас
    shapes.append(textbox(
        910, "Frame", emu(0.45), emu(1.3), emu(7.5), emu(3.5),
        [
            paragraph([text_run("СЕГМЕНТЫ (4):", bold=True, size=13, color=GOLD)]),
            paragraph([text_run("• Банки и МФО", size=12, color=NAVY)]),
            paragraph([text_run("• Ритейл и транспорт", size=12, color=NAVY)]),
            paragraph([text_run("• Государственные организации", size=12, color=NAVY)]),
            paragraph([text_run("• B2C-аудитория Татарстана", size=12, color=NAVY)]),
            paragraph([text_run("", size=6)]),
            paragraph([text_run("ПОЗИЦИОНИРОВАНИЕ:", bold=True, size=13, color=GOLD)]),
            paragraph([text_run("«Аккредитованная КБС с соответствием 572-ФЗ из коробки»", bold=True, size=13, color=NAVY)]),
            paragraph([text_run("", size=6)]),
            paragraph([text_run("СТРУКТУРА БЮДЖЕТА — 186 млн ₽", bold=True, size=13, color=GOLD)]),
            paragraph([text_run("• 40 % — формирование доверия", bold=True, size=12, color=NAVY)]),
            paragraph([text_run("• 30 % — B2B-продажи и партнёрство", size=12, color=NAVY)]),
            paragraph([text_run("• 20 % — демозоны и пилоты", size=12, color=NAVY)]),
            paragraph([text_run("• 10 % — программы ранних пользователей", size=12, color=NAVY)]),
        ],
        fill_color=BG_LIGHT,
    ))

    # Правая колонка — эффект (4 плитки 2x2)
    shapes.append(textbox(
        920, "EffectTitle", emu(8.2), emu(1.3), emu(4.6), emu(0.4),
        [paragraph([text_run("ЭКОНОМИЧЕСКИЙ ЭФФЕКТ ЗА 3 ГОДА", bold=True, size=12, color=GOLD)])],
    ))
    eff = [
        ("832", "млн ₽\nвыручка"),
        ("355", "млн ₽\nмарж. эффект"),
        ("180 %", "ROI"),
        ("18", "мес.\nокупаемость"),
    ]
    eff_w = emu(2.2)
    eff_h = emu(1.5)
    for i, (num, lbl) in enumerate(eff):
        col = i % 2
        row = i // 2
        x = emu(8.2) + col * (eff_w + emu(0.2))
        y = emu(1.8) + row * (eff_h + emu(0.2))
        shapes.append(textbox(
            921 + i, f"Eff{i}", x, y, eff_w, eff_h,
            [
                paragraph([text_run(num, bold=True, size=32, color=GOLD)], align="c"),
                paragraph([text_run(lbl.split("\n")[0], size=11, color=WHITE)], align="c"),
                paragraph([text_run(lbl.split("\n")[1] if "\n" in lbl else "", size=11, color=WHITE)], align="c"),
            ],
            fill_color=NAVY,
        ))

    # Дорожная карта (внизу)
    shapes.append(textbox(
        930, "RoadmapTitle", emu(0.45), emu(5.0), emu(12.3), emu(0.4),
        [paragraph([text_run("ДОРОЖНАЯ КАРТА", bold=True, size=13, color=GOLD)])],
    ))
    road = [
        ("2026 H1", "Подготовительный", NAVY),
        ("2026 H2 – 2027 H1", "Пилотный (8–12 объектов)", GOLD),
        ("2027 H2 – 2028 H2", "Масштабирование (35–45)", GOLD),
        ("2028 H2", "Закрепление (≥ 60 внедрений)", NAVY),
    ]
    rw = emu(3.0)
    rg = emu(0.1)
    for i, (period, name, color) in enumerate(road):
        x = emu(0.45) + i * (rw + rg)
        shapes.append(textbox(
            931 + i, f"Road{i}", x, emu(5.5), rw, emu(1.0),
            [
                paragraph([text_run(period, bold=True, size=12, color=GOLD if color == NAVY else WHITE)], align="c"),
                paragraph([text_run(name, size=11, color=WHITE)], align="c"),
            ],
            fill_color=color,
        ))

    # Дисклеймер
    shapes.append(textbox(
        940, "Disclaimer", emu(0.45), emu(6.7), emu(12.3), emu(0.4),
        [paragraph([text_run("Прогнозные оценки автора; ставка дисконтирования 18 %; уточняются с финансовым департаментом банка", size=10, color=GREY)])],
    ))
    return make_slide_xml("".join(shapes))


# =====================================================================
# СЛАЙД 10 — Заключение
# =====================================================================
def slide_10() -> str:
    shapes = [page_header("Заключение", 10)]

    # Главная плашка
    shapes.append(textbox(
        1010, "MainResult", emu(2.0), emu(1.5), emu(9.3), emu(1.2),
        [
            paragraph([text_run("✓ Цель работы достигнута", bold=True, size=22, color=WHITE)], align="c"),
            paragraph([text_run("✓ Поставленные задачи решены", bold=True, size=22, color=GOLD)], align="c"),
        ],
        fill_color=NAVY,
    ))

    # Практическая значимость
    shapes.append(textbox(
        1020, "SignificanceTitle", emu(0.5), emu(3.0), emu(12.3), emu(0.5),
        [paragraph([text_run("ПРАКТИЧЕСКАЯ ЗНАЧИМОСТЬ", bold=True, size=18, color=GOLD)], align="c")],
    ))
    items = [
        ("📋", "Стратегия продвижения Face2 на 2026–2028 гг."),
        ("📊", "Система KPI и расчёт экономического эффекта (ROI 180 %)"),
        ("⚠", "Матрица рисков с мерами их снижения"),
        ("👥", "Организационный план внедрения мероприятий"),
    ]
    for i, (icon, text) in enumerate(items):
        shapes.append(textbox(
            1021 + i, f"Sig{i}", emu(2.0), emu(3.6) + i * emu(0.5), emu(9.3), emu(0.45),
            [paragraph([text_run(f"{icon}  ", size=15), text_run(text, size=15, color=NAVY)])],
        ))

    # «Спасибо за внимание»
    shapes.append(textbox(
        1030, "Thanks", emu(0.5), emu(6.0), emu(12.3), emu(1.0),
        [
            paragraph([text_run("СПАСИБО ЗА ВНИМАНИЕ", bold=True, size=40, color=GOLD)], align="c"),
            paragraph([text_run("Готова ответить на вопросы комиссии", size=14, color=NAVY)], align="c"),
        ],
    ))
    return make_slide_xml("".join(shapes))


# =====================================================================
# Сборка пакета OOXML
# =====================================================================
SLIDE_BUILDERS = [slide_1, slide_2, slide_3, slide_4, slide_5,
                  slide_6, slide_7, slide_8, slide_9, slide_10]


def make_content_types() -> str:
    overrides = "".join(
        f'<Override PartName="/ppt/slides/slide{i+1}.xml" '
        f'ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>'
        for i in range(len(SLIDE_BUILDERS))
    )
    return f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>
<Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/>
<Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/>
<Override PartName="/ppt/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>
<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
{overrides}
</Types>"""


ROOT_RELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/>
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
</Relationships>"""


CORE_PROPS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
xmlns:dc="http://purl.org/dc/elements/1.1/"
xmlns:dcterms="http://purl.org/dc/terms/"
xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
<dc:title>ВКР: Стратегия продвижения Face2 АК Барс Банк</dc:title>
<dc:creator>Студент(ка) ВКР</dc:creator>
<cp:lastModifiedBy>VKR Author</cp:lastModifiedBy>
<dcterms:created xsi:type="dcterms:W3CDTF">2026-05-24T12:00:00Z</dcterms:created>
<dcterms:modified xsi:type="dcterms:W3CDTF">2026-05-24T12:00:00Z</dcterms:modified>
</cp:coreProperties>"""


APP_PROPS = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties">
<TotalTime>0</TotalTime>
<Words>0</Words>
<Application>VKR Builder</Application>
<Slides>{len(SLIDE_BUILDERS)}</Slides>
<Notes>0</Notes>
<HiddenSlides>0</HiddenSlides>
<MMClips>0</MMClips>
</Properties>"""


def make_presentation_xml() -> str:
    sld_id_list = "".join(
        f'<p:sldId id="{256 + i}" r:id="rId{i+2}"/>'
        for i in range(len(SLIDE_BUILDERS))
    )
    return f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:presentation xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
saveSubsetFonts="1">
<p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/></p:sldMasterIdLst>
<p:sldIdLst>{sld_id_list}</p:sldIdLst>
<p:sldSz cx="{SLIDE_W}" cy="{SLIDE_H}" type="screen16x9"/>
<p:notesSz cx="6858000" cy="9144000"/>
<p:defaultTextStyle/>
</p:presentation>"""


def make_presentation_rels() -> str:
    rels = ['<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="slideMasters/slideMaster1.xml"/>']
    for i in range(len(SLIDE_BUILDERS)):
        rels.append(
            f'<Relationship Id="rId{i+2}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide{i+1}.xml"/>'
        )
    rels_xml = "".join(rels)
    return f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">{rels_xml}</Relationships>"""


SLIDE_MASTER_XML = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldMaster xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
<p:cSld>
<p:bg><p:bgRef idx="1001"><a:schemeClr val="bg1"/></p:bgRef></p:bg>
<p:spTree>
<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>
</p:spTree>
</p:cSld>
<p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>
<p:sldLayoutIdLst><p:sldLayoutId id="2147483649" r:id="rId1"/></p:sldLayoutIdLst>
<p:txStyles>
<p:titleStyle><a:lvl1pPr><a:defRPr sz="4400" b="1"><a:solidFill><a:srgbClr val="{NAVY}"/></a:solidFill><a:latin typeface="Calibri"/></a:defRPr></a:lvl1pPr></p:titleStyle>
<p:bodyStyle><a:lvl1pPr><a:defRPr sz="1800"><a:solidFill><a:srgbClr val="{NAVY}"/></a:solidFill><a:latin typeface="Calibri"/></a:defRPr></a:lvl1pPr></p:bodyStyle>
<p:otherStyle><a:defPPr><a:defRPr lang="ru-RU"/></a:defPPr></p:otherStyle>
</p:txStyles>
</p:sldMaster>"""


SLIDE_MASTER_RELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="../theme/theme1.xml"/>
</Relationships>"""


SLIDE_LAYOUT_XML = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldLayout xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
type="blank" preserve="1">
<p:cSld name="Blank">
<p:spTree>
<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>
</p:spTree>
</p:cSld>
<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>
</p:sldLayout>"""


SLIDE_LAYOUT_RELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="../slideMasters/slideMaster1.xml"/>
</Relationships>"""


THEME_XML = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="VKR Theme">
<a:themeElements>
<a:clrScheme name="VKR">
<a:dk1><a:srgbClr val="000000"/></a:dk1>
<a:lt1><a:srgbClr val="FFFFFF"/></a:lt1>
<a:dk2><a:srgbClr val="{NAVY}"/></a:dk2>
<a:lt2><a:srgbClr val="EEECE1"/></a:lt2>
<a:accent1><a:srgbClr val="{NAVY}"/></a:accent1>
<a:accent2><a:srgbClr val="{GOLD}"/></a:accent2>
<a:accent3><a:srgbClr val="{GREEN}"/></a:accent3>
<a:accent4><a:srgbClr val="{RED}"/></a:accent4>
<a:accent5><a:srgbClr val="{LIGHT_BLUE}"/></a:accent5>
<a:accent6><a:srgbClr val="{LIGHT_GREEN}"/></a:accent6>
<a:hlink><a:srgbClr val="0000FF"/></a:hlink>
<a:folHlink><a:srgbClr val="800080"/></a:folHlink>
</a:clrScheme>
<a:fontScheme name="VKR">
<a:majorFont><a:latin typeface="Calibri"/><a:ea typeface=""/><a:cs typeface=""/></a:majorFont>
<a:minorFont><a:latin typeface="Calibri"/><a:ea typeface=""/><a:cs typeface=""/></a:minorFont>
</a:fontScheme>
<a:fmtScheme name="VKR">
<a:fillStyleLst>
<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
</a:fillStyleLst>
<a:lnStyleLst>
<a:ln w="9525"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/></a:ln>
<a:ln w="25400"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/></a:ln>
<a:ln w="38100"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/></a:ln>
</a:lnStyleLst>
<a:effectStyleLst>
<a:effectStyle><a:effectLst/></a:effectStyle>
<a:effectStyle><a:effectLst/></a:effectStyle>
<a:effectStyle><a:effectLst/></a:effectStyle>
</a:effectStyleLst>
<a:bgFillStyleLst>
<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
</a:bgFillStyleLst>
</a:fmtScheme>
</a:themeElements>
</a:theme>"""


SLIDE_RELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>
</Relationships>"""


def build_pptx(out_path: str) -> None:
    files: List[Tuple[str, str]] = [
        ("[Content_Types].xml", make_content_types()),
        ("_rels/.rels", ROOT_RELS),
        ("docProps/core.xml", CORE_PROPS),
        ("docProps/app.xml", APP_PROPS),
        ("ppt/presentation.xml", make_presentation_xml()),
        ("ppt/_rels/presentation.xml.rels", make_presentation_rels()),
        ("ppt/theme/theme1.xml", THEME_XML),
        ("ppt/slideMasters/slideMaster1.xml", SLIDE_MASTER_XML),
        ("ppt/slideMasters/_rels/slideMaster1.xml.rels", SLIDE_MASTER_RELS),
        ("ppt/slideLayouts/slideLayout1.xml", SLIDE_LAYOUT_XML),
        ("ppt/slideLayouts/_rels/slideLayout1.xml.rels", SLIDE_LAYOUT_RELS),
    ]

    for i, builder in enumerate(SLIDE_BUILDERS, start=1):
        files.append((f"ppt/slides/slide{i}.xml", builder()))
        files.append((f"ppt/slides/_rels/slide{i}.xml.rels", SLIDE_RELS))

    os.makedirs(os.path.dirname(out_path), exist_ok=True)
    with zipfile.ZipFile(out_path, "w", zipfile.ZIP_DEFLATED) as zf:
        for path, content in files:
            zf.writestr(path, content)
    print(f"OK: {out_path} ({os.path.getsize(out_path) // 1024} KB, {len(SLIDE_BUILDERS)} slides)")


if __name__ == "__main__":
    build_pptx(OUT)
