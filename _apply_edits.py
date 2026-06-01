#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Применить три правки к «ВКР ГОТОВОЕ.md»:

1. Удалить преамбулу до «# ВВЕДЕНИЕ» (служебное описание структуры/оформления).
2. Извлечь 7 блоков рисунков (ASCII-схема + Источник + подпись) в файл
   «Рисунки.md» и удалить их из основного файла. Текстовые отсылки
   «...на рисунке N.N.N» в абзацах сохраняются.
3. Перевернуть три широкие таблицы (2.2.5, 3.3.1, 3.3.6) в более узкий
   формат, чтобы они помещались по ширине страницы A4 (поля 30/10/20/20 мм).
"""

from __future__ import annotations

import re
from pathlib import Path

HERE = Path(__file__).parent
SRC = HERE / "ВКР ГОТОВОЕ.md"
DST_FIG = HERE / "Рисунки.md"

# ---------------------------------------------------------------------------
# 0. Чтение
# ---------------------------------------------------------------------------

text = SRC.read_text(encoding="utf-8")
lines = text.split("\n")

# ---------------------------------------------------------------------------
# 1. Удалить преамбулу до «# ВВЕДЕНИЕ»
# ---------------------------------------------------------------------------

i_intro = next(i for i, l in enumerate(lines) if l.strip() == "# ВВЕДЕНИЕ")
removed_preamble = lines[:i_intro]
lines = lines[i_intro:]
print(f"[1] Удалена преамбула: {len(removed_preamble)} строк до «# ВВЕДЕНИЕ»")

# ---------------------------------------------------------------------------
# 2. Извлечь блоки рисунков
# ---------------------------------------------------------------------------

CAPTION_RE = re.compile(r"^Рис\.\s+(\d+\.\d+\.\d+)\.\s+(.+)$")
SOURCE_RE = re.compile(r"^Источник:")
# Символы box-drawing/блочные — индикаторы ASCII-схемы.
ART_CHARS = set("┌┐└┘─│├┤┬┴┼▼▲▶◄►◀█▌▐░▒▓")
# Префиксы строк, которые НЕ относятся к содержимому блока рисунка.
NON_FIG_PREFIXES = ("|", "Таблица", "# ", "## ", "### ", "- ")
# Порог длины: строки длиннее считаются обычным текстовым абзацем.
TEXT_LINE_THRESHOLD = 200


def is_figure_content(line: str) -> bool:
    """Строка является частью блока рисунка."""
    s = line.strip()
    if not s:
        return True
    if SOURCE_RE.match(s):
        return True
    if any(c in ART_CHARS for c in s):
        return True
    # После нормализации пробелов в _docx_extract.py подписи к схемам
    # выглядят как короткие строки без структурных префиксов.
    if len(s) < TEXT_LINE_THRESHOLD and not s.startswith(NON_FIG_PREFIXES):
        return True
    return False


# Найти все строки-подписи (Рис. N.N.N. ...).
caption_indices = [
    i for i, l in enumerate(lines) if CAPTION_RE.match(l.strip())
]
print(f"[2] Найдено подписей рисунков: {len(caption_indices)}")

# Для каждой подписи определим границы блока.
figure_blocks: list[tuple[int, int, str, str]] = []
for cap_i in caption_indices:
    cap = lines[cap_i].strip()
    m = CAPTION_RE.match(cap)
    num, name = m.group(1), m.group(2)

    j = cap_i - 1
    while j >= 0 and is_figure_content(lines[j]):
        j -= 1
    start = j + 1
    figure_blocks.append((start, cap_i, num, name))

# Построить Рисунки.md.
fig_doc: list[str] = ["# РИСУНКИ", ""]
for start, end, num, name in figure_blocks:
    fig_doc.append(f"## Рис. {num}. {name}")
    fig_doc.append("")
    block_lines = lines[start:end + 1]
    while block_lines and not block_lines[0].strip():
        block_lines.pop(0)
    while block_lines and not block_lines[-1].strip():
        block_lines.pop()
    fig_doc.extend(block_lines)
    fig_doc.append("")
    fig_doc.append("---")
    fig_doc.append("")

DST_FIG.write_text("\n".join(fig_doc).rstrip() + "\n", encoding="utf-8")
print(f"[2] Записан файл: {DST_FIG.name} "
      f"({DST_FIG.stat().st_size:,} байт)")

# Удалить блоки из исходного md (с конца к началу, чтобы индексы не плыли).
removed_total = 0
for start, end, num, name in reversed(figure_blocks):
    # Расширим end до пустой строки после подписи (если есть).
    extended_end = end
    if extended_end + 1 < len(lines) and lines[extended_end + 1].strip() == "":
        extended_end += 1
    removed_total += extended_end - start + 1
    del lines[start:extended_end + 1]

print(f"[2] Удалено из ВКР ГОТОВОЕ.md: {removed_total} строк (7 блоков)")

# ---------------------------------------------------------------------------
# 3. Перестроить таблицы 2.2.5, 3.3.1, 3.3.6 в длинный формат
# ---------------------------------------------------------------------------

text = "\n".join(lines)


def replace_table(text: str, anchor_caption: str, new_table: str) -> str:
    idx = text.find(anchor_caption)
    if idx == -1:
        raise SystemExit(f"Не найдена якорная подпись: {anchor_caption!r}")

    after = idx
    nl = text.find("\n", after)
    tbl_start = None
    while nl != -1:
        line_start = nl + 1
        line_end = text.find("\n", line_start)
        if line_end == -1:
            line_end = len(text)
        line = text[line_start:line_end]
        if line.lstrip().startswith("|"):
            tbl_start = line_start
            break
        nl = line_end
    if tbl_start is None:
        raise SystemExit(f"Таблица не найдена после {anchor_caption!r}")

    pos = tbl_start
    while True:
        ln_end = text.find("\n", pos)
        if ln_end == -1:
            ln_end = len(text)
        line = text[pos:ln_end]
        if line.lstrip().startswith("|"):
            pos = ln_end + 1 if ln_end != len(text) else ln_end
            if pos >= len(text):
                tbl_end = pos
                break
        else:
            tbl_end = pos
            break

    replacement = new_table.rstrip() + "\n"
    return text[:tbl_start] + replacement + text[tbl_end:]


# ---------- 2.2.5 ----------
TBL_225_PARAMS = [
    ("Тип игрока", "Банк-разработчик", "Банк-разработчик",
     "Технологический вендор", "Производитель полного цикла"),
    ("Аккредитация Минцифры (КБС)", "№ 4-А от 15.11.2023 [27]",
     "Через дочерние структуры", "Партнёр ЦБТ [40]", "Есть [41]"),
    ("Право регистрации БПд в ЕБС (115-ФЗ)",
     "№ 101 в перечне [35]", "Есть", "Нет", "Нет"),
    ("Реестр отечественного ПО",
     "Есть [27]", "Есть", "Есть [40]", "Есть [41]"),
    ("ФСБ и ФСТЭК",
     "Есть [27]", "Есть", "Через ПАК «ВЛ-БИО» [40]", "Есть, СКЗИ [41]"),
    ("Независимая верификация (NIST/iBeta/CVPR)",
     "Не зафиксирована", "Не зафиксирована",
     "Подтверждена [40]", "Не зафиксирована"),
    ("Заявленная точность",
     "99,5% [27]", "Косвенно через VisionLabs",
     "Лидерство NIST FRTE [40]", "≈ 99,9% [41]"),
    ("Время сканирования",
     "≈ 2 с [27]", "6 с (цель 2 с к 2025) [38]",
     "≈ 2,5 с (LUNA)", "≈ 0,3 с [41]"),
    ("Производство терминалов",
     "Партнёрское", "Нет", "Нет (модуль ПАК)", "Есть (OGATE) [41]"),
    ("Продуктовая линейка",
     "Face2Pay/Pass/Check-in + КБС — по банковским сценариям",
     "СберPay Smile — преимущественно эквайринг",
     "LUNA Platform/Cars/ID/Pass/Kiosk + Deepfake Detection — по форм-факторам",
     "OGATE + OGATE Control + КБС OVISION"),
    ("География",
     "Татарстан + федер. пилоты (Минцифры, Минтранс, НСПК)",
     "Вся РФ и новые территории [38]",
     "38 стран, 9 представительств [40]", "Вся РФ"),
    ("Антифрод (дипфейки)",
     "Отдельное решение публично не зафиксировано",
     "Liveness + AI-фрод + ПИН [38]",
     "Продукт Deepfake Detection [40]", "Антиспуфинг + дипфейки [41]"),
    ("Банковская/инфраструктурная экосистема",
     "АББ, Карта жителя РТ, СБП, АБО [26]",
     "Сбер, СберPay, СберID",
     "ЦБТ, ЦРТ, Нетрис, ITV, ВК ИТС, Macroscop [40]",
     "Без банковской экосистемы"),
]
PLAYERS_225 = ["Face2 (АК БАРС БАНК)", "Сбер", "VisionLabs", "OVISION"]


def build_table_225() -> str:
    out = ["| Параметр | Игрок | Значение |", "| --- | --- | --- |"]
    for row in TBL_225_PARAMS:
        param = row[0]
        for j, player in enumerate(PLAYERS_225):
            value = row[j + 1]
            out.append(f"| {param} | {player} | {value} |")
    return "\n".join(out)


text = replace_table(text, "Таблица 2.2.5", build_table_225())
print("[3a] Таблица 2.2.5: 13×5 → 52×3 (длинный формат)")


# ---------- 3.3.1 ----------
TBL_331_DIRS = [
    "А. B2B длинного цикла (уровень 1)",
    "Б. B2B/B2G короткого цикла (уровень 2)",
    "В. B2C-слой",
]
TBL_331_STAGES = [
    "Осведомлённость", "Интерес", "Доверие", "Пилот", "Принятие", "Лояльность",
]
TBL_331_DATA = [
    ("Цитируемость в деловых СМИ; число публикаций экспертного контента",
     "Число входящих RFP; число запросов на демо",
     "Число подтверждённых референсов; результаты NIST FRTE и iBeta Level 2",
     "Число активных пилотов; средняя длительность пилота",
     "Число подписанных контрактов уровня 1; выручка КБС под ключ",
     "Доля контрактов с расширением модулей; NPS клиентов уровня 1"),
    ("Уникальные посетители face2.ru; охват контекстной рекламы",
     "Заявки через CTA-механику; загрузки типовых КП по шести сегментам",
     "Число просмотров кейсового и регуляторного контента; CTR доказательного контента",
     "Число подключений в пилотном статусе",
     "Число коммерческих подключений; выручка Face2Pay, Face2Pass, Face2Check-in",
     "Доля повторных продаж модулей; LTV подключённого объекта"),
    ("Регистрации в ЕБС через приложение АБО АББ; охват в «Карте жителя РТ»",
     "Активации Face2Pay в метрополитене; число первых транзакций",
     "Доля повторных транзакций; снижение оттока после первой неудачной попытки",
     "— (этап не применим к B2C)",
     "Доля активных пользователей; средняя частота транзакций на пользователя",
     "NPS B2C; доля рекомендующих"),
]


def build_table_331() -> str:
    out = ["| Направление | Этап воронки | KPI |",
           "| --- | --- | --- |"]
    for di, dname in enumerate(TBL_331_DIRS):
        for si, sname in enumerate(TBL_331_STAGES):
            out.append(f"| {dname} | {sname} | {TBL_331_DATA[di][si]} |")
    return "\n".join(out)


text = replace_table(text, "Таблица 3.3.1", build_table_331())
print("[3b] Таблица 3.3.1: 3×7 → 18×3 (длинный формат)")


# ---------- 3.3.6 ----------
TBL_336_ACTS = [
    "Переоснащение face2.ru под CTA-механику",
    "Контентный конвейер (3 типа контента)",
    "Заявки на NIST FRTE и iBeta Level 2",
    "Участие в отраслевых конференциях",
    "Развитие партнёрской сети интеграторов и вендоров",
    "Пилоты уровня 1 (КБС под ключ)",
    "Подключения объектов уровня 2 (Face2Pay/Pass/Check-in)",
    "Регуляторные обзоры в контенте",
    "PR-размещения в деловых федеральных СМИ",
    "Референсная программа фазы 3",
]
TBL_336_ROLES = [
    "Маркетинговая служба",
    "Команда Face2",
    "PR-функция",
    "Юридическая служба",
    "CISO",
    "Продажи B2B",
    "Партнёрское подразделение",
    "Внешние партнёры",
]
TBL_336_DATA = [
    ["A", "R", "C", "C", "I", "I", "I", "R (агентство)"],
    ["A", "C", "R", "C", "C", "I", "I", "R (студия)"],
    ["I", "A/R", "C", "C", "C", "I", "I", "R (лаборатория)"],
    ["A", "C", "R", "I", "C", "C", "C", "R (организаторы)"],
    ["C", "C", "I", "I", "I", "C", "A/R", "I"],
    ["C", "R", "I", "C", "C", "A", "C", "I"],
    ["A", "R", "I", "I", "I", "C", "R", "I"],
    ["C", "I", "C", "A/R", "C", "I", "I", "I"],
    ["A", "C", "R", "C", "C", "I", "I", "R (агентство)"],
    ["A", "R", "C", "I", "I", "C", "C", "I"],
]


def build_table_336() -> str:
    out = ["| Мероприятие | Роль | RACI |", "| --- | --- | --- |"]
    for ai, act in enumerate(TBL_336_ACTS):
        for ri, role in enumerate(TBL_336_ROLES):
            out.append(f"| {act} | {role} | {TBL_336_DATA[ai][ri]} |")
    return "\n".join(out)


text = replace_table(text, "Таблица 3.3.6", build_table_336())
print("[3c] Таблица 3.3.6: 10×9 → 80×3 (длинный формат)")


# ---------------------------------------------------------------------------
# 4. Сохранение
# ---------------------------------------------------------------------------

text = re.sub(r"\n{3,}", "\n\n", text)
if not text.endswith("\n"):
    text += "\n"

SRC.write_text(text, encoding="utf-8")
n_lines = text.count("\n") + 1
print(f"\n[итог] {SRC.name}: {SRC.stat().st_size:,} байт, {n_lines} строк")
print(f"[итог] {DST_FIG.name}: {DST_FIG.stat().st_size:,} байт")
