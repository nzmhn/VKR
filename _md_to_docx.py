#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Конвертер markdown → Word (.docx) для файла «ВКР ГОТОВОЕ.md».

Поддерживаемые элементы markdown:
  - заголовки уровней # и ##;
  - обычные абзацы с инлайн-форматированием **bold** и *italic*;
  - маркированные списки (- ...);
  - цитаты (> ...);
  - таблицы с шапкой;
  - блоки кода ``` ``` (отображаются моноширинно с сохранением пробелов);
  - горизонтальные линии --- (как разрыв страницы между главами).

Скрипт собирает .docx как ZIP с OOXML без внешних зависимостей
(python-docx и pandoc недоступны).
Оформление: Times New Roman 14 пт, междострочный 1,5, абзацный отступ 1,25 см,
поля 30/10/20/20 мм, формат A4. Таблицы — TNR 12 пт, одинарный интервал.
"""

from __future__ import annotations

import re
import zipfile
from html import escape as _xe
from pathlib import Path

# ---------------------------------------------------------------------------
# Константы геометрии и стиля (как в _build_chapter2_docx.py)
# ---------------------------------------------------------------------------

TWIPS_PER_CM = 567

PAGE_W_TW = 11906
PAGE_H_TW = 16838
MARGIN_LEFT_TW = 30 * TWIPS_PER_CM // 10
MARGIN_RIGHT_TW = 10 * TWIPS_PER_CM // 10
MARGIN_TOP_TW = 20 * TWIPS_PER_CM // 10
MARGIN_BOTTOM_TW = 20 * TWIPS_PER_CM // 10

CONTENT_W_TW = PAGE_W_TW - MARGIN_LEFT_TW - MARGIN_RIGHT_TW
CONTENT_W_CM = CONTENT_W_TW / TWIPS_PER_CM

# English Metric Units для DrawingML: 1 см = 360 000 EMU.
EMU_PER_CM = 360000
CONTENT_W_EMU = int(CONTENT_W_CM * EMU_PER_CM)
# Максимальная высота рисунка, чтобы он помещался на странице A4.
MAX_FIG_H_CM = 20.0

FONT_MAIN = "Times New Roman"
FONT_MONO = "Courier New"
SZ_BODY_HP = 28        # 14 pt
SZ_TABLE_HP = 24       # 12 pt
SZ_MONO_HP = 18        # 9 pt — для ASCII-схем

CLR_BLACK = "000000"
CLR_LIGHT_GRAY = "F2F2F2"

INDENT_FIRST_TW = int(1.25 * TWIPS_PER_CM)
INDENT_LIST_TW = int(0.75 * TWIPS_PER_CM)

LINE_BODY = 360       # 1.5x
LINE_SINGLE = 240     # 1.0x

SPC_BEFORE_HEAD = 240
SPC_AFTER_HEAD = 120


# ---------------------------------------------------------------------------
# OOXML-хелперы
# ---------------------------------------------------------------------------

def xe(text: str) -> str:
    if text is None:
        return ""
    return _xe(text, quote=False)


def run(text: str, *, sz_hp: int = SZ_BODY_HP, bold: bool = False,
        italic: bool = False, color: str = CLR_BLACK,
        font: str = FONT_MAIN, preserve_space: bool = True,
        vert_align: str | None = None) -> str:
    rpr = ['<w:rPr>']
    rpr.append(
        f'<w:rFonts w:ascii="{font}" w:hAnsi="{font}" '
        f'w:cs="{font}" w:eastAsia="{font}"/>'
    )
    if bold:
        rpr.append('<w:b/><w:bCs/>')
    if italic:
        rpr.append('<w:i/><w:iCs/>')
    rpr.append(f'<w:color w:val="{color}"/>')
    rpr.append(f'<w:sz w:val="{sz_hp}"/><w:szCs w:val="{sz_hp}"/>')
    if vert_align in ("subscript", "superscript"):
        rpr.append(f'<w:vertAlign w:val="{vert_align}"/>')
    rpr.append('<w:lang w:val="ru-RU"/>')
    rpr.append('</w:rPr>')
    space = ' xml:space="preserve"' if preserve_space else ""
    return (
        f'<w:r>{"".join(rpr)}'
        f'<w:t{space}>{xe(text)}</w:t></w:r>'
    )


def tab_run() -> str:
    return '<w:r><w:tab/></w:r>'


# Инлайн-разметка формул: _{...} и ^{...} (или односимвольные _x / ^x)
# превращаются в подстрочные/надстрочные индексы.
_SUBSUP_RE = re.compile(r"(_\{[^}]*\}|\^\{[^}]*\}|_[^\s_^]|\^[^\s_^])")


def formula_runs(text: str, *, sz_hp: int = SZ_BODY_HP,
                 italic: bool = False) -> str:
    """Собрать runs для текста формулы с поддержкой индексов и степеней."""
    out: list[str] = []
    pos = 0
    for m in _SUBSUP_RE.finditer(text):
        if m.start() > pos:
            out.append(run(text[pos:m.start()], sz_hp=sz_hp, italic=italic))
        tok = m.group(0)
        kind = "subscript" if tok[0] == "_" else "superscript"
        inner = tok[2:-1] if tok[1] == "{" else tok[1:]
        out.append(run(inner, sz_hp=sz_hp, italic=italic, vert_align=kind))
        pos = m.end()
    if pos < len(text):
        out.append(run(text[pos:], sz_hp=sz_hp, italic=italic))
    if not out:
        out.append(run(text, sz_hp=sz_hp, italic=italic))
    return "".join(out)


def formula_paragraph(eq: str, number: str = "") -> str:
    """Формула отдельной строкой: по центру, номер прижат к правому полю.

    Реализовано через центрирующий и правый табуляторы (требование
    Положения: формула отдельной строкой, номер в круглых скобках справа)."""
    center_pos = CONTENT_W_TW // 2
    right_pos = CONTENT_W_TW
    ppr_xml = (
        '<w:pPr>'
        f'<w:tabs>'
        f'<w:tab w:val="center" w:pos="{center_pos}"/>'
        f'<w:tab w:val="right" w:pos="{right_pos}"/>'
        '</w:tabs>'
        f'<w:spacing w:before="120" w:after="120" '
        f'w:line="{LINE_BODY}" w:lineRule="auto"/>'
        '<w:jc w:val="left"/>'
        '<w:rPr>'
        f'<w:rFonts w:ascii="{FONT_MAIN}" w:hAnsi="{FONT_MAIN}" '
        f'w:cs="{FONT_MAIN}" w:eastAsia="{FONT_MAIN}"/>'
        f'<w:sz w:val="{SZ_BODY_HP}"/><w:szCs w:val="{SZ_BODY_HP}"/>'
        '<w:lang w:val="ru-RU"/>'
        '</w:rPr>'
        '</w:pPr>'
    )
    content = tab_run() + formula_runs(eq, italic=True)
    if number:
        content += tab_run() + run(number)
    return f'<w:p>{ppr_xml}{content}</w:p>'


def where_block(entries: list[str]) -> str:
    """Блок пояснений символов: со слова «где», каждый символ — с новой строки."""
    parts: list[str] = []
    for i, entry in enumerate(entries):
        prefix = "где " if i == 0 else ""
        parts.append(
            para(
                (run(prefix) if prefix else "") + formula_runs(entry, italic=False),
                align="both", indent_first=0,
                line=LINE_BODY, line_rule="auto",
                space_before=0, space_after=0,
            )
        )
    return "".join(parts)


def ppr(*, align: str = "both", indent_first: int = 0, indent_left: int = 0,
        line: int = LINE_BODY, line_rule: str = "auto",
        space_before: int = 0, space_after: int = 0,
        keep_next: bool = False, keep_lines: bool = False) -> str:
    parts = ['<w:pPr>']
    if keep_next:
        parts.append('<w:keepNext/>')
    if keep_lines:
        parts.append('<w:keepLines/>')
    parts.append(
        f'<w:spacing w:before="{space_before}" w:after="{space_after}" '
        f'w:line="{line}" w:lineRule="{line_rule}"/>'
    )
    if indent_first or indent_left:
        attrs = []
        if indent_left:
            attrs.append(f'w:left="{indent_left}"')
        if indent_first:
            attrs.append(f'w:firstLine="{indent_first}"')
        parts.append(f'<w:ind {" ".join(attrs)}/>')
    parts.append(f'<w:jc w:val="{align}"/>')
    parts.append(
        '<w:rPr>'
        f'<w:rFonts w:ascii="{FONT_MAIN}" w:hAnsi="{FONT_MAIN}" '
        f'w:cs="{FONT_MAIN}" w:eastAsia="{FONT_MAIN}"/>'
        f'<w:sz w:val="{SZ_BODY_HP}"/><w:szCs w:val="{SZ_BODY_HP}"/>'
        '<w:lang w:val="ru-RU"/>'
        '</w:rPr>'
    )
    parts.append('</w:pPr>')
    return "".join(parts)


def para(content: str = "", *, align: str = "both",
         indent_first: int = INDENT_FIRST_TW, indent_left: int = 0,
         line: int = LINE_BODY, line_rule: str = "auto",
         space_before: int = 0, space_after: int = 0,
         keep_next: bool = False, keep_lines: bool = False) -> str:
    return (
        '<w:p>'
        f'{ppr(align=align, indent_first=indent_first, indent_left=indent_left, line=line, line_rule=line_rule, space_before=space_before, space_after=space_after, keep_next=keep_next, keep_lines=keep_lines)}'
        f'{content}'
        '</w:p>'
    )


def empty_para(line: int = LINE_BODY) -> str:
    return para("", indent_first=0, line=line)


def page_break_para() -> str:
    return (
        '<w:p>'
        f'{ppr(indent_first=0)}'
        '<w:r><w:br w:type="page"/></w:r>'
        '</w:p>'
    )


# ---------------------------------------------------------------------------
# Инлайн-форматирование
# ---------------------------------------------------------------------------

# Поддерживаем **bold**, *italic*, `code`.
_TOKEN_RE = re.compile(
    r"(\*\*[^*]+?\*\*|(?<![\w*])\*[^*\n]+?\*(?![\w*])|`[^`\n]+?`)"
)


def runs_from_text(text: str, *, sz_hp: int = SZ_BODY_HP) -> str:
    out = []
    pos = 0
    for m in _TOKEN_RE.finditer(text):
        if m.start() > pos:
            chunk = text[pos:m.start()]
            if chunk:
                out.append(run(chunk, sz_hp=sz_hp))
        tok = m.group(0)
        if tok.startswith("**"):
            out.append(run(tok[2:-2], sz_hp=sz_hp, bold=True))
        elif tok.startswith("`"):
            out.append(run(tok[1:-1], sz_hp=sz_hp, font=FONT_MONO))
        else:  # *italic*
            out.append(run(tok[1:-1], sz_hp=sz_hp, italic=True))
        pos = m.end()
    if pos < len(text):
        out.append(run(text[pos:], sz_hp=sz_hp))
    if not out:
        out.append(run(text, sz_hp=sz_hp))
    return "".join(out)


# ---------------------------------------------------------------------------
# Блочные стили
# ---------------------------------------------------------------------------

def heading_h1(text: str) -> str:
    """# заголовок: ЗАГЛАВНЫМИ, по центру, полужирный 14 пт."""
    return para(
        run(text.upper(), bold=True),
        align="center", indent_first=0,
        space_before=SPC_BEFORE_HEAD, space_after=SPC_AFTER_HEAD,
        keep_next=True, keep_lines=True,
    )


def heading_h2(text: str) -> str:
    """## заголовок: по центру, полужирный 14 пт."""
    return para(
        run(text, bold=True),
        align="center", indent_first=0,
        space_before=200, space_after=120,
        keep_next=True, keep_lines=True,
    )


def body_paragraph(text: str) -> str:
    return para(
        runs_from_text(text),
        align="both",
        indent_first=INDENT_FIRST_TW,
        line=LINE_BODY, line_rule="auto",
    )


def list_item(text: str) -> str:
    """Маркированный пункт (без спец-нумерации, простой буллет «—»)."""
    content = run("— ", sz_hp=SZ_BODY_HP) + runs_from_text(text)
    return para(
        content, align="both",
        indent_first=0, indent_left=INDENT_LIST_TW,
        line=LINE_BODY, line_rule="auto",
    )


def quote_paragraph(text: str) -> str:
    """Цитата (>): курсив, отступ слева 1,25 см."""
    return para(
        runs_from_text(text),
        align="both",
        indent_first=0, indent_left=INDENT_FIRST_TW,
        line=LINE_BODY, line_rule="auto",
        space_before=60, space_after=60,
    )


def code_line_para(line: str) -> str:
    """Строка моноширинного блока — без переноса слов, single-spaced.

    Сохраняем все пробелы; не выравниваем по ширине (left)."""
    if line == "":
        return empty_para(LINE_SINGLE)
    return para(
        run(line, sz_hp=SZ_MONO_HP, font=FONT_MONO),
        align="left", indent_first=0,
        line=LINE_SINGLE, line_rule="auto",
        space_before=0, space_after=0,
        keep_lines=True,
    )


# ---------------------------------------------------------------------------
# Таблицы
# ---------------------------------------------------------------------------

def _cell_borders() -> str:
    return (
        '<w:tcBorders>'
        '<w:top w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:left w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:bottom w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:right w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '</w:tcBorders>'
    )


def _cell_para(text: str, *, bold: bool = False, align: str = "left") -> str:
    if not text:
        return para(
            run("", sz_hp=SZ_TABLE_HP, bold=bold),
            align=align, indent_first=0,
            line=LINE_SINGLE, line_rule="auto",
            space_before=0, space_after=0,
        )
    # В ячейке используем тот же runs_from_text (для **bold** и т.п.).
    return para(
        runs_from_text(text, sz_hp=SZ_TABLE_HP) if not bold
            else run(text, sz_hp=SZ_TABLE_HP, bold=True),
        align=align, indent_first=0,
        line=LINE_SINGLE, line_rule="auto",
        space_before=0, space_after=0,
    )


def _adaptive_widths(rows: list[list[str]], n_cols: int,
                     total_cm: float) -> list[float]:
    max_len = [0] * n_cols
    for r in rows:
        for ci in range(min(len(r), n_cols)):
            cell = r[ci] or ""
            words = cell.replace("\n", " ").split()
            longest_word = max((len(w) for w in words), default=0)
            cell_total = len(cell)
            effective = max(longest_word, min(cell_total // 2, 30))
            max_len[ci] = max(max_len[ci], effective)
    if sum(max_len) == 0:
        return [total_cm / n_cols] * n_cols
    min_cm = 1.2
    raw = [max(ml, 1) for ml in max_len]
    total_raw = sum(raw)
    widths = [max(min_cm, total_cm * r / total_raw) for r in raw]
    s = sum(widths)
    widths = [w * total_cm / s for w in widths]
    return widths


def render_table(header: list[str], rows: list[list[str]],
                 aligns: list[str] | None = None) -> str:
    n_cols = max([len(header)] + [len(r) for r in rows]) if rows or header else 1
    widths_cm = _adaptive_widths([header] + rows, n_cols, CONTENT_W_CM)
    widths_tw = [int(w * TWIPS_PER_CM) for w in widths_cm]
    total_w = sum(widths_tw)

    # Выравнивания ячеек тела по столбцам (left/center/right).
    if aligns is None:
        aligns = ["left"] * n_cols
    elif len(aligns) < n_cols:
        aligns = list(aligns) + ["left"] * (n_cols - len(aligns))

    grid = ''.join(f'<w:gridCol w:w="{w}"/>' for w in widths_tw)
    tbl_pr = (
        '<w:tblPr>'
        f'<w:tblW w:w="{total_w}" w:type="dxa"/>'
        '<w:jc w:val="center"/>'
        '<w:tblBorders>'
        '<w:top w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:left w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:bottom w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:right w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:insideH w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:insideV w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '</w:tblBorders>'
        '<w:tblLook w:val="04A0" w:firstRow="1" w:lastRow="0" '
        'w:firstColumn="1" w:lastColumn="0" w:noHBand="0" w:noVBand="1"/>'
        '</w:tblPr>'
    )

    rows_xml = []
    all_rows = ([header] if header else []) + rows
    for ridx, row in enumerate(all_rows):
        is_header = (ridx == 0 and header)
        row = list(row) + [""] * (n_cols - len(row))
        cells_xml = []
        for cidx, cell in enumerate(row):
            # Заливка шапок таблиц снята (PR #71): w:fill="auto" вместо
            # серого CLR_LIGHT_GRAY (#F2F2F2), чтобы шапка не отличалась
            # цветом от тела таблицы.
            shading = (
                '<w:shd w:val="clear" w:color="auto" w:fill="auto"/>'
                if is_header else ''
            )
            cells_xml.append(
                '<w:tc>'
                '<w:tcPr>'
                f'<w:tcW w:w="{widths_tw[cidx]}" w:type="dxa"/>'
                f'{_cell_borders()}'
                f'{shading}'
                '<w:vAlign w:val="center"/>'
                '</w:tcPr>'
                f'{_cell_para(str(cell), bold=False, align=("center" if is_header else aligns[cidx]))}'
                '</w:tc>'
            )
        row_pr = (
            '<w:trPr>'
            + ('<w:tblHeader/>' if is_header else '')
            + '<w:cantSplit/>'
            + '</w:trPr>'
        )
        rows_xml.append(f'<w:tr>{row_pr}{"".join(cells_xml)}</w:tr>')

    return (
        '<w:tbl>'
        f'{tbl_pr}'
        f'<w:tblGrid>{grid}</w:tblGrid>'
        f'{"".join(rows_xml)}'
        '</w:tbl>'
    ) + empty_para(LINE_SINGLE)


# ---------------------------------------------------------------------------
# Рисунки: встраивание PNG (DrawingML pic:pic), без внешних зависимостей
# ---------------------------------------------------------------------------

import struct


def png_size_px(path: Path) -> tuple[int, int]:
    """Ширина/высота PNG из заголовка IHDR (big-endian), без PIL."""
    try:
        data = path.read_bytes()[:24]
        if data[:8] == b"\x89PNG\r\n\x1a\n" and data[12:16] == b"IHDR":
            w, h = struct.unpack(">II", data[16:24])
            if w > 0 and h > 0:
                return int(w), int(h)
    except Exception:
        pass
    return 1600, 1000  # запасное соотношение 16:10


class ImageRegistry:
    """Накопитель встраиваемых изображений с присвоением rId и путей media/."""

    def __init__(self, base_dir: Path):
        self.base_dir = Path(base_dir)
        self.items: list[dict] = []
        self._by_key: dict[str, dict] = {}
        # styles/settings/fontTable/numbering уже заняли rId1..rId4.
        self._rid_n = 4
        self._doc_id = 1000

    def register(self, rel_path: str) -> dict:
        abspath = (self.base_dir / rel_path)
        key = str(abspath)
        if key in self._by_key:
            return self._by_key[key]
        self._rid_n += 1
        self._doc_id += 1
        idx = len(self.items) + 1
        item = {
            "rid": f"rId{self._rid_n}",
            "target": f"media/image{idx}.png",
            "abspath": abspath,
            "doc_id": self._doc_id,
            "rel_path": rel_path,
        }
        self._by_key[key] = item
        self.items.append(item)
        return item


def image_paragraph(reg: "ImageRegistry", rel_path: str, alt: str = "") -> str:
    """Центрированный абзац с инлайн-рисунком, вписанным по ширине набора."""
    item = reg.register(rel_path)
    w_px, h_px = png_size_px(item["abspath"])
    aspect = (w_px / h_px) if h_px else 1.6
    w_cm = CONTENT_W_CM
    h_cm = w_cm / aspect
    if h_cm > MAX_FIG_H_CM:
        h_cm = MAX_FIG_H_CM
        w_cm = h_cm * aspect
    cx = int(w_cm * EMU_PER_CM)
    cy = int(h_cm * EMU_PER_CM)
    rid = item["rid"]
    did = item["doc_id"]
    name = Path(rel_path).name
    a_ns = "http://schemas.openxmlformats.org/drawingml/2006/main"
    pic_ns = "http://schemas.openxmlformats.org/drawingml/2006/picture"
    drawing = (
        '<w:r><w:drawing>'
        '<wp:inline distT="0" distB="0" distL="0" distR="0">'
        f'<wp:extent cx="{cx}" cy="{cy}"/>'
        '<wp:effectExtent l="0" t="0" r="0" b="0"/>'
        f'<wp:docPr id="{did}" name="Picture {did}" descr="{xe(alt)}"/>'
        f'<wp:cNvGraphicFramePr><a:graphicFrameLocks xmlns:a="{a_ns}" '
        'noChangeAspect="1"/></wp:cNvGraphicFramePr>'
        f'<a:graphic xmlns:a="{a_ns}">'
        f'<a:graphicData uri="{pic_ns}">'
        f'<pic:pic xmlns:pic="{pic_ns}">'
        '<pic:nvPicPr>'
        f'<pic:cNvPr id="{did}" name="{xe(name)}" descr="{xe(alt)}"/>'
        '<pic:cNvPicPr/>'
        '</pic:nvPicPr>'
        '<pic:blipFill>'
        f'<a:blip r:embed="{rid}"/>'
        '<a:stretch><a:fillRect/></a:stretch>'
        '</pic:blipFill>'
        '<pic:spPr>'
        f'<a:xfrm><a:off x="0" y="0"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm>'
        '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>'
        '</pic:spPr>'
        '</pic:pic>'
        '</a:graphicData>'
        '</a:graphic>'
        '</wp:inline>'
        '</w:drawing></w:r>'
    )
    return para(
        drawing, align="center", indent_first=0,
        line=LINE_SINGLE, space_before=120, space_after=60,
        keep_next=True, keep_lines=True,
    )


# Подпись рисунка: строки, начинающиеся с «Рис. N» — по центру, без отступа.
_FIG_CAPTION_RE = re.compile(r"^Рис\.\s*\d")


def figure_caption_paragraph(text: str) -> str:
    return para(
        run(text, sz_hp=SZ_TABLE_HP),
        align="center", indent_first=0,
        line=LINE_SINGLE, space_before=0, space_after=200,
        keep_lines=True,
    )


# ---------------------------------------------------------------------------
# Парсер markdown
# ---------------------------------------------------------------------------

def _is_table_separator(line: str) -> bool:
    s = line.strip()
    return bool(re.match(r"^\|[\s\-\|:]+\|$", s)) and "-" in s


def _split_table_row(line: str) -> list[str]:
    s = line.strip()
    if s.startswith("|"):
        s = s[1:]
    if s.endswith("|"):
        s = s[:-1]
    return [c.strip() for c in s.split("|")]


def _parse_table_alignments(sep_line: str, n_cols: int) -> list[str]:
    """Разобрать разделитель markdown-таблицы и вернуть список выравниваний
    ('left' / 'center' / 'right') длиной n_cols.

    Правила:
      `:---:` → center, `---:` → right, `:---` → left, `---` → left (по умолчанию).
    """
    cells = _split_table_row(sep_line)
    aligns: list[str] = []
    for c in cells:
        c = c.strip()
        starts = c.startswith(":")
        ends = c.endswith(":")
        if starts and ends:
            aligns.append("center")
        elif ends:
            aligns.append("right")
        else:
            aligns.append("left")
    if len(aligns) < n_cols:
        aligns = aligns + ["left"] * (n_cols - len(aligns))
    return aligns[:n_cols]


def parse_markdown(md_path: Path) -> list[dict]:
    text = md_path.read_text(encoding="utf-8")
    lines = text.split("\n")
    blocks: list[dict] = []
    para_buf: list[str] = []
    quote_buf: list[str] = []
    list_buf: list[str] = []

    def flush_para():
        nonlocal para_buf
        if para_buf:
            t = " ".join(s.strip() for s in para_buf if s.strip())
            if t:
                blocks.append({"type": "para", "text": t})
            para_buf = []

    def flush_quote():
        nonlocal quote_buf
        if quote_buf:
            for q in quote_buf:
                blocks.append({"type": "quote", "text": q})
            quote_buf = []

    def flush_list():
        nonlocal list_buf
        if list_buf:
            for li in list_buf:
                blocks.append({"type": "li", "text": li})
            list_buf = []

    def flush_all():
        flush_para()
        flush_quote()
        flush_list()

    i = 0
    n = len(lines)
    while i < n:
        ln = lines[i]
        stripped = ln.strip()

        # Блок кода ```...``` (включая ```formula для формул)
        if stripped.startswith("```"):
            flush_all()
            info = stripped[3:].strip().lower()
            j = i + 1
            code_lines: list[str] = []
            while j < n and not lines[j].strip().startswith("```"):
                code_lines.append(lines[j])
                j += 1
            if info == "formula":
                blocks.append({"type": "formula", "lines": code_lines})
            else:
                blocks.append({"type": "code", "lines": code_lines})
            i = j + 1 if j < n else j
            continue

        # Картинка: ![alt](path)
        m_img = re.match(r"^!\[(.*?)\]\((.*?)\)\s*$", stripped)
        if m_img:
            flush_all()
            blocks.append({"type": "image", "alt": m_img.group(1).strip(),
                           "path": m_img.group(2).strip()})
            i += 1
            continue

        # Заголовки.
        if stripped.startswith("# "):
            flush_all()
            blocks.append({"type": "h1", "text": stripped[2:].strip()})
            i += 1
            continue
        if stripped.startswith("## "):
            flush_all()
            blocks.append({"type": "h2", "text": stripped[3:].strip()})
            i += 1
            continue
        if stripped.startswith("### "):
            flush_all()
            blocks.append({"type": "h2", "text": stripped[4:].strip()})
            i += 1
            continue

        # Горизонтальная линия
        if stripped == "---":
            flush_all()
            blocks.append({"type": "hr"})
            i += 1
            continue

        # Цитаты
        if stripped.startswith(">"):
            flush_para()
            flush_list()
            qtext = stripped[1:].strip()
            quote_buf.append(qtext)
            i += 1
            continue
        else:
            flush_quote()

        # Таблица
        if stripped.startswith("|") and i + 1 < n and \
                _is_table_separator(lines[i + 1]):
            flush_all()
            header = _split_table_row(lines[i])
            sep_line = lines[i + 1]
            aligns = _parse_table_alignments(sep_line, len(header))
            i += 2
            data_rows: list[list[str]] = []
            while i < n and lines[i].strip().startswith("|"):
                data_rows.append(_split_table_row(lines[i]))
                i += 1
            blocks.append({"type": "table", "header": header,
                           "rows": data_rows, "aligns": aligns})
            continue

        # Маркированный список (- ...)
        if stripped.startswith("- "):
            flush_para()
            list_buf.append(stripped[2:].strip())
            i += 1
            continue
        else:
            flush_list()

        # Пустая строка → конец абзаца
        if not stripped:
            flush_para()
            i += 1
            continue

        # Обычный текст
        para_buf.append(ln)
        i += 1

    flush_all()
    return blocks


# ---------------------------------------------------------------------------
# Рендер
# ---------------------------------------------------------------------------

def render_blocks(blocks: list[dict], reg: "ImageRegistry") -> str:
    parts: list[str] = []
    for idx, b in enumerate(blocks):
        t = b["type"]
        if t == "h1":
            # Если это заголовок главы (# 1., # 2., # 3., # Введение,
            # # Заключение, # Вывод по главе...) — добавим разрыв страницы
            # перед ним, кроме первого h1.
            txt = b["text"]
            is_chapter = bool(re.match(r"^(\d+\.\s|Введение|Заключение|Вывод по главе|Список использованных)", txt))
            if is_chapter and any(p["type"] not in ("hr",) for p in blocks[:idx]):
                # Только если это не первый смысловой блок документа.
                if idx > 0:
                    parts.append(page_break_para())
            parts.append(heading_h1(txt))
        elif t == "h2":
            parts.append(heading_h2(b["text"]))
        elif t == "para":
            if _FIG_CAPTION_RE.match(b["text"]):
                parts.append(figure_caption_paragraph(b["text"]))
            else:
                parts.append(body_paragraph(b["text"]))
        elif t == "image":
            parts.append(image_paragraph(reg, b["path"], b.get("alt", "")))
        elif t == "quote":
            parts.append(quote_paragraph(b["text"]))
        elif t == "li":
            parts.append(list_item(b["text"]))
        elif t == "code":
            for ln in b["lines"]:
                parts.append(code_line_para(ln))
            parts.append(empty_para(LINE_SINGLE))
        elif t == "formula":
            # Внутри блока:
            #   строка вида "EQ ||| (3.3.1)"  — формула с номером справа;
            #   строка "EQ"                    — формула без номера;
            #   строка "где| symbol — ..."     — первая строка пояснения;
            #   строка "| symbol — ..."        — последующие пояснения.
            where_entries: list[str] = []
            for ln in b["lines"]:
                s = ln.rstrip()
                if not s.strip():
                    continue
                if s.lstrip().startswith("|"):
                    where_entries.append(s.lstrip()[1:].strip())
                    continue
                # это строка формулы — сперва выгрузим накопленные пояснения
                if where_entries:
                    parts.append(where_block(where_entries))
                    where_entries = []
                if "|||" in s:
                    eq, num = s.split("|||", 1)
                    parts.append(formula_paragraph(eq.strip(), num.strip()))
                else:
                    parts.append(formula_paragraph(s.strip(), ""))
            if where_entries:
                parts.append(where_block(where_entries))
        elif t == "table":
            parts.append(render_table(b["header"], b["rows"],
                                      aligns=b.get("aligns")))
        elif t == "hr":
            # горизонтальные линии используем как «мягкий» разделитель —
            # просто пустой абзац (разрыв страницы добавляем перед h1)
            parts.append(empty_para(LINE_SINGLE))
        else:
            pass
    return "".join(parts)


# ---------------------------------------------------------------------------
# Сборка ZIP-контейнера .docx
# ---------------------------------------------------------------------------

CONTENT_TYPES_XML = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="png" ContentType="image/png"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
  <Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>
  <Override PartName="/word/fontTable.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml"/>
  <Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>
  <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
  <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
</Types>'''

ROOT_RELS_XML = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
</Relationships>'''

DOC_RELS_XML = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable" Target="fontTable.xml"/>
  <Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/>
</Relationships>'''

CORE_XML = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
                   xmlns:dc="http://purl.org/dc/elements/1.1/"
                   xmlns:dcterms="http://purl.org/dc/terms/"
                   xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
  <dc:title>ВКР ГОТОВОЕ</dc:title>
  <dc:creator>Автор</dc:creator>
  <cp:lastModifiedBy>Автор</cp:lastModifiedBy>
</cp:coreProperties>'''

APP_XML = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"
            xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
  <Application>VKR-md2docx</Application>
</Properties>'''

SETTINGS_XML = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:zoom w:percent="100"/>
  <w:defaultTabStop w:val="708"/>
  <w:characterSpacingControl w:val="doNotCompress"/>
  <w:compat>
    <w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/>
  </w:compat>
  <w:themeFontLang w:val="ru-RU"/>
</w:settings>'''

FONT_TABLE_XML = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:fonts xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:font w:name="Times New Roman">
    <w:panose1 w:val="02020603050405020304"/>
    <w:charset w:val="CC"/>
    <w:family w:val="roman"/>
    <w:pitch w:val="variable"/>
  </w:font>
  <w:font w:name="Courier New">
    <w:panose1 w:val="02070309020205020404"/>
    <w:charset w:val="CC"/>
    <w:family w:val="modern"/>
    <w:pitch w:val="fixed"/>
  </w:font>
</w:fonts>'''

NUMBERING_XML = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>'''

STYLES_XML = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:docDefaults>
    <w:rPrDefault>
      <w:rPr>
        <w:rFonts w:ascii="{FONT_MAIN}" w:hAnsi="{FONT_MAIN}" w:cs="{FONT_MAIN}" w:eastAsia="{FONT_MAIN}"/>
        <w:sz w:val="{SZ_BODY_HP}"/>
        <w:szCs w:val="{SZ_BODY_HP}"/>
        <w:lang w:val="ru-RU" w:eastAsia="ru-RU" w:bidi="ar-SA"/>
      </w:rPr>
    </w:rPrDefault>
    <w:pPrDefault>
      <w:pPr>
        <w:spacing w:before="0" w:after="0" w:line="{LINE_BODY}" w:lineRule="auto"/>
        <w:jc w:val="both"/>
      </w:pPr>
    </w:pPrDefault>
  </w:docDefaults>
  <w:style w:type="paragraph" w:default="1" w:styleId="Normal">
    <w:name w:val="Normal"/>
    <w:qFormat/>
  </w:style>
</w:styles>'''


def build_document_xml(body_xml: str) -> str:
    sect_pr = (
        '<w:sectPr>'
        f'<w:pgSz w:w="{PAGE_W_TW}" w:h="{PAGE_H_TW}"/>'
        f'<w:pgMar w:top="{MARGIN_TOP_TW}" w:right="{MARGIN_RIGHT_TW}" '
        f'w:bottom="{MARGIN_BOTTOM_TW}" w:left="{MARGIN_LEFT_TW}" '
        f'w:header="720" w:footer="720" w:gutter="0"/>'
        '<w:cols w:space="708"/>'
        '<w:docGrid w:linePitch="360"/>'
        '</w:sectPr>'
    )
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        '<w:document '
        'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
        'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" '
        'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
        'xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" '
        'xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas" '
        'xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup" '
        'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">'
        '<w:body>'
        f'{body_xml}'
        f'{sect_pr}'
        '</w:body>'
        '</w:document>'
    )


def write_docx(out_path: Path, body_xml: str,
               reg: "ImageRegistry | None" = None) -> None:
    doc_xml = build_document_xml(body_xml)
    # Динамические связи документа: к базовым (styles/settings/fontTable/
    # numbering) добавляем по одной связи на каждое встроенное изображение.
    img_rels = ""
    for it in (reg.items if reg else []):
        img_rels += (
            f'\n  <Relationship Id="{it["rid"]}" '
            'Type="http://schemas.openxmlformats.org/officeDocument/2006/'
            f'relationships/image" Target="{it["target"]}"/>'
        )
    doc_rels = DOC_RELS_XML.replace(
        "</Relationships>", f"{img_rels}\n</Relationships>"
    )
    with zipfile.ZipFile(out_path, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("[Content_Types].xml", CONTENT_TYPES_XML)
        zf.writestr("_rels/.rels", ROOT_RELS_XML)
        zf.writestr("word/_rels/document.xml.rels", doc_rels)
        zf.writestr("word/document.xml", doc_xml)
        zf.writestr("word/styles.xml", STYLES_XML)
        zf.writestr("word/settings.xml", SETTINGS_XML)
        zf.writestr("word/fontTable.xml", FONT_TABLE_XML)
        zf.writestr("word/numbering.xml", NUMBERING_XML)
        zf.writestr("docProps/core.xml", CORE_XML)
        zf.writestr("docProps/app.xml", APP_XML)
        # Встроенные изображения.
        for it in (reg.items if reg else []):
            zf.write(it["abspath"], f"word/{it['target']}")


# ---------------------------------------------------------------------------
# main
# ---------------------------------------------------------------------------

def main() -> None:
    here = Path(__file__).parent
    src = here / "ВКР ГОТОВОЕ.md"
    out = here / "ВКР ГОТОВОЕ.docx"
    blocks = parse_markdown(src)
    reg = ImageRegistry(here)
    body_xml = render_blocks(blocks, reg)
    write_docx(out, body_xml, reg)
    print(f"Записан файл: {out} ({out.stat().st_size:,} байт)")
    print(f"Встроено изображений: {len(reg.items)}")
    by_type: dict[str, int] = {}
    for b in blocks:
        by_type[b["type"]] = by_type.get(b["type"], 0) + 1
    print("Блоки:")
    for k in sorted(by_type):
        print(f"  {k}: {by_type[k]}")


if __name__ == "__main__":
    main()
