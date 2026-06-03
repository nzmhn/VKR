#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Конвертер Word (.docx) → Markdown для файла «raschet_B2B.docx».

В окружении недоступен python-docx, поэтому .docx (это ZIP с XML) разбирается
стандартной библиотекой. Тело документа обходится по порядку: абзацы и таблицы.

Эвристики восстановления структуры:
  * # (H1)  — ВВЕДЕНИЕ, ГЛАВА N, ЗАКЛЮЧЕНИЕ, СПИСОК…, ПРИЛОЖЕНИЯ и т. п.;
  * ## (H2) — нумерованные параграфы вида «1.1.», «2.3.»;
  * ### (H3)— прочие центрированные полужирные строки;
  * списки  — абзацы, начинающиеся с тире «–»/«-»/«•»;
  * таблицы — реальные таблицы Word → таблицы Markdown;
  * инлайн **полужирный** и *курсив* берётся из run-ов (w:b / w:i).
"""

from __future__ import annotations

import re
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path

HERE = Path(__file__).parent
SRC = HERE / "raschet_B2B.docx"
DST = HERE / "raschet_B2B.md"

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


def q(tag: str) -> str:
    return f"{{{W}}}{tag}"


_BULLET_RE = re.compile(r"^\s*[\u2013\u2014\u2012\u2022\-]\s+")
_H2_RE = re.compile(r"^\d+(?:\.\d+)+\.?\s")
_H1_RE = re.compile(
    r"^(ВВЕДЕНИЕ|ЗАКЛЮЧЕНИЕ|ПРИЛОЖЕНИЯ|ГЛАВА\b|СПИСОК\b|ОГЛАВЛЕНИЕ|СОДЕРЖАНИЕ|РАСЧ[ЕЁ]Т)"
)


def esc(text: str) -> str:
    """Экранирование символов, значимых в Markdown."""
    return text.replace("\\", "\\\\").replace("|", "\\|")


def run_text(run: ET.Element) -> str:
    """Текст одного run с учётом w:t, w:tab, w:br."""
    parts: list[str] = []
    for child in run.iter():
        tag = child.tag
        if tag == q("t"):
            parts.append(child.text or "")
        elif tag == q("tab"):
            parts.append("\t")
        elif tag in (q("br"), q("cr")):
            parts.append(" ")
    return "".join(parts)


def render_run(run: ET.Element) -> str:
    text = run_text(run)
    if not text:
        return ""
    rpr = run.find(q("rPr"))
    bold = italic = False
    if rpr is not None:
        b = rpr.find(q("b"))
        i = rpr.find(q("i"))
        bold = b is not None and b.get(q("val"), "true") not in ("false", "0")
        italic = i is not None and i.get(q("val"), "true") not in ("false", "0")
    # Сохраняем ведущие/замыкающие пробелы вне маркеров.
    lead = text[: len(text) - len(text.lstrip())]
    trail = text[len(text.rstrip()):]
    core = text.strip()
    if not core:
        return text
    core = esc(core)
    if bold and italic:
        core = f"***{core}***"
    elif bold:
        core = f"**{core}**"
    elif italic:
        core = f"*{core}*"
    return f"{lead}{core}{trail}"


def paragraph_runs_text(p: ET.Element) -> str:
    """Собрать форматированный текст абзаца (с инлайн-разметкой)."""
    out: list[str] = []
    for r in p.findall(q("r")):
        out.append(render_run(r))
    text = "".join(out)
    # Сворачиваем разрывы в соседних маркерах вида **a** **b** -> **a b**? нет, оставляем.
    return re.sub(r"[ \t]+", " ", text).strip()


def paragraph_plain(p: ET.Element) -> str:
    return re.sub(r"\s+", " ", "".join(t.text or "" for t in p.iter(q("t")))).strip()


def is_centered_bold(p: ET.Element) -> bool:
    ppr = p.find(q("pPr"))
    centered = False
    if ppr is not None:
        jc = ppr.find(q("jc"))
        centered = jc is not None and jc.get(q("val")) == "center"
    runs = p.findall(q("r"))
    if not runs:
        return False
    bold_all = True
    has_text = False
    for r in runs:
        if not (run_text(r).strip()):
            continue
        has_text = True
        rpr = r.find(q("rPr"))
        b = rpr.find(q("b")) if rpr is not None else None
        if not (b is not None and b.get(q("val"), "true") not in ("false", "0")):
            bold_all = False
    return centered and bold_all and has_text


def convert_paragraph(p: ET.Element) -> str:
    plain = paragraph_plain(p)
    if not plain:
        return ""
    formatted = paragraph_runs_text(p)

    # Заголовки.
    if _H1_RE.match(plain):
        return f"# {plain}"
    if _H2_RE.match(plain):
        return f"## {plain}"
    if is_centered_bold(p) and len(plain) < 120:
        return f"### {plain}"

    # Списки.
    if _BULLET_RE.match(plain):
        item = _BULLET_RE.sub("", formatted).strip()
        return f"- {item}"

    return formatted


def convert_table(tbl: ET.Element) -> str:
    rows: list[list[str]] = []
    for tr in tbl.findall(q("tr")):
        cells: list[str] = []
        for tc in tr.findall(q("tc")):
            cell_lines = []
            for p in tc.findall(q("p")):
                txt = paragraph_runs_text(p)
                if txt:
                    cell_lines.append(txt)
            cells.append("<br>".join(cell_lines))
        if cells:
            rows.append(cells)
    if not rows:
        return ""
    width = max(len(r) for r in rows)
    rows = [r + [""] * (width - len(r)) for r in rows]
    lines = []
    header = rows[0]
    lines.append("| " + " | ".join(header) + " |")
    lines.append("| " + " | ".join(["---"] * width) + " |")
    for r in rows[1:]:
        lines.append("| " + " | ".join(r) + " |")
    return "\n".join(lines)


def main() -> None:
    z = zipfile.ZipFile(SRC)
    xml = z.read("word/document.xml")
    root = ET.fromstring(xml)
    body = root.find(q("body"))
    assert body is not None

    blocks: list[str] = []
    for el in body:
        if el.tag == q("p"):
            md = convert_paragraph(el)
            if md:
                blocks.append(md)
        elif el.tag == q("tbl"):
            md = convert_table(el)
            if md:
                blocks.append(md)

    # Объединяем последовательные пункты списка без пустой строки между ними.
    out_lines: list[str] = []
    prev_is_list = False
    for b in blocks:
        is_list = b.startswith("- ")
        if out_lines:
            if is_list and prev_is_list:
                out_lines.append(b)
            else:
                out_lines.append("")
                out_lines.append(b)
        else:
            out_lines.append(b)
        prev_is_list = is_list

    DST.write_text("\n".join(out_lines) + "\n", encoding="utf-8")
    print(f"Готово: {DST.name}")
    print(f"Блоков: {len(blocks)}, строк: {len(out_lines)}")


if __name__ == "__main__":
    main()
