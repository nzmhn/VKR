#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Конвертер «analiz_glava3.docx» → «analiz_glava3.md».

Работает только на стандартной библиотеке Python (zipfile + ElementTree),
без сторонних зависимостей. Структура восстанавливается по стилям абзацев:

  * Heading1/Title  -> #
  * Heading2        -> ##
  * Heading3        -> ###
  * списки (numPr)  -> «- » с отступом по уровню (ilvl);
                       нумерованные списки -> «1. » и т. п.;
  * таблицы Word    -> таблицы Markdown;
  * инлайн **bold** / *italic* берётся из свойств run-ов.
"""

from __future__ import annotations

import re
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

HERE = Path(__file__).parent
SRC = HERE / "analiz_glava3.docx"
DST = HERE / "analiz_glava3.md"

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


def q(tag: str) -> str:
    return f"{{{W}}}{tag}"


def build_numbering(zf: zipfile.ZipFile) -> dict:
    """numId -> {ilvl -> ('bullet'|'decimal', lvlText)}."""
    try:
        root = ET.fromstring(zf.read("word/numbering.xml"))
    except KeyError:
        return {}

    abstract: dict[str, dict[int, tuple[str, str]]] = {}
    for an in root.findall(q("abstractNum")):
        aid = an.get(q("abstractNumId"))
        levels: dict[int, tuple[str, str]] = {}
        for lvl in an.findall(q("lvl")):
            ilvl = int(lvl.get(q("ilvl"), "0"))
            fmt_el = lvl.find(q("numFmt"))
            txt_el = lvl.find(q("lvlText"))
            fmt = fmt_el.get(q("val")) if fmt_el is not None else "bullet"
            txt = txt_el.get(q("val")) if txt_el is not None else "•"
            levels[ilvl] = (fmt, txt)
        abstract[aid] = levels

    num_map: dict[str, dict[int, tuple[str, str]]] = {}
    for num in root.findall(q("num")):
        nid = num.get(q("numId"))
        aref = num.find(q("abstractNumId"))
        if aref is not None:
            num_map[nid] = abstract.get(aref.get(q("val")), {})
    return num_map


def run_text(r: ET.Element) -> str:
    """Текст одного run с учётом w:t/w:tab/w:br."""
    parts: list[str] = []
    for child in r:
        tag = child.tag
        if tag == q("t"):
            parts.append(child.text or "")
        elif tag == q("tab"):
            parts.append("\t")
        elif tag in (q("br"), q("cr")):
            parts.append("\n")
    return "".join(parts)


def run_md(r: ET.Element) -> str:
    text = run_text(r)
    if text == "":
        return ""
    rpr = r.find(q("rPr"))
    bold = italic = False
    if rpr is not None:
        b = rpr.find(q("b"))
        i = rpr.find(q("i"))
        bold = b is not None and b.get(q("val"), "true") not in ("false", "0")
        italic = i is not None and i.get(q("val"), "true") not in ("false", "0")
    lead = text[: len(text) - len(text.lstrip(" "))]
    trail = text[len(text.rstrip(" ")):]
    core = text.strip(" ")
    if core:
        esc = core.replace("|", "\\|")
        if bold and italic:
            esc = f"***{esc}***"
        elif bold:
            esc = f"**{esc}**"
        elif italic:
            esc = f"*{esc}*"
        return f"{lead}{esc}{trail}"
    return text


def para_inline(p: ET.Element) -> str:
    parts: list[str] = []
    # Учитываем run-ы как напрямую, так и внутри гиперссылок.
    for el in p.iter():
        if el.tag == q("r"):
            parts.append(run_md(el))
    out = "".join(parts)
    out = re.sub(r"[ \t]{2,}", " ", out)
    return out.strip()


def para_plain(p: ET.Element) -> str:
    """Текст абзаца без инлайн-форматирования (для заголовков)."""
    parts: list[str] = []
    for el in p.iter():
        if el.tag == q("r"):
            parts.append(run_text(el))
    out = "".join(parts)
    out = re.sub(r"[ \t]{2,}", " ", out)
    return out.strip()


def get_pstyle(p: ET.Element) -> str | None:
    ppr = p.find(q("pPr"))
    if ppr is None:
        return None
    st = ppr.find(q("pStyle"))
    return st.get(q("val")) if st is not None else None


def get_numpr(p: ET.Element):
    ppr = p.find(q("pPr"))
    if ppr is None:
        return None
    numpr = ppr.find(q("numPr"))
    if numpr is None:
        return None
    ilvl_el = numpr.find(q("ilvl"))
    numid_el = numpr.find(q("numId"))
    ilvl = int(ilvl_el.get(q("val"))) if ilvl_el is not None else 0
    numid = numid_el.get(q("val")) if numid_el is not None else None
    return (numid, ilvl)


HEADING_LEVEL = {
    "Title": 1,
    "Heading1": 1,
    "Heading2": 2,
    "Heading3": 3,
    "Heading4": 4,
    "Heading5": 5,
    "Heading6": 6,
}


def para_to_md(p: ET.Element, numbering: dict, counters: dict) -> str | None:
    style = get_pstyle(p)
    text = para_inline(p)

    if style in HEADING_LEVEL:
        head = para_plain(p)
        if not head:
            return None
        lvl = HEADING_LEVEL[style]
        return f"{'#' * lvl} {head}"

    numpr = get_numpr(p)
    if numpr is not None and text:
        numid, ilvl = numpr
        levels = numbering.get(numid, {})
        fmt, _ = levels.get(ilvl, ("bullet", "•"))
        indent = "  " * ilvl
        if fmt == "decimal":
            key = (numid, ilvl)
            counters[key] = counters.get(key, 0) + 1
            return f"{indent}{counters[key]}. {text}"
        return f"{indent}- {text}"

    if not text:
        return None
    return text


def cell_to_md(tc: ET.Element) -> str:
    parts: list[str] = []
    for p in tc.findall(q("p")):
        t = para_inline(p)
        if t:
            parts.append(t)
    return "<br>".join(parts).replace("|", "\\|")


def table_to_md(tbl: ET.Element) -> str:
    rows = tbl.findall(q("tr"))
    if not rows:
        return ""
    grid = [[cell_to_md(tc) for tc in r.findall(q("tc"))] for r in rows]
    n_cols = max(len(r) for r in grid)

    def fmt_row(cells: list[str]) -> str:
        cells = cells + [""] * (n_cols - len(cells))
        return "| " + " | ".join(cells) + " |"

    lines = [fmt_row(grid[0]), "| " + " | ".join(["---"] * n_cols) + " |"]
    for r in grid[1:]:
        lines.append(fmt_row(r))
    return "\n".join(lines)


def main() -> None:
    zf = zipfile.ZipFile(str(SRC))
    numbering = build_numbering(zf)
    root = ET.fromstring(zf.read("word/document.xml"))
    body = root.find(q("body"))

    out: list[str] = []
    counters: dict = {}
    prev_was_list = False

    for block in list(body):
        tag = block.tag
        if tag == q("tbl"):
            out.append("")
            out.append(table_to_md(block))
            out.append("")
            prev_was_list = False
            counters = {}
            continue
        if tag != q("p"):
            continue

        md = para_to_md(block, numbering, counters)
        if md is None:
            continue

        is_list = bool(re.match(r"^\s*(?:[-*]|\d+\.)\s", md))
        is_heading = md.startswith("#")

        if is_heading:
            out.append("")
            out.append(md)
            out.append("")
            counters = {}
        elif is_list:
            if not prev_was_list and out and out[-1] != "":
                out.append("")
            out.append(md)
        else:
            if prev_was_list:
                out.append("")
            out.append(md)
            out.append("")
            counters = {}

        prev_was_list = is_list

    text = "\n".join(out)
    text = re.sub(r"\n{3,}", "\n\n", text).strip() + "\n"
    DST.write_text(text, encoding="utf-8")

    n_lines = text.count("\n") + 1
    print(f"OK: {DST.name} ({DST.stat().st_size:,} bytes, {n_lines} lines)")


if __name__ == "__main__":
    main()
