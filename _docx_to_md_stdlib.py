#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Конвертер Word (.docx) → Markdown на стандартной библиотеке (без python-docx).

Используется, когда python-docx недоступен (нет доступа к PyPI). Файл .docx —
это zip-архив с XML; разбираем word/document.xml напрямую.

Особенности:
  * правки (track changes): вставки <w:ins> принимаются (текст включается),
    удаления <w:del>/<w:delText> отклоняются (текст отбрасывается) — то есть
    формируется «чистая» версия документа с принятыми изменениями;
  * заголовки восстанавливаются по эвристике: абзац по центру и полужирный
      – # (H1): ВВЕДЕНИЕ, ГЛАВА N…, ЗАКЛЮЧЕНИЕ, СПИСОК…, ПРИЛОЖЕНИЯ;
      – ## (H2): нумерованные параграфы «1.1.», «2.3.» и т. п.;
      – ### (H3): прочие центрированные полужирные строки;
  * маркированные списки — абзацы, начинающиеся с тире «–»/«-»/«•»;
  * таблицы Word → таблицы Markdown;
  * инлайн **полужирный** и *курсив* берутся из run-ов.

Запуск:  python3 _docx_to_md_stdlib.py <файл.docx> [файл.md]
"""

from __future__ import annotations

import re
import sys
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


def q(tag: str) -> str:
    return f"{{{W}}}{tag}"


_BULLET_RE = re.compile(r"^\s*[\u2013\u2014\u2012•\-]\s+")
_H2_RE = re.compile(r"^\d+(?:\.\d+)+\.?\s")
# Заголовок главы вида «1. ТЕОРЕТИЧЕСКИЕ …» (одиночный номер + заглавный текст),
# но не нумерованный параграф «1.1.» (его ловит _H2_RE).
_CHAPTER_RE = re.compile(r"^\d+\.\s+[А-ЯA-Z]")
_H1_RE = re.compile(
    r"^(ВВЕДЕНИЕ|ЗАКЛЮЧЕНИЕ|ПРИЛОЖЕНИЯ|ГЛАВА\b|СПИСОК\b|ОГЛАВЛЕНИЕ|СОДЕРЖАНИЕ)"
)


def run_text(r: ET.Element) -> str:
    """Собрать текст run-а: <w:t>, табуляции и переводы строк."""
    parts: list[str] = []
    for el in r.iter():
        tag = el.tag
        if tag == q("t"):
            parts.append(el.text or "")
        elif tag == q("tab"):
            parts.append("\t")
        elif tag in (q("br"), q("cr")):
            parts.append(" ")
    return "".join(parts)


def is_bold(r: ET.Element) -> bool:
    rpr = r.find(q("rPr"))
    if rpr is None:
        return False
    b = rpr.find(q("b"))
    if b is None:
        return False
    return b.get(q("val"), "true") not in ("0", "false", "none")


def is_italic(r: ET.Element) -> bool:
    rpr = r.find(q("rPr"))
    if rpr is None:
        return False
    i = rpr.find(q("i"))
    if i is None:
        return False
    return i.get(q("val"), "true") not in ("0", "false", "none")


def collect_runs(node: ET.Element, runs: list[ET.Element]) -> None:
    """Рекурсивно собрать run-ы в порядке следования.

    Содержимое <w:del> (удалённое в правках) пропускается; <w:ins>
    (вставленное) обрабатывается прозрачно — правки принимаются.
    """
    for child in node:
        tag = child.tag
        if tag == q("del"):
            continue  # удалённый текст отбрасываем
        if tag == q("r"):
            runs.append(child)
        elif tag == q("ins"):
            collect_runs(child, runs)  # вставку принимаем
        elif tag in (q("smartTag"), q("hyperlink"), q("sdt"), q("sdtContent")):
            collect_runs(child, runs)


def para_alignment(p: ET.Element) -> str | None:
    ppr = p.find(q("pPr"))
    if ppr is None:
        return None
    jc = ppr.find(q("jc"))
    return jc.get(q("val")) if jc is not None else None


def inline_from_runs(runs: list[ET.Element]) -> str:
    parts: list[str] = []
    for r in runs:
        t = run_text(r)
        if t == "":
            continue
        lead = t[: len(t) - len(t.lstrip())]
        trail = t[len(t.rstrip()):]
        core = t.strip()
        esc = core.replace("|", "\\|") if core else core
        if core:
            if is_bold(r) and is_italic(r):
                esc = f"***{esc}***"
            elif is_bold(r):
                esc = f"**{esc}**"
            elif is_italic(r):
                esc = f"*{esc}*"
        parts.append(f"{lead}{esc}{trail}")
    out = "".join(parts)
    out = re.sub(r"[ \t]{2,}", " ", out)
    return out.strip()


def para_plain_text(runs: list[ET.Element]) -> str:
    return re.sub(r"[ \t]{2,}", " ", "".join(run_text(r) for r in runs)).strip()


def classify_heading(p: ET.Element, runs: list[ET.Element], text: str) -> int | None:
    centered = para_alignment(p) == "center"
    nonempty = [r for r in runs if run_text(r).strip()]
    if not nonempty:
        return None
    has_bold = any(is_bold(r) for r in nonempty)
    all_bold = has_bold and all(is_bold(r) for r in nonempty)
    if not (centered and all_bold):
        return None
    if _H1_RE.match(text):
        return 1
    if _H2_RE.match(text):
        return 2
    if _CHAPTER_RE.match(text):
        return 1
    return 3


def para_to_md(p: ET.Element) -> str | None:
    runs: list[ET.Element] = []
    collect_runs(p, runs)
    text = para_plain_text(runs)
    if not text:
        return None

    level = classify_heading(p, runs, text)
    if level is not None:
        return f"{'#' * level} {text}"

    if _BULLET_RE.match(text):
        body = inline_from_runs(runs)
        body = _BULLET_RE.sub("", body, count=1)
        return f"- {body}"

    return inline_from_runs(runs) or text.replace("|", "\\|")


def cell_text(tc: ET.Element) -> str:
    chunks: list[str] = []
    for p in tc.findall(q("p")):
        runs: list[ET.Element] = []
        collect_runs(p, runs)
        txt = para_plain_text(runs)
        if txt:
            chunks.append(txt)
    return "<br>".join(chunks).replace("|", "\\|")


def table_to_md(tbl: ET.Element) -> str:
    rows = tbl.findall(q("tr"))
    if not rows:
        return ""
    grid = [[cell_text(c) for c in r.findall(q("tc"))] for r in rows]
    n_cols = max((len(r) for r in grid), default=0)
    if n_cols == 0:
        return ""

    def fmt_row(cells: list[str]) -> str:
        cells = cells + [""] * (n_cols - len(cells))
        return "| " + " | ".join(cells) + " |"

    lines = [fmt_row(grid[0]), "| " + " | ".join(["---"] * n_cols) + " |"]
    for r in grid[1:]:
        lines.append(fmt_row(r))
    return "\n".join(lines)


def convert(src: Path, dst: Path) -> None:
    xml = zipfile.ZipFile(str(src)).read("word/document.xml")
    root = ET.fromstring(xml)
    body = root.find(q("body"))
    if body is None:
        raise SystemExit("Не найден <w:body> в document.xml")

    out: list[str] = []
    prev_was_list = False

    for block in body:
        tag = block.tag
        if tag == q("tbl"):
            out.append("")
            out.append(table_to_md(block))
            out.append("")
            prev_was_list = False
            continue
        if tag != q("p"):
            continue

        md = para_to_md(block)
        if md is None:
            continue

        is_list = md.startswith("- ")
        is_heading = md.startswith("#")

        if is_heading:
            out.append("")
            out.append(md)
            out.append("")
        elif is_list:
            if not prev_was_list and out and out[-1] != "":
                out.append("")
            out.append(md)
        else:
            if prev_was_list:
                out.append("")
            out.append(md)
            out.append("")

        prev_was_list = is_list

    text = "\n".join(out)
    text = re.sub(r"\n{3,}", "\n\n", text).strip() + "\n"
    dst.write_text(text, encoding="utf-8")

    n_lines = text.count("\n") + 1
    print(f"OK: {dst.name} ({dst.stat().st_size:,} байт, {n_lines} строк)")


def main() -> None:
    if len(sys.argv) < 2:
        raise SystemExit("Использование: python3 _docx_to_md_stdlib.py <файл.docx> [файл.md]")
    src = Path(sys.argv[1])
    dst = Path(sys.argv[2]) if len(sys.argv) > 2 else src.with_suffix(".md")
    convert(src, dst)


if __name__ == "__main__":
    main()
