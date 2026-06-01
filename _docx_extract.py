#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Минимальный конвертер .docx -> Markdown через стандартную библиотеку.

Не использует python-docx: распаковывает zip-архив и парсит word/document.xml.
"""
from __future__ import annotations

import re
import sys
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NS = {"w": W}


def q(tag: str) -> str:
    return f"{{{W}}}{tag}"


def run_text(r) -> tuple[str, bool, bool]:
    """Текст run + флаги bold/italic."""
    rpr = r.find(q("rPr"))
    bold = italic = False
    if rpr is not None:
        b = rpr.find(q("b"))
        if b is not None and b.attrib.get(q("val"), "true") not in ("0", "false"):
            bold = True
        i = rpr.find(q("i"))
        if i is not None and i.attrib.get(q("val"), "true") not in ("0", "false"):
            italic = True
    parts = []
    for child in r.iter():
        tag = child.tag
        if tag == q("t"):
            parts.append(child.text or "")
        elif tag == q("tab"):
            parts.append("\t")
        elif tag == q("br"):
            parts.append("\n")
    return "".join(parts), bold, italic


def para_to_md(p) -> str | None:
    ppr = p.find(q("pPr"))
    align = None
    if ppr is not None:
        jc = ppr.find(q("jc"))
        if jc is not None:
            align = jc.attrib.get(q("val"))

    # Сборка инлайн-текста с форматированием
    runs = p.findall(q("r"))
    inline_parts = []
    plain_parts = []
    any_bold_run = False
    all_bold = True
    has_text = False
    for r in runs:
        text, bold, italic = run_text(r)
        if not text:
            continue
        plain_parts.append(text)
        if text.strip():
            has_text = True
            if bold:
                any_bold_run = True
            else:
                all_bold = False
        core = text.strip()
        lead = text[: len(text) - len(text.lstrip())]
        trail = text[len(text.rstrip()):]
        if core:
            esc = core.replace("|", "\\|")
            if bold and italic:
                esc = f"***{esc}***"
            elif bold:
                esc = f"**{esc}**"
            elif italic:
                esc = f"*{esc}*"
            inline_parts.append(f"{lead}{esc}{trail}")
        else:
            inline_parts.append(text)

    plain = "".join(plain_parts).strip()
    if not plain:
        return None
    inline = re.sub(r"[ \t]{2,}", " ", "".join(inline_parts)).strip()

    centered = align == "center"
    is_h1 = centered and any_bold_run and all_bold and re.match(
        r"^(ВВЕДЕНИЕ|ЗАКЛЮЧЕНИЕ|ПРИЛОЖЕНИЯ|ПРИЛОЖЕНИЕ|ГЛАВА\b|СПИСОК\b|ОГЛАВЛЕНИЕ|СОДЕРЖАНИЕ|РЕФЕРАТ|АННОТАЦИЯ)",
        plain,
    )
    is_h2 = centered and any_bold_run and all_bold and re.match(
        r"^\d+(?:\.\d+)+\.?\s", plain
    )
    if is_h1:
        return f"# {plain}"
    if is_h2:
        return f"## {plain}"
    if centered and any_bold_run and all_bold and has_text and len(plain) < 120:
        return f"### {plain}"

    if re.match(r"^\s*[\u2013\u2014\u2012\u2022\-]\s+", plain):
        body = re.sub(r"^\s*[\u2013\u2014\u2012\u2022\-]\s+", "", inline, count=1)
        return f"- {body}"

    return inline


def cell_text(tc) -> str:
    parts = []
    for p in tc.findall(q("p")):
        md = para_to_md(p)
        if md:
            parts.append(md)
    return "<br>".join(parts).replace("|", "\\|")


def table_to_md(tbl) -> str:
    rows = tbl.findall(q("tr"))
    if not rows:
        return ""
    n_cols = max(len(r.findall(q("tc"))) for r in rows)

    def fmt(cells):
        cells = cells + [""] * (n_cols - len(cells))
        return "| " + " | ".join(cells) + " |"

    lines = []
    header = [cell_text(c) for c in rows[0].findall(q("tc"))]
    lines.append(fmt(header))
    lines.append("| " + " | ".join(["---"] * n_cols) + " |")
    for r in rows[1:]:
        lines.append(fmt([cell_text(c) for c in r.findall(q("tc"))]))
    return "\n".join(lines)


def convert(src: Path, dst: Path) -> None:
    with zipfile.ZipFile(src) as z:
        with z.open("word/document.xml") as f:
            tree = ET.parse(f)
    body = tree.getroot().find(q("body"))

    out: list[str] = []
    prev_was_list = False
    for child in list(body):
        if child.tag == q("p"):
            md = para_to_md(child)
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
        elif child.tag == q("tbl"):
            out.append("")
            out.append(table_to_md(child))
            out.append("")
            prev_was_list = False

    text = "\n".join(out)
    text = re.sub(r"\n{3,}", "\n\n", text).strip() + "\n"
    dst.write_text(text, encoding="utf-8")
    n = text.count("\n") + 1
    print(f"OK: {dst.name} ({dst.stat().st_size:,} байт, {n} строк)")


if __name__ == "__main__":
    src = Path(sys.argv[1])
    dst = src.with_suffix(".md")
    convert(src, dst)
