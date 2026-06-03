#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Сравнение содержимого двух .docx по абзацам (правки приняты).

Извлекает плоский текст из каждого документа (вставки <w:ins> принимаются,
удаления <w:del> отбрасываются), нормализует пробелы и сравнивает абзацы
через difflib. Печатает блоки различий в формате, удобном для чтения.

Запуск: python3 _compare_docx.py <a.docx> <b.docx>
"""
from __future__ import annotations

import re
import sys
import zipfile
import difflib
import xml.etree.ElementTree as ET

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


def q(tag: str) -> str:
    return f"{{{W}}}{tag}"


def run_text(r: ET.Element) -> str:
    parts: list[str] = []
    for el in r.iter():
        if el.tag == q("t"):
            parts.append(el.text or "")
        elif el.tag == q("tab"):
            parts.append(" ")
        elif el.tag in (q("br"), q("cr")):
            parts.append(" ")
    return "".join(parts)


def collect_runs(node: ET.Element, runs: list[ET.Element]) -> None:
    for child in node:
        tag = child.tag
        if tag == q("del"):
            continue
        if tag == q("r"):
            runs.append(child)
        elif tag in (q("ins"), q("smartTag"), q("hyperlink"), q("sdt"), q("sdtContent")):
            collect_runs(child, runs)


def norm(s: str) -> str:
    s = s.replace("\u00a0", " ")
    s = re.sub(r"\s+", " ", s)
    return s.strip()


def para_text(p: ET.Element) -> str:
    runs: list[ET.Element] = []
    collect_runs(p, runs)
    return norm("".join(run_text(r) for r in runs))


def extract_units(path: str) -> list[str]:
    """Вернуть список логических блоков: абзацы и строки таблиц."""
    xml = zipfile.ZipFile(path).read("word/document.xml")
    body = ET.fromstring(xml).find(q("body"))
    units: list[str] = []
    if body is None:
        return units
    for block in body:
        if block.tag == q("p"):
            t = para_text(block)
            if t:
                units.append(t)
        elif block.tag == q("tbl"):
            for tr in block.findall(q("tr")):
                cells = []
                for tc in tr.findall(q("tc")):
                    ct = norm(" ".join(para_text(p) for p in tc.findall(q("p"))))
                    cells.append(ct)
                row = " | ".join(c for c in cells)
                if row.strip(" |"):
                    units.append("[таблица] " + row)
    return units


def main() -> None:
    a_path, b_path = sys.argv[1], sys.argv[2]
    a = extract_units(a_path)
    b = extract_units(b_path)
    print(f"Файл A: {a_path}  — блоков: {len(a)}")
    print(f"Файл B: {b_path}  — блоков: {len(b)}")
    print("=" * 80)

    sm = difflib.SequenceMatcher(a=a, b=b, autojunk=False)
    ratio = sm.ratio()
    n_diff = 0
    block_no = 0
    for tag, i1, i2, j1, j2 in sm.get_opcodes():
        if tag == "equal":
            continue
        block_no += 1
        print(f"\n### Различие #{block_no}  [{tag}]")
        if tag in ("replace", "delete"):
            for k in range(i1, i2):
                n_diff += 1
                print(f"  − [только в A, абзац {k+1}] {a[k]}")
        if tag in ("replace", "insert"):
            for k in range(j1, j2):
                n_diff += 1
                print(f"  + [только в B, абзац {k+1}] {b[k]}")

    print("\n" + "=" * 80)
    print(f"Схожесть последовательностей: {ratio*100:.2f}%")
    print(f"Блоков с различиями: {block_no}; строк-различий: {n_diff}")
    if block_no == 0:
        print("РАЗЛИЧИЙ НЕТ — содержимое идентично (после приёма правок и нормализации).")


if __name__ == "__main__":
    main()
