#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Перенос правок из «вкр_Эльвина_правки.docx» в «ВКР ГОТОВОЕ.docx».

Логика:
  * из обоих документов извлекаются логические блоки (абзацы и строки таблиц)
    в двух видах — «сырой» (raw, как в XML, правки приняты) и нормализованный
    (для сопоставления через difflib);
  * difflib выравнивает блоки; для каждого блока-замены формируется пара
    (старый сырой текст -> новый сырой текст);
  * пары применяются к word/document.xml документа «ВКР ГОТОВОЕ» как точечные
    строковые замены `>OLD<` -> `>NEW<` (абзацы здесь — одиночные run-ы, текст
    непрерывен), с проверкой, что каждая OLD-строка встречается ровно один раз.

Режимы:
  python3 _apply_pravki_to_gotovoe.py            # dry-run: только отчёт
  python3 _apply_pravki_to_gotovoe.py --write     # записать ВКР ГОТОВОЕ.docx (+.md)
"""
from __future__ import annotations

import re
import sys
import shutil
import zipfile
import difflib
import xml.etree.ElementTree as ET
from pathlib import Path

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

GOTOVOE = "ВКР ГОТОВОЕ.docx"
PRAVKI = "вкр_Эльвина_правки.docx"
GOTOVOE_MD = "ВКР ГОТОВОЕ.md"


def q(tag: str) -> str:
    return f"{{{W}}}{tag}"


def run_text(r: ET.Element) -> str:
    parts: list[str] = []
    for el in r.iter():
        if el.tag == q("t"):
            parts.append(el.text or "")
        elif el.tag in (q("tab"), q("br"), q("cr")):
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
    return re.sub(r"\s+", " ", s.replace("\u00a0", " ")).strip()


def para_raw(p: ET.Element) -> str:
    runs: list[ET.Element] = []
    collect_runs(p, runs)
    return "".join(run_text(r) for r in runs)


def extract(path: str) -> list[str]:
    """Список сырых текстов блоков (абзацы и строки таблиц)."""
    xml = zipfile.ZipFile(path).read("word/document.xml")
    body = ET.fromstring(xml).find(q("body"))
    units: list[str] = []
    if body is None:
        return units
    for block in body:
        if block.tag == q("p"):
            t = para_raw(block)
            if norm(t):
                units.append(t)
        elif block.tag == q("tbl"):
            for tr in block.findall(q("tr")):
                cells = [norm(" ".join(para_raw(p) for p in tc.findall(q("p"))))
                         for tc in tr.findall(q("tc"))]
                if norm(" ".join(cells)):
                    # для таблиц правок здесь нет точечных замен на уровне строки,
                    # но блок нужен для корректного выравнивания difflib
                    units.append("\u0001TBL\u0001" + " | ".join(cells))
    return units


def build_pairs() -> list[tuple[str, str]]:
    a = extract(GOTOVOE)
    b = extract(PRAVKI)
    na = [norm(x) for x in a]
    nb = [norm(x) for x in b]
    sm = difflib.SequenceMatcher(a=na, b=nb, autojunk=False)
    pairs: list[tuple[str, str]] = []
    for tag, i1, i2, j1, j2 in sm.get_opcodes():
        if tag != "replace":
            if tag in ("insert", "delete"):
                print(f"  ВНИМАНИЕ: блок {tag} A[{i1}:{i2}] B[{j1}:{j2}] — "
                      f"требует ручной обработки")
            continue
        if (i2 - i1) != (j2 - j1):
            print(f"  ВНИМАНИЕ: replace разной длины A[{i1}:{i2}] B[{j1}:{j2}]")
        # позиционное сопоставление внутри блока замены
        for off in range(min(i2 - i1, j2 - j1)):
            old_raw, new_raw = a[i1 + off], b[j1 + off]
            if old_raw.startswith("\u0001TBL\u0001") or new_raw.startswith("\u0001TBL\u0001"):
                # строка таблицы: вытащим изменившуюся ячейку отдельно ниже
                continue
            if norm(old_raw) != norm(new_raw):
                pairs.append((old_raw, new_raw))
    return pairs


# Точечная правка внутри таблицы (Различие #8): убрать слово «лобовой».
TABLE_EDITS = [
    ("Прямой лобовой конкуренции мало", "Прямой конкуренции мало"),
]


def main() -> None:
    write = "--write" in sys.argv
    pairs = build_pairs() + TABLE_EDITS

    xml = zipfile.ZipFile(GOTOVOE).read("word/document.xml").decode("utf-8")
    md = Path(GOTOVOE_MD).read_text(encoding="utf-8")

    print(f"Всего пар замен: {len(pairs)}")
    print("=" * 80)
    ok = True
    for n, (old, new) in enumerate(pairs, 1):
        c_xml = xml.count(old)
        c_md = md.count(old)
        flag = "" if c_xml == 1 else "  <-- ПРОБЛЕМА (xml != 1)"
        print(f"\n[{n}] вхождений: xml={c_xml}, md={c_md}{flag}")
        print(f"   OLD: {norm(old)[:140]}")
        print(f"   NEW: {norm(new)[:140]}")
        if c_xml != 1:
            ok = False

    print("\n" + "=" * 80)
    if not ok:
        print("Есть проблемы с однозначностью — запись НЕ выполняется.")
        sys.exit(1)

    if not write:
        print("DRY-RUN ок. Для записи запустите с флагом --write")
        return

    # применяем к XML
    for old, new in pairs:
        xml = xml.replace(old, new, 1)
    # применяем к .md (где найдено)
    for old, new in pairs:
        if md.count(old) >= 1:
            md = md.replace(old, new, 1)

    # пересобираем docx: копия архива с заменой document.xml
    src = zipfile.ZipFile(GOTOVOE)
    tmp = GOTOVOE + ".tmp"
    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as out:
        for item in src.infolist():
            data = src.read(item.filename)
            if item.filename == "word/document.xml":
                data = xml.encode("utf-8")
            out.writestr(item, data)
    src.close()
    shutil.move(tmp, GOTOVOE)
    Path(GOTOVOE_MD).write_text(md, encoding="utf-8")
    print(f"Записано: {GOTOVOE} и {GOTOVOE_MD}")


if __name__ == "__main__":
    main()
