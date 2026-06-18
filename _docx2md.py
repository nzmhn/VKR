#!/usr/bin/env python3
"""Convert DOCX files to Markdown using only Python stdlib.

Handles:
- Headings (Heading 1..6)
- Paragraphs with bold/italic/underline/strike runs
- Bullet/numbered lists (basic)
- Tables (GitHub-flavored markdown)
- Hyperlinks
- Code-style runs
- Images: extracted to <basename>_assets/ and referenced with ![](...) (alt = filename)

Usage:
    python3 _docx2md.py input.docx [input2.docx ...]
"""
from __future__ import annotations

import os
import re
import sys
import zipfile
import xml.etree.ElementTree as ET

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
A = "http://schemas.openxmlformats.org/drawingml/2006/main"
PIC = "http://schemas.openxmlformats.org/drawingml/2006/picture"
WP = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"

NS = {"w": W, "r": R, "a": A, "pic": PIC, "wp": WP}


def qn(tag: str) -> str:
    prefix, local = tag.split(":", 1)
    return f"{{{NS[prefix]}}}{local}"


def load_relationships(z: zipfile.ZipFile) -> dict:
    rels = {}
    try:
        data = z.read("word/_rels/document.xml.rels")
    except KeyError:
        return rels
    root = ET.fromstring(data)
    for rel in root:
        rid = rel.attrib.get("Id")
        target = rel.attrib.get("Target")
        rtype = rel.attrib.get("Type", "")
        rels[rid] = {"target": target, "type": rtype}
    return rels


def load_numbering(z: zipfile.ZipFile) -> dict:
    """Return mapping numId -> {ilvl -> 'bullet'|'decimal'}"""
    out = {}
    try:
        data = z.read("word/numbering.xml")
    except KeyError:
        return out
    root = ET.fromstring(data)
    abstract = {}
    for an in root.findall(qn("w:abstractNum")):
        aid = an.attrib.get(qn("w:abstractNumId"))
        levels = {}
        for lvl in an.findall(qn("w:lvl")):
            ilvl = lvl.attrib.get(qn("w:ilvl"))
            fmt_el = lvl.find(qn("w:numFmt"))
            fmt = fmt_el.attrib.get(qn("w:val")) if fmt_el is not None else "bullet"
            levels[ilvl] = fmt
        abstract[aid] = levels
    for n in root.findall(qn("w:num")):
        nid = n.attrib.get(qn("w:numId"))
        ref = n.find(qn("w:abstractNumId"))
        if ref is not None:
            aid = ref.attrib.get(qn("w:val"))
            out[nid] = abstract.get(aid, {})
    return out


def text_of_run(r) -> str:
    parts = []
    for child in r.iter():
        tag = child.tag
        if tag == qn("w:t"):
            parts.append(child.text or "")
        elif tag == qn("w:tab"):
            parts.append("\t")
        elif tag == qn("w:br"):
            parts.append("\n")
    return "".join(parts)


def run_props(r):
    rpr = r.find(qn("w:rPr"))
    bold = italic = underline = strike = code = False
    if rpr is not None:
        if rpr.find(qn("w:b")) is not None:
            bold = True
        if rpr.find(qn("w:i")) is not None:
            italic = True
        if rpr.find(qn("w:u")) is not None:
            underline = True
        if rpr.find(qn("w:strike")) is not None:
            strike = True
        rfonts = rpr.find(qn("w:rFonts"))
        if rfonts is not None:
            ascii_font = (rfonts.attrib.get(qn("w:ascii")) or "").lower()
            if "mono" in ascii_font or "consolas" in ascii_font or "courier" in ascii_font:
                code = True
    return bold, italic, underline, strike, code


def render_run(r) -> str:
    txt = text_of_run(r)
    if not txt:
        return ""
    bold, italic, underline, strike, code = run_props(r)
    # Escape markdown specials minimally
    txt = txt.replace("\\", "\\\\")
    if code:
        txt = f"`{txt}`"
        return txt
    if strike:
        txt = f"~~{txt}~~"
    if bold and italic:
        txt = f"***{txt}***"
    elif bold:
        txt = f"**{txt}**"
    elif italic:
        txt = f"*{txt}*"
    return txt


def extract_image_from_drawing(drawing, rels, z, assets_dir, base_assets_rel):
    """Find image reference in a w:drawing element, save the image file, return md ref."""
    # blip is at a:blip or pic:blipFill/a:blip
    blip = None
    for el in drawing.iter():
        if el.tag.endswith("}blip"):
            blip = el
            break
    if blip is None:
        return ""
    rid = blip.attrib.get(qn("r:embed")) or blip.attrib.get(qn("r:link"))
    if not rid or rid not in rels:
        return ""
    target = rels[rid]["target"]
    # target like "media/image1.png"
    full = "word/" + target if not target.startswith("word/") else target
    try:
        data = z.read(full)
    except KeyError:
        return ""
    fname = os.path.basename(target)
    os.makedirs(assets_dir, exist_ok=True)
    out_path = os.path.join(assets_dir, fname)
    with open(out_path, "wb") as f:
        f.write(data)
    return f"![{fname}]({base_assets_rel}/{fname})"


def render_paragraph(p, rels, z, numbering, assets_dir, base_assets_rel):
    pPr = p.find(qn("w:pPr"))
    style = ""
    indent_level = 0
    list_fmt = None
    if pPr is not None:
        ps = pPr.find(qn("w:pStyle"))
        if ps is not None:
            style = (ps.attrib.get(qn("w:val")) or "").lower()
        numPr = pPr.find(qn("w:numPr"))
        if numPr is not None:
            ilvl_el = numPr.find(qn("w:ilvl"))
            numId_el = numPr.find(qn("w:numId"))
            ilvl = ilvl_el.attrib.get(qn("w:val")) if ilvl_el is not None else "0"
            numId = numId_el.attrib.get(qn("w:val")) if numId_el is not None else None
            indent_level = int(ilvl) if ilvl.isdigit() else 0
            if numId and numId in numbering:
                list_fmt = numbering[numId].get(ilvl, "bullet")
            else:
                list_fmt = "bullet"

    # Build inline content (runs, hyperlinks, drawings)
    parts = []

    def walk(el):
        for child in el:
            tag = child.tag
            if tag == qn("w:r"):
                # Check for drawing inside run
                drawing = child.find(qn("w:drawing"))
                if drawing is not None:
                    img_md = extract_image_from_drawing(
                        drawing, rels, z, assets_dir, base_assets_rel
                    )
                    if img_md:
                        parts.append(img_md)
                    # also include any text in same run
                txt = render_run(child)
                if txt:
                    parts.append(txt)
            elif tag == qn("w:hyperlink"):
                rid = child.attrib.get(qn("r:id"))
                inner = []
                for r in child.findall(qn("w:r")):
                    inner.append(render_run(r))
                inner_text = "".join(inner).strip()
                if rid and rid in rels:
                    href = rels[rid]["target"]
                    parts.append(f"[{inner_text}]({href})")
                else:
                    parts.append(inner_text)
            elif tag == qn("w:smartTag") or tag == qn("w:fldSimple"):
                walk(child)
            elif tag == qn("w:ins") or tag == qn("w:del"):
                walk(child)

    walk(p)
    text = "".join(parts).strip()

    # Heading detection
    m = re.match(r"heading(\d+)", style)
    if m:
        level = max(1, min(6, int(m.group(1))))
        if not text:
            return ""
        return "#" * level + " " + text
    if style == "title":
        return "# " + text if text else ""
    if style == "subtitle":
        return "## " + text if text else ""

    # List
    if list_fmt is not None:
        prefix = "  " * indent_level
        if list_fmt == "bullet":
            return f"{prefix}- {text}"
        else:
            return f"{prefix}1. {text}"

    return text


def render_table(tbl, rels, z, numbering, assets_dir, base_assets_rel):
    rows = []
    for tr in tbl.findall(qn("w:tr")):
        cells = []
        for tc in tr.findall(qn("w:tc")):
            cell_lines = []
            for p in tc.findall(qn("w:p")):
                line = render_paragraph(
                    p, rels, z, numbering, assets_dir, base_assets_rel
                )
                if line:
                    cell_lines.append(line)
            cell_text = " <br> ".join(cell_lines).replace("|", "\\|")
            cells.append(cell_text or " ")
        rows.append(cells)
    if not rows:
        return ""
    width = max(len(r) for r in rows)
    rows = [r + [" "] * (width - len(r)) for r in rows]
    out = []
    out.append("| " + " | ".join(rows[0]) + " |")
    out.append("| " + " | ".join(["---"] * width) + " |")
    for r in rows[1:]:
        out.append("| " + " | ".join(r) + " |")
    return "\n".join(out)


def convert(docx_path: str) -> str:
    base = os.path.splitext(os.path.basename(docx_path))[0]
    out_dir = os.path.dirname(os.path.abspath(docx_path))
    assets_rel = f"{base}_assets"
    assets_dir = os.path.join(out_dir, assets_rel)

    with zipfile.ZipFile(docx_path) as z:
        rels = load_relationships(z)
        numbering = load_numbering(z)
        doc_xml = z.read("word/document.xml")
        root = ET.fromstring(doc_xml)
        body = root.find(qn("w:body"))

        blocks = []
        for child in body:
            tag = child.tag
            if tag == qn("w:p"):
                line = render_paragraph(
                    child, rels, z, numbering, assets_dir, assets_rel
                )
                blocks.append(line)
            elif tag == qn("w:tbl"):
                tbl_md = render_table(
                    child, rels, z, numbering, assets_dir, assets_rel
                )
                if tbl_md:
                    blocks.append("")
                    blocks.append(tbl_md)
                    blocks.append("")
            elif tag == qn("w:sectPr"):
                continue

    # Collapse multiple blank lines
    md = "\n\n".join(b for b in blocks if b is not None)
    md = re.sub(r"\n{3,}", "\n\n", md)
    # Merge adjacent identical inline markers from neighbouring runs:
    # **a****b**     -> **ab**
    # *a**b*         -> *ab*  (only between same single asterisks)
    # ***a******b*** -> ***ab***
    md = re.sub(r"\*{6}", "", md)
    md = re.sub(r"\*{4}", "", md)
    md = re.sub(r"(?<!\*)\*{2}(?!\*)(?<=\S)\s*(?<!\*)\*{2}(?!\*)", "", md)
    # Drop empty bold/italic groups (only when there's whitespace between markers)
    md = re.sub(r"\*\*\s+\*\*", "", md)
    md = re.sub(r"(?<!\*)\*\s+\*(?!\*)", "", md)
    md = md.strip() + "\n"

    out_path = os.path.join(out_dir, base + ".md")
    with open(out_path, "w", encoding="utf-8") as f:
        f.write(md)
    return out_path


def main():
    paths = sys.argv[1:]
    if not paths:
        print("usage: _docx2md.py <docx>...", file=sys.stderr)
        sys.exit(1)
    for p in paths:
        try:
            out = convert(p)
            print(f"OK  {p} -> {out}")
        except Exception as e:
            print(f"ERR {p}: {e}", file=sys.stderr)
            raise


if __name__ == "__main__":
    main()
