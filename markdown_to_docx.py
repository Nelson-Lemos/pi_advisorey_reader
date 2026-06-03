import re
from pathlib import Path
from typing import Optional
from docx import Document
from docx.shared import Pt, Inches, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement


def _set_cell_shading(cell, color_hex):
    shading = OxmlElement("w:shd")
    shading.set(qn("w:fill"), color_hex)
    shading.set(qn("w:val"), "clear")
    cell._tc.get_or_add_tcPr().append(shading)


def _add_hyperlink(paragraph, text, url):
    part = paragraph.part
    r_id = part.relate_to(url, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink", is_external=True)
    hyperlink = OxmlElement("w:hyperlink")
    hyperlink.set(qn("r:id"), r_id)
    new_run = OxmlElement("w:r")
    rPr = OxmlElement("w:rPr")
    c = OxmlElement("w:color")
    c.set(qn("w:val"), "0563C1")
    rPr.append(c)
    u = OxmlElement("w:u")
    u.set(qn("w:val"), "single")
    rPr.append(u)
    new_run.append(rPr)
    t = OxmlElement("w:t")
    t.text = text
    new_run.append(t)
    hyperlink.append(new_run)
    paragraph._p.append(hyperlink)


def _parse_inline(text, paragraph):
    pattern = re.compile(
        r'\[([^\]]+)\]\(([^)]+)\)'  # [text](url)
        r'|(?<!\*)\*\*(.+?)\*\*(?!\*)'  # **bold**
        r'|(?<!\*)\*(.+?)\*(?!\*)'  # *italic*
        r'|`([^`]+)`'  # `code`
    )
    pos = 0
    for m in pattern.finditer(text):
        start = m.start()
        if start > pos:
            run = paragraph.add_run(text[pos:start])
            run.font.size = Pt(11)
        pos = m.end()
        if m.group(1) is not None and m.group(2) is not None:
            _add_hyperlink(paragraph, m.group(1), m.group(2))
        elif m.group(3) is not None:
            run = paragraph.add_run(m.group(3))
            run.bold = True
            run.font.size = Pt(11)
        elif m.group(4) is not None:
            run = paragraph.add_run(m.group(4))
            run.italic = True
            run.font.size = Pt(11)
        elif m.group(5) is not None:
            run = paragraph.add_run(m.group(5))
            run.font.name = "Courier New"
            run.font.size = Pt(10)
            run.font.color.rgb = RGBColor(0xE0, 0x5C, 0x5C)
    if pos < len(text):
        run = paragraph.add_run(text[pos:])
        run.font.size = Pt(11)


def _parse_table(doc, lines, i):
    header_line = lines[i].strip()
    sep_line = lines[i + 1].strip() if i + 1 < len(lines) else ""
    if not re.match(r'^[\s\|:\-]+$', sep_line):
        return 0
    header_cells = [c.strip() for c in header_line.split("|") if c.strip()]
    num_cols = len(header_cells)
    if num_cols == 0:
        return 0
    rows_data = []
    j = i + 2
    while j < len(lines):
        line = lines[j].strip()
        if not line or line.startswith("#") or line.startswith("```") or line.startswith("---") or line.startswith("***"):
            break
        if "|" not in line:
            break
        cells = [c.strip() for c in line.split("|")]
        if len(cells) >= 2:
            cells = cells[1:-1] if cells[0] == "" and cells[-1] == "" else cells
            cells = [c.strip() for c in cells]
            while len(cells) < num_cols:
                cells.append("")
            rows_data.append(cells[:num_cols])
        j += 1
    table = doc.add_table(rows=1 + len(rows_data), cols=num_cols)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.style = "Table Grid"
    for ci, text in enumerate(header_cells):
        cell = table.rows[0].cells[ci]
        cell.text = ""
        p = cell.paragraphs[0]
        run = p.add_run(text)
        run.bold = True
        run.font.size = Pt(10)
        _set_cell_shading(cell, "2B2B2B")
    for ri, row_cells in enumerate(rows_data):
        for ci, text in enumerate(row_cells):
            cell = table.rows[ri + 1].cells[ci]
            cell.text = ""
            p = cell.paragraphs[0]
            _parse_inline(text, p)
            if ri % 2 == 1:
                _set_cell_shading(cell, "1E1E1E")
    return j - i


def _parse_code_block(doc, lines, i):
    lang_line = lines[i].strip()
    language = lang_line[3:].strip() if len(lang_line) > 3 else ""
    j = i + 1
    code_lines = []
    while j < len(lines):
        if lines[j].strip().startswith("```"):
            break
        code_lines.append(lines[j].rstrip())
        j += 1
    if code_lines:
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(4)
        p.paragraph_format.space_after = Pt(4)
        pPr = p._p.get_or_add_pPr()
        shd = OxmlElement("w:shd")
        shd.set(qn("w:val"), "clear")
        shd.set(qn("w:color"), "auto")
        shd.set(qn("w:fill"), "1E1E1E")
        pPr.append(shd)
        code_text = "\n".join(code_lines)
        run = p.add_run(code_text)
        run.font.name = "Consolas"
        run.font.size = Pt(9)
        run.font.color.rgb = RGBColor(0xD4, 0xD4, 0xD4)
        if language:
            p2 = doc.add_paragraph()
            run2 = p2.add_run(language)
            run2.font.size = Pt(8)
            run2.font.color.rgb = RGBColor(0x88, 0x88, 0x88)
            run2.italic = True
            p2.paragraph_format.space_before = Pt(0)
            p2.paragraph_format.space_after = Pt(2)
    return j - i + 1


def _parse_unordered_list(doc, lines, i):
    start = i
    while i < len(lines):
        stripped = lines[i].strip()
        if not stripped:
            i += 1
            continue
        if not re.match(r'^[\-\*\+]\s', stripped):
            break
        text = re.sub(r'^[\-\*\+]\s+', "", stripped)
        p = doc.add_paragraph(style="List Bullet")
        _parse_inline(text, p)
        i += 1
    return i - start


def _parse_ordered_list(doc, lines, i):
    start = i
    while i < len(lines):
        stripped = lines[i].strip()
        if not stripped:
            i += 1
            continue
        if not re.match(r'^\d+[\.\)]\s', stripped):
            break
        text = re.sub(r'^\d+[\.\)]\s+', "", stripped)
        p = doc.add_paragraph(style="List Number")
        _parse_inline(text, p)
        i += 1
    return i - start


def _parse_blockquote(doc, lines, i):
    start = i
    while i < len(lines):
        stripped = lines[i].strip()
        if not stripped.startswith(">"):
            break
        text = re.sub(r'^>\s?', "", stripped)
        p = doc.add_paragraph()
        p.paragraph_format.left_indent = Cm(1.27)
        pPr = p._p.get_or_add_pPr()
        shd = OxmlElement("w:shd")
        shd.set(qn("w:val"), "clear")
        shd.set(qn("w:color"), "auto")
        shd.set(qn("w:fill"), "1A1A2E")
        pPr.append(shd)
        side = OxmlElement("w:pBdr")
        left = OxmlElement("w:left")
        left.set(qn("w:val"), "single")
        left.set(qn("w:sz"), "12")
        left.set(qn("w:space"), "8")
        left.set(qn("w:color"), "7C5CFC")
        side.append(left)
        pPr.append(side)
        _parse_inline(text, p)
        i += 1
    return i - start


def markdown_to_docx(markdown_text, docx_path, images_dict=None):
    doc = Document()

    default_section = doc.sections[0]
    default_section.top_margin = Cm(2.54)
    default_section.bottom_margin = Cm(2.54)
    default_section.left_margin = Cm(3.17)
    default_section.right_margin = Cm(3.17)

    lines = markdown_text.split("\n")
    i = 0
    in_empty_block = False

    while i < len(lines):
        line = lines[i]
        stripped = line.strip()

        if not stripped:
            if not in_empty_block:
                doc.add_paragraph()
                in_empty_block = True
            i += 1
            continue
        in_empty_block = False

        if stripped.startswith("```"):
            consumed = _parse_code_block(doc, lines, i)
            i += consumed
            continue

        if stripped == "---" or stripped == "***" or stripped == "___":
            p = doc.add_paragraph()
            pPr = p._p.get_or_add_pPr()
            pBdr = OxmlElement("w:pBdr")
            bottom = OxmlElement("w:bottom")
            bottom.set(qn("w:val"), "single")
            bottom.set(qn("w:sz"), "6")
            bottom.set(qn("w:space"), "1")
            bottom.set(qn("w:color"), "444444")
            pBdr.append(bottom)
            pPr.append(pBdr)
            i += 1
            continue

        if "|" in stripped and i + 1 < len(lines) and re.match(r'^[\s\|:\-]+$', lines[i + 1].strip()):
            consumed = _parse_table(doc, lines, i)
            if consumed > 0:
                i += consumed
                continue

        heading_match = re.match(r'^(#{1,6})\s+(.+)$', stripped)
        if heading_match:
            level = len(heading_match.group(1))
            text = heading_match.group(2)
            p = doc.add_heading(text, level=min(level, 4))
            i += 1
            continue

        if re.match(r'^[\-\*\+]\s', stripped):
            consumed = _parse_unordered_list(doc, lines, i)
            if consumed > 0:
                i += consumed
                continue

        if re.match(r'^\d+[\.\)]\s', stripped):
            consumed = _parse_ordered_list(doc, lines, i)
            if consumed > 0:
                i += consumed
                continue

        if stripped.startswith(">"):
            consumed = _parse_blockquote(doc, lines, i)
            if consumed > 0:
                i += consumed
                continue

        img_match = re.match(r'^!\[([^\]]*)\]\(([^)]+)\)$', stripped)
        if img_match:
            alt = img_match.group(1)
            img_path = img_match.group(2)
            img_key = None
            if images_dict:
                for key in images_dict:
                    if key.endswith(img_path) or img_path in key:
                        img_key = key
                        break
            if img_key and images_dict and img_key in images_dict:
                img_data = images_dict[img_key]
                import io
                from docx.shared import Emu
                try:
                    image_stream = io.BytesIO(img_data)
                    p = doc.add_paragraph()
                    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    run = p.add_run()
                    run.add_picture(image_stream, width=Inches(5.0))
                except Exception:
                    p = doc.add_paragraph()
                    run = p.add_run(f"[{alt}]" if alt else "[IMAGEM]")
                    run.italic = True
                    run.font.color.rgb = RGBColor(0x88, 0x88, 0x88)
            else:
                actual_path = Path(img_path)
                if actual_path.exists():
                    p = doc.add_paragraph()
                    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    run = p.add_run()
                    run.add_picture(str(actual_path), width=Inches(5.0))
                else:
                    p = doc.add_paragraph()
                    run = p.add_run(f"[{alt}]" if alt else "[IMAGEM]")
                    run.italic = True
                    run.font.color.rgb = RGBColor(0x88, 0x88, 0x88)
            i += 1
            continue

        inline_img = re.match(r'^!\[([^\]]*)\]\(([^)]+)\)', stripped)
        if inline_img:
            i += 1
            continue

        p = doc.add_paragraph()
        _parse_in_text = stripped
        if _parse_in_text:
            _parse_inline(_parse_in_text, p)
        i += 1

    doc.save(str(docx_path))
