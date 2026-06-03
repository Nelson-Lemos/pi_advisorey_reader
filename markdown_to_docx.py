# -*- coding: utf-8 -*-
"""
markdown_to_docx.py  v3.0
Converte Markdown (output do Datalab/Marker) para DOCX bem formatado.
- Limpa tags HTML residuais sem remover < > matemáticos
- Tabelas com cabeçalho colorido e linhas alternadas
- Títulos hierárquicos com formatação inline preservada
- Listas, citações, código, hiperligações, imagens, placeholders
- Separadores e quebras de página
"""

import re
import io
from pathlib import Path
from docx import Document
from docx.shared import Pt, Inches, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

# ── Constantes de estilo ───────────────────────────────────────────────
HEADING_COLORS = {
    1: RGBColor(0x1F, 0x38, 0x64),
    2: RGBColor(0x2E, 0x74, 0xB5),
    3: RGBColor(0x2E, 0x74, 0xB5),
    4: RGBColor(0x40, 0x40, 0x40),
    5: RGBColor(0x60, 0x60, 0x60),
    6: RGBColor(0x80, 0x80, 0x80),
}
HEADING_SPACING = {1: (14, 8), 2: (12, 6), 3: (10, 4), 4: (8, 3), 5: (6, 2), 6: (4, 2)}
HEADING_FONT_SIZES = {1: 20, 2: 16, 3: 14, 4: 12, 5: 11, 6: 10}

# Tags HTML que convertemos para markdown antes da limpeza final
_HTML_TO_MD = [
    (re.compile(r"<br\s*/?>", re.IGNORECASE), "\n"),
    (re.compile(r"<hr\s*/?>", re.IGNORECASE), "\n---\n"),
    (re.compile(r"<(?:b|strong)>(.*?)</(?:b|strong)>", re.IGNORECASE | re.DOTALL), r"**\1**"),
    (re.compile(r"<(?:i|em)>(.*?)</(?:i|em)>", re.IGNORECASE | re.DOTALL), r"*\1*"),
    (re.compile(r"<code>(.*?)</code>", re.IGNORECASE | re.DOTALL), r"`\1`"),
    (re.compile(r"<u>(.*?)</u>", re.IGNORECASE | re.DOTALL), r"__\1__"),
    (re.compile(r"<s>(.*?)</s>", re.IGNORECASE | re.DOTALL), r"~~\1~~"),
]

# Expressão para detetar tags HTML válidas (exclui <, >, <=, >=, etc.)
_TAG_RE = re.compile(r"</?[a-zA-Z][a-zA-Z0-9]*\b[^>]*>")


def _clean_html(text: str) -> str:
    """Remove tags HTML residuais sem remover < > matemáticos."""
    text = text.replace("&amp;", "&")
    text = text.replace("&lt;", "<")
    text = text.replace("&gt;", ">")
    text = text.replace("&nbsp;", " ")
    text = text.replace("&quot;", '"')
    text = text.replace("&#39;", "'")
    text = text.replace("&apos;", "'")

    for pattern, replacement in _HTML_TO_MD:
        text = pattern.sub(replacement, text)

    text = _TAG_RE.sub("", text)
    text = re.sub(r"[ \t]{2,}", " ", text)
    return text.strip()


# ── Helpers XML ────────────────────────────────────────────────────────

def _set_cell_bg(cell, color_hex: str):
    shd = OxmlElement("w:shd")
    shd.set(qn("w:fill"), color_hex)
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    cell._tc.get_or_add_tcPr().append(shd)


def _set_cell_border(cell, color_hex="CCCCCC"):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    tcBorders = OxmlElement("w:tcBorders")
    for side in ("top", "left", "bottom", "right"):
        border = OxmlElement(f"w:{side}")
        border.set(qn("w:val"), "single")
        border.set(qn("w:sz"), "4")
        border.set(qn("w:space"), "0")
        border.set(qn("w:color"), color_hex)
        tcBorders.append(border)
    tcPr.append(tcBorders)


def _add_hyperlink(paragraph, text: str, url: str):
    part = paragraph.part
    r_id = part.relate_to(
        url,
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink",
        is_external=True,
    )
    hl = OxmlElement("w:hyperlink")
    hl.set(qn("r:id"), r_id)
    r = OxmlElement("w:r")
    rPr = OxmlElement("w:rPr")
    color = OxmlElement("w:color")
    color.set(qn("w:val"), "0563C1")
    rPr.append(color)
    u = OxmlElement("w:u")
    u.set(qn("w:val"), "single")
    rPr.append(u)
    sz = OxmlElement("w:sz")
    sz.set(qn("w:val"), "22")
    rPr.append(sz)
    r.append(rPr)
    t = OxmlElement("w:t")
    t.text = text
    r.append(t)
    hl.append(r)
    paragraph._p.append(hl)


def _set_para_spacing(para, before=0, after=4):
    para.paragraph_format.space_before = Pt(before)
    para.paragraph_format.space_after = Pt(after)


def _add_page_break(doc):
    p = doc.add_paragraph()
    run = p.add_run()
    run.add_break(docx.enum.text.WD_BREAK.PAGE)


# ── Inline formatting ─────────────────────────────────────────────────

_INLINE_RE = re.compile(
    r"\[(carimbo|assinatura|logo\s*marca)\]"
    r"|\[([^\]]+)\]\(([^)]+)\)"
    r"|\*\*\*(.+?)\*\*\*"
    r"|\*\*(.+?)\*\*"
    r"|__(.+?)__"
    r"|~~(.+?)~~"
    r"|\*(.+?)\*"
    r"|`([^`]+)`"
)


def _parse_inline(text: str, paragraph, font_size=11):
    """Aplica formatação inline: bold, italic, code, hiperligações, placeholders."""
    text = _clean_html(text)
    pos = 0
    for m in _INLINE_RE.finditer(text):
        if m.start() > pos:
            run = paragraph.add_run(text[pos : m.start()])
            run.font.size = Pt(font_size)
        pos = m.end()

        if m.group(1):
            ptype = m.group(1)
            lbl = f"[{ptype}]"
            run = paragraph.add_run(lbl)
            run.bold = True
            run.font.size = Pt(font_size)
            if ptype == "carimbo":
                run.font.color.rgb = RGBColor(0xC0, 0x39, 0x2B)
            elif ptype == "assinatura":
                run.font.color.rgb = RGBColor(0x2E, 0x86, 0xC1)
            elif "marca" in ptype or ptype == "logo":
                run.font.color.rgb = RGBColor(0x6C, 0x6C, 0x6C)
        elif m.group(2) and m.group(3):
            _add_hyperlink(paragraph, m.group(2), m.group(3))
        elif m.group(4):
            run = paragraph.add_run(m.group(4))
            run.bold = True
            run.italic = True
            run.font.size = Pt(font_size)
        elif m.group(5):
            run = paragraph.add_run(m.group(5))
            run.bold = True
            run.font.size = Pt(font_size)
        elif m.group(6):
            run = paragraph.add_run(m.group(6))
            run.underline = True
            run.font.size = Pt(font_size)
        elif m.group(7):
            run = paragraph.add_run(m.group(7))
            run.font.strike = True
            run.font.size = Pt(font_size)
        elif m.group(8):
            run = paragraph.add_run(m.group(8))
            run.italic = True
            run.font.size = Pt(font_size)
        elif m.group(9):
            run = paragraph.add_run(m.group(9))
            run.font.name = "Courier New"
            run.font.size = Pt(font_size - 1)
            run.font.color.rgb = RGBColor(0xC7, 0x25, 0x4E)

    if pos < len(text):
        run = paragraph.add_run(text[pos:])
        run.font.size = Pt(font_size)


# ── Table parser ──────────────────────────────────────────────────────

def _split_table_row(line: str):
    """Divide uma linha de tabela markdown em células, ignorando pipes extremos."""
    parts = [c.strip() for c in line.split("|")]
    if parts and parts[0] == "":
        parts = parts[1:]
    if parts and parts[-1] == "":
        parts = parts[:-1]
    return parts


def _parse_table(doc: Document, lines: list, i: int) -> int:
    """Renderiza uma tabela markdown com cabeçalho, bordas e linhas alternadas."""
    header_raw = _clean_html(lines[i].strip())
    sep_line = lines[i + 1].strip() if i + 1 < len(lines) else ""

    if not re.match(r"^[\s|:\-]+$", sep_line) or "|" not in sep_line:
        return 0

    header_cells = _split_table_row(header_raw)
    num_cols = len(header_cells)
    if num_cols == 0:
        return 0

    col_aligns = []
    for cell in _split_table_row(sep_line):
        c = cell.strip()
        if c.startswith(":") and c.endswith(":"):
            col_aligns.append("center")
        elif c.endswith(":"):
            col_aligns.append("right")
        else:
            col_aligns.append("left")
    while len(col_aligns) < num_cols:
        col_aligns.append("left")

    rows_data = []
    j = i + 2
    while j < len(lines):
        line = lines[j].strip()
        if not line or "|" not in line:
            break
        if line.startswith("#") or line.startswith("```"):
            break
        cells = _split_table_row(line)
        while len(cells) < num_cols:
            cells.append("")
        rows_data.append(cells[:num_cols])
        j += 1

    table = doc.add_table(rows=1 + len(rows_data), cols=num_cols)
    table.style = "Table Grid"
    table.alignment = WD_TABLE_ALIGNMENT.LEFT
    table.autofit = True

    for ci, cell_text in enumerate(header_cells):
        cell = table.rows[0].cells[ci]
        cell.text = ""
        p = cell.paragraphs[0]
        p.paragraph_format.space_before = Pt(2)
        p.paragraph_format.space_after = Pt(2)
        _parse_inline(cell_text, p, font_size=10)
        for run in p.runs:
            run.bold = True
            run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
        _set_cell_bg(cell, "1F3864")
        _set_cell_border(cell, "1F3864")

    for ri, row_cells in enumerate(rows_data):
        bg = "EBF3FB" if ri % 2 == 0 else "FFFFFF"
        for ci, cell_text in enumerate(row_cells):
            cell = table.rows[ri + 1].cells[ci]
            cell.text = ""
            p = cell.paragraphs[0]
            p.paragraph_format.space_before = Pt(1)
            p.paragraph_format.space_after = Pt(1)
            _parse_inline(cell_text, p, font_size=10)
            _set_cell_bg(cell, bg)
            _set_cell_border(cell, "BDD7EE")

    return j - i


# ── Code block ────────────────────────────────────────────────────────

def _parse_code_block(doc: Document, lines: list, i: int) -> int:
    lang = lines[i].strip()[3:].strip()
    j = i + 1
    code_lines = []
    while j < len(lines):
        if lines[j].strip().startswith("```"):
            j += 1
            break
        code_lines.append(lines[j])
        j += 1

    if code_lines:
        p = doc.add_paragraph()
        pPr = p._p.get_or_add_pPr()
        shd = OxmlElement("w:shd")
        shd.set(qn("w:val"), "clear")
        shd.set(qn("w:color"), "auto")
        shd.set(qn("w:fill"), "F4F4F4")
        pPr.append(shd)
        p.paragraph_format.left_indent = Cm(0.5)
        p.paragraph_format.space_before = Pt(6)
        p.paragraph_format.space_after = Pt(6)
        run = p.add_run("\n".join(code_lines))
        run.font.name = "Courier New"
        run.font.size = Pt(9)
        run.font.color.rgb = RGBColor(0x24, 0x29, 0x2E)
        if lang:
            p2 = doc.add_paragraph()
            r2 = p2.add_run(f"  {lang}")
            r2.font.size = Pt(8)
            r2.font.color.rgb = RGBColor(0x88, 0x88, 0x88)
            r2.italic = True
            p2.paragraph_format.space_before = Pt(0)
            p2.paragraph_format.space_after = Pt(4)
    return j - i


# ── List parsers ──────────────────────────────────────────────────────

def _parse_unordered_list(doc: Document, lines: list, i: int) -> int:
    start = i
    while i < len(lines):
        s = lines[i].strip()
        if not s:
            i += 1
            continue
        if not re.match(r"^[\-\*\+]\s", s):
            break
        text = re.sub(r"^[\-\*\+]\s+", "", s)
        p = doc.add_paragraph(style="List Bullet")
        _parse_inline(text, p)
        _set_para_spacing(p, 0, 2)
        i += 1
    return i - start


def _parse_ordered_list(doc: Document, lines: list, i: int) -> int:
    start = i
    while i < len(lines):
        s = lines[i].strip()
        if not s:
            i += 1
            continue
        if not re.match(r"^\d+[\.\)]\s", s):
            break
        text = re.sub(r"^\d+[\.\)]\s+", "", s)
        p = doc.add_paragraph(style="List Number")
        _parse_inline(text, p)
        _set_para_spacing(p, 0, 2)
        i += 1
    return i - start


# ── Blockquote ────────────────────────────────────────────────────────

def _parse_blockquote(doc: Document, lines: list, i: int) -> int:
    while i < len(lines):
        s = lines[i].strip()
        if not s.startswith(">"):
            break
        text = re.sub(r"^>\s?", "", s)
        p = doc.add_paragraph()
        p.paragraph_format.left_indent = Cm(1.0)
        pPr = p._p.get_or_add_pPr()
        shd = OxmlElement("w:shd")
        shd.set(qn("w:val"), "clear")
        shd.set(qn("w:color"), "auto")
        shd.set(qn("w:fill"), "F0F4FF")
        pPr.append(shd)
        bdr = OxmlElement("w:pBdr")
        left = OxmlElement("w:left")
        left.set(qn("w:val"), "single")
        left.set(qn("w:sz"), "16")
        left.set(qn("w:space"), "8")
        left.set(qn("w:color"), "4472C4")
        bdr.append(left)
        pPr.append(bdr)
        _parse_inline(text, p)
        _set_para_spacing(p, 2, 2)
        i += 1
    return 1


# ── Heading parser ────────────────────────────────────────────────────

def _parse_heading(doc: Document, text: str, level: int):
    """Cria um título com formatação inline preservada."""
    heading = doc.add_heading(level=level)
    heading.clear()
    _parse_inline(text, heading, font_size=HEADING_FONT_SIZES.get(level, 12))
    for run in heading.runs:
        run.font.color.rgb = HEADING_COLORS.get(level, HEADING_COLORS[4])
    bef, aft = HEADING_SPACING.get(level, (6, 2))
    _set_para_spacing(heading, bef, aft)


# ── Placeholder / Image rendered ──────────────────────────────────────

def _render_image(doc, alt: str, img_path: str, images_dict: dict = None) -> bool:
    """Tenta inserir uma imagem a partir de images_dict, caminho local, ou fallback."""
    inserted = False
    if images_dict:
        key = next(
            (k for k in images_dict if k.endswith(img_path) or img_path in k),
            None,
        )
        if key:
            try:
                data = images_dict[key]
                if isinstance(data, str):
                    import base64
                    data = base64.b64decode(data)
                p = doc.add_paragraph()
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                p.add_run().add_picture(io.BytesIO(data), width=Inches(5.0))
                inserted = True
            except Exception:
                pass
    if not inserted:
        actual = Path(img_path)
        if actual.exists():
            try:
                p = doc.add_paragraph()
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                p.add_run().add_picture(str(actual), width=Inches(5.0))
                inserted = True
            except Exception:
                pass
    if not inserted:
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = p.add_run(f"[{alt or 'IMAGEM'}]")
        run.italic = True
        run.font.color.rgb = RGBColor(0x88, 0x88, 0x88)
    return True


# ── Separador horizontal ──────────────────────────────────────────────

def _render_horizontal_rule(doc):
    p = doc.add_paragraph()
    pPr = p._p.get_or_add_pPr()
    pBdr = OxmlElement("w:pBdr")
    bot = OxmlElement("w:bottom")
    bot.set(qn("w:val"), "single")
    bot.set(qn("w:sz"), "6")
    bot.set(qn("w:space"), "1")
    bot.set(qn("w:color"), "AAAAAA")
    pBdr.append(bot)
    pPr.append(pBdr)
    _set_para_spacing(p, 4, 4)


# ── Main converter ────────────────────────────────────────────────────

def markdown_to_docx(markdown_text: str, docx_path: str, images_dict: dict = None):
    """
    Converte Markdown para DOCX formatado.
    markdown_text : texto markdown (output do Datalab/Marker)
    docx_path     : caminho de saída .docx
    images_dict   : {nome: bytes} — imagens (opcional)
    """
    markdown_text = re.sub(r"<!--.*?-->", "", markdown_text, flags=re.DOTALL)

    doc = Document()
    section = doc.sections[0]
    section.top_margin = Cm(2.54)
    section.bottom_margin = Cm(2.54)
    section.left_margin = Cm(2.54)
    section.right_margin = Cm(2.54)

    default_font = doc.styles["Normal"].font
    default_font.name = "Calibri"
    default_font.size = Pt(11)

    for level in range(1, 7):
        style = doc.styles[f"Heading {level}"]
        style.font.color.rgb = HEADING_COLORS.get(level, HEADING_COLORS[4])
        style.font.bold = True

    lines = markdown_text.split("\n")
    i = 0
    had_content = False

    while i < len(lines):
        raw = lines[i]
        stripped = raw.strip()

        if not stripped:
            if had_content:
                doc.add_paragraph()
            i += 1
            continue

        had_content = True

        if stripped.startswith("```"):
            consumed = _parse_code_block(doc, lines, i)
            i += consumed if consumed > 0 else 1
            continue

        if re.match(r"^[-\*_]{3,}$", stripped):
            _render_horizontal_rule(doc)
            i += 1
            continue

        if "|" in stripped and i + 1 < len(lines):
            next_stripped = lines[i + 1].strip()
            if re.match(r"^[\s|:\-]+$", next_stripped) and "|" in next_stripped:
                consumed = _parse_table(doc, lines, i)
                if consumed > 0:
                    i += consumed
                    continue

        hm = re.match(r"^(#{1,6})\s+(.+)$", stripped)
        if hm:
            level = len(hm.group(1))
            text = hm.group(2)
            _parse_heading(doc, text, level)
            i += 1
            continue

        if re.match(r"^[\-\*\+]\s", stripped):
            consumed = _parse_unordered_list(doc, lines, i)
            if consumed > 0:
                i += consumed
                continue

        if re.match(r"^\d+[\.\)]\s", stripped):
            consumed = _parse_ordered_list(doc, lines, i)
            if consumed > 0:
                i += consumed
                continue

        if stripped.startswith(">"):
            consumed = _parse_blockquote(doc, lines, i)
            if consumed > 0:
                i += 1
                continue

        img_m = re.match(r"^!\[([^\]]*)\]\(([^)]+)\)$", stripped)
        if img_m:
            _render_image(doc, img_m.group(1), img_m.group(2), images_dict)
            i += 1
            continue

        ph_m = re.match(r"^\[(carimbo|assinatura|logo\s*marca)\]$", stripped)
        if ph_m:
            ptype = ph_m.group(1)
            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            pPr = p._p.get_or_add_pPr()
            shd = OxmlElement("w:shd")
            shd.set(qn("w:val"), "clear")
            shd.set(qn("w:color"), "auto")
            bg_colors = {
                "carimbo": "FFF5F5",
                "assinatura": "F0F8FF",
                "logo marca": "F5F5F5",
            }
            shd.set(qn("w:fill"), bg_colors.get(ptype, "F9F9F9"))
            pPr.append(shd)
            bdr = OxmlElement("w:pBdr")
            left = OxmlElement("w:left")
            left.set(qn("w:val"), "single")
            left.set(qn("w:sz"), "12")
            left.set(qn("w:space"), "6")
            border_colors = {
                "carimbo": "C0392B",
                "assinatura": "2E86C1",
                "logo marca": "6C6C6C",
            }
            left.set(qn("w:color"), border_colors.get(ptype, "999999"))
            bdr.append(left)
            pPr.append(bdr)
            _parse_inline(stripped, p, font_size=10)
            _set_para_spacing(p, 3, 3)
            i += 1
            continue

        p = doc.add_paragraph()
        _parse_inline(stripped, p, font_size=11)
        _set_para_spacing(p, 0, 4)
        i += 1

    doc.save(str(docx_path))
