# -*- coding: utf-8 -*-
"""
markdown_to_docx.py  v2.0
Converte Markdown (output do Datalab/Marker) para DOCX bem formatado.
- Limpa tags HTML residuais
- Tabelas com cabeçalho colorido e linhas alternadas
- Títulos hierárquicos com espaçamento
- Listas, citações, código, hiperligações, imagens
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


# ── Limpeza de HTML residual ──────────────────────────────────────────────────

def _clean_html(text: str) -> str:
    """Remove tags HTML residuais do texto (comuns no output do Datalab)."""
    # Substituir entidades HTML comuns
    text = text.replace("&amp;",  "&")
    text = text.replace("&lt;",   "<")
    text = text.replace("&gt;",   ">")
    text = text.replace("&nbsp;", " ")
    text = text.replace("&quot;", '"')
    text = text.replace("&#39;",  "'")
    text = text.replace("&apos;", "'")
    # Remover tags <br>, <br/>, <hr>
    text = re.sub(r"<br\s*/?>", "\n", text, flags=re.IGNORECASE)
    text = re.sub(r"<hr\s*/?>", "\n---\n", text, flags=re.IGNORECASE)
    # Converter <b>/<strong> → **bold**
    text = re.sub(r"<(?:b|strong)>(.*?)</(?:b|strong)>", r"**\1**", text, flags=re.IGNORECASE | re.DOTALL)
    # Converter <i>/<em> → *italic*
    text = re.sub(r"<(?:i|em)>(.*?)</(?:i|em)>", r"*\1*", text, flags=re.IGNORECASE | re.DOTALL)
    # Converter <code> → `code`
    text = re.sub(r"<code>(.*?)</code>", r"`\1`", text, flags=re.IGNORECASE | re.DOTALL)
    # Remover todas as outras tags HTML
    text = re.sub(r"<[^>]+>", "", text)
    # Limpar espaços múltiplos (mas não newlines)
    text = re.sub(r"[ \t]{2,}", " ", text)
    return text.strip()


# ── Helpers XML ───────────────────────────────────────────────────────────────

def _set_cell_bg(cell, color_hex: str):
    shd = OxmlElement("w:shd")
    shd.set(qn("w:fill"),  color_hex)
    shd.set(qn("w:val"),   "clear")
    shd.set(qn("w:color"), "auto")
    cell._tc.get_or_add_tcPr().append(shd)


def _set_cell_border(cell, color_hex="CCCCCC"):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    tcBorders = OxmlElement("w:tcBorders")
    for side in ("top", "left", "bottom", "right"):
        border = OxmlElement(f"w:{side}")
        border.set(qn("w:val"),   "single")
        border.set(qn("w:sz"),    "4")
        border.set(qn("w:space"), "0")
        border.set(qn("w:color"), color_hex)
        tcBorders.append(border)
    tcPr.append(tcBorders)


def _add_hyperlink(paragraph, text: str, url: str):
    part = paragraph.part
    r_id = part.relate_to(
        url,
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink",
        is_external=True
    )
    hl = OxmlElement("w:hyperlink")
    hl.set(qn("r:id"), r_id)
    r = OxmlElement("w:r")
    rPr = OxmlElement("w:rPr")
    color = OxmlElement("w:color"); color.set(qn("w:val"), "0563C1"); rPr.append(color)
    u = OxmlElement("w:u"); u.set(qn("w:val"), "single"); rPr.append(u)
    r.append(rPr)
    t = OxmlElement("w:t"); t.text = text; r.append(t)
    hl.append(r)
    paragraph._p.append(hl)


def _set_para_spacing(para, before=0, after=4):
    para.paragraph_format.space_before = Pt(before)
    para.paragraph_format.space_after  = Pt(after)


# ── Inline formatting ─────────────────────────────────────────────────────────

def _parse_inline(text: str, paragraph, font_size=11):
    """Aplica formatação inline: bold, italic, code, hiperligações, placeholders."""
    text = _clean_html(text)
    pattern = re.compile(
        r"\[(carimbo|assinatura|logo\s*marca)\]"  # special placeholders
        r"|\[([^\]]+)\]\(([^)]+)\)"               # [text](url)
        r"|(?<!\*)\*\*(.+?)\*\*(?!\*)"            # **bold**
        r"|\*\*\*(.+?)\*\*\*"                     # ***bold italic***
        r"|(?<!\*)\*(.+?)\*(?!\*)"                # *italic*
        r"|`([^`]+)`"                             # `code`
    )
    pos = 0
    for m in pattern.finditer(text):
        if m.start() > pos:
            run = paragraph.add_run(text[pos:m.start()])
            run.font.size = Pt(font_size)
        pos = m.end()
        if m.group(1):                              # placeholder classificado
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
        elif m.group(2) and m.group(3):             # hyperlink
            _add_hyperlink(paragraph, m.group(2), m.group(3))
        elif m.group(4):                            # bold
            run = paragraph.add_run(m.group(4))
            run.bold = True; run.font.size = Pt(font_size)
        elif m.group(5):                            # bold italic
            run = paragraph.add_run(m.group(5))
            run.bold = True; run.italic = True; run.font.size = Pt(font_size)
        elif m.group(6):                            # italic
            run = paragraph.add_run(m.group(6))
            run.italic = True; run.font.size = Pt(font_size)
        elif m.group(7):                            # code
            run = paragraph.add_run(m.group(7))
            run.font.name  = "Courier New"
            run.font.size  = Pt(10)
            run.font.color.rgb = RGBColor(0xC7, 0x25, 0x4E)
    if pos < len(text):
        run = paragraph.add_run(text[pos:])
        run.font.size = Pt(font_size)


# ── Table parser ──────────────────────────────────────────────────────────────

def _parse_table(doc: Document, lines: list, i: int) -> int:
    """Renderiza uma tabela markdown com cabeçalho, bordas e linhas alternadas."""
    header_line = _clean_html(lines[i].strip())
    sep_line    = lines[i + 1].strip() if i + 1 < len(lines) else ""

    if not re.match(r"^[\s|:\-]+$", sep_line):
        return 0

    # Parse alignment from separator
    col_aligns = []
    for cell in sep_line.split("|"):
        c = cell.strip()
        if c.startswith(":") and c.endswith(":"):
            col_aligns.append("center")
        elif c.endswith(":"):
            col_aligns.append("right")
        else:
            col_aligns.append("left")

    header_cells = [_clean_html(c.strip()) for c in header_line.split("|") if c.strip()]
    num_cols = len(header_cells)
    if num_cols == 0:
        return 0

    rows_data = []
    j = i + 2
    while j < len(lines):
        line = lines[j].strip()
        if not line or "|" not in line:
            break
        if line.startswith("#") or line.startswith("```"):
            break
        cells = [_clean_html(c.strip()) for c in line.split("|")]
        if cells and cells[0] == "":
            cells = cells[1:]
        if cells and cells[-1] == "":
            cells = cells[:-1]
        while len(cells) < num_cols:
            cells.append("")
        rows_data.append(cells[:num_cols])
        j += 1

    # Create table
    table = doc.add_table(rows=1 + len(rows_data), cols=num_cols)
    table.style     = "Table Grid"
    table.alignment = WD_TABLE_ALIGNMENT.LEFT

    # Header row — dark blue background, white text
    for ci, cell_text in enumerate(header_cells):
        cell = table.rows[0].cells[ci]
        cell.text = ""
        p   = cell.paragraphs[0]
        run = p.add_run(cell_text)
        run.bold       = True
        run.font.size  = Pt(10)
        run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
        p.paragraph_format.space_before = Pt(2)
        p.paragraph_format.space_after  = Pt(2)
        _set_cell_bg(cell, "1F3864")   # dark navy
        _set_cell_border(cell, "1F3864")

    # Data rows
    for ri, row_cells in enumerate(rows_data):
        bg = "EBF3FB" if ri % 2 == 0 else "FFFFFF"  # light blue / white alternating
        for ci, cell_text in enumerate(row_cells):
            cell = table.rows[ri + 1].cells[ci]
            cell.text = ""
            p = cell.paragraphs[0]
            p.paragraph_format.space_before = Pt(1)
            p.paragraph_format.space_after  = Pt(1)
            _parse_inline(cell_text, p, font_size=10)
            _set_cell_bg(cell, bg)
            _set_cell_border(cell, "BDD7EE")

    # Add space after table
    doc.add_paragraph()
    return j - i


# ── Code block ────────────────────────────────────────────────────────────────

def _parse_code_block(doc: Document, lines: list, i: int) -> int:
    lang = lines[i].strip()[3:].strip()
    j    = i + 1
    code_lines = []
    while j < len(lines):
        if lines[j].strip().startswith("```"):
            j += 1
            break
        code_lines.append(lines[j].rstrip())
        j += 1

    if code_lines:
        p   = doc.add_paragraph()
        pPr = p._p.get_or_add_pPr()
        shd = OxmlElement("w:shd")
        shd.set(qn("w:val"),   "clear")
        shd.set(qn("w:color"), "auto")
        shd.set(qn("w:fill"),  "F4F4F4")
        pPr.append(shd)
        p.paragraph_format.left_indent   = Cm(0.5)
        p.paragraph_format.space_before  = Pt(4)
        p.paragraph_format.space_after   = Pt(4)
        run = p.add_run("\n".join(code_lines))
        run.font.name  = "Courier New"
        run.font.size  = Pt(9)
        run.font.color.rgb = RGBColor(0x24, 0x29, 0x2E)
        if lang:
            p2  = doc.add_paragraph()
            r2  = p2.add_run(f"  {lang}")
            r2.font.size  = Pt(8)
            r2.font.color.rgb = RGBColor(0x88, 0x88, 0x88)
            r2.italic = True
            p2.paragraph_format.space_before = Pt(0)
            p2.paragraph_format.space_after  = Pt(2)
    return j - i


# ── List parsers ──────────────────────────────────────────────────────────────

def _parse_unordered_list(doc: Document, lines: list, i: int) -> int:
    start = i
    while i < len(lines):
        s = lines[i].strip()
        if not s:
            i += 1; continue
        if not re.match(r"^[\-\*\+]\s", s):
            break
        text = re.sub(r"^[\-\*\+]\s+", "", s)
        p    = doc.add_paragraph(style="List Bullet")
        _parse_inline(text, p)
        _set_para_spacing(p, 0, 2)
        i += 1
    return i - start


def _parse_ordered_list(doc: Document, lines: list, i: int) -> int:
    start = i
    while i < len(lines):
        s = lines[i].strip()
        if not s:
            i += 1; continue
        if not re.match(r"^\d+[\.\)]\s", s):
            break
        text = re.sub(r"^\d+[\.\)]\s+", "", s)
        p    = doc.add_paragraph(style="List Number")
        _parse_inline(text, p)
        _set_para_spacing(p, 0, 2)
        i += 1
    return i - start


# ── Blockquote ────────────────────────────────────────────────────────────────

def _parse_blockquote(doc: Document, lines: list, i: int) -> int:
    start = i
    while i < len(lines):
        s = lines[i].strip()
        if not s.startswith(">"):
            break
        text = re.sub(r"^>\s?", "", s)
        p    = doc.add_paragraph()
        p.paragraph_format.left_indent = Cm(1.0)
        pPr  = p._p.get_or_add_pPr()
        shd  = OxmlElement("w:shd")
        shd.set(qn("w:val"),   "clear")
        shd.set(qn("w:color"), "auto")
        shd.set(qn("w:fill"),  "F0F4FF")
        pPr.append(shd)
        bdr  = OxmlElement("w:pBdr")
        left = OxmlElement("w:left")
        left.set(qn("w:val"),   "single")
        left.set(qn("w:sz"),    "16")
        left.set(qn("w:space"), "8")
        left.set(qn("w:color"), "4472C4")
        bdr.append(left); pPr.append(bdr)
        _parse_inline(text, p)
        _set_para_spacing(p, 2, 2)
        i += 1
    return i - start


# ── Main converter ────────────────────────────────────────────────────────────

def markdown_to_docx(markdown_text: str, docx_path: str, images_dict: dict = None):
    """
    Converte Markdown para DOCX formatado.
    markdown_text : texto markdown (output do Datalab)
    docx_path     : caminho de saída .docx
    images_dict   : {nome: bytes} — imagens do Datalab (opcional)
    """
    # Limpar HTML residual do texto completo primeiro
    markdown_text = re.sub(r"<!--.*?-->", "", markdown_text, flags=re.DOTALL)  # comentários HTML

    doc     = Document()
    section = doc.sections[0]
    section.top_margin    = Cm(2.54)
    section.bottom_margin = Cm(2.54)
    section.left_margin   = Cm(2.54)
    section.right_margin  = Cm(2.54)

    # Estilos de título personalizados
    from docx.shared import RGBColor as RGB
    heading_colors = {
        1: RGB(0x1F, 0x38, 0x64),  # navy escuro
        2: RGB(0x2E, 0x74, 0xB5),  # azul médio
        3: RGB(0x2E, 0x74, 0xB5),  # azul médio
        4: RGB(0x40, 0x40, 0x40),  # cinza escuro
    }

    lines         = markdown_text.split("\n")
    i             = 0
    empty_streak  = 0

    while i < len(lines):
        raw     = lines[i]
        stripped = raw.strip()

        # ── Linha vazia ───────────────────────────────────────────────────────
        if not stripped:
            empty_streak += 1
            if empty_streak == 1:
                doc.add_paragraph()
            i += 1
            continue
        empty_streak = 0

        # ── Bloco de código ───────────────────────────────────────────────────
        if stripped.startswith("```"):
            consumed = _parse_code_block(doc, lines, i)
            i += consumed if consumed > 0 else 1
            continue

        # ── Separador horizontal ──────────────────────────────────────────────
        if re.match(r"^[-\*_]{3,}$", stripped):
            p   = doc.add_paragraph()
            pPr = p._p.get_or_add_pPr()
            pBdr = OxmlElement("w:pBdr")
            bot  = OxmlElement("w:bottom")
            bot.set(qn("w:val"),   "single")
            bot.set(qn("w:sz"),    "6")
            bot.set(qn("w:space"), "1")
            bot.set(qn("w:color"), "AAAAAA")
            pBdr.append(bot); pPr.append(pBdr)
            i += 1
            continue

        # ── Tabela ────────────────────────────────────────────────────────────
        if "|" in stripped and i + 1 < len(lines):
            next_stripped = lines[i + 1].strip()
            if re.match(r"^[\s|:\-]+$", next_stripped) and "|" in next_stripped:
                consumed = _parse_table(doc, lines, i)
                if consumed > 0:
                    i += consumed
                    continue

        # ── Título ────────────────────────────────────────────────────────────
        hm = re.match(r"^(#{1,6})\s+(.+)$", stripped)
        if hm:
            level = min(len(hm.group(1)), 4)
            text  = _clean_html(hm.group(2))
            p     = doc.add_heading(text, level=level)
            # Aplicar cor ao título
            for run in p.runs:
                run.font.color.rgb = heading_colors.get(level, RGB(0x1F, 0x38, 0x64))
            spaces = {1: (12, 6), 2: (10, 4), 3: (8, 3), 4: (6, 2)}
            bef, aft = spaces.get(level, (6, 2))
            _set_para_spacing(p, bef, aft)
            i += 1
            continue

        # ── Lista não ordenada ────────────────────────────────────────────────
        if re.match(r"^[\-\*\+]\s", stripped):
            consumed = _parse_unordered_list(doc, lines, i)
            if consumed > 0:
                i += consumed
                continue

        # ── Lista ordenada ────────────────────────────────────────────────────
        if re.match(r"^\d+[\.\)]\s", stripped):
            consumed = _parse_ordered_list(doc, lines, i)
            if consumed > 0:
                i += consumed
                continue

        # ── Citação ───────────────────────────────────────────────────────────
        if stripped.startswith(">"):
            consumed = _parse_blockquote(doc, lines, i)
            if consumed > 0:
                i += consumed
                continue

        # ── Imagem ────────────────────────────────────────────────────────────
        img_m = re.match(r"^!\[([^\]]*)\]\(([^)]+)\)$", stripped)
        if img_m:
            alt, img_path = img_m.group(1), img_m.group(2)
            inserted = False
            if images_dict:
                key = next((k for k in images_dict
                            if k.endswith(img_path) or img_path in k), None)
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
                p   = doc.add_paragraph()
                run = p.add_run(f"[{alt or 'IMAGEM'}]")
                run.italic = True
                run.font.color.rgb = RGBColor(0x88, 0x88, 0x88)
            i += 1
            continue

        # ── Placeholder de imagem classificada ────────────────────────────────
        ph_m = re.match(r"^\[(carimbo|assinatura|logo\s*marca)\]$", stripped)
        if ph_m:
            ptype = ph_m.group(1)
            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            bg_colors = {"carimbo": "FFF5F5", "assinatura": "F0F8FF", "logo marca": "F5F5F5"}
            pPr = p._p.get_or_add_pPr()
            shd = OxmlElement("w:shd")
            shd.set(qn("w:val"), "clear"); shd.set(qn("w:color"), "auto")
            shd.set(qn("w:fill"), bg_colors.get(ptype, "F9F9F9"))
            pPr.append(shd)
            bdr = OxmlElement("w:pBdr")
            left = OxmlElement("w:left")
            left.set(qn("w:val"), "single"); left.set(qn("w:sz"), "12")
            left.set(qn("w:space"), "6")
            border_colors = {"carimbo": "C0392B", "assinatura": "2E86C1", "logo marca": "6C6C6C"}
            left.set(qn("w:color"), border_colors.get(ptype, "999999"))
            bdr.append(left); pPr.append(bdr)
            _parse_inline(stripped, p, font_size=10)
            _set_para_spacing(p, 3, 3)
            i += 1
            continue

        # ── Parágrafo normal ──────────────────────────────────────────────────
        p = doc.add_paragraph()
        _parse_inline(stripped, p, font_size=11)
        _set_para_spacing(p, 0, 3)
        i += 1

    doc.save(str(docx_path))