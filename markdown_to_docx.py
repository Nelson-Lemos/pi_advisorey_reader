# -*- coding: utf-8 -*-
"""
markdown_to_docx.py  v3.0  —  Fidelidade máxima ao documento original
======================================================================
Converte o output Markdown do Datalab para DOCX preservando:
- Tabelas financeiras com todas as colunas e linhas
- Negrito real (não asteriscos visíveis)
- Títulos hierárquicos
- Listas, citações, separadores
- Remove ruído OCR (carimbos, descrições de imagens)
- Sem asteriscos ou símbolos markdown visíveis no output final
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


# ── Filtros de ruído OCR ──────────────────────────────────────────────────────

STAMP_PATTERNS = [
    re.compile(r'^\s*\[(?:Circular|circular|Faint|faint|A circular|Three circular|Two circular|Logo|logo|A blank|This image|An image|Image of|Photo|Picture)[^\]]*\]\s*$', re.IGNORECASE),
    re.compile(r'^\s*(?:Circular|circular|Faint|faint) stamp of\b', re.IGNORECASE),
    re.compile(r'^\s*A (?:circular|faint|blue|red|black|green) (?:ink |blue )?stamp\b', re.IGNORECASE),
    re.compile(r'^\s*(?:Three|Two|A) circular (?:blue|red|black) ink stamps?\b', re.IGNORECASE),
    re.compile(r'^\s*(?:Logo of|This image shows|An image of|Photo of|Picture of)\b', re.IGNORECASE),
    re.compile(r'^\s*A blank,?\s+aged\b', re.IGNORECASE),
    re.compile(r'^\s*\[?(?:stamp|seal|logo|emblem|signature)\b[^\]]*\]?\s*$', re.IGNORECASE),
]

def _is_noise_line(line):
    s = line.strip()
    if not s:
        return False
    for pat in STAMP_PATTERNS:
        if pat.match(s):
            return True
    return False


def _clean_markdown(text):
    """Remove noise, fix HTML entities, normalise whitespace."""
    # HTML comments
    text = re.sub(r'<!--.*?-->', '', text, flags=re.DOTALL)
    # HTML entities
    text = text.replace('&amp;', '&').replace('&lt;', '<').replace('&gt;', '>') \
               .replace('&nbsp;', ' ').replace('&quot;', '"').replace('&#39;', "'") \
               .replace('&apos;', "'").replace('&ldquo;', '"').replace('&rdquo;', '"') \
               .replace('&lsquo;', "'").replace('&rsquo;', "'").replace('&mdash;', '—') \
               .replace('&ndash;', '–').replace('&hellip;', '…')
    # <br> → newline
    text = re.sub(r'<br\s*/?>', '\n', text, flags=re.IGNORECASE)
    # <sup>/<sub> — keep text, remove tags
    text = re.sub(r'<su[pb][^>]*>(.*?)</su[pb]>', r'\1', text, flags=re.IGNORECASE | re.DOTALL)
    # Remove remaining HTML tags
    text = re.sub(r'<[^>]+>', '', text)
    # Filter noise lines (stamps, image descriptions)
    lines = text.split('\n')
    lines = [l for l in lines if not _is_noise_line(l)]
    text = '\n'.join(lines)
    # Collapse 3+ blank lines → 2
    text = re.sub(r'\n{4,}', '\n\n\n', text)
    return text.strip()


# ── XML helpers ───────────────────────────────────────────────────────────────

def _cell_bg(cell, hex_color):
    shd = OxmlElement('w:shd')
    shd.set(qn('w:fill'), hex_color)
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    cell._tc.get_or_add_tcPr().append(shd)


def _cell_borders(cell, color='AAAAAA'):
    tcPr = cell._tc.get_or_add_tcPr()
    tcBorders = OxmlElement('w:tcBorders')
    for side in ('top', 'left', 'bottom', 'right'):
        b = OxmlElement(f'w:{side}')
        b.set(qn('w:val'), 'single')
        b.set(qn('w:sz'), '4')
        b.set(qn('w:space'), '0')
        b.set(qn('w:color'), color)
        tcBorders.append(b)
    tcPr.append(tcBorders)


def _add_hyperlink(paragraph, text, url):
    try:
        r_id = paragraph.part.relate_to(
            url,
            'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink',
            is_external=True)
        hl = OxmlElement('w:hyperlink')
        hl.set(qn('r:id'), r_id)
        r = OxmlElement('w:r')
        rPr = OxmlElement('w:rPr')
        c = OxmlElement('w:color'); c.set(qn('w:val'), '0563C1'); rPr.append(c)
        u = OxmlElement('w:u'); u.set(qn('w:val'), 'single'); rPr.append(u)
        r.append(rPr)
        t = OxmlElement('w:t'); t.text = text; r.append(t)
        hl.append(r)
        paragraph._p.append(hl)
    except Exception:
        paragraph.add_run(text)


def _para_spacing(para, before=0, after=3):
    para.paragraph_format.space_before = Pt(before)
    para.paragraph_format.space_after = Pt(after)


# ── Inline formatting ─────────────────────────────────────────────────────────

def _add_inline(text, paragraph, size=11):
    """Parse **bold**, *italic*, `code`, [link](url) and plain text."""
    # Clean text first
    text = re.sub(r'<[^>]+>', '', text)
    pat = re.compile(
        r'\[([^\]]+)\]\(([^)]+)\)'        # [text](url)
        r'|\*\*\*(.+?)\*\*\*'             # ***bold italic***
        r'|\*\*(.+?)\*\*'                 # **bold**
        r'|\*(.+?)\*'                     # *italic*
        r'|`([^`]+)`'                     # `code`
    )
    pos = 0
    for m in pat.finditer(text):
        if m.start() > pos:
            run = paragraph.add_run(text[pos:m.start()])
            run.font.size = Pt(size)
        pos = m.end()
        if m.group(1) and m.group(2):
            _add_hyperlink(paragraph, m.group(1), m.group(2))
        elif m.group(3):
            r = paragraph.add_run(m.group(3))
            r.bold = True; r.italic = True; r.font.size = Pt(size)
        elif m.group(4):
            r = paragraph.add_run(m.group(4))
            r.bold = True; r.font.size = Pt(size)
        elif m.group(5):
            r = paragraph.add_run(m.group(5))
            r.italic = True; r.font.size = Pt(size)
        elif m.group(6):
            r = paragraph.add_run(m.group(6))
            r.font.name = 'Courier New'; r.font.size = Pt(9)
            r.font.color.rgb = RGBColor(0xC7, 0x25, 0x4E)
    if pos < len(text):
        run = paragraph.add_run(text[pos:])
        run.font.size = Pt(size)


# ── Table renderer ────────────────────────────────────────────────────────────

def _parse_table(doc, lines, start):
    """Render a markdown table into a Word table. Returns lines consumed."""
    header_raw = lines[start].strip()
    sep = lines[start + 1].strip() if start + 1 < len(lines) else ''
    if not re.match(r'^[\s|:\-]+$', sep):
        return 0

    def split_cells(row):
        cells = row.split('|')
        # Remove leading/trailing empty from | at start/end
        if cells and cells[0].strip() == '':
            cells = cells[1:]
        if cells and cells[-1].strip() == '':
            cells = cells[:-1]
        return [c.strip() for c in cells]

    headers = split_cells(header_raw)
    ncols = len(headers)
    if ncols == 0:
        return 0

    # Parse alignment from separator
    align = []
    for c in split_cells(sep):
        if c.startswith(':') and c.endswith(':'):
            align.append(WD_ALIGN_PARAGRAPH.CENTER)
        elif c.endswith(':'):
            align.append(WD_ALIGN_PARAGRAPH.RIGHT)
        else:
            align.append(WD_ALIGN_PARAGRAPH.LEFT)
    while len(align) < ncols:
        align.append(WD_ALIGN_PARAGRAPH.LEFT)

    rows_data = []
    j = start + 2
    while j < len(lines):
        row = lines[j].strip()
        if not row or '|' not in row:
            break
        if row.startswith('#') or row.startswith('```'):
            break
        cells = split_cells(row)
        while len(cells) < ncols:
            cells.append('')
        rows_data.append(cells[:ncols])
        j += 1

    # Build table
    tbl = doc.add_table(rows=1 + len(rows_data), cols=ncols)
    tbl.style = 'Table Grid'
    tbl.alignment = WD_TABLE_ALIGNMENT.LEFT

    # Header row
    for ci, htext in enumerate(headers):
        cell = tbl.rows[0].cells[ci]
        cell.text = ''
        p = cell.paragraphs[0]
        p.alignment = align[ci] if ci < len(align) else WD_ALIGN_PARAGRAPH.LEFT
        _add_inline(htext, p, size=9)
        for run in p.runs:
            run.bold = True
            run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
        p.paragraph_format.space_before = Pt(2)
        p.paragraph_format.space_after = Pt(2)
        _cell_bg(cell, '1F3864')
        _cell_borders(cell, '1F3864')

    # Data rows
    for ri, row_cells in enumerate(rows_data):
        bg = 'EBF3FB' if ri % 2 == 0 else 'FFFFFF'
        for ci, ctext in enumerate(row_cells):
            cell = tbl.rows[ri + 1].cells[ci]
            cell.text = ''
            p = cell.paragraphs[0]
            p.alignment = align[ci] if ci < len(align) else WD_ALIGN_PARAGRAPH.LEFT
            _add_inline(ctext, p, size=9)
            p.paragraph_format.space_before = Pt(1)
            p.paragraph_format.space_after = Pt(1)
            _cell_bg(cell, bg)
            _cell_borders(cell, 'BDD7EE')

    doc.add_paragraph()  # space after table
    return j - start


# ── Code block ────────────────────────────────────────────────────────────────

def _parse_code(doc, lines, start):
    lang = lines[start].strip()[3:].strip()
    j = start + 1
    code = []
    while j < len(lines):
        if lines[j].strip().startswith('```'):
            j += 1
            break
        code.append(lines[j].rstrip())
        j += 1
    if code:
        p = doc.add_paragraph()
        p.paragraph_format.left_indent = Cm(0.5)
        p.paragraph_format.space_before = Pt(4)
        p.paragraph_format.space_after = Pt(4)
        pPr = p._p.get_or_add_pPr()
        shd = OxmlElement('w:shd')
        shd.set(qn('w:val'), 'clear'); shd.set(qn('w:color'), 'auto'); shd.set(qn('w:fill'), 'F4F4F4')
        pPr.append(shd)
        r = p.add_run('\n'.join(code))
        r.font.name = 'Courier New'; r.font.size = Pt(9)
        r.font.color.rgb = RGBColor(0x24, 0x29, 0x2E)
    return j - start


# ── Lists ─────────────────────────────────────────────────────────────────────

def _parse_ulist(doc, lines, start):
    i = start
    while i < len(lines):
        s = lines[i].strip()
        if not s:
            i += 1; continue
        if not re.match(r'^[-*+•]\s', s):
            break
        text = re.sub(r'^[-*+•]\s+', '', s)
        p = doc.add_paragraph(style='List Bullet')
        _add_inline(text, p)
        _para_spacing(p, 0, 2)
        i += 1
    return i - start


def _parse_olist(doc, lines, start):
    i = start
    while i < len(lines):
        s = lines[i].strip()
        if not s:
            i += 1; continue
        if not re.match(r'^\d+[.)]\s', s):
            break
        text = re.sub(r'^\d+[.)]\s+', '', s)
        p = doc.add_paragraph(style='List Number')
        _add_inline(text, p)
        _para_spacing(p, 0, 2)
        i += 1
    return i - start


# ── Blockquote ────────────────────────────────────────────────────────────────

def _parse_quote(doc, lines, start):
    i = start
    while i < len(lines):
        s = lines[i].strip()
        if not s.startswith('>'):
            break
        text = re.sub(r'^>\s?', '', s)
        p = doc.add_paragraph()
        p.paragraph_format.left_indent = Cm(1.0)
        pPr = p._p.get_or_add_pPr()
        shd = OxmlElement('w:shd')
        shd.set(qn('w:val'), 'clear'); shd.set(qn('w:color'), 'auto'); shd.set(qn('w:fill'), 'F0F4FF')
        pPr.append(shd)
        bdr = OxmlElement('w:pBdr')
        left = OxmlElement('w:left')
        left.set(qn('w:val'), 'single'); left.set(qn('w:sz'), '16')
        left.set(qn('w:space'), '8'); left.set(qn('w:color'), '4472C4')
        bdr.append(left); pPr.append(bdr)
        _add_inline(text, p)
        _para_spacing(p, 2, 2)
        i += 1
    return i - start


# ── Main converter ────────────────────────────────────────────────────────────

HEADING_COLORS = {
    1: RGBColor(0x1F, 0x38, 0x64),
    2: RGBColor(0x2E, 0x74, 0xB5),
    3: RGBColor(0x2E, 0x74, 0xB5),
    4: RGBColor(0x40, 0x40, 0x40),
}
HEADING_SPACING = {1: (14, 6), 2: (10, 4), 3: (8, 3), 4: (6, 2)}


def markdown_to_docx(markdown_text: str, docx_path: str, images_dict: dict = None):
    """
    Converte Markdown (output Datalab/Marker) para DOCX fiel ao original.
    - Remove carimbos e ruído OCR automaticamente
    - Converte **negrito** e *itálico* em formatação Word real (sem asteriscos)
    - Tabelas financeiras com cabeçalho colorido
    - Títulos, listas, citações, separadores
    """
    markdown_text = _clean_markdown(markdown_text)

    doc = Document()
    sec = doc.sections[0]
    sec.top_margin = sec.bottom_margin = Cm(2.54)
    sec.left_margin = sec.right_margin = Cm(2.54)

    lines = markdown_text.split('\n')
    i = 0
    blank_streak = 0

    while i < len(lines):
        raw = lines[i]
        s = raw.strip()

        # Blank line
        if not s:
            blank_streak += 1
            if blank_streak == 1:
                doc.add_paragraph()
            i += 1
            continue
        blank_streak = 0

        # Code block
        if s.startswith('```'):
            n = _parse_code(doc, lines, i)
            i += n if n > 0 else 1
            continue

        # Horizontal rule
        if re.match(r'^[-*_]{3,}$', s):
            p = doc.add_paragraph()
            pPr = p._p.get_or_add_pPr()
            pBdr = OxmlElement('w:pBdr')
            bot = OxmlElement('w:bottom')
            bot.set(qn('w:val'), 'single'); bot.set(qn('w:sz'), '6')
            bot.set(qn('w:space'), '1'); bot.set(qn('w:color'), 'AAAAAA')
            pBdr.append(bot); pPr.append(pBdr)
            i += 1
            continue

        # Table
        if '|' in s and i + 1 < len(lines):
            nxt = lines[i + 1].strip()
            if re.match(r'^[\s|:\-]+$', nxt) and '|' in nxt:
                n = _parse_table(doc, lines, i)
                if n > 0:
                    i += n
                    continue

        # Heading
        hm = re.match(r'^(#{1,6})\s+(.+)$', s)
        if hm:
            level = min(len(hm.group(1)), 4)
            text = re.sub(r'<[^>]+>', '', hm.group(2))
            # Strip trailing markdown formatting from heading text
            text = re.sub(r'\*+$', '', text).strip()
            p = doc.add_heading(text, level=level)
            for run in p.runs:
                run.font.color.rgb = HEADING_COLORS.get(level, HEADING_COLORS[4])
                run.bold = True
            bef, aft = HEADING_SPACING.get(level, (6, 2))
            _para_spacing(p, bef, aft)
            i += 1
            continue

        # Unordered list
        if re.match(r'^[-*+•]\s', s):
            n = _parse_ulist(doc, lines, i)
            if n > 0:
                i += n
                continue

        # Ordered list
        if re.match(r'^\d+[.)]\s', s):
            n = _parse_olist(doc, lines, i)
            if n > 0:
                i += n
                continue

        # Blockquote
        if s.startswith('>'):
            n = _parse_quote(doc, lines, i)
            if n > 0:
                i += n
                continue

        # Image (inline or block)
        img_m = re.match(r'^!\[([^\]]*)\]\(([^)]+)\)$', s)
        if img_m:
            alt, img_path = img_m.group(1), img_m.group(2)
            inserted = False
            if images_dict:
                key = next((k for k in images_dict if k.endswith(img_path) or img_path in k), None)
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
                ap = Path(img_path)
                if ap.exists():
                    try:
                        p = doc.add_paragraph()
                        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                        p.add_run().add_picture(str(ap), width=Inches(5.0))
                        inserted = True
                    except Exception:
                        pass
            if not inserted:
                # Show image placeholder — always keep, skip only stamp/noise descriptions
                label = alt if alt and not _is_noise_line(alt) else "IMAGEM"
                p = doc.add_paragraph()
                r = p.add_run(f'[ {label} ]')
                r.italic = True
                r.font.color.rgb = RGBColor(0x44, 0x72, 0xC4)
            i += 1
            continue

        # Normal paragraph — convert inline markdown to Word formatting
        p = doc.add_paragraph()
        _add_inline(s, p, size=11)
        _para_spacing(p, 0, 3)
        i += 1

    doc.save(str(docx_path))