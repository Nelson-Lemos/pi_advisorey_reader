# -*- coding: utf-8 -*-
"""
PDF Translator Pro v5.0
- Fiel ao documento original (tabelas, paragrafos, fontes, numeracao)
- Traduz apenas o texto (se pedido)
- Imagens substituidas por [LOGOMARCA] ou [CARIMBO]
"""

from fastapi import FastAPI, UploadFile, File, BackgroundTasks, HTTPException
from fastapi.staticfiles import StaticFiles
from fastapi.responses import FileResponse, HTMLResponse
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel
from typing import List
import uuid, json, zipfile, shutil, asyncio
from pathlib import Path
from datetime import datetime
import logging

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

#  TESSERACT PATH (Windows) 
import os
_TESS_PATHS = [
    r"C:\Program Files\Tesseract-OCR\tesseract.exe",
    r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe",
    r"C:\Users\HP ZBOOK 15 G3\AppData\Local\Tesseract-OCR\tesseract.exe",
]
def _setup_tesseract():
    try:
        import pytesseract
        for p in _TESS_PATHS:
            if os.path.exists(p):
                pytesseract.pytesseract.tesseract_cmd = p
                logger.info("Tesseract found: %s", p)
                return True
        # Try PATH
        import shutil
        if shutil.which("tesseract"):
            logger.info("Tesseract found in PATH")
            return True
        logger.warning("Tesseract not found in common paths")
        return False
    except ImportError:
        return False

_TESSERACT_OK = _setup_tesseract()

#  PADDLEOCR SETUP 
_PADDLE_OCR = None
def _get_paddle():
    global _PADDLE_OCR
    if _PADDLE_OCR is not None:
        return _PADDLE_OCR
    try:
        from paddleocr import PaddleOCR
        _PADDLE_OCR = PaddleOCR(use_angle_cls=True, lang="en",
                                 show_log=False, use_gpu=False)
        logger.info("PaddleOCR initialised")
        return _PADDLE_OCR
    except ImportError:
        logger.info("PaddleOCR not installed (optional). Using Tesseract.")
        return None
    except Exception as e:
        logger.warning("PaddleOCR init failed: %s", e)
        return None

app = FastAPI(title="PDF Translator Pro", version="5.0.0")
app.add_middleware(CORSMiddleware, allow_origins=["*"], allow_credentials=True,
                   allow_methods=["*"], allow_headers=["*"])

BASE_DIR   = Path(__file__).resolve().parent
STATIC_DIR = BASE_DIR / "static"
UPLOAD_DIR = BASE_DIR / "uploads"
OUTPUT_DIR = BASE_DIR / "outputs"
JOBS_DIR   = BASE_DIR / "jobs"

for d in [STATIC_DIR, UPLOAD_DIR, OUTPUT_DIR, JOBS_DIR]:
    d.mkdir(parents=True, exist_ok=True)

logger.info("BASE_DIR: %s", BASE_DIR)
logger.info("index.html exists: %s", (STATIC_DIR / "index.html").exists())

app.mount("/static", StaticFiles(directory=str(STATIC_DIR)), name="static")


class ProcessRequest(BaseModel):
    job_id: str
    target_language: str = "pt"
    image_mode: str      = "placeholder"
    output_pdf: bool     = False
    remove_images: bool  = False


def save_job(job):
    (JOBS_DIR / f"{job['job_id']}.json").write_text(
        json.dumps(job, ensure_ascii=False, indent=2), encoding="utf-8")

def load_job(job_id):
    p = JOBS_DIR / f"{job_id}.json"
    if not p.exists():
        raise HTTPException(404, "Job not found")
    return json.loads(p.read_text(encoding="utf-8"))

def list_jobs():
    out = []
    for p in sorted(JOBS_DIR.glob("*.json"), key=lambda x: x.stat().st_mtime, reverse=True):
        try:
            out.append(json.loads(p.read_text(encoding="utf-8")))
        except Exception:
            pass
    return out


#  TRANSLATION 

def translate_text(text, lang, translator):
    """Translate a single string. Returns original if translation fails."""
    if not text or not text.strip():
        return text
    try:
        MAX = 4800
        if len(text) <= MAX:
            r = translator.translate(text)
            return r.strip() if r else text
        # Split long text
        parts, buf = [], ""
        for sentence in text.replace(". ", ".|||").split("|||"):
            if len(buf) + len(sentence) < MAX:
                buf += sentence
            else:
                if buf:
                    r = translator.translate(buf)
                    parts.append(r.strip() if r else buf)
                buf = sentence
        if buf:
            r = translator.translate(buf)
            parts.append(r.strip() if r else buf)
        return " ".join(parts)
    except Exception as e:
        import time
        msg = str(e)
        if "429" in msg or "TooMany" in msg:
            time.sleep(4)
            try:
                r = translator.translate(text[:MAX])
                return r.strip() if r else text
            except Exception:
                pass
        logger.warning("translate error: %s", e)
        return text


#  IMAGE CLASSIFICATION 

def classify_image(bbox, page_w, page_h):
    x0, y0, x1, y1 = bbox
    w = x1 - x0
    h = y1 - y0
    area_ratio = (w * h) / max(page_w * page_h, 1)
    aspect = w / max(h, 1)
    in_header = y0 < page_h * 0.20
    in_footer = y1 > page_h * 0.80
    is_small  = area_ratio < 0.10
    is_square = 0.5 < aspect < 2.0

    if is_small and (in_header or in_footer):
        return "[LOGOMARCA]"
    if is_small and is_square:
        return "[CARIMBO]"
    return "[FIGURA]"


#  CORE: PDF -> DOCX with faithful structure 

#  OCR PIPELINE 

def is_scanned_pdf(pdf_path):
    """
    Returns True if the PDF is image-based (scanned).
    Checks if pages have very little extractable text.
    """
    try:
        import fitz
        doc = fitz.open(str(pdf_path))
        text_pages = 0
        for page in doc:
            if len(page.get_text().strip()) > 30:
                text_pages += 1
        total = len(doc)
        doc.close()
        ratio = text_pages / max(total, 1)
        logger.info("PDF text ratio: %.0f%% (%d/%d pages have text)",
                    ratio * 100, text_pages, total)
        return ratio < 0.3   # less than 30% pages with text = scanned
    except Exception as e:
        logger.warning("Could not check PDF type: %s", e)
        return False


def ocr_pdf_to_searchable(pdf_path, out_pdf_path, lang_hint="por+eng"):
    """
    Run OCR on a scanned PDF and produce a searchable PDF.
    Uses pytesseract + PyMuPDF to overlay text on the original images.
    Returns path to the searchable PDF.
    """
    import fitz
    try:
        import pytesseract
        from PIL import Image
        import io
    except ImportError:
        raise RuntimeError("pytesseract and Pillow are required for OCR. "
                           "Run: pip install pytesseract Pillow")

    logger.info("Starting OCR on: %s", pdf_path.name)
    doc_in  = fitz.open(str(pdf_path))
    doc_out = fitz.open()   # new empty PDF

    for pn, page in enumerate(doc_in):
        logger.info("  OCR page %d/%d...", pn + 1, len(doc_in))

        # Render page to image at 300 DPI for good OCR quality
        mat = fitz.Matrix(300 / 72, 300 / 72)
        pix = page.get_pixmap(matrix=mat, colorspace=fitz.csRGB)
        img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)

        # Run Tesseract OCR - get both text and layout data
        try:
            ocr_data = pytesseract.image_to_data(
                img, lang=lang_hint,
                output_type=pytesseract.Output.DICT,
                config="--psm 6"   # assume uniform block of text
            )
        except Exception as e:
            logger.warning("  Tesseract error on page %d: %s", pn + 1, e)
            # Add blank page and continue
            new_page = doc_out.new_page(width=page.rect.width, height=page.rect.height)
            new_page.insert_image(new_page.rect, pixmap=pix)
            continue

        # Create new page with same dimensions
        pw = page.rect.width
        ph = page.rect.height
        new_page = doc_out.new_page(width=pw, height=ph)

        # Insert the original page image as background
        new_page.insert_image(new_page.rect, pixmap=pix)

        # Overlay invisible OCR text so the PDF is searchable
        scale_x = pw / pix.width
        scale_y = ph / pix.height
        n = len(ocr_data["text"])
        for i in range(n):
            word = (ocr_data["text"][i] or "").strip()
            conf = int(ocr_data["conf"][i] or 0)
            if not word or conf < 30:   # skip low-confidence words
                continue
            x = ocr_data["left"][i]   * scale_x
            y = ocr_data["top"][i]    * scale_y
            w = ocr_data["width"][i]  * scale_x
            h = ocr_data["height"][i] * scale_y
            font_size = max(h * 0.85, 6)
            # Insert as invisible text (render mode 3 = invisible)
            try:
                new_page.insert_text(
                    fitz.Point(x, y + h),
                    word + " ",
                    fontsize=font_size,
                    render_mode=3,   # invisible - text is there for search/copy
                )
            except Exception:
                pass   # skip problematic characters

    doc_in.close()
    doc_out.save(str(out_pdf_path), deflate=True, garbage=4)
    doc_out.close()
    logger.info("OCR complete -> %s (%.1f KB)",
                out_pdf_path.name, out_pdf_path.stat().st_size / 1024)
    return out_pdf_path


def ocr_pdf_to_docx_direct(pdf_path, docx_path, lang="pt",
                            do_translate=True, image_mode="placeholder",
                            lang_hint="por+eng"):
    """
    Full OCR pipeline for scanned PDFs.
    Tries PaddleOCR first (better quality for complex layouts).
    Falls back to Tesseract if PaddleOCR not available.
    """
    import fitz
    import numpy as np
    from PIL import Image
    from docx import Document
    from docx.shared import Pt, Cm
    from docx.oxml.ns import qn
    from docx.oxml import OxmlElement
    from collections import defaultdict
    import io

    # Setup translator
    translator = None
    if do_translate and lang not in ("", "none"):
        try:
            from deep_translator import GoogleTranslator
            translator = GoogleTranslator(source="auto", target=lang)
        except ImportError:
            pass

    def tr(text):
        if translator and text and text.strip():
            return translate_text(text, lang, translator)
        return text

    doc_pdf   = fitz.open(str(pdf_path))
    doc_out   = Document()
    for sec in doc_out.sections:
        sec.top_margin    = Cm(2.54)
        sec.bottom_margin = Cm(2.54)
        sec.left_margin   = Cm(3.17)
        sec.right_margin  = Cm(3.17)

    total_pages = len(doc_pdf)
    paddle      = _get_paddle()
    engine      = "PaddleOCR" if paddle else "Tesseract"
    logger.info("OCR engine: %s | Pages: %d", engine, total_pages)

    for pn, page in enumerate(doc_pdf):
        logger.info("  OCR page %d/%d...", pn + 1, total_pages)
        if pn > 0:
            doc_out.add_page_break()

        # Render page at 300 DPI
        mat = fitz.Matrix(300 / 72, 300 / 72)
        pix = page.get_pixmap(matrix=mat, colorspace=fitz.csRGB)
        img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
        ph_pts = page.rect.height
        scale_y = ph_pts / pix.height

        lines_out = []   # list of (y_pos, font_pt, text)

        #  PaddleOCR path 
        if paddle:
            try:
                img_np  = np.array(img)
                results = paddle.ocr(img_np, cls=True)
                if results and results[0]:
                    # Sort by vertical position
                    sorted_res = sorted(results[0],
                                        key=lambda r: r[0][0][1])  # top-left Y
                    for item in sorted_res:
                        box, (text, conf) = item[0], item[1]
                        if not text.strip() or conf < 0.3:
                            continue
                        # Estimate font size from box height
                        h_px  = abs(box[2][1] - box[0][1])
                        h_pts = h_px * scale_y
                        font_pt = min(max(round(h_pts * 0.72), 8), 36)
                        y_pos   = box[0][1] * scale_y
                        lines_out.append((y_pos, font_pt, text.strip()))
            except Exception as e:
                logger.warning("  PaddleOCR page %d error: %s, trying Tesseract", pn+1, e)
                lines_out = []   # reset, fall through to Tesseract

        #  Tesseract fallback 
        if not lines_out:
            try:
                import pytesseract
                data = pytesseract.image_to_data(
                    img, lang=lang_hint,
                    output_type=pytesseract.Output.DICT,
                    config="--psm 6 --oem 3"
                )
                n = len(data["text"])
                # Group into lines by line_num + block_num
                groups = defaultdict(list)
                for i in range(n):
                    word = (data["text"][i] or "").strip()
                    conf = int(data["conf"][i] or 0)
                    if not word or conf < 30:
                        continue
                    key = (data["block_num"][i], data["par_num"][i], data["line_num"][i])
                    h_pts = data["height"][i] * scale_y
                    y_pts = data["top"][i]    * scale_y
                    groups[key].append((word, h_pts, y_pts))

                for key in sorted(groups.keys()):
                    words  = groups[key]
                    text   = " ".join(w[0] for w in words)
                    avg_h  = sum(w[1] for w in words) / max(len(words), 1)
                    y_pos  = words[0][2]
                    font_pt = min(max(round(avg_h * 0.72), 8), 36)
                    lines_out.append((y_pos, font_pt, text.strip()))
            except Exception as e:
                logger.warning("  Tesseract page %d error: %s", pn+1, e)
                doc_out.add_paragraph("[OCR failed on page %d: %s]" % (pn+1, e))
                continue

        #  Write to DOCX 
        for y_pos, font_pt, text in lines_out:
            if not text:
                continue
            translated = tr(text)
            is_heading = font_pt >= 14 and len(translated) < 150
            if is_heading:
                level = 1 if font_pt >= 20 else (2 if font_pt >= 16 else 3)
                p = doc_out.add_heading(translated, level=level)
            else:
                p   = doc_out.add_paragraph()
                run = p.add_run(translated)
                run.font.size = Pt(font_pt)
            p.paragraph_format.space_before = Pt(0)
            p.paragraph_format.space_after  = Pt(max(round(font_pt * 0.2), 2))

    doc_pdf.close()
    doc_out.save(str(docx_path))
    logger.info("OCR->DOCX complete (%s): %s", engine, docx_path.name)



def pdf_to_docx(pdf_path, docx_path, lang="pt", do_translate=True, image_mode="placeholder"):
    """
    Convert a text-based PDF to DOCX with real selectable text.
    Extracts text blocks with PyMuPDF, preserves formatting,
    translates if requested, and replaces images with markers.
    """
    import fitz
    from docx import Document
    from docx.shared import Pt, RGBColor, Cm
    from docx.oxml.ns import qn
    from docx.oxml import OxmlElement
    from collections import defaultdict

    # Setup translator
    translator = None
    if do_translate and lang not in ("", "none"):
        try:
            from deep_translator import GoogleTranslator
            translator = GoogleTranslator(source="auto", target=lang)
        except ImportError:
            pass

    def tr(text):
        if translator and text and text.strip():
            return translate_text(text, lang, translator)
        return text

    def make_placeholder(doc, label):
        p   = doc.add_paragraph()
        run = p.add_run(label)
        run.font.size      = Pt(9)
        run.font.italic    = True
        run.font.color.rgb = RGBColor(0x88, 0x88, 0x88)
        pPr = p._p.get_or_add_pPr()
        shd = OxmlElement("w:shd")
        shd.set(qn("w:val"), "clear")
        shd.set(qn("w:color"), "auto")
        shd.set(qn("w:fill"), "F0F0F0")
        pPr.append(shd)
        p.paragraph_format.space_before = Pt(2)
        p.paragraph_format.space_after  = Pt(2)
        return p

    doc_out = Document()
    for sec in doc_out.sections:
        sec.top_margin    = Cm(2.54)
        sec.bottom_margin = Cm(2.54)
        sec.left_margin   = Cm(3.17)
        sec.right_margin  = Cm(3.17)

    doc_pdf = fitz.open(str(pdf_path))
    logger.info("Extracting text from %d pages...", len(doc_pdf))

    for pn, page in enumerate(doc_pdf):
        if pn > 0:
            doc_out.add_page_break()

        pw = page.rect.width
        ph = page.rect.height

        # Get blocks sorted in reading order
        raw = page.get_text("dict", flags=fitz.TEXT_PRESERVE_WHITESPACE)["blocks"]
        blocks = sorted(raw, key=lambda b: (round(b["bbox"][1] / 12) * 12, b["bbox"][0]))

        for block in blocks:
            btype = block.get("type", -1)

            #  IMAGE 
            if btype == 1:
                if image_mode == "remove":
                    continue
                bbox = block["bbox"]
                x0,y0,x1,y1 = bbox
                w = x1-x0; h = y1-y0
                area = (w*h)/(pw*ph)
                aspect = w/max(h,1)
                in_header = y0 < ph*0.2
                in_footer = y1 > ph*0.8
                if area < 0.10 and (in_header or in_footer):
                    label = "[LOGOMARCA]"
                elif area < 0.10 and 0.5 < aspect < 2.0:
                    label = "[CARIMBO]"
                elif area > 0.30:
                    label = "[FIGURA]"
                else:
                    label = "[LOGOMARCA]"
                if image_mode == "placeholder":
                    make_placeholder(doc_out, label)
                continue

            #  TEXT 
            if btype != 0:
                continue

            # Collect all lines in block, preserving span formatting
            for line in block.get("lines", []):
                spans = line.get("spans", [])
                if not spans:
                    continue

                # Merge spans on same line into one paragraph
                # Group by dominant formatting
                line_text = ""
                dom_size  = 11.0
                dom_bold  = False
                dom_italic = False
                dom_font  = "Calibri"

                for span in spans:
                    t = span.get("text", "")
                    if not t.strip():
                        continue
                    line_text  += t
                    sz = span.get("size", 11.0)
                    if sz > dom_size:   # dominant = largest
                        dom_size   = sz
                        dom_bold   = bool(span.get("flags", 0) & 16)
                        dom_italic = bool(span.get("flags", 0) & 2)
                        dom_font   = span.get("font", "Calibri")

                line_text = line_text.strip()
                if not line_text:
                    continue

                font_pt    = min(max(round(dom_size), 7), 72)
                translated = tr(line_text)

                # Heading detection
                is_heading = (font_pt >= 13 and dom_bold and len(translated) < 200)                              or font_pt >= 16

                if is_heading:
                    level = 1 if font_pt >= 20 else (2 if font_pt >= 15 else 3)
                    p = doc_out.add_heading(translated, level=level)
                else:
                    p   = doc_out.add_paragraph()
                    run = p.add_run(translated)
                    run.bold      = dom_bold
                    run.italic    = dom_italic
                    run.font.size = Pt(font_pt)
                    # Preserve font family
                    clean_font = dom_font.split("+")[-1].strip()
                    if clean_font:
                        try:
                            from docx.oxml.ns import qn as qn2
                            rPr = run._r.get_or_add_rPr()
                            rFonts = OxmlElement("w:rFonts")
                            rFonts.set(qn2("w:ascii"),    clean_font)
                            rFonts.set(qn2("w:hAnsi"),    clean_font)
                            rFonts.set(qn2("w:cs"),       clean_font)
                            rPr.append(rFonts)
                        except Exception:
                            pass

                p.paragraph_format.space_before = Pt(0)
                p.paragraph_format.space_after  = Pt(round(font_pt * 0.2))

    doc_pdf.close()

    # Post-process: handle translation if not done inline
    doc_out.save(str(docx_path))
    logger.info("Text extraction complete -> %s", docx_path.name)


def _translate_docx_inplace(docx_path, translator, lang):
    """Translate all text in an existing DOCX file in-place."""
    from docx import Document
    doc = Document(str(docx_path))
    done = 0

    # Paragraphs
    for para in doc.paragraphs:
        if para.text.strip():
            translated = translate_text(para.text, lang, translator)
            if translated and translated != para.text:
                # Replace text preserving runs formatting
                if len(para.runs) == 1:
                    para.runs[0].text = translated
                else:
                    # Clear all runs, put text in first
                    for i, run in enumerate(para.runs):
                        run.text = translated if i == 0 else ""
            done += 1

    # Tables
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for para in cell.paragraphs:
                    if para.text.strip():
                        translated = translate_text(para.text, lang, translator)
                        if translated and translated != para.text:
                            if len(para.runs) == 1:
                                para.runs[0].text = translated
                            elif para.runs:
                                for i, run in enumerate(para.runs):
                                    run.text = translated if i == 0 else ""
                        done += 1

    doc.save(str(docx_path))
    logger.info("Translated %d paragraphs in DOCX", done)


def _replace_images_in_docx(docx_path):
    """
    Remove ALL images from DOCX and replace with text placeholders.
    Works on paragraphs, tables, headers, footers.
    Dramatically reduces file size.
    """
    from docx import Document
    from docx.oxml.ns import qn
    from docx.oxml import OxmlElement
    from docx.shared import Pt, RGBColor
    import zipfile, shutil, os

    # Step 1: Remove image files from the DOCX zip (reduces size immediately)
    docx_str  = str(docx_path)
    temp_path = docx_str + ".tmp"
    image_exts = {".png", ".jpg", ".jpeg", ".gif", ".bmp", ".tiff", ".emf", ".wmf"}
    removed_media = 0

    with zipfile.ZipFile(docx_str, 'r') as zin:
        with zipfile.ZipFile(temp_path, 'w', zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                # Skip media files
                ext = os.path.splitext(item.filename)[1].lower()
                if item.filename.startswith('word/media/') and ext in image_exts:
                    removed_media += 1
                    continue   # don't copy to output
                zout.writestr(item, zin.read(item.filename))

    os.replace(temp_path, docx_str)
    logger.info("Removed %d image files from DOCX archive", removed_media)

    # Step 2: Replace drawing XML elements with placeholder text runs
    doc = Document(docx_str)
    replaced = 0

    def replace_drawings_in_element(parent_el):
        """Recursively find w:drawing elements and replace with text."""
        nonlocal replaced
        ns_w = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'

        # Find all paragraphs under this element
        for p_el in parent_el.iter('{%s}p' % ns_w):
            drawings = p_el.findall('.//{%s}drawing' % ns_w)
            if not drawings:
                continue

            # Determine placeholder label from position in doc
            # (simple: alternate between LOGOMARCA and CARIMBO based on size)
            for drawing in drawings:
                # Try to get image dimensions from drawing XML
                # cx/cy are in EMUs (914400 EMUs = 1 inch)
                extent = drawing.find('.//{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}extent')
                label = "[LOGOMARCA]"
                if extent is not None:
                    try:
                        cx = int(extent.get('cx', 0))
                        cy = int(extent.get('cy', 0))
                        # Small square image = stamp/seal
                        if cx > 0 and cy > 0:
                            aspect = cx / cy
                            size_in = cx / 914400
                            if size_in < 2.0 and 0.5 < aspect < 2.0:
                                label = "[CARIMBO]"
                            elif size_in > 4.0:
                                label = "[FIGURA]"
                    except Exception:
                        pass

                # Remove the drawing element from its parent run
                parent_r = drawing.getparent()
                if parent_r is not None:
                    parent_r.remove(drawing)

                # Create a new run with placeholder text in this paragraph
                new_r = OxmlElement('w:r')
                rpr   = OxmlElement('w:rPr')
                # italic
                i_el  = OxmlElement('w:i')
                rpr.append(i_el)
                # color grey
                color_el = OxmlElement('w:color')
                color_el.set(qn('w:val'), '888888')
                rpr.append(color_el)
                # size 9pt = 18 half-points
                sz_el = OxmlElement('w:sz')
                sz_el.set(qn('w:val'), '18')
                rpr.append(sz_el)
                new_r.append(rpr)
                t_el  = OxmlElement('w:t')
                t_el.text = label
                t_el.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
                new_r.append(t_el)
                p_el.append(new_r)
                replaced += 1

    # Process main document body
    replace_drawings_in_element(doc.element.body)

    # Process tables
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                replace_drawings_in_element(cell._element)

    # Process headers and footers
    for section in doc.sections:
        for hf in [section.header, section.footer,
                   section.even_page_header, section.even_page_footer,
                   section.first_page_header, section.first_page_footer]:
            try:
                if hf and hf._element is not None:
                    replace_drawings_in_element(hf._element)
            except Exception:
                pass

    doc.save(docx_str)
    size_kb = docx_path.stat().st_size / 1024
    logger.info("Images replaced: %d placeholders, final size: %.1f KB",
                replaced, size_kb)


#  FILE PROCESSOR 

def process_file(file_info, out_dir, lang, image_mode, do_translate):
    pdf = Path(file_info["upload_path"])
    result = dict(file_info)
    try:
        if not pdf.exists():
            raise FileNotFoundError(str(pdf))

        docx_name = pdf.stem + "_convertido.docx"
        docx_path = out_dir / docx_name

        # Detect if PDF is scanned (image-based)
        scanned = is_scanned_pdf(pdf)
        result["pdf_type"] = "scanned" if scanned else "text"

        if scanned:
            logger.info("Scanned PDF detected -> using OCR pipeline: %s", pdf.name)
            # Option A: OCR directly to DOCX (faster, good for simple layouts)
            ocr_pdf_to_docx_direct(
                pdf, docx_path,
                lang=lang,
                do_translate=do_translate,
                image_mode=image_mode,
                lang_hint="por+eng"
            )
        else:
            logger.info("Text PDF detected -> using pdf2docx: %s", pdf.name)
            pdf_to_docx(pdf, docx_path, lang=lang,
                        do_translate=do_translate, image_mode=image_mode)

        if not docx_path.exists():
            raise RuntimeError("DOCX not created: " + str(docx_path))

        size_kb = docx_path.stat().st_size / 1024
        logger.info("Done: %s (%.1f KB)", docx_name, size_kb)

        result.update({
            "output_path":  str(docx_path),
            "output_name":  docx_name,
            "size_kb":      round(size_kb, 1),
            "status":       "completed",
        })
        return result

    except Exception as e:
        logger.error("Error processing %s: %s", file_info["name"], e)
        result["error"]  = str(e)
        result["status"] = "failed"
        return result


async def run_job(job_id, lang, image_mode, do_translate):
    job     = load_job(job_id)
    out_dir = OUTPUT_DIR / job_id
    out_dir.mkdir(parents=True, exist_ok=True)
    job["status"]     = "processing"
    job["started_at"] = datetime.now().isoformat()
    save_job(job)

    loop = asyncio.get_event_loop()
    for idx, fi in enumerate(job["files"]):
        fi["status"] = "processing"
        save_job(job)

        updated = await loop.run_in_executor(
            None, process_file, fi, out_dir, lang, image_mode, do_translate)

        job["files"][idx] = updated
        if updated.get("status") == "completed":
            job["processed"] += 1
        else:
            job["failed"] += 1
        job["percentage"] = round(
            (job["processed"] + job["failed"]) / job["total"] * 100, 1)
        save_job(job)
        await asyncio.sleep(0.1)

    # Create ZIP
    zip_path = OUTPUT_DIR / (job_id + "_all.zip")
    added = 0
    with zipfile.ZipFile(str(zip_path), "w", zipfile.ZIP_DEFLATED) as zf:
        for fi in job["files"]:
            p = fi.get("output_path", "")
            if p and Path(p).exists():
                zf.write(p, fi["output_name"])
                added += 1

    logger.info("ZIP: %d files, %.1f KB", added,
                zip_path.stat().st_size / 1024 if zip_path.exists() else 0)

    job["status"]       = "completed"
    job["zip_path"]     = str(zip_path)
    job["output_dir"]   = str(out_dir)
    job["completed_at"] = datetime.now().isoformat()
    save_job(job)


#  ROUTES 

@app.get("/", response_class=HTMLResponse)
async def root():
    index = STATIC_DIR / "index.html"
    if not index.exists():
        return HTMLResponse(
            "<h2>index.html not found</h2><p>Place in: <code>" + str(STATIC_DIR) + "</code></p>",
            status_code=404)
    return HTMLResponse(content=index.read_text(encoding="utf-8"))


@app.get("/api/health")
async def health():
    deps = {}
    for lib, label in [("fitz", "PyMuPDF"), ("pytesseract", "Tesseract"),
                        ("deep_translator", "deep-translator"),
                        ("docx", "python-docx"), ("pdf2docx", "pdf2docx")]:
        try:
            __import__(lib)
            deps[label] = "ok"
        except ImportError:
            deps[label] = "not installed"
    return {
        "status": "ok",
        "base_dir":    str(BASE_DIR),
        "output_dir":  str(OUTPUT_DIR),
        "index_html_exists": (STATIC_DIR / "index.html").exists(),
        "dependencies": deps,
    }


@app.post("/api/upload")
async def upload_files(files: List[UploadFile] = File(...)):
    job_id = str(uuid.uuid4())
    up_dir = UPLOAD_DIR / job_id
    up_dir.mkdir(parents=True, exist_ok=True)
    pdfs   = []

    for upload in files:
        filename  = upload.filename or ("file_" + str(uuid.uuid4()) + ".pdf")
        safe_name = "".join(c for c in filename
                            if c.isalnum() or c in "._- ()").strip() or "file.pdf"
        dest    = up_dir / safe_name
        content = await upload.read()
        dest.write_bytes(content)

        if safe_name.lower().endswith(".zip"):
            try:
                with zipfile.ZipFile(str(dest)) as zf:
                    for name in zf.namelist():
                        if name.lower().endswith(".pdf") and not name.startswith("__MACOSX"):
                            ep = up_dir / Path(name).name
                            ep.write_bytes(zf.read(name))
                            pdfs.append({"id": str(uuid.uuid4()), "name": ep.name,
                                         "size": ep.stat().st_size,
                                         "upload_path": str(ep), "status": "pending"})
                dest.unlink()
            except zipfile.BadZipFile:
                raise HTTPException(400, safe_name + " is not a valid ZIP")
        elif safe_name.lower().endswith(".pdf"):
            pdfs.append({"id": str(uuid.uuid4()), "name": safe_name,
                         "size": len(content), "upload_path": str(dest), "status": "pending"})

    if not pdfs:
        raise HTTPException(400, "No PDF found.")

    job = {"job_id": job_id, "status": "pending", "total": len(pdfs),
           "processed": 0, "failed": 0, "percentage": 0.0, "files": pdfs,
           "created_at": datetime.now().isoformat(),
           "target_language": "pt", "image_mode": "placeholder"}
    save_job(job)
    return {"job_id": job_id, "files_count": len(pdfs), "files": pdfs}


@app.post("/api/process")
async def start_processing(req: ProcessRequest, background_tasks: BackgroundTasks):
    job = load_job(req.job_id)
    if job["status"] == "processing":
        raise HTTPException(400, "Already processing")

    # do_translate = False if language is "none" or "original"
    do_translate = req.target_language not in ("none", "original", "")
    image_mode   = "remove" if req.remove_images else req.image_mode

    job.update({"target_language": req.target_language, "image_mode": image_mode,
                "status": "processing", "processed": 0, "failed": 0, "percentage": 0.0})
    for f in job["files"]:
        f["status"] = "pending"
    save_job(job)

    background_tasks.add_task(run_job, req.job_id, req.target_language,
                              image_mode, do_translate)
    return {"status": "started", "job_id": req.job_id}


@app.get("/api/jobs/{job_id}")
async def get_job(job_id: str):
    return load_job(job_id)


@app.get("/api/jobs")
async def get_all_jobs():
    return list_jobs()


@app.get("/api/outputs/{job_id}")
async def list_outputs(job_id: str):
    job     = load_job(job_id)
    out_dir = OUTPUT_DIR / job_id
    files   = []
    if out_dir.exists():
        for f in sorted(out_dir.iterdir()):
            if f.suffix in (".docx", ".pdf"):
                files.append({"name": f.name, "path": str(f),
                               "size_kb": round(f.stat().st_size / 1024, 1)})
    return {"job_id": job_id, "output_dir": str(out_dir), "files": files,
            "job_files": [{"name": fi.get("name"), "output_path": fi.get("output_path",""),
                           "output_name": fi.get("output_name",""), "status": fi.get("status"),
                           "error": fi.get("error","")} for fi in job.get("files",[])]}


@app.get("/api/download/{job_id}/zip")
async def download_zip(job_id: str):
    job = load_job(job_id)
    zp  = Path(job.get("zip_path", ""))
    if not zp.exists():
        raise HTTPException(404, "ZIP not ready yet")
    return FileResponse(str(zp), filename="conversao_" + job_id[:8] + ".zip",
                        media_type="application/zip")


@app.get("/api/download/{job_id}/file/{file_id}")
async def download_file(job_id: str, file_id: str):
    job = load_job(job_id)
    fi  = next((f for f in job["files"] if f["id"] == file_id), None)
    if not fi:
        raise HTTPException(404, "File not found")
    op = Path(fi.get("output_path", ""))
    if not op.exists():
        raise HTTPException(404, "DOCX not ready yet")
    return FileResponse(str(op), filename=fi["output_name"],
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document")


@app.delete("/api/jobs/{job_id}")
async def delete_job(job_id: str):
    load_job(job_id)
    for d in [UPLOAD_DIR / job_id, OUTPUT_DIR / job_id]:
        if d.exists():
            shutil.rmtree(d)
    zp = OUTPUT_DIR / (job_id + "_all.zip")
    if zp.exists():
        zp.unlink()
    (JOBS_DIR / (job_id + ".json")).unlink(missing_ok=True)
    return {"deleted": True}


if __name__ == "__main__":
    import uvicorn
    uvicorn.run("main:app", host="0.0.0.0", port=8000, reload=True,
                reload_dirs=[str(BASE_DIR)])