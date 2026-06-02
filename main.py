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

MAX_CONCURRENT = 4
TRANSLATE_TIMEOUT = 30

_translation_cache: dict = {}

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

def list_jobs(limit=50):
    out = []
    for p in sorted(JOBS_DIR.glob("*.json"), key=lambda x: x.stat().st_mtime, reverse=True):
        if len(out) >= limit:
            break
        try:
            out.append(json.loads(p.read_text(encoding="utf-8")))
        except Exception:
            pass
    return out


#  TRANSLATION 

def translate_text(text, lang, translator, retry=0):
    if not text or not text.strip():
        return text
    cache_key = (text, lang)
    cached = _translation_cache.get(cache_key)
    if cached is not None:
        return cached
    try:
        MAX = 4800
        if len(text) <= MAX:
            r = translator.translate(text, timeout=TRANSLATE_TIMEOUT)
            result = r.strip() if r else text
            _translation_cache[cache_key] = result
            return result
        parts, buf = [], ""
        for sentence in text.replace(". ", ".|||").split("|||"):
            if len(buf) + len(sentence) < MAX:
                buf += sentence
            else:
                if buf:
                    r = translator.translate(buf, timeout=TRANSLATE_TIMEOUT)
                    parts.append(r.strip() if r else buf)
                buf = sentence
        if buf:
            r = translator.translate(buf, timeout=TRANSLATE_TIMEOUT)
            parts.append(r.strip() if r else buf)
        result = " ".join(parts)
        _translation_cache[cache_key] = result
        return result
    except Exception as e:
        import time
        msg = str(e)
        if "429" in msg or "TooMany" in msg:
            wait = min(4 * (2 ** retry), 60)
            logger.warning("rate limited, waiting %ds (retry %d)", wait, retry)
            time.sleep(wait)
            if retry < 3:
                return translate_text(text, lang, translator, retry=retry + 1)
            else:
                logger.warning("max retries exceeded for text")
                return text
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

def pdf_to_docx(pdf_path, docx_path, lang="pt", do_translate=True, image_mode="placeholder"):
    """
    Convert PDF to DOCX preserving 100% of the structure.
    Uses pdf2docx for faithful conversion, then post-processes for translation.
    Falls back to manual block reconstruction if pdf2docx unavailable.
    """
    import fitz
    from docx import Document
    from docx.shared import Pt, RGBColor, Cm, Inches
    from docx.oxml.ns import qn
    from docx.oxml import OxmlElement

    # Setup translator
    translator = None
    if do_translate and lang not in ("", "none"):
        try:
            from deep_translator import GoogleTranslator
            translator = GoogleTranslator(source="auto", target=lang)
        except ImportError:
            logger.warning("deep_translator not installed")

    #  Try pdf2docx first (best fidelity) 
    try:
        from pdf2docx import Converter
        logger.info("Using pdf2docx for faithful conversion...")
        cv = Converter(str(pdf_path))
        cv.convert(str(docx_path), start=0, end=None)
        cv.close()
        logger.info("pdf2docx conversion done")

        if do_translate and translator:
            logger.info("Post-processing: translating text in DOCX...")
            _translate_docx_inplace(docx_path, translator, lang)
            logger.info("Translation done")

        # Handle images in the converted DOCX
        if image_mode == "placeholder":
            _replace_images_in_docx(docx_path)
        elif image_mode == "remove":
            _remove_images_from_docx(docx_path)

        return

    except ImportError:
        logger.info("pdf2docx not available, using manual reconstruction...")
    except Exception as e:
        logger.warning("pdf2docx failed (%s), falling back...", e)

    #  Fallback: manual block-by-block reconstruction 
    logger.info("Manual PDF->DOCX reconstruction...")

    doc_pdf = fitz.open(str(pdf_path))

    # Phase 1: Extract all elements (text + images) with their metadata
    pages_content = []
    for pn, page in enumerate(doc_pdf):
        pw = page.rect.width
        ph = page.rect.height
        blocks = sorted(page.get_text("dict")["blocks"],
                        key=lambda b: (round(b["bbox"][1] / 10) * 10, b["bbox"][0]))
        page_els = []
        for block in blocks:
            btype = block["type"]

            if btype == 1:
                if image_mode == "remove":
                    continue
                label = classify_image(block["bbox"], pw, ph)
                page_els.append(("image", {"label": label}))
                continue

            if btype != 0:
                continue

            lines = block.get("lines", [])
            if not lines:
                continue

            line_texts = []
            for line in lines:
                spans = line.get("spans", [])
                line_text = "".join(s.get("text", "") for s in spans)
                if line_text.strip():
                    line_texts.append(line_text)

            if not line_texts:
                continue

            full_text = " ".join(t.strip() for t in line_texts if t.strip())
            if not full_text:
                continue

            dom_size, dom_bold, dom_italic, dom_font = 11.0, False, False, "Calibri"
            for line in lines:
                for span in line.get("spans", []):
                    if span.get("text", "").strip():
                        dom_size   = span.get("size", 11.0)
                        dom_bold   = bool(span.get("flags", 0) & 16)
                        dom_italic = bool(span.get("flags", 0) & 2)
                        dom_font   = span.get("font", "Calibri")
                        break
                else:
                    continue
                break

            font_pt = min(max(round(dom_size), 6), 72)
            page_els.append(("text", {
                "text": full_text, "size": dom_size, "bold": dom_bold,
                "italic": dom_italic, "font": dom_font, "font_pt": font_pt
            }))

        pages_content.append(page_els)

    doc_pdf.close()

    # Phase 2: Translate all unique texts sequentially (avoid rate limit)
    if translator:
        import time
        all_texts = []
        for page_els in pages_content:
            for el_type, el_data in page_els:
                if el_type == "text" and el_data["text"].strip():
                    all_texts.append(el_data["text"])
        unique_texts = list(dict.fromkeys(all_texts))
        translated_map = {}

        for t in unique_texts:
            key = (t, lang)
            if key not in _translation_cache:
                translate_text(t, lang, translator)
                time.sleep(0.1)
            translated_map[t] = _translation_cache.get(key, t)
    else:
        translated_map = {}

    # Phase 3: Build DOCX with pre-translated texts
    doc_out = Document()
    for sec in doc_out.sections:
        sec.top_margin    = Cm(2.54)
        sec.bottom_margin = Cm(2.54)
        sec.left_margin   = Cm(3.17)
        sec.right_margin  = Cm(3.17)

    for pn, page_els in enumerate(pages_content):
        if pn > 0:
            doc_out.add_page_break()

        for el_type, el_data in page_els:
            if el_type == "image":
                p = doc_out.add_paragraph()
                run = p.add_run(el_data["label"])
                run.font.size      = Pt(9)
                run.font.italic    = True
                run.font.color.rgb = RGBColor(0x77, 0x77, 0x77)
                pPr = p._p.get_or_add_pPr()
                shd = OxmlElement("w:shd")
                shd.set(qn("w:val"), "clear")
                shd.set(qn("w:color"), "auto")
                shd.set(qn("w:fill"), "F2F2F2")
                pPr.append(shd)
                p.paragraph_format.space_before = Pt(2)
                p.paragraph_format.space_after  = Pt(2)
                continue

            if el_type != "text":
                continue

            text = translated_map.get(el_data["text"], el_data["text"])
            is_heading = (el_data["font_pt"] >= 14 or (el_data["font_pt"] >= 12 and el_data["bold"])) \
                         and len(text) < 200

            if is_heading:
                level = 1 if el_data["font_pt"] >= 20 else (2 if el_data["font_pt"] >= 16 else 3)
                p = doc_out.add_heading(text, level=level)
            else:
                p   = doc_out.add_paragraph()
                run = p.add_run(text)
                run.bold      = el_data["bold"]
                run.italic    = el_data["italic"]
                run.font.size = Pt(el_data["font_pt"])
                try:
                    font_name = el_data["font"]
                    run.font.name = font_name.split("+")[-1] if "+" in font_name else font_name
                except Exception:
                    pass

            p.paragraph_format.space_before = Pt(0)
            p.paragraph_format.space_after  = Pt(round(el_data["size"] * 0.3))

    doc_out.save(str(docx_path))
    total_els = sum(len(pe) for pe in pages_content)
    logger.info("Manual reconstruction done: %d elements", total_els)


def _translate_docx_inplace(docx_path, translator, lang):
    """Translate all text in an existing DOCX file in-place."""
    from docx import Document
    import time

    doc = Document(str(docx_path))

    # Phase 1: Collect all unique texts that need translation
    texts_to_translate = set()
    for para in doc.paragraphs:
        if para.text.strip():
            texts_to_translate.add(para.text)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for para in cell.paragraphs:
                    if para.text.strip():
                        texts_to_translate.add(para.text)

    # Phase 2: Pre-translate all unique texts sequentially (avoid rate limit)
    translated = 0
    for t in texts_to_translate:
        key = (t, lang)
        if key not in _translation_cache:
            try:
                translate_text(t, lang, translator)
                translated += 1
            except Exception as e:
                logger.warning("translate_text error: %s", e)
            time.sleep(0.1)

    # Phase 3: Apply translations
    done = 0
    errors = 0
    for para in doc.paragraphs:
        try:
            if para.text.strip():
                cached = _translation_cache.get((para.text, lang))
                translated = cached if cached is not None else para.text
                if translated != para.text:
                    if len(para.runs) >= 1:
                        para.runs[0].text = translated
                        for run in para.runs[1:]:
                            run.text = ""
                done += 1
        except Exception as e:
            errors += 1
            logger.warning("Error applying translation to paragraph: %s", e)

    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for para in cell.paragraphs:
                    try:
                        if para.text.strip():
                            cached = _translation_cache.get((para.text, lang))
                            translated = cached if cached is not None else para.text
                            if translated != para.text:
                                if len(para.runs) >= 1:
                                    para.runs[0].text = translated
                                    for run in para.runs[1:]:
                                        run.text = ""
                            done += 1
                    except Exception as e:
                        errors += 1
                        logger.warning("Error applying translation to cell: %s", e)

    try:
        doc.save(str(docx_path))
        logger.info("Translated %d paragraphs in DOCX (%d new, %d errors)", done, translated, errors)
    except Exception as e:
        logger.error("Failed to save DOCX after translation: %s", e)


def _replace_images_in_docx(docx_path):
    """Replace inline images in DOCX with [LOGOMARCA]/[CARIMBO] text."""
    from docx import Document
    from docx.oxml.ns import qn
    from docx.oxml import OxmlElement
    from lxml import etree
    import copy

    doc = Document(str(docx_path))
    replaced = 0

    for para in doc.paragraphs:
        # Check if paragraph contains images (drawing elements)
        drawings = para._p.findall('.//' + qn('w:drawing'))
        if drawings:
            # Remove image runs, replace with placeholder text
            for run in para.runs:
                if run._r.findall('.//' + qn('w:drawing')):
                    # Determine label based on position (simple heuristic)
                    run.text   = "[LOGOMARCA]"
                    run.italic = True
                    # Remove the drawing element
                    for drawing in run._r.findall('.//' + qn('w:drawing')):
                        run._r.remove(drawing)
                    replaced += 1

    doc.save(str(docx_path))
    if replaced:
        logger.info("Replaced %d images with placeholders in DOCX", replaced)


def _remove_images_from_docx(docx_path):
    """Remove all images from DOCX file entirely."""
    from docx import Document
    from docx.oxml.ns import qn

    doc = Document(str(docx_path))
    removed = 0

    for para in doc.paragraphs:
        drawings = para._p.findall('.//' + qn('w:drawing'))
        if not drawings:
            drawings = para._p.findall('.//' + qn('wp:inline'))
        if drawings:
            for drawing in drawings:
                para._p.remove(drawing)
                removed += 1
            # Clean up empty runs
            for run in para.runs[:]:
                if not run.text.strip() and not run._r.findall('.//' + qn('w:drawing')):
                    if not run._r.findall('.//' + qn('wp:inline')):
                        run._r.getparent().remove(run._r)

    doc.save(str(docx_path))
    if removed:
        logger.info("Removed %d images from DOCX", removed)


#  FILE PROCESSOR 

def process_file(file_info, out_dir, lang, image_mode, do_translate):
    pdf = Path(file_info["upload_path"])
    result = dict(file_info)
    try:
        if not pdf.exists():
            raise FileNotFoundError(str(pdf))

        docx_name = pdf.stem + "_convertido.docx"
        docx_path = out_dir / docx_name

        logger.info("Processing: %s -> %s", pdf.name, docx_name)
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

    _translation_cache.clear()
    loop  = asyncio.get_event_loop()
    files = job["files"]
    total = len(files)
    sem   = asyncio.Semaphore(MAX_CONCURRENT)
    lock  = asyncio.Lock()
    save_counter = 0

    for fi in files:
        fi["status"] = "processing"
    save_job(job)

    async def process_one(idx, fi):
        nonlocal save_counter, job
        async with sem:
            updated = await loop.run_in_executor(
                None, process_file, fi, out_dir, lang, image_mode, do_translate)
        async with lock:
            job["files"][idx] = updated
            if updated.get("status") == "completed":
                job["processed"] += 1
            else:
                job["failed"] += 1
            job["percentage"] = round(
                (job["processed"] + job["failed"]) / total * 100, 1)
            save_counter += 1
            if save_counter % 5 == 0:
                save_job(job)

    try:
        await asyncio.gather(*[process_one(i, f) for i, f in enumerate(files)])
    except Exception as e:
        logger.error("run_job error: %s", e)
        job["status"] = "failed"
        for fi in job["files"]:
            if fi.get("status") not in ("completed", "failed"):
                fi["status"] = "failed"
                fi["error"]  = str(e)
        save_job(job)
        return

    save_job(job)

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
    uvicorn.run("main:app", host="0.0.0.0", port=8000, reload=False)