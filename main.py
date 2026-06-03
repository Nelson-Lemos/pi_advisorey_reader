from fastapi import FastAPI, UploadFile, File, BackgroundTasks, HTTPException
from fastapi.staticfiles import StaticFiles
from fastapi.responses import FileResponse, HTMLResponse
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel
from typing import List
import uuid, json, zipfile, shutil, asyncio, re, time, io
from pathlib import Path
from datetime import datetime
import logging

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

MAX_CONCURRENT = 4
TRANSLATE_TIMEOUT = 30
FILE_TIMEOUT_SECONDS = 180
MAX_UPLOAD_BYTES = 100 * 1024 * 1024
MAX_ZIP_MEMBERS = 100
MAX_ZIP_TOTAL_BYTES = 100 * 1024 * 1024

_translation_cache: dict = {}

app = FastAPI(title="PDF Tradutor Pro", version="6.0.0")
app.add_middleware(CORSMiddleware, allow_origins=["*"], allow_credentials=True,
                   allow_methods=["*"], allow_headers=["*"])

BASE_DIR   = Path(__file__).resolve().parent
STATIC_DIR = BASE_DIR / "static"
UPLOAD_DIR = BASE_DIR / "uploads"
OUTPUT_DIR = BASE_DIR / "outputs"
JOBS_DIR   = BASE_DIR / "jobs"

for d in [STATIC_DIR, UPLOAD_DIR, OUTPUT_DIR, JOBS_DIR]:
    d.mkdir(parents=True, exist_ok=True)

app.mount("/static", StaticFiles(directory=str(STATIC_DIR)), name="static")


class ProcessRequest(BaseModel):
    job_id: str
    target_language: str = "pt"
    image_mode: str      = "placeholder"
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


def sanitize_filename(filename):
    return "".join(c for c in filename if c.isalnum() or c in "._- ()").strip() or "file.pdf"


def translate_text(text, lang, translator, retry=0, cache=None):
    if not text or not text.strip():
        return text
    cache = _translation_cache if cache is None else cache
    cache_key = (text, lang)
    cached = cache.get(cache_key)
    if cached is not None:
        return cached
    try:
        MAX = 4800
        if len(text) <= MAX:
            r = translator.translate(text, timeout=TRANSLATE_TIMEOUT)
            result = r.strip() if r else text
            cache[cache_key] = result
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
        cache[cache_key] = result
        return result
    except Exception as e:
        msg = str(e)
        if "429" in msg or "TooMany" in msg:
            wait = min(4 * (2 ** retry), 60)
            logger.warning("rate limited, waiting %ds (retry %d)", wait, retry)
            time.sleep(wait)
            if retry < 3:
                return translate_text(text, lang, translator, retry=retry + 1, cache=cache)
            else:
                logger.warning("max retries exceeded")
                return text
        logger.warning("translate error: %s", e)
        return text


def translate_markdown(markdown_text, lang, translator):
    lines = markdown_text.split("\n")
    translated_lines = []
    in_code_block = False
    for line in lines:
        stripped = line.strip()
        if stripped.startswith("```"):
            in_code_block = not in_code_block
            translated_lines.append(line)
            continue
        if in_code_block:
            translated_lines.append(line)
            continue
        if stripped.startswith("#") or stripped.startswith(">") or stripped.startswith("```"):
            translated_lines.append(line)
            continue
        if stripped.startswith("!["):
            translated_lines.append(line)
            continue
        if re.match(r'^[\-\*\+]\s', stripped) or re.match(r'^\d+[\.\)]\s', stripped):
            translated_lines.append(line)
            continue
        if "|" in stripped and re.match(r'^[\s\|:\-a-zA-Z0-9]+$', stripped):
            translated_lines.append(line)
            continue
        if stripped == "---" or stripped == "***":
            translated_lines.append(line)
            continue
        if not stripped:
            translated_lines.append(line)
            continue
        translated = translate_text(stripped, lang, translator)
        if translated and translated != stripped:
            indent = line[:len(line) - len(line.lstrip())]
            translated_lines.append(indent + translated)
        else:
            translated_lines.append(line)
    return "\n".join(translated_lines)


def convert_pdf_to_docx(pdf_path, docx_path, lang="pt", do_translate=True, image_mode="placeholder"):
    logger.info("Converting PDF to Markdown via Marker: %s", pdf_path.name)

    from marker.converters.pdf import PdfConverter
    from marker.models import create_model_dict
    from marker.output import text_from_rendered

    converter = PdfConverter(
        artifact_dict=create_model_dict(),
    )
    rendered = converter(str(pdf_path))
    markdown_text, metadata, images_by_block = text_from_rendered(rendered)

    total_chars = len(markdown_text)
    logger.info("Marker extracted %d chars", total_chars)

    image_mode_effective = "remove" if image_mode == "remove" else image_mode

    if do_translate and lang not in ("", "none", "original"):
        from deep_translator import GoogleTranslator
        translator = GoogleTranslator(source="auto", target=lang)
        markdown_text = translate_markdown(markdown_text, lang, translator)
        logger.info("Translation done")

    if image_mode_effective == "remove":
        markdown_text = re.sub(r'!\[[^\]]*\]\([^)]+\)', "", markdown_text)
        images_used = {}
    elif image_mode_effective == "placeholder":
        def _replace_img(m):
            alt = m.group(1) or "IMAGEM"
            return f"[{alt}]"
        markdown_text = re.sub(r'!\[([^\]]*)\]\([^)]+\)', _replace_img, markdown_text)
        images_used = {}
    else:
        images_used = images_by_block

    from markdown_to_docx import markdown_to_docx
    markdown_to_docx(markdown_text, str(docx_path), images_dict=images_used)

    logger.info("DOCX created: %s", docx_path.name)
    return {
        "total_chars": total_chars,
        "extraction_method": "marker",
        "pdf_type": "scanned" if total_chars > 0 else "text",
    }


def process_file(file_info, out_dir, lang, image_mode, do_translate):
    pdf = Path(file_info["upload_path"])
    result = dict(file_info)
    try:
        if not pdf.exists():
            raise FileNotFoundError(str(pdf))

        docx_name = pdf.stem + "_convertido.docx"
        docx_path = out_dir / docx_name

        logger.info("Processing: %s -> %s", pdf.name, docx_name)
        meta = convert_pdf_to_docx(pdf, docx_path, lang=lang,
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
            "progress":     100,
            "pdf_type":     meta.get("pdf_type", "text"),
            "chars_extracted": meta.get("total_chars", 0),
            "extraction_method": meta.get("extraction_method", "marker"),
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

    loop  = asyncio.get_event_loop()
    files = job["files"]
    total = len(files)
    sem   = asyncio.Semaphore(MAX_CONCURRENT)
    lock  = asyncio.Lock()
    save_counter = 0

    for fi in files:
        fi["status"] = "pending"
        fi["progress"] = 0
    save_job(job)

    async def process_one(idx, fi):
        nonlocal save_counter, job
        async with sem:
            async with lock:
                job["current_file"] = fi["name"]
                job["files"][idx]["status"] = "processing"
                job["files"][idx]["progress"] = 10
                save_job(job)
            try:
                future = loop.run_in_executor(
                    None, process_file, fi, out_dir, lang, image_mode, do_translate)
                updated = await asyncio.wait_for(future, timeout=FILE_TIMEOUT_SECONDS)
            except asyncio.TimeoutError:
                logger.error("Timeout processing %s after %.1fs", fi["name"], FILE_TIMEOUT_SECONDS)
                updated = dict(fi)
                updated.update({
                    "status": "failed",
                    "progress": 100,
                    "error": "Exceeded 3 minutes per file",
                })
        async with lock:
            updated.setdefault("progress", 100)
            job["files"][idx] = updated
            if updated.get("status") == "completed":
                job["processed"] += 1
            else:
                job["failed"] += 1
            job["percentage"] = round(
                (job["processed"] + job["failed"]) / total * 100, 1)
            save_counter += 1
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
    for lib, label in [("marker", "Marker"), ("deep_translator", "deep-translator"),
                        ("docx", "python-docx")]:
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
        "engine": "Marker AI",
    }


@app.post("/api/upload")
async def upload_files(files: List[UploadFile] = File(...)):
    job_id = str(uuid.uuid4())
    up_dir = UPLOAD_DIR / job_id
    up_dir.mkdir(parents=True, exist_ok=True)
    pdfs   = []

    for upload in files:
        filename  = upload.filename or ("file_" + str(uuid.uuid4()) + ".pdf")
        safe_name = sanitize_filename(filename)
        dest    = up_dir / safe_name
        content = await upload.read()
        if len(content) > MAX_UPLOAD_BYTES:
            raise HTTPException(413, safe_name + " is too large")
        dest.write_bytes(content)

        if safe_name.lower().endswith(".zip"):
            try:
                with zipfile.ZipFile(str(dest)) as zf:
                    pdf_infos = [
                        info for info in zf.infolist()
                        if info.filename.lower().endswith(".pdf")
                        and not info.filename.startswith("__MACOSX")
                        and not info.is_dir()
                    ]
                    if len(pdf_infos) > MAX_ZIP_MEMBERS:
                        raise HTTPException(413, safe_name + " has too many PDFs")
                    total_uncompressed = sum(info.file_size for info in pdf_infos)
                    if total_uncompressed > MAX_ZIP_TOTAL_BYTES:
                        raise HTTPException(413, safe_name + " is too large after extraction")
                    for info in pdf_infos:
                            ep_name = sanitize_filename(Path(info.filename).name)
                            ep = up_dir / ep_name
                            if ep.exists():
                                ep = up_dir / (str(uuid.uuid4())[:8] + "_" + ep_name)
                            ep.write_bytes(zf.read(info))
                            pdfs.append({"id": str(uuid.uuid4()), "name": ep.name,
                                         "size": ep.stat().st_size,
                                         "upload_path": str(ep), "status": "pending",
                                         "progress": 0})
                dest.unlink()
            except zipfile.BadZipFile:
                raise HTTPException(400, safe_name + " is not a valid ZIP")
        elif safe_name.lower().endswith(".pdf"):
            pdfs.append({"id": str(uuid.uuid4()), "name": safe_name,
                         "size": len(content), "upload_path": str(dest), "status": "pending",
                         "progress": 0})

    if not pdfs:
        raise HTTPException(400, "No PDF found.")

    job = {"job_id": job_id, "status": "pending", "total": len(pdfs),
           "processed": 0, "failed": 0, "percentage": 0.0, "files": pdfs,
           "current_file": "",
           "created_at": datetime.now().isoformat(),
           "target_language": "pt", "image_mode": "placeholder"}
    save_job(job)
    return {"job_id": job_id, "files_count": len(pdfs), "files": pdfs}


@app.post("/api/process")
async def start_processing(req: ProcessRequest, background_tasks: BackgroundTasks):
    job = load_job(req.job_id)
    if job["status"] == "processing":
        raise HTTPException(400, "Already processing")

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
