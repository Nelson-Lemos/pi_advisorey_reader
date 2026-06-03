# -*- coding: utf-8 -*-
"""
Universal PDF Translator Pro v7.0
==================================
Motor 1: Marker local (instalar marker-pdf)
Motor 2: Datalab API (OCR na nuvem — recomendado para PDFs escaneados)

Saída: DOCX formatado via markdown_to_docx.py
Tradução: Google Translate com cache e retry automático
"""

from fastapi import FastAPI, UploadFile, File, BackgroundTasks, HTTPException
from fastapi.staticfiles import StaticFiles
from fastapi.responses import FileResponse, HTMLResponse
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel
from typing import List
import uuid, json, zipfile, shutil, asyncio, re, time, os
from pathlib import Path
from datetime import datetime
import logging

logging.basicConfig(level=logging.INFO, format="%(levelname)s:%(name)s: %(message)s")
logger = logging.getLogger(__name__)

# ── Limites e configuração ────────────────────────────────────────────────────
MAX_CONCURRENT       = 4
FILE_TIMEOUT_SECONDS = 3600  # 1 hora por ficheiro (documentos grandes)
MAX_UPLOAD_BYTES     = 100 * 1024 * 1024   # 100 MB
MAX_ZIP_MEMBERS      = 100
MAX_ZIP_TOTAL_BYTES  = 200 * 1024 * 1024

# ── Datalab API ───────────────────────────────────────────────────────────────
DATALAB_API_KEY  = os.environ.get("DATALAB_API_KEY", "nyeSwMeyBnFJAvcL0zrp_EJl4LgSVIw6V3ZHMFqniZM")
DATALAB_CONV_URL = "https://www.datalab.to/api/v1/convert"
DATALAB_OCR_URL  = "https://www.datalab.to/api/v1/ocr"

# ── Cache de traduções ────────────────────────────────────────────────────────
_translation_cache: dict = {}

# ── FastAPI ───────────────────────────────────────────────────────────────────
app = FastAPI(title="Universal PDF Translator Pro", version="7.0.0")
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

MIME_MAP = {
    ".pdf":  "application/pdf",
    ".docx": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    ".doc":  "application/msword",
    ".xlsx": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    ".pptx": "application/vnd.openxmlformats-officedocument.presentationml.presentation",
    ".html": "text/html", ".htm": "text/html", ".txt": "text/plain",
    ".png":  "image/png", ".jpg": "image/jpeg", ".jpeg": "image/jpeg",
    ".tiff": "image/tiff", ".bmp": "image/bmp",
}
IMAGE_EXTS = {".png", ".jpg", ".jpeg", ".tiff", ".bmp", ".webp"}


# ── Modelos Pydantic ──────────────────────────────────────────────────────────

class ProcessRequest(BaseModel):
    job_id:          str
    target_language: str  = "pt"
    image_mode:      str  = "placeholder"   # placeholder | keep | remove
    remove_images:   bool = False
    engine:          str  = "datalab"        # datalab | marker
    datalab_mode:    str  = "accurate"       # fast | balanced | 
    datalab_output:  str  = "markdown"       # markdown | html | json


# ── Job helpers ───────────────────────────────────────────────────────────────

def sanitize_filename(filename: str) -> str:
    return "".join(c for c in filename if c.isalnum() or c in "._- ()").strip() or "file.pdf"

def save_job(job):
    (JOBS_DIR / f"{job['job_id']}.json").write_text(
        json.dumps(job, ensure_ascii=False, indent=2), encoding="utf-8")

def load_job(job_id: str):
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


# ── Tradução com cache (NLLB-200 local) ─────────────────────────────────────────

def _translate_chunk(text: str, translator) -> str:
    """Traduz um chunk com cache. Max 4000 chars por chamada (limite NLLB)."""
    if not text or not text.strip():
        return text

    MAX_CHARS = 4000

    # Split if too long
    if len(text) > MAX_CHARS:
        sentences = re.split(r"(?<=[.!?])\s+", text)
        parts, buf = [], ""
        for s in sentences:
            if len(buf) + len(s) + 1 <= MAX_CHARS:
                buf = (buf + " " + s).strip()
            else:
                if buf:
                    parts.append(_translate_chunk(buf, translator))
                buf = s
        if buf:
            parts.append(_translate_chunk(buf, translator))
        return " ".join(parts)

    cached = _translation_cache.get(text)
    if cached:
        return cached
    try:
        r = translator.translate(text)
        result = r.strip() if r else text
        _translation_cache[text] = result
        return result
    except Exception as e:
        logger.error("Translation error: %s — keeping original", e)
        return text


def translate_markdown(md_text: str, translator) -> str:
    """
    Traduz markdown preservando estrutura.
    Para documentos grandes (>50k chars), divide em páginas para evitar
    rate limit do Google Translate.
    """
    # Para documentos muito grandes, dividir em secções por página markdown
    PAGE_SEP = "\n\n---\n\n"
    CHUNK_LIMIT = 60000  # split point for large docs

    if len(md_text) > CHUNK_LIMIT:
        # Dividir em blocos de ~40k chars respeitando parágrafos
        blocks = []
        current = []
        current_len = 0
        for para in md_text.split("\n\n"):
            if current_len + len(para) > CHUNK_LIMIT and current:
                blocks.append("\n\n".join(current))
                current = [para]
                current_len = len(para)
            else:
                current.append(para)
                current_len += len(para)
        if current:
            blocks.append("\n\n".join(current))

        total_blocks = len(blocks)
        logger.info("Translating large doc: %d blocks, %d chars total", total_blocks, len(md_text))
        translated_blocks = []
        for i, block in enumerate(blocks):
            logger.info("Translating block %d/%d (%d chars)...", i+1, total_blocks, len(block))
            try:
                translated_blocks.append(_translate_markdown_block(block, translator))
            except Exception as e:
                logger.error("Block %d/%d failed: %s — keeping original", i+1, total_blocks, e)
                translated_blocks.append(block)
        return "\n\n".join(translated_blocks)

    return _translate_markdown_block(md_text, translator)


def _translate_markdown_block(md_text: str, translator) -> str:
    """Traduz um bloco de markdown preservando estrutura."""
    lines = md_text.split("\n")
    out, batch, code_block = [], [], False

    def flush():
        if not batch:
            return
        chunk = "\n".join(batch)
        out.extend(_translate_chunk(chunk, translator).split("\n"))
        batch.clear()

    for line in lines:
        s = line.strip()
        if s.startswith("```"):
            flush(); out.append(line); code_block = not code_block; continue
        if code_block:
            out.append(line); continue
        if not s:
            flush(); out.append(line); continue
        if s in ("---", "***", "___"):
            flush(); out.append(line); continue
        if re.match(r"^!\[", s):
            flush(); out.append(line); continue
        if "|" in s and re.match(r"^[\s|:\-]+$", s):
            flush(); out.append(line); continue
        m = re.match(r"^(#{1,6}\s+)(.*)", line)
        if m:
            flush()
            out.append(m.group(1) + _translate_chunk(m.group(2), translator))
            continue
        batch.append(line)
        if len("\n".join(batch)) >= 1200:
            flush()

    flush()
    return "\n".join(out)


def translate_html(html_text: str, translator) -> str:
    try:
        from bs4 import BeautifulSoup
        soup = BeautifulSoup(html_text, "html.parser")
        for tag in soup.find_all(string=True):
            if tag.parent.name in ("script", "style", "code", "pre"):
                continue
            if tag.strip():
                tag.replace_with(_translate_chunk(str(tag), translator))
        return str(soup)
    except Exception as e:
        logger.warning("HTML translation failed: %s", e)
        return html_text


def translate_json_obj(obj, translator):
    if isinstance(obj, dict):
        return {k: translate_json_obj(v, translator) for k, v in obj.items()}
    if isinstance(obj, list):
        return [translate_json_obj(i, translator) for i in obj]
    if isinstance(obj, str) and obj.strip() and len(obj) > 3:
        return _translate_chunk(obj, translator)
    return obj


def get_translator(lang: str):
    if not lang or lang in ("none", "original", ""):
        return None
    try:
        from nllb_translator import create_translator
        return create_translator(lang)
    except Exception as e:
        logger.error("Falha ao criar tradutor NLLB: %s", e)
        return None


# ── Polling Datalab ───────────────────────────────────────────────────────────

def _datalab_poll(check_url: str, headers: dict, label: str, max_minutes=12) -> dict:
    import requests
    for attempt in range(max_minutes * 15):
        time.sleep(4)
        r = requests.get(check_url, headers=headers, timeout=30)
        r.raise_for_status()
        result = r.json()
        status = result.get("status", "")
        if attempt % 5 == 0:
            logger.info("Datalab %s: %s (attempt %d)", label, status, attempt + 1)
        if status == "complete":
            return result
        if status == "failed":
            raise RuntimeError(f"Datalab {label} failed: {result.get('error', '')}")
    raise TimeoutError(f"Datalab {label}: timeout after {max_minutes} min")


# ── Cache de classificação de imagens ──────────────────────────────────────────
_classification_cache: dict = {}

def classify_image(img_bytes: bytes, img_filename: str, headers: dict,
                   block_metadata: dict = None) -> str:
    """
    Classifica uma imagem como carimbo, assinatura, logo_marca ou normal.
    Usa new_block_types (se disponível) ou OCR Datalab como fallback.
    """
    cache_key = f"{img_filename}:{len(img_bytes)}"
    if cache_key in _classification_cache:
        return _classification_cache[cache_key]

    if block_metadata and isinstance(block_metadata, dict):
        for key in (img_filename, img_filename.split("/")[-1],
                     img_filename.split("\\")[-1]):
            bt = block_metadata.get(key) or block_metadata.get("all", {}).get(key)
            if bt:
                bt_lower = bt.lower()
                if bt_lower in ("stamp", "carimbo", "rubber_stamp"):
                    _classification_cache[cache_key] = "carimbo"
                    return "carimbo"
                if bt_lower in ("logo", "logotipo", "logo_marca", "watermark"):
                    _classification_cache[cache_key] = "logo_marca"
                    return "logo_marca"
                if bt_lower in ("signature", "assinatura", "rubrica", "handwritten"):
                    _classification_cache[cache_key] = "assinatura"
                    return "assinatura"
                if bt_lower in ("figure", "picture", "image", "photo", "illustration",
                                "diagram", "chart", "graphic"):
                    _classification_cache[cache_key] = "normal"
                    return "normal"
                break

    # Fallback: OCR via endpoint Datalab
    try:
        import io, requests
        files = {"file": (img_filename, io.BytesIO(img_bytes), "image/png")}
        resp = requests.post(
            DATALAB_OCR_URL,
            files=files,
            data={"langs": "auto"},
            headers=headers,
            timeout=(30, 120),
        )
        resp.raise_for_status()
        data = resp.json()
        text_parts = []
        for page in data.get("pages", []):
            for block in page.get("text_lines", []):
                t = block.get("text", "").strip()
                if t:
                    text_parts.append(t)
        ocr_text = " ".join(text_parts).lower()

        if re.search(r"\d{2}[/\-\.]\d{2}[/\-\.]\d{2,4}", ocr_text) or \
           any(w in ocr_text for w in ["carimbo", "selo", "certifico", "certificamos",
                "republica", "ministerio", "governo", "oficial", "autentic",
                "conferido", "valido", "protocolo"]):
            _classification_cache[cache_key] = "carimbo"
            return "carimbo"

        if any(w in ocr_text for w in ["assinatura", "assinado", "assinante",
                "firma", "rubrica", "nome"]):
            _classification_cache[cache_key] = "assinatura"
            return "assinatura"

        if any(w in ocr_text for w in ["logo", "logotipo", "marca", "corporation",
                "inc", "ltd", "ltda", "company", "enterprise", "group"]):
            _classification_cache[cache_key] = "logo_marca"
            return "logo_marca"

        _classification_cache[cache_key] = "normal"
        return "normal"

    except Exception as e:
        logger.warning("Falha ao classificar %s via OCR: %s — mantendo como normal",
                       img_filename, e)
        _classification_cache[cache_key] = "normal"
        return "normal"


# ── Motor 1: Datalab API ──────────────────────────────────────────────────────

def convert_with_datalab(src_path: Path, out_path: Path,
                          lang="pt", do_translate=True,
                          output_format="markdown", mode="balanced",
                          image_mode="placeholder") -> dict:
    """
    Converte qualquer documento via Datalab API:
    - Imagens/PDFs escaneados → OCR endpoint (Chandra/Surya)
    - PDFs digitais/Office    → Marker endpoint (force_ocr=true)
    Depois converte markdown → DOCX via markdown_to_docx.py
    """
    import requests, base64

    key = DATALAB_API_KEY.strip()
    if not key:
        raise ValueError("DATALAB_API_KEY not set")

    headers = {"X-API-Key": key}
    ext     = src_path.suffix.lower()
    mime    = MIME_MAP.get(ext, "application/octet-stream")
    is_img  = ext in IMAGE_EXTS

    session = requests.Session()

    def _post_file(url, files, data=None):
        return session.post(
            url,
            files=files,
            data=data,
            headers=headers,
            timeout=(60, 3600),
        )

    # ── Submissão com retry automático ───────────────────────────────────────
    def _submit(attempt=0):
        with open(str(src_path), "rb") as f:
            if is_img:
                return _post_file(
                    DATALAB_OCR_URL,
                    files={"file": (src_path.name, f, mime)},
                    data={"langs": "auto"},
                ), "OCR"
            extra_data = {}
            if image_mode == "keep":
                extra_data["extras"] = "new_block_types"
            return _post_file(
                DATALAB_CONV_URL,
                files={"file": (src_path.name, f, mime)},
                data={
                    "output_format":           output_format,
                    "mode":                    mode,
                    "force_ocr":               "true",
                    "disable_image_extraction": "false",
                    **extra_data,
                },
            ), "Marker"

    resp, label = None, "Marker"
    max_attempts = 5
    for attempt in range(max_attempts):
        try:
            if is_img:
                logger.info("Datalab OCR: %s (attempt %d)", src_path.name, attempt + 1)
            else:
                logger.info("Datalab Marker: %s (mode=%s, format=%s, force_ocr=true, attempt=%d)",
                            src_path.name, mode, output_format, attempt + 1)
            resp, label = _submit(attempt)
            resp.raise_for_status()
            break
        except Exception as e:
            wait = 10 * (attempt + 1)
            logger.warning("Upload attempt %d failed: %s — waiting %ds", attempt + 1, e, wait)
            if attempt == max_attempts - 1:
                raise
            time.sleep(wait)

    data = resp.json()
    if not data.get("success"):
        raise RuntimeError(f"Datalab {label} submit failed: {data.get('error', data)}")
    logger.info("Datalab %s: submitted (id=%s)", label, data.get("request_id", "?"))

    # ── Polling ───────────────────────────────────────────────────────────────
    result = _datalab_poll(data["request_check_url"], headers, label)
    pages  = result.get("page_count", "?")
    score  = result.get("parse_quality_score", "?")
    logger.info("Datalab %s: complete — %s pages, quality=%s", label, pages, score)

    # ── Extrair metadados de blocos (new_block_types) ─────────────────────────
    block_metadata = {}
    if image_mode == "keep":
        for candidate in ("block_types", "blocks", "metadata", "block_metadata"):
            raw = result.get(candidate, {})
            if raw:
                if isinstance(raw, dict):
                    block_metadata = raw
                elif isinstance(raw, str):
                    try:
                        import json as _json
                        block_metadata = _json.loads(raw)
                    except Exception:
                        pass
                break
        if not block_metadata and "json" in result:
            json_data = result.get("json", {})
            if isinstance(json_data, dict):
                block_metadata = json_data.get("block_types", {}) or \
                                 json_data.get("blocks", {}) or block_metadata

    # ── Extracção de texto ────────────────────────────────────────────────────
    if is_img:
        lines = []
        for page in result.get("pages", []):
            for block in page.get("text_lines", []):
                t = block.get("text", "").strip()
                if t:
                    lines.append(t)
            lines.append("")
        text_content  = "\n".join(lines)
        output_format = "markdown"
    elif output_format == "json":
        import json as _json
        text_content = _json.dumps(result.get("json", {}), ensure_ascii=False, indent=2)
    else:
        text_content = result.get(output_format, "") or result.get("markdown", "")

    if not text_content or not text_content.strip():
        raise RuntimeError("Datalab returned empty content — document may be empty or protected")

    total_chars = len(text_content)
    logger.info("Datalab: extracted %d chars", total_chars)

    # ── Tratamento de imagens no markdown ─────────────────────────────────────
    images = result.get("images", {})

    # Construir dicionário de bytes de imagem antecipadamente para classificação
    img_dict_bytes = {}
    if images:
        for img_name, img_b64 in images.items():
            try:
                img_dict_bytes[img_name] = base64.b64decode(img_b64)
            except Exception:
                pass

    if output_format == "markdown" and not is_img:
        if image_mode == "remove":
            text_content = re.sub(r"!\[[^\]]*\]\([^)]+\)", "", text_content)
        elif image_mode == "placeholder":
            text_content = re.sub(
                r"!\[([^\]]*)\]\([^)]+\)",
                lambda m: f"[{m.group(1) or 'IMAGEM'}]",
                text_content
            )
        elif image_mode == "keep" and img_dict_bytes:
            # Classificar cada imagem e substituir carimbos/assinaturas/logos
            classifications = {}
            for img_key, img_bytes_val in img_dict_bytes.items():
                classifications[img_key] = classify_image(
                    img_bytes_val, img_key, headers, block_metadata
                )
            normal_img_keys = {k for k, v in classifications.items() if v == "normal"}

            # Substituir no markdown imagens classificadas como especiais
            def _replace_classified(match):
                alt = match.group(1)
                img_path = match.group(2)
                img_key = next(
                    (k for k in classifications
                     if k.endswith(img_path) or img_path in k),
                    None
                )
                if img_key:
                    cls = classifications.get(img_key, "normal")
                    if cls == "carimbo":
                        return "[carimbo]"
                    if cls == "assinatura":
                        return "[assinatura]"
                    if cls == "logo_marca":
                        return "[logo marca]"
                return match.group(0)

            text_content = re.sub(
                r"!\[([^\]]*)\]\(([^)]+)\)",
                _replace_classified,
                text_content
            )
            logger.info(
                "Datalab: %d/%d imagens classificadas como normais",
                len(normal_img_keys), len(img_dict_bytes)
            )

            # Filtrar img_dict_bytes para conter apenas imagens normais
            img_dict_bytes = {k: v for k, v in img_dict_bytes.items()
                              if k in normal_img_keys}

    # ── Tradução ──────────────────────────────────────────────────────────────
    if do_translate and lang not in ("", "none", "original"):
        translator = get_translator(lang)
        if translator:
            logger.info("Datalab: translating to %s...", lang)
            if output_format == "html":
                text_content = translate_html(text_content, translator)
            elif output_format == "json":
                import json as _json
                text_content = _json.dumps(
                    translate_json_obj(result.get("json", {}), translator),
                    ensure_ascii=False, indent=2)
            else:
                text_content = translate_markdown(text_content, translator)
            logger.info("Datalab: translation complete")

    # ── Guardar imagens normais em disco (apenas modo keep) ───────────────────
    if image_mode == "keep" and img_dict_bytes:
        img_dir = out_path.parent / (out_path.stem + "_images")
        img_dir.mkdir(exist_ok=True)
        for img_name, img_bytes in img_dict_bytes.items():
            (img_dir / img_name).write_bytes(img_bytes)
        logger.info("Datalab: saved %d normal images to disk", len(img_dict_bytes))

    # ── Converter markdown → DOCX ─────────────────────────────────────────────
    saved_as_docx = False
    if output_format == "markdown":
        try:
            from markdown_to_docx import markdown_to_docx
            docx_path = out_path.with_suffix(".docx")
            markdown_to_docx(
                text_content,
                str(docx_path),
                images_dict=img_dict_bytes if image_mode == "keep" else None
            )
            out_path      = docx_path
            saved_as_docx = True
            logger.info("Datalab: markdown → DOCX OK")
        except ImportError:
            logger.warning("markdown_to_docx.py not found — saving as .md")
        except Exception as e:
            logger.warning("markdown_to_docx failed (%s) — saving as .md", e)

    if not saved_as_docx:
        ext_map  = {"markdown": ".md", "html": ".html", "json": ".json"}
        out_path = out_path.with_suffix(ext_map.get(output_format, ".md"))
        out_path.write_text(text_content, encoding="utf-8")

    size_kb = round(out_path.stat().st_size / 1024, 1)
    logger.info("Datalab: saved → %s (%.1f KB)", out_path.name, size_kb)

    return {
        "total_chars":       total_chars,
        "extraction_method": f"datalab_{label.lower()}",
        "pdf_type":          "scanned",
        "page_count":        pages,
        "quality_score":     score,
        "output_path":       str(out_path),
        "output_name":       out_path.name,
        "size_kb":           size_kb,
    }


# ── Motor 2: Marker local ─────────────────────────────────────────────────────

def convert_with_marker(src_path: Path, out_path: Path,
                         lang="pt", do_translate=True,
                         image_mode="placeholder") -> dict:
    """
    Converte PDF usando Marker instalado localmente.
    Requer: pip install marker-pdf
    """
    logger.info("Marker local: converting %s", src_path.name)

    from marker.converters.pdf import PdfConverter
    from marker.models import create_model_dict
    from marker.output import text_from_rendered

    converter = PdfConverter(artifact_dict=create_model_dict())
    rendered  = converter(str(src_path))
    markdown_text, metadata, images_by_block = text_from_rendered(rendered)

    total_chars = len(markdown_text)
    logger.info("Marker: extracted %d chars", total_chars)

    # Tratamento de imagens
    if image_mode == "remove":
        markdown_text  = re.sub(r"!\[[^\]]*\]\([^)]+\)", "", markdown_text)
        images_by_block = {}
    elif image_mode == "placeholder":
        markdown_text  = re.sub(
            r"!\[([^\]]*)\]\([^)]+\)",
            lambda m: f"[{m.group(1) or 'IMAGEM'}]",
            markdown_text
        )
        images_by_block = {}

    # Tradução
    if do_translate and lang not in ("", "none", "original"):
        translator = get_translator(lang)
        if translator:
            logger.info("Marker: translating to %s...", lang)
            markdown_text = translate_markdown(markdown_text, translator)
            logger.info("Marker: translation complete")

    # Converter markdown → DOCX
    try:
        from markdown_to_docx import markdown_to_docx
        docx_path = out_path.with_suffix(".docx")
        markdown_to_docx(markdown_text, str(docx_path), images_dict=images_by_block)
        out_path = docx_path
        logger.info("Marker: markdown → DOCX OK")
    except ImportError:
        logger.warning("markdown_to_docx.py not found — saving as .md")
        out_path = out_path.with_suffix(".md")
        out_path.write_text(markdown_text, encoding="utf-8")
    except Exception as e:
        logger.warning("markdown_to_docx failed (%s) — saving as .md", e)
        out_path = out_path.with_suffix(".md")
        out_path.write_text(markdown_text, encoding="utf-8")

    size_kb = round(out_path.stat().st_size / 1024, 1)
    logger.info("Marker: saved → %s (%.1f KB)", out_path.name, size_kb)

    return {
        "total_chars":       total_chars,
        "extraction_method": "marker_local",
        "pdf_type":          "scanned" if total_chars > 0 else "text",
        "output_path":       str(out_path),
        "output_name":       out_path.name,
        "size_kb":           size_kb,
    }


# ── Processador de ficheiros ──────────────────────────────────────────────────

def process_file(file_info: dict, out_dir: Path, lang: str,
                 image_mode: str, do_translate: bool,
                 engine: str, datalab_mode: str, datalab_output: str) -> dict:

    src    = Path(file_info["upload_path"])
    result = dict(file_info)

    try:
        if not src.exists():
            raise FileNotFoundError(str(src))

        out_stem = src.stem + "_convertido"
        out_path = out_dir / out_stem  # extensão definida pelo motor

        logger.info("Processing [%s]: %s", engine.upper(), src.name)

        if engine == "datalab":
            meta = convert_with_datalab(
                src, out_path, lang=lang,
                do_translate=do_translate,
                output_format=datalab_output,
                mode=datalab_mode,
                image_mode=image_mode,
            )
        elif engine == "marker":
            meta = convert_with_marker(
                src, out_path, lang=lang,
                do_translate=do_translate,
                image_mode=image_mode,
            )
        else:
            raise ValueError(f"Unknown engine: {engine}")

        final_path = Path(meta["output_path"])
        if not final_path.exists():
            raise RuntimeError(f"Output not created: {final_path}")

        logger.info("Done: %s (%.1f KB)", meta["output_name"], meta["size_kb"])
        result.update({
            "output_path":       meta["output_path"],
            "output_name":       meta["output_name"],
            "size_kb":           meta["size_kb"],
            "status":            "completed",
            "progress":          100,
            "pdf_type":          meta.get("pdf_type", "unknown"),
            "chars_extracted":   meta.get("total_chars", 0),
            "extraction_method": meta.get("extraction_method", engine),
            "page_count":        meta.get("page_count", "?"),
            "quality_score":     meta.get("quality_score", "?"),
        })

    except Exception as e:
        logger.error("Error [%s] %s: %s", engine, file_info["name"], e)
        result["error"]    = str(e)
        result["status"]   = "failed"
        result["progress"] = 100

    return result


# ── Job runner com concorrência ───────────────────────────────────────────────

async def run_job(job_id: str, lang: str, image_mode: str, do_translate: bool,
                  engine: str, datalab_mode: str, datalab_output: str):

    job     = load_job(job_id)
    out_dir = OUTPUT_DIR / job_id
    out_dir.mkdir(parents=True, exist_ok=True)

    job["status"]     = "processing"
    job["started_at"] = datetime.now().isoformat()
    for fi in job["files"]:
        fi["status"] = "pending"; fi["progress"] = 0
    save_job(job)

    loop  = asyncio.get_event_loop()
    total = len(job["files"])
    sem   = asyncio.Semaphore(MAX_CONCURRENT)
    lock  = asyncio.Lock()

    async def process_one(idx, fi):
        async with sem:
            async with lock:
                job["current_file"]            = fi["name"]
                job["files"][idx]["status"]    = "processing"
                job["files"][idx]["progress"]  = 10
                save_job(job)
            try:
                import functools
                fn = functools.partial(
                    process_file, fi, out_dir, lang, image_mode, do_translate,
                    engine, datalab_mode, datalab_output
                )
                updated = await asyncio.wait_for(
                    loop.run_in_executor(None, fn),
                    timeout=FILE_TIMEOUT_SECONDS
                )
            except asyncio.TimeoutError:
                logger.error("Timeout: %s after %ds", fi["name"], FILE_TIMEOUT_SECONDS)
                updated = dict(fi)
                updated.update({"status": "failed", "progress": 100,
                                 "error": f"Timeout após {FILE_TIMEOUT_SECONDS}s"})
        async with lock:
            job["files"][idx] = updated
            if updated.get("status") == "completed":
                job["processed"] += 1
            else:
                job["failed"] += 1
            job["percentage"] = round(
                (job["processed"] + job["failed"]) / total * 100, 1)
            save_job(job)

    try:
        await asyncio.gather(*[process_one(i, f) for i, f in enumerate(job["files"])])
    except Exception as e:
        logger.error("run_job gather error: %s", e)

    # Criar ZIP
    zip_path = OUTPUT_DIR / (job_id + "_all.zip")
    added = 0
    with zipfile.ZipFile(str(zip_path), "w", zipfile.ZIP_DEFLATED) as zf:
        for fi in job["files"]:
            p = fi.get("output_path", "")
            if p and Path(p).exists():
                zf.write(p, fi["output_name"])
                added += 1
    logger.info("ZIP: %d files (%.1f KB)", added,
                zip_path.stat().st_size / 1024 if zip_path.exists() else 0)

    job["status"]       = "completed"
    job["zip_path"]     = str(zip_path)
    job["output_dir"]   = str(out_dir)
    job["completed_at"] = datetime.now().isoformat()
    save_job(job)


# ── Rotas ─────────────────────────────────────────────────────────────────────

@app.get("/", response_class=HTMLResponse)
async def root():
    index = STATIC_DIR / "index.html"
    if not index.exists():
        return HTMLResponse(f"<h2>index.html not found</h2><p>{STATIC_DIR}</p>", status_code=404)
    return HTMLResponse(content=index.read_text(encoding="utf-8"))


@app.get("/api/health")
async def health():
    deps = {}
    checks = [
        ("marker",         "Marker local"),
        ("transformers",   "transformers (NLLB)"),
        ("docx",           "python-docx"),
        ("fitz",           "PyMuPDF"),
        ("pdf2docx",       "pdf2docx"),
        ("bs4",            "beautifulsoup4"),
        ("openpyxl",       "openpyxl"),
        ("pptx",           "python-pptx"),
    ]
    for lib, label in checks:
        try:
            __import__(lib)
            deps[label] = "ok"
        except ImportError:
            deps[label] = "not installed"

    # Verificar markdown_to_docx.py
    deps["markdown_to_docx"] = "ok" if (BASE_DIR / "markdown_to_docx.py").exists() else "not found"
    # Verificar chave Datalab
    deps["Datalab API key"]  = "ok" if DATALAB_API_KEY else "not set"

    return {
        "status":       "ok",
        "version":      "7.0.0",
        "base_dir":     str(BASE_DIR),
        "dependencies": deps,
        "engines":      ["datalab", "marker"],
        "datalab_key":  "configured" if DATALAB_API_KEY else "missing",
    }


@app.post("/api/upload")
async def upload_files(files: List[UploadFile] = File(...)):
    job_id = str(uuid.uuid4())
    up_dir = UPLOAD_DIR / job_id
    up_dir.mkdir(parents=True, exist_ok=True)
    accepted = []

    for upload in files:
        filename  = upload.filename or "file"
        safe_name = sanitize_filename(filename)
        dest      = up_dir / safe_name
        content   = await upload.read()

        if len(content) > MAX_UPLOAD_BYTES:
            raise HTTPException(413, f"{safe_name} é demasiado grande (max 100MB)")
        dest.write_bytes(content)
        ext = dest.suffix.lower()

        if ext == ".zip":
            try:
                with zipfile.ZipFile(str(dest)) as zf:
                    pdf_infos = [
                        i for i in zf.infolist()
                        if i.filename.lower().endswith(".pdf")
                        and not i.filename.startswith("__MACOSX")
                        and not i.is_dir()
                    ]
                    if len(pdf_infos) > MAX_ZIP_MEMBERS:
                        raise HTTPException(413, f"{safe_name} tem demasiados PDFs")
                    if sum(i.file_size for i in pdf_infos) > MAX_ZIP_TOTAL_BYTES:
                        raise HTTPException(413, f"{safe_name} é demasiado grande após extracção")
                    for info in pdf_infos:
                        ep_name = sanitize_filename(Path(info.filename).name)
                        ep = up_dir / ep_name
                        if ep.exists():
                            ep = up_dir / (str(uuid.uuid4())[:8] + "_" + ep_name)
                        ep.write_bytes(zf.read(info))
                        accepted.append({
                            "id": str(uuid.uuid4()), "name": ep.name,
                            "size": ep.stat().st_size, "type": "pdf",
                            "upload_path": str(ep), "status": "pending", "progress": 0
                        })
                dest.unlink()
            except zipfile.BadZipFile:
                raise HTTPException(400, f"{safe_name} não é um ZIP válido")
        elif ext == ".pdf":
            accepted.append({
                "id": str(uuid.uuid4()), "name": safe_name,
                "size": len(content), "type": "pdf",
                "upload_path": str(dest), "status": "pending", "progress": 0
            })

    if not accepted:
        raise HTTPException(400, "Nenhum PDF encontrado. Envie ficheiros .pdf ou .zip com PDFs.")

    job = {
        "job_id": job_id, "status": "pending",
        "total": len(accepted), "processed": 0, "failed": 0,
        "percentage": 0.0, "current_file": "",
        "files": accepted,
        "created_at": datetime.now().isoformat(),
        "target_language": "pt", "image_mode": "placeholder",
        "engine": "datalab",
    }
    save_job(job)
    return {"job_id": job_id, "files_count": len(accepted), "files": accepted}


@app.post("/api/process")
async def start_processing(req: ProcessRequest, background_tasks: BackgroundTasks):
    job = load_job(req.job_id)
    if job["status"] == "processing":
        raise HTTPException(400, "Já em processamento")

    do_translate = req.target_language not in ("none", "original", "")
    image_mode   = "remove" if req.remove_images else req.image_mode

    job.update({
        "target_language": req.target_language,
        "image_mode":      image_mode,
        "engine":          req.engine,
        "datalab_mode":    req.datalab_mode,
        "datalab_output":  req.datalab_output,
        "status":          "processing",
        "processed":       0,
        "failed":          0,
        "percentage":      0.0,
    })
    for f in job["files"]:
        f["status"] = "pending"
    save_job(job)

    background_tasks.add_task(
        run_job, req.job_id, req.target_language,
        image_mode, do_translate,
        req.engine, req.datalab_mode, req.datalab_output
    )
    return {"status": "started", "job_id": req.job_id}



@app.get("/api/languages")
async def get_languages():
    return {
        "af":"Afrikaans","sq":"Albanês","am":"Amárico","ar":"Árabe","hy":"Arménio",
        "az":"Azerbaijanês","eu":"Basco","be":"Bielorusso","bn":"Bengalês","bs":"Bósnio",
        "bg":"Búlgaro","ca":"Catalão","zh-cn":"Chinês Simplificado","zh-tw":"Chinês Tradicional",
        "hr":"Croata","cs":"Checo","da":"Dinamarquês","nl":"Neerlandês","en":"Inglês",
        "et":"Estoniano","tl":"Filipino","fi":"Finlandês","fr":"Francês","gl":"Galego",
        "ka":"Georgiano","de":"Alemão","el":"Grego","gu":"Gujarati","ht":"Haitiano",
        "ha":"Hauça","iw":"Hebraico","hi":"Hindi","hu":"Húngaro","is":"Islandês",
        "ig":"Igbo","id":"Indonésio","ga":"Irlandês","it":"Italiano","ja":"Japonês",
        "jw":"Javanês","kn":"Canarês","kk":"Cazaque","km":"Khmer","ko":"Coreano",
        "ku":"Curdo","ky":"Quirguiz","lo":"Laociano","la":"Latim","lv":"Letão",
        "lt":"Lituano","mk":"Macedônio","ms":"Malaio","ml":"Malaiala","mt":"Maltês",
        "mi":"Maori","mr":"Marata","mn":"Mongol","my":"Birmanês","ne":"Nepalês",
        "no":"Norueguês","ps":"Pashto","fa":"Persa","pl":"Polaco","pt":"Português",
        "pa":"Punjabi","ro":"Romeno","ru":"Russo","sm":"Samoano","sr":"Sérvio",
        "si":"Cingalês","sk":"Eslovaco","sl":"Esloveno","so":"Somali","es":"Espanhol",
        "sw":"Suaíli","sv":"Sueco","tg":"Tajique","ta":"Tâmil","te":"Telugu",
        "th":"Tailandês","tr":"Turco","uk":"Ucraniano","ur":"Urdu","uz":"Uzbeque",
        "vi":"Vietnamita","cy":"Galês","xh":"Xhosa","yi":"Iídiche","yo":"Ioruba","zu":"Zulu"
    }

@app.get("/api/jobs/{job_id}")
async def get_job(job_id: str):
    return load_job(job_id)


@app.get("/api/jobs")
async def get_all_jobs():
    return list_jobs()


@app.get("/api/download/{job_id}/zip")
async def download_zip(job_id: str):
    job = load_job(job_id)
    zp  = Path(job.get("zip_path", ""))
    if not zp.exists():
        raise HTTPException(404, "ZIP não pronto ainda")
    return FileResponse(str(zp), filename=f"traducao_{job_id[:8]}.zip",
                        media_type="application/zip")


@app.get("/api/download/{job_id}/file/{file_id}")
async def download_file(job_id: str, file_id: str):
    job = load_job(job_id)
    fi  = next((f for f in job["files"] if f["id"] == file_id), None)
    if not fi:
        raise HTTPException(404, "Ficheiro não encontrado")
    op = Path(fi.get("output_path", ""))
    if not op.exists():
        raise HTTPException(404, "Ficheiro não pronto")
    ext = op.suffix.lower()
    MIME = {
        ".docx": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        ".pdf":  "application/pdf",
        ".md":   "text/markdown",
        ".html": "text/html",
        ".json": "application/json",
        ".txt":  "text/plain",
    }
    return FileResponse(str(op), filename=fi["output_name"],
                        media_type=MIME.get(ext, "application/octet-stream"))


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