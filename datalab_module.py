# -*- coding: utf-8 -*-
"""
datalab_module.py  —  Integração Datalab/Marker API
====================================================
Adiciona conversão OCR via Datalab ao main.py baseado em Marker local.

COMO USAR:
----------
1. Copiar este ficheiro para a mesma pasta do main.py
2. No main.py, adicionar no topo:
       from datalab_module import convert_with_datalab, datalab_available
3. Na função convert_pdf_to_docx(), adicionar no início:
       if use_datalab:
           return convert_with_datalab(pdf_path, docx_path, lang, do_translate)
4. Adicionar use_datalab=False ao ProcessRequest e passar o parâmetro

A chave API já está configurada abaixo.
"""

import os
import re
import time
import base64
import logging
from pathlib import Path

logger = logging.getLogger(__name__)

# ── Configuração ──────────────────────────────────────────────────────────────
DATALAB_API_KEY  = os.environ.get("DATALAB_API_KEY", "wUPsGG53rk_SqbznrRfo8z_bvDyfv1KXVtlVUcMEmFU")
DATALAB_CONV_URL = "https://www.datalab.to/api/v1/convert"
DATALAB_OCR_URL  = "https://www.datalab.to/api/v1/ocr"

# Cache de traduções — evita traduzir o mesmo texto duas vezes
_translation_cache: dict = {}

MIME_MAP = {
    ".pdf":  "application/pdf",
    ".docx": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    ".doc":  "application/msword",
    ".xlsx": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    ".pptx": "application/vnd.openxmlformats-officedocument.presentationml.presentation",
    ".ppt":  "application/vnd.ms-powerpoint",
    ".html": "text/html",
    ".htm":  "text/html",
    ".txt":  "text/plain",
    ".png":  "image/png",
    ".jpg":  "image/jpeg",
    ".jpeg": "image/jpeg",
    ".tiff": "image/tiff",
    ".bmp":  "image/bmp",
}
IMAGE_EXTS = {".png", ".jpg", ".jpeg", ".tiff", ".bmp", ".webp"}


def datalab_available() -> bool:
    """Verifica se a chave API está configurada."""
    return bool(DATALAB_API_KEY and DATALAB_API_KEY.strip())


# ── Tradução com cache e retry ────────────────────────────────────────────────

def _translate_chunk(text: str, translator, retry: int = 0) -> str:
    """Traduz um chunk de texto com retry automático."""
    if not text or not text.strip():
        return text
    cache_key = text
    if cache_key in _translation_cache:
        return _translation_cache[cache_key]
    try:
        time.sleep(0.3)  # pausa educada entre chamadas
        r = translator.translate(text)
        result = r.strip() if r else text
        _translation_cache[cache_key] = result
        return result
    except Exception as e:
        msg = str(e)
        wait = min(5 * (retry + 1), 30)
        logger.warning("translate error (tentativa %d): %s — aguardando %ds", retry + 1, msg, wait)
        time.sleep(wait)
        if retry < 3:
            return _translate_chunk(text, translator, retry + 1)
        logger.error("Tradução falhou após %d tentativas, mantendo original", retry + 1)
        return text


def _translate_markdown(md_text: str, translator) -> str:
    """
    Traduz markdown preservando:
    - Blocos de código (``` ```)
    - Títulos (# ## ###) — traduz só o texto, mantém #
    - Linhas vazias e separadores
    - Imagens ![...](...) e links [...](...) — mantém intactos
    - Linhas de tabela |---|---| — mantém intactas
    Agrupa linhas normais em batches de ~1200 chars para eficiência.
    """
    lines = md_text.split("\n")
    out = []
    code_block = False
    batch = []

    def flush():
        if not batch:
            return
        chunk = "\n".join(batch)
        translated = _translate_chunk(chunk, translator)
        out.extend(translated.split("\n"))
        batch.clear()

    for line in lines:
        stripped = line.strip()

        # Bloco de código
        if stripped.startswith("```"):
            flush()
            out.append(line)
            code_block = not code_block
            continue
        if code_block:
            out.append(line)
            continue

        # Linha vazia
        if not stripped:
            flush()
            out.append(line)
            continue

        # Separador markdown
        if stripped in ("---", "***", "___"):
            flush()
            out.append(line)
            continue

        # Imagem markdown — não traduzir
        if re.match(r"^!\[", stripped):
            flush()
            out.append(line)
            continue

        # Linha de tabela |---|---|
        if "|" in stripped and re.match(r"^[\s|:\-]+$", stripped):
            flush()
            out.append(line)
            continue

        # Título — traduz só o conteúdo, mantém os #
        heading = re.match(r"^(#{1,6}\s+)(.*)", line)
        if heading:
            flush()
            prefix, text = heading.groups()
            out.append(prefix + _translate_chunk(text, translator))
            continue

        # Linha normal — acumular em batch
        batch.append(line)
        if len("\n".join(batch)) >= 1200:
            flush()

    flush()
    return "\n".join(out)


def _translate_html(html_text: str, translator) -> str:
    """Traduz nós de texto em HTML, preservando tags e atributos."""
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


def _translate_json_obj(obj, translator):
    """Traduz recursivamente os valores de texto num objecto JSON."""
    if isinstance(obj, dict):
        return {k: _translate_json_obj(v, translator) for k, v in obj.items()}
    if isinstance(obj, list):
        return [_translate_json_obj(i, translator) for i in obj]
    if isinstance(obj, str) and obj.strip() and len(obj) > 3:
        return _translate_chunk(obj, translator)
    return obj


def _get_translator(lang: str):
    """Cria e devolve um GoogleTranslator para o idioma dado."""
    if not lang or lang in ("none", "original", ""):
        return None
    try:
        from deep_translator import GoogleTranslator
        t = GoogleTranslator(source="auto", target=lang)
        logger.info("Tradutor pronto: auto → %s", lang)
        return t
    except ImportError:
        logger.error("deep_translator não instalado — pip install deep-translator")
        return None
    except Exception as e:
        logger.error("Erro ao criar tradutor: %s", e)
        return None


# ── Polling helper ────────────────────────────────────────────────────────────

def _poll(check_url: str, headers: dict, label: str, max_minutes: int = 12) -> dict:
    """Faz polling até o job estar completo. Lança TimeoutError se exceder o tempo."""
    import requests
    max_attempts = max_minutes * 15  # ~4s por tentativa
    for attempt in range(max_attempts):
        time.sleep(4)
        resp = requests.get(check_url, headers=headers, timeout=30)
        resp.raise_for_status()
        result = resp.json()
        status = result.get("status", "")
        if attempt % 5 == 0:
            logger.info("Datalab %s: %s (tentativa %d/%d)", label, status, attempt + 1, max_attempts)
        if status == "complete":
            return result
        if status == "failed":
            raise RuntimeError(f"Datalab {label} falhou: {result.get('error', '')}")
    raise TimeoutError(f"Datalab {label}: timeout após {max_minutes} minutos")


# ── Conversor principal ───────────────────────────────────────────────────────

def convert_with_datalab(
    src_path: Path,
    out_path: Path,
    lang: str = "pt",
    do_translate: bool = True,
    output_format: str = "markdown",   # markdown | html | json
    mode: str = "balanced",            # fast | balanced | accurate
    api_key: str = "",
    image_mode: str = "placeholder",   # placeholder | keep | remove
) -> dict:
    """
    Converte qualquer documento usando a API Datalab/Marker.

    Fluxo:
      1. Envio do ficheiro para Datalab
      2. Polling até conversão completa
      3. Extracção do texto (markdown / html / json)
      4. Tradução com cache (se pedido)
      5. Gravação do ficheiro de saída

    Devolve dict com metadados (chars, páginas, qualidade, etc.)
    compatível com o formato esperado pelo main.py do Marker local.
    """
    import requests
    key = (api_key or DATALAB_API_KEY).strip()
    if not key:
        raise ValueError(
            "Chave API Datalab não configurada. "
            "Defina DATALAB_API_KEY ou passe api_key..."
        )

    headers  = {"X-API-Key": key}
    ext      = src_path.suffix.lower()
    mime     = MIME_MAP.get(ext, "application/octet-stream")
    is_image = ext in IMAGE_EXTS

    session = requests.Session()

    def _post_file(url, files, data=None):
        return session.post(
            url,
            files=files,
            data=data,
            headers=headers,
            timeout=(60, 3600),
        )

    # ── 1. Submissão ──────────────────────────────────────────────────────────
    max_attempts = 5
    for attempt in range(max_attempts):
        try:
            if is_image:
                logger.info("Datalab OCR: a enviar imagem %s (attempt %d)", src_path.name, attempt + 1)
                with open(str(src_path), "rb") as f:
                    resp = _post_file(
                        DATALAB_OCR_URL,
                        files={"file": (src_path.name, f, mime)},
                        data={"langs": "auto"},
                    )
                label = "OCR"
            else:
                logger.info("Datalab Marker: a enviar %s (mode=%s, format=%s, force_ocr=true, attempt=%d)",
                            src_path.name, mode, output_format, attempt + 1)
                with open(str(src_path), "rb") as f:
                    resp = _post_file(
                        DATALAB_CONV_URL,
                        files={"file": (src_path.name, f, mime)},
                        data={
                            "output_format": output_format,
                            "mode": mode,
                            "force_ocr": "true",          # <-- força OCR em PDFs escaneados
                            "disable_image_extraction": "false",
                        },
                    )
                label = "Marker"

            resp.raise_for_status()
            break
        except Exception as e:
            wait = 10 * (attempt + 1)
            logger.warning("Datalab submission attempt %d failed: %s — waiting %ds", attempt + 1, e, wait)
            if attempt == max_attempts - 1:
                raise
            time.sleep(wait)

    data = resp.json()
    if not data.get("success"):
        raise RuntimeError(f"Datalab {label} submissão falhou: {data.get('error', data)}")

    logger.info("Datalab %s: submetido (request_id=%s)", label, data.get("request_id", "?"))

    # ── 2. Polling ────────────────────────────────────────────────────────────
    result = _poll(data["request_check_url"], headers, label)

    pages   = result.get("page_count", "?")
    score   = result.get("parse_quality_score", "?")
    logger.info("Datalab %s: completo — %s páginas, qualidade=%s", label, pages, score)

    # ── 3. Extracção de texto ─────────────────────────────────────────────────
    if is_image:
        # OCR devolve lista de páginas com linhas de texto
        lines = []
        for page in result.get("pages", []):
            for block in page.get("text_lines", []):
                t = block.get("text", "").strip()
                if t:
                    lines.append(t)
            lines.append("")  # separador de página
        text_content  = "\n".join(lines)
        output_format = "markdown"  # tratar como texto simples
    else:
        text_content = result.get(output_format, "") or result.get("markdown", "")
        if output_format == "json":
            import json as _json
            text_content = _json.dumps(
                result.get("json", {}), ensure_ascii=False, indent=2)

    if not text_content or not text_content.strip():
        raise RuntimeError(
            "Datalab devolveu conteúdo vazio. "
            "O documento pode estar vazio, corrompido ou protegido por senha."
        )

    total_chars = len(text_content)
    logger.info("Datalab: extraídos %d caracteres", total_chars)

    # ── 4. Tratamento de imagens ──────────────────────────────────────────────
    if output_format == "markdown" and not is_image:
        if image_mode == "remove":
            text_content = re.sub(r"!\[[^\]]*\]\([^)]+\)", "", text_content)
        elif image_mode == "placeholder":
            text_content = re.sub(
                r"!\[([^\]]*)\]\([^)]+\)",
                lambda m: f"[{m.group(1) or 'IMAGEM'}]",
                text_content
            )
        # "keep" → não fazer nada, manter links de imagem no markdown

    # Guardar imagens base64 se existirem
    images = result.get("images", {})
    if images and image_mode == "keep":
        img_dir = out_path.parent / (out_path.stem + "_images")
        img_dir.mkdir(exist_ok=True)
        for img_name, img_b64 in images.items():
            try:
                (img_dir / img_name).write_bytes(base64.b64decode(img_b64))
            except Exception as e:
                logger.warning("Não foi possível guardar imagem %s: %s", img_name, e)
        logger.info("Datalab: %d imagens guardadas em %s", len(images), img_dir.name)

    # ── 5. Tradução ───────────────────────────────────────────────────────────
    if do_translate and lang not in ("", "none", "original"):
        translator = _get_translator(lang)
        if translator:
            logger.info("Datalab: a traduzir para %s...", lang)
            if output_format == "html":
                text_content = _translate_html(text_content, translator)
            elif output_format == "json":
                import json as _json
                obj = result.get("json", {})
                obj = _translate_json_obj(obj, translator)
                text_content = _json.dumps(obj, ensure_ascii=False, indent=2)
            else:
                text_content = _translate_markdown(text_content, translator)
            logger.info("Datalab: tradução completa")

    # ── 6. Guardar ficheiro ───────────────────────────────────────────────────
    # Se output é markdown, tentar converter directamente para DOCX
    saved_as_docx = False
    if output_format == "markdown":
        try:
            from markdown_to_docx import markdown_to_docx as _md2docx
            docx_path = out_path.with_suffix(".docx")
            # Passar imagens base64 convertidas para bytes se mode=keep
            img_dict = None
            if images and image_mode == "keep":
                img_dict = {}
                for img_name, img_b64 in images.items():
                    try:
                        img_dict[img_name] = base64.b64decode(img_b64)
                    except Exception:
                        pass
            _md2docx(text_content, str(docx_path), images_dict=img_dict)
            out_path = docx_path
            saved_as_docx = True
            logger.info("Datalab: markdown → DOCX via markdown_to_docx OK")
        except ImportError:
            logger.info("markdown_to_docx.py não encontrado — guardando como .md")
        except Exception as e:
            logger.warning("markdown_to_docx falhou (%s) — guardando como .md", e)

    if not saved_as_docx:
        ext_map  = {"markdown": ".md", "html": ".html", "json": ".json"}
        out_path = out_path.with_suffix(ext_map.get(output_format, ".md"))
        out_path.write_text(text_content, encoding="utf-8")

    size_kb = out_path.stat().st_size / 1024
    logger.info("Datalab: guardado → %s (%.1f KB)", out_path.name, size_kb)

    return {
        "total_chars":        total_chars,
        "extraction_method":  f"datalab_{label.lower()}",
        "pdf_type":           "scanned",
        "page_count":         pages,
        "quality_score":      score,
        "output_format":      "docx" if saved_as_docx else output_format,
        "output_path":        str(out_path),
        "output_name":        out_path.name,
        "size_kb":            round(size_kb, 1),
    }


# ── Exemplo de integração no main.py ─────────────────────────────────────────
"""
PASSO 1 — No topo do main.py, adicionar:

    from datalab_module import convert_with_datalab, datalab_available

PASSO 2 — Adicionar campo ao ProcessRequest:

    class ProcessRequest(BaseModel):
        job_id: str
        target_language: str = "pt"
        image_mode: str      = "placeholder"
        remove_images: bool  = False
        use_datalab: bool    = False          # <-- ADICIONAR
        datalab_mode: str    = "balanced"     # <-- ADICIONAR
        datalab_output: str  = "markdown"     # <-- ADICIONAR

PASSO 3 — No início de convert_pdf_to_docx(), adicionar:

    def convert_pdf_to_docx(pdf_path, docx_path, lang, do_translate, image_mode,
                             use_datalab=False, datalab_mode="balanced",
                             datalab_output="markdown"):
        if use_datalab and datalab_available():
            return convert_with_datalab(
                src_path=pdf_path,
                out_path=docx_path,
                lang=lang,
                do_translate=do_translate,
                output_format=datalab_output,
                mode=datalab_mode,
                image_mode=image_mode,
            )
        # ... resto do código Marker local existente ...

PASSO 4 — Em process_file(), passar os novos parâmetros:

    meta = convert_pdf_to_docx(pdf, docx_path, lang=lang,
                                do_translate=do_translate,
                                image_mode=image_mode,
                                use_datalab=do_use_datalab,
                                datalab_mode=datalab_mode,
                                datalab_output=datalab_output)

PASSO 5 — Em start_processing(), passar ao run_job:

    background_tasks.add_task(run_job, req.job_id, req.target_language,
                              image_mode, do_translate,
                              req.use_datalab, req.datalab_mode, req.datalab_output)
"""