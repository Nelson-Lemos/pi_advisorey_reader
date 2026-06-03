# -*- coding: utf-8 -*-
"""
nllb_translator.py  —  Tradução via NLLB-200 Distilled 600M (Meta AI)
======================================================================
Substitui o Google Translate (deep-translator) pelo modelo NLLB-200
da Meta, executado localmente em CPU com carregamento singleton.

Modelo: facebook/nllb-200-distilled-600M
"""

import logging
import re

logger = logging.getLogger(__name__)

# ── Singleton: modelo carregado uma única vez ────────────────────────────────
_model = None
_tokenizer = None


def _load_model():
    global _model, _tokenizer
    if _model is not None:
        return _model, _tokenizer
    from transformers import AutoModelForSeq2SeqLM, AutoTokenizer
    logger.info("A carregar NLLB-200 Distilled 600M (~600 MB RAM)...")
    _tokenizer = AutoTokenizer.from_pretrained("facebook/nllb-200-distilled-600M")
    _model = AutoModelForSeq2SeqLM.from_pretrained(
        "facebook/nllb-200-distilled-600M",
        low_cpu_mem_usage=True,
    )
    _model.eval()
    logger.info("NLLB-200 carregado com sucesso")
    return _model, _tokenizer


# ── Mapa de códigos ISO → FLORES-200 ─────────────────────────────────────────
LANG_CODE_MAP = {
    "pt": "por_Latn", "pt-br": "por_Latn",
    "en": "eng_Latn", "en-us": "eng_Latn", "en-gb": "eng_Latn",
    "es": "spa_Latn",
    "fr": "fra_Latn",
    "de": "deu_Latn",
    "it": "ita_Latn",
    "nl": "nld_Latn",
    "ru": "rus_Cyrl",
    "zh-cn": "zho_Hans", "zh-tw": "zho_Hant",
    "ja": "jpn_Jpan",
    "ko": "kor_Hang",
    "ar": "ara_Arab",
    "hi": "hin_Deva",
    "bn": "ben_Beng",
    "af": "afr_Latn",
    "sq": "als_Latn",
    "am": "amh_Ethi",
    "hy": "hye_Armn",
    "az": "azj_Latn",
    "eu": "eus_Latn",
    "be": "bel_Cyrl",
    "bs": "bos_Latn",
    "bg": "bul_Cyrl",
    "ca": "cat_Latn",
    "hr": "hrv_Latn",
    "cs": "ces_Latn",
    "da": "dan_Latn",
    "et": "est_Latn",
    "tl": "tgl_Latn",
    "fi": "fin_Latn",
    "gl": "glg_Latn",
    "ka": "kat_Geor",
    "el": "ell_Grek",
    "gu": "guj_Gujr",
    "ht": "hat_Latn",
    "ha": "hau_Latn",
    "iw": "heb_Hebr",
    "hu": "hun_Latn",
    "is": "isl_Latn",
    "ig": "ibo_Latn",
    "id": "ind_Latn",
    "ga": "gle_Latn",
    "jw": "jav_Latn",
    "kn": "kan_Knda",
    "kk": "kaz_Cyrl",
    "km": "khm_Khmr",
    "ku": "kur_Latn",
    "ky": "kir_Cyrl",
    "lo": "lao_Laoo",
    "la": "lat_Latn",
    "lv": "lav_Latn",
    "lt": "lit_Latn",
    "mk": "mkd_Cyrl",
    "ms": "msa_Latn",
    "ml": "mal_Mlym",
    "mt": "mlt_Latn",
    "mi": "mri_Latn",
    "mr": "mar_Deva",
    "mn": "khk_Cyrl",
    "my": "mya_Mymr",
    "ne": "npi_Deva",
    "no": "nob_Latn",
    "ps": "pus_Arab",
    "fa": "pes_Arab",
    "pl": "pol_Latn",
    "pa": "pan_Guru",
    "ro": "ron_Latn",
    "sm": "smo_Latn",
    "sr": "srp_Cyrl",
    "si": "sin_Sinh",
    "sk": "slk_Latn",
    "sl": "slv_Latn",
    "so": "som_Latn",
    "sw": "swh_Latn",
    "sv": "swe_Latn",
    "tg": "tgk_Cyrl",
    "ta": "tam_Taml",
    "te": "tel_Telu",
    "th": "tha_Thai",
    "tr": "tur_Latn",
    "uk": "ukr_Cyrl",
    "ur": "urd_Arab",
    "uz": "uzn_Latn",
    "vi": "vie_Latn",
    "cy": "cym_Latn",
    "xh": "xho_Latn",
    "yi": "ydd_Hebr",
    "yo": "yor_Latn",
    "zu": "zul_Latn",
}


def get_flores_code(iso_code: str) -> str:
    """Converte código ISO 639-1 para código FLORES-200."""
    code = LANG_CODE_MAP.get(iso_code.lower().strip())
    if code:
        return code
    logger.warning("Código de idioma não mapeado: %s — usar por_Latn", iso_code)
    return "por_Latn"


# ── Tradutor NLLB ────────────────────────────────────────────────────────────

class NLLBTranslator:
    """Wrapper para o modelo NLLB-200 com interface .translate(text)."""

    def __init__(self, target_lang: str):
        self.model, self.tokenizer = _load_model()
        self.target_code = get_flores_code(target_lang)
        logger.debug("NLLB tradutor pronto: → %s (%s)", target_lang, self.target_code)

    def translate(self, text: str) -> str:
        if not text or not text.strip():
            return text
        import torch
        try:
            inputs = self.tokenizer(
                text,
                return_tensors="pt",
                truncation=True,
                max_length=512,
            )
            with torch.no_grad():
                translated = self.model.generate(
                    **inputs,
                    forced_bos_token_id=self.tokenizer.lang_code_to_id[self.target_code],
                    max_length=512,
                    num_beams=2,
                    early_stopping=True,
                )
            return self.tokenizer.batch_decode(translated, skip_special_tokens=True)[0]
        except Exception as e:
            logger.error("Erro na tradução NLLB: %s", e)
            return text


# ── Função de conveniência (compatível com get_translator do main.py) ────────

def create_translator(lang: str):
    """Cria e devolve um NLLBTranslator para o idioma dado."""
    if not lang or lang in ("none", "original", ""):
        return None
    try:
        t = NLLBTranslator(lang)
        logger.info("NLLB-200 tradutor pronto: → %s", lang)
        return t
    except Exception as e:
        logger.error("Falha ao criar tradutor NLLB: %s", e)
        return None
