"""
hf_translator.py  —  Tradução cloud via Google Translate (deep-translator)
========================================================================
Usa o Google Translate através da biblioteca deep-translator.
Gratuito, sem API key, sem download de modelos.

Fallback automático para NLLB local se a API falhar.
"""

import os
import time
import logging

logger = logging.getLogger(__name__)

_translation_cache: dict = {}


class GoogleTranslatorWrapper:
    """Tradutor via Google Translate (cloud, gratuito)."""

    def __init__(self, target_lang: str):
        self.target_lang = target_lang

    def translate(self, text: str) -> str:
        if not text or not text.strip():
            return text

        cache_key = text
        if cache_key in _translation_cache:
            return _translation_cache[cache_key]

        for attempt in range(3):
            try:
                from deep_translator import GoogleTranslator
                t = GoogleTranslator(source="auto", target=self.target_lang)
                result = t.translate(text)
                if result:
                    _translation_cache[cache_key] = result
                    return result
                return text
            except Exception as e:
                logger.warning("Google Translate error (attempt %d/3): %s", attempt + 1, e)
                if attempt < 2:
                    time.sleep(3)
                    continue
                return text

        return text


def create_translator(lang: str):
    """Cria tradutor cloud. Fallback para NLLB local se necessário."""
    if not lang or lang in ("none", "original", ""):
        return None

    try:
        t = GoogleTranslatorWrapper(lang)
        logger.info("Tradutor: Google Translate (cloud) → %s", lang)
        return t
    except Exception as e:
        logger.warning("Erro ao criar tradutor cloud: %s", e)

    logger.info("Fallback para NLLB local")
    try:
        from nllb_translator import create_translator as local_create
        return local_create(lang)
    except Exception as e:
        logger.error("Erro ao criar tradutor local: %s", e)
        return None
