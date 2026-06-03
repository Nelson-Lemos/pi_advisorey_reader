# Plano: Substituir Google Translate pelo NLLB-200 Distilled 600M

## Ficheiros

### 1. CRIAR: `nllb_translator.py`

Módulo com:
- Singleton do modelo `facebook/nllb-200-distilled-600M`
- Classe `NLLBTranslator` com método `.translate(text)`
- Mapa ISO → FLORES-200 (`get_flores_code()`)
- Função `create_translator(lang)` compatível com `get_translator()` do main.py

### 2. MODIFICAR: `main.py`

**Remover:**
```python
from deep_translator import GoogleTranslator
```

**Substituir `get_translator()` por:**
```python
def get_translator(lang: str):
    from nllb_translator import create_translator
    return create_translator(lang)
```

**Ajustar `_translate_chunk()`:**
- `MAX_CHARS = 4000` (era 1500 — NLLB suporta mais tokens)
- Remover `time.sleep(0.5)` (desnecessário para modelo local)
- Remover retry logic (modelo local não falha por rede)
- Manter cache (`_translation_cache`)

**Manter inalterado:**
- `translate_markdown()` — preserva estrutura markdown
- `translate_html()` — traduz HTML via BeautifulSoup
- `translate_json_obj()` — traduz JSON recursivamente
- `_translation_cache` — cache de traduções

### 3. MODIFICAR: `requirements.txt`

```diff
- deep-translator
+ transformers
+ torch
+ sentencepiece
+ accelerate
```

## Fluxo final

```
Markdown (após conversão Datalab)
       ↓
translate_markdown(md_text, NLLBTranslator("pt"))
  ├── Preserva markdown (cabeçalhos, listas, imagens, código)
  ├── Divide textos longos em chunks de 4000 chars
  └── Cada chunk → NLLBTranslator.translate(chunk)  (CPU local)
       ↓
markdown_to_docx() → DOCX traduzido com NLLB-200
```

## Para implementar

```
python -m pip install transformers torch sentencepiece accelerate
```

Depois criar `nllb_translator.py` com o conteúdo do plano, modificar `main.py` e `requirements.txt` conforme descrito acima.
