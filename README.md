# 📄 PDF Tradutor Pro

Sistema web completo para converter e traduzir PDFs em lote para DOCX automaticamente.

## ✨ Funcionalidades

- **Upload em lote** — 1 a 100+ PDFs, ficheiros individuais, múltiplos ou ZIP
- **Detecção automática** — distingue PDFs de texto e PDFs digitalizados (escaneados)
- **OCR inteligente** — extrai texto de PDFs com imagens via Tesseract
- **Tradução automática** — para Português, Inglês, Espanhol, Francês, Alemão e mais
- **Geração de DOCX** — preserva negrito, itálico, tamanhos de fonte e estrutura
- **Fila de processamento** — progresso em tempo real com estatísticas
- **Download em ZIP** — todos os DOCX traduzidos num único ficheiro
- **Histórico** — trabalhos anteriores com opção de re-download

## 🚀 Instalação e Arranque

### 1. Clonar / extrair o projecto

```bash
cd pdf-translator
```

### 2. Criar ambiente virtual (recomendado)

```bash
python -m venv venv
source venv/bin/activate        # Linux/Mac
venv\Scripts\activate           # Windows
```

### 3. Instalar dependências Python

```bash
pip install -r requirements.txt
```

### 4. Instalar Tesseract OCR (para PDFs digitalizados)

**Ubuntu/Debian:**
```bash
sudo apt install tesseract-ocr tesseract-ocr-por tesseract-ocr-eng
```

**macOS:**
```bash
brew install tesseract tesseract-lang
```

**Windows:**
Descarregar instalador em: https://github.com/UB-Mannheim/tesseract/wiki

### 5. Iniciar o servidor

```bash
python main.py
```

Ou com uvicorn directamente:
```bash
uvicorn main:app --host 0.0.0.0 --port 8000 --reload
```

### 6. Abrir no browser

```
http://localhost:8000
```

## 📁 Estrutura do Projecto

```
pdf-translator/
├── main.py              # Backend FastAPI
├── requirements.txt     # Dependências Python
├── static/
│   └── index.html       # Frontend completo (HTML/CSS/JS)
├── uploads/             # PDFs enviados (gerado automaticamente)
├── outputs/             # DOCX traduzidos (gerado automaticamente)
└── jobs/                # Estado dos trabalhos em JSON
```

## 🔄 Fluxo de Processamento

```
Upload PDFs/ZIP
      ↓
Análise do tipo (texto ou digitalizado)
      ↓
Extracção de texto (PyMuPDF / pdfminer)
      ↓ (se digitalizado)
OCR via Tesseract
      ↓
Tradução via Google Translate (deep-translator)
      ↓
Geração de DOCX (python-docx)
      ↓
Compactação em ZIP
      ↓
Download
```

## 🌐 API Endpoints

| Método | Endpoint | Descrição |
|--------|----------|-----------|
| `POST` | `/api/upload` | Enviar ficheiros PDF/ZIP |
| `POST` | `/api/process` | Iniciar processamento |
| `GET`  | `/api/jobs/{id}` | Estado de um trabalho |
| `GET`  | `/api/jobs` | Listar todos os trabalhos |
| `GET`  | `/api/download/{id}/zip` | Descarregar ZIP com todos os DOCX |
| `GET`  | `/api/download/{id}/file/{fid}` | Descarregar DOCX individual |
| `DELETE` | `/api/jobs/{id}` | Eliminar trabalho e ficheiros |
| `GET`  | `/api/health` | Estado da API e dependências |

## ⚙️ Idiomas Suportados

| Código | Idioma |
|--------|--------|
| `pt` | Português |
| `en` | Inglês |
| `es` | Espanhol |
| `fr` | Francês |
| `de` | Alemão |
| `it` | Italiano |
| `zh-cn` | Chinês (Simplificado) |
| `ar` | Árabe |
| `ru` | Russo |
| `ja` | Japonês |

## 📝 Notas

- A tradução usa a API gratuita do Google Translate via `deep-translator`
- Para uso intensivo, considere uma API paga (DeepL, OpenAI, etc.)
- O OCR requer o Tesseract instalado no sistema operativo
- Os ficheiros processados são guardados localmente em `outputs/`
