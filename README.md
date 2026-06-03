# PDF Tradutor Pro

Aplicacao web local para converter PDFs em DOCX, traduzir o texto e remover imagens/logos do documento final.

## Objetivo

O fluxo principal e:

1. Enviar um ou mais PDFs, ou um ZIP com PDFs.
2. Converter cada PDF para DOCX preservando tabelas, fontes, colunas e espacamento com `pdf2docx`.
3. Traduzir texto em paragrafos e tabelas no DOCX.
4. Remover imagens/logos/carimbos do DOCX final por padrao.
5. Disponibilizar download individual e ZIP com todos os DOCX.

> Nota: PDF nao e um formato editavel. A preservacao de layout depende da qualidade do PDF original. `pdf2docx` e usado como caminho principal porque preserva melhor a infraestrutura visual do ficheiro do que reconstrucoes manuais.

## Limites Operacionais

- Tempo maximo por ficheiro: 3 minutos.
- Upload maximo por ficheiro: 100 MB.
- ZIP maximo: 100 PDFs.
- ZIP maximo descompactado: 100 MB.
- Jobs, uploads e outputs sao artefatos locais e nao devem ser versionados.

## Instalar

```bash
python -m venv venv
venv\Scripts\activate
pip install -r requirements.txt
```

Para OCR no futuro, instale tambem o Tesseract no sistema operacional. O fluxo atual prioriza PDFs com texto extraivel e conversao fiel via `pdf2docx`.

## Rodar

```bash
python main.py
```

Depois abra:

```text
http://localhost:8000
```

## API

| Metodo | Endpoint | Descricao |
| --- | --- | --- |
| `POST` | `/api/upload` | Envia PDFs ou ZIP |
| `POST` | `/api/process` | Inicia conversao/traducao |
| `GET` | `/api/jobs/{id}` | Consulta estado e progresso |
| `GET` | `/api/jobs` | Lista historico local |
| `GET` | `/api/download/{id}/zip` | Baixa ZIP com DOCX |
| `GET` | `/api/download/{id}/file/{fid}` | Baixa DOCX individual |
| `DELETE` | `/api/jobs/{id}` | Remove job e ficheiros |
| `GET` | `/api/health` | Estado da API e dependencias |

## Testes

```bash
python -m pytest tests/test_backend.py -q
```

## Estrutura

```text
main.py              Backend FastAPI
static/index.html    Frontend HTML/CSS/JS
requirements.txt     Dependencias Python
tests/               Testes de backend
uploads/             PDFs enviados, gerado localmente
outputs/             DOCX/ZIP gerados, localmente
jobs/                Estado dos jobs, localmente
```
