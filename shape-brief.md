# Design Brief: Icones Lucide + Barra de Progresso por Arquivo

## 1. Feature Summary

Substituir todos os emojis da interface por ícones SVG Lucide, e adicionar uma barra de progresso individual para cada arquivo na file-list durante o processamento. O app já tem uma barra de progresso geral no progress-card; esta é complementar, por arquivo.

## 2. Primary User Action

O usuário acompanha visualmente o progresso individual de cada PDF na fila de processamento, vendo exatamente onde cada arquivo está (upload → extração → tradução → DOCX), em vez de apenas uma barra geral.

## 3. Design Direction

- **Color strategy:** Restrained (mesma do app — um acento roxo ≤15%). Segue o DESIGN.md.
- **Theme:** Profissionais em ambiente de escritório, dark mode sóbrio (já definido em PRODUCT.md).
- **References:** Segue fielmente o DESIGN.md — The Professional Studio. Lucide como sistema de ícones (substituindo emojis).

## 4. Scope

- **Fidelity:** Production-ready
- **Breadth:** Todo o frontend (icones em toda a interface) + file-list (barra por arquivo)
- **Interactivity:** Shipped-quality component
- **Time intent:** Polish until it ships

## 5. Layout Strategy

- **Icones Lucide:** Substituir cada emoji por seu equivalente Lucide, mantendo posição e tamanho atuais. Os ícones devem herdar a cor do texto do elemento pai via `currentColor`.
- **Barra por arquivo:** Cada `.file-item` ganha uma mini barra de progresso (altura ~4px) abaixo do nome do arquivo, com transição suave. O status do arquivo (pending/processing/completed/failed) continua existindo como badge.

## 6. Key States

- **Default (pré-upload):** File-list vazia, sem alterações.
- **Files added (pending):** Lista de arquivos, cada um sem barra (ou barra em 0%), status "pending".
- **Processing (por arquivo):** Barra individual animando de 0% → 100%. Status muda para "processing".
- **Completed:** Barra cheia (verde), status "completed". Ícone de check.
- **Failed:** Barra para com highlight vermelho, status "failed". Ícone de erro.
- **Mixed:** Alguns concluídos, outros processando, outros falhos — cada um com seu estado independente.

## 7. Interaction Model

- O backend já envia updates de progresso via polling (`/api/jobs/{id}`) com `processed_count` e `total_count`. Vamos expandir o payload para incluir progresso individual por arquivo (opcional — se não vier, usar progresso uniforme: `100% / total_count` por arquivo).
- A barra de cada arquivo atualiza em tempo real conforme o backend reporta.

## 8. Content Requirements

- **Icon mapping (emoji → Lucide):**
  - 📄 (logo) → `file-text`
  - ☁ (upload) → `cloud-upload`
  - ✅ (success) → `check-circle`
  - ❌ (error/delete) → `x-circle` / `trash-2`
  - ⬇ (download) → `download`
  - 🗑 (delete) → `trash-2`
  - 🔄 (processing) → `loader` (com animação spin)
  - País flags → `globe` ou `languages`
  - ⚙ (options) → `settings`
  - 📂 (empty) → `folder-open`
  - ⏳ (pending) → `clock`
- **Progress bar:** Sem copy adicional — a barra visual + % já comunica.

## 9. Recommended References

- interaction-design.md
- motion-design.md (para a animação das barras e transições de ícones)

## 10. Open Questions

- O backend precisa expandir o payload do job status para incluir progresso individual? Atualmente `processed_count` e `total_count` são os únicos indicadores. Podemos derivar progresso uniforme (`index / total * 100`) sem mudanças no backend.
- Como carregar Lucide? CDN via `<script src="https://unpkg.com/lucide@latest">` + `lucide.createIcons()` é a abordagem mais simples para vanilla HTML.
