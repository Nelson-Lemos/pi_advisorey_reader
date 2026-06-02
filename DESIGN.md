---
name: PDF Tradutor Pro
description: Conversor e tradutor de PDF para DOCX
colors:
  primary: "#7c5cfc"
  primary-deep: "#9333ea"
  primary-light: "#a78bfa"
  primary-dim: "rgba(124,92,252,0.12)"
  primary-glow: "rgba(124,92,252,0.25)"
  neutral-bg: "#0a0a0f"
  neutral-surface: "#111118"
  neutral-surface2: "#18181f"
  neutral-border: "#23232f"
  neutral-border2: "#2e2e3d"
  text-primary: "#e8e8f0"
  text-secondary: "#a0a0b8"
  text-muted: "#5a5a72"
  success: "#22c55e"
  success-dim: "rgba(34,197,94,0.12)"
  warning: "#f59e0b"
  warning-dim: "rgba(245,158,11,0.12)"
  error: "#ef4444"
  error-dim: "rgba(239,68,68,0.12)"
  info: "#3b82f6"
typography:
  display:
    fontFamily: "Syne, sans-serif"
    fontSize: "clamp(1.25rem, 2vw, 1.375rem)"
    fontWeight: 700
    lineHeight: 1.2
    letterSpacing: "-0.5px"
  body:
    fontFamily: "DM Sans, sans-serif"
    fontSize: "0.9375rem"
    fontWeight: 400
    lineHeight: 1.5
  label:
    fontFamily: "DM Mono, monospace"
    fontSize: "0.75rem"
    fontWeight: 400
    lineHeight: 1.4
    letterSpacing: "0.06em"
rounded:
  sm: "6px"
  md: "8px"
  lg: "12px"
  xl: "14px"
  xxl: "20px"
spacing:
  xs: "4px"
  sm: "8px"
  md: "12px"
  lg: "16px"
  xl: "20px"
  xxl: "24px"
  section: "40px"
components:
  button-primary:
    backgroundColor: "{colors.primary}"
    textColor: "#fff"
    rounded: "{rounded.md}"
    padding: "12px 24px"
    size: "15px 500"
  button-primary-hover:
    backgroundColor: "{colors.primary-deep}"
    textColor: "#fff"
    rounded: "{rounded.md}"
    padding: "12px 24px"
  button-secondary:
    backgroundColor: "{colors.neutral-surface2}"
    textColor: "{colors.text-secondary}"
    rounded: "{rounded.md}"
    padding: "12px 24px"
    border: "1px solid {colors.neutral-border}"
  button-ghost:
    backgroundColor: "transparent"
    textColor: "{colors.text-muted}"
    rounded: "{rounded.md}"
    padding: "12px 16px"
  tab-active:
    backgroundColor: "{colors.primary}"
    textColor: "#fff"
    rounded: "{rounded.md}"
    padding: "9px 20px"
  tab-inactive:
    backgroundColor: "transparent"
    textColor: "{colors.text-muted}"
    rounded: "{rounded.md}"
    padding: "9px 20px"
  input-select:
    backgroundColor: "{colors.neutral-bg}"
    textColor: "{colors.text-primary}"
    rounded: "{rounded.md}"
    border: "1px solid {colors.neutral-border2}"
    padding: "9px 12px"
  toggle-off:
    backgroundColor: "{colors.neutral-border2}"
    rounded: "100px"
    size: "44px 24px"
  toggle-on:
    backgroundColor: "{colors.primary}"
    rounded: "100px"
    size: "44px 24px"
---

# Design System: PDF Tradutor Pro

## 1. Overview

**Creative North Star: "The Professional Studio"**

Um estúdio escuro mas acolhedor — silencioso, iluminado por um único foco de luz violeta. O fundo não compete: superfícies em camadas tonais criam profundidade sem sombras dramáticas. Cada elemento está onde deve estar. O roxo aparece com intenção cirúrgica: botões, abas ativas, destaque de progresso. O resto é tipografia nítida em DM Sans e Syne, espaçamento generoso, bordas sutis.

Este sistema rejeita explicitamente o visual "dark gamer" — sem neons, sem gradientes exagerados, sem brilho artificial. A paleta é controlada, profissional. O app parece uma ferramenta de escritório que respeita o usuário.

**Key Characteristics:**
- Camadas tonais para profundidade (bg → surface → surface2), não sombras pesadas
- Um único acento roxo ($accent #7c5cfc) usado em ≤15% da superfície
- Tipografia como elemento decorativo principal — hierarquia via peso e tamanho
- Espaçamento generoso e consistente (padding base 24-40px)
- Cantos visivelmente arredondados (12px componentes, 20px cartões)

## 2. Colors

Paleta escura controlada com um acento violeta. Fundos tendem a azul muito sutil para evitar o cinza genérico.

### Primary
- **Violeta Profissional** (#7c5cfc): acento principal. Botões primários, abas ativas, destaque de progresso, foco de inputs. Usado com moderação cirúrgica.
- **Violeta Profundo** (#9333ea): variante para hover de botões primários.
- **Violeta Claro** (#a78bfa): cor secundária para detalhes (counters, badges, labels variantes).

### Neutral
- **Fundo Base** (#0a0a0f): background da página. Tom preto-azulado muito escuro.
- **Superfície** (#111118): cartões, containers, abas inativas.
- **Superfície Elevada** (#18181f): file-list header, option-group, hover states.
- **Borda** (#23232f): borda default de cartões e containers.
- **Borda Forte** (#2e2e3d): borda de hover, inputs, drop-zone dashed.
- **Texto Principal** (#e8e8f0): headings, labels principais.
- **Texto Secundário** (#a0a0b8): metadados, subtítulos.
- **Texto Leve** (#5a5a72): placeholders, hints, informações auxiliares.

### Semantic
- **Verde Sucesso** (#22c55e): operações concluídas, status completed, botão de download, status-dot.
- **Âmbar Alerta** (#f59e0b): status processing, atenção.
- **Vermelho Erro** (#ef4444): status failed, botão de exclusão hover.
- **Azul Info** (#3b82f6): uso futuro para informational toasts.

Cada cor semântica tem sua variante `-dim` (12% opacidade) para backgrounds sutis de status pills e badges.

### Named Rules
**The Single Voice Rule.** O acento roxo é a única cor de destaque na interface. Cores semânticas (verde, âmbar, vermelho) aparecem apenas em status e ações específicas — nunca competem com o roxo. Se mais de uma cor não-neutra está competindo pela atenção, o design quebrou a regra.

## 3. Typography

**Display Font:** Syne (sans-serif, com fallback system-ui)
**Body Font:** DM Sans (sans-serif, com fallback system-ui)
**Label/Mono Font:** DM Mono (monospace, com fallback Courier New)

**Character:** Contraste entre a Syne geométrica e expressiva (títulos) e a DM Sans humanista e legível (corpo). A DM Mono traz precisão técnica para metadados e status. O conjunto passa confiança: nem muito casual, nem muito rígido.

### Hierarchy
- **Display** (Syne 700, clamp(1.25rem, 2vw, 1.375rem), 1.2): Logo e títulos de cartões de destaque.
- **Headline** (Syne 600, 1.125rem, 1.3): Título do progress-card e seções.
- **Title** (Syne 600, 0.875rem, 1.3): File-list header.
- **Body** (DM Sans 400, 0.9375rem, 1.5): Texto corrido, labels de campo, botões. Máximo 75ch.
- **Label** (DM Mono 400/500, 0.6875rem–0.75rem, 1.4, letter-spacing 0.06em): Option-label, stats, status, file-size. Uppercase para option-label.

### Named Rules
**The Syne Ceiling Rule.** Syne é exclusiva para títulos e o logo. Nunca use Syne em corpo de texto, botões (exceto se visualmente destacado), ou labels. DM Sans ou DM Mono para todo o resto.

## 4. Elevation

Profundidade é criada exclusivamente por camadas tonais (luminância) — não por sombras. Cada nível de superfície é uma cor distinta da paleta neutral:

| Layer | Token | Uso |
|---|---|---|
| 0 (fundo) | `--bg` #0a0a0f | Background da página |
| 1 (superfície) | `--surface` #111118 | Cartões, containers |
| 2 (elevado) | `--surface2` #18181f | Headers internos, option-groups, hover |

Não há box-shadows no sistema de elevação. A única exceção é o brilho sutil (`box-shadow`) nos botões primários para destacar a ação principal, e no toast para temporária separação visual. O resto da profundidade vem da diferença de tom entre as camadas.

### Named Rules
**The Flat-By-Default Rule.** Superfícies são planas em repouso. Nenhuma sombra ambiental. O único relevo aceitável é o glow em elementos interativos primários (btn-primary, tab active) e o shadow fugaz no toast.

## 5. Components

### Buttons
- **Shape:** Cantos suavemente arredondados (10px, arredondamento médio para ação).
- **Primary:** Fundo gradiente (--accent → #9333ea), texto branco, glow sutil (0 4px 20px var(--accent-glow)). Hover: levanta 1px, glow intensificado. Disabled: 40% opacidade.
- **Secondary:** Fundo surface2, texto text2, borda 1px border. Hover: fundo surface, texto text.
- **Ghost:** Fundo transparente, texto text3. Hover: fundo surface2.
- **Download:** Fundo verde sólido (#22c55e), texto branco, glow verde. Exclusivo para o botão de download no result-card.
- **Small** (.btn-sm): Padding reduzido (6px 12px), fonte 12px, raio 8px. Usado em ações de histórico.

### Inputs / Selects
- **Style:** Fundo bg, borda 1px border2, raio 8px. Chevron customizado via SVG inline.
- **Focus:** Borda muda para --accent.
- **Label:** Uppercase, tracking 0.08em, cor text3, fonte 11px, DM Sans 600.

### Toggle
- **Shape:** Pill arredondado (100px), 44×24px.
- **Off:** Fundo border2. Círculo interno branco à esquerda.
- **On:** Fundo accent. Círculo desliza 20px para direita.
- **Transition:** 0.2s ease em background e transform.

### Tabs / Navigation
- **Container:** Fundo surface, borda 1px border, raio 14px, padding 4px, gap 4px.
- **Tab Inativa:** Fundo transparente, texto text3. Hover: fundo surface2, texto text2.
- **Tab Ativa:** Fundo accent, texto branco, glow. Transição 0.2s.

### Cards / Containers
- **Upload Card:** Fundo surface, borda 1px border, raio 20px, padding 40px.
- **Progress Card:** Mesmo estilo, padding 32px 36px.
- **Result Card:** Fundo com gradiente sutil verde→roxo (5% opacidade), borda verde 20% opacidade.
- **History Card:** Fundo surface, borda border, raio 12px, padding 20px 24px. Hover: borda border2.
- **Drop Zone:** Borda 2px dashed border2, fundo bg. Hover/drag: borda accent, fundo accent-dim.

### Chips / Badges
- **Type Badge:** Fundo surface2, borda border, raio 6px, DM Mono 11px, text3.
- **Status Pills:** File-status pills usam cor semântica + variante dim como fundo. Raio 6px, DM Mono 12px.
- **File Count:** Pill com fundo accent-dim, borda accent 20%, texto accent2, raio 100px.
- **API Status:** Pill verde com borda 1px, raio 100px, texto DM Mono 12px. Ponto pulsante.

### Progress Bar
- **Wrap:** Altura 8px, fundo surface2, raio 100px.
- **Fill:** Gradiente horizontal (--accent → #c084fc), transição 0.5s ease, shimmer animado.
- **Stats:** File-status com ícones semânticos e valores em Syne.

### File List
- **Item:** Padding 12px 20px, borda inferior border. Hover: fundo surface2.
- **Icon:** Gradiente vermelho-laranja (#ef4444 → #f97316), raio 8px, 36×36px.
- **Actions:** Botão ghost que fica vermelho no hover para exclusão.

### Toast
- **Container:** Fixo bottom-right, gap 8px, z-index 9999.
- **Toast item:** Fundo surface2, borda border2, raio 10px, padding 12px 18px, box-shadow: --shadow. Max-width 320px.
- **Variants:** Borda verde (success) ou vermelha (error) a 30% opacidade.
- **Entrada:** slide-up 0.3s ease. Auto-dismiss via JS.

### Spinner
- 16×16px, borda 2px, white 30% opacidade + top-color branco, animação spin 0.6s linear.

## 6. Do's and Don'ts

### Do:
- **Do** usar camadas tonais para profundidade. bg → surface → surface2 é o único sistema de elevação.
- **Do** limitar o uso do roxo a ≤15% da tela. Raridade é a força do acento.
- **Do** usar DM Sans para corpo, Syne para títulos, DM Mono para metadados — cada fonte no seu papel.
- **Do** manter cantos consistentemente arredondados: 12px para componentes, 20px para cartões grandes.
- **Do** usar padding generoso (40px em cartões, 24px no app wrapper) para criar respiro.
- **Do** usar o gradiente nos botões primários — é o único gradiente permitido no sistema.

### Don't:
- **Don't** usar neons, gradientes múltiplos, glassmorphism, ou brilho exagerado — o visual "dark gamer" é explicitamente proibido.
- **Don't** adicionar segundas cores de acento que competem com o roxo. Cores semânticas são para status, não para decoração.
- **Don't** usar sombras para profundidade. O sistema é baseado em camadas tonais.
- **Don't** usar Syne em corpo de texto, botões, ou labels — é exclusiva para títulos.
- **Don't** usar border-left ou border-right maior que 1px como faixa decorativa.
- **Don't** criar cards aninhados.
- **Don't** usar emoji como ícone principal em botões ou ações — substitua por SVGs inline ou elementos visuais consistentes.
