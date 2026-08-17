# AGENTS.md — Sara / V4 propostas

Fonte da verdade operacional: [`sistema-proposta.html`](sistema-proposta.html).

Este repo é o kit da Sara (pipe E.E, farmer, propostas e scripts). Agentes devem **seguir o sistema**, não reinventar a cada pedido.

## Antes de qualquer proposta

1. Exigir **brief mínimo** (cliente, produtos do catálogo, preços finais, referência visual, narrativa em 1 linha).
2. Se faltar campo obrigatório → **parar e perguntar**. Não inventar preço, escopo ou visual.
3. Abrir a referência visual pedida e **copiar CSS/estrutura de slides** dela.

## Travado (não reabrir no meio do trabalho)

| Tema | Decisão |
|------|---------|
| Visual deck cliente | Vermelho V4 · Outfit + Plus Jakarta · slides full-bleed (padrão Modular / Martins / Motéis / Estação Delta) |
| Hub interno | Dark + Inter (`index.html`, `farmer-sara.html`, score, planos) — não misturar com deck cliente |
| Nomes de produto | Só catálogo: **EC · Assessoria de Growth · CRM · Social Media · SDR IA · E-commerce B2B** |
| Preço | Vem do brief. Não “calcular” nem arredondar sem pedido |
| Kit padrão | `proposta-*.html` + `script-*-pitch.html`. Word/PDF/PPTX só se brief pedir |
| Variantes | Não criar `-v2` / visual paralelo sem brief explícito |

## Sequência canônica de slides

Capa → Dor/contexto → Diagnóstico → Produto(s) → Como opera → Âncora de mercado (opcional) → Investimento → Próximos passos.

## Script de pitch

- Conversa, não monólogo
- Perguntas abertas (evitar “faz sentido?”)
- Objeções **antes** do preço
- Gate de avanço antes de abrir valor
- Um preço / uma decisão (ou A/B se brief pedir)
- Pedir a venda + próximo passo concreto

## Onde puxar referência

- **Growth + EC + CRM bonus** → `proposta-modular-estruturas.html`
- **Growth pós-EE** → `proposta-martins-magazine-growth.html`
- **Growth + SDR IA** → `proposta-moteis-fortaleza-consultiva.html`
- **EC pura** → `proposta-estacao-delta-ec.html`
- **EC + E-com + Growth** → `proposta-isoluz.html`
- **G6** → usar conteúdo/preço; **não** copiar o visual Sora para propostas novas

## Arquivos e Pages

- Manter HTML de proposta/script na **raiz** (GitHub Pages).
- Assets (logo, docx) em `assets/` quando possível.
- Não remover `.nojekyll`.
- Não mover arquivos publicados sem atualizar todos os links.

## Anti-retrabalho

- “Igual à G6” sem versão = ambíguo → pedir qual arquivo.
- Mudança de tipografia/layout depois do 1º draft = fora, salvo erro.
- Pedidos de Word/PDF depois do HTML = ciclo separado, sem refazer o deck.
- Pipe/safra (`index.html`, `farmer-sara.html`) só muda com dados novos — não no fluxo de proposta.
