# Estrutura da entrega — Deck HTML G6 × V4

Documento enxuto para reutilizar/editar com IA (baixo token).

- Deck: https://sarapizzico-hue.github.io/Sara/g6/
- Formato: HTML standalone 1600×900 · IBM Plex · navegável · GitHub Pages

## 1) Tese
Do comercial artesanal à máquina de receita.  
Não vender horas/PDF. Vender **capacidade instalada** (processo + CRM em uso + gestão + retenção).

## 2) Arco
Dor → vazamento com dados → encaixe → visão → promessa/formato → produto → fases → sucesso (Receita Previsível) → recap → investimento.

## 3) Ordem dos slides (entrega atual = 18)

1. **Capa** — tese de transformação, brand-first. “Do comercial artesanal à máquina de receita”.
2. **Contexto / Dor G6** — “Não é falta de lead. É comercial reativo sem sistema.”
3. **Vazando** — ~1.000 entradas × 3.755 saídas (jan–jun). CPL R$7,51. WhatsApp ~1min. Churn real não conciliado.
4. **Dores × G6** — 6 dores nomeadas × o que a EC fecha.
5. **Visão** — fundação antes de meta. Sistema → Gestão → Receita.
6. **Promessa** — projeto fechado; V4 instala fundação. Não é PDF.
7. **Formato** — 12 semanas · 106,5h · V4+G6 · presencial F1+F5 · handover.
8. **Produto** — fundação operacional + controle/adesão.
9. **Fases (visão)** — 6 fases (18 / 15 / 15,5 / 20 / 22 / 10) + 6h QA = 106,5h.
10. **F1 Diagnóstico** — 18h · entrevistas · auditoria CRM · presencial loja+PAP.
11. **F2 Arquitetura** — 15h · dual-track loja×PAP · pipeline · scripts · SLA.
12. **F3 CRM** — 15,5h paralelo F2 · painel de receita · campos + motivo de perda.
13. **F4 Estratégica** — 20h · playbook · comissão loja≠PAP · breakeven.
14. **F5 Instalação** — 22h presencial · roleplay · Mateus herda rotina.
15. **F6 Handover** — 10h · plano 60 dias · retenção/save · anti-regressão.
16. **Sucesso** — título **Receita Previsível** + 8 critérios observáveis.
17. **Recap** — Dores × Dados.
18. **Investimento** — R$ 89.635,85 → R$ 71.582,32 (condição comercial; não “desconto”).

## 4) Dados canônicos
- ~1.000 ativações/mês
- Meta Forms 1.269 leads · CPL R$ 7,51
- WhatsApp ~1 minuto
- Saídas: 415, 560, 614, 746, 715, 705 (total 3.755)
- 5 loja + 2 PAP · Mateus ~30 dias
- 12 semanas · 106,5h · 6 fases · dual-track
- Preço: R$ 89.635,85 → R$ 71.582,32
- Não mencionar sucessão/saída de sócios

## 5) Spec HTML mínima
- Self-contained (CSS + SVG + JS + logos base64)
- 1600×900 · fitDeck · `.slide.active`
- IBM Plex Sans/Mono
- Tokens: `#e50914` / `#280001` / `#ffebc8`
- Tipos: `red | dark | white`
- Componentes: eyebrow, leak-grid, map-grid, phase-rail, success-grid, price-layout, split-2

## 6) Como editar com IA (poucos tokens)
1. Não manda o HTML inteiro com base64.
2. Manda este documento + o slide alvo (“no slide 16, trocar X por Y”).
3. Preserve a ordem do arco.
4. Sucesso = título “Receita Previsível”.
5. Preço = “condição comercial”, nunca “desconto”.

## 7) Variante enxuta (13 slides)
Capa → Dor G6 → Vazando → Sangria → Dores×G6 → CRM → Visão → Produto → Fases → Sucesso → Personalização → Recap → Investimento  
Modelo: https://sarapizzico-hue.github.io/Sara/g6-modelo/

## 8) Prompt curto
```
Crie/edite deck HTML V4 standalone 1600×900 (IBM Plex, #e50914/#280001, fitDeck, N/TOTAL).
Tese: do comercial artesanal à máquina de receita. Capacidade instalada, não PDF.
Use a ordem e os dados canônicos deste documento.
Altere apenas o(s) slide(s) indicado(s). Preserve o restante.
```
