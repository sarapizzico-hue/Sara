# Visual e kit

## Deck cliente canônico (travado)

**Layout** puxa o palco Colli da ABC71 unificada. **Informação** puxa o arco das últimas propostas (Muller / G6 / Isoluz).

Âncora para duplicar: `modelo-proposta-esqueleto.html`.

| Token | Valor |
|---|---|
| Canvas | 1600×900, `fitDeck` (`translate(-50%,-50%) scale`) |
| Display | Montserrat 700/800 (títulos). Corpo 400/500/600. *Não* usar 800 no texto de card |
| Métricas / counter | JetBrains Mono 500/700, tabular-nums |
| Capa | `#90191e` → `#7c191d` → `#65181c` |
| Fundo slide | `#351010` → `#290c0c` → `#210909` + grid 80px |
| Vermelho | `#ef3340` / `#ff6b73` |
| Âmbar / verde | `#f4b942` / `#58d68d` (EE e condição) |
| Tipos | `cover` · (conteúdo) · `divider` |
| Chrome no slide | `topbar` 7px · `slide-head` + `section-chip` · `footer` 42px |
| Chrome na sala | progresso topo, pílula, nav de atos, ← 01/N →, tela cheia |

Self-contained — **um arquivo**, sem iframes. Fontes via Google Fonts (não copiar os TTF da ABC71).

HTML na **raiz** (GitHub Pages). Assets em `assets/`. Não remover `.nojekyll`.

## O que copiar de onde

| Camada | Fonte | O quê |
|---|---|---|
| Palco, capa, divisor, cards, métricas, nav | `modelo-proposta-esqueleto.html` (sistema ABC71) | CSS + chrome |
| Tese, reframe, KPIs, produtos, ciclo, preço | Brief + Muller / G6 / Isoluz | Texto desta conta |
| Nomes de produto | catálogo da skill | EC · EQV · Growth · CRM · Social · SDR IA · E-com B2B |

## O que não copiar

| Origem | Por quê |
|---|---|
| Conteúdo da ABC71 (SWOT, personas, SEO, semanas) | Outro tipo de peça — não é proposta comercial |
| Iframes / 44 slides da unificada | Overkill; um HTML só |
| TTF hospedados na ABC71 | Usar Google Fonts |
| `proposta-g6-internet.html` (Sora) | Visual antigo |
| `index.html` / `farmer-sara.html` | Hub interno dark + Inter |
| Cream / IBM Plex da Muller | Só se o brief pedir a âncora antiga |
| Modular / Martins (Outfit, full-bleed) | Só se o brief citar |

## Como editar com pouco token

1. Duplicar `modelo-proposta-esqueleto.html` → `proposta-{cliente}-{oferta}.html`.
2. Trocar só os `.ph` e os títulos. Não reabrir o CSS.
3. Preserve o arco: capa → diagnóstico → (divisor) → produto(s) → ciclo → pronto → investimento.
4. Preço = “condição comercial”, nunca “desconto”.
5. Número sem fonte não entra.

## Kit de arquivos

| Entrega | Nome | Quando |
|---|---|---|
| Esqueleto (layout + arco) | `modelo-proposta-esqueleto.html` | Duplicar sempre |
| Proposta deck | `proposta-{slug-cliente}-{oferta}.html` | Sempre |
| Script pitch | `script-{slug-cliente}-pitch.html` | Sempre, salvo brief não |
| Word / PDF / PPTX | `assets/` | Só se brief pedir |

## Spec mínima

1. Duplicar `modelo-proposta-esqueleto.html`.
2. Cada slide: `data-act`, `data-title`, `aria-label`.
3. Cumprir `dinamica.md` (`goTo`, hash, pílula, nav de atos).
4. `fitDeck` sem scroll interno.
5. Browser: teclado, `#slide-03`, saltar pelos atos, overflow, preço visível.

## Hub interno vs. deck

- Deck cliente: vinho Colli + Montserrat (`modelo-proposta-esqueleto.html`)
- Ferramenta Sara: dark + Inter — não misturar e não alterar no fluxo de proposta
