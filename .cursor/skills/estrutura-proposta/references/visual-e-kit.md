# Visual e kit

## Deck cliente canônico (travado)

Copiar CSS/JS da âncora do brief. Não inventar tema.

**Âncora viva no `main`:** `proposta-dental-muller.html` (mesmo sistema de `proposta-g6-modelo-13-slides.html` e `g6-modelo/`).

| Token | Valor |
|---|---|
| Canvas | 1600×900, scale no viewport (`#deck` + `--deck-scale`) |
| Título | IBM Plex Sans 600 · letter-spacing negativa |
| Mono / counter | IBM Plex Mono |
| Vermelho | `#e50914` |
| Dark | `#280001` → `#130001` |
| Cream | `#ffebc8` |
| Hero capa | radial `#ff5a2c` → `#e50914` → `#b00610` → `#280001` |
| Tipos de slide | `red` · `dark` · `white` (creme/`#f5f4f2` no default) |
| Eyebrow | pill uppercase, letter-spacing ~0.09em |
| Counter | `01 / N` canto |

Interação: setas, espaço, swipe, dots. Footer com wordmark V4. Self-contained.

HTML na **raiz** (GitHub Pages). Assets (logo, docx) em `assets/`. Não remover `.nojekyll`. Não mover arquivo publicado sem atualizar links.

## O que não copiar

| Origem | Por quê |
|---|---|
| `proposta-g6-internet.html` (Sora) | Conteúdo/preço ok; visual **não** é o padrão novo |
| `index.html` / `farmer-sara.html` | Hub interno dark + Inter |
| Modular / Martins / Motéis (Outfit + Plus Jakarta, full-bleed) | Só se o brief citar essa âncora |
| Isoluz (Manrope, palco 16:9 próprio) | Deck próprio; puxar **conteúdo** de EC+e-com+Growth, não o tema, salvo brief |

## Como editar com pouco token

1. Não manda o HTML inteiro com base64.
2. Manda a ficha (`references/brief.md`) + o slide alvo (“no slide 6, trocar X por Y”).
3. Preserve a ordem do arco.
4. Preço = “condição comercial”, nunca “desconto”.
5. Sucesso de EC G6 = título **Receita Previsível** se o brief for G6/EC profunda.

## Kit de arquivos

| Entrega | Nome | Quando |
|---|---|---|
| Proposta deck | `proposta-{slug-cliente}-{oferta}.html` | Sempre |
| Script pitch | `script-{slug-cliente}-pitch.html` | Sempre, salvo brief não |
| Word | `assets/{nome}.docx` | Se brief pedir |
| PDF/PPTX | só com pedido | Ciclo separado |
| Logo | `assets/` | Sempre que houver mídia |

Slug: minúsculas, hífen (`proposta-dental-muller.html`).

Não criar `-v2` / `-novo` sem arquivar a anterior no brief. G6 Social foi exceção (dois preços).

## Spec mínima de implementação

Ao criar arquivo novo:

1. Abrir a âncora visual do brief.
2. Duplicar o arquivo (ou copiar `<style>` + JS de navegação + chrome: topline, footer, counter, dots).
3. Trocar `<title>`, tese, slides. Manter classes `.slide.red|dark|white`, `.eyebrow`, `.lede`.
4. Conferir `fitDeck`: o retângulo 1600×900 cabe no viewport sem scroll interno do slide.
5. Abrir no browser: percorrer **todos** os slides com teclado, checar overflow de texto, preço visível, counter correto.

Componentes úteis (G6 / modelo): eyebrow, leak-grid, map-grid, phase-rail, success-grid, price-layout, split-2. Reuse se a âncora já tiver; não inventar um terceiro grid.

## Hub interno vs. deck

- Deck cliente: Colli 16:9, IBM Plex, vermelho V4
- Ferramenta Sara (pipe, farmer, score, plano): dark + Inter — **não misturar** e **não alterar** no fluxo de proposta
