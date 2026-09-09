# Dinâmica de sala (Colli unificada)

Estudo do casco [ABC71 — deck unificado](https://relatorios.collieassociados.com/abc71-semana-4-deck-unificado/#slide-01).

**Pegar o palco e a dinâmica. Não pegar o conteúdo da conta** (SWOT, personas, SEO, Drawflow, semanas de projeto).

Informação do deck continua o arco das últimas propostas — ver `atos.md` e `modelo-proposta-esqueleto.html`.

Visual = vinho Colli + Montserrat + JetBrains Mono (`visual-e-kit.md`). Cream/IBM Plex da Muller fica como âncora antiga, só se o brief pedir.

## O que a dinâmica resolve na sala

A ABC71 não é um HTML que “vira página”. É um palco com **orientação permanente**: onde estou no arco, quanto falta, e como pular de bloco sem perder o cliente.

Isso casa com os três atos da proposta. Na ABC71 os blocos são Momento / Aquisição / Mídia / Plano. **Na nossa proposta os blocos são Narrativa / Dados / Produto / Decisão.**

## Pegar

| Peça | O que faz | Como adaptar |
|---|---|---|
| **Pílula de contexto** | Sempre visível: bloco + título do slide | `Narrativa` · tese da capa. Atualiza a cada `goTo` |
| **Nav por ato** | B1 B2 B3… saltam para o início do bloco | 4 botões: Narrativa · Dados · Produto · Decisão. No combo curto, Dados pode colapsar no Diagnóstico |
| **Hash `#slide-01`** | Link direto, ensaio, “volta nesse slide” | `history.replaceState` no `goTo`. Ler o hash no boot |
| **`window.goTo(i)`** | API estável (casco, PDF, ensaio) | Exportar `goTo`. Todo slide com `data-title` e `data-act` |
| **Barra de progresso** | Quanto da reunião já passou | Topo, 4px, gradiente vermelho. Dental Muller já tem na base — subir para o topo |
| **Home / End / F** | Início, fim, tela cheia | Além de setas, espaço, PageUp/Down (Muller já tem as setas) |
| **Chip de seção** | `01 Momento` no canto do slide | `02 Dados` — índice **local do ato**, não o número global |
| **Divisor de ato** | Subcapa: número grande + kicker + título + trilho do que vem | No combo (`modelo-proposta-esqueleto.html`): um divisor antes de Produto. Em EC 13–18: divisores antes de Dados, Produto e Decisão |
| **Profundidade sob demanda** | Card denso abre modal; Esc fecha; setas navegam **dentro** do modal, não o deck | Usar em 6 dores, detalhe de fase, ou “o que está incluso”. O slide da sala fica respirável |
| **`aria-label` + print 1600×900** | Acessível e exportável | Manifest `__DECK_EXPORT_MANIFEST__` se o brief pedir PDF depois |

## Não pegar

- Conteúdo, teses, números ou personas da ABC71
- Casco com **vários iframes** (44 slides) — um único HTML na raiz
- SWOT / TOWS / comitê de compra como arco de proposta
- Toolbar “PDF após aprovação” até o brief pedir PDF

## Contrato JS mínimo

Todo deck novo (a partir desta skill) deve cumprir:

```js
// Cada <section class="slide"> tem:
//   data-act="narrativa|dados|produto|decisao"
//   data-title="Capa"
//   aria-label="Capa — Do comercial artesanal à máquina de receita"

function goTo(index) {
  // 1. ativa o slide
  // 2. atualiza counter 01 / N  (JetBrains Mono)
  // 3. atualiza pílula: ato + data-title
  // 4. atualiza barra de progresso
  // 5. marca o botão de ato
  // 6. history.replaceState(null, '', '#slide-' + pad(index+1))
  // 7. document.title = pad(index+1) + ' / ' + pad(total) + ' — ' + cliente
}
window.goTo = goTo;

// Boot: se location.hash casar /slide-(\d+)/, abrir nesse índice
// Teclado: ← → espaço PageUp PageDown Home End
// F = fullscreen (se document.fullscreenEnabled)
// Se um dialog/modal estiver aberto: setas NÃO avançam o deck; Esc fecha
```

Nav persistente (fora do canvas 1600×900, como a ABC71 — não compete com o conteúdo):

- topo: progresso
- topo-esq: pílula de contexto
- baixo-esq: atos (Narrativa · Dados · Produto · Decisão)
- baixo-centro: ← `01 / 12` →
- baixo-dir: Tela cheia (PDF só se brief pedir)

No combo curto, os quatro atos ainda existem: Diagnóstico = Narrativa+Dados; cada produto = Produto; Investimento = Decisão. Os botões saltam para o **primeiro slide daquele ato**.

## Divisor (quando usar)

Slide `dark` ou `red`, sem cards. Função: respirar e anunciar o ato — não vender.

```
eyebrow / kicker    Ato 3 · Produto
título              O que fica instalado no time
sub                 Formato · fases · o que “pronto” significa
trilho              Diagnóstico → Arquitetura → CRM → Instalação → Handover
```

Um divisor não leva dado novo. Se o título do divisor for genérico o suficiente para qualquer empresa, reescrever com o vocabulário da conta.

## Modal de profundidade (quando usar)

Só quando o slide da sala ficaria ilegível: seis dores, quatro fases com encaixe, lista “incluso”.

- O slide mostra o mapa (títulos). O zoom mostra o texto.
- Esc / clique fora fecha e devolve o foco.
- Setas dentro do modal: próximo card. Setas com modal fechado: próximo slide.

Não usar modal para esconder o preço ou o reframe — esses ficam na sala.

## Relação com o que já temos

Dental Muller / G6 modelo já têm: `fitDeck` 1600×900, setas, espaço, Home/End, swipe, counter, dots, progresso na base.

Falta (passar a exigir): hash, pílula de ato, nav de atos, `data-act`, `window.goTo`, tela cheia, progresso no topo, divisor em decks longos, modal só se o conteúdo pedir.

Âncora de **layout + dinâmica**: `modelo-proposta-esqueleto.html` (palco ABC71). Âncora de **informação**: Muller / G6 / Isoluz + brief.
