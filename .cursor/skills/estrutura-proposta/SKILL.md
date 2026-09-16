---
name: estrutura-proposta
description: Monta e edita propostas comerciais V4 (deck HTML 16:9 Colli, dinâmica de sala, script de pitch). Use when the user pede proposta, deck, pitch, investimento, estruturação comercial, EQV, Growth, CRM, ou adaptação de um cliente novo; também ao estudar dinâmica Colli/ABC71 unificada para o palco (hash, atos, pílula) sem copiar conteúdo.
icon: book-open
color: red
---

# Estrutura de proposta V4

Skill da Sara / V4 Company. Ensina o agente a **construir proposta junto com a Sara**, não a inventar um deck genérico.

Frase-mestra: **nada solto no deck.** O cliente precisa se reconhecer (narrativa), ver a conta que dói (dados) e entender o que fica instalado no time (produto) — **antes** de ver preço.

Vende **capacidade instalada**. Não vende PDF, retainer vago nem cardápio de horas.

## Quando usar

- Pedido de proposta, deck, HTML comercial, pitch, investimento, condição comercial
- Conta nova ou adaptação de um caso âncora (G6, Isoluz, Dental Muller, Modular…)
- Edição de slide, tese, reframe, fases, preço, script de sala
- Dúvida de produto, visual, nome de arquivo ou “igual à última”

Não usar em pipe/safra/farmer (`index.html`, `farmer-sara.html`) — isso é ferramenta interna, não proposta.

## Como construir junto (ordem obrigatória)

Não pule etapas. Se o brief estiver incompleto, **pare e preencha a ficha com a Sara**. Não invente preço, escopo, tese ou visual.

1. Ler esta skill inteira.
2. Abrir `references/brief.md` e exigir o **brief mínimo**.
3. Escolher o **formato** em `references/formatos.md` (curto combo · EC profunda · 14 slides STEP).
4. Montar os **três atos + decisão** com `references/atos.md`.
5. Duplicar `modelo-proposta-esqueleto.html` (layout Colli + arco da proposta). Preencher os `.ph` com o brief. Não copiar conteúdo da ABC71.
6. Nomear produtos só pelo catálogo em `references/catalogo.md`.
7. Entregar o **kit**: `proposta-{cliente}-{oferta}.html` + `script-{cliente}-pitch.html` (salvo brief dizer não).
8. Rodar o checklist do final desta skill **antes** de declarar pronto.

Detalhe longo fica nas references. Carregue só o arquivo da etapa em que estiver.

## Brief mínimo (sem isso não começa)

Obrigatório:

| Campo | Exemplo de formato |
|---|---|
| Cliente | Dental Muller |
| Produto(s) do catálogo | EQV + Assessoria de Growth |
| Preços finais | âncora → condição única (ou A/B se brief pedir) |
| Referência visual | `modelo-proposta-esqueleto.html` (layout) · conteúdo no padrão Muller/G6 |
| Narrativa em 1 linha | Do [jeito atual] ao [estado instalado] |

Se faltar campo → listar o que falta e esperar. Não “chutar” ticket, desconto, prazo ou headline.

Campos que fecham a sala (preencher juntos quando existirem): reframe, evidências desta conta, vazamento com fonte, 4–6 dores, formato (prazo/carga/quem executa), critérios de “pronto”, por que agora.

Ficha completa: `references/brief.md`.

## Método — três atos + uma decisão

A lógica não muda. O que muda por conta é narrativa, dados e produto.

| Ato | Função na sala | Pergunta que o slide responde |
|---|---|---|
| **1. Narrativa** | Cliente se reconhece. Tese + reframe + visão + promessa | O que esta reunião decide? Vocês não têm X. Têm Y. |
| **2. Dados** | A conta que explica a estagnação. Entrada que já funciona → vazamento → pergunta que o CRM não responde | Quanto está vazando? |
| **3. Produto** | O que instala. Formato, escopo, fases/ciclo, papéis, “pronto” | O que fica rodando no time? |
| **Fecho** | Recap dores+dados → **um preço** → por que agora → próximo passo | Quanto — e o que fecha hoje? |

Na sala a ordem é: narrativa (capa/contexto) → dados → narrativa de solução → produto → recap → investimento. **Não abrir preço antes da conta e do encaixe.**

Teste da narrativa: se apagar o nome do cliente e o deck ainda servir para qualquer empresa, está genérico demais.

## Escolher o formato

Decida pelo produto, não por gosto de layout.

| Situação | Formato | Âncora |
|---|---|---|
| Combo de produtos (EQV+Growth, EC+e-com+Growth) | **Curto** 7–10 slides no layout `modelo-proposta-esqueleto.html` | Informação: `proposta-dental-muller.html`, `proposta-isoluz.html` |
| EC de instalação (fases, dual-track, handover) | **EC profunda** 13–18 slides: capa → dor → vazamento → dores×produto → visão → formato → fases → pronto → recap → investimento | `ESTRUTURA-ENTREGA-G6.md`, `g6-modelo/`, `proposta-g6-estruturacao-comercial.html` |
| Precisa de STEP / papéis / ondas / pedidos da semana | **Expandir o combo** (não o esqueleto de 8) | Brief pede STEP; copiar ciclo + papéis da Hanaro / Preço Popular |

Slides vazios não entram. Gate da sala (objeções, “faz sentido?”) fica no **script**, não no deck do cliente.

Mapa de cada slide: `references/formatos.md`. Dinâmica (hash, pílula, atos): `references/dinamica.md`.

## Visual travado

Híbrido travado nesta skill:

- **Layout / palco:** sistema Colli da ABC71 unificada (vinho, Montserrat, JetBrains Mono, slide-head, chip de seção, divisor, chrome de sala). Âncora: `modelo-proposta-esqueleto.html`
- **Informação:** arco das últimas propostas (Muller / G6 / Isoluz) — tese Do X ao Y, reframe, KPIs desta conta, 1 slide por produto do catálogo, ciclo, critérios de pronto, um preço. **Não** SWOT, personas, SEO ou semanas de projeto da ABC71

- Deck **1600×900** com `fitDeck`
- Self-contained: **um** HTML na raiz (sem casco de iframes)
- Navegação: setas / espaço / Home / End, hash `#slide-01`, pílula **ato + título**, nav de atos, `window.goTo`, tela cheia
- HTML na **raiz** (GitHub Pages). Não remover `.nojekyll`

Não copiar: visual Sora da G6 internet; tema dark Inter do hub Farmer; cream/IBM Plex da Muller **salvo o brief pedir essa âncora antiga**; conteúdo da ABC71.

Spec: `references/visual-e-kit.md`. Dinâmica: `references/dinamica.md`.

## Produto e preço

Nomes canônicos (não renomear por cliente):

**EC · EQV · Assessoria de Growth · CRM · Social Media · SDR IA · E-commerce B2B**

- Core resolve a dor #1. Essencial (CRM, retenção, dual-track) entra como o que fecha 100% — não como “se quiser a gente também faz”
- Um preço. Desconto já embutido. Chamar de **condição comercial**, nunca “desconto”
- Agente não calcula nem arredonda preço. Valor vem do brief
- A/B de pacote só se o brief pedir
- Critérios de sucesso são observáveis (pipeline usado, ritual acontecendo). Não são garantia de ROI/receita

Catálogo, combos e “onde puxar”: `references/catalogo.md`.

## Script de pitch (sempre, salvo brief dizer não)

Arquivo: `script-{cliente}-pitch.html`

- Conversa, não monólogo
- Perguntas **abertas** — evitar “faz sentido?”
- Objeções **antes** do preço (escopo, time, timing, outro decisor)
- Gate de avanço antes de abrir valor (lição ALNV)
- Um preço / uma decisão
- Pedir a venda + próximo passo concreto (assinatura, acessos, kickoff)
- Âncora de tom: `script-g6-pitch.html`, `script-isoluz-pitch.html`, `script-mbflex-pitch.html`

## Arquivos

| Entrega | Nome | Quando |
|---|---|---|
| Deck | `proposta-{cliente}-{oferta}.html` | Sempre |
| Script | `script-{cliente}-pitch.html` | Sempre, salvo brief não |
| Word/PDF/PPTX | `assets/` ou raiz se já for padrão do deal | Só se brief pedir — ciclo separado, sem refazer o deck |
| Logo | `assets/` | Sempre que houver mídia |

Não criar `-v2` / visual paralelo sem arquivar a anterior no brief.

Editar com IA: não mande o HTML inteiro com base64. Mande a ficha + “no slide N, trocar X por Y”. Preserve o arco.

## Regras da sala (fazer / evitar)

**Fazer**

- Nomear a dor com o vocabulário do cliente
- Mostrar o que já funciona, depois o vazamento
- Encaixar cada fase/produto **neste** cliente (canal, gestor, ritual)
- Critério de “pronto” visível para o sponsor
- Recap dores + dados imediatamente antes do preço
- Um preço · uma decisão

**Evitar**

- Deck que serve para qualquer empresa se apagar o logo
- Abrir investimento antes da conta e do encaixe
- Produto como lista de entregáveis sem instalação
- Várias formas de pagamento na mesma tela
- Upsell “se quiser a gente também faz”
- Projeção ou case soltos, sem amarrar na tese
- Número sem fonte
- Prometer taxa de case maduro como se fosse desta conta
- “Igual à G6” sem dizer qual arquivo

## Checklist antes de entregar

- [ ] Brief completo (cliente, produtos, preços, visual, tese)
- [ ] Tese no formato Do X ao Y, no idioma da conta
- [ ] Reframe: não é o sintoma, é a causa raiz
- [ ] Pelo menos um número desta operação, com fonte
- [ ] Cada produto do brief tem bloco próprio e nome de catálogo
- [ ] Recap ou ciclo fecha **antes** do slide de investimento
- [ ] Um preço (ou A/B explícito no brief) · condição comercial, não desconto
- [ ] Próximo passo concreto
- [ ] Visual 1600×900 Colli (`modelo-proposta-esqueleto.html`), não reinventado
- [ ] Dinâmica: hash `#slide-01`, pílula de ato, `data-act` + `window.goTo`
- [ ] `proposta-*.html` na raiz + script de pitch
- [ ] Hub Farmer / pipe **não** alterados neste fluxo
