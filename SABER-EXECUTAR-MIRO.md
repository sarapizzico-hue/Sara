# Saber → Executar · rotina Sara · Miro + Flow

Fonte: rotina da Sara Pizzico. Sistema: **Flow**. Cockpit = “Saber para Executar”.

## Ordem real (não inverter)
Cliente entra → **GC** → **Kick-off** (Sara entra com o time e se apresenta) → **5 semanas** (1 contato/semana para conversar e tirar dúvida) → **S3** o time apresenta o comercial → Sara + time montam o **pré-pit** (contexto + o que houve na EE, puxado no **Notebook LM**) → **Account Planning valida o produto** (NBO) → reunião para **finalizar a proposta com narrativa** → apresentam → **fecha em call** ou **follow** (marca devolutiva se precisar) → ganho **ou drop com motivo** para reabrir no momento certo.

## 5 semanas (manual no card do Cockpit)
Cada semana é um par **data + 1 linha**. Sem automação: a Sara marca o contato na mão.

| Semana | O que a Sara faz | Campos Flow |
|--------|------------------|--------------|
| 1 | Kick-off: se apresenta com o time. 1º contato. | `semana_1_data` · `semana_1_obs` |
| 2 | Pulso: conversa e dúvida. Sem preço. | `semana_2_data` · `semana_2_obs` |
| 3 | Depois do comercial: pré-pit + Notebook LM. | `semana_3_data` · `semana_3_obs` |
| 4 | Account Planning valida produto. Reunião de narrativa. | `semana_4_data` · `semana_4_obs` |
| 5 | Apresentam. Fecha em call / follow / drop. | `semana_5_data` · `semana_5_obs` |

`semana_ee` (select 1–5) continua só como filtro da semana atual. O histórico vive nos cinco pares.

S4/S5 escorregam se o pré-pit ou o AP não fechou. O contato semanal **não** para.

## Pré-pit (S3)
Contexto da conta + o que houve na EE. Fonte: Notebook LM.  
Ainda **não** valida produto. Libera o Account Planning.

Ferramenta: `documento-pre-pit.html`

## Account Planning (depois do pré-pit)
Valida o produto. Gera a **NBO**. Um foco: retenção · otimização de campanhas · novos produtos.

## Depois da call
- **Fecha em call** → win → kickoff do produto · `data_finalizacao` + `valor_fechado_tcv` / `valor_fechado_mrr`.
- **Follow** → mesmo dia; `follow_obs` + `termometro` + `data_devolutiva` se pediu tempo.
- **Drop** → `data_finalizacao` + `motivo_drop` + `abertura_futura`. Não apaga o contexto.

## Colunas extras no Flow (além do Cockpit)
Cadência: `data_entrada` · `data_gc` · `kickoff_sara` · `data_kickoff` · `semana_ee` · `semana_1..5_data/obs` · `notebook_lm_url` · `status_ciclo`  
Pré-pit: `prepit_status` · `prepit_url` · `o_que_houve_ee`  
AP/NBO: `foco_ap` · `nbo` · `nbo_step` · `valor_estimado_tcv` · `valor_estimado_mrr` · `nbo_status`  
Proposta: `proposta_url` · `valor_tcv` · `valor_mrr` · `forma_pgto` · `termometro` (Quente / Morno / Frio)  
Close: `data_reuniao` · `resultado_pitch` · `follow_obs` · `data_devolutiva` · `data_finalizacao` · `valor_fechado_tcv` · `valor_fechado_mrr` · `motivo_drop` · `abertura_futura`

### Onde vai o R$
1. **NBO** — `valor_estimado_tcv` / `valor_estimado_mrr` (hipótese, não inventar).
2. **Construção / Revisão** — `valor_tcv` / `valor_mrr` + `forma_pgto` (número do deck).
3. **Pitch** — os mesmos `valor_tcv` / `valor_mrr` (o que foi apresentado).
4. **Ganho** — `valor_fechado_tcv` / `valor_fechado_mrr` + `data_finalizacao`. Drop não leva valor fechado.

MVP na semana 1 da conta: GC, kick-off, `semana_1_data` + obs, próxima ação.

## Migração Cockpit → Expansão
Depois do AP validar o produto (NBO comercial). Pré-pit Pronto + narrativa pronta + gates ok.  
Retenção / otimização sem venda permanece no Cockpit.

## Pipe comercial (desenho para o Miro)
Modelo atual: NBO → Diagnóstico → Oportunidade → POC → Revisão → Reunião agendada → Reunião realizada → Follow → Ganho/Drop.

Proposto (só oportunidade):
Cockpit (GC → KO → 5 sem manuais → pré-pit S3 → AP)  
↓ NBO validada  
NBO → Construção da Proposta → Revisão da Proposta → Reunião Agendada → Reunião Realizada → Follow → Ganho/Drop

Sai do pipe: Diagnóstico, Oportunidade (redundante), Construção POC.
Página para copiar: `pipeline-comercial-miro.html` (Frame 2 = campos em cada etapa · Frame 4 = card de exemplo).
