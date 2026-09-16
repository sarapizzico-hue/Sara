# Saber → Executar · rotina Sara · Miro + Flow

Fonte: rotina da Sara Pizzico. Sistema: **Flow**. Cockpit = “Saber para Executar”.

## Ordem real (não inverter)
Cliente entra → **GC** → **Kick-off** (Sara entra com o time e se apresenta) → **5 semanas** (1 contato/semana para conversar e tirar dúvida) → **S3** o time apresenta o comercial → Sara + time montam o **pré-pit** (contexto + o que houve na EE, puxado no **Notebook LM**) → **Account Planning valida o produto** (NBO) → reunião para **finalizar a proposta com narrativa** → apresentam → **fecha em call** ou **follow** (marca devolutiva se precisar) → ganho **ou drop com motivo** para reabrir no momento certo.

## 5 semanas
| Semana | O que a Sara faz |
|--------|------------------|
| 1 | Kick-off: se apresenta com o time. 1º contato. |
| 2 | Pulso: conversa e dúvida. Sem preço. |
| 3 | Depois do comercial: pré-pit + Notebook LM. |
| 4 | Account Planning valida produto. Reunião de narrativa. |
| 5 | Apresentam. Fecha em call / follow / drop. |

S4/S5 escorregam se o pré-pit ou o AP não fechou. O contato semanal **não** para.

## Pré-pit (S3)
Contexto da conta + o que houve na EE. Fonte: Notebook LM.  
Ainda **não** valida produto. Libera o Account Planning.

Ferramenta: `documento-pre-pit.html`

## Account Planning (depois do pré-pit)
Valida o produto. Gera a **NBO**. Um foco: retenção · otimização de campanhas · novos produtos.

## Depois da call
- **Fecha em call** → win → kickoff do produto.
- **Follow** → mesmo dia; se pediu tempo, **data de devolutiva** no Flow.
- **Drop** → `motivo_drop` + `abertura_futura` (gatilho para reabrir). Não apaga o contexto.

## Colunas extras no Flow (além do Cockpit)
Cadência: `data_gc` · `kickoff_sara` · `semana_ee` · `contato_semanal` · `notebook_lm_url` · `status_ciclo`  
Pré-pit: `prepit_status` · `prepit_url` · `o_que_houve_ee`  
AP/NBO: `foco_ap` · `nbo` · `nbo_step` · `nbo_status` (Em contexto → Validada → Em proposta → Em pitch → Follow → Ganhou/Drop)  
Close: `resultado_pitch` · `data_devolutiva` · `motivo_drop` · `abertura_futura`

MVP na semana 1 da conta: GC, kick-off, semana EE, contato semanal, próxima ação.

## Migração Cockpit → Expansão
Depois do AP validar o produto (NBO comercial). Pré-pit Pronto + narrativa pronta + gates ok.  
Retenção / otimização sem venda permanece no Cockpit.

## Pipe comercial (desenho para o Miro)
Modelo atual: NBO → Diagnóstico → Oportunidade → POC → Revisão → Reunião agendada → Reunião realizada → Follow → Ganho/Drop.

Proposto (só oportunidade):
Cockpit (GC → KO → 5 sem → pré-pit S3 → AP)  
↓ NBO validada  
NBO → Construção da Proposta → Revisão da Proposta → Reunião Agendada → Reunião Realizada → Follow → Ganho/Drop

Sai do pipe: Diagnóstico, Oportunidade (redundante), Construção POC.
Página para copiar: `pipeline-comercial-miro.html`
