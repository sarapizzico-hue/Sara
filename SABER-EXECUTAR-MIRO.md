# Saber → Executar · processo para Miro + Flow

Material de alinhamento (Erick Cardoso × Sara Pizzico).  
Sistema de registro: **Flow**. Cockpit = “Saber para Executar”.  
Não é catálogo: Account Planning gera **NBO**; NBO alimenta o **pipe comercial**.

## 1) Tese
A base ativa não vira pitch por feeling.  
O time planeja a conta → escolhe a Next Best Offer → registra no Flow → só então pré-pit, pitch e pipe de Expansão.

## 2) Arco (copiar no Miro, 8 frames)
Base ativa → Account Planning → NBO → Cockpit Saber p/ Executar (Flow) → Documento de Pré-pit → Pitch da semana → Pipe Expansão (ou permanece no Cockpit) → Executar, com Pipe Review por cima.

## 3) Account Planning
Coordenado pelo time. Um foco por ciclo (não os três ao mesmo tempo):

| Foco | Quando | Output típico |
|------|--------|----------------|
| Retenção | Saúde amarela/vermelha, risco de churn, CSAT baixo | NBO de save / reengajar — **não** migra para Expansão |
| Otimização de campanhas | Entrega estável, mídia/CRM já rodando, gap de performance | NBO de ajuste; só vira pipe se for produto/upgrade |
| Novos produtos | Gates ok + tese + WTP | NBO comercial → **migra** para pipe de Expansão |

## 4) Next Best Offer (NBO)
1 frase: a próxima oferta **desta** conta, agora.  
Campos: texto · STEP (Ter/Saber/Executar/Combo) · valor estimado · timing · o que **não** oferecer · status.

Status: Hipótese → Validada → Em pré-pit → Em pitch → No pipe → Ganhou / Drop.

## 5) Cockpit “Saber para Executar” (Flow)
Acompanha execução + base ativa + pipe comercial.  
Registra **métricas** (CSAT, NPS, lead score, % Flow preenchido, projeção) e **projetos** (NBO, pré-pit, pitch, modo de entrega).

Card no Cockpit ≠ card no pipe de Expansão. Ligação obrigatória (ID / URL).

## 6) Colunas adicionais no Flow (pedido Erick)
Além do que já existe no Cockpit E.E (cliente, squad, status, TCV, temperatura, pitch, obs):

**Account Planning / NBO**  
`foco_ap` · `nbo` · `nbo_step` · `nbo_valor` · `nbo_timing` · `nbo_nao_oferecer` · `nbo_status`

**Pré-pit**  
`prepit_status` · `prepit_url` · `modelo_negocio` · `publico_icp` · `alinhamentos_pendentes`

**Comercial**  
`lead_score` · `gates_ok` · `pitch_semana` · `projecao_cenario` · `tcv_nbo` · `origem_card` · `destino_pipe`

**Entrega**  
`modo_execucao` (Padrão · Assistida · Canal específico) · `canal_especifico` · `diagnostico_antecipado` · `estrategia_entrega`

**Higiene (Pipe Review)**  
`flow_preenchimento` (%) · `proxima_acao` · `dono` · `data_proxima`

## 7) Documento de Pré-pit (Sara)
Consolidar **antes** da proposta final. Sem pré-pit “Pronto”, não agenda pitch de preço.

Blocos:
1. Contexto da conta (contrato atual, semana EE, saúde, expectativa)
2. Modelo de negócio (como fatura, canais, unidades, margem)
3. Público / ICP
4. Diagnóstico (gargalo #1, tese, o que não oferecer)
5. NBO + STEP + economia (tier, margem, BE)
6. Alinhamentos necessários (decisor, budget, ops/capacidade, canal, case)
7. Roteiro do pitch (abertura, prova, oferta, ask)
8. Riscos e plano B (execução assistida / canal específico)

Ferramenta: `documento-pre-pit.html`

## 8) Migração Cockpit SABER → pipe de Expansão
**Migra** quando NBO é comercial (novo produto / upgrade / Executar), gates ok, lead score ≥ 70 (ou Plano S3 fechado), pré-pit Pronto, timing Agora ou 30d.

**Não migra** (fica no Cockpit): retenção, otimização sem venda, score < 70, gate falhou, pré-pit incompleto.

Como: clonar/mover no Flow · manter ID de origem · destino = Expansão Executar **ou** Expansão Ter/Saber · status origem = Migrado · não apagar histórico.

## 9) Pipe Review (coordenação)
Rituais para o Flow estar **100% preenchido**, revisar projeções, alinhar pitches da semana e antecipar diagnósticos + estratégia de entrega.

- Diário 15 min: P0, contratos na rua, silêncio
- Semanal 45–60 min: preenchimento Flow · NBO · pré-pit · pitches · projeção chão vs P/R/O · modo de execução
- Mensal: QBR leve / portfolio — retenção vs expansão

Regra: card incompleto não entra em projeção nem em pauta de pitch.

## 10) Gates inegociáveis
Decisor na sala + pago/pagando. Sem isso: diagnostica, **não apresenta preço**.
