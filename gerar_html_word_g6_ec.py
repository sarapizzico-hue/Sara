#!/usr/bin/env python3
"""Gera HTML no formato Microsoft Word (abre direto no Word → Salvar como .docx)."""

from html import escape
from pathlib import Path

OUT_HTML = Path("/workspace/documento-g6-internet-estruturacao-comercial-word.html")
OUT_DOC = Path("/workspace/documento-g6-internet-estruturacao-comercial.doc")


def td(text, bold=False):
    b = "font-weight:700;" if bold else ""
    return f'<td style="{b}">{escape(text)}</td>'


def table(headers, rows):
    head = "".join(f"<th>{escape(h)}</th>" for h in headers)
    body = []
    for row in rows:
        cells = []
        for i, val in enumerate(row):
            cells.append(td(val, bold=i == 0))
        body.append("<tr>" + "".join(cells) + "</tr>")
    return (
        '<table class="grid" border="1" cellspacing="0" cellpadding="0">'
        f"<thead><tr>{head}</tr></thead><tbody>{''.join(body)}</tbody></table>"
    )


def h1(text):
    return f"<h1>{escape(text)}</h1>"


def h2(text):
    return f"<h2>{escape(text)}</h2>"


def p(text, cls=""):
    c = f' class="{cls}"' if cls else ""
    return f"<p{c}>{escape(text)}</p>"


def quote(text):
    return f'<p class="quote">{escape(text)}</p>'


def callout(label, text):
    return (
        '<table class="callout" border="0" cellspacing="0" cellpadding="0"><tr><td>'
        f'<p class="callout-label">{escape(label.upper())}</p>'
        f"<p>{escape(text)}</p>"
        "</td></tr></table>"
    )


def ul(items):
    lis = "".join(f"<li>{escape(i)}</li>" for i in items)
    return f"<ul>{lis}</ul>"


def phase(title, intro, entregas, encaixe):
    return (
        h2(title)
        + p(intro, "justify")
        + '<p class="label">Entregas</p>'
        + ul(entregas)
        + f'<p class="muted"><i>{escape(encaixe)}</i></p>'
    )


parts = []

parts.append(
    """<!--[if !mso]><!-->
<div class="open-bar">
  <b>Para abrir no Word:</b> clique com o botão direito neste arquivo →
  <b>Abrir com → Microsoft Word</b>.
  Depois: <b>Arquivo → Salvar como → Documento do Word (.docx)</b>.
  No Windows, o arquivo <code>.doc</code> ao lado abre direto no Word com um duplo clique.
</div>
<!--<![endif]-->"""
)

# Capa
parts.append(
    """<table class="cover" border="0" cellspacing="0" cellpadding="0"><tr><td>
<p class="cover-kicker">G6 INTERNET  ×  V4 COMPANY</p>
<p class="cover-title">Estruturação Comercial</p>
<p class="cover-sub">Do comercial artesanal à máquina de receita</p>
<p class="cover-lead">Documento consolidado da proposta apresentada e da reunião comercial de 14 de agosto de 2026 — diagnóstico, escopo, fases, investimento, decisões e próximos passos.</p>
</td></tr></table>"""
)

parts.append(
    table(
        ["Campo", "Informação"],
        [
            ["Cliente", "G6 Internet"],
            ["Produto", "Estruturação Comercial (Sales V4) — projeto fechado com implementação assistida"],
            ["Duração", "12 semanas  ·  106,5 horas  ·  dual-track loja × PAP"],
            ["Investimento", "R$ 71.582,32  (referência do produto: R$ 89.635,85)  ·  12× recorrente"],
            ["Reunião", "14/08/2026  ·  [Comercial] V4 Company :::: G6 Internet"],
            ["Participantes", "Laurelise Santos (G6)  ·  Rafaela Tolomeu Messias  ·  Sara Pizzico  ·  Denis Orosco"],
            ["Pendência", "Validação da proposta com o sócio Geraldo antes do fechamento"],
            ["Fontes", "Deck de 18 slides da Estruturação Comercial + ata/transcrição da reunião (Google Docs)"],
        ],
    )
)

parts.append(
    callout(
        "Tese da apresentação",
        "A G6 não tem problema de lead. Tem um comercial reativo sem sistema. "
        "Entra cerca de 1.000 clientes por mês e a base quase não cresce. "
        "A Estruturação Comercial instala a fundação — processo, CRM em uso, gestão e rotina — "
        "para o novo gestor (Mateus) herdar uma máquina pronta, e não improvisar processo, peça e meta ao mesmo tempo.",
    )
)

parts.append(h1("Sumário"))
parts.append(
    ul(
        [
            "1. Resumo executivo",
            "2. Contexto da G6 — o que foi diagnosticado",
            "3. A conta do vazamento — dados da operação e da reunião",
            "4. Dores × o que a Estruturação Comercial fecha",
            "5. Visão estratégica e promessa do projeto",
            "6. Formato, escopo e o que fica instalado",
            "7. As seis fases (com entregas e encaixe G6)",
            "8. Três pilares: sistemas, processos e pessoas",
            "9. Critérios de Receita Previsível",
            "10. Investimento, condição comercial e viabilidade",
            "11. Ata da reunião de 14/08/2026 — decisões e próximos passos",
            "12. Pauta complementar: lançamento G6 Móvel / Chip",
            "13. Como levar este caso ao Geraldo",
            "14. Fontes e materiais de apoio",
        ]
    )
)

parts.append(h1("1. Resumo executivo"))
parts.append(
    p(
        "Em 14 de agosto de 2026, a V4 Company apresentou à G6 Internet a proposta de "
        "Estruturação Comercial: um projeto de 12 semanas, com 106,5 horas, para construir "
        "e instalar a fundação comercial da operação — loja e porta a porta (PAP) — com "
        "diagnóstico presencial, arquitetura, CRM como painel de receita, camada estratégica, "
        "imersão/treino e handover de 60 dias.",
        "justify",
    )
)
parts.append(
    p(
        "A leitura central: a G6 já gera demanda (cerca de 1.000 ativações/mês, WhatsApp em ~1 minuto, "
        "Meta Forms com 1.269 leads no ano e CPL de R$ 7,51). O que falta é sistema. O CRM existe, "
        "mas não governa. Loja e PAP vendem sem padrão. A meta desce quando aperta. Ativação, save "
        "e retenção ficam fora do processo. Entre janeiro e junho de 2026, a operação somou 3.755 cancelamentos.",
        "justify",
    )
)
parts.append(
    p(
        "Na reunião, Rafaela traduziu esse vazamento em conta de gestão: cada cliente custa cerca de "
        "R$ 800 (CAC + instalação) e precisa ficar no mínimo 8 meses para se pagar. Sobre os 3.755 "
        "cancelamentos, a conta de padaria aponta ~R$ 3 milhões já investidos para adquirir/instalar "
        "e ~R$ 4 milhões de receita potencial deixada na mesa — cerca de R$ 7 milhões em seis meses. "
        "O investimento de marketing com a V4 + a Estruturação Comercial não chega a 10% desse vazamento.",
        "justify",
    )
)
parts.append(
    p(
        "Laurelise (Laura) reconheceu o encaixe com a dor combinada na conversa anterior, afirmou "
        "vontade de fazer o projeto e deixou pendente a validação com o sócio Geraldo. O cronograma "
        "de 12 semanas e o diagnóstico in loco (18h) foram alinhados. O valor da condição comercial "
        "G6 é R$ 71.582,32, parcelável em 12 recorrências (sem cartão), para viabilizar o start junto "
        "com o contrato de marketing já em andamento.",
        "justify",
    )
)
parts.append(h2("O que ficou alinhado vs. o que ficou pendente"))
parts.append(
    table(
        ["Status", "Ponto"],
        [
            ["Alinhado", "Roteiro de 12 semanas e as seis fases (diagnóstico → handover)"],
            ["Alinhado", "Diagnóstico operacional presencial, com 18h provisionadas"],
            ["Alinhado", "Estruturação Comercial é custo único de implementação (não se repete)"],
            ["Alinhado", "Acompanhamento comercial após o projeto entra no executado de marketing, sem custo extra"],
            ["Alinhado", "Identidade do G6 Móvel segue o padrão visual da G6 Internet"],
            ["Pendente", "Validação da proposta e do investimento com Geraldo (Laura leva na terça)"],
            ["Pendente", "Kickoff com estrutura detalhada do plano (Denis)"],
            ["Pendente", "Reunião específica de lançamento móvel com Duarte / time de execução (Natã agenda)"],
        ],
    )
)

parts.append(h1("2. Contexto da G6 — o que foi diagnosticado"))
parts.append(quote("Headline do deck: “Vocês não têm problema de lead. Têm um comercial reativo sem sistema.”"))
parts.append(
    p(
        "A proposta foi montada a partir do contexto que Laura passou, das dores da operação e da "
        "movimentação do Mateus (novo gestor comercial, entrada em cerca de 30 dias). O mundo ideal "
        "apresentado: o gestor chega no meio do processo, treina com a V4 e herda uma máquina já "
        "estruturada — em vez de dividir tempo entre estruturar, trocar peça e bater meta.",
        "justify",
    )
)
parts.append(h2("Quatro falhas de sistema"))
parts.append(
    table(
        ["Sintoma", "O que isso significa na G6"],
        [
            ["Entra ~1.000 / sai ~800", "Volume alto, retenção frágil. O crescimento líquido vaza todo mês. A batalha de aquisição não vira base."],
            ["CRM existe, mas não governa", "Sem motivo de perda, sem próximo passo, sem previsibilidade de pipeline. Dado solto. Não é painel de receita."],
            ["Meta vira comodidade", "Quando aperta, a meta desce para o time “bater” e receber comissão. Cultura de acomodação cascateada pela liderança."],
            ["Loja e PAP sem padrão", "5 internos + 2 porta a porta. Cada canal vende de um jeito, sem playbook, sem deparo, sem objeção mapeada."],
        ],
    )
)
parts.append('<p class="label">Leitura do deck</p>')
parts.append(quote("Demanda entra. Processo não segura. Receita escapa entre ativação, follow-up e retenção."))
parts.append(h2("O que a operação já tem — e o que ainda não segura"))
parts.append(
    ul(
        [
            "Time comercial de 7 pessoas, mais o Mateus na chegada.",
            "Loja funcionando de um jeito e PAP de outro — nenhum dos dois com processo e meta claros.",
            "Porta a porta sem deparo: região visitada, oportunidades mapeadas, o que trouxe, o que não trouxe, objeções.",
            "Cliente oculto / mensagens de atendimento longas demais para um produto de alta intenção (primeiro plano de internet).",
            "Base instalada sem jornada de upsell/upgrade — o time olha só “vendeu, acabou”.",
            "Casos de “miséria”: cliente instalado que não pagou / não ativou — comissão apagada, receita nenhuma.",
            "Cultura de atendimento forte (calor humano, proximidade). O comercial, porém, herdou o ritmo do atendimento e ficou reativo — falta a gana de meta sem perder qualidade de conversa.",
            "Chegada do chip/móvel: se o time não estiver treinado para vender para a base, a G6 gasta de novo para ofertar um produto que a carteira já poderia absorver.",
        ]
    )
)
parts.append(
    callout(
        "Por que o diagnóstico precisa ser presencial",
        "Rafaela relatou que, de fora, a leitura era “ajuste de time”. Presencialmente, viu o perfil "
        "da casa: atendimento excelente, ritmo mais calmo, comercial seguindo a mesma cadência do "
        "atendimento. Sem estar in loco, o Denis estruturaria um plano genérico. A Fase 1 (18h) "
        "existe para ele ver cultura, resistência à mudança, loja, PAP, cidade e o dia a dia — "
        "não só entrevistar.",
    )
)

parts.append(h1("3. A conta do vazamento — dados da operação e da reunião"))
parts.append(quote("A G6 tem capacidade de vender. O que a operação não tem é capacidade de segurar — e de medir por que perde."))
parts.append(h2("3.1 Números do deck"))
parts.append(
    table(
        ["Lado", "Indicador", "Dado"],
        [
            ["Entradas", "Ativações / mês", "~1.000 (número citado pela operação)"],
            ["Entradas", "Meta Forms (ano)", "1.269 leads  ·  CPL R$ 7,51"],
            ["Entradas", "WhatsApp", "Resposta em cerca de 1 minuto"],
            ["Saídas", "Cancelamentos jan–jun/2026", "3.755"],
            ["Saídas", "Janeiro", "415"],
            ["Saídas", "Fevereiro", "560"],
            ["Saídas", "Março", "614"],
            ["Saídas", "Abril", "746"],
            ["Saídas", "Maio", "715"],
            ["Saídas", "Junho", "705"],
        ],
    )
)
parts.append(
    callout(
        "Pergunta que o deck deixou na mesa",
        "Qual é a taxa real de churn da G6? Para calcular, falta a carteira ativa no início de cada mês "
        "e o cancelamento por coorte, motivo, cidade e tempo de casa. Esse dado ainda não existe de "
        "forma conciliada — e sem ele, toda decisão de investimento é chute. A Estruturação Comercial "
        "instala exatamente essa inteligência no sistema.",
    )
)
parts.append(h2("3.2 Conta de padaria apresentada na reunião (Rafaela → Laura / Geraldo)"))
parts.append(
    p(
        "Esta conta não está no slide de vazamento do deck; foi falada na reunião para traduzir o "
        "impacto financeiro ao Geraldo. É estimativa de gestão, não conciliação contábil — e foi "
        "apresentada assim, de propósito, porque o tracking fino ainda não existe.",
        "muted",
    )
)
parts.append(
    table(
        ["Premissa", "Número", "Leitura"],
        [
            ["CAC + instalação por cliente", "≈ R$ 800", "Custo médio para colocar o cliente dentro"],
            ["Payback mínimo", "8 meses", "Tempo para o cliente se pagar (zero a zero)"],
            ["Cancelamentos em 6 meses", "3.755", "Jan 415 → jun 705, série crescente no 1º semestre"],
            ["Investimento queimado (3.755 × R$ 800)", "≈ R$ 3 milhões", "Dinheiro já posto em aquisição + instalação"],
            ["Receita potencial deixada na mesa", "≈ R$ 4 milhões", "O que esses clientes deixariam ao ano"],
            ["Vazamento total estimado em 6 meses", "≈ R$ 7 milhões", "Investimento perdido + receita não realizada"],
            ["Contrato V4 marketing (12 meses)", "≈ R$ 134 mil", "12× R$ 11.180 — já fechado"],
            ["Meta de crescimento citada (Geraldo)", "5% ao ano", "Leitura: ~R$ 3 mi/ano a mais, ~R$ 300 mil/mês"],
        ],
    )
)
parts.append(h2("3.3 O “mínimo que já paga o projeto” (cenário 2%)"))
parts.append(
    p(
        "Rafaela ancorou o ROI sem superlativo: se a estruturação melhorar só 2% da perda — “que não é nada” "
        "frente ao que o método costuma mover — a G6 puxaria cerca de 75 clientes. Conta apresentada:",
        "justify",
    )
)
parts.append(
    table(
        ["Item", "Valor"],
        [
            ["Clientes retidos no cenário 2%", "≈ 75"],
            ["Custo já investido nesses 75 (não perdido)", "≈ R$ 60.000"],
            ["LTV / o que pagam em 12 meses", "≈ R$ 90.000"],
            ["Receita retida (60 + 90)", "≈ R$ 150.000"],
            ["Leitura", "Paga ~2 meses do contrato V4 + parte da Estruturação Comercial"],
        ],
    )
)
parts.append(
    p(
        "A tese: se a G6 não arruma o comercial agora, continua perdendo no longo prazo, porque o gasto "
        "para colocar o cliente dentro exige no mínimo 8 meses de permanência. Sem estrutura, nem se "
        "sabe quantos dos 415 que saíram em janeiro já tinham se pago.",
        "justify",
    )
)

parts.append(h1("4. Dores × o que a Estruturação Comercial fecha"))
parts.append(
    table(
        ["Dor da G6", "O que o produto fecha"],
        [
            ["Processo comercial inconsistente — loja e PAP vendem sem padrão; não há jornada nem critério de etapa.", "Arquitetura dual-track loja × PAP, com critérios de etapa, papéis e jornada comercial alvo."],
            ["Cada vendedor vende de um jeito — performance depende de pessoa, não de método.", "Scripts, cadências, BPMN e matriz CHA. Treino com roleplay, não PDF na gaveta."],
            ["CRM não representa a realidade — pipeline não reflete ativação, perda nem follow-up.", "CRM como painel de receita: campos obrigatórios, motivos de perda, próxima ação, dashboard."],
            ["Gestor não conduz reunião comercial — Mateus precisa herdar rituais prontos.", "Rituais diário / semanal / pipeline review + onboarding do gestor na Fase 4 e 5."],
            ["Sem cadência e passagem de bastão — vendeu e acabou; ativação, save e retenção soltos.", "SLA marketing ↔ comercial ↔ operação; rotinas de ativação, follow-up e save no processo."],
            ["Treino não vira rotina — artefato sem instalação assistida regride.", "Fase 5 presencial (22h): imersão, roleplay, checklist semanal e correção de desvios."],
        ],
    )
)

parts.append(h1("5. Visão estratégica e promessa do projeto"))
parts.append(quote("A ordem certa: fundação antes de pressão de meta."))
parts.append(
    table(
        ["#", "Pilar", "O que instala"],
        [
            ["01", "Sistema — máquina comercial", "Processo, CRM e rituais transformam esforço individual em operação previsível."],
            ["02", "Gestão — gestor executa", "Com a casa estruturada, a liderança cobra comportamento e conversão — não improvisa processo."],
            ["03", "Receita — menos vazamento", "Ativação, follow-up e retenção entram no método. O CAC deixa de queimar no escuro."],
        ],
    )
)
parts.append(callout("Frase de ouro do deck", "Sem sistema, meta vira pressão. Com sistema, meta vira consequência."))
parts.append(h2("Promessa"))
parts.append(
    p(
        "Em um projeto fechado, a V4 constrói e instala a fundação comercial da G6. "
        "Processo, CRM em uso, scripts, cadências, indicadores e rituais de gestão — "
        "operando de verdade pelo time de loja e PAP.",
        "justify",
    )
)
parts.append('<p class="strong-red">Não é entregar PDF. É deixar a operação rodando.</p>')
parts.append(
    p(
        "Na reunião, Rafaela reforçou o mesmo ponto: a ideia não é auditar, entregar um documento "
        "e deixar a G6 aplicar. É auditar, implementar, treinar, criar rotina, garantir que está "
        "rodando para o Mateus — e só então passar o bastão.",
        "justify",
    )
)

parts.append(h1("6. Formato, escopo e o que fica instalado"))
parts.append(
    table(
        ["Bloco", "Definição"],
        [
            ["Duração", "12 semanas — diagnosticar, construir, instalar e fazer handover com adesão."],
            ["Carga V4", "106,5h (média/alta): análise, desenho, CRM, QA e instalação — não só calls. +6h de gestão de projeto, QA e alinhamentos internos."],
            ["Execução", "V4 constrói e conduz a instalação; o time G6 opera no dia a dia."],
            ["Ritmo", "Encontros semanais com sponsor/gestor e checkpoints com o time comercial (Denis organiza)."],
            ["Presencial", "Fase 1 (diagnóstico) e Fase 5 (instalação/treino) — onde a adesão se decide."],
            ["Saída", "Rotina de gestão instalada + plano de continuidade de 60 dias."],
            ["Janela operacional", "Setembro–novembro de execução + plano de 60 dias; leitura de cobertura até janeiro, com Mateus já operando a máquina."],
        ],
    )
)
parts.append(h2("Encaixe com a chegada do Mateus"))
parts.append(
    p(
        "Primeiros ~30 dias: Fases 1–2 (e início do CRM) avançam antes do gestor entrar. "
        "Quando o Mateus chega, o projeto já está em treinamento do time e dele. "
        "O handover (passagem de bastão do Denis) coincide com o fim do onboarding interno. "
        "O treinamento do gestor não vira problema da diretoria — é troca Denis ↔ Mateus.",
        "justify",
    )
)
parts.append(h2("O que o produto instala"))
parts.append(
    table(
        ["Fundação operacional", "Controle e adesão"],
        [
            ["Arquitetura loja × PAP com critérios de etapa", "CRM como painel de receita (não agenda)"],
            ["Scripts, cadência e papéis claros por canal", "Higiene, motivos de perda e dashboard de uso"],
            ["Rituais de gestão (diária / semanal / pipeline review)", "Treinamento com instalação assistida + roleplay"],
            ["Onboarding operacional de vendedores", "Rotinas de ativação, follow-up e save no processo"],
        ],
    )
)
parts.append(quote("O produto não vende documento. Vende capacidade instalada no time."))

parts.append(h1("7. As seis fases (com entregas e encaixe G6)"))
parts.append(p("Seis fases. Uma fundação comercial instalada. Total: 100,5h de fases + 6h de gestão/QA = 106,5h.", "justify"))
parts.append(
    table(
        ["Fase", "Horas", "Nome", "O que resolve"],
        [
            ["01", "18h", "Diagnóstico operacional", "Operação real, não documento bonito para o problema errado. Presencial."],
            ["02", "15h", "Arquitetura comercial", "Sistema operacional que o time consegue executar."],
            ["03", "15,5h", "CRM e controle", "Em paralelo à F2 — método vira workflow. Brian (CRM) entra aqui."],
            ["04", "20h", "Camada estratégica", "Playbook, comissão e breakeven. ~9ª semana / chegada do Mateus."],
            ["05", "22h", "Instalação da rotina", "Adesão: imersão, roleplay e rituais assistidos. Presencial."],
            ["06", "10h", "Handover 60 dias", "Continuidade + retenção + riscos de regressão."],
        ],
    )
)

parts.append(
    phase(
        "Fase 01 — Diagnóstico operacional  ·  18h  ·  presencial",
        "Entender a operação real da G6 antes de desenhar qualquer playbook. Não é só entrevista: "
        "é assistir a operação. 18 horas provisionadas para conversar com Laura, com a gestão, "
        "com o time de loja e com o PAP — o suficiente para mapear a rotina e os pontos da equação "
        "(onde 1+1 está dando zero).",
        [
            "Diagnóstico comercial detalhado",
            "Entrevistas com sponsor, gestor, top e bottom performer",
            "Auditoria do CRM / RD atual",
            "Mapa de gargalos por etapa",
            "Priorização das dores que impedem venda previsível",
        ],
        "Encaixe G6: presencial no onboarding — time fechado, loja + PAP.",
    )
)
parts.append(
    phase(
        "Fase 02 — Arquitetura operacional comercial  ·  15h",
        "Denis volta da imersão e constrói o plano de melhoria a partir da auditoria: o que a G6 "
        "acerta e o que precisa correção, para um comercial mais fluido, com menos “expresso” e "
        "mais passagem de jornada.",
        [
            "ICP / segmentos e jornada comercial alvo",
            "Pipeline e critérios de etapa · papéis",
            "SLA marketing ↔ comercial ↔ gestor (e comercial → financeiro/operação)",
            "Metas de atividade e indicadores",
            "Scripts · cadências · BPMN · matriz CHA",
        ],
        "Encaixe G6: dual-track loja × PAP.",
    )
)
parts.append(
    phase(
        "Fase 03 — CRM e controle  ·  15,5h  ·  paralelo à Fase 2",
        "O analista de CRM (Brian) entra com a análise do Denis para reorganizar — ou, se a Fase 1 "
        "concluir que o CRM atual não é o ideal, implementar outro (licenças por conta da G6; "
        "configuração e implementação inclusas). Faturamento é a última etapa: cada etapa da jornada "
        "precisa ter métrica de sucesso (analogia do bolo / 250 g de farinha).",
        [
            "Reorganização do pipeline no CRM",
            "Campos obrigatórios · motivos de perda",
            "Atividades obrigatórias e próximas ações",
            "Dashboard de gestão e conversão (o “monitor” do gestor — Mateus)",
            "Regras de higiene · rotina de registro",
        ],
        "Encaixe G6: origem, tempo de resposta, status, dor, capacidade, produto, próxima ação e motivo de perda.",
    )
)
parts.append(
    phase(
        "Fase 04 — Camada estratégica  ·  20h",
        "Processo, gestão, metas, incentivo e viabilidade financeira. Playbook apresentado à diretoria "
        "e à gestão; política de comissionamento; breakeven da operação (“se eu inputar X de marketing "
        "com este time e este formato, aonde eu chego?”). Perspectiva: por volta da 9ª semana, com o "
        "Mateus já em onboarding — ele é onboardado na filosofia do projeto, não no jeito antigo.",
        [
            "Playbook comercial orientado ao uso (não à gaveta)",
            "Política de comissionamento",
            "Breakeven da operação comercial",
        ],
        "Encaixe G6: comissão loja ≠ PAP. Meta como piso. Breakeven amarra CAC, payback e retenção.",
    )
)
parts.append(
    phase(
        "Fase 05 — Instalação da rotina  ·  22h  ·  presencial",
        "Volta presencial a Minas. É a fase que garante que os artefatos sejam usados. Denis conduz "
        "o comercial; Brian conduz o treino de CRM. Roleplay na loja e no PAP. Primeiras rotinas de "
        "pipeline review. Checklist semanal para correção de desvios. Se amanhã a G6 contratar mais "
        "10 vendedores, já existe o pacote de recebimento, treino e documentos.",
        [
            "Onboarding operacional de vendedores",
            "Imersão do time + treinamento do gestor no CRM",
            "Roleplay loja e PAP",
            "Pipeline review e reunião semanal assistidas",
            "Checklist semanal · correção de desvios",
        ],
        "Encaixe G6: imersão presencial + roleplay. Mateus herda a rotina.",
    )
)
parts.append(
    phase(
        "Fase 06 — Handover e continuidade  ·  10h",
        "Últimos 7 a 10 dias. Transferência da propriedade intelectual: tudo documentado, editado e "
        "corrigido conforme a G6 pediu. Score de maturidade antes/depois. Rotinas de retenção e "
        "expansão (o escoamento da base). Mapa de risco de regressão (“se voltarmos a fazer o que "
        "fazíamos, o que quebra?”). Plano tático dos 60 dias seguintes — cobertura de operação até "
        "janeiro; a partir daí, replicar o processo.",
        [
            "Reunião final de transferência",
            "Score de maturidade antes/depois",
            "Plano de 60 dias",
            "Rotinas de retenção e expansão",
            "Riscos de regressão · evolução",
        ],
        "Encaixe G6: motivos de perda, save/upsell e ritual de gestão.",
    )
)

parts.append(h1("8. Três pilares: sistemas, processos e pessoas"))
parts.append(
    p(
        "Denis Orosco — consultor de receita da V4, responsável pela frente de produtos Sales da unidade — "
        "estruturou a fala em três pilares. Essa camada não está como slide separado no deck, mas foi "
        "o fio da apresentação das fases.",
        "justify",
    )
)
parts.append(
    table(
        ["Pilar", "Pergunta que o projeto responde", "Na prática G6"],
        [
            ["Sistemas", "O CRM atual serve? Vale manter ou trocar? Qual a melhor arquitetura da stack?", "Governança do RD/CRM, campos, motivos de perda, dashboard tático do Mateus. Se precisar trocar, implementação inclusa."],
            ["Processos", "Como os fluxos se conectam? Onde está a passagem de bastão venda → operação?", "SLA, jornada loja × PAP, cadência, script, higiene, rituais semanais, checklist de desvio."],
            ["Pessoas", "Como capacito o time, a nova liderança e a direção para gerir o novo método?", "Imersão, roleplay, treino do Mateus, treino da direção na cobrança tático-operacional."],
        ],
    )
)
parts.append(h2("Credenciais apresentadas (Denis)"))
parts.append(
    ul(
        [
            "12 anos em áreas que conectam marketing e comercial.",
            "Google Expert 2025 (um dos 170 profissionais certificados no país) — linha forte de gestão de tráfego, para o comercial não ser atrito do investimento de mídia.",
            "Lean / Six Sigma: Black Belt certificado (único player da V4 Company com essa formação no momento da reunião); rota para Master Black Belt. Mapeamento com roteiro DMAIC.",
            "Win by Design (Vale do Silício) — certificação em arquitetura de receita; menos de 300 profissionais certificados no Brasil.",
            "2 anos na V4 (24º mês em agosto/2026). 1º ano em estruturas complexas — conta de R$ 1,4 bi/ano, 3.200 vendedores, 260 gerentes, 6 estados. Desde setembro anterior na frente dos produtos Sales (produto prototipado por ele).",
        ]
    )
)
parts.append(p("Postura na conta: mediador. A máquina a G6 executa; a V4 ajuda a achar o melhor caminho e instala.", "muted"))

parts.append(h1("9. Critérios de Receita Previsível"))
parts.append(p("Receita Previsível não é slogan. São os critérios observáveis para dizer que a fundação está instalada:", "justify"))
parts.append(
    table(
        ["#", "Critério de sucesso"],
        [
            ["1", "Pipeline configurado e usado — não só “existente”"],
            ["2", "Reunião semanal acontecendo de verdade"],
            ["3", "Follow-up sendo executado"],
            ["4", "Gestor usando indicadores"],
            ["5", "Time nos scripts e critérios"],
            ["6", "Onboarding documentado"],
            ["7", "Rotinas de retenção em operação"],
            ["8", "Sponsor com previsibilidade"],
        ],
    )
)
parts.append(h2("Recap do deck — o que está em jogo"))
parts.append(
    table(
        ["Dores", "Dados"],
        [
            ["Comercial reativo: entra ~1.000 / base quase não cresce", "~1.000 ativações/mês · Meta Forms 1.269 leads · CPL R$ 7,51"],
            ["3.755 cancelamentos em 6 meses — sem churn real conciliado", "Saídas: 415 · 560 · 614 · 746 · 715 · 705 (jan–jun)"],
            ["CRM sem governança — não é painel de receita", "WhatsApp responde em ~1 minuto"],
            ["Loja e PAP sem padrão, script nem cadência", "5 vendedores loja + 2 PAP · Mateus em ~30 dias"],
            ["Ativação, save e retenção fora do processo", "Projeto: 12 semanas · 106,5h · dual-track loja × PAP"],
        ],
    )
)
parts.append(callout("Fechamento do recap", "A Estruturação Comercial instala a máquina que fecha essas dores — com método, adesão e gestão."))

parts.append(h1("10. Investimento, condição comercial e viabilidade"))
parts.append(
    """<table class="price" border="0" cellspacing="0" cellpadding="0">
<tr>
<td class="price-left" width="50%">
<p class="cover-kicker">VALOR DE REFERÊNCIA DO PRODUTO</p>
<p class="strike">R$ 89.635,85</p>
<p class="cover-kicker">CONDIÇÃO COMERCIAL G6  ·  BASE ATIVA</p>
<p class="price-now">R$ 71.582,32</p>
<p class="cover-lead">12 semanas · 106,5h · implementação assistida<br>dual-track loja × PAP · presenciais nas fases críticas</p>
</td>
<td class="price-right" width="50%">
<p class="callout-label">FORMA APRESENTADA NA REUNIÃO</p>
<p class="price-month">12 recorrências  ·  R$ 5.965,19 / mês</p>
<p>Pagamento recorrente (sem cartão de crédito), para não competir com a liberação do cartão do contrato de marketing.</p>
<p>Redução de R$ 18.053,53 em relação à referência, por Laura ser cliente da base ativa.</p>
<p>Custo único de implementação. Não se repete no ano 2. O acompanhamento do comercial segue no executado de marketing, sem custo adicional.</p>
</td>
</tr>
</table>"""
)
parts.append(h2("O que está incluso / por que agora / resultado esperado"))
parts.append(
    table(
        ["Bloco", "Conteúdo"],
        [
            ["Incluso", "Capacidade instalada: diagnóstico → arquitetura → CRM como painel de receita → estratégica → instalação → handover 60 dias."],
            ["Por que agora", "Parar o vazamento. Com 3.755 saídas em 6 meses e churn real ainda não conciliado, estruturar o comercial é a decisão que dá previsibilidade."],
            ["Resultado esperado", "Máquina comercial: processo, rituais e gestão rodando — para vender, ativar e reter com método."],
        ],
    )
)
parts.append(h2("Como a conta fecha no ano (para o Geraldo)"))
parts.append(
    p(
        "Na reunião, Laura somou o marketing já fechado com a Estruturação Comercial. Rafaela pediu "
        "para levar ao Geraldo o depara com os R$ 7 milhões — senão ele olha preço, não investimento.",
        "justify",
    )
)
parts.append(
    table(
        ["Linha", "Base de cálculo", "Total"],
        [
            ["Marketing V4 (já fechado)", "12 × R$ 11.180", "R$ 134.160"],
            ["Estruturação Comercial (proposta)", "R$ 71.582,32  (ou 12 × R$ 5.965,19)", "R$ 71.582,32"],
            ["Pacote anual marketing + EC", "Laura arredondou o mensal para ~R$ 17.180", "≈ R$ 205.700  (Rafaela citou R$ 204 mil)"],
            ["+ mídia (~R$ 10 mil/mês)", "R$ 120.000 no ano", "≈ R$ 324.000 no ano"],
            ["Vazamento estimado em 6 meses", "Conta de padaria da reunião", "≈ R$ 7.000.000"],
            ["Relação investimento × vazamento", "324 mil / 7 milhões", "Menos de 5%  (Rafaela: “nem 10%”)"],
        ],
    )
)
parts.append(
    callout(
        "Argumento de decisão levado ao Geraldo",
        "Ou a G6 mexe na máquina agora para evitar outro semestre de vazamento, ou continua sentada "
        "esperando o comercial mudar sozinho. Daqui a 3 meses a V4 entrega lead qualificado e o furo "
        "do balde continua do lado comercial — e a pauta vira briga de conversão, não expansão. "
        "O investimento para buscar receita previsível está mais barato do que um mês perdendo ~800 clientes.",
    )
)
parts.append(h2("O que é único vs. o que renova"))
parts.append(
    ul(
        [
            "Estruturação Comercial: custo único. Implementa, o executado acompanha, não gera nova fatura dessa frente.",
            "Marketing (execução): 12 meses. Depois a G6 renova ou não. Mídia, criativo e operação de growth estão nesse bloco.",
            "Estruturação estratégica prévia (plano de marketing já feito): também é etapa única de entrada — auditar e planejar antes de executar, para não repetir erro.",
        ]
    )
)
parts.append(
    p(
        "Rafaela não negocia valor na mesa de estratégia. Dinheiro fica com a Sara. Na reunião, o "
        "pedido explícito a Laura: transparência se o número não couber — “calculamos que se fosse X "
        "a gente conseguiria” — para a V4 tentar viabilizar o start agora, não daqui a alguns meses.",
        "justify",
    )
)

parts.append(h1("11. Ata da reunião de 14/08/2026"))
parts.append(
    p(
        "Fonte: observações Gemini + transcrição da call “[Comercial] V4 Company :::: G6 Internet”, "
        "14 de agosto de 2026. Duração aproximada: 1h23. Gravação anexa no Google Docs.",
        "muted",
    )
)
parts.append(h2("Participantes e papéis"))
parts.append(
    table(
        ["Pessoa", "Papel na reunião"],
        [
            ["Laurelise Santos (Laura)", "Sponsor G6. Precisa validar com Geraldo. Time comercial de 7 pessoas + Mateus."],
            ["Rafaela Tolomeu Messias", "Conduziu a leitura estratégica, a conta de vazamento e o investimento."],
            ["Sara Pizzico", "Comercial V4. Dona da negociação de valor (não esteve na fala de preço; Rafaela redireciona para ela)."],
            ["Denis Orosco", "Consultor de receita / Sales. Detalhou fases, credenciais e formato. Saiu ~1h07 por outra agenda."],
            ["Geraldo (ausente)", "Sócio. Dono da validação financeira. Laura conversa com ele na terça."],
            ["Mateus (ausente)", "Novo gestor comercial. Entra em ~30 dias. Herda a máquina."],
            ["Brian (citado)", "Analista de CRM. Entra na Fase 3 e no treino da Fase 5. Duarte já o apresentou à G6."],
            ["Duarte / Natã (citados)", "Execução de marketing / agenda. Pauta de lançamento móvel na semana seguinte."],
        ],
    )
)
parts.append(h2("Resumo da ata (Gemini)"))
parts.append(
    p(
        "A reunião abordou a reestruturação comercial para estancar perdas financeiras e planejar o "
        "lançamento de novos serviços. A organização identificou falhas graves na governança de sistemas "
        "e processos, com prejuízo acumulado estimado de R$ 7 milhões em 6 meses. Decidiu-se o plano de "
        "12 semanas focado em sistemas, processos e capacitação. A viabilidade financeira do projeto foi "
        "contextualizada; o planejamento do serviço móvel foi iniciado como pauta paralela.",
        "justify",
    )
)
parts.append(h2("Decisões registradas"))
parts.append(
    table(
        ["Tipo", "Decisão"],
        [
            ["Precisa de mais conversa", "Validação da proposta de reestruturação comercial com o sócio Geraldo antes do fechamento definitivo."],
            ["Alinhada", "Adoção do roteiro de 12 semanas da V4, incluindo as seis fases de diagnóstico, arquitetura, implementação e treinamento."],
            ["Alinhada", "Diagnóstico operacional in loco, com 18 horas provisionadas para mapeamento."],
            ["Alinhada", "Identidade visual do G6 Móvel segue o mesmo padrão da G6 Internet, para credibilidade e reconhecimento."],
        ],
    )
)
parts.append(h2("Próximas etapas (com dono)"))
parts.append(
    table(
        ["Dono", "Ação"],
        [
            ["Denis Orosco", "Realizar o diagnóstico operacional in loco (18h): mapear rotinas, entrevistar gestão, observar loja e campo."],
            ["Denis Orosco", "Apresentar a estrutura detalhada do plano no kickoff."],
            ["Denis Orosco", "Desenvolver a arquitetura operacional e comercial: metas, indicadores, scripts e cadências."],
            ["Denis Orosco", "Treinar o time comercial e o gestor Mateus na nova estrutura e na usabilidade do CRM."],
            ["Laura", "Apresentar a proposta ao Geraldo e demais sócios; validar investimento e aceitação do projeto (terça)."],
            ["Rafaela", "Briefing ao Duarte sobre o lançamento G6 Móvel: PDV, abrangência inicial, nome e comunicação."],
            ["Rafaela / Natã", "Agendar reunião na semana seguinte com o time de execução para lançamento do produto."],
        ],
    )
)
parts.append(h2("Linha do tempo da conversa (pauta comercial)"))
parts.append(
    table(
        ["Min", "O que aconteceu"],
        [
            ["04:34", "Rafaela abre o contexto: dores, Mateus, receita minando, dados soltos. Objetivo: entregar máquina pronta ao gestor."],
            ["06:46", "Espelho: sem receita previsível; entra ~1.000 / sai ~800; CRM sem governança; meta que desce; loja e PAP sem processo."],
            ["09:52", "Conta do vazamento: R$ 800, 8 meses, 3.755 cancelamentos, ~R$ 3 mi + ~R$ 4 mi = ~R$ 7 mi em 6 meses."],
            ["12:20", "Cenário 2% (~75 clientes, ~R$ 150 mil retidos). Sem estrutura, o balde continua furado."],
            ["16:54", "Denis se apresenta como consultor de receita / mediador da máquina Sales."],
            ["19:27", "Jornada: mensagens longas, sem upsell na base, risco na chegada do chip."],
            ["21:26", "Promessa do projeto fechado (não é PDF). 12 semanas, 106h, F1 e F5 presenciais, handover 60 dias."],
            ["26:15", "Credenciais Denis: Google Expert, Black Belt, Win by Design, conta de R$ 1,4 bi."],
            ["31:41", "Três pilares: sistemas, processos, pessoas. Roadmap das 6 fases."],
            ["34:39", "Por que presencial: cultura reativa vista in loco. 18h de diagnóstico."],
            ["40:22", "Fases 2 e 3: arquitetura + CRM (Brian), campos, dashboard, métricas por etapa."],
            ["54:37", "Fase 4: playbook, comissão, breakeven, onboarding do Mateus ~9ª semana."],
            ["55:44", "Fase 5: volta a Minas, imersão, roleplay, rituais, checklist."],
            ["57:43", "Fase 6: transferência de IP, maturidade, retenção, risco de regressão, plano 60 dias."],
            ["59:49", "Laura confirma encaixe com a dor da quarta-feira. Quer fazer; precisa dos sócios. Antecipa resistência do time de 7 + Mateus."],
            ["01:02", "Rafaela: marketing + EC < 10% dos R$ 7 mi. Valor não se negocia com ela — vai para a Sara."],
            ["01:05", "Investimento: de R$ 89.635,85 para R$ 71.582,32; 12× recorrente, sem cartão."],
            ["01:07", "Denis se despede. Laura: precisa validar; soma ~R$ 17 mil/mês com o marketing. Rafaela ancora R$ 204 mil vs. R$ 7 mi."],
            ["01:11", "EC é custo único; acompanhamento já está no executado. Marketing renova ou não aos 12 meses."],
            ["01:13", "Laura fala com Geraldo na terça e puxa a V4 se precisar. Rafaela pede transparência para viabilizar."],
            ["01:15", "Pauta extra: lançamento móvel (nome, PDV, Piranga Sul / Tejubá, mesma cara da G6)."],
        ],
    )
)
parts.append(h2("Posição da Laura na mesa"))
parts.append(
    p(
        "Laura confirmou que o escopo responde à pergunta que ela fez na reunião anterior (“vocês vão "
        "olhar o comercial também?”). Enxerga a mudança como corte de era: a G6 não é mais a empresa "
        "pequena em que se conhecia cada funcionário. Já prevê que parte do time de 7 “espanne”, e que "
        "a chegada do Mateus sozinha já gera atrito. Assume a mudança porque “não dá para ser pai e mãe "
        "o tempo todo”. Sentimento declarado: “quero fazer” — com o caveat de repassar aos sócios. "
        "O valor não tinha sido pensado antes; surgiu da conversa daquela semana.",
        "justify",
    )
)

parts.append(h1("12. Pauta complementar: lançamento G6 Móvel / Chip"))
parts.append(
    p(
        "Depois da proposta, Laura pediu ajuda no lançamento do serviço móvel, que deve sair mais "
        "rápido do que o previsto. Esta pauta não faz parte do escopo da Estruturação Comercial — "
        "foi encaminhada para o time de execução (Duarte) em reunião específica na semana seguinte. "
        "Entra neste documento porque foi apresentada na mesma call e se conecta com a tese comercial: "
        "sem time treinado para vender na base, o chip vira novo CAC em cima de carteira já paga.",
        "justify",
    )
)
parts.append(
    table(
        ["Tema", "O que ficou"],
        [
            ["Nome em avaliação", "G6 Chip  ·  G6 Móvel  ·  G6 Celular  ·  G6 5G"],
            ["Linha de comunicação sugerida (Laura)", "“A qualidade e a constância que você conhece na sua casa, agora no seu bolso.”"],
            ["Marca", "Mesma cara da G6 Internet. Credibilidade de mercado (analogia: a Vivo não muda de cara entre casa e celular)."],
            ["Praça 1", "Piranga Sul (transcrição: “Pirango Sul”) — ~6.000 habitantes. Lançamento pequeno para validar."],
            ["Praça 2", "Tejubá (transcrição) — ~90 mil habitantes, dos quais ~20 mil já são clientes G6. Venda nova + base."],
            ["Materiais de PDV", "Wind banner na porta da loja; totem/display de papelão; wobblers (“bolachinha”). Produzir o kit para todas as lojas mesmo usando só a primeira agora."],
            ["CRM / base", "Neste momento Laura não precisa de disparo massivo na base toda. Validar off na cidade pequena; campanha/disparo só daquela região se fizer sentido."],
            ["Próximo passo", "Rafaela passa o briefing ao Duarte e pede ao Natã a agenda da semana seguinte."],
        ],
    )
)

parts.append(h1("13. Como levar este caso ao Geraldo"))
parts.append(p("Roteiro que a Rafaela pediu para Laura usar — útil também para a Sara no follow-up:", "justify"))
parts.append(
    ul(
        [
            "Não abrir pelo preço. Abrir pelos R$ 7 milhões em 6 meses e pelos 3.755 cancelamentos.",
            "Mostrar que cada cliente precisa de 8 meses para se pagar (CAC + instalação ≈ R$ 800).",
            "Deparar os ~R$ 204 mil no ano (marketing + EC) — ou ~R$ 324 mil com mídia — com o vazamento. É menos de 10%.",
            "Deixar claro: Estruturação Comercial é custo único; não vira mensalidade eterna.",
            "Explicar o timing do Mateus: os primeiros 30 dias estruturam a casa; ele herda máquina + treino, em vez de improvisar processo.",
            "Cenário mínimo (2%): ~R$ 150 mil retidos já ajudam a pagar o projeto. Qualquer melhoria real tende a ser maior.",
            "Risco de esperar 3 meses: a V4 entrega lead, o comercial continua furado, a relação vira cobrança de conversão.",
            "Se o número não couber, voltar com o teto que cabe — a V4 tenta viabilizar o start agora (recorrência já foi o gesto para não travar no cartão).",
        ]
    )
)
parts.append(
    callout(
        "Frase de fechamento (script interno)",
        "O que a gente propõe não é mais esforço comercial. É previsibilidade: parar o vazamento com "
        "método instalado na loja e no PAP — em doze semanas.",
    )
)

parts.append(h1("14. Fontes e materiais de apoio"))
parts.append(
    table(
        ["Material", "Onde"],
        [
            ["Ata e transcrição da reunião 14/08/2026", "https://docs.google.com/document/d/12M9kFqxpMApDL0yZMOgtmz1UbBXQvFDMvXZCsnZZSAE/edit"],
            ["Deck Estruturação Comercial (18 slides)", "proposta-g6-estruturacao-comercial.html  ·  /g6/  ·  /g6-ec/"],
            ["PDF do deck", "G6-Internet_Estruturacao-Comercial_V4.pdf"],
            ["Script interno de narrativa (slide a slide)", "script-narrativa-g6-estruturacao-comercial.docx"],
            ["Proposta Growth + CRM (12 meses, contexto anterior)", "proposta-g6-internet.html"],
        ],
    )
)
parts.append(
    p(
        "Documento consolidado pela V4 Company a partir do deck apresentado e da reunião comercial "
        "de 14 de agosto de 2026. Valores e prazos conforme falados na mesa e registrados no deck. "
        "A conta de R$ 7 milhões é estimativa de gestão (não conciliação contábil), e foi apresentada "
        "assim porque o churn real ainda não está conciliado — exatamente o que o projeto se propõe a instalar.",
        "muted",
    )
)
parts.append('<p class="footer-brand">G6 Internet × V4 Company  ·  Estruturação Comercial  ·  Confidencial</p>')

BODY = "\n".join(parts)

HTML = f"""<!DOCTYPE html>
<html xmlns:o="urn:schemas-microsoft-com:office:office"
      xmlns:w="urn:schemas-microsoft-com:office:word"
      xmlns="http://www.w3.org/TR/REC-html40">
<head>
<meta charset="utf-8">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<meta name="ProgId" content="Word.Document">
<meta name="Generator" content="Microsoft Word 15">
<meta name="Originator" content="Microsoft Word 15">
<title>G6 Internet × V4 — Estruturação Comercial</title>
<!--[if gte mso 9]>
<xml>
  <o:DocumentProperties>
    <o:Title>G6 Internet × V4 — Estruturação Comercial</o:Title>
    <o:Author>V4 Company</o:Author>
    <o:Description>Documento consolidado da proposta e da reunião de 14/08/2026</o:Description>
  </o:DocumentProperties>
  <w:WordDocument>
    <w:View>Print</w:View>
    <w:Zoom>100</w:Zoom>
    <w:DoNotOptimizeForBrowser/>
    <w:ValidateAgainstSchemas/>
  </w:WordDocument>
</xml>
<![endif]-->
<style>
  @page WordSection1 {{
    size: 21cm 29.7cm;
    margin: 1.8cm 2cm 2cm 2cm;
    mso-header-margin: 1cm;
    mso-footer-margin: 1.2cm;
    mso-paper-source: 0;
  }}
  div.WordSection1 {{ page: WordSection1; }}
  body {{
    font-family: Calibri, Arial, sans-serif;
    font-size: 11pt;
    color: #1A1A1A;
    line-height: 1.35;
    margin: 24px auto;
    max-width: 820px;
    background: #f4f0ef;
  }}
  div.WordSection1 {{
    background: #fff;
    padding: 28px 32px 48px;
  }}
  h1 {{
    font-size: 18pt;
    color: #A10F14;
    border-bottom: 1.5pt solid #A10F14;
    padding-bottom: 4pt;
    margin: 22pt 0 8pt;
    page-break-after: avoid;
  }}
  h2 {{
    font-size: 13pt;
    color: #E50914;
    margin: 14pt 0 6pt;
    page-break-after: avoid;
  }}
  p {{ margin: 0 0 8pt; }}
  p.justify {{ text-align: justify; }}
  p.quote {{
    color: #A10F14;
    font-style: italic;
    margin: 6pt 0 10pt;
  }}
  p.muted {{ color: #4A4A4A; font-size: 10.5pt; font-style: italic; }}
  p.label {{ color: #E50914; font-weight: 700; margin: 8pt 0 4pt; }}
  p.strong-red {{ color: #A10F14; font-weight: 700; }}
  p.footer-brand {{
    color: #A10F14;
    font-weight: 700;
    text-align: center;
    font-size: 9pt;
    margin-top: 16pt;
  }}
  ul {{ margin: 4pt 0 10pt 18pt; padding: 0; }}
  li {{ margin: 0 0 3pt; color: #4A4A4A; }}
  table.grid {{
    width: 100%;
    border-collapse: collapse;
    margin: 8pt 0 12pt;
    font-size: 10pt;
  }}
  table.grid th {{
    background: #A10F14;
    color: #fff;
    text-align: left;
    font-size: 9pt;
    letter-spacing: .04em;
    text-transform: uppercase;
    padding: 6pt 8pt;
    border: 1pt solid #A10F14;
  }}
  table.grid td {{
    border: 1pt solid #E6D4D4;
    padding: 6pt 8pt;
    vertical-align: top;
  }}
  table.grid tr:nth-child(even) td {{ background: #F7F5F4; }}
  table.callout {{
    width: 100%;
    margin: 8pt 0 12pt;
    border-left: 4pt solid #A10F14;
    background: #F8EFEF;
  }}
  table.callout td {{ padding: 10pt 12pt; }}
  .callout-label {{
    color: #A10F14;
    font-size: 8.5pt;
    font-weight: 700;
    letter-spacing: .08em;
    margin: 0 0 4pt;
  }}
  table.cover {{
    width: 100%;
    background: #280001;
    margin: 0 0 14pt;
  }}
  table.cover td {{ padding: 22pt 20pt; }}
  .cover-kicker {{
    color: #F5B5B8;
    font-size: 9pt;
    font-weight: 700;
    letter-spacing: .12em;
    margin: 0 0 8pt;
  }}
  .cover-title {{
    color: #fff;
    font-size: 28pt;
    font-weight: 700;
    margin: 0 0 6pt;
    line-height: 1.1;
  }}
  .cover-sub {{
    color: #FFD0D0;
    font-size: 14pt;
    font-style: italic;
    margin: 0 0 10pt;
  }}
  .cover-lead {{
    color: #E8D0D0;
    font-size: 11pt;
    margin: 0;
  }}
  table.price {{ width: 100%; margin: 8pt 0 12pt; }}
  td.price-left {{
    background: #280001;
    color: #fff;
    padding: 16pt;
    vertical-align: top;
    width: 50%;
  }}
  td.price-right {{
    background: #F8EFEF;
    border: 1.5pt solid #A10F14;
    padding: 14pt;
    vertical-align: top;
    width: 50%;
  }}
  .strike {{
    color: #C9A0A0;
    text-decoration: line-through;
    font-size: 14pt;
    margin: 0 0 8pt;
  }}
  .price-now {{
    color: #fff;
    font-size: 24pt;
    font-weight: 700;
    margin: 0 0 8pt;
  }}
  .price-month {{
    font-size: 13pt;
    font-weight: 700;
    color: #1A1A1A;
    margin: 0 0 8pt;
  }}
  .open-bar {{
    background: #fff3cd;
    border: 1pt solid #e0c36a;
    padding: 10pt 12pt;
    margin: 0 0 14pt;
    font-size: 10.5pt;
  }}
  .open-bar code {{
    background: #fff;
    padding: 1pt 4pt;
  }}
</style>
</head>
<body>
<div class="WordSection1">
{BODY}
</div>
</body>
</html>
"""

OUT_HTML.write_text(HTML, encoding="utf-8")
OUT_DOC.write_text(HTML, encoding="utf-8")
print("Wrote", OUT_HTML, OUT_HTML.stat().st_size)
print("Wrote", OUT_DOC, OUT_DOC.stat().st_size)
