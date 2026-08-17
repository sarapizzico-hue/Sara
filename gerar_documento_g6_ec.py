#!/usr/bin/env python3
"""Gera a proposta comercial (Word + HTML Word) da Estruturação Comercial G6."""

from html import escape
from pathlib import Path
from shutil import copy2

from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor

RED = "A10F14"
RED2 = "E50914"
INK = "1A1A1A"
MUT = "4A4A4A"
CREAM = "F7F5F4"
LINE = "E6D4D4"
WHITE = "FFFFFF"
SOFT = "F8EFEF"
DARK = "280001"

OUT_DOCX = Path("/workspace/documento-g6-internet-estruturacao-comercial.docx")
OUT_HTML = Path("/workspace/documento-g6-internet-estruturacao-comercial-word.html")
OUT_DOC = Path("/workspace/documento-g6-internet-estruturacao-comercial.doc")
G6_WORD = Path("/workspace/g6-word")


def set_run_font(run, name="Calibri", size=11, bold=False, color=INK, italic=False):
    run.font.name = name
    run._element.rPr.rFonts.set(qn("w:eastAsia"), name)
    run.font.size = Pt(size)
    run.bold = bold
    run.italic = italic
    run.font.color.rgb = RGBColor.from_string(color)


def shade_cell(cell, hex_color):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:fill"), hex_color)
    shd.set(qn("w:val"), "clear")
    tcPr.append(shd)


def set_cell_borders(cell, color="E6D4D4", sz="4"):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    tcBorders = OxmlElement("w:tcBorders")
    for edge in ("top", "left", "bottom", "right"):
        el = OxmlElement(f"w:{edge}")
        el.set(qn("w:val"), "single")
        el.set(qn("w:sz"), sz)
        el.set(qn("w:space"), "0")
        el.set(qn("w:color"), color)
        tcBorders.append(el)
    tcPr.append(tcBorders)


def set_cell_margins(cell, top=80, bottom=80, left=120, right=120):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    tcMar = OxmlElement("w:tcMar")
    for m, val in (("top", top), ("left", left), ("bottom", bottom), ("right", right)):
        node = OxmlElement(f"w:{m}")
        node.set(qn("w:w"), str(val))
        node.set(qn("w:type"), "dxa")
        tcMar.append(node)
    tcPr.append(tcMar)


def set_table_width(table, width_cm):
    table.autofit = False
    table.allow_autofit = False
    tbl = table._tbl
    tblPr = tbl.tblPr if tbl.tblPr is not None else OxmlElement("w:tblPr")
    tblW = OxmlElement("w:tblW")
    tblW.set(qn("w:w"), str(int(width_cm * 567)))
    tblW.set(qn("w:type"), "dxa")
    tblPr.append(tblW)


def prevent_row_split(row):
    tr = row._tr
    trPr = tr.get_or_add_trPr()
    cant = OxmlElement("w:cantSplit")
    trPr.append(cant)


def cell_para(cell, text, size=11, bold=False, color=INK, align="left", space_after=4):
    cell.text = ""
    p = cell.paragraphs[0]
    p.alignment = {
        "left": WD_ALIGN_PARAGRAPH.LEFT,
        "center": WD_ALIGN_PARAGRAPH.CENTER,
        "right": WD_ALIGN_PARAGRAPH.RIGHT,
    }[align]
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(space_after)
    p.paragraph_format.line_spacing = 1.15
    run = p.add_run(text)
    set_run_font(run, size=size, bold=bold, color=color)
    return p


def add_cell_run(cell, text, size=11, bold=False, color=INK, italic=False):
    p = cell.paragraphs[0] if cell.paragraphs[0].text == "" and len(cell.paragraphs) == 1 else cell.add_paragraph()
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(2)
    run = p.add_run(text)
    set_run_font(run, size=size, bold=bold, color=color, italic=italic)
    return p


def add_p(doc, text, size=11, bold=False, color=INK, space_before=0, space_after=8, align="left", italic=False):
    p = doc.add_paragraph()
    p.alignment = {
        "left": WD_ALIGN_PARAGRAPH.LEFT,
        "center": WD_ALIGN_PARAGRAPH.CENTER,
        "right": WD_ALIGN_PARAGRAPH.RIGHT,
        "justify": WD_ALIGN_PARAGRAPH.JUSTIFY,
    }[align]
    p.paragraph_format.space_before = Pt(space_before)
    p.paragraph_format.space_after = Pt(space_after)
    p.paragraph_format.line_spacing = 1.15
    run = p.add_run(text)
    set_run_font(run, size=size, bold=bold, color=color, italic=italic)
    return p


def add_heading_styled(doc, text, level=1):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(16 if level == 1 else 12)
    p.paragraph_format.space_after = Pt(8)
    if level == 1:
        pPr = p._p.get_or_add_pPr()
        pBdr = OxmlElement("w:pBdr")
        bottom = OxmlElement("w:bottom")
        bottom.set(qn("w:val"), "single")
        bottom.set(qn("w:sz"), "12")
        bottom.set(qn("w:space"), "4")
        bottom.set(qn("w:color"), RED)
        pBdr.append(bottom)
        pPr.append(pBdr)
    run = p.add_run(text)
    set_run_font(run, size=16 if level == 1 else 13, bold=True, color=RED if level == 1 else RED2)
    return p


def add_bullet(doc, text, size=11):
    p = doc.add_paragraph(style="List Bullet")
    p.paragraph_format.space_before = Pt(1)
    p.paragraph_format.space_after = Pt(3)
    p.paragraph_format.left_indent = Cm(0.75)
    r = p.add_run(text)
    set_run_font(r, size=size, color=MUT)
    return p


def callout(doc, label, text):
    table = doc.add_table(rows=1, cols=1)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_width(table, 16.5)
    cell = table.cell(0, 0)
    shade_cell(cell, SOFT)
    set_cell_borders(cell, RED, "12")
    set_cell_margins(cell, 100, 100, 140, 140)
    cell.text = ""
    p1 = cell.paragraphs[0]
    p1.paragraph_format.space_after = Pt(4)
    r1 = p1.add_run(label.upper())
    set_run_font(r1, size=9, bold=True, color=RED)
    p2 = cell.add_paragraph()
    p2.paragraph_format.space_after = Pt(0)
    r2 = p2.add_run(text)
    set_run_font(r2, size=11, color=INK)
    doc.add_paragraph().paragraph_format.space_after = Pt(6)


def make_table(doc, headers, rows, col_widths=None):
    table = doc.add_table(rows=1 + len(rows), cols=len(headers))
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_width(table, 16.5)
    for i, h in enumerate(headers):
        cell = table.rows[0].cells[i]
        shade_cell(cell, RED)
        set_cell_borders(cell, RED, "4")
        set_cell_margins(cell, 70, 70, 90, 90)
        cell_para(cell, h, size=9, bold=True, color=WHITE, space_after=0)
    for r_i, row in enumerate(rows):
        for c_i, val in enumerate(row):
            cell = table.rows[r_i + 1].cells[c_i]
            shade_cell(cell, CREAM if r_i % 2 else WHITE)
            set_cell_borders(cell, LINE, "4")
            set_cell_margins(cell, 70, 70, 90, 90)
            cell_para(cell, val, size=10, bold=c_i == 0, space_after=0)
        prevent_row_split(table.rows[r_i + 1])
    prevent_row_split(table.rows[0])
    if col_widths:
        for row in table.rows:
            for i, w in enumerate(col_widths):
                row.cells[i].width = Cm(w)
    spacer = doc.add_paragraph()
    spacer.paragraph_format.space_after = Pt(8)
    return table


def add_page_number(paragraph):
    run = paragraph.add_run()
    fld_begin = OxmlElement("w:fldChar")
    fld_begin.set(qn("w:fldCharType"), "begin")
    instr = OxmlElement("w:instrText")
    instr.set(qn("xml:space"), "preserve")
    instr.text = " PAGE "
    fld_end = OxmlElement("w:fldChar")
    fld_end.set(qn("w:fldCharType"), "end")
    run._r.append(fld_begin)
    run._r.append(instr)
    run._r.append(fld_end)


def build_docx():
    doc = Document()
    section = doc.sections[0]
    section.page_width = Cm(21.0)
    section.page_height = Cm(29.7)
    section.left_margin = Cm(2.0)
    section.right_margin = Cm(2.0)
    section.top_margin = Cm(1.8)
    section.bottom_margin = Cm(2.0)

    header = section.header
    header.is_linked_to_previous = False
    hp = header.paragraphs[0]
    hp.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    r = hp.add_run("G6 Internet  ×  V4 Company  ·  Proposta comercial")
    set_run_font(r, size=8, color=RED, bold=True)

    footer = section.footer
    footer.is_linked_to_previous = False
    fp = footer.paragraphs[0]
    fp.alignment = WD_ALIGN_PARAGRAPH.LEFT
    r = fp.add_run("Estruturação Comercial  ·  Confidencial  ·  Página ")
    set_run_font(r, size=8, color=MUT)
    add_page_number(fp)

    cover = doc.add_table(rows=1, cols=1)
    set_table_width(cover, 17.0)
    c = cover.cell(0, 0)
    shade_cell(c, DARK)
    set_cell_borders(c, DARK, "0")
    set_cell_margins(c, 280, 280, 220, 220)
    c.text = ""
    p = c.paragraphs[0]
    r = p.add_run("PROPOSTA COMERCIAL")
    set_run_font(r, size=11, bold=True, color="F5B5B8")
    p2 = c.add_paragraph()
    p2.paragraph_format.space_before = Pt(10)
    r = p2.add_run("Estruturação Comercial")
    set_run_font(r, size=32, bold=True, color=WHITE)
    p3 = c.add_paragraph()
    p3.paragraph_format.space_after = Pt(12)
    r = p3.add_run("G6 Internet  ×  V4 Company")
    set_run_font(r, size=16, color="FFD0D0")
    p4 = c.add_paragraph()
    r = p4.add_run(
        "Projeto para instalar a fundação comercial da G6: processo, CRM em uso, "
        "gestão e rotina — loja e porta a porta — em 12 semanas."
    )
    set_run_font(r, size=12, color="E8D0D0")

    add_p(doc, "", size=8, space_after=4)
    make_table(
        doc,
        ["Campo", "Informação"],
        [
            ["Cliente", "G6 Internet"],
            ["Produto", "Estruturação Comercial (Sales V4) — projeto fechado com implementação assistida"],
            ["Prazo", "12 semanas  ·  106,5 horas  ·  dual-track loja × porta a porta"],
            ["Investimento", "R$ 71.582,32  (referência: R$ 89.635,85)  ·  12 parcelas recorrentes"],
            ["Condução", "V4 Company  ·  Denis Orosco (receita / Sales)  ·  Brian (CRM)"],
            ["Sponsor G6", "Laurelise Santos  ·  validação societária com Geraldo"],
            ["Gestor comercial", "Mateus — entrada prevista em cerca de 30 dias"],
        ],
        [4.4, 12.1],
    )

    callout(
        doc,
        "Leitura central",
        "A G6 não tem problema de lead. Tem um comercial reativo, sem sistema. "
        "A demanda entra; o processo não segura. Esta proposta instala a máquina "
        "para vender, ativar e reter com previsibilidade — e entrega essa máquina "
        "pronta para o novo gestor operar.",
    )

    add_heading_styled(doc, "1. Objetivo")
    add_p(
        doc,
        "Construir e instalar a fundação comercial da G6 Internet em um projeto fechado de "
        "12 semanas: diagnóstico presencial, arquitetura de processo, CRM como painel de receita, "
        "camada estratégica (playbook, comissão e breakeven), imersão/treino do time e handover "
        "com plano de continuidade de 60 dias.",
        align="justify",
    )
    add_p(
        doc,
        "O resultado esperado não é um documento para a gaveta. É a operação rodando: loja e "
        "porta a porta com método, rituais de gestão e o Mateus herdando uma máquina pronta — "
        "em vez de improvisar processo, troca de peça e meta ao mesmo tempo.",
        align="justify",
    )

    add_heading_styled(doc, "2. Diagnóstico da operação")
    add_p(
        doc,
        "A G6 gera demanda. Meta Forms entregou 1.269 leads no ano com CPL de R$ 7,51. "
        "O WhatsApp responde em cerca de 1 minuto. A operação cita cerca de 1.000 ativações "
        "por mês. O que falta é sistema para converter, ativar, seguir e reter.",
        align="justify",
    )
    make_table(
        doc,
        ["Sintoma", "Efeito na operação"],
        [
            [
                "Volume alto, retenção frágil",
                "Entra cerca de 1.000 clientes e sai uma faixa próxima de 800. O crescimento líquido vaza todo mês.",
            ],
            [
                "CRM existe, mas não governa",
                "Sem motivo de perda, sem próximo passo, sem previsibilidade de pipeline. Não funciona como painel de receita.",
            ],
            [
                "Meta vira comodidade",
                "Quando aperta, a meta desce para o time “bater” e receber comissão. A liderança cascateia acomodação.",
            ],
            [
                "Loja e PAP sem padrão",
                "5 internos + 2 porta a porta. Cada canal vende de um jeito, sem playbook, deparo nem objeção mapeada.",
            ],
            [
                "Pós-venda solto",
                "Vendeu e acabou. Ativação, save, follow-up e upsell na base ficam fora do processo.",
            ],
            [
                "Gestão sem ritual pronto",
                "O novo gestor precisa herdar reunião comercial, indicadores e critérios — não inventar do zero.",
            ],
        ],
        [5.2, 11.3],
    )
    add_p(doc, "Pontos da operação que o projeto precisa endereçar", size=12, bold=True, color=RED2, space_before=4)
    add_bullet(doc, "Time comercial de 7 pessoas, mais o Mateus na chegada.")
    add_bullet(doc, "Porta a porta sem deparo: região, oportunidades, o que trouxe, o que não trouxe, objeções.")
    add_bullet(doc, "Atendimento inicial longo demais para um produto de alta intenção (primeiro plano de internet).")
    add_bullet(doc, "Base instalada sem jornada de upgrade — o time olha só a venda nova.")
    add_bullet(doc, "Casos de instalação sem pagamento/ativação: comissão apagada e receita nenhuma.")
    add_bullet(doc, "Cultura de atendimento forte; o comercial herdou o ritmo do atendimento e ficou reativo.")
    add_bullet(doc, "Chegada do serviço móvel/chip: sem time treinado para vender na base, a G6 gasta de novo para ofertar um produto à carteira que já pagou para entrar.")
    callout(
        doc,
        "Por que o diagnóstico é presencial",
        "A Fase 1 acontece in loco (18 horas) para mapear cultura, resistência à mudança, loja, "
        "porta a porta e o dia a dia da unidade. Sem isso, o plano fica genérico e não cola na G6.",
    )

    add_heading_styled(doc, "3. Impacto financeiro")
    add_p(
        doc,
        "A G6 tem capacidade de vender. O que a operação ainda não tem é capacidade de segurar — "
        "nem de medir, com rigor, por que perde.",
        align="justify",
    )
    add_heading_styled(doc, "3.1 Cancelamentos no 1º semestre de 2026", 2)
    make_table(
        doc,
        ["Mês", "Cancelamentos"],
        [
            ["Janeiro", "415"],
            ["Fevereiro", "560"],
            ["Março", "614"],
            ["Abril", "746"],
            ["Maio", "715"],
            ["Junho", "705"],
            ["Total jan–jun", "3.755"],
        ],
        [8.25, 8.25],
    )
    add_p(
        doc,
        "A taxa real de churn ainda não está conciliada. Falta carteira ativa no início de cada mês "
        "e o cancelamento por coorte, motivo, cidade e tempo de casa. Sem esse dado, decisão de "
        "investimento em retenção continua no escuro. O projeto instala essa inteligência no sistema.",
        align="justify",
        italic=True,
        color=MUT,
    )

    add_heading_styled(doc, "3.2 Conta de gestão (estimativa)", 2)
    add_p(
        doc,
        "Premissas usadas para dimensionar o vazamento — estimativa de gestão, não conciliação contábil, "
        "justamente porque o tracking fino ainda não existe:",
        align="justify",
    )
    make_table(
        doc,
        ["Premissa", "Número", "Leitura"],
        [
            ["CAC + instalação por cliente", "≈ R$ 800", "Custo médio para colocar o cliente dentro"],
            ["Payback mínimo", "8 meses", "Tempo para o cliente se pagar"],
            ["Cancelamentos em 6 meses", "3.755", "Série crescente no 1º semestre"],
            ["Investimento queimado (3.755 × R$ 800)", "≈ R$ 3 milhões", "Aquisição + instalação já desembolsadas"],
            ["Receita potencial não realizada", "≈ R$ 4 milhões", "O que esses clientes deixariam ao ano"],
            ["Vazamento estimado em 6 meses", "≈ R$ 7 milhões", "Investimento perdido + receita deixada na mesa"],
        ],
        [6.2, 3.8, 6.5],
    )

    add_heading_styled(doc, "3.3 Sensibilidade mínima (2%)", 2)
    add_p(
        doc,
        "Se a estruturação retiver apenas 2% da perda — patamar conservador — a G6 segura cerca de "
        "75 clientes. Isso equivale a aproximadamente R$ 60 mil de custo já investido que deixa de "
        "ser perdido e R$ 90 mil de receita em 12 meses: cerca de R$ 150 mil retidos, o suficiente "
        "para cobrir dois meses do contrato de marketing com a V4 e parte desta estruturação.",
        align="justify",
    )
    add_p(
        doc,
        "Enquanto o comercial não tiver processo, o gasto para colocar o cliente dentro continua "
        "exigindo no mínimo 8 meses de permanência — sem visibilidade de quantos saem antes de se pagar.",
        align="justify",
    )

    add_heading_styled(doc, "4. A solução")
    add_p(
        doc,
        "A ordem de construção é sistema → gestão → receita. Sem sistema, meta vira pressão. "
        "Com sistema, meta vira consequência.",
        align="justify",
    )
    make_table(
        doc,
        ["Pilar", "O que instala"],
        [
            ["01 · Sistema", "Processo, CRM e rituais transformam esforço individual em operação previsível."],
            ["02 · Gestão", "Com a casa estruturada, a liderança cobra comportamento e conversão — não improvisa processo."],
            ["03 · Receita", "Ativação, follow-up e retenção entram no método. O CAC deixa de queimar no escuro."],
        ],
        [4.2, 12.3],
    )
    add_p(
        doc,
        "Três frentes andam juntas: sistemas (stack e CRM — manter ou trocar, com implementação inclusa), "
        "processos (jornada loja × PAP, SLA, passagem de bastão venda → operação) e pessoas "
        "(capacitação do time, do Mateus e da direção na nova forma de cobrar).",
        align="justify",
    )

    add_heading_styled(doc, "5. Escopo e entregas")
    make_table(
        doc,
        ["Fundação operacional", "Controle e adesão"],
        [
            ["Arquitetura loja × PAP com critérios de etapa", "CRM como painel de receita (não agenda)"],
            ["Scripts, cadência e papéis claros por canal", "Higiene, motivos de perda e dashboard de uso"],
            ["Rituais de gestão (diária / semanal / pipeline review)", "Treinamento com instalação assistida e roleplay"],
            ["Onboarding operacional de vendedores", "Rotinas de ativação, follow-up e save no processo"],
        ],
        [8.25, 8.25],
    )
    add_p(
        doc,
        "O produto vende capacidade instalada no time. Artefato sem instalação assistida regride "
        "ao comportamento antigo — por isso as fases 1 e 5 são presenciais.",
        align="justify",
    )

    add_heading_styled(doc, "6. Fases do projeto")
    add_p(
        doc,
        "Seis fases, 100,5 horas de execução mais 6 horas de gestão de projeto, QA e alinhamentos internos. Total: 106,5 horas.",
        align="justify",
    )
    make_table(
        doc,
        ["Fase", "Horas", "Nome", "Função"],
        [
            ["01", "18h", "Diagnóstico operacional", "Operação real, presencial. Time fechado, loja + PAP."],
            ["02", "15h", "Arquitetura comercial", "Sistema que o time consegue executar. Dual-track loja × PAP."],
            ["03", "15,5h", "CRM e controle", "Paralelo à Fase 2. Método vira workflow. Brian entra aqui."],
            ["04", "20h", "Camada estratégica", "Playbook, comissão e breakeven. Onboarding do Mateus (~9ª semana)."],
            ["05", "22h", "Instalação da rotina", "Presencial. Imersão, roleplay, rituais assistidos."],
            ["06", "10h", "Handover 60 dias", "Continuidade, retenção e mapa de risco de regressão."],
        ],
        [1.6, 2.0, 4.8, 8.1],
    )

    add_heading_styled(doc, "Fase 1 — Diagnóstico operacional · 18h · presencial", 2)
    add_p(
        doc,
        "Mapear a operação real antes de desenhar playbook: entrevistas com sponsor, gestão, "
        "top e bottom performer; auditoria do CRM/RD; mapa de gargalos por etapa; priorização "
        "das dores que impedem venda previsível.",
        align="justify",
    )

    add_heading_styled(doc, "Fase 2 — Arquitetura operacional comercial · 15h", 2)
    add_p(
        doc,
        "Desenhar o sistema executável: ICP e segmentos, jornada-alvo, pipeline com critérios "
        "de etapa e papéis, SLA marketing ↔ comercial ↔ operação, metas de atividade, scripts, "
        "cadências, BPMN e matriz CHA. Dual-track loja × porta a porta.",
        align="justify",
    )

    add_heading_styled(doc, "Fase 3 — CRM e controle · 15,5h · paralelo à Fase 2", 2)
    add_p(
        doc,
        "Reorganizar o pipeline; campos obrigatórios e motivos de perda; atividades e próxima "
        "ação; dashboard de gestão e conversão (monitor do gestor); regras de higiene. "
        "Campos G6: origem, tempo de resposta, status, dor, capacidade, produto, próxima ação "
        "e motivo de perda. Se o diagnóstico concluir que o CRM atual não serve, a implementação "
        "de outro está inclusa (licenças por conta da G6).",
        align="justify",
    )

    add_heading_styled(doc, "Fase 4 — Camada estratégica · 20h", 2)
    add_p(
        doc,
        "Playbook orientado ao uso, política de comissionamento (loja ≠ PAP), meta como piso e "
        "breakeven amarrado a CAC, payback e retenção. O Mateus é onboardado na filosofia do "
        "projeto, não no jeito antigo.",
        align="justify",
    )

    add_heading_styled(doc, "Fase 5 — Instalação da rotina · 22h · presencial", 2)
    add_p(
        doc,
        "Imersão do time, treino do gestor no CRM, roleplay de loja e PAP, pipeline review e "
        "reunião semanal assistidas, checklist de desvios. Entrega também o pacote de onboarding "
        "para novas contratações.",
        align="justify",
    )

    add_heading_styled(doc, "Fase 6 — Handover e continuidade · 10h", 2)
    add_p(
        doc,
        "Transferência da documentação, score de maturidade antes/depois, rotinas de retenção "
        "e expansão, riscos de regressão e plano tático dos 60 dias seguintes — cobertura de "
        "operação até janeiro, com o Mateus já no comando da máquina.",
        align="justify",
    )

    add_heading_styled(doc, "7. Formato de execução")
    make_table(
        doc,
        ["Item", "Definição"],
        [
            ["Duração", "12 semanas (diagnóstico, construção, instalação e handover)."],
            ["Carga V4", "106,5 horas — análise, desenho, CRM, QA e instalação, não só calls."],
            ["Papéis", "A V4 constrói e conduz a instalação; o time G6 opera no dia a dia."],
            ["Ritmo", "Encontros semanais com sponsor/gestor e checkpoints com o time comercial."],
            ["Presencial", "Fases 1 e 5 — onde a adesão se decide."],
            ["Janela", "Execução em setembro–novembro, mais plano de 60 dias; operação rodando em janeiro."],
            ["Chegada do Mateus", "Os primeiros ~30 dias avançam diagnóstico e arquitetura. Ele entra no treino da máquina já desenhada."],
        ],
        [3.8, 12.7],
    )

    add_heading_styled(doc, "8. Critérios de sucesso")
    add_p(
        doc,
        "A fundação está instalada quando estes pontos estão em uso — não apenas documentados:",
        align="justify",
    )
    make_table(
        doc,
        ["#", "Critério"],
        [
            ["1", "Pipeline configurado e usado"],
            ["2", "Reunião comercial semanal acontecendo"],
            ["3", "Follow-up sendo executado"],
            ["4", "Gestor usando indicadores"],
            ["5", "Time nos scripts e critérios de etapa"],
            ["6", "Onboarding de vendedores documentado"],
            ["7", "Rotinas de retenção em operação"],
            ["8", "Sponsor com previsibilidade de receita"],
        ],
        [1.6, 14.9],
    )

    add_heading_styled(doc, "9. Investimento")
    price = doc.add_table(rows=1, cols=2)
    set_table_width(price, 16.5)
    left, right = price.cell(0, 0), price.cell(0, 1)
    shade_cell(left, DARK)
    shade_cell(right, SOFT)
    set_cell_borders(left, DARK, "4")
    set_cell_borders(right, RED, "8")
    set_cell_margins(left, 160, 160, 160, 160)
    set_cell_margins(right, 140, 140, 140, 140)
    left.text = ""
    p = left.paragraphs[0]
    r = p.add_run("VALOR DE REFERÊNCIA")
    set_run_font(r, size=8, bold=True, color="F5B5B8")
    p2 = left.add_paragraph()
    r = p2.add_run("R$ 89.635,85")
    set_run_font(r, size=16, color="C9A0A0")
    p2.runs[0].font.strike = True
    p3 = left.add_paragraph()
    r = p3.add_run("CONDIÇÃO COMERCIAL G6")
    set_run_font(r, size=8, bold=True, color="F5B5B8")
    p4 = left.add_paragraph()
    r = p4.add_run("R$ 71.582,32")
    set_run_font(r, size=26, bold=True, color=WHITE)
    p5 = left.add_paragraph()
    r = p5.add_run("12 semanas · 106,5h · implementação assistida\nloja × PAP · presenciais nas fases 1 e 5")
    set_run_font(r, size=10, color="E8D0D0")
    right.text = ""
    p = right.paragraphs[0]
    r = p.add_run("CONDIÇÕES")
    set_run_font(r, size=8, bold=True, color=RED)
    add_cell_run(right, "12 recorrências de R$ 5.965,19", size=14, bold=True, color=INK)
    add_cell_run(right, "Pagamento recorrente, sem cartão de crédito — para não competir com a liberação do cartão do contrato de marketing.", size=10, color=MUT)
    add_cell_run(right, "Redução de R$ 18.053,53 sobre a referência, por a G6 ser cliente da base ativa.", size=10, color=MUT)
    add_cell_run(right, "Custo único de implementação. Não se repete no ano 2. O acompanhamento comercial segue no executado de marketing, sem custo adicional desta frente.", size=10, color=MUT)
    spacer = doc.add_paragraph()
    spacer.paragraph_format.space_after = Pt(10)

    add_heading_styled(doc, "9.1 Enquadramento no ano", 2)
    make_table(
        doc,
        ["Linha", "Base", "Total"],
        [
            ["Marketing V4 (já contratado)", "12 × R$ 11.180", "R$ 134.160"],
            ["Estruturação Comercial (esta proposta)", "12 × R$ 5.965,19", "R$ 71.582,32"],
            ["Pacote anual marketing + estruturação", "≈ R$ 17.180 / mês", "≈ R$ 205.700"],
            ["+ mídia (~R$ 10 mil/mês)", "R$ 120.000 no ano", "≈ R$ 324.000"],
            ["Vazamento estimado em 6 meses", "Conta de gestão", "≈ R$ 7.000.000"],
            ["Investimento anual × vazamento", "324 mil / 7 milhões", "Menos de 5%"],
        ],
        [6.2, 5.3, 5.0],
    )
    callout(
        doc,
        "Leitura de decisão",
        "O investimento para instalar previsibilidade é menor do que um mês perdendo cerca de "
        "800 clientes. Adiar a estruturação mantém o furo do balde: o marketing entrega lead "
        "qualificado e a conversão/retenção continua sem sistema.",
    )
    add_p(
        doc,
        "A Estruturação Comercial é etapa única. O marketing de execução (12 meses) renova ou "
        "não ao fim do ciclo. Negociação de valor desta proposta fica com Sara Pizzico.",
        align="justify",
    )

    add_heading_styled(doc, "10. Próximos passos")
    make_table(
        doc,
        ["Responsável", "Ação"],
        [
            ["G6 (Laura / Geraldo)", "Validar a proposta e o investimento para formalizar o start."],
            ["V4 (Denis)", "Kickoff com o detalhamento da Fase 1 presencial (18h)."],
            ["V4 (Denis + Brian)", "Executar as seis fases conforme o cronograma de 12 semanas."],
            ["V4 + G6", "Onboardar o Mateus na máquina já em construção."],
        ],
        [5.0, 11.5],
    )
    add_p(
        doc,
        "Documento elaborado pela V4 Company para a G6 Internet. Valores e prazos conforme "
        "condição comercial da base ativa. A conta de R$ 7 milhões é estimativa de gestão, "
        "apresentada assim porque o churn real ainda não está conciliado — exatamente o que "
        "este projeto se propõe a instalar no sistema.",
        size=9,
        color=MUT,
        italic=True,
        space_before=12,
    )
    add_p(
        doc,
        "G6 Internet × V4 Company  ·  Estruturação Comercial  ·  Confidencial",
        size=9,
        color=RED,
        bold=True,
        align="center",
        space_before=8,
    )
    doc.save(OUT_DOCX)
    print("Wrote", OUT_DOCX)


def html_table(headers, rows):
    head = "".join(f"<th>{escape(h)}</th>" for h in headers)
    body = []
    for row in rows:
        cells = "".join(
            f'<td style="{"font-weight:700;" if i == 0 else ""}">{escape(val)}</td>'
            for i, val in enumerate(row)
        )
        body.append(f"<tr>{cells}</tr>")
    return (
        '<table class="grid" border="1" cellspacing="0" cellpadding="0">'
        f"<thead><tr>{head}</tr></thead><tbody>{''.join(body)}</tbody></table>"
    )


def build_html():
    parts = []
    parts.append(
        """<table class="cover" border="0" cellspacing="0" cellpadding="0"><tr><td>
<p class="cover-kicker">PROPOSTA COMERCIAL</p>
<p class="cover-title">Estruturação Comercial</p>
<p class="cover-sub">G6 Internet  ×  V4 Company</p>
<p class="cover-lead">Projeto para instalar a fundação comercial da G6: processo, CRM em uso, gestão e rotina — loja e porta a porta — em 12 semanas.</p>
</td></tr></table>"""
    )
    parts.append(
        html_table(
            ["Campo", "Informação"],
            [
                ["Cliente", "G6 Internet"],
                ["Produto", "Estruturação Comercial (Sales V4) — projeto fechado com implementação assistida"],
                ["Prazo", "12 semanas  ·  106,5 horas  ·  dual-track loja × porta a porta"],
                ["Investimento", "R$ 71.582,32  (referência: R$ 89.635,85)  ·  12 parcelas recorrentes"],
                ["Condução", "V4 Company  ·  Denis Orosco (receita / Sales)  ·  Brian (CRM)"],
                ["Sponsor G6", "Laurelise Santos  ·  validação societária com Geraldo"],
                ["Gestor comercial", "Mateus — entrada prevista em cerca de 30 dias"],
            ],
        )
    )
    parts.append(
        '<table class="callout" border="0" cellspacing="0" cellpadding="0"><tr><td>'
        '<p class="callout-label">LEITURA CENTRAL</p>'
        "<p>A G6 não tem problema de lead. Tem um comercial reativo, sem sistema. "
        "A demanda entra; o processo não segura. Esta proposta instala a máquina "
        "para vender, ativar e reter com previsibilidade — e entrega essa máquina "
        "pronta para o novo gestor operar.</p>"
        "</td></tr></table>"
    )

    def h1(t):
        return f"<h1>{escape(t)}</h1>"

    def h2(t):
        return f"<h2>{escape(t)}</h2>"

    def p(t, cls="justify"):
        return f'<p class="{cls}">{escape(t)}</p>'

    def ul(items):
        return "<ul>" + "".join(f"<li>{escape(i)}</li>" for i in items) + "</ul>"

    parts += [
        h1("1. Objetivo"),
        p(
            "Construir e instalar a fundação comercial da G6 Internet em um projeto fechado de "
            "12 semanas: diagnóstico presencial, arquitetura de processo, CRM como painel de receita, "
            "camada estratégica (playbook, comissão e breakeven), imersão/treino do time e handover "
            "com plano de continuidade de 60 dias."
        ),
        p(
            "O resultado esperado não é um documento para a gaveta. É a operação rodando: loja e "
            "porta a porta com método, rituais de gestão e o Mateus herdando uma máquina pronta — "
            "em vez de improvisar processo, troca de peça e meta ao mesmo tempo."
        ),
        h1("2. Diagnóstico da operação"),
        p(
            "A G6 gera demanda. Meta Forms entregou 1.269 leads no ano com CPL de R$ 7,51. "
            "O WhatsApp responde em cerca de 1 minuto. A operação cita cerca de 1.000 ativações "
            "por mês. O que falta é sistema para converter, ativar, seguir e reter."
        ),
        html_table(
            ["Sintoma", "Efeito na operação"],
            [
                ["Volume alto, retenção frágil", "Entra cerca de 1.000 clientes e sai uma faixa próxima de 800. O crescimento líquido vaza todo mês."],
                ["CRM existe, mas não governa", "Sem motivo de perda, sem próximo passo, sem previsibilidade de pipeline. Não funciona como painel de receita."],
                ["Meta vira comodidade", "Quando aperta, a meta desce para o time “bater” e receber comissão. A liderança cascateia acomodação."],
                ["Loja e PAP sem padrão", "5 internos + 2 porta a porta. Cada canal vende de um jeito, sem playbook, deparo nem objeção mapeada."],
                ["Pós-venda solto", "Vendeu e acabou. Ativação, save, follow-up e upsell na base ficam fora do processo."],
                ["Gestão sem ritual pronto", "O novo gestor precisa herdar reunião comercial, indicadores e critérios — não inventar do zero."],
            ],
        ),
        '<p class="label">Pontos da operação que o projeto precisa endereçar</p>',
        ul(
            [
                "Time comercial de 7 pessoas, mais o Mateus na chegada.",
                "Porta a porta sem deparo: região, oportunidades, o que trouxe, o que não trouxe, objeções.",
                "Atendimento inicial longo demais para um produto de alta intenção (primeiro plano de internet).",
                "Base instalada sem jornada de upgrade — o time olha só a venda nova.",
                "Casos de instalação sem pagamento/ativação: comissão apagada e receita nenhuma.",
                "Cultura de atendimento forte; o comercial herdou o ritmo do atendimento e ficou reativo.",
                "Chegada do serviço móvel/chip: sem time treinado para vender na base, a G6 gasta de novo para ofertar um produto à carteira que já pagou para entrar.",
            ]
        ),
        '<table class="callout" border="0" cellspacing="0" cellpadding="0"><tr><td>'
        '<p class="callout-label">POR QUE O DIAGNÓSTICO É PRESENCIAL</p>'
        "<p>A Fase 1 acontece in loco (18 horas) para mapear cultura, resistência à mudança, loja, "
        "porta a porta e o dia a dia da unidade. Sem isso, o plano fica genérico e não cola na G6.</p>"
        "</td></tr></table>",
        h1("3. Impacto financeiro"),
        p(
            "A G6 tem capacidade de vender. O que a operação ainda não tem é capacidade de segurar — "
            "nem de medir, com rigor, por que perde."
        ),
        h2("3.1 Cancelamentos no 1º semestre de 2026"),
        html_table(
            ["Mês", "Cancelamentos"],
            [
                ["Janeiro", "415"],
                ["Fevereiro", "560"],
                ["Março", "614"],
                ["Abril", "746"],
                ["Maio", "715"],
                ["Junho", "705"],
                ["Total jan–jun", "3.755"],
            ],
        ),
        p(
            "A taxa real de churn ainda não está conciliada. Falta carteira ativa no início de cada mês "
            "e o cancelamento por coorte, motivo, cidade e tempo de casa. Sem esse dado, decisão de "
            "investimento em retenção continua no escuro. O projeto instala essa inteligência no sistema.",
            "muted",
        ),
        h2("3.2 Conta de gestão (estimativa)"),
        p(
            "Premissas usadas para dimensionar o vazamento — estimativa de gestão, não conciliação contábil, "
            "justamente porque o tracking fino ainda não existe:"
        ),
        html_table(
            ["Premissa", "Número", "Leitura"],
            [
                ["CAC + instalação por cliente", "≈ R$ 800", "Custo médio para colocar o cliente dentro"],
                ["Payback mínimo", "8 meses", "Tempo para o cliente se pagar"],
                ["Cancelamentos em 6 meses", "3.755", "Série crescente no 1º semestre"],
                ["Investimento queimado (3.755 × R$ 800)", "≈ R$ 3 milhões", "Aquisição + instalação já desembolsadas"],
                ["Receita potencial não realizada", "≈ R$ 4 milhões", "O que esses clientes deixariam ao ano"],
                ["Vazamento estimado em 6 meses", "≈ R$ 7 milhões", "Investimento perdido + receita deixada na mesa"],
            ],
        ),
        h2("3.3 Sensibilidade mínima (2%)"),
        p(
            "Se a estruturação retiver apenas 2% da perda — patamar conservador — a G6 segura cerca de "
            "75 clientes. Isso equivale a aproximadamente R$ 60 mil de custo já investido que deixa de "
            "ser perdido e R$ 90 mil de receita em 12 meses: cerca de R$ 150 mil retidos, o suficiente "
            "para cobrir dois meses do contrato de marketing com a V4 e parte desta estruturação."
        ),
        p(
            "Enquanto o comercial não tiver processo, o gasto para colocar o cliente dentro continua "
            "exigindo no mínimo 8 meses de permanência — sem visibilidade de quantos saem antes de se pagar."
        ),
        h1("4. A solução"),
        p("A ordem de construção é sistema → gestão → receita. Sem sistema, meta vira pressão. Com sistema, meta vira consequência."),
        html_table(
            ["Pilar", "O que instala"],
            [
                ["01 · Sistema", "Processo, CRM e rituais transformam esforço individual em operação previsível."],
                ["02 · Gestão", "Com a casa estruturada, a liderança cobra comportamento e conversão — não improvisa processo."],
                ["03 · Receita", "Ativação, follow-up e retenção entram no método. O CAC deixa de queimar no escuro."],
            ],
        ),
        p(
            "Três frentes andam juntas: sistemas (stack e CRM — manter ou trocar, com implementação inclusa), "
            "processos (jornada loja × PAP, SLA, passagem de bastão venda → operação) e pessoas "
            "(capacitação do time, do Mateus e da direção na nova forma de cobrar)."
        ),
        h1("5. Escopo e entregas"),
        html_table(
            ["Fundação operacional", "Controle e adesão"],
            [
                ["Arquitetura loja × PAP com critérios de etapa", "CRM como painel de receita (não agenda)"],
                ["Scripts, cadência e papéis claros por canal", "Higiene, motivos de perda e dashboard de uso"],
                ["Rituais de gestão (diária / semanal / pipeline review)", "Treinamento com instalação assistida e roleplay"],
                ["Onboarding operacional de vendedores", "Rotinas de ativação, follow-up e save no processo"],
            ],
        ),
        p(
            "O produto vende capacidade instalada no time. Artefato sem instalação assistida regride "
            "ao comportamento antigo — por isso as fases 1 e 5 são presenciais."
        ),
        h1("6. Fases do projeto"),
        p("Seis fases, 100,5 horas de execução mais 6 horas de gestão de projeto, QA e alinhamentos internos. Total: 106,5 horas."),
        html_table(
            ["Fase", "Horas", "Nome", "Função"],
            [
                ["01", "18h", "Diagnóstico operacional", "Operação real, presencial. Time fechado, loja + PAP."],
                ["02", "15h", "Arquitetura comercial", "Sistema que o time consegue executar. Dual-track loja × PAP."],
                ["03", "15,5h", "CRM e controle", "Paralelo à Fase 2. Método vira workflow. Brian entra aqui."],
                ["04", "20h", "Camada estratégica", "Playbook, comissão e breakeven. Onboarding do Mateus (~9ª semana)."],
                ["05", "22h", "Instalação da rotina", "Presencial. Imersão, roleplay, rituais assistidos."],
                ["06", "10h", "Handover 60 dias", "Continuidade, retenção e mapa de risco de regressão."],
            ],
        ),
        h2("Fase 1 — Diagnóstico operacional · 18h · presencial"),
        p(
            "Mapear a operação real antes de desenhar playbook: entrevistas com sponsor, gestão, "
            "top e bottom performer; auditoria do CRM/RD; mapa de gargalos por etapa; priorização "
            "das dores que impedem venda previsível."
        ),
        h2("Fase 2 — Arquitetura operacional comercial · 15h"),
        p(
            "Desenhar o sistema executável: ICP e segmentos, jornada-alvo, pipeline com critérios "
            "de etapa e papéis, SLA marketing ↔ comercial ↔ operação, metas de atividade, scripts, "
            "cadências, BPMN e matriz CHA. Dual-track loja × porta a porta."
        ),
        h2("Fase 3 — CRM e controle · 15,5h · paralelo à Fase 2"),
        p(
            "Reorganizar o pipeline; campos obrigatórios e motivos de perda; atividades e próxima "
            "ação; dashboard de gestão e conversão (monitor do gestor); regras de higiene. "
            "Campos G6: origem, tempo de resposta, status, dor, capacidade, produto, próxima ação "
            "e motivo de perda. Se o diagnóstico concluir que o CRM atual não serve, a implementação "
            "de outro está inclusa (licenças por conta da G6)."
        ),
        h2("Fase 4 — Camada estratégica · 20h"),
        p(
            "Playbook orientado ao uso, política de comissionamento (loja ≠ PAP), meta como piso e "
            "breakeven amarrado a CAC, payback e retenção. O Mateus é onboardado na filosofia do "
            "projeto, não no jeito antigo."
        ),
        h2("Fase 5 — Instalação da rotina · 22h · presencial"),
        p(
            "Imersão do time, treino do gestor no CRM, roleplay de loja e PAP, pipeline review e "
            "reunião semanal assistidas, checklist de desvios. Entrega também o pacote de onboarding "
            "para novas contratações."
        ),
        h2("Fase 6 — Handover e continuidade · 10h"),
        p(
            "Transferência da documentação, score de maturidade antes/depois, rotinas de retenção "
            "e expansão, riscos de regressão e plano tático dos 60 dias seguintes — cobertura de "
            "operação até janeiro, com o Mateus já no comando da máquina."
        ),
        h1("7. Formato de execução"),
        html_table(
            ["Item", "Definição"],
            [
                ["Duração", "12 semanas (diagnóstico, construção, instalação e handover)."],
                ["Carga V4", "106,5 horas — análise, desenho, CRM, QA e instalação, não só calls."],
                ["Papéis", "A V4 constrói e conduz a instalação; o time G6 opera no dia a dia."],
                ["Ritmo", "Encontros semanais com sponsor/gestor e checkpoints com o time comercial."],
                ["Presencial", "Fases 1 e 5 — onde a adesão se decide."],
                ["Janela", "Execução em setembro–novembro, mais plano de 60 dias; operação rodando em janeiro."],
                ["Chegada do Mateus", "Os primeiros ~30 dias avançam diagnóstico e arquitetura. Ele entra no treino da máquina já desenhada."],
            ],
        ),
        h1("8. Critérios de sucesso"),
        p("A fundação está instalada quando estes pontos estão em uso — não apenas documentados:"),
        html_table(
            ["#", "Critério"],
            [
                ["1", "Pipeline configurado e usado"],
                ["2", "Reunião comercial semanal acontecendo"],
                ["3", "Follow-up sendo executado"],
                ["4", "Gestor usando indicadores"],
                ["5", "Time nos scripts e critérios de etapa"],
                ["6", "Onboarding de vendedores documentado"],
                ["7", "Rotinas de retenção em operação"],
                ["8", "Sponsor com previsibilidade de receita"],
            ],
        ),
        h1("9. Investimento"),
        """<table class="price" border="0" cellspacing="0" cellpadding="0">
<tr>
<td class="price-left" width="50%">
<p class="cover-kicker">VALOR DE REFERÊNCIA</p>
<p class="strike">R$ 89.635,85</p>
<p class="cover-kicker">CONDIÇÃO COMERCIAL G6</p>
<p class="price-now">R$ 71.582,32</p>
<p class="cover-lead">12 semanas · 106,5h · implementação assistida<br>loja × PAP · presenciais nas fases 1 e 5</p>
</td>
<td class="price-right" width="50%">
<p class="callout-label">CONDIÇÕES</p>
<p class="price-month">12 recorrências de R$ 5.965,19</p>
<p>Pagamento recorrente, sem cartão de crédito — para não competir com a liberação do cartão do contrato de marketing.</p>
<p>Redução de R$ 18.053,53 sobre a referência, por a G6 ser cliente da base ativa.</p>
<p>Custo único de implementação. Não se repete no ano 2. O acompanhamento comercial segue no executado de marketing, sem custo adicional desta frente.</p>
</td>
</tr>
</table>""",
        h2("9.1 Enquadramento no ano"),
        html_table(
            ["Linha", "Base", "Total"],
            [
                ["Marketing V4 (já contratado)", "12 × R$ 11.180", "R$ 134.160"],
                ["Estruturação Comercial (esta proposta)", "12 × R$ 5.965,19", "R$ 71.582,32"],
                ["Pacote anual marketing + estruturação", "≈ R$ 17.180 / mês", "≈ R$ 205.700"],
                ["+ mídia (~R$ 10 mil/mês)", "R$ 120.000 no ano", "≈ R$ 324.000"],
                ["Vazamento estimado em 6 meses", "Conta de gestão", "≈ R$ 7.000.000"],
                ["Investimento anual × vazamento", "324 mil / 7 milhões", "Menos de 5%"],
            ],
        ),
        '<table class="callout" border="0" cellspacing="0" cellpadding="0"><tr><td>'
        '<p class="callout-label">LEITURA DE DECISÃO</p>'
        "<p>O investimento para instalar previsibilidade é menor do que um mês perdendo cerca de "
        "800 clientes. Adiar a estruturação mantém o furo do balde: o marketing entrega lead "
        "qualificado e a conversão/retenção continua sem sistema.</p>"
        "</td></tr></table>",
        p(
            "A Estruturação Comercial é etapa única. O marketing de execução (12 meses) renova ou "
            "não ao fim do ciclo. Negociação de valor desta proposta fica com Sara Pizzico."
        ),
        h1("10. Próximos passos"),
        html_table(
            ["Responsável", "Ação"],
            [
                ["G6 (Laura / Geraldo)", "Validar a proposta e o investimento para formalizar o start."],
                ["V4 (Denis)", "Kickoff com o detalhamento da Fase 1 presencial (18h)."],
                ["V4 (Denis + Brian)", "Executar as seis fases conforme o cronograma de 12 semanas."],
                ["V4 + G6", "Onboardar o Mateus na máquina já em construção."],
            ],
        ),
        p(
            "Documento elaborado pela V4 Company para a G6 Internet. Valores e prazos conforme "
            "condição comercial da base ativa. A conta de R$ 7 milhões é estimativa de gestão, "
            "apresentada assim porque o churn real ainda não está conciliado — exatamente o que "
            "este projeto se propõe a instalar no sistema.",
            "muted",
        ),
        '<p class="footer-brand">G6 Internet × V4 Company  ·  Estruturação Comercial  ·  Confidencial</p>',
    ]

    body = "\n".join(parts)
    html = f"""<!DOCTYPE html>
<html xmlns:o="urn:schemas-microsoft-com:office:office"
      xmlns:w="urn:schemas-microsoft-com:office:word"
      xmlns="http://www.w3.org/TR/REC-html40">
<head>
<meta charset="utf-8">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<meta name="ProgId" content="Word.Document">
<meta name="Generator" content="Microsoft Word 15">
<title>G6 Internet × V4 — Proposta de Estruturação Comercial</title>
<!--[if gte mso 9]>
<xml>
  <w:WordDocument>
    <w:View>Print</w:View>
    <w:Zoom>100</w:Zoom>
    <w:DoNotOptimizeForBrowser/>
  </w:WordDocument>
</xml>
<![endif]-->
<style>
  @page WordSection1 {{ size: 21cm 29.7cm; margin: 1.8cm 2cm 2cm 2cm; }}
  div.WordSection1 {{ page: WordSection1; }}
  body {{
    font-family: Calibri, Arial, sans-serif; font-size: 11pt; color: #1A1A1A;
    line-height: 1.35; margin: 24px auto; max-width: 820px; background: #f4f0ef;
  }}
  div.WordSection1 {{ background: #fff; padding: 28px 32px 48px; }}
  h1 {{ font-size: 16pt; color: #A10F14; border-bottom: 1.5pt solid #A10F14; padding-bottom: 4pt; margin: 18pt 0 8pt; }}
  h2 {{ font-size: 12.5pt; color: #E50914; margin: 12pt 0 6pt; }}
  p {{ margin: 0 0 8pt; }}
  p.justify {{ text-align: justify; }}
  p.muted {{ color: #4A4A4A; font-size: 10.5pt; font-style: italic; }}
  p.label {{ color: #E50914; font-weight: 700; margin: 8pt 0 4pt; }}
  p.footer-brand {{ color: #A10F14; font-weight: 700; text-align: center; font-size: 9pt; margin-top: 16pt; }}
  ul {{ margin: 4pt 0 10pt 18pt; padding: 0; }}
  li {{ margin: 0 0 3pt; color: #4A4A4A; }}
  table.grid {{ width: 100%; border-collapse: collapse; margin: 8pt 0 12pt; font-size: 10pt; }}
  table.grid th {{ background: #A10F14; color: #fff; text-align: left; font-size: 9pt; padding: 6pt 8pt; border: 1pt solid #A10F14; }}
  table.grid td {{ border: 1pt solid #E6D4D4; padding: 6pt 8pt; vertical-align: top; }}
  table.callout {{ width: 100%; margin: 8pt 0 12pt; border-left: 4pt solid #A10F14; background: #F8EFEF; }}
  table.callout td {{ padding: 10pt 12pt; }}
  .callout-label {{ color: #A10F14; font-size: 8.5pt; font-weight: 700; letter-spacing: .08em; margin: 0 0 4pt; }}
  table.cover {{ width: 100%; background: #280001; margin: 0 0 14pt; }}
  table.cover td {{ padding: 22pt 20pt; }}
  .cover-kicker {{ color: #F5B5B8; font-size: 9pt; font-weight: 700; letter-spacing: .12em; margin: 0 0 8pt; }}
  .cover-title {{ color: #fff; font-size: 26pt; font-weight: 700; margin: 0 0 6pt; }}
  .cover-sub {{ color: #FFD0D0; font-size: 14pt; margin: 0 0 10pt; }}
  .cover-lead {{ color: #E8D0D0; font-size: 11pt; margin: 0; }}
  table.price {{ width: 100%; margin: 8pt 0 12pt; }}
  td.price-left {{ background: #280001; color: #fff; padding: 16pt; vertical-align: top; width: 50%; }}
  td.price-right {{ background: #F8EFEF; border: 1.5pt solid #A10F14; padding: 14pt; vertical-align: top; width: 50%; }}
  .strike {{ color: #C9A0A0; text-decoration: line-through; font-size: 14pt; margin: 0 0 8pt; }}
  .price-now {{ color: #fff; font-size: 24pt; font-weight: 700; margin: 0 0 8pt; }}
  .price-month {{ font-size: 13pt; font-weight: 700; margin: 0 0 8pt; }}
</style>
</head>
<body>
<div class="WordSection1">
{body}
</div>
</body>
</html>
"""
    OUT_HTML.write_text(html, encoding="utf-8")
    OUT_DOC.write_text(html, encoding="utf-8")
    print("Wrote", OUT_HTML)
    print("Wrote", OUT_DOC)


def publish_g6_word():
    G6_WORD.mkdir(exist_ok=True)
    copy2(OUT_DOCX, G6_WORD / "G6-Internet-Estruturacao-Comercial.docx")
    copy2(OUT_DOC, G6_WORD / "G6-Internet-Estruturacao-Comercial.doc")
    print("Copied into", G6_WORD)


if __name__ == "__main__":
    build_docx()
    build_html()
    publish_g6_word()
