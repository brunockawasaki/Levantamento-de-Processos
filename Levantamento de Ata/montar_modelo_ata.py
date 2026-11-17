# -*- coding: utf-8 -*-
"""Gera os modelos de Ata de Reuniao a partir do 'Mod Pauta-sumula de Reuniao' (.doc/.docx).

O modelo e um formulario feito so de tabelas, sem logo. Este script refaz o formulario em .docx
(campos, ordem, largura das colunas e margens do modelo), em Arial 10 (regra PO.GG.QLD.002).

Uso:
  python montar_modelo_ata.py <saida_com_marcadores.docx> [<saida_modelo_texto_guia.docx>]

Campos do modelo (na ordem):
  Ata de Reuniao (titulo, fundo cinza)
  Departamento: | Assunto / Objetivo: | Local / Data:  Responsavel:
  Participantes: (2 colunas)
  Pauta da Reuniao:
  Ata da Reuniao
  Pendencias | Responsavel | Data | OK
  Data da Expedicao: | Registro feito por:
  Data da proxima reuniao:
  Documentos Anexos:

Marcadores: {{DEPARTAMENTO}} {{ASSUNTO}} {{LOCAL_DATA}} {{RESPONSAVEL}} {{PART1}} {{PART2}} (linha repetida)
            {{PAUTA}} {{SUMULA}} {{ACAO}} {{ACAO_RESP}} {{ACAO_DATA}} {{ACAO_OK}} (linha repetida)
            {{EXPEDICAO}} {{REGISTRO}} {{PROXIMA}} {{ANEXOS}}
O 2o arquivo de saida e o modelo em branco, com texto-guia no lugar dos marcadores.
"""
import sys
import docx
from docx.shared import Pt, Twips
from docx.oxml import parse_xml
from docx.oxml.ns import qn

W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
TOTAL = 10690  # largura das tabelas do modelo (twips)

GUIA = {
    "{{DEPARTAMENTO}}": "Informar o departamento ou setor responsável pela reunião.",
    "{{ASSUNTO}}": "Informar o assunto tratado e o objetivo da reunião.",
    "{{LOCAL_DATA}}": "Informar o local e a data da reunião.",
    "{{RESPONSAVEL}}": "Informar quem conduziu a reunião.",
    "{{PART1}}": "Nome do participante",
    "{{PART2}}": "Nome do participante",
    "{{PAUTA}}": "Listar os assuntos previstos para a reunião, na ordem em que serão tratados.",
    "{{SUMULA}}": "Registrar o que foi discutido e o que foi decidido em cada item da pauta, de forma objetiva.",
    "{{EXPEDICAO}}": "DD/MM/AAAA",
    "{{REGISTRO}}": "Nome de quem redigiu a ata",
    "{{PROXIMA}}": "DD/MM/AAAA",
    "{{ANEXOS}}": "Relacionar os documentos anexados, quando houver.",
}


def esc(t):
    return t.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


def rpr(bold=False, size=20):
    return ('<w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/>%s<w:color w:val="000000"/>'
            '<w:sz w:val="%d"/><w:szCs w:val="%d"/></w:rPr>' % ("<w:b/><w:bCs/>" if bold else "", size, size))


def run(text, bold=False, size=20):
    return '<w:r>%s<w:t xml:space="preserve">%s</w:t></w:r>' % (rpr(bold, size), esc(text))


def para(runs, jc="left", before=40, after=40, keep=False):
    return ('<w:p><w:pPr>%s<w:spacing w:before="%d" w:after="%d" w:line="240" w:lineRule="auto"/><w:jc w:val="%s"/></w:pPr>%s</w:p>'
            % ("<w:keepNext/>" if keep else "", before, after, jc, "".join(runs)))


def spacer():
    return parse_xml('<w:p %s><w:pPr><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/></w:pPr>'
                     '<w:r>%s</w:r></w:p>' % (W, rpr()))


def cell(width, paras, shade=None, span=None):
    sh = '<w:shd w:val="clear" w:color="auto" w:fill="%s"/>' % shade if shade else ""
    gs = '<w:gridSpan w:val="%d"/>' % span if span else ""
    return '<w:tc><w:tcPr><w:tcW w:w="%d" w:type="dxa"/>%s%s</w:tcPr>%s</w:tc>' % (width, gs, sh, "".join(paras))


def row(cells, height=None, header=False):
    pr = ""
    if height or header:
        pr = "<w:trPr><w:cantSplit/>%s%s</w:trPr>" % (
            '<w:trHeight w:val="%d" w:hRule="atLeast"/>' % height if height else "",
            "<w:tblHeader/>" if header else "")
    return "<w:tr>%s%s</w:tr>" % (pr, "".join(cells))


def table(widths, rows):
    borders = "".join('<w:%s w:val="single" w:sz="4" w:space="0" w:color="000000"/>' % s
                      for s in ("top", "left", "bottom", "right", "insideH", "insideV"))
    return parse_xml(
        '<w:tbl %s><w:tblPr><w:tblW w:w="%d" w:type="dxa"/><w:tblBorders>%s</w:tblBorders><w:tblLayout w:type="fixed"/>'
        '<w:tblCellMar><w:left w:w="100" w:type="dxa"/><w:right w:w="100" w:type="dxa"/></w:tblCellMar></w:tblPr>'
        '<w:tblGrid>%s</w:tblGrid>%s</w:tbl>' % (W, sum(widths), borders,
                                                  "".join('<w:gridCol w:w="%d"/>' % w for w in widths), "".join(rows)))


def build(dst, guia):
    d = docx.Document()
    sec = d.sections[0]
    sec.page_width, sec.page_height = Twips(12240), Twips(15840)  # carta, como o modelo
    sec.top_margin, sec.bottom_margin = Twips(1417), Twips(1258)
    sec.left_margin, sec.right_margin = Twips(900), Twips(540)
    st = d.styles["Normal"]
    st.font.name = "Arial"
    st.font.size = Pt(10)
    st.element.rPr.rFonts.set(qn("w:eastAsia"), "Arial")
    st.paragraph_format.space_after = Pt(0)
    sect = d.element.body.find(qn("w:sectPr"))

    def val(token):
        return GUIA[token] if guia else token

    def labeled(width, label, token, span=None, bold=False, jc="left"):
        return cell(width, [para([run(label + " ", bold), run(val(token))], jc=jc)], span=span)

    def titulo(width, text, size=20, span=None):
        return cell(width, [para([run(text, True, size)], jc="center", keep=True)], span=span)

    # --- titulo (fundo cinza)
    t1 = table([TOTAL], [row([cell(TOTAL, [para([run("Ata de Reunião", False, 32)], jc="center", before=100, after=100)],
                                   shade="E6E6E6")], 700)])
    # --- dados da reuniao
    t2 = table([6550, 4140], [
        row([labeled(TOTAL, "Departamento:", "{{DEPARTAMENTO}}", span=2, bold=True, jc="center")], 280),
        row([labeled(TOTAL, "Assunto / Objetivo:", "{{ASSUNTO}}", span=2)], 280),
        row([labeled(6550, "Local / Data:", "{{LOCAL_DATA}}"), labeled(4140, "Responsável:", "{{RESPONSAVEL}}")], 280),
    ])
    # --- participantes (2 colunas, uma linha por par de nomes)
    if guia:
        linhas_p = [[val("{{PART1}}"), val("{{PART2}}")]] * 2
    else:
        linhas_p = [["{{PART1}}", "{{PART2}}"]]
    t3 = table([5110, 5580], [row([titulo(TOTAL, "Participantes:", span=2)], 280)] +
               [row([cell(5110, [para([run(a)])]), cell(5580, [para([run(b)])])], 280) for a, b in linhas_p])
    # --- pauta e ata
    t4 = table([TOTAL], [
        row([titulo(TOTAL, "Pauta da Reunião:")], 280),
        row([cell(TOTAL, [para([run(val("{{PAUTA}}"))])])], 1100),
    ])
    t5 = table([TOTAL], [
        row([titulo(TOTAL, "Ata da Reunião", 24)], 280),
        row([cell(TOTAL, [para([run(val("{{SUMULA}}"))])])], 1100),
    ])
    # --- pendencias
    W4 = [6190, 1980, 1800, 720]
    cab = row([cell(w, [para([run(t, True)])]) for w, t in zip(W4, ["Pendências", "Responsável", "Data", "OK"])], 280, header=True)
    if guia:
        linhas = [["Ação ou output definido na reunião", "Nome do responsável", "DD/MM/AAAA", ""]] * 2 + [["", "", "", ""]]
    else:
        linhas = [["{{ACAO}}", "{{ACAO_RESP}}", "{{ACAO_DATA}}", "{{ACAO_OK}}"]]
    t6 = table(W4, [cab] + [row([cell(w, [para([run(v)])]) for w, v in zip(W4, ln)], 280) for ln in linhas])
    # --- expedicao / registro, proxima reuniao, anexos
    t7 = table([5290, 5400], [
        row([labeled(5290, "Data da Expedição:", "{{EXPEDICAO}}"), labeled(5400, "Registro feito por:", "{{REGISTRO}}")], 1000),
        row([labeled(TOTAL, "Data da próxima reunião:", "{{PROXIMA}}", span=2)], 280),
        row([cell(TOTAL, [para([run("Documentos Anexos:")], keep=True), para([run(val("{{ANEXOS}}"))])], span=2)], 280),
    ])
    for t in (t1, t2, t3, t4, t5, t6, t7):
        sect.addprevious(t)
        sect.addprevious(spacer())

    d.core_properties.title = "Modelo de Ata de Reunião - Grupo Gadens"
    d.save(dst)
    print("ok ->", dst)


if __name__ == "__main__":
    if len(sys.argv) not in (2, 3):
        raise SystemExit(__doc__)
    build(sys.argv[1], False)
    if len(sys.argv) == 3:
        build(sys.argv[2], True)
