# -*- coding: utf-8 -*-
"""Gera o modelo de Politica de Departamento com marcadores {{...}} a partir do template 'GG00 - Modelo de Politica de Departamento.docx'.

Nao mexe na formatacao do template (estilos, numeracao, logo, sumario, tabela de revisao): so troca os textos de exemplo
por marcadores e deixa blocos repetiveis (subitens do item 4, anexos, termos, praticas obrigatorias).
O numero da politica (GGXX) fica fixo, com o XX em vermelho: quem define a numeracao e o Compliance.

Uso:
  python montar_modelo_politica.py "<GG00 - Modelo de Politica de Departamento.docx>" "<MODELO POLITICA com marcadores.docx>"

Marcadores (corpo): {{NOME_PROCESSO}} {{NOME_PROC_MAI}} {{OBJETIVO}} {{APLICACAO}} {{RESPONSAVEL}} {{VISAO_GERAL}}
  {{SUB_TITULO}}+{{SUB_TEXTO}} (bloco repetido) {{ANEXO}} {{TERMO}}+{{DEFINICAO}} (paragrafo repetido)
  {{PRATICA}}+{{PORQUE}} (bloco repetido) {{DUVIDAS}} {{ELAB_EM}} {{ELAB_POR}} {{REV_EM}} {{REV_POR}} {{APR_EM}} {{APR_POR}}
Cabecalho: {{NOME_PROCESSO}} {{STATUS}} {{PROXIMA_REVISAO}}      Rodape: {{ELAB_POR}} {{REV_POR}} {{APR_POR}} {{APR_EM}}
"""
import sys, zipfile, copy
from lxml import etree

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"
NS = {"w": W}


def q(tag):
    return "{%s}%s" % (W, tag)


def texto(el):
    return "".join(t.text or "" for t in el.iter(q("t")))


def limpa_paragrafo(p):
    """Tira tudo do paragrafo, menos o pPr; devolve o rPr da 1a run (formatacao do texto)."""
    ppr = p.find(q("pPr"))
    run = p.find(q("r"))
    rpr = copy.deepcopy(run.find(q("rPr"))) if run is not None and run.find(q("rPr")) is not None else None
    for ch in list(p):
        if ch is not ppr:
            p.remove(ch)
    return rpr


def poe_texto(p, txt):
    rpr = limpa_paragrafo(p)
    r = etree.SubElement(p, q("r"))
    if rpr is not None:
        r.append(rpr)
    t = etree.SubElement(r, q("t"))
    t.text = txt
    t.set(XML_SPACE, "preserve")
    return r


def poe_negrito(run):
    rpr = run.find(q("rPr"))
    if rpr is None:
        rpr = etree.Element(q("rPr")); run.insert(0, rpr)
    if rpr.find(q("b")) is None:
        pos = 1 if rpr.find(q("rFonts")) is not None else 0
        rpr.insert(pos, etree.Element(q("bCs"))); rpr.insert(pos, etree.Element(q("b")))


def confere(el, comeco):
    assert texto(el).strip().startswith(comeco), (comeco, texto(el)[:60])


def monta(origem, destino):
    zin = zipfile.ZipFile(origem)
    doc = etree.fromstring(zin.read("word/document.xml"))
    cab = etree.fromstring(zin.read("word/header1.xml"))
    rod = etree.fromstring(zin.read("word/footer1.xml"))
    body = doc.find(q("body"))
    E = list(body)

    # ---------- corpo ----------
    confere(E[0], "Nome do Processo Completo"); poe_texto(E[0], "{{NOME_PROCESSO}}")
    confere(E[8], "Objetivo desta pol");        poe_texto(E[8], "{{OBJETIVO}}")
    confere(E[11], "EX. Esta Pol");             poe_texto(E[11], "{{APLICACAO}}")
    confere(E[14], "EX. Processo realizado");   poe_texto(E[14], "{{RESPONSAVEL}}")

    # item 4: visao geral + subitens (4.1, 4.2...) em bloco repetivel
    confere(E[17], "Delinear o processo Macro")
    modelo_texto = copy.deepcopy(E[17])
    poe_texto(E[17], "{{VISAO_GERAL}}")
    sub_titulo = copy.deepcopy(modelo_texto); poe_negrito(poe_texto(sub_titulo, "{{SUB_TITULO}}"))
    esp = sub_titulo.find(q("pPr")).find(q("spacing"))
    esp.set(q("before"), "160")  # respiro antes de cada subitem (4.1, 4.2...)
    sub_texto = copy.deepcopy(modelo_texto);  poe_texto(sub_texto, "{{SUB_TEXTO}}")
    E[17].addnext(sub_titulo); sub_titulo.addnext(sub_texto)

    # anexos: um paragrafo por linha
    confere(E[20], "É possível fazer o tema desta"); confere(E[21], "Se o anexo mudar")
    anexo = copy.deepcopy(modelo_texto); poe_texto(anexo, "{{ANEXO}}")
    E[20].addprevious(anexo); body.remove(E[20]); body.remove(E[21])

    # termos e definicoes: fica so o paragrafo do 1o termo (negrito + ': ' + definicao)
    confere(E[24], "(Aqui se colocam os termos"); body.remove(E[24])
    confere(E[25], "Comitê de Compliance")
    runs = E[25].findall(q("r"))
    runs[0].find(q("t")).text = "{{TERMO}}"
    runs[2].find(q("t")).text = "{{DEFINICAO}}"
    for r in runs[3:]:
        E[25].remove(r)
    for i in range(26, 36):
        body.remove(E[i])

    # praticas obrigatorias: regra numerada (negrito/italico) + paragrafo do porque, em bloco repetivel
    confere(E[38], "Aqui se coloca as partes"); confere(E[39], "Item de obrigatoriedade")
    poe_texto(E[38], "{{PRATICA}}"); poe_texto(E[39], "{{PORQUE}}")
    for i in (40, 41, 42, 43, 45):  # bookmarkEnd solto, exemplos "XXXXXXX" e paragrafos vazios; o 44 fica de respiro antes de DUVIDAS
        body.remove(E[i])

    confere(E[47], "Contatar área tal"); poe_texto(E[47], "{{DUVIDAS}}")

    # tabela de revisao: so a coluna REV. 0 recebe dados
    tabela = E[51]
    linhas = tabela.findall(q("tr"))
    ordem = [None, "ELAB_EM", "ELAB_POR", "REV_EM", "REV_POR", "APR_EM", "APR_POR"]
    assert len(linhas) == len(ordem)
    for tr, marc in zip(linhas, ordem):
        if marc:
            poe_texto(tr.findall(q("tc"))[1].find(q("p")), "{{%s}}" % marc)

    # titulo do item 4 (e a linha do sumario): "(NOME DO PROCESSSO)" -> marcador
    n = 0
    for t in doc.iter(q("t")):
        if t.text and "(NOME DO PROCESSSO)" in t.text:
            t.text = t.text.replace("(NOME DO PROCESSSO)", "{{NOME_PROC_MAI}}"); n += 1
    assert n == 2, n  # titulo + sumario

    # ---------- cabecalho ----------
    feitos = set()
    for p in list(cab.iter(q("p"))):
        tx = texto(p)
        if tx == "Nome do Processo":
            poe_texto(p, "{{NOME_PROCESSO}}"); feitos.add("nome")
        elif tx.startswith("Nº: "):
            xx = [r for r in p.findall(q("r")) if texto(r) == "XX"]
            assert len(xx) == 1
            rpr = xx[0].find(q("rPr"))
            if rpr is None:
                rpr = etree.Element(q("rPr")); xx[0].insert(0, rpr)
            for c in rpr.findall(q("color")):
                rpr.remove(c)
            cor = etree.Element(q("color")); cor.set(q("val"), "FF0000")
            pos = 0
            for i, ch in enumerate(rpr):
                if etree.QName(ch).localname in ("rStyle", "rFonts", "b", "bCs", "i", "iCs", "caps", "smallCaps", "strike", "dstrike", "outline", "shadow", "emboss", "imprint", "noProof", "snapToGrid", "vanish", "webHidden"):
                    pos = i + 1
            rpr.insert(pos, cor)
            feitos.add("numero")
        elif tx.startswith("Status: "):
            rs = p.findall(q("r"))
            rs[1].find(q("t")).text = "{{STATUS}}"
            seguinte = p.getnext()
            assert seguinte is not None and texto(seguinte).startswith("Revisão/Vigente")
            seguinte.getparent().remove(seguinte)
            feitos.add("status")
        elif tx.startswith("Próxima revisão: "):
            rs = p.findall(q("r"))
            rs[1].find(q("t")).text = "{{PROXIMA_REVISAO}}"
            p.remove(rs[2])
            feitos.add("proxima")
    assert feitos == {"nome", "numero", "status", "proxima"}, feitos

    # ---------- rodape ----------
    marcas = {"Elaborado por:": "ELAB_POR", "Revisada por:": "REV_POR", "Aprovado por:": "APR_POR", "Data de Aprovação:": "APR_EM"}
    n = 0
    for tc in rod.iter(q("tc")):
        ps = tc.findall(q("p"))
        rotulo = texto(ps[0])
        assert rotulo in marcas, rotulo
        poe_texto(ps[1], "{{%s}}" % marcas[rotulo]); n += 1
    assert n == 4

    # ---------- grava ----------
    novos = {
        "word/document.xml": etree.tostring(doc, xml_declaration=True, encoding="UTF-8", standalone=True),
        "word/header1.xml": etree.tostring(cab, xml_declaration=True, encoding="UTF-8", standalone=True),
        "word/footer1.xml": etree.tostring(rod, xml_declaration=True, encoding="UTF-8", standalone=True),
    }
    with zipfile.ZipFile(destino, "w", zipfile.ZIP_DEFLATED) as zout:
        for info in zin.infolist():
            zout.writestr(info.filename, novos.get(info.filename, zin.read(info.filename)))
    print("ok ->", destino)


if __name__ == "__main__":
    if len(sys.argv) != 3:
        raise SystemExit(__doc__)
    monta(sys.argv[1], sys.argv[2])
