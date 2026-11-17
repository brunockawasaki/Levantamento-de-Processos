# -*- coding: utf-8 -*-
"""Coloca o marcador {{ELABORADOR}} nos modelos de Word (POP, IT e Regimento) no lugar do nome fixo do cabecalho.

Feito em 13/11/2026 a pedido do Bruno: o campo "Elaborador:" do cabecalho vinha escrito "Bruno Kawasaki" em todos os
documentos. Agora o formulario preenche com a resposta de "Nome de quem preenche".
  - cabecalho: "Elaborador: Bruno Kawasaki" -> "Elaborador: {{ELABORADOR}}"  (2 copias: caixa de texto e compatibilidade)
  - tabela Controle de emissao: "Elaboração:" -> "Elaboração: {{ELABORADOR}}"
  - "Revisor: Bruno Kawasaki" e "Validador: ..." (IT e Regimento) NAO mudam (Processos revisa).

Uso: python ajustar_elaborador_modelo.py <modelo.docx> [<modelo.docx> ...]
Pode rodar de novo: se o marcador ja estiver la, nao faz nada.
"""
import re, sys, zipfile, shutil, os

CABECALHO = re.compile(r"(Elaborador: </w:t></w:r><w:sdt>.*?<w:sdtContent>.*?<w:t[^>]*>)Bruno Kawasaki(</w:t>)", re.S)
TABELA = "<w:t>Elaboração:</w:t>"
TABELA_NOVA = '<w:t xml:space="preserve">Elaboração: {{ELABORADOR}}</w:t>'


def ajusta_xml(xml):
    ja = xml.count("{{ELABORADOR}}")
    if ja:
        return xml, 0, 0, ja
    xml, n1 = CABECALHO.subn(r"\1{{ELABORADOR}}\2", xml)
    n2 = xml.count(TABELA)
    if n2:
        xml = xml.replace(TABELA, TABELA_NOVA)
    return xml, n1, n2, 0


def processa(caminho):
    tmp = caminho + ".tmp"
    with zipfile.ZipFile(caminho) as zin, zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            dados = zin.read(item.filename)
            if item.filename == "word/document.xml":
                xml, n1, n2, ja = ajusta_xml(dados.decode("utf-8"))
                if ja:
                    print("  ja tinha o marcador (%d): nada a fazer" % ja)
                else:
                    assert n1 == 2, "esperava 2 'Elaborador:' no cabecalho, achei %d" % n1
                    assert n2 == 1, "esperava 1 'Elaboração:' na tabela, achei %d" % n2
                    print("  cabecalho trocado: %d | tabela: %d" % (n1, n2))
                dados = xml.encode("utf-8")
            zout.writestr(item, dados)
    os.replace(tmp, caminho)


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    for c in sys.argv[1:]:
        print(os.path.basename(c))
        processa(c)
