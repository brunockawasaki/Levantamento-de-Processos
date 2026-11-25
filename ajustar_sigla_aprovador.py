# -*- coding: utf-8 -*-
"""Sigla do setor e aprovador nos Words de POP, IT e Regimento (pedido do Bruno, 25/11/2026).

Problema: os modelos saiam sempre com a sigla da Qualidade (PO.GG.QLD.001; IT e Regimento com PRO) e com o nome do
aprovador escrito fixo (Alessandro Braz / Carlos Vargas). Regra nova:
  - codigo: <TIPO>.GG.<sigla do setor da "Area Dona do Processo">.XXX   (o numero de sequencia e dado por Processos)
  - aprovador: o formulario ganha o campo "Nome do aprovador"; vazio sai "(a definir)"
  - "Revisor: Bruno Kawasaki" e "Validador: ..." (IT e Regimento) NAO mudam.

Marcadores nos modelos: {{SIGLA_SETOR}} (documento e cabecalho) e {{APROVADOR}} (2 lugares no corpo).
(nao usar {{SIGLA}}: esse ja e o marcador da tabela de siglas/glossario.)

Uso:
  python ajustar_sigla_aprovador.py modelo  <modelo.docx> [<modelo.docx> ...]
  python ajustar_sigla_aprovador.py pop     <formulario_POP.html> <modelo_POP.docx>
Pode rodar de novo: se o marcador ja estiver la, nao faz nada.
"""
import re, sys, zipfile, os, base64

# ---------------------------------------------------------------- modelos Word
def _sub(xml, padrao, repl, esperado, nome):
    novo, n = re.subn(padrao, repl, xml, flags=re.S)
    assert n == esperado, "%s: esperava %d, achei %d" % (nome, esperado, n)
    return novo


def ajusta_documento(xml, nome):
    if "{{SIGLA_SETOR}}" in xml:
        return xml, False
    if "QLD" in xml and "Alessandro Braz" in xml:                      # POP: "PO.GG.QLD." + "001"
        xml = _sub(xml, r"(<w:t[^>]*>)PO\.GG\.QLD\.(</w:t>(?:(?!<w:t[ >]).){0,600}?<w:t[^>]*>)001(</w:t>)",
                   r"\1PO.GG.{{SIGLA_SETOR}}.\2XXX\3", 2, "codigo POP")
        xml = _sub(xml, r"(<w:t[^>]*>)Alessandro Braz, Carlos Vargas(</w:t>)", r"\1{{APROVADOR}}\2", 2, "aprovador POP")
    else:                                                                # IT / Regimento: "IT" ".GG." "PRO" ".001"
        xml = _sub(xml, r"(<w:t[^>]*>\.GG\.</w:t>(?:(?!<w:t[ >]).){0,400}?<w:t[^>]*>)PRO(</w:t>(?:(?!<w:t[ >]).){0,400}?<w:t[^>]*>)\.001(</w:t>)",
                   r"\1{{SIGLA_SETOR}}\2.XXX\3", 2, "codigo IT/RE")
        xml = _sub(xml, r"(<w:t[^>]*>)Carlos Vargas(</w:t>)", r"\1{{APROVADOR}}\2", 2, "aprovador IT/RE")
    return xml, True


def ajusta_cabecalho(xml):
    if "{{SIGLA_SETOR}}" in xml:
        return xml, False
    if "PO.GG.QLD.001" in xml:
        return xml.replace("PO.GG.QLD.001", "PO.GG.{{SIGLA_SETOR}}.XXX"), True
    novo, n = re.subn(r"(<w:t[^>]*>\.GG\.</w:t>(?:(?!<w:t[ >]).){0,400}?<w:t[^>]*>)PRO(</w:t>(?:(?!<w:t[ >]).){0,400}?<w:t[^>]*>)\.001(</w:t>)",
                      r"\1{{SIGLA_SETOR}}\2.XXX\3", xml, flags=re.S)
    return novo, n > 0


def ajusta_extras(xml, parte):
    """26/11/2026 (novo PO.GG.PRO.002): sigla da EMPRESA no codigo, Revisao 00, data do dia e status 'Criacao'.
    parte: 'doc' (document.xml), 'cab' (header) ou 'rod' (footer). Marcadores novos: {{SIGLA_EMPRESA}} e {{DATA_DOC}}."""
    if "{{SIGLA_EMPRESA}}" in xml or "{{DATA_DOC}}" in xml:
        return xml, False
    if parte in ("doc", "cab"):
        xml = xml.replace("PO.GG.{{SIGLA_SETOR}}", "PO.{{SIGLA_EMPRESA}}.{{SIGLA_SETOR}}")
        xml = re.sub(r"(<w:t[^>]*>)\.GG\.(</w:t>(?:(?!<w:t[ >]).){0,400}?<w:t[^>]*>\{\{SIGLA_SETOR\}\})",
                     r"\1.{{SIGLA_EMPRESA}}.\2", xml, flags=re.S)
    if parte == "doc":
        assert xml.count("{{SIGLA_EMPRESA}}") == 2, "codigo (empresa) no corpo: %d" % xml.count("{{SIGLA_EMPRESA}}")
        xml = _sub(xml, r"(<w:t[^>]*>Data: </w:t>(?:(?!<w:t[ >]).){0,600}?<w:t[^>]*>)\d{2}/\d{2}/\d{4}(</w:t>)", r"\1{{DATA_DOC}}\2", 2, "data capa")
        xml = _sub(xml, r"(<w:t[^>]*>Revisão: </w:t>(?:(?!<w:t[ >]).){0,3000}?<w:t[^>]*>)\d+(</w:t>)", r"\g<1>00\2", 2, "revisao capa")
        xml = _sub(xml, r"(<w:t[^>]*>Status</w:t>(?:(?!<w:t[ >]).){0,600}?<w:t[^>]*>: </w:t>(?:(?!<w:t[ >]).){0,3000}?<w:t[^>]*>)(?:Aprovado|Em aprovação)(</w:t>)",
                   r"\1Criação\2", 2, "status capa")
        xml = _sub(xml, r"(<w:t[^>]*>)0(</w:t>)", r"\g<1>00\2", 1, "revisao no controle de emissao")
    elif parte == "cab":
        assert "{{SIGLA_EMPRESA}}" in xml, "codigo (empresa) no cabecalho"
    elif parte == "rod":
        xml = _sub(xml, r"(<w:t[^>]*>Revisão:0</w:t>(?:(?!<w:t[ >]).){0,600}?<w:t[^>]*>)\d(</w:t>)", r"\g<1>0\2", 1, "revisao rodape")
        xml = _sub(xml, r"(<w:t[^>]*>Data: </w:t>(?:(?!<w:t[ >]).){0,600}?<w:t[^>]*>)\d{2}/\d{2}/\d{4}(</w:t>)", r"\1{{DATA_DOC}}\2", 1, "data rodape")
    return xml, True


RPR_ARIAL = '<w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/></w:rPr>'


def ajusta_status_e_controle(xml):
    """26/11/2026: (1) Status da capa = 'Em elaboracao' (texto que o cabecalho ja usa);
    (2) POP: controle de emissao ganha as linhas Revisor e Validador (como IT e Regimento; Tabela 6 do PO.GG.PRO.002)."""
    mudou = False
    novo, n = re.subn(r"(<w:t[^>]*>Status</w:t>(?:(?!<w:t[ >]).){0,600}?<w:t[^>]*>: </w:t>(?:(?!<w:t[ >]).){0,3000}?<w:t[^>]*>)(?:Aprovado|Em aprovação|Criação)(</w:t>)",
                      r"\1Em elaboração\2", xml, flags=re.S)
    if n:
        assert n == 2, "status: achei %d" % n
        mudou = novo != xml
        xml = novo
    antigo = '<w:br/><w:t>Aprovação</w:t></w:r>'
    if xml.count(antigo) == 1 and "Revisor: Bruno Kawasaki</w:t>" not in xml and "Elaboração: {{ELABORADOR}}" in xml:   # so o POP tem esse formato
        xml = xml.replace(antigo, ('<w:br/><w:t>Revisor: Bruno Kawasaki</w:t></w:r><w:r>%s<w:br/><w:t>Validador: Carlos Vargas</w:t></w:r>'
                                   '<w:r>%s<w:br/><w:t>Aprovação</w:t></w:r>') % (RPR_ARIAL, RPR_ARIAL))
        antigo_d = '<w:br/><w:t>DD/MM/AA</w:t></w:r><w:r>%s<w:t>AA</w:t></w:r>' % RPR_ARIAL
        assert xml.count(antigo_d) == 1, "datas do POP: achei %d" % xml.count(antigo_d)
        xml = xml.replace(antigo_d, antigo_d + ('<w:r>%s<w:br/><w:t>DD/MM/AAAA</w:t></w:r>' % RPR_ARIAL) * 2)
        mudou = True
    return xml, mudou


def ajusta_rev_cabecalho(xml):
    """Cabecalho: 'Documento: PO.xx.xxx.XXX Rev.01 (30/06/2026)' -> 'Rev.00 ({{DATA_DOC}})'. O texto vem partido em varios runs."""
    if "{{DATA_DOC}}" in xml:
        return xml, False
    i = xml.find(">Rev.</w:t>")
    if i < 0:
        return xml, False
    toks = list(re.finditer(r"(<w:t[^>]*>)([^<]*)(</w:t>)", xml[i + len(">Rev.</w:t>"):]))
    base = i + len(">Rev.</w:t>")
    acum, usados = "", []
    for m in toks:
        acum += m.group(2)
        usados.append(m)
        if ")" in acum:
            break
    assert ")" in acum, "cabecalho: nao achei o ')' da revisao"
    resto = acum.split(")", 1)[1]           # texto depois do ')' (normalmente vazio)
    novo = []
    for k, m in enumerate(usados):
        texto = ("00 ({{DATA_DOC}})" + resto) if k == 0 else ""
        novo.append((base + m.start(2), base + m.end(2), texto))
    for ini, fim, texto in reversed(novo):
        xml = xml[:ini] + texto + xml[fim:]
    return xml, True


def processa_modelo(caminho):
    tmp = caminho + ".tmp"
    feito = {"doc": 0, "cab": 0, "extras": 0}
    with zipfile.ZipFile(caminho) as zin, zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            dados = zin.read(item.filename)
            if item.filename == "word/document.xml":
                xml, mudou = ajusta_documento(dados.decode("utf-8"), caminho)
                feito["doc"] += mudou
                xml, mudou = ajusta_extras(xml, "doc")
                feito["extras"] += mudou
                xml, mudou = ajusta_status_e_controle(xml)
                feito["extras"] += mudou
                dados = xml.encode("utf-8")
            elif re.match(r"word/header\d+\.xml$", item.filename):
                xml, mudou = ajusta_cabecalho(dados.decode("utf-8"))
                feito["cab"] += mudou
                if "{{SIGLA_SETOR}}" in xml:
                    xml, mudou = ajusta_extras(xml, "cab")
                    feito["extras"] += mudou
                    xml, mudou = ajusta_rev_cabecalho(xml)
                    feito["extras"] += mudou
                dados = xml.encode("utf-8")
            elif re.match(r"word/footer\d+\.xml$", item.filename):
                xml = dados.decode("utf-8")
                if "Revisão:0" in xml:
                    xml, mudou = ajusta_extras(xml, "rod")
                    feito["extras"] += mudou
                dados = xml.encode("utf-8")
            zout.writestr(item, dados)
    os.replace(tmp, caminho)
    return feito


processa = processa_modelo   # nome usado pelos montar_modelo_*.py

# ---------------------------------------------------------------- formulario de POP (html)
SIGLAS_JS = r'''  // ---------- Sigla do setor (codigo do documento) ----------
  // A sigla vem da "Area Dona do Processo". Setor sem sigla conhecida sai "XXX" (Processos define). Digitar a propria sigla (3 letras) tambem vale.
  var SIGLAS_SETOR = [["gestao de projetos de incorporacao","GPI"],["novos negocios","NNG"],["incorporacao","INC"],["suprimentos","SUP"],["gmo","GMO"],["credito imobiliario","CIM"],["projetos complementares","PJC"],["orcamento","ORC"],["planejamento","PLA"],["qualidade","QLD"],["contabilidade","CON"],["fiscal","FIS"],["juridico","JUR"],["financeiro","FIN"],["processos","PRO"],["seguranca do trabalho","SST"],["controladoria","CNT"],["recursos humanos","RCH"],["administrativo","ADM"],["arquitetura","ARQ"],["comercial make it","CMI"],["comercial gadens","CGD"],["comercial","COM"],["contas a pagar","CAP"],["contas a receber","CAR"],["tesouraria","TES"],["marketing","MKT"],["infraestrutura","INF"],["projetos e sistemas","PES"],["engenharia de obras","OBR"]];
  function siglaDoSetor(texto){
    var t = String(texto || "").replace(/\s+/g, " ").trim();
    if(/^[A-Za-z]{3}$/.test(t)) return t.toUpperCase();
    var n = t.normalize("NFD").replace(/[̀-ͯ]/g, "").toLowerCase();
    for(var i = 0; i < SIGLAS_SETOR.length; i++){ if(n.indexOf(SIGLAS_SETOR[i][0]) >= 0) return SIGLAS_SETOR[i][1]; }
    return "XXX";
  }
  function marcadoresCabecalho(xml){
    xml = xml.split("{{SIGLA_SETOR}}").join(siglaDoSetor((document.getElementById("id_setor") || {}).value));
    var aprov = String((document.getElementById("id_aprovador") || {}).value || "").replace(/\s+/g, " ").trim();
    return xml.split("{{APROVADOR}}").join(escapeHtml(aprov || "(a definir)"));
  }
  function marcadoresPartes(zip, enc, dec){
    Object.keys(zip).forEach(function(k){
      if(/^word\/(header|footer)\d+\.xml$/.test(k)){ zip[k] = enc.encode(marcadoresCabecalho(dec.decode(zip[k]))); }
    });
  }
  (function(){
    var s = document.getElementById("id_setor"), d = document.getElementById("id_setor_sigla");
    if(!s || !d) return;
    function mostra(){
      var v = siglaDoSetor(s.value);
      d.textContent = s.value.trim() === "" ? "" : (v === "XXX" ? "Sigla no código do documento: XXX (setor sem sigla cadastrada; Processos define)." : "Sigla no código do documento: " + v + (v === "COM" ? " (Comercial de forma geral)." : "."));
    }
    s.addEventListener("input", mostra); mostra();
  })();

'''

CAMPO_HTML = '''      <div class="id-field">
        <label for="id_aprovador">Nome do aprovador
          <span class="help-icon" tabindex="0">?<span class="tooltip-text">Quem vai aprovar o documento no seu setor. Se ainda não souber, deixe em branco: sai "(a definir)" no Word.</span></span>
        </label>
        <input type="text" id="id_aprovador" placeholder="Ex.: Maria Souza" spellcheck="true">
      </div>
'''


def troca(s, antigo, novo, nome, vezes=1):
    assert s.count(antigo) == vezes, "%s: esperava %d, achei %d" % (nome, vezes, s.count(antigo))
    return s.replace(antigo, novo)


EMPRESA_JS = r'''  // ---------- Sigla da empresa (codigo do documento, Tabela 4 do PO.GG.PRO.002) e data do dia ----------
  var SIGLAS_EMPRESA = {"Grupo Gadens":"GG","Gadens Incorporadora":"IG","Construtora Make It":"CM","Autoshopping Linha Verde":"AS","Triunfo do Brasil":"TB","Shopping Cidade":"SC","ULTJ":"UL"};
  function siglaDaEmpresa(nome){ return SIGLAS_EMPRESA[String(nome || "").trim()] || "GG"; }
  function dataHoje(){
    var d = new Date(), p = function(n){ return (n < 10 ? "0" : "") + n; };
    return p(d.getDate()) + "/" + p(d.getMonth() + 1) + "/" + d.getFullYear();
  }
'''
LINHA_SETOR = '    xml = xml.split("{{SIGLA_SETOR}}").join(siglaDoSetor((document.getElementById("id_setor") || {}).value));\n'
LINHAS_NOVAS = ('    xml = xml.split("{{SIGLA_EMPRESA}}").join(siglaDaEmpresa((document.getElementById("id_empresa") || {}).value));\n'
                '    xml = xml.split("{{DATA_DOC}}").join(dataHoje());\n')


def patch_js_empresa(s):
    """Acrescenta no formulario as funcoes e as 2 trocas de marcador. Idempotente. Devolve (texto, mudou)."""
    if "siglaDaEmpresa" in s:
        return s, False
    s = troca(s, "  function marcadoresCabecalho(xml){\n", EMPRESA_JS + "  function marcadoresCabecalho(xml){\n", "funcoes empresa/data")
    s = troca(s, LINHA_SETOR, LINHA_SETOR + LINHAS_NOVAS, "trocas empresa/data")
    return s, True


def processa_pop(caminho_html, caminho_docx):
    s = open(caminho_html, encoding="utf-8", newline="").read()
    crlf = "\r\n" in s
    s = s.replace("\r\n", "\n")
    s, _mudou = patch_js_empresa(s)
    if _mudou:
        print("  JS: sigla da empresa e data do dia adicionados")
    if "id_aprovador" in s:
        print("  formulario ja tem o campo do aprovador: nada a fazer no HTML")
    else:
        # campo novo logo depois do select de empresa
        s = troca(s, '          <option value="Grupo Gadens">Grupo Gadens</option>\n        </select>\n      </div>\n',
                  '          <option value="Grupo Gadens">Grupo Gadens</option>\n        </select>\n      </div>\n' + CAMPO_HTML, "campo")
        s = troca(s, '<input type="text" id="id_setor" placeholder="Ex.: Contas a receber" spellcheck="true">',
                  '<input type="text" id="id_setor" placeholder="Ex.: Contas a receber" spellcheck="true">\n        <small id="id_setor_sigla" style="display:block;margin-top:4px;color:#767171;font-size:12px;"></small>', "dica da sigla")
        s = troca(s, 'var idFieldsIds = ["id_nome","id_email","id_setor","id_empresa","id_data","id_tarefa"];',
                  'var idFieldsIds = ["id_nome","id_email","id_setor","id_empresa","id_aprovador","id_data","id_tarefa"];', "ids")
        s = troca(s, '        empresa: document.getElementById("id_empresa").value,\n',
                  '        empresa: document.getElementById("id_empresa").value,\n        aprovador: document.getElementById("id_aprovador").value,\n', "salvar")
        s = troca(s, '      document.getElementById("id_empresa").value = state.identificacao.empresa || "";\n',
                  '      document.getElementById("id_empresa").value = state.identificacao.empresa || "";\n      document.getElementById("id_aprovador").value = state.identificacao.aprovador || "";\n', "restaurar")
        s = troca(s, '    lines.push("Empresa do Grupo: " + (document.getElementById("id_empresa").value || "-"));\n',
                  '    lines.push("Empresa do Grupo: " + (document.getElementById("id_empresa").value || "-"));\n    lines.push("Nome do aprovador: " + (document.getElementById("id_aprovador").value || "-"));\n', "resumo")
        s = troca(s, '  // ---------- Identificação: listeners ----------\n', SIGLAS_JS + '  // ---------- Identificação: listeners ----------\n', "helpers")
        s = troca(s, '    xml = xml.split("{{ELABORADOR}}").join(escapeHtml(nomeElab || "(não informado)"));\n',
                  '    xml = xml.split("{{ELABORADOR}}").join(escapeHtml(nomeElab || "(não informado)"));\n    xml = marcadoresCabecalho(xml);\n', "corpo", vezes=1)
        s = troca(s, '      zip["word/document.xml"] = new TextEncoder().encode(filledXml);\n',
                  '      zip["word/document.xml"] = new TextEncoder().encode(filledXml);\n      marcadoresPartes(zip, new TextEncoder(), new TextDecoder("utf-8"));\n', "cabecalho")
        print("  HTML ajustado")
    # modelo embutido
    b64 = base64.b64encode(open(caminho_docx, "rb").read()).decode("ascii")
    s, n = re.subn(r'(var TEMPLATE_DOCX_B64 = ")[A-Za-z0-9+/=]+(";)', lambda m: m.group(1) + b64 + m.group(2), s)
    assert n == 1, "TEMPLATE_DOCX_B64: achei %d" % n
    print("  modelo Word embutido atualizado")
    if crlf:
        s = s.replace("\n", "\r\n")
    open(caminho_html, "w", encoding="utf-8", newline="").write(s)


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    modo = sys.argv[1]
    if modo == "modelo":
        for c in sys.argv[2:]:
            print(os.path.basename(c), processa_modelo(c))
    elif modo == "pop":
        processa_pop(sys.argv[2], sys.argv[3])
    else:
        sys.exit(__doc__)
