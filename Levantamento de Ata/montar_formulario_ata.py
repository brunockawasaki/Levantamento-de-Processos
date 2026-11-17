# -*- coding: utf-8 -*-
"""Monta o 'Formulario de Ata de Reuniao' a partir do formulario de POP.

Uso:
  python montar_formulario_ata.py <formulario_POP.html> <MODELO_ATA_com_marcadores.docx> <saida.html>

- <formulario_POP.html>: formulario interativo de POP (fonte do layout, voz, progresso, tema).
- <MODELO_ATA_com_marcadores.docx>: gerado por montar_modelo_ata.py (embutido no HTML em base64).
As perguntas seguem os campos do modelo de Ata (Departamento, Assunto/Objetivo, Local/Data, Responsavel,
Participantes, Pauta, Ata, Outputs/Pendencias, Documentos anexos, Proxima reuniao). Data da expedicao e Registro feito por
sao preenchidos pelo motor. Se o formulario de POP mudar, basta
rodar de novo (o script avisa se algum trecho nao for encontrado).
"""
import sys, re, json, base64


def rep(s, old, new):
    n = s.count(old)
    if n != 1:
        raise SystemExit("Trecho aparece %d vezes (esperava 1) no formulario de POP:\n%s" % (n, old[:220]))
    return s.replace(old, new)


def rep_re(s, pattern, new, flags=re.S):
    out, n = re.subn(pattern, lambda m: new, s, count=1, flags=flags)
    if n != 1:
        raise SystemExit("NAO ACHEI (regex) no formulario de POP:\n" + pattern[:220])
    return out


# ------------------------------------------------------------------ exemplos (ata publicada)
F = "Ata de Reunião · Almoxarifado Central (Contabilidade/Fiscal)"
EXEMPLOS = {
    "q3": [
        {"fonte": F, "texto":
            "Andressa Gazolla\nEdgard Machado\nCarlos Vargas\nCleber Colin\nFelipe Schimit\nPaulo Tavares"},
    ],
    "q4": [
        {"fonte": F, "texto":
            "A pauta trouxe recomendações do time de Contabilidade/Fiscal, para evitar a não cobertura de materiais no sentido fiscal e contábil, utilizando a luz da reforma tributária.\n"
            "O almoxarifado central seria o comprador de materiais, devendo a partir disso, distribuir para as obras.\n"
            "Outra opção, é o almoxarifado central ser um prestador de serviços, que realizariam a armazenagem dos materiais.\n"
            "Estrutura física deverá ser equivalente a suprir a necessidade de todo o Grupo."},
    ],
    "q5": [
        {"fonte": F, "texto":
            "O material contábil será compartilhado pelo time de Contabilidade/Fiscal com os demais participantes para análise, bem como, será necessário nova rodada para definição do tema."},
    ],
    "q6": [
        {"fonte": F, "linhas": [
            ["Agendar conversa com a Thá para entender como o processo era feito e avaliar se contribuiria para o Grupo de alguma forma", "Andressa", ""],
            ["Verificar se a conversa entre Jurídico e Thá aconteceu, para definir data", "Bruno", ""],
        ]},
    ],
}
EXEMPLOS_NOTAS = {
    "q3": "Na ata publicada os nomes aparecem em uma linha, separados por ponto e vírgula. Aqui, escreva um nome por linha: o formulário coloca dois nomes por linha da tabela no Word.",
    "q5": "A ata é o resumo do que foi discutido e decidido. Na ata de exemplo, só a decisão foi registrada.",
    "q6": "A ata publicada não traz a data de entrega. Aqui ela é um campo próprio, para cada pendência ter prazo. Se não houver prazo definido, deixe em branco.",
}

# ------------------------------------------------------------------ perguntas
QUESTIONS_JS = r'''  var QUESTIONS = [
    {id:"q1", num:1, title:"Onde foi a reunião?",
      dica:'Ex.: "Sala de reuniões da sede" ou "Microsoft Teams". A data da reunião você informou lá em cima, na identificação.',
      label:"Ata · Local / Data", type:"text"},
    {id:"q2", num:2, title:"Quem foi o responsável pela reunião?",
      dica:'Quem conduziu a reunião. Ex.: "Carlos Vargas".',
      label:"Ata · Responsável", type:"text"},
    {id:"q3", num:3, title:"Quem participou da reunião?",
      dica:'Escreva um nome por linha. O formulário coloca os nomes em duas colunas, como no modelo.',
      label:"Ata · Participantes", type:"text", autoGrow:true},
    {id:"q4", num:4, title:"Quais assuntos foram tratados na reunião?",
      dica:'Escreva um assunto por linha, na ordem em que foram discutidos. Cada linha vira um parágrafo na ata.',
      label:"Ata · Pauta da Reunião", type:"text", required:true, autoGrow:true},
    {id:"q5", num:5, title:"O que foi discutido e decidido? (ata)",
      dica:'Resuma o que foi conversado e o que ficou decidido, de forma objetiva. Se nada foi decidido, escreva "não houve decisão".',
      label:"Ata da Reunião", type:"text", required:true, autoGrow:true},
    {id:"q6", num:6, title:"O que ficou pendente? Qual ação (output) alguém precisa entregar, quem entrega e até quando?",
      dica:'Uma linha por pendência: a ação ou output, o responsável e a data de entrega. Ex.: "Enviar o material de análise" — "Maria" — "30/10/2026". A coluna OK fica em branco no Word, para marcar quando for entregue. Se não houver, deixe a tabela vazia.',
      label:"Ata · Pendências", type:"table", cellType:"textarea",
      columns:["Pendência (Ação)","Responsável","Data"],
      example:["Enviar o material para análise","Nome do responsável","DD/MM/AAAA"]},
    {id:"q7", num:7, title:"Algum documento foi anexado à ata?",
      dica:'Um documento por linha. Ex.: "Apresentação da reunião.pdf". Se não houver, escreva "não há".',
      label:"Ata · Documentos Anexos", type:"text", autoGrow:true},
    {id:"q8", num:8, title:"Quando será a próxima reunião?",
      dica:'Ex.: "20/11/2026" ou "a definir". Se não houver próxima reunião, escreva "não há".',
      label:"Ata · Data da próxima reunião", type:"text"},
    {id:"q9", num:9, title:"Tem mais alguma coisa que quem for revisar a ata precisa saber?",
      dica:"Pendências sem responsável, pontos que ficaram em aberto, ideias para a próxima reunião. Isso não vai para o Word.",
      label:null, type:"text"}
  ];
'''

# ------------------------------------------------------------------ motor do Word
ENGINE_JS = r'''function oneLine(s){ return String(s || "").replace(/\s*[\r\n]+\s*/g, " ").trim(); }

  function formatDateBR(iso){
    var m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(iso || "");
    return m ? (m[3] + "/" + m[2] + "/" + m[1]) : "";
  }

  function buildParticipantesRows(){
    var nomes = getTextLines("q3").map(function(s){ return s.replace(/[\s.;,]+$/, ""); })
                                  .filter(function(s){ return s !== ""; });
    var rows = [];
    for(var i = 0; i < nomes.length; i += 2) rows.push([nomes[i], nomes[i + 1] || ""]);
    return rows;
  }

  function buildLocalData(){
    var partes = [oneLine(document.getElementById("q1").value), formatDateBR(document.getElementById("id_data").value)];
    return partes.filter(function(p){ return p !== ""; }).join(" – ");
  }

  function buildFilledDocumentXml(xml){
    var NI = "(não informado)";
    xml = fillToken(xml, "DEPARTAMENTO", oneLine(document.getElementById("id_setor").value), NI);
    xml = fillToken(xml, "ASSUNTO", oneLine(document.getElementById("id_tarefa").value), NI);
    xml = fillToken(xml, "LOCAL_DATA", buildLocalData(), NI);
    xml = fillToken(xml, "RESPONSAVEL", oneLine(document.getElementById("q2").value), NI);
    var parts = buildParticipantesRows();
    if(parts.length === 0) parts = [[NI, ""]];
    xml = substituteRepeatingRow(xml, ["PART1", "PART2"], parts);
    xml = fillToken(xml, "PAUTA", document.getElementById("q4").value, NI);
    xml = fillToken(xml, "SUMULA", document.getElementById("q5").value, NI);
    var rows = getTableRowsFiltered("q6").map(function(r){ return [r[0], r[1], r[2], ""]; });
    if(rows.length === 0) rows = [[NI, "", "", ""]];
    xml = substituteRepeatingRow(xml, ["ACAO", "ACAO_RESP", "ACAO_DATA", "ACAO_OK"], rows);
    xml = fillToken(xml, "EXPEDICAO", formatDateBR(todayISO()), NI);
    xml = fillToken(xml, "REGISTRO", oneLine(document.getElementById("id_nome").value), NI);
    xml = fillToken(xml, "PROXIMA", oneLine(document.getElementById("q8").value), NI);
    xml = fillToken(xml, "ANEXOS", document.getElementById("q7").value, NI);
    return xml;
  }
'''

# cada linha do texto vira um paragrafo (evita o espacamento esticado do texto justificado)
WT_JS = r'''  function fillToken(xml, token, value, emptyText){
    var lit = "{{" + token + "}}";
    var found = findEnclosingBlock(xml, lit, /<w:p(?=[ >])/, "</w:p>");
    if(!found) return xml;
    var lines = String(value || "").split(/\r?\n/).map(function(l){ return l.trim(); }).filter(function(l){ return l !== ""; });
    if(lines.length === 0) lines = [emptyText || ""];
    var tokenTag = new RegExp('<w:t[^>]*>' + escapeRegex(lit) + '</w:t>');
    var out = lines.map(function(l){
      return found.block.replace(tokenTag, function(){ return '<w:t xml:space="preserve">' + escapeHtml(l) + '</w:t>'; });
    }).join("");
    return xml.slice(0, found.start) + out + xml.slice(found.end);
  }

'''


def main(pop_html, modelo_docx, out_html):
    s = open(pop_html, encoding="utf-8").read()

    m = re.search(r'(var TEMPLATE_DOCX_B64\s*=\s*")([^"]*)(")', s)
    if not m:
        raise SystemExit("Nao achei TEMPLATE_DOCX_B64")
    s = s[:m.start(2)] + "<<B64>>" + s[m.end(2):]

    # --- identidade
    s = rep_re(s, r"<title>.*?</title>", "<title>Ata de Reunião | Grupo Gadens</title>")
    s = rep(s, "<h1>G-Docs Formulário de Levantamento de Processo</h1>", "<h1>G-Docs Formulário de Ata de Reunião</h1>")
    s = rep_re(s, r"<p>Conte como o seu trabalho é feito.*?</p>",
               "<p>Conte o que aconteceu na reunião — o formulário monta a ata no modelo do Grupo Gadens.</p>")
    s = rep(s, 'var STORAGE_KEY = "gadens_levantamento_pop_v1";', 'var STORAGE_KEY = "gadens_ata_v2";')
    s = rep(s, 'var THEME_KEY = "gadens_pop_theme";', 'var THEME_KEY = "gadens_ata_theme";')

    s = rep_re(s, r'<div class="card instrucoes">.*?</ul>\s*</div>', '''<div class="card instrucoes">
    <h2>Antes de começar</h2>
    <p style="font-size:.92rem; margin:0 0 8px;">Este formulário ajuda você a escrever a Ata de Reunião no modelo do Grupo Gadens: departamento, assunto, local e data, responsável, participantes, pauta, ata, pendências com responsável e data, documentos anexos e próxima reunião. Você responde com as suas palavras; o formulário monta o Word.</p>
    <ul>
      <li>Escreva em frases curtas, uma ideia por linha. Cada linha vira um parágrafo na ata.</li>
      <li>O botão "?" de cada pergunta mostra um exemplo de ata já feita.</li>
      <li>Em "Data da reunião", informe o dia em que a reunião aconteceu. A data da expedição da ata é a de hoje, colocada automaticamente.</li>
      <li>Nas pendências, diga a ação (output) que deve ser entregue, quem entrega e até quando.</li>
      <li>Nas tabelas, a linha cinza em itálico é só um exemplo — ela não é salva.</li>
      <li>Suas respostas ficam salvas neste computador enquanto você preenche. Gere o Word antes de fechar a página.</li>
    </ul>
  </div>''')
    s = rep(s, "Área Dona do Processo\n          <span class=\"help-icon\" tabindex=\"0\">?<span class=\"tooltip-text\">Área principal e executora majoritária do processo</span></span>",
               "Departamento\n          <span class=\"help-icon\" tabindex=\"0\">?<span class=\"tooltip-text\">Departamento ou setor responsável pela reunião</span></span>")
    s = rep(s, 'placeholder="Ex.: Contas a receber"', 'placeholder="Ex.: Suprimentos"')
    s = rep(s, 'Nome da tarefa ou rotina descrita\n          <span class="help-icon" tabindex="0">?<span class="tooltip-text">Título do documento</span></span>',
               'Assunto / Objetivo da reunião\n          <span class="help-icon" tabindex="0">?<span class="tooltip-text">Também é usado no nome do arquivo</span></span>')
    s = rep(s, 'placeholder="Use o nome pelo qual ela é conhecida no dia a dia"', 'placeholder="Ex.: Almoxarifado Central"')
    s = rep(s, 'Data de preenchimento <span class="req-mark">*</span>', 'Data da reunião <span class="req-mark">*</span>')
    s = rep_re(s, r'<div class="card terminou">.*?</div>', '''<div class="card terminou">
    <h2>Terminou?</h2>
    <p>Releia rapidinho suas respostas e clique em "Enviar Respostas (gerar Word)" — isso monta a ata já no modelo. Abra o Word e dê uma conferida antes de compartilhar.</p>
    <p>O nome de quem preenche vai em "Registro feito por".</p>
  </div>''')
    s = rep(s, "Grupo Gadens — Levantamento de Processo (base para POP)", "Grupo Gadens — Ata de Reunião")

    # --- perguntas
    s = rep_re(s, r"  var QUESTIONS = \[.*?\n  \];\n", QUESTIONS_JS)

    # --- obrigatorias
    s = rep(s, '{label:"Área Dona do Processo", el:', '{label:"Departamento", el:')
    s = rep(s, '{label:"Nome da tarefa ou rotina", el:', '{label:"Assunto / Objetivo da reunião", el:')
    s = rep(s, '{label:"Data de preenchimento", el:', '{label:"Data da reunião", el:')
    s = rep(s, 'var q1 = document.getElementById("q1");\n    if(q1.value.trim() === "") missing.push({label:"Pergunta 1 — Objetivo do trabalho", el: q1});',
               'var q4 = document.getElementById("q4");\n    if(q4.value.trim() === "") missing.push({label:"Pergunta 4 — Pauta da reunião", el: q4});')
    s = rep(s, 'if(!isQuestionFilled(qById["q9"])) missing.push({label:"Pergunta 9 — Passo a passo (pelo menos 1 linha preenchida)", el: document.getElementById("wrap_q9")});',
               'if(document.getElementById("q5").value.trim() === "") missing.push({label:"Pergunta 5 — Ata da reunião", el: document.getElementById("q5")});')

    # --- exportacao .txt
    s = rep(s, '"Sem nome da tarefa"', '"Sem assunto da reunião"')
    s = rep(s, 'lines.push("LEVANTAMENTO DE PROCESSO — BASE PARA POP");', 'lines.push("ATA DE REUNIÃO");')
    s = rep(s, 'lines.push("Área Dona do Processo: "', 'lines.push("Departamento: "')
    s = rep(s, 'lines.push("Nome da tarefa ou rotina descrita: "', 'lines.push("Assunto / Objetivo da reunião: "')
    s = rep(s, 'lines.push("Data de preenchimento: "', 'lines.push("Data da reunião: "')

    # --- motor do Word
    s = rep_re(s, r"function buildAplicacaoText\(\)\{.*?\n  \}\n", "@@ENGINE@@\n")
    s = rep_re(s, r"  function buildFilledDocumentXml\(xml\)\{.*?\n    return xml;\n  \}\n", "")
    s = s.replace("@@ENGINE@@\n", ENGINE_JS)
    # quebras de linha dentro das celulas / paragrafos
    s = rep(s, "  function substituteSimpleToken(xml, token, value){", WT_JS + "  function substituteSimpleToken(xml, token, value){")
    s = rep(s, """          var re = new RegExp('<w:t[^>]*>' + escapeRegex("{{" + tok + "}}") + '</w:t>');
          var val = (rowValues[i] || "").trim();
          block = block.replace(re, '<w:t xml:space="preserve">' + escapeHtml(val) + '</w:t>');""",
               """          block = fillToken(block, tok, rowValues[i], "");""")
    # rodape (Registro feito por e Data da expedicao)
    s = rep(s, '      zip["word/document.xml"] = new TextEncoder().encode(filledXml);',
               '''      zip["word/document.xml"] = new TextEncoder().encode(filledXml);
''')
    s = rep(s, 'var filename = "POP - " + tarefa', 'var filename = "Ata - " + tarefa')
    s = rep(s, "Abra e revise antes de enviar à equipe de Processos — código do documento, fluxo Bizagi, aprovações e controle de emissão continuam sendo preenchidos por eles.",
               "Abra e revise antes de compartilhar a ata.")

    # --- modo por voz
    s = rep(s, 'title:"Qual é a área dona do processo?", dica:"Exemplo: Contas a receber."',
               'title:"Qual é o departamento responsável pela reunião?", dica:"Exemplo: Suprimentos."')
    s = rep(s, 'title:"Para qual empresa do grupo é esse processo?"', 'title:"Para qual empresa do grupo é essa reunião?"')
    s = rep(s, 'title:"Qual é o nome da tarefa ou rotina que você vai descrever?", dica:"Use o nome pelo qual ela é conhecida no dia a dia."',
               'title:"Qual foi o assunto ou objetivo da reunião?", dica:"Esse assunto também será usado no nome do arquivo."')

    # --- exemplos
    s = rep_re(s, r"  var EXEMPLOS = \{.*?\n  var EXEMPLOS_NOTAS = \{.*?\};\n",
               "  var EXEMPLOS = " + json.dumps(EXEMPLOS, ensure_ascii=False) + ";\n  var EXEMPLOS_NOTAS = " + json.dumps(EXEMPLOS_NOTAS, ensure_ascii=False) + ";\n")
    s = rep(s, "// Cada exemplo é um trecho de POP já aprovado", "// Cada exemplo é um trecho de ata já feita")
    s = rep(s, '(lista.length > 1 ? "Exemplos de POPs já publicados" : "Exemplo de POP já publicado")',
               '(lista.length > 1 ? "Exemplos de atas já feitas" : "Exemplo de ata já feita")')

    b64 = base64.b64encode(open(modelo_docx, "rb").read()).decode("ascii")
    s = s.replace("<<B64>>", b64)
    open(out_html, "w", encoding="utf-8").write(s)
    print("ok ->", out_html, len(s), "chars")


if __name__ == "__main__":
    if len(sys.argv) != 4:
        raise SystemExit(__doc__)
    main(*sys.argv[1:])
