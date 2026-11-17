# -*- coding: utf-8 -*-
"""Monta o 'Formulario de Politica de Departamento' a partir do formulario de POP.

Uso:
  python montar_formulario_politica.py <formulario_POP.html> <MODELO_POLITICA_com_marcadores.docx> <saida.html>

- <formulario_POP.html>: formulario interativo de POP (fonte do layout, voz, progresso, tema).
- <MODELO_POLITICA_com_marcadores.docx>: gerado por montar_modelo_politica.py (embutido no HTML em base64).
As perguntas seguem as secoes do template GG00 (Objetivo, Aplicacao, Responsavel, Processos, Anexos, Termos, Praticas
obrigatorias, Duvidas, Aprovacao e revisao). Alem do corpo, o motor tambem preenche cabecalho, rodape e sumario do Word.
O numero da politica (GGXX) nao e perguntado: fica fixo, com o XX em vermelho (o Compliance define a numeracao).
Se o formulario de POP mudar, basta rodar de novo (o script avisa se algum trecho nao for encontrado).
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


# ------------------------------------------------------------------ exemplos (politica publicada, palavra por palavra)
F = "Política de Departamento · GG01 – Política Corporativa de Gestão de Pagamentos"
SEM_ANEXOS = "Esta Política não possui anexos vinculados no momento de sua publicação."
EXEMPLOS = {
    "q1": [{"fonte": F, "texto":
        "Esta Política estabelece as diretrizes corporativas para a gestão dos pagamentos do Grupo Gadens, definindo os princípios, responsabilidades, controles e regras aplicáveis ao processamento das obrigações financeiras da Companhia.\n"
        "Seu propósito é assegurar que todos os pagamentos sejam realizados de forma planejada, íntegra, rastreável e em conformidade com a legislação vigente, os contratos firmados, as políticas internas e a estratégia financeira do Grupo."}],
    "q2": [{"fonte": F, "texto":
        "Esta Política aplica-se a todas as empresas integrantes do Grupo Gadens, incluindo holdings, controladas, SPEs, SCPs, empreendimentos, corporativo/CSC e demais unidades de negócio.\n"
        "Suas disposições são obrigatórias para todos os colaboradores, administradores, gestores e terceiros envolvidos, direta ou indiretamente, nas atividades relacionadas à contratação, recebimento, validação, aprovação, registro, programação, processamento e controle dos pagamentos."}],
    "q3": [{"fonte": F, "texto":
        "O processo de Gestão de Pagamentos é de responsabilidade da área de Contas a Pagar, subordinada à Diretoria Financeira do Grupo Gadens, cabendo a essa área a condução, o acompanhamento e a manutenção desta Política, em conjunto com as demais áreas envolvidas, conforme suas atribuições descritas no item 6 (Papéis e Responsabilidades)."}],
    "q6": [{"fonte": F, "texto": SEM_ANEXOS}],
    "q7": [{"fonte": F, "linhas": [
        ["Área Demandante", "área responsável pela contratação do bem ou serviço, validação da entrega e solicitação do pagamento."],
        ["Central de Notas", "área responsável pelo recebimento, conferência inicial, validação fiscal e registro dos documentos fiscais no sistema corporativo."],
        ["Competência", "período contábil ao qual pertence a despesa, independentemente da data de seu pagamento."],
    ]}],
    "q8": [{"fonte": F, "linhas": [
        ["Toda aprovação de pagamento deve observar a Matriz Corporativa de Alçadas, sem exceção", ""],
        ["Nenhum pagamento pode ser processado sem documentação fiscal e comprobatória completa", ""],
        ["A segregação de funções entre contratação, aprovação, execução e conciliação deve ser sempre observada", ""],
    ]}],
    "q9": [{"fonte": F, "texto":
        "Dúvidas relacionadas a esta Política devem ser encaminhadas à área de Contas a Pagar, por meio dos canais oficiais de comunicação da Companhia."}],
}
EXEMPLOS_NOTAS = {
    "q6": "A política publicada não tem anexos. Se a sua também não tiver, deixe a tabela vazia: o Word sai com esta mesma frase.",
    "q7": "Na política publicada os termos aparecem na seção \"Definições\". Aqui, escreva um termo por linha da tabela.",
    "q8": "Na política publicada as práticas aparecem em lista, sem o motivo. O modelo da política pede a prática e o motivo (o que ela previne): preencha os dois.",
}

# ------------------------------------------------------------------ perguntas
QUESTIONS_JS = r'''  var QUESTIONS = [
    {id:"q1", num:1, title:"Qual é o objetivo desta política?",
      dica:'Diga qual processo a política trata e qual é o fim pretendido. Escreva em frases curtas: cada linha vira um parágrafo.',
      label:"Política · Objetivo", type:"text", required:true, autoGrow:true},
    {id:"q2", num:2, title:"A quem e a quais processos esta política se aplica?",
      dica:'Diga quais empresas, áreas ou empreendimentos ela alcança e quem precisa cumpri-la. Cada linha vira um parágrafo.',
      label:"Política · Aplicação", type:"text", required:true, autoGrow:true},
    {id:"q3", num:3, title:"Qual área é responsável pelo processo?",
      dica:'Diga qual área ou departamento conduz o processo e mantém esta política.',
      label:"Política · Responsável pelo processo", type:"text", required:true, autoGrow:true},
    {id:"q4", num:4, title:"Como é o processo, de forma geral?",
      dica:'Descreva o processo em visão geral (macro), antes de entrar nos detalhes. Os detalhes vão na próxima pergunta. Cada linha vira um parágrafo.',
      label:"Política · Processo (visão geral)", type:"text", required:true, autoGrow:true},
    {id:"q5", num:5, title:"Quer detalhar o processo em partes? (itens 4.1, 4.2…)",
      dica:'Uma linha por parte: o título e o texto. O Word numera sozinho como 4.1, 4.2 e assim por diante. Se não precisar dividir, deixe a tabela vazia.',
      label:"Política · Partes do processo", type:"table", cellType:"textarea",
      columns:["Título da parte","Texto da parte"],
      example:["Nome da parte","Como funciona esta parte do processo"]},
    {id:"q6", num:6, title:"A política tem anexos?",
      dica:'Um anexo por linha: o nome e onde ele fica (link). Se o anexo mudar, a política precisa ser atualizada. Se não houver anexos, deixe a tabela vazia.',
      label:"Política · Anexos", type:"table", cellType:"textarea",
      columns:["Anexo","Onde fica (link)"],
      example:["Nome do anexo","Link do documento"]},
    {id:"q7", num:7, title:"Quais termos precisam de definição?",
      dica:'Um termo por linha, com a definição em palavras simples. Serve para o leitor não ficar com dúvida. Se não houver, deixe vazio.',
      label:"Política · Termos e definições", type:"table", cellType:"textarea",
      columns:["Termo","Definição"],
      example:["Termo","O que ele significa"]},
    {id:"q8", num:8, title:"Quais práticas são obrigatórias?",
      dica:'São as partes do processo que não podem ser ignoradas porque previnem ineficiência, retrabalho, risco ou perda financeira. Uma prática por linha, com o motivo (o que ela previne).',
      label:"Política · Práticas obrigatórias", type:"table", cellType:"textarea",
      columns:["Prática obrigatória","Por que é obrigatória"],
      example:["Regra que não pode ser ignorada","O que ela previne"]},
    {id:"q9", num:9, title:"A quem a pessoa deve recorrer em caso de dúvida?",
      dica:'Diga qual área contatar e por qual canal. Cada linha vira um parágrafo.',
      label:"Política · Dúvidas", type:"text", autoGrow:true},
    {id:"q10", num:10, title:"Quem é o gestor do processo (quem revisa)?",
      dica:'Nome de quem revisa a política. Se ainda não souber, deixe em branco: o Word sai com "a definir".',
      label:"Política · Revisada por", type:"text"},
    {id:"q11", num:11, title:"Quem é o diretor da área (quem aprova)?",
      dica:'Nome de quem aprova a política. Se ainda não souber, deixe em branco: o Word sai com "a definir".',
      label:"Política · Aprovado por", type:"text"},
    {id:"q12", num:12, title:"Em que data a política foi revisada?",
      dica:'Formato DD/MM/AAAA. Se ainda não foi revisada, deixe em branco.',
      label:"Política · Revisado em", type:"text"},
    {id:"q13", num:13, title:"Em que data a política foi aprovada?",
      dica:'Formato DD/MM/AAAA. Com a data, o formulário calcula a próxima revisão: 12 meses depois, como manda a Governança das Políticas (GG01-C: "revisadas anualmente"). Se ainda não foi aprovada, deixe em branco: a próxima revisão sai "a definir".',
      label:"Política · Data de aprovação", type:"text"},
    {id:"q14", num:14, title:"Tem mais alguma coisa que quem for revisar a política precisa saber?",
      dica:"Pontos em aberto, dúvidas sobre o conteúdo, ideias para a próxima revisão. Isso não vai para o Word.",
      label:null, type:"text"}
  ];
'''

# ------------------------------------------------------------------ motor do Word
ENGINE_JS = r'''function oneLine(s){ return String(s || "").replace(/\s*[\r\n]+\s*/g, " ").trim(); }

  function formatDateBR(iso){
    var m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(iso || "");
    return m ? (m[3] + "/" + m[2] + "/" + m[1]) : "";
  }

  function pad2(n){ return (n < 10 ? "0" : "") + n; }

  function parseDateBR(s){
    var m = /^\s*(\d{1,2})[\/.\-](\d{1,2})[\/.\-](\d{2}|\d{4})\s*$/.exec(s || "");
    if(!m) return null;
    var d = +m[1], mo = +m[2], y = +m[3];
    if(y < 100) y += 2000;
    if(mo < 1 || mo > 12 || d < 1 || d > 31) return null;
    return {d: d, m: mo, y: y};
  }

  // data digitada: se for uma data valida, padroniza em DD/MM/AAAA; senao devolve o texto como foi digitado
  function dataOuTexto(s){
    var o = parseDateBR(s);
    return o ? (pad2(o.d) + "/" + pad2(o.m) + "/" + o.y) : oneLine(s);
  }

  // proxima revisao = 12 meses apos a aprovacao (DD/MM/AA); 29/02 vira 28/02
  function proximaRevisao(s){
    var o = parseDateBR(s);
    if(!o) return "";
    var d = (o.m === 2 && o.d === 29) ? 28 : o.d;
    return pad2(d) + "/" + pad2(o.m) + "/" + String(o.y + 1).slice(2);
  }

  function valoresPolitica(){
    var AD = "a definir";
    var nome = oneLine(document.getElementById("id_tarefa").value) || "(não informado)";
    var elabEm = formatDateBR(document.getElementById("id_data").value) || formatDateBR(todayISO());
    var aprovEm = oneLine(document.getElementById("q13").value);
    return {
      NOME_PROCESSO: nome,
      NOME_PROC_MAI: nome.toLocaleUpperCase("pt-BR"),
      STATUS: "Criação",
      PROXIMA_REVISAO: proximaRevisao(aprovEm) || AD,
      ELAB_EM: elabEm,
      ELAB_POR: oneLine(document.getElementById("id_nome").value) || "(não informado)",
      REV_EM: dataOuTexto(document.getElementById("q12").value) || AD,
      REV_POR: oneLine(document.getElementById("q10").value) || AD,
      APR_EM: dataOuTexto(aprovEm) || AD,
      APR_POR: oneLine(document.getElementById("q11").value) || AD
    };
  }

  function trocaTodos(xml, valores, chaves){
    chaves.forEach(function(k){ xml = xml.split("{{" + k + "}}").join(escapeHtml(valores[k])); });
    return xml;
  }

  // cabecalho e rodape do Word
  function buildFilledPartXml(xml){
    return trocaTodos(xml, valoresPolitica(), ["NOME_PROCESSO", "STATUS", "PROXIMA_REVISAO", "ELAB_POR", "REV_POR", "APR_POR", "APR_EM"]);
  }

  // pede ao Word para atualizar o sumario ao abrir
  function addUpdateFields(xml){
    if(xml.indexOf("w:updateFields") !== -1) return xml;
    if(xml.indexOf("<w:hdrShapeDefaults") !== -1) return xml.replace("<w:hdrShapeDefaults", '<w:updateFields w:val="true"/><w:hdrShapeDefaults');
    return xml.replace("<w:compat>", '<w:updateFields w:val="true"/><w:compat>');
  }

  // repete um bloco de 2 paragrafos (A = titulo/regra, B = texto, 1 paragrafo por linha)
  function repeatBlock(xml, tokA, tokB, rows, emptyA){
    var litA = "{{" + tokA + "}}", litB = "{{" + tokB + "}}";
    var fa = findEnclosingBlock(xml, litA, /<w:p(?=[ >])/, "</w:p>");
    var fb = findEnclosingBlock(xml, litB, /<w:p(?=[ >])/, "</w:p>");
    if(!fa || !fb) return xml;
    function fillOne(block, lit, text){
      var tag = new RegExp('<w:t[^>]*>' + escapeRegex(lit) + '</w:t>');
      return block.replace(tag, function(){ return '<w:t xml:space="preserve">' + escapeHtml(text) + '</w:t>'; });
    }
    var out = "";
    if(rows.length === 0){
      if(emptyA) out = fillOne(fa.block, litA, emptyA);
    } else {
      rows.forEach(function(r){
        out += fillOne(fa.block, litA, r.a);
        String(r.b || "").split(/\r?\n/).map(function(l){ return l.trim(); }).filter(function(l){ return l !== ""; })
          .forEach(function(l){ out += fillOne(fb.block, litB, l); });
      });
    }
    return xml.slice(0, fa.start) + out + xml.slice(fb.end);
  }

  // repete um paragrafo com dois marcadores (termo em negrito + definicao)
  function repeatTerms(xml, rows){
    var fa = findEnclosingBlock(xml, "{{TERMO}}", /<w:p(?=[ >])/, "</w:p>");
    if(!fa) return xml;
    var out;
    if(rows.length === 0){
      out = fa.block.replace("{{TERMO}}", "(não informado)").replace("{{DEFINICAO}}", "")
        .replace(/<w:r(?: [^>]*)?>(?:(?!<\/w:r>)[\s\S])*?<w:t[^>]*>: <\/w:t><\/w:r>/, "");
    } else {
      out = rows.map(function(r){
        return fa.block
          .replace(new RegExp('<w:t[^>]*>' + escapeRegex("{{TERMO}}") + '</w:t>'), function(){ return '<w:t xml:space="preserve">' + escapeHtml(r[0]) + '</w:t>'; })
          .replace(new RegExp('<w:t[^>]*>' + escapeRegex("{{DEFINICAO}}") + '</w:t>'), function(){ return '<w:t xml:space="preserve">' + escapeHtml(r[1]) + '</w:t>'; });
      }).join("");
    }
    return xml.slice(0, fa.start) + out + xml.slice(fa.end);
  }

  function buildFilledDocumentXml(xml){
    var NI = "(não informado)";
    xml = trocaTodos(xml, valoresPolitica(), ["NOME_PROCESSO", "NOME_PROC_MAI", "ELAB_EM", "ELAB_POR", "REV_EM", "REV_POR", "APR_EM", "APR_POR"]);
    xml = fillToken(xml, "OBJETIVO", document.getElementById("q1").value, NI);
    xml = fillToken(xml, "APLICACAO", document.getElementById("q2").value, NI);
    xml = fillToken(xml, "RESPONSAVEL", document.getElementById("q3").value, NI);
    xml = fillToken(xml, "VISAO_GERAL", document.getElementById("q4").value, NI);
    var subs = getTableRowsFiltered("q5").map(function(r, i){
      return {a: "4." + (i + 1) + " " + (oneLine(r[0]) || "(sem título)"), b: r[1]};
    });
    xml = repeatBlock(xml, "SUB_TITULO", "SUB_TEXTO", subs, "");
    var anexos = getTableRowsFiltered("q6").map(joinRow);
    xml = fillToken(xml, "ANEXO", anexos.join("\n"), "Esta Política não possui anexos vinculados no momento de sua publicação.");
    xml = repeatTerms(xml, getTableRowsFiltered("q7").map(function(r){ return [oneLine(r[0]), oneLine(r[1])]; }));
    var prat = getTableRowsFiltered("q8").map(function(r){ return {a: oneLine(r[0]) || NI, b: r[1]}; });
    xml = repeatBlock(xml, "PRATICA", "PORQUE", prat, NI);
    xml = fillToken(xml, "DUVIDAS", document.getElementById("q9").value, NI);
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

ZIP_JS = '''      zip["word/document.xml"] = new TextEncoder().encode(filledXml);
      ["word/header1.xml", "word/footer1.xml"].forEach(function(nome){
        if(zip[nome]) zip[nome] = new TextEncoder().encode(buildFilledPartXml(new TextDecoder("utf-8").decode(zip[nome])));
      });
      if(zip["word/settings.xml"]) zip["word/settings.xml"] = new TextEncoder().encode(addUpdateFields(new TextDecoder("utf-8").decode(zip["word/settings.xml"])));
'''


def main(pop_html, modelo_docx, out_html):
    s = open(pop_html, encoding="utf-8").read()

    m = re.search(r'(var TEMPLATE_DOCX_B64\s*=\s*")([^"]*)(")', s)
    if not m:
        raise SystemExit("Nao achei TEMPLATE_DOCX_B64")
    s = s[:m.start(2)] + "<<B64>>" + s[m.end(2):]

    # --- identidade
    s = rep_re(s, r"<title>.*?</title>", "<title>Política de Departamento | Grupo Gadens</title>")
    s = rep(s, "<h1>G-Docs Formulário de Levantamento de Processo</h1>", "<h1>G-Docs Formulário de Política de Departamento</h1>")
    s = rep_re(s, r"<p>Conte como o seu trabalho é feito.*?</p>",
               "<p>Conte como é o processo — o formulário monta a Política de Departamento no modelo do Grupo Gadens.</p>")
    s = rep(s, 'var STORAGE_KEY = "gadens_levantamento_pop_v1";', 'var STORAGE_KEY = "gadens_politica_v1";')
    s = rep(s, 'var THEME_KEY = "gadens_pop_theme";', 'var THEME_KEY = "gadens_politica_theme";')

    s = rep_re(s, r'<div class="card instrucoes">.*?</ul>\s*</div>', '''<div class="card instrucoes">
    <h2>Antes de começar</h2>
    <p style="font-size:.92rem; margin:0 0 8px;">Este formulário ajuda você a escrever uma Política de Departamento no modelo do Grupo Gadens: objetivo, aplicação, responsável, processo, anexos, termos e definições, práticas obrigatórias, dúvidas e aprovação. Você responde com as suas palavras; o formulário monta o Word.</p>
    <ul>
      <li>Escreva em frases curtas. Cada linha vira um parágrafo na política.</li>
      <li>O botão "?" de cada pergunta mostra como preencher e, quando há, um exemplo de política já publicada.</li>
      <li>O número da política (GGXX) sai em vermelho no Word: quem define a numeração é o setor de Compliance.</li>
      <li>A próxima revisão é calculada pela data de aprovação, 12 meses depois. Sem a data, sai "a definir".</li>
      <li>Nas tabelas, a linha cinza em itálico é só um exemplo — ela não é salva.</li>
      <li>Ao abrir o Word, ele pergunta se pode atualizar os campos: responda "Sim", é para atualizar o sumário.</li>
      <li>Suas respostas ficam salvas neste computador enquanto você preenche. Gere o Word antes de fechar a página.</li>
    </ul>
  </div>''')
    s = rep(s, 'Nome da tarefa ou rotina descrita\n          <span class="help-icon" tabindex="0">?<span class="tooltip-text">Título do documento</span></span>',
               'Nome do processo (título da política)\n          <span class="help-icon" tabindex="0">?<span class="tooltip-text">Vira o título da política e o nome do arquivo</span></span>')
    s = rep(s, 'placeholder="Use o nome pelo qual ela é conhecida no dia a dia"', 'placeholder="Ex.: Gestão de Pagamentos"')
    s = rep_re(s, r'<div class="card terminou">.*?</div>', '''<div class="card terminou">
    <h2>Terminou?</h2>
    <p>Releia rapidinho suas respostas e clique em "Enviar Respostas (gerar Word)" — isso monta a política já no modelo. Abra o Word e dê uma conferida antes de enviar ao setor de Compliance.</p>
    <p>O nome de quem preenche vai em "Elaborado por".</p>
  </div>''')
    s = rep(s, "Grupo Gadens — Levantamento de Processo (base para POP)", "Grupo Gadens — Política de Departamento")

    # --- perguntas
    s = rep_re(s, r"  var QUESTIONS = \[.*?\n  \];\n", QUESTIONS_JS)

    # --- obrigatorias
    s = rep(s, '{label:"Nome da tarefa ou rotina", el:', '{label:"Nome do processo", el:')
    s = rep(s, 'var q1 = document.getElementById("q1");\n    if(q1.value.trim() === "") missing.push({label:"Pergunta 1 — Objetivo do trabalho", el: q1});',
               'var q1 = document.getElementById("q1");\n    if(q1.value.trim() === "") missing.push({label:"Pergunta 1 — Objetivo da política", el: q1});\n'
               '    var q2 = document.getElementById("q2");\n    if(q2.value.trim() === "") missing.push({label:"Pergunta 2 — Aplicação", el: q2});\n'
               '    var q3 = document.getElementById("q3");\n    if(q3.value.trim() === "") missing.push({label:"Pergunta 3 — Responsável pelo processo", el: q3});\n'
               '    var q4 = document.getElementById("q4");\n    if(q4.value.trim() === "") missing.push({label:"Pergunta 4 — Visão geral do processo", el: q4});')
    s = rep(s, 'if(!isQuestionFilled(qById["q9"])) missing.push({label:"Pergunta 9 — Passo a passo (pelo menos 1 linha preenchida)", el: document.getElementById("wrap_q9")});',
               '/* as demais perguntas são opcionais */')

    # --- exportacao .txt
    s = rep(s, '"Sem nome da tarefa"', '"Sem nome do processo"')
    s = rep(s, 'lines.push("LEVANTAMENTO DE PROCESSO — BASE PARA POP");', 'lines.push("POLÍTICA DE DEPARTAMENTO");')
    s = rep(s, 'lines.push("Nome da tarefa ou rotina descrita: "', 'lines.push("Nome do processo: "')

    # --- motor do Word
    s = rep_re(s, r"function buildAplicacaoText\(\)\{.*?\n  \}\n", "@@ENGINE@@\n")
    s = rep_re(s, r"  function buildFilledDocumentXml\(xml\)\{.*?\n    return xml;\n  \}\n", "")
    s = s.replace("@@ENGINE@@\n", ENGINE_JS)
    s = rep(s, "  function substituteSimpleToken(xml, token, value){", WT_JS + "  function substituteSimpleToken(xml, token, value){")
    s = rep(s, """          var re = new RegExp('<w:t[^>]*>' + escapeRegex("{{" + tok + "}}") + '</w:t>');
          var val = (rowValues[i] || "").trim();
          block = block.replace(re, '<w:t xml:space="preserve">' + escapeHtml(val) + '</w:t>');""",
               """          block = fillToken(block, tok, rowValues[i], "");""")
    s = rep(s, '      zip["word/document.xml"] = new TextEncoder().encode(filledXml);', ZIP_JS.rstrip("\n"))
    s = rep(s, 'var filename = "POP - " + tarefa', 'var filename = "Política - " + tarefa')
    s = rep(s, "Abra e revise antes de enviar à equipe de Processos — código do documento, fluxo Bizagi, aprovações e controle de emissão continuam sendo preenchidos por eles.",
               "Abra e revise antes de enviar ao setor de Compliance — o número da política (GGXX, em vermelho) é definido por ele. O Word vai perguntar se pode atualizar os campos: responda Sim, é o sumário.")

    # --- modo por voz
    s = rep(s, 'title:"Qual é o nome da tarefa ou rotina que você vai descrever?", dica:"Use o nome pelo qual ela é conhecida no dia a dia."',
               'title:"Qual é o nome do processo desta política?", dica:"Esse nome vira o título da política e o nome do arquivo."')

    # --- exemplos
    s = rep_re(s, r"  var EXEMPLOS = \{.*?\n  var EXEMPLOS_NOTAS = \{.*?\};\n",
               "  var EXEMPLOS = " + json.dumps(EXEMPLOS, ensure_ascii=False) + ";\n  var EXEMPLOS_NOTAS = " + json.dumps(EXEMPLOS_NOTAS, ensure_ascii=False) + ";\n")
    s = rep(s, "// Cada exemplo é um trecho de POP já aprovado", "// Cada exemplo é um trecho de política já publicada")
    s = rep(s, '(lista.length > 1 ? "Exemplos de POPs já publicados" : "Exemplo de POP já publicado")',
               '(lista.length > 1 ? "Exemplos de políticas já publicadas" : "Exemplo de política já publicada")')

    b64 = base64.b64encode(open(modelo_docx, "rb").read()).decode("ascii")
    s = s.replace("<<B64>>", b64)
    open(out_html, "w", encoding="utf-8").write(s)
    print("ok ->", out_html, len(s), "chars")


if __name__ == "__main__":
    if len(sys.argv) != 4:
        raise SystemExit(__doc__)
    main(*sys.argv[1:])
