# -*- coding: utf-8 -*-
"""Monta o 'Formulario de Levantamento de Regimento' a partir do formulario de POP.

Uso:
  python montar_formulario_regimento.py <formulario_POP.html> <modelo_RE_com_marcadores.docx> <saida.html>

- <formulario_POP.html>: formulario interativo de POP (fonte do layout, voz, progresso, tema).
- <modelo_RE_com_marcadores.docx>: gerado por montar_modelo_regimento.py (embutido no HTML em base64).
As perguntas seguem as 9 secoes do RE.GG.PRO.001; os exemplos do botao "?" vem do RE.GG.NNG.001
(Comites de Novos Negocios). Se o formulario de POP mudar, basta rodar de novo (o script avisa se
algum trecho nao for encontrado).
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


# ------------------------------------------------------------------ exemplos (RE.GG.NNG.001)
F = "RE.GG.NNG.001 · Comitês de Novos Negócios"
DD = "RE.IG.PJC.001 · Comitê de Due Diligence (Incorporadora Gadens)"
CE = "RE.GG.PJC.006 · Comitê Extraordinário de Revisão de Projeto (CERP)"
AN = "RE.CM.PJC.003 · Comitês de Anteprojeto (Make It)"
EXEMPLOS = {
    "q1": [
        {"fonte": F, "texto":
            "V0: Avaliar estrategicamente terrenos identificados como oportunidades de negócio, promovendo o alinhamento entre as áreas envolvidas e validando as premissas iniciais técnicas, comerciais, mercadológicas e econômico-financeiras que suportam a viabilidade preliminar do empreendimento.\n"
            "V2: Apresentar à Diretoria os resultados consolidados das análises realizadas na fase final de Due Diligence das oportunidades de aquisição ou permuta de terrenos."},
        {"fonte": DD, "texto": "Este regimento define o funcionamento dos Comitês de Desenvolvimento da Due Diligence da Incorporadora Gadens. Este Regimento tem como propósito garantir alinhamento entre setores, validação formal de entregas, registro de decisões e controle de riscos antes da passagem de fase."},
        {"fonte": CE, "texto": "Estabelecer o fluxo de amostragem e deliberação formal das alterações de projeto identificadas após o início de obra, por meio do Comitê Extraordinário de Revisão de Projeto (CERP), registrando as soluções alteradas, os itens modificados e os projetos revisados, e formalizando sua liberação para execução."},
    ],
    "q2": [
        {"fonte": F, "texto": "Esta etapa tem como finalidade concluir a análise técnica, financeira e operacional do empreendimento, permitindo a deliberação final da diretoria quanto à viabilidade do negócio e à decisão de prosseguir (“GO”) ou não prosseguir (“NO GO”) com a aquisição do terreno."},
        {"fonte": DD, "texto": "Este comitê é uma instância deliberativa e consultiva. Não substitui reuniões operacionais do dia a dia. Toda decisão relevante de escopo, custo, prazo ou risco deve ser registrada em ata e tratada formalmente neste fórum."},
        {"fonte": AN + " · Papéis (Diretoria de Incorporação)", "texto": "Deliberador. Aprova especificações e fachada; autoriza avanço ao PE"},
    ],
    "q3": [
        {"fonte": F, "texto": "Aplica-se às incorporadoras Gadens e Make It."},
        {"fonte": CE, "texto": "Este regimento aplica-se a todos os empreendimentos da Gadens Incorporadora e Make It Construtora, a partir do início de obra (pós Comitê EX / Fase EX)."},
        {"fonte": AN, "texto": "Este regimento aplica-se à Make It, abrangendo os setores de Incorporação, Arquitetura, Legalização, Projetos Complementares, Orçamentos, Planejamento, Gestão de Obras, Diretoria de Engenharia e Diretoria de Incorporação, envolvidos no ciclo de desenvolvimento da fase de Projeto Anteprojeto."},
    ],
    "q4": [
        {"fonte": F, "texto":
            "Novos Negócios (responsável por demandar, organizar e conduzir as etapas de V0 e V1)\n"
            "Incorporação Gadens e Make It (responsável por demandar, organizar e conduzir a etapa de V2)\n"
            "Comercial\nJurídico\nControladoria\nFinanceiro\nDiretoria"},
        {"fonte": CE, "texto": "Áreas envolvidas: Engenharia de Obra, Qualidade, Projetos Complementares, Arquitetura e Orçamentos."},
        {"fonte": DD, "texto": "As áreas envolvidas nos comitês são: Incorporação, Arquitetura e Legalização, Projetos Complementares, Orçamentos, Planejamento e Controle (acompanhamento), Diretoria de Engenharia, Diretoria de Incorporação e demais Diretorias."},
    ],
    "q5": [
        {"fonte": F, "texto":
            "A Etapa de V0, é recorrente em reuniões semanais, conduzidas pela área de Novos Negócios, responsável por organizar a pauta, consolidar os materiais e encaminhar previamente as informações necessárias às áreas participantes."},
        {"fonte": DD, "texto": "Incorporação, responsável pela convocação e pela ata, conforme a tabela de papéis do título 5 (Comitê Due)."},
        {"fonte": CE, "texto": "Analista de Projetos (Arquitetura / Complementares), responsável pela convocação, conforme a tabela de papéis do título 5. Convoca o comitê sob demanda, definindo data, pauta e participantes obrigatórios e condicionais."},
    ],
    "q6": [
        {"fonte": CE + " · Papéis (resumido)", "texto":
            "Engenheiro Residente: Obrigatório.\nRepresentante da Qualidade: Obrigatório.\nRepresentante de Orçamentos: Participação se necessário.\nGestor(a) de Projetos Complementares: Participação se necessário."},
    ],
    "q7": [
        {"fonte": F, "texto": "A Etapa de V0, será realizado semanalmente, às terças-feiras, das 9h às 10h, podendo ocorrer convocações extraordinárias conforme necessidade e disponibilidade das áreas envolvidas."},
        {"fonte": DD, "texto": "Não há periodicidade recorrente: este comitê ocorre uma única vez por empreendimento, nesta fase do ciclo de desenvolvimento."},
        {"fonte": CE, "texto": "Não há periodicidade fixa: o comitê é convocado sob demanda, sempre que houver alteração de projeto identificada em campo ou solicitada pela equipe técnica após o início de obra."},
    ],
    "q9": [
        {"fonte": F, "texto": "Reunião preferencialmente presencial, podendo ocorrer de forma virtual quando necessário."},
        {"fonte": DD, "texto": "Formato preferencialmente presencial trazendo o modelo online para exceções."},
        {"fonte": CE, "texto": "Preferencialmente presencial trazendo o modelo online para exceções."},
    ],
    "q10": [
        {"fonte": F, "texto": "A Etapa de V0: As oportunidades a serem avaliadas deverão ser previamente estruturadas pela área de Novos Negócios, contendo as premissas iniciais do negócio, incluindo informações do terreno, análise preliminar de mercado, estimativas iniciais de custos e demais dados relevantes para a avaliação da viabilidade preliminar."},
        {"fonte": DD, "texto":
            "Equipe e Projetistas Envolvidos (Responsabilidade: Incorporação e Projetos Complementares).\n"
            "Lista da equipe interna com analistas responsáveis por disciplina\n"
            "Lista de projetistas contratados ou em processo de contratação, por disciplina\n"
            "Status de gastos da Due Diligence: tabela por disciplina com valores comprometidos\n"
            "Linha do tempo da Due Diligence"},
        {"fonte": AN, "texto":
            "Status da Fase EP (Responsabilidade: Incorporação e Projetos Complementares).\n"
            "Confirmação dos itens aprovados no Comitê EP\n"
            "Pendências do EP: resolvidas ou justificativa formal de não atendimento"},
    ],
    "q11": [
        {"fonte": F, "texto": "A Etapa de V0 e V1: As decisões do comitê serão registradas em ata, elaborada pela área de Novos Negócios. Na V0, a ata conterá as principais premissas discutidas e a deliberação quanto à continuidade ou não da oportunidade."},
        {"fonte": DD, "texto": "Toda decisão relevante de escopo, custo, prazo ou risco deve ser registrada em ata."},
        {"fonte": CE, "texto": "Não há modelo de ata previamente padronizado. A ata deve conter, no mínimo: empreendimento e identificação da ocorrência; data, horário e participantes (com indicação de obrigatoriedade atendida); amostragem apresentada (item alterado, solução, disciplina e projeto revisado); classe da alteração (conforme PO 04.007); decisão do comitê (aprovado, aprovado com ressalvas ou reprovado); ações, responsáveis e prazos definidos; e data de liberação em CDE. A ata deve ser elaborada e distribuída pelo Analista responsável pela convocação em até 2 dias úteis após a reunião."},
    ],
    "q12": [
        {"fonte": DD + " · Saídas (Ata de reunião)", "texto": "Registro formal das deliberações, decisões, pendências e encaminhamentos do Comitê, distribuída em até 2 dias úteis após a reunião."},
        {"fonte": DD + " · Saídas (Prazos de devolução)", "texto": "Pendências abordadas em reunião por ausência de informação ou complemento. Deverá ser programada entrega aos participantes."},
    ],
    "q13": [
        {"fonte": F + " · Etapa de V0", "linhas": [
            ["Diretoria", "Avaliar a viabilidade econômico-financeira do negócio com base nas premissas apresentadas pelas áreas envolvidas, considerando indicadores de margem, exposição de caixa e demais aspectos estratégicos, deliberando sobre o avanço ou não da oportunidade."],
            ["Jurídico", "Identificar riscos jurídicos relacionados ao terreno, vendedores e/ou permutantes, bem como eventuais restrições contratuais;\nApresentar o resultado das consultas e certidões públicas relacionadas ao imóvel e às partes envolvidas, se aplicável."],
        ]},
        {"fonte": DD, "linhas": [
            ["Incorporação", "Responsável pela convocação e ata. Apresenta o empreendimento, viabilidade e decisões comerciais"],
            ["Gestão de Obras", "Participante consultivo. Validador técnico das soluções pré-concebidas em projeto e análise de riscos"],
            ["Diretoria de Engenharia", "Deliberador. Aprova e questiona soluções e estratégias"],
        ]},
        {"fonte": CE, "linhas": [
            ["Engenheiro Residente", "Obrigatório. Identificar ou relatar a alteração em obra; validar a viabilidade construtiva da solução revisada; formalizar a necessidade de revisão."],
            ["Representante de Orçamentos", "Participação se necessário. Apresentar o impacto financeiro da alteração, quando houver variação de custo."],
        ]},
    ],
    "q14": [
        {"fonte": F + " · Etapa de V0", "linhas": [
            ["Relatório de Engenharia", "Avaliação preliminar das condições técnicas do terreno, incluindo possíveis restrições construtivas.", "Engenharia"],
            ["Prazos Estimados", "Estimativa preliminar de prazos para aprovação de projetos e desenvolvimento do empreendimento.", "Incorporação"],
        ]},
        {"fonte": CE, "linhas": [
            ["Apontamento Construflow", "Referente à alteração identificada ou solicitada, cadastrado conforme PO 04.004.", "Solicitante (Engenheiro Residente / Analistas / Gestores)"],
            ["Pauta do comitê", "Com a amostragem a ser apresentada.", "Analista de Projetos (Arquitetura / Complementares)"],
        ]},
        {"fonte": AN, "linhas": [
            ["Orçamento do Anteprojeto", "Apresentação de custos de itens específicos.", "Orçamentos"],
        ]},
    ],
    "q15": [
        {"fonte": F + " · Etapa de V2", "linhas": [
            ["Decisão do Comitê (Go / No Go / Ajustes)", "Deliberação formal do comitê quanto à continuidade do empreendimento, podendo resultar em aprovação (Go), reprovação (No Go) ou solicitação de ajustes adicionais.", "Comitê"],
            ["Ata da Reunião", "Registro formal das discussões, premissas avaliadas e decisão tomada no comitê, garantindo rastreabilidade e governança do processo.", "Incorporação"],
        ]},
        {"fonte": DD, "linhas": [
            ["Decisão de Comitê Due", "Decisão da Diretoria sobre viabilidade e prosseguimento da aquisição do terreno.", "Diretoria de Incorporação e Diretoria de Engenharia"],
            ["Ata de reunião", "Registro formal das deliberações, decisões, pendências e encaminhamentos do Comitê, distribuída em até 2 dias úteis após a reunião.", "Analista responsável pela convocação"],
        ]},
        {"fonte": CE, "linhas": [
            ["Liberação em CDE", "Atualização do status do projeto revisado no Autodoc para “Liberado para Obra”, após deliberação favorável do comitê.", "Analista de Projetos Complementares / Arquitetura"],
        ]},
    ],
    "q16": [
        {"fonte": F, "texto": "Modelo de apresentação de oportunidade de terreno\nModelo de Ata padrão."},
        {"fonte": DD, "texto": "PO.IG.PJC.001 – Controle de Alterações de Projeto."},
        {"fonte": CE, "texto": "PO 04.004 – Apontamentos no Construflow;\nPO 04.007 – Escala de classificação de alterações (Épico, Marco, Estória ou Atividade)."},
    ],
}
EXEMPLOS_NOTAS = {
    "q1": "Quando a reunião tem etapas (como V0, V1 e V2 no Comitê de Novos Negócios), escreva um objetivo por etapa, começando pelo nome da etapa.",
    "q4": "Nos Regimentos publicados as áreas aparecem em uma frase corrida. Aqui, escreva uma por linha: o formulário monta a lista no Word.",
    "q6": "Nenhum Regimento publicado fixa um mínimo de presentes. O exemplo mostra como os Regimentos indicam quem é obrigatório e quem participa só se necessário. Esta pergunta existe para cobrir essa lacuna.",
    "q8": "Os Regimentos publicados não tratam desse ponto. Esta pergunta existe para cobrir essa lacuna. Se não houver regra, escreva \"não há regra\".",
    "q13": "Se a reunião tem etapas, repita a mesma área quantas vezes precisar e comece o texto com a etapa. Ex.: “V0 — avaliar …”.",
}

# ------------------------------------------------------------------ perguntas
T = "Regimento · Seção "
QUESTIONS_JS = r'''  var QUESTIONS = [
    {id:"q1", num:1, title:"Para que existe essa reunião ou comitê? O que ela quer alcançar?",
      dica:'Escreva em poucas frases. Se a reunião tem etapas (ex.: V0, V1, V2), escreva uma etapa por linha, começando pelo nome dela.',
      label:"Regimento · Seção 1 — Objetivo", type:"text", required:true, autoGrow:true},
    {id:"q2", num:2, title:"O que a reunião decide? Quais respostas ela pode dar e como a decisão é tomada?",
      dica:'Ex.: aprovar (Go), reprovar (No Go) ou pedir ajustes. Diga quem tem a palavra final ou se vale a maioria. Se não houver decisão, escreva "não há decisão".',
      label:"Regimento · Seção 1 — Objetivo", type:"text", autoGrow:true},
    {id:"q3", num:3, title:"Para quais empresas do grupo essa reunião vale?",
      dica:'Ex.: "Aplica-se às incorporadoras Gadens e Make It." Se valer para todas, escreva "Grupo Gadens".',
      label:"Regimento · Seção 2 — Aplicação", type:"text"},
    {id:"q4", num:4, title:"Quais áreas participam da reunião? O que cada uma faz, em uma frase? A presença é obrigatória ou só quando necessário?",
      dica:'Uma área por linha. Ex.: "Jurídico (obrigatório)"; "Marketing (só quando necessário)". Indique qual área organiza e conduz a reunião.',
      label:"Regimento · Seção 3 — Áreas envolvidas", type:"text", autoGrow:true},
    {id:"q5", num:5, title:"Quem convoca e agenda a reunião? O que dá a largada para ela acontecer?",
      dica:'Ex.: "A área de Novos Negócios convoca, organiza a pauta e envia o material com 2 dias úteis de antecedência." Diga também se a reunião acontece quando alguém pede ou depois de outra etapa.',
      label:"Regimento · Seção 4 — Governança · Convocação/Agendamento", type:"text", autoGrow:true},
    {id:"q6", num:6, title:"Quantas pessoas ou áreas precisam estar presentes para a reunião valer? E se faltarem, o que acontece?",
      dica:'Ex.: "Precisa de Diretoria e Jurídico presentes. Sem eles, a reunião é remarcada." Se não houver mínimo, escreva "não há mínimo".',
      label:"Regimento · Seção 4 — Governança · Convocação/Agendamento", type:"text", autoGrow:true},
    {id:"q7", num:7, title:"Com que frequência a reunião acontece? Em que dia da semana e horário?",
      dica:'Ex.: "Semanalmente, às terças-feiras, das 9h às 10h." Diga também se pode haver reunião extra.',
      label:"Regimento · Seção 4 — Governança · Periodicidade", type:"text", autoGrow:true},
    {id:"q8", num:8, title:"E se a reunião não puder acontecer na data? O que se faz?",
      dica:'Ex.: "Remarcar para a mesma semana" ou "O assunto vai para a próxima reunião". Se não houver regra, escreva "não há regra".',
      label:"Regimento · Seção 4 — Governança · Periodicidade", type:"text", autoGrow:true},
    {id:"q9", num:9, title:"A reunião é presencial ou online? Existe alguma exceção?",
      dica:'Ex.: "Preferencialmente presencial, podendo ser online quando necessário." Se for presencial, diga o local.',
      label:"Regimento · Seção 4 — Governança · Formato", type:"text"},
    {id:"q10", num:10, title:"O que é tratado em cada reunião? Que material precisa ser levado?",
      dica:'Diga os assuntos da pauta e quem prepara o material. Se a reunião tem etapas, escreva uma etapa por linha, começando por "A Etapa de V0:".',
      label:"Regimento · Seção 4 — Governança · Pauta/Material de Apoio", type:"text", autoGrow:true},
    {id:"q11", num:11, title:"O que fica registrado depois da reunião? Quem faz a ata e onde ela é guardada?",
      dica:'Ex.: "A ata é feita pela área de Incorporação e guardada na pasta do comitê no SharePoint. Ela traz as decisões e os próximos passos."',
      label:"Regimento · Seção 4 — Governança · Registro/Ata", type:"text", autoGrow:true},
    {id:"q12", num:12, title:"Qual o prazo para enviar a ata? Existe prazo para as áreas resolverem as pendências que a reunião deixou?",
      dica:'Ex.: "A ata é enviada em até 2 dias úteis. As pendências são devolvidas em até 5 dias úteis." Se não houver prazo, escreva "não há prazo".',
      label:"Regimento · Seção 4 — Governança · Registro/Ata", type:"text", autoGrow:true},
    {id:"q13", num:13, title:"O que cada área faz na reunião? E o que decide?",
      dica:'Uma área por linha. Se a reunião tem etapas, repita a área e comece o texto com a etapa (ex.: "V0 — avaliar…"). Pode escrever mais de uma responsabilidade na mesma célula, uma por linha.',
      label:"Regimento · Seção 5 — Papéis e Responsabilidades", type:"table", cellType:"textarea", required:true,
      columns:["Área","Responsabilidades"],
      example:["Jurídico","Identificar riscos jurídicos do terreno e das partes envolvidas."]},
    {id:"q14", num:14, title:"O que precisa estar pronto para a reunião acontecer? Quem entrega?",
      dica:'Um item por linha: o que é, como ele é (bem resumido) e qual setor entrega. Ex.: estudos, orçamento, relatórios, apresentações.',
      label:"Regimento · Seção 6 — Entradas (Inputs)", type:"table", cellType:"textarea",
      columns:["Input","Descrição","Responsável"],
      example:["Estudo de Mercado","Análise de demanda, concorrência e potencial de venda.","Comercial"]},
    {id:"q15", num:15, title:"O que a reunião entrega no fim? Quem é responsável por cada item?",
      dica:'Um item por linha. Pense em: a decisão (ex.: Go / No Go / Ajustes), a ata e os próximos passos.',
      label:"Regimento · Seção 7 — Saídas (Outputs)", type:"table", cellType:"textarea",
      columns:["Output","Descrição","Responsável"],
      example:["Ata da Reunião","Registro das discussões e da decisão tomada.","Incorporação"]},
    {id:"q16", num:16, title:"Existe algum documento, modelo, norma ou procedimento que a reunião usa como base?",
      dica:'Um por linha. Ex.: "Modelo de Ata padrão"; "PO.GG.QLD.002 - Regras de elaboração de documento". Se não houver, escreva "não há".',
      label:"Regimento · Seção 8 — Referências", type:"text", autoGrow:true},
    {id:"q17", num:17, title:"Tem mais alguma coisa que quem for escrever o Regimento precisa saber?",
      dica:"Exceções, casos em que a reunião muda, problemas que já aconteceram, ideias de melhoria — tudo ajuda. Isso não vai para o Word.",
      label:null, type:"text"}
  ];
'''

# ------------------------------------------------------------------ motor do Word
ENGINE_JS = r'''function buildAplicacaoText(){
    var texto = document.getElementById("q3").value.trim();
    if(texto) return texto;
    return document.getElementById("id_empresa").value.trim();
  }

  function joinText(ids){
    return ids.map(function(id){ return (document.getElementById(id).value || "").trim(); })
              .filter(function(t){ return t !== ""; }).join("\n");
  }

  // Itens de lista: cada um termina em ";" e o ultimo em "."
  function punctuateItems(items){
    var clean = items.map(function(s){ return s.replace(/[\s.;,]+$/, ""); }).filter(function(s){ return s !== ""; });
    return clean.map(function(s, i){ return s + (i === clean.length - 1 ? "." : ";"); });
  }

  function buildFilledDocumentXml(xml){
    xml = substituteSimpleToken(xml, "TITULO", document.getElementById("id_tarefa").value);
    xml = substituteSimpleToken(xml, "TITULO", document.getElementById("id_tarefa").value);
    xml = fillToken(xml, "OBJETIVO", joinText(["q1","q2"]), "(não informado)");
    xml = fillToken(xml, "APLICACAO", buildAplicacaoText(), "(não informado)");
    xml = substituteRepeatingParagraph(xml, "AREA_ITEM", punctuateItems(getTextLines("q4")));

    xml = fillToken(xml, "GOV_CONVOCACAO", joinText(["q5","q6"]), "(não informado)");
    xml = fillToken(xml, "GOV_PERIODICIDADE", joinText(["q7","q8"]), "(não informado)");
    xml = fillToken(xml, "GOV_FORMATO", joinText(["q9"]), "(não informado)");
    xml = fillToken(xml, "GOV_PAUTA", joinText(["q10"]), "(não informado)");
    xml = fillToken(xml, "GOV_ATA", joinText(["q11","q12"]), "(não informado)");

    xml = substituteRepeatingRow(xml, ["AREA", "AREA_RESP"], getTableRowsFiltered("q13"));
    xml = substituteRepeatingRow(xml, ["INPUT", "INPUT_DESC", "INPUT_RESP"], getTableRowsFiltered("q14"));
    xml = substituteRepeatingRow(xml, ["OUTPUT", "OUTPUT_DESC", "OUTPUT_RESP"], getTableRowsFiltered("q15"));
    xml = substituteRepeatingParagraph(xml, "REFERENCIA_ITEM", punctuateItems(getTextLines("q16")));
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
    s = rep_re(s, r"<title>.*?</title>", "<title>Levantamento de Regimento | Grupo Gadens</title>")
    s = rep(s, "<h1>G-Docs Formulário de Levantamento de Processo</h1>", "<h1>G-Docs Formulário de Levantamento de Regimento</h1>")
    s = rep_re(s, r"<p>Conte como o seu trabalho é feito.*?</p>",
               "<p>Conte como a reunião ou o comitê funciona — a equipe de Processos transforma em Regimento.</p>")
    s = rep(s, 'var STORAGE_KEY = "gadens_levantamento_pop_v1";', 'var STORAGE_KEY = "gadens_levantamento_re_v1";')
    s = rep(s, 'var THEME_KEY = "gadens_pop_theme";', 'var THEME_KEY = "gadens_re_theme";')

    s = rep_re(s, r'<div class="card instrucoes">.*?</ul>\s*</div>', '''<div class="card instrucoes">
    <h2>Antes de começar</h2>
    <p style="font-size:.92rem; margin:0 0 8px;">Este formulário ajuda você a escrever o Regimento — o documento oficial que explica como uma reunião ou comitê funciona: para que serve, quem participa, quando acontece e o que entrega. Você responde com as suas palavras; o formulário monta o Word no modelo oficial.</p>
    <ul>
      <li>Escreva como se estivesse explicando a reunião a um colega novo. Não precisa "escrever bonito".</li>
      <li>O botão "?" de cada pergunta mostra um exemplo do Regimento dos Comitês de Novos Negócios.</li>
      <li>Se a reunião tem etapas (como V0, V1 e V2), escreva uma etapa por linha e comece cada linha com o nome da etapa.</li>
      <li>Se uma pergunta não se aplicar, escreva "não se aplica"; se não souber, escreva "não sei" e siga em frente.</li>
      <li>Nas tabelas, a linha cinza em itálico é só um exemplo — ela não é salva.</li>
      <li>Suas respostas ficam salvas neste computador enquanto você preenche. Gere o Word antes de fechar a página.</li>
    </ul>
  </div>''')
    s = rep(s, "Área Dona do Processo\n          <span class=\"help-icon\" tabindex=\"0\">?<span class=\"tooltip-text\">Área principal e executora majoritária do processo</span></span>",
               "Área responsável pela reunião\n          <span class=\"help-icon\" tabindex=\"0\">?<span class=\"tooltip-text\">Área que organiza e conduz a reunião ou comitê</span></span>")
    s = rep(s, 'placeholder="Ex.: Contas a receber"', 'placeholder="Ex.: Novos Negócios"')
    s = rep(s, 'Nome da tarefa ou rotina descrita\n          <span class="help-icon" tabindex="0">?<span class="tooltip-text">Título do documento</span></span>',
               'Nome da reunião ou comitê\n          <span class="help-icon" tabindex="0">?<span class="tooltip-text">Será o título do Regimento</span></span>')
    s = rep(s, 'placeholder="Use o nome pelo qual ela é conhecida no dia a dia"', 'placeholder="Ex.: Comitês de Novos Negócios"')
    s = rep_re(s, r'<div class="card terminou">.*?</div>', '''<div class="card terminou">
    <h2>Terminou?</h2>
    <p>Releia rapidinho suas respostas e clique em "Enviar Respostas (gerar Word)" — isso monta o rascunho do Regimento já no modelo oficial. Abra o Word, dê uma conferida e envie para a equipe de Processos.</p>
    <p>Código do documento, aprovações e controle de emissão são preenchidos pela equipe de Processos/Qualidade — você não precisa se preocupar com essa parte.</p>
  </div>''')
    s = rep(s, "Grupo Gadens — Levantamento de Processo (base para POP)", "Grupo Gadens — Levantamento de Regimento")

    # --- perguntas
    s = rep_re(s, r"  var QUESTIONS = \[.*?\n  \];\n", QUESTIONS_JS)

    # --- obrigatorias
    s = rep(s, '{label:"Área Dona do Processo", el:', '{label:"Área responsável pela reunião", el:')
    s = rep(s, '{label:"Nome da tarefa ou rotina", el:', '{label:"Nome da reunião ou comitê", el:')
    s = rep(s, '{label:"Pergunta 1 — Objetivo do trabalho", el: q1}', '{label:"Pergunta 1 — Objetivo da reunião", el: q1}')
    s = rep(s, 'if(!isQuestionFilled(qById["q9"])) missing.push({label:"Pergunta 9 — Passo a passo (pelo menos 1 linha preenchida)", el: document.getElementById("wrap_q9")});',
               'if(!isQuestionFilled(qById["q13"])) missing.push({label:"Pergunta 13 — Papéis (pelo menos 1 área preenchida)", el: document.getElementById("wrap_q13")});')

    # --- exportacao .txt
    s = rep(s, '"Sem nome da tarefa"', '"Sem nome da reunião"')
    s = rep(s, 'lines.push("LEVANTAMENTO DE PROCESSO — BASE PARA POP");', 'lines.push("LEVANTAMENTO DE REGIMENTO");')
    s = rep(s, 'lines.push("Área Dona do Processo: "', 'lines.push("Área responsável pela reunião: "')
    s = rep(s, 'lines.push("Nome da tarefa ou rotina descrita: "', 'lines.push("Nome da reunião ou comitê: "')

    # --- motor do Word
    s = rep_re(s, r"function buildAplicacaoText\(\)\{.*?\n  \}\n", "@@ENGINE@@\n")
    s = rep_re(s, r"  function buildFilledDocumentXml\(xml\)\{.*?\n    return xml;\n  \}\n", "")
    s = s.replace("@@ENGINE@@\n", ENGINE_JS)
    # quebras de linha dentro das celulas
    s = rep(s, "  function substituteSimpleToken(xml, token, value){", WT_JS + "  function substituteSimpleToken(xml, token, value){")
    s = rep(s, """          var re = new RegExp('<w:t[^>]*>' + escapeRegex("{{" + tok + "}}") + '</w:t>');
          var val = (rowValues[i] || "").trim();
          block = block.replace(re, '<w:t xml:space="preserve">' + escapeHtml(val) + '</w:t>');""",
               """          block = fillToken(block, tok, rowValues[i], "");""")
    s = rep(s, 'var filename = "POP - " + tarefa', 'var filename = "Regimento - " + tarefa')
    s = rep(s, "Abra e revise antes de enviar à equipe de Processos — código do documento, fluxo Bizagi, aprovações e controle de emissão continuam sendo preenchidos por eles.",
               "Abra e revise antes de enviar à equipe de Processos — código do documento, aprovações e controle de emissão continuam sendo preenchidos por eles.")

    # --- modo por voz
    s = rep(s, 'title:"Qual é a área dona do processo?", dica:"Exemplo: Contas a receber."',
               'title:"Qual é a área responsável pela reunião?", dica:"Exemplo: Novos Negócios."')
    s = rep(s, 'title:"Para qual empresa do grupo é esse processo?"', 'title:"Para qual empresa do grupo é essa reunião?"')
    s = rep(s, 'title:"Qual é o nome da tarefa ou rotina que você vai descrever?", dica:"Use o nome pelo qual ela é conhecida no dia a dia."',
               'title:"Qual é o nome da reunião ou comitê?", dica:"Esse nome será o título do Regimento."')

    # --- exemplos
    s = rep_re(s, r"  var EXEMPLOS = \{.*?\n  var EXEMPLOS_NOTAS = \{.*?\};\n",
               "  var EXEMPLOS = " + json.dumps(EXEMPLOS, ensure_ascii=False) + ";\n  var EXEMPLOS_NOTAS = " + json.dumps(EXEMPLOS_NOTAS, ensure_ascii=False) + ";\n")
    s = rep(s, "// Cada exemplo é um trecho de POP já aprovado", "// Cada exemplo é um trecho de Regimento já aprovado")
    s = rep(s, '(lista.length > 1 ? "Exemplos de POPs já publicados" : "Exemplo de POP já publicado")',
               '(lista.length > 1 ? "Exemplos de Regimentos já publicados" : "Exemplo de Regimento já publicado")')

    b64 = base64.b64encode(open(modelo_docx, "rb").read()).decode("ascii")
    s = s.replace("<<B64>>", b64)
    open(out_html, "w", encoding="utf-8").write(s)
    print("ok ->", out_html, len(s), "bytes")


if __name__ == "__main__":
    if len(sys.argv) != 4:
        raise SystemExit(__doc__)
    main(*sys.argv[1:])
