# -*- coding: utf-8 -*-
"""Monta o 'Formulario de Levantamento de IT' (versao por blocos) a partir do formulario de POP.

Uso:
  python montar_formulario_it.py <formulario_POP.html> <modelo_IT_com_marcadores.docx> <saida.html>

- <formulario_POP.html>: formulario interativo de POP (fonte do layout, voz, progresso, tema).
- <modelo_IT_com_marcadores.docx>: gerado por montar_modelo_it.py (embutido no HTML em base64).
As perguntas e os exemplos sao baseados na IT.GG.CON.001 (Cadastro de CNO).
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


# ------------------------------------------------------------------ exemplos (IT.GG.CON.001)
FONTE = "IT.GG.CON.001 · Cadastro de CNO"
EXEMPLOS = {
    "q1": [{"fonte": FONTE, "texto": "Demonstrar o procedimento de abertura de CNO, para orientação das áreas envolvidas."}],
    "q2": [{"fonte": FONTE, "texto": "Aplicável para todo o Grupo Gadens, especificamente para abertura de Cadastro Nacional de Obras."}],
    "q3": [{"fonte": FONTE, "texto": "PO.GG.QLD.002 - Regras de elaboração de documento Grupo Gadens."}],
    "q4": [{"fonte": FONTE, "linhas": [
        ["CNO", "Cadastro Nacional de Obras"],
        ["GOV", "Plataforma do Governo (Brasil)"],
        ["UF", "Unidade Federativa"],
        ["CNPJ", "Cadastro Nacional da Pessoa Jurídica"],
        ["CNAE", "Classificação Nacional de Atividades Econômicas"],
        ["ART", "Anotação de Responsabilidade Técnica"],
    ]}],
    "etapas": [
        {"fonte": FONTE + " · Título + Texto com imagem", "texto":
            "TÍTULO (Etapa): Acesso ao cadastro\n\n"
            "TEXTO: Acessar o site https://cav.receita.fazenda.gov.br, sendo necessário entrar com o GOV, utilizando o certificado digital e selecionar a SPE correspondente.\n"
            "[logo abaixo: print da tela de login]"},
        {"fonte": FONTE + " · Texto com imagem", "texto":
            "TEXTO: Irá abrir a tela de Cadastro Nacional de Obras (CNO), sendo necessário para a abertura de um novo CNO, selecionar “Inscrever ou Alterar Obra”.\n"
            "[logo abaixo: print das opções da tela]"},
        {"fonte": FONTE + " · Tabela Campo / Descrição", "linhas": [
            ["Número do alvará", "Código de registro único impresso em uma autorização oficial emitida pelo governo (como uma prefeitura) ou pela Justiça."],
            ["Proprietário", "Pessoa física ou jurídica que consta legalmente na matrícula do terreno ou imóvel no Cartório de Registro de Imóveis"],
            ["Município", "Cidade onde o imóvel está localizado e qual prefeitura emitiu e fiscaliza o documento"],
        ]},
        {"fonte": FONTE + " · Lista de opções", "texto":
            "TEXTO: As opções serão:\n"
            "LISTA:\nResidencial unifamiliar\nResidencial multifamiliar\nComercial salas e lojas\nGalpão industrial\nCaso popular\nConjunto habitacional popular\n"
            "[logo abaixo: print das opções]"},
        {"fonte": FONTE + " · Título + Texto com imagem", "texto":
            "TÍTULO (Etapa): Resumo\n\n"
            "TEXTO: Por fim, ao final do cadastro, na tela “Resumo”, as informações repassadas serão demonstradas, sendo que ao final da página, teremos o campo “Finalizar” para concluir o processo.\n"
            "[logo abaixo: print do botão Finalizar]"},
    ],
}
EXEMPLOS_NOTAS = {
    "etapas": "Monte na ordem em que a pessoa vai fazer: um Título, depois blocos de Texto com imagem. Use Lista quando houver opções e Tabela quando houver campos de um sistema para explicar.",
}

# ------------------------------------------------------------------ pecas de JS/HTML novas
QUESTIONS_JS = r'''  var QUESTIONS = [
    {id:"q1", num:1, title:"Para que serve essa instrução? O que ela ensina a fazer?",
      dica:'Ex.: "Demonstrar o procedimento de abertura de CNO, para orientação das áreas envolvidas."',
      label:"IT · Seção 1 — Objetivo", type:"text", required:true},
    {id:"q2", num:2, title:"Para quem ela vale? Quais empresas ou áreas usam e para qual tarefa?",
      dica:'Ex.: "Aplicável para todo o Grupo Gadens, especificamente para abertura de Cadastro Nacional de Obras."',
      label:"IT · Seção 2 — Aplicação", type:"text"},
    {id:"q3", num:3, title:"Existe algum documento, norma ou site que a pessoa precisa seguir?",
      dica:'Um por linha. Ex.: "PO.GG.QLD.002 - Regras de elaboração de documento Grupo Gadens". Se não houver, escreva "não há".',
      label:"IT · Seção 3 — Referências", type:"text"},
    {id:"q4", num:4, title:"Quais siglas ou termos aparecem na instrução e precisam de explicação?",
      dica:"Uma sigla por linha. Ex.: CNO — Cadastro Nacional de Obras.",
      label:"IT · Seção 4 — Definições", type:"table", cellType:"textarea",
      columns:["Termo ou sigla","O que significa"], example:["CNO","Cadastro Nacional de Obras"]},
    {id:"q5", num:6, title:"Tem mais alguma coisa que quem for escrever a IT precisa saber?",
      dica:"Observações para a equipe de Processos (não vai para o Word). Ex.: quem pode ajudar, o que costuma dar errado, ideias de melhoria.",
      label:null, type:"text"}
  ];
'''

ETAPAS_HTML_JS = r'''
  // ---------- Etapas: blocos (título, texto com imagem, lista, tabela) ----------
  var etpList = null, etpSeq = 0;
  var BLOCK_LABEL = {titulo:"Título", texto:"Texto com imagem", lista:"Lista de opções", tabela:"Tabela Campo / Descrição"};
  var MIC_SVG = '<svg width="16" height="16" viewBox="0 0 24 24" fill="currentColor" aria-hidden="true"><path d="M12 14a3 3 0 0 0 3-3V5a3 3 0 0 0-6 0v6a3 3 0 0 0 3 3z"/><path d="M19 11a1 1 0 0 0-2 0 5 5 0 0 1-10 0 1 1 0 0 0-2 0 7 7 0 0 0 6 6.92V20H9a1 1 0 0 0 0 2h6a1 1 0 0 0 0-2h-2v-2.08A7 7 0 0 0 19 11z"/></svg>';

  qById.etapas = {id:"etapas", num:5, type:"etapas", columns:["Campo","Descrição"],
    title:"Passo a passo da atividade (Etapas)",
    dica:'Monte o passo a passo na ordem. Clique em "+ Título" para começar uma etapa (ex.: Acesso ao cadastro) e em "+ Texto com imagem" para explicar o que fazer. Logo abaixo do texto, clique em "+ Imagem ou print" (ou cole um print com Ctrl+V). Use "+ Lista" para opções e "+ Tabela" para explicar os campos de um sistema.'};

  function buildEtapasCard(){
    var wrap = document.createElement("div");
    wrap.className = "bloco";
    wrap.setAttribute("data-block", "etapas");
    wrap.innerHTML =
      '<div class="section-bar"><h3>5. Passo a passo da atividade (Etapas) <span class="req-mark">*</span> ' +
      '<button type="button" class="btn-ex" data-qid="etapas" aria-label="Ver exemplos de resposta" title="Ver exemplos de resposta">?</button></h3>' +
      '<span class="tag">[IT · Seção 5 — Etapas da Instrução de Trabalho]</span></div>' +
      '<p class="dica">' + escapeHtml(qById.etapas.dica) + '</p>' +
      '<div id="etapas-list"></div>' +
      '<div class="etp-add-bar">' +
      '<button type="button" class="etp-add" data-add="titulo">+ Título</button>' +
      '<button type="button" class="etp-add etp-add-main" data-add="texto">+ Texto com imagem</button>' +
      '<button type="button" class="etp-add" data-add="lista">+ Lista de opções</button>' +
      '<button type="button" class="etp-add" data-add="tabela">+ Tabela Campo / Descrição</button>' +
      '</div>' +
      '<p class="img-note">As imagens são reduzidas e ficam salvas só neste navegador, enquanto você preenche. Gere o Word antes de fechar a página.</p>' +
      '<p class="img-warning hidden">Não foi possível salvar todas as imagens automaticamente (espaço cheio). Gere o Word antes de fechar a página, ou remova imagens que não precisa.</p>';
    var anchor = container.querySelector('[data-block="q5"]');
    container.insertBefore(wrap, anchor);
    etpList = document.getElementById("etapas-list");
  }
  buildEtapasCard();

  function addTableLine(el, vals){
    var tbody = el.querySelector("tbody");
    var tr = document.createElement("tr");
    tr.className = "etp-row";
    tr.innerHTML = '<td><textarea rows="1" class="cell-textarea autogrow" spellcheck="true" aria-label="Campo"></textarea></td>' +
      '<td><textarea rows="1" class="cell-textarea autogrow" spellcheck="true" aria-label="Descrição"></textarea></td>' +
      '<td class="col-del"><button type="button" class="etp-del-line" title="Remover linha" aria-label="Remover linha">×</button></td>';
    var t = tr.querySelectorAll("textarea");
    t[0].value = (vals && vals[0]) || "";
    t[1].value = (vals && vals[1]) || "";
    tbody.appendChild(tr);
    t.forEach(autoGrowTextarea);
  }

  function addBlock(type, data){
    data = data || {};
    etpSeq++;
    var id = "b" + Date.now().toString(36) + etpSeq;
    var el = document.createElement("div");
    el.className = "etp-block etp-" + type;
    el.setAttribute("data-type", type);
    el.setAttribute("data-id", id);
    var tid = "etp_" + id;
    var html = '<div class="etp-head"><span class="etp-label"></span><span class="etp-ctrl">' +
      '<button type="button" class="etp-up" title="Subir" aria-label="Subir">↑</button>' +
      '<button type="button" class="etp-down" title="Descer" aria-label="Descer">↓</button>' +
      '<button type="button" class="etp-del" title="Remover este bloco" aria-label="Remover este bloco">×</button></span></div>';
    if(type === "titulo"){
      html += '<div class="etp-title-row"><select class="etp-level" aria-label="Nível do título">' +
        '<option value="1">Etapa</option><option value="2">Subetapa</option><option value="3">Item</option></select>' +
        '<input type="text" class="etp-title" placeholder="Ex.: Acesso ao cadastro" spellcheck="true"></div>';
    } else if(type === "tabela"){
      html += '<table class="dyn-table etp-table"><thead><tr><th>Campo</th><th>Descrição</th><th class="col-del"></th></tr></thead><tbody></tbody></table>' +
        '<button type="button" class="etp-add-line">+ adicionar linha</button>';
    } else {
      var ph = type === "lista" ? "Um item por linha. Ex.: Residencial unifamiliar" : "Escreva aqui o que a pessoa deve fazer nesta parte…";
      html += '<div class="text-field-wrap"><textarea id="' + tid + '" class="etp-text autogrow" spellcheck="true" placeholder="' + ph + '"></textarea>' +
        '<button type="button" class="btn-mic" data-qid="' + tid + '" title="Ditar por voz" aria-label="Ditar por voz">' + MIC_SVG + '</button></div>';
    }
    if(type !== "titulo"){
      html += '<div class="img-cell etp-imgs"><div class="img-thumbs"></div>' +
        '<label class="btn-img">+ Imagem ou print<input type="file" class="img-file" accept="image/*" multiple></label>' +
        '<span class="img-hint">ou cole um print aqui (Ctrl+V) ou arraste uma imagem</span></div>';
    }
    el.innerHTML = html;
    el._imgs = (data.imgs || []).slice();
    if(type === "titulo"){
      el.querySelector(".etp-level").value = String(data.level || 1);
      el.querySelector(".etp-title").value = data.title || "";
    } else if(type === "tabela"){
      var rows = (data.rows && data.rows.length) ? data.rows : [["", ""]];
      rows.forEach(function(r){ addTableLine(el, r); });
    } else {
      el.querySelector(".etp-text").value = data.text || "";
    }
    etpList.appendChild(el);
    var ta = el.querySelector(".etp-text");
    if(ta) autoGrowTextarea(ta);
    renderThumbs(el);
    renumberEtapas();
    setupMicButtons();
    return el;
  }

  function renumberEtapas(){
    if(!etpList) return;
    var c = [0, 0, 0];
    etpList.querySelectorAll(".etp-block").forEach(function(el){
      var lab = el.querySelector(".etp-label");
      var type = el.getAttribute("data-type");
      if(type === "titulo"){
        var L = parseInt(el.querySelector(".etp-level").value, 10) || 1;
        c[L - 1]++;
        for(var i = L; i < 3; i++) c[i] = 0;
        var parts = ["5"];
        for(var j = 0; j < L; j++) parts.push(c[j] || 1);
        lab.textContent = "Título " + parts.join(".") + ".";
      } else {
        lab.textContent = BLOCK_LABEL[type];
      }
    });
  }

  function collectEtapas(){
    var out = [];
    if(!etpList) return out;
    etpList.querySelectorAll(".etp-block").forEach(function(el){
      var type = el.getAttribute("data-type");
      var b = {type: type, imgs: (el._imgs || []).slice()};
      if(type === "titulo"){
        b.level = parseInt(el.querySelector(".etp-level").value, 10) || 1;
        b.title = el.querySelector(".etp-title").value;
      } else if(type === "tabela"){
        b.rows = [];
        el.querySelectorAll("tr.etp-row").forEach(function(tr){
          var t = tr.querySelectorAll("textarea");
          b.rows.push([t[0].value, t[1].value]);
        });
      } else {
        b.text = el.querySelector(".etp-text").value;
      }
      out.push(b);
    });
    return out;
  }

  function blockHasContent(b){
    if(b.imgs && b.imgs.length) return true;
    if(b.type === "titulo") return (b.title || "").trim() !== "";
    if(b.type === "tabela") return b.rows.some(function(r){ return r[0].trim() !== "" || r[1].trim() !== ""; });
    return (b.text || "").trim() !== "";
  }
  function etapasFilled(){ return collectEtapas().some(blockHasContent); }

  function restoreEtapas(list){
    etpList.innerHTML = "";
    if(list && list.length){
      list.forEach(function(b){ addBlock(b.type, b); });
    } else {
      addBlock("titulo", {level: 1});
      addBlock("texto", {});
    }
  }

  document.getElementById("questions-container").addEventListener("click", function(e){
    var addBtn = e.target.closest ? e.target.closest(".etp-add") : null;
    if(addBtn){
      var el = addBlock(addBtn.getAttribute("data-add"), addBtn.getAttribute("data-add") === "titulo" ? {level: 1} : {});
      var first = el.querySelector("input.etp-title, textarea");
      if(first) first.focus();
      scheduleSave(); updateProgress();
      return;
    }
    var blk = e.target.closest ? e.target.closest(".etp-block") : null;
    if(!blk) return;
    if(e.target.closest(".etp-up")){
      if(blk.previousElementSibling) blk.parentNode.insertBefore(blk, blk.previousElementSibling);
      renumberEtapas(); scheduleSave(); return;
    }
    if(e.target.closest(".etp-down")){
      if(blk.nextElementSibling) blk.parentNode.insertBefore(blk.nextElementSibling, blk);
      renumberEtapas(); scheduleSave(); return;
    }
    if(e.target.closest(".etp-del")){
      var b = collectEtapas()[Array.prototype.indexOf.call(etpList.children, blk)];
      if(b && blockHasContent(b) && !confirm("Remover este bloco e o que foi escrito nele?")) return;
      (blk._imgs || []).forEach(function(id){ delete IMG_STORE[id]; });
      blk.parentNode.removeChild(blk);
      if(!etpList.children.length) addBlock("texto", {});
      renumberEtapas(); scheduleSave(); updateProgress(); return;
    }
    if(e.target.closest(".etp-add-line")){
      addTableLine(blk, null); scheduleSave(); return;
    }
    var dl = e.target.closest(".etp-del-line");
    if(dl){
      var tr = dl.closest("tr");
      var rows = blk.querySelectorAll("tr.etp-row");
      if(rows.length > 1){ tr.parentNode.removeChild(tr); }
      else { tr.querySelectorAll("textarea").forEach(function(t){ t.value = ""; }); }
      scheduleSave(); updateProgress();
    }
  });

  document.getElementById("questions-container").addEventListener("change", function(e){
    if(e.target && e.target.classList && e.target.classList.contains("etp-level")){
      renumberEtapas(); scheduleSave();
    }
  });
'''

IMG_JS = r'''
  // ---------- Imagens (comuns a todos os blocos) ----------
  var IMG_KEY = STORAGE_KEY + "_imgs";
  var IMG_STORE = {};
  var imgSeq = 0;
  function newImgId(){ imgSeq++; return "i" + Date.now().toString(36) + imgSeq; }

  function loadImgStore(){
    try{ var raw = localStorage.getItem(IMG_KEY); if(raw) IMG_STORE = JSON.parse(raw) || {}; }catch(err){ IMG_STORE = {}; }
  }

  function readImageFile(file, cb){
    var fr = new FileReader();
    fr.onload = function(){
      var img = new Image();
      img.onload = function(){
        var max = 1400;
        var s = Math.min(1, max / Math.max(img.width, img.height));
        var w = Math.max(1, Math.round(img.width * s)), h = Math.max(1, Math.round(img.height * s));
        var c = document.createElement("canvas");
        c.width = w; c.height = h;
        var ctx = c.getContext("2d");
        ctx.fillStyle = "#fff"; ctx.fillRect(0, 0, w, h);
        ctx.drawImage(img, 0, 0, w, h);
        cb({data: c.toDataURL("image/jpeg", 0.8), w: w, h: h});
      };
      img.onerror = function(){ cb(null); };
      img.src = fr.result;
    };
    fr.onerror = function(){ cb(null); };
    fr.readAsDataURL(file);
  }

  function renderThumbs(el){
    var box = el.querySelector(".img-thumbs");
    if(!box) return;
    box.innerHTML = "";
    (el._imgs || []).forEach(function(id){
      var im = IMG_STORE[id];
      if(!im) return;
      var wrap = document.createElement("span");
      wrap.className = "img-thumb";
      var t = document.createElement("img");
      t.src = im.data; t.alt = "Imagem anexada";
      var rm = document.createElement("button");
      rm.type = "button"; rm.className = "btn-img-remove"; rm.setAttribute("data-img", id);
      rm.title = "Remover imagem"; rm.setAttribute("aria-label", "Remover imagem"); rm.textContent = "×";
      wrap.appendChild(t); wrap.appendChild(rm);
      box.appendChild(wrap);
    });
  }

  function attachImages(el, files){
    var list = Array.prototype.slice.call(files || []).filter(function(f){ return f && /^image\//.test(f.type); });
    list.forEach(function(f){
      readImageFile(f, function(im){
        if(!im){ alert("Não foi possível ler essa imagem. Tente outro arquivo (JPG ou PNG)."); return; }
        var id = newImgId();
        IMG_STORE[id] = im;
        el._imgs = (el._imgs || []).concat([id]);
        renderThumbs(el);
        scheduleSave(); updateProgress();
      });
    });
  }

  container.addEventListener("change", function(e){
    if(e.target && e.target.classList && e.target.classList.contains("img-file")){
      var el = e.target.closest(".etp-block");
      if(el) attachImages(el, e.target.files);
      e.target.value = "";
    }
  });

  container.addEventListener("click", function(e){
    var rm = e.target.closest ? e.target.closest(".btn-img-remove") : null;
    if(!rm) return;
    var el = rm.closest(".etp-block");
    var id = rm.getAttribute("data-img");
    el._imgs = (el._imgs || []).filter(function(x){ return x !== id; });
    delete IMG_STORE[id];
    renderThumbs(el);
    scheduleSave(); updateProgress();
  });

  container.addEventListener("dragover", function(e){
    if(e.target.closest && e.target.closest(".etp-imgs")) e.preventDefault();
  });
  container.addEventListener("drop", function(e){
    var zone = e.target.closest ? e.target.closest(".etp-imgs") : null;
    if(!zone) return;
    e.preventDefault();
    attachImages(zone.closest(".etp-block"), e.dataTransfer.files);
  });

  document.addEventListener("paste", function(e){
    var a = document.activeElement;
    var el = (a && a.closest) ? a.closest(".etp-block") : null;
    if(!el || !el.querySelector(".img-file")) return;
    var cd = e.clipboardData;
    if(!cd || !cd.items) return;
    var files = [], hasText = false;
    for(var i = 0; i < cd.items.length; i++){
      var it = cd.items[i];
      if(it.kind === "file" && /^image\//.test(it.type)){ var f = it.getAsFile(); if(f) files.push(f); }
      if(it.kind === "string" && it.type === "text/plain") hasText = true;
    }
    if(files.length && !hasText){
      e.preventDefault();
      attachImages(el, files);
    }
  });

  function showImgWarning(on){
    document.querySelectorAll(".img-warning").forEach(function(el){ el.classList.toggle("hidden", !on); });
  }
'''

ENGINE_JS = r'''function buildAplicacaoText(){
    var empresa = document.getElementById("id_empresa").value.trim();
    var onde = document.getElementById("q2").value.trim();
    if(onde) return onde;
    return empresa ? "Empresa: " + empresa + "." : "";
  }

  var docMedia = [];
  var XML_A = 'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"';
  var XML_PIC = 'xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"';

  function substituteTokenAll(xml, token, value){
    var re = new RegExp('<w:t[^>]*>\\{\\{' + token + '\\}\\}</w:t>', "g");
    var safe = '<w:t xml:space="preserve">' + escapeHtml(value) + '</w:t>';
    return xml.replace(re, function(){ return safe; });
  }

  function substituteBlock(xml, token, innerXml){
    var found = findEnclosingBlock(xml, "{{" + token + "}}", /<w:p(?=[ >])/, "</w:p>");
    if(!found) return xml;
    return xml.slice(0, found.start) + innerXml + xml.slice(found.end);
  }

  // Lista: cada item termina em ponto e virgula; o ultimo, em ponto.
  function normalizeList(items){
    var clean = items.map(function(s){ return String(s).trim().replace(/[.;,\s]+$/, ""); }).filter(function(s){ return s !== ""; });
    return clean.map(function(s, i){ return s + (i === clean.length - 1 ? "." : ";"); });
  }

  function capFirstText(s){ s = String(s || "").trim(); return s ? s.charAt(0).toUpperCase() + s.slice(1) : s; }
  function endPunct(s){ s = String(s || "").trim(); return s === "" ? s : (/[.;:!?]$/.test(s) ? s : s + "."); }
  function linesOf(s){ return String(s || "").split(/\r?\n/).map(function(l){ return l.trim(); }).filter(function(l){ return l !== ""; }); }

  function xr(o){
    o = o || {};
    return '<w:rPr><w:rFonts w:ascii="Arial" w:eastAsia="Calibri" w:hAnsi="Arial" w:cs="Arial"/>' +
      (o.b ? '<w:b/><w:bCs/>' : '') + (o.caps ? '<w:caps/>' : '') +
      '<w:color w:val="000000" w:themeColor="text1"/><w:sz w:val="20"/><w:szCs w:val="20"/><w:lang w:eastAsia="pt-BR"/></w:rPr>';
  }
  function xRun(text, o){
    return '<w:r>' + xr(o) + '<w:t xml:space="preserve">' + escapeHtml(text) + '</w:t></w:r>';
  }
  function xHeading(level, text){
    return '<w:p><w:pPr><w:pStyle w:val="PargrafodaLista"/><w:keepNext/><w:numPr><w:ilvl w:val="' + level + '"/><w:numId w:val="1"/></w:numPr>' +
      '<w:spacing w:after="0" w:line="360" w:lineRule="auto"/><w:jc w:val="both"/>' + xr({b:1, caps:1}) + '</w:pPr>' +
      xRun(text, {b:1, caps:1}) + '</w:p>';
  }
  function xBody(runsXml){
    return '<w:p><w:pPr><w:spacing w:after="0" w:line="360" w:lineRule="auto"/><w:ind w:firstLine="851"/><w:jc w:val="both"/>' + xr({}) + '</w:pPr>' + runsXml + '</w:p>';
  }
  function xBullet(text){
    return '<w:p><w:pPr><w:pStyle w:val="PargrafodaLista"/><w:numPr><w:ilvl w:val="0"/><w:numId w:val="2"/></w:numPr>' +
      '<w:spacing w:after="0" w:line="360" w:lineRule="auto"/><w:jc w:val="both"/>' + xr({}) + '</w:pPr>' + xRun(text) + '</w:p>';
  }
  function xTable(rows){
    function cell(text, head){
      return '<w:tc><w:tcPr><w:tcW w:w="4513" w:type="dxa"/></w:tcPr><w:p><w:pPr><w:spacing w:after="0" w:line="360" w:lineRule="auto"/>' +
        '<w:jc w:val="' + (head ? 'center' : 'both') + '"/>' + xr({b: head}) + '</w:pPr>' + xRun(text, {b: head}) + '</w:p></w:tc>';
    }
    var x = '<w:tbl><w:tblPr><w:tblStyle w:val="Tabelacomgrade"/><w:tblW w:w="9026" w:type="dxa"/>' +
      '<w:tblLook w:val="04A0" w:firstRow="1" w:lastRow="0" w:firstColumn="1" w:lastColumn="0" w:noHBand="0" w:noVBand="1"/></w:tblPr>' +
      '<w:tblGrid><w:gridCol w:w="4513"/><w:gridCol w:w="4513"/></w:tblGrid>' +
      '<w:tr><w:trPr><w:tblHeader/></w:trPr>' + cell("CAMPO", true) + cell("DESCRIÇÃO", true) + '</w:tr>';
    rows.forEach(function(r){
      x += '<w:tr><w:trPr><w:cantSplit/></w:trPr>' + cell(r[0].trim(), false) + cell(r[1].trim(), false) + '</w:tr>';
    });
    return x + '</w:tbl><w:p><w:pPr><w:spacing w:after="0"/>' + xr({}) + '</w:pPr></w:p>';
  }
  function dataUrlToBytes(u){
    var bin = atob(String(u).split(",")[1] || "");
    var a = new Uint8Array(bin.length);
    for(var i = 0; i < bin.length; i++){ a[i] = bin.charCodeAt(i); }
    return a;
  }
  function xImage(im){
    var maxW = 4680000, maxH = 4320000;            // 13 cm de largura, no maximo 12 cm de altura
    var cx = maxW, cy = Math.round(maxW * im.h / im.w);
    if(cy > maxH){ cy = maxH; cx = Math.round(maxH * im.w / im.h); }
    var n = docMedia.length + 1;
    var rid = "rIdItImg" + n, fname = "it_img" + n + ".jpg";
    docMedia.push({rid: rid, name: fname, bytes: dataUrlToBytes(im.data)});
    return '<w:p><w:pPr><w:spacing w:before="240" w:after="0" w:line="360" w:lineRule="auto"/><w:jc w:val="center"/>' + xr({}) + '</w:pPr>' +
      '<w:r><w:rPr><w:noProof/></w:rPr><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0">' +
      '<wp:extent cx="' + cx + '" cy="' + cy + '"/><wp:effectExtent l="0" t="0" r="0" b="0"/>' +
      '<wp:docPr id="' + (5000 + n) + '" name="Imagem ' + n + '"/>' +
      '<wp:cNvGraphicFramePr><a:graphicFrameLocks ' + XML_A + ' noChangeAspect="1"/></wp:cNvGraphicFramePr>' +
      '<a:graphic ' + XML_A + '><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">' +
      '<pic:pic ' + XML_PIC + '><pic:nvPicPr><pic:cNvPr id="0" name="' + fname + '"/><pic:cNvPicPr/></pic:nvPicPr>' +
      '<pic:blipFill><a:blip r:embed="' + rid + '"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>' +
      '<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="' + cx + '" cy="' + cy + '"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>' +
      '</pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>';
  }

  function buildEtapasXml(){
    var out = [];
    function images(b){ (b.imgs || []).forEach(function(id){ if(IMG_STORE[id]) out.push(xImage(IMG_STORE[id])); }); }
    collectEtapas().forEach(function(b){
      if(b.type === "titulo"){
        if((b.title || "").trim() !== "") out.push(xHeading(b.level, b.title.trim()));
      } else if(b.type === "texto"){
        linesOf(b.text).forEach(function(l){ out.push(xBody(xRun(endPunct(capFirstText(l))))); });
        images(b);
      } else if(b.type === "lista"){
        normalizeList(linesOf(b.text)).forEach(function(t){ out.push(xBullet(capFirstText(t))); });
        images(b);
      } else if(b.type === "tabela"){
        var rows = b.rows.filter(function(r){ return r[0].trim() !== "" || r[1].trim() !== ""; });
        if(rows.length) out.push(xTable(rows));
        images(b);
      }
    });
    if(!out.length) out.push(xBody(xRun("(não informado)")));
    return out.join("");
  }

  function buildFilledDocumentXml(xml){
    docMedia = [];
    var tarefa = document.getElementById("id_tarefa").value.trim();
    xml = substituteTokenAll(xml, "TITULO", tarefa ? tarefa.toUpperCase() : "TÍTULO DA INSTRUÇÃO DE TRABALHO");
    xml = substituteSimpleToken(xml, "OBJETIVO", endPunct(capFirstText(document.getElementById("q1").value)));
    xml = substituteSimpleToken(xml, "APLICACAO", endPunct(capFirstText(buildAplicacaoText())));
    var refs = getTextLines("q3");
    xml = substituteRepeatingParagraph(xml, "REFERENCIA_ITEM", refs.length ? normalizeList(refs) : ["(não informado)"]);
    xml = substituteRepeatingRow(xml, ["SIGLA", "SIGLA_DEF"], getTableRowsFiltered("q4"));
    xml = substituteBlock(xml, "ETAPAS", buildEtapasXml());
    return xml;
  }
'''

HANDLER_JS = r'''      var zip = fflate.unzipSync(templateBytes);
      var docXmlBytes = zip["word/document.xml"];
      if(!docXmlBytes) throw new Error("word/document.xml não encontrado no template.");
      var enc = new TextEncoder(), dec = new TextDecoder("utf-8");
      var xmlText = dec.decode(docXmlBytes);
      var filledXml = buildFilledDocumentXml(xmlText);
      zip["word/document.xml"] = enc.encode(filledXml);
      if(docMedia.length){
        var relsXml = dec.decode(zip["word/_rels/document.xml.rels"]);
        var add = docMedia.map(function(m){
          return '<Relationship Id="' + m.rid + '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/' + m.name + '"/>';
        }).join("");
        zip["word/_rels/document.xml.rels"] = enc.encode(relsXml.replace("</Relationships>", add + "</Relationships>"));
        var ct = dec.decode(zip["[Content_Types].xml"]);
        if(!/Extension="jpe?g"/i.test(ct)){
          ct = ct.replace(/(<Types[^>]*>)/, '$1<Default Extension="jpg" ContentType="image/jpeg"/>');
        }
        zip["[Content_Types].xml"] = enc.encode(ct);
        docMedia.forEach(function(m){ zip["word/media/" + m.name] = [m.bytes, {level: 0}]; });
      }
      var outBytes = fflate.zipSync(zip, {level: 6});
      var blob = new Blob([outBytes], {type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document"});
      var tarefa = sanitizeFileName(document.getElementById("id_tarefa").value);
      var data = document.getElementById("id_data").value || todayISO();
      var filename = "IT - " + tarefa + " - " + data + ".docx";
      var url = URL.createObjectURL(blob);
      var a = document.createElement("a");
      a.href = url;
      a.download = filename;
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
      setTimeout(function(){ URL.revokeObjectURL(url); }, 1000);
      statusEl.textContent = "Word gerado: " + filename + ". Abra e revise: confira a ordem das etapas e as imagens. Código do documento, aprovações e controle de emissão continuam sendo preenchidos pela equipe de Processos.";
      statusEl.classList.remove("hidden");
    }catch(err){'''

CSS_JS = r'''
  /* ---- Etapas em blocos (IT) ---- */
  .etp-block{ border:1px solid var(--borda); border-radius:6px; padding:10px 12px; margin:10px 0; background:#fff; }
  .etp-block.etp-titulo{ border-left:4px solid var(--vermelho); background:#fafafa; }
  :root[data-theme="dark"] .etp-block{ background:#242424; }
  :root[data-theme="dark"] .etp-block.etp-titulo{ background:#2a2222; }
  .etp-head{ display:flex; justify-content:space-between; align-items:center; margin-bottom:6px; }
  .etp-label{ font-weight:600; font-size:.85rem; color:var(--cinza); }
  :root[data-theme="dark"] .etp-label{ color:#cfcfcf; }
  .etp-ctrl button{ border:1px solid var(--borda); background:transparent; color:inherit; border-radius:4px; width:26px; height:26px; cursor:pointer; margin-left:4px; font-size:14px; line-height:1; }
  .etp-ctrl button:hover{ background:rgba(0,0,0,.06); }
  .etp-title-row{ display:flex; gap:8px; }
  .etp-title-row select{ flex:0 0 130px; }
  .etp-title-row input{ flex:1; min-width:0; }
  .etp-title-row select, .etp-title-row input{ padding:8px; border:1px solid var(--borda); border-radius:4px; font:inherit; background:#fff; color:inherit; }
  :root[data-theme="dark"] .etp-title-row select, :root[data-theme="dark"] .etp-title-row input{ background:#2e2e2e; color:#e8e8e8; }
  .etp-add-bar{ display:flex; flex-wrap:wrap; gap:8px; margin:10px 0 6px; }
  .etp-add, .etp-add-line{ background:#fff; border:1px dashed var(--cinza); color:var(--cinza); border-radius:4px; padding:7px 12px; font-size:.85rem; cursor:pointer; }
  .etp-add-main{ border-style:solid; border-color:var(--vermelho); color:var(--vermelho); font-weight:600; }
  .etp-add:hover, .etp-add-line:hover{ background:#f2f2f2; }
  :root[data-theme="dark"] .etp-add, :root[data-theme="dark"] .etp-add-line{ background:#242424; color:#e8e8e8; }
  .etp-del-line{ border:0; background:transparent; color:var(--vermelho); font-size:18px; cursor:pointer; }
  .etp-table{ margin-bottom:6px; }
  .img-cell{ display:flex; flex-wrap:wrap; gap:8px; align-items:center; margin-top:8px; padding:8px; border:1px dashed var(--borda); border-radius:4px; }
  .img-thumbs{ display:flex; flex-wrap:wrap; gap:6px; }
  .img-thumb{ position:relative; display:inline-block; border:1px solid var(--borda); border-radius:4px; overflow:hidden; line-height:0; }
  .img-thumb img{ width:84px; height:62px; object-fit:cover; display:block; }
  .btn-img-remove{ position:absolute; top:0; right:0; width:18px; height:18px; border:0; background:rgba(0,0,0,.65); color:#fff; font-size:13px; line-height:18px; cursor:pointer; padding:0; }
  .btn-img{ display:inline-block; border:1px solid var(--cinza); color:var(--cinza); border-radius:4px; padding:6px 12px; font-size:.82rem; cursor:pointer; background:#fff; }
  .btn-img:hover{ background:#f2f2f2; }
  .btn-img input{ display:none; }
  :root[data-theme="dark"] .btn-img{ background:#242424; color:#e8e8e8; }
  .img-hint{ font-size:.78rem; color:var(--dica); }
  .img-note{ font-size:.82rem; color:var(--dica); margin:4px 0 8px; }
  .img-warning{ font-size:.82rem; color:#b00020; margin:0 0 8px; }
  .img-warning.hidden{ display:none; }
  :root[data-theme="dark"] .img-warning{ color:#ff9d9d; }
'''


def main(pop_html, modelo_docx, out_html):
    s = open(pop_html, encoding="utf-8").read()

    m = re.search(r'(var TEMPLATE_DOCX_B64\s*=\s*")([^"]*)(")', s)
    if not m:
        raise SystemExit("Nao achei TEMPLATE_DOCX_B64")
    s = s[:m.start(2)] + "<<B64>>" + s[m.end(2):]

    # --- identidade
    s = rep_re(s, r"<title>.*?</title>", "<title>Levantamento de IT — Instrução de Trabalho | Grupo Gadens</title>")
    s = rep(s, "<h1>G-Docs Formulário de Levantamento de Processo</h1>", "<h1>G-Docs Formulário de Levantamento de IT</h1>")
    s = rep_re(s, r"<p>Conte como o seu trabalho é feito.*?</p>",
               "<p>Conte como a atividade é feita, passo a passo, com prints — a equipe de Processos transforma em IT (Instrução de Trabalho).</p>")
    s = rep(s, 'var STORAGE_KEY = "gadens_levantamento_pop_v1";', 'var STORAGE_KEY = "gadens_levantamento_it_v2";')
    s = rep(s, 'var THEME_KEY = "gadens_pop_theme";', 'var THEME_KEY = "gadens_it_theme";')

    s = rep_re(s, r'<div class="card instrucoes">.*?</ul>\s*</div>', '''<div class="card instrucoes">
    <h2>Antes de começar</h2>
    <p style="font-size:.92rem; margin:0 0 8px;">Este formulário ajuda você a escrever a IT (Instrução de Trabalho) — o documento oficial que ensina, passo a passo, como fazer uma atividade. Você escreve e anexa os prints; o formulário monta o Word no modelo oficial.</p>
    <ul>
      <li>Nas partes 1 a 4, responda com poucas palavras. O botão "?" de cada pergunta mostra um exemplo de uma IT pronta (Cadastro de CNO).</li>
      <li>Na parte 5 (Etapas), escreva o passo e, logo abaixo, anexe o print ou a imagem. Clique em "+ Imagem ou print", cole um print com Ctrl+V ou arraste a imagem.</li>
      <li>Use "+ Título" para começar uma nova etapa, "+ Lista" para opções e "+ Tabela" para explicar os campos de um sistema.</li>
      <li>Escreva como se estivesse ensinando um colega novo. Não precisa "escrever bonito".</li>
      <li>Suas respostas e imagens ficam salvas neste computador enquanto você preenche. Gere o Word antes de fechar a página.</li>
    </ul>
  </div>''')
    s = rep(s, 'Nome da tarefa ou rotina descrita\n          <span class="help-icon" tabindex="0">?<span class="tooltip-text">Título do documento</span></span>',
            'Nome da atividade\n          <span class="help-icon" tabindex="0">?<span class="tooltip-text">Será o título da IT</span></span>')
    s = rep(s, 'placeholder="Use o nome pelo qual ela é conhecida no dia a dia"', 'placeholder="Ex.: Cadastro de CNO"')
    s = rep_re(s, r'<div class="card terminou">.*?</div>', '''<div class="card terminou">
    <h2>Terminou?</h2>
    <p>Releia rapidinho, confira se os prints estão nas etapas certas e clique em "Enviar Respostas (gerar Word)" — isso monta o rascunho da IT já no modelo oficial. Abra o Word, dê uma conferida e envie para a equipe de Processos.</p>
    <p>Código do documento, aprovações e controle de emissão são preenchidos pela equipe de Processos/Qualidade — você não precisa se preocupar com essa parte.</p>
  </div>''')
    s = rep(s, "Grupo Gadens — Levantamento de Processo (base para POP)", "Grupo Gadens — Levantamento de IT (base para Instrução de Trabalho)")

    # --- perguntas e card de etapas
    s = rep_re(s, r"  var QUESTIONS = \[.*?\n  \];\n", QUESTIONS_JS)
    s = rep(s, "  // ---------- Linhas de tabela ----------", IMG_JS + ETAPAS_HTML_JS + "\n  // ---------- Linhas de tabela ----------")

    # --- progresso, salvar, carregar
    s = rep(s, "    var total = QUESTIONS.length + 1;\n    var filled = 0;\n    if(isIdentificacaoFilled()) filled++;",
               "    var total = QUESTIONS.length + 2;\n    var filled = 0;\n    if(isIdentificacaoFilled()) filled++;\n    if(etapasFilled()) filled++;")
    s = rep(s, "      respostas: {}\n    };\n    QUESTIONS.forEach(function(q){\n      if(q.type === \"text\"){\n        state.respostas[q.id] = document.getElementById(q.id).value;\n      } else {\n        state.respostas[q.id] = collectTableRows(q.id);\n      }\n    });\n    try{\n      localStorage.setItem(STORAGE_KEY, JSON.stringify(state));\n    }catch(err){ /* localStorage indisponível: segue sem autosave */ }",
               "      respostas: {},\n      etapas: collectEtapas()\n    };\n    var usadas = {};\n    state.etapas.forEach(function(b){\n      b.imgs = (b.imgs || []).filter(function(id){ return IMG_STORE[id]; });\n      b.imgs.forEach(function(id){ usadas[id] = IMG_STORE[id]; });\n    });\n    QUESTIONS.forEach(function(q){\n      if(q.type === \"text\"){\n        state.respostas[q.id] = document.getElementById(q.id).value;\n      } else {\n        state.respostas[q.id] = collectTableRows(q.id);\n      }\n    });\n    try{\n      localStorage.setItem(STORAGE_KEY, JSON.stringify(state));\n    }catch(err){ /* localStorage indisponível: segue sem autosave */ }\n    try{\n      localStorage.setItem(IMG_KEY, JSON.stringify(usadas));\n      showImgWarning(false);\n    }catch(err){ showImgWarning(true); }")
    s = rep(s, "  function loadState(){\n    var raw = null;", "  function loadState(){\n    loadImgStore();\n    var raw = null;")
    s = rep(s, "        if(saved && saved.length){\n          saved.forEach(function(rowVals){ addRow(q.id, rowVals); });\n        } else {\n          addRow(q.id, null);\n        }\n      }\n    });\n  }",
               "        if(saved && saved.length){\n          saved.forEach(function(rowVals){ addRow(q.id, rowVals); });\n        } else {\n          addRow(q.id, null);\n        }\n      }\n    });\n    restoreEtapas(state && state.etapas ? state.etapas : null);\n  }")

    # --- obrigatorias
    s = rep(s, '{label:"Nome da tarefa ou rotina", el:', '{label:"Nome da atividade", el:')
    s = rep(s, '{label:"Pergunta 1 — Objetivo do trabalho", el: q1}', '{label:"Pergunta 1 — Objetivo da instrução", el: q1}')
    s = rep(s, 'if(!isQuestionFilled(qById["q9"])) missing.push({label:"Pergunta 9 — Passo a passo (pelo menos 1 linha preenchida)", el: document.getElementById("wrap_q9")});',
               'if(!etapasFilled()) missing.push({label:"Pergunta 5 — Passo a passo (escreva pelo menos uma etapa)", el: document.getElementById("etapas-list")});')

    # --- exportacao .txt
    s = rep(s, '"Sem nome da tarefa"', '"Sem nome da atividade"')
    s = rep(s, 'lines.push("LEVANTAMENTO DE PROCESSO — BASE PARA POP");', 'lines.push("LEVANTAMENTO DE IT — BASE PARA INSTRUÇÃO DE TRABALHO");')
    s = rep(s, 'lines.push("Nome da tarefa ou rotina descrita: "', 'lines.push("Nome da atividade: "')

    # --- motor do Word
    s = rep_re(s, r"function buildAplicacaoText\(\)\{.*?\n  \}\n", "@@ENGINE@@\n")
    s = rep_re(s, r"  function buildFilledDocumentXml\(xml\)\{.*?\n    return xml;\n  \}\n", "")
    s = s.replace("@@ENGINE@@\n", ENGINE_JS)
    s = rep_re(s, r"      var zip = fflate\.unzipSync\(templateBytes\);.*?\n    \}catch\(err\)\{", HANDLER_JS)
    s = rep(s, "      try{ localStorage.removeItem(STORAGE_KEY); }catch(err){}", "      try{ localStorage.removeItem(STORAGE_KEY); localStorage.removeItem(IMG_KEY); }catch(err){}")

    # --- microfone: aceitar botoes criados depois (blocos novos)
    s = rep(s, "    micButtons.forEach(function(btn){\n      btn.addEventListener(\"click\", function(){\n        if(btn.classList.contains(\"listening\")){",
               "    micButtons.forEach(function(btn){\n      if(btn.getAttribute(\"data-bound\")) return;\n      btn.setAttribute(\"data-bound\", \"1\");\n      btn.addEventListener(\"click\", function(){\n        if(btn.classList.contains(\"listening\")){")

    # --- modo por voz
    s = rep(s, 'title:"Para qual empresa do grupo é esse processo?"', 'title:"Para qual empresa do grupo é essa atividade?"')
    s = rep(s, 'title:"Qual é o nome da tarefa ou rotina que você vai descrever?", dica:"Use o nome pelo qual ela é conhecida no dia a dia."',
               'title:"Qual é o nome da atividade que você vai descrever?", dica:"Esse nome será o título da IT."')
    s = rep(s, "Terminamos. Revise as respostas no formulário e depois gere o Word.",
               "Terminamos. Agora escreva as etapas e anexe os prints no formulário, e depois gere o Word.")
    s = rep(s, '<span>No Modo por voz, o formulário faz as perguntas em voz alta e você responde falando. Funciona no Chrome e no Edge.</span>',
               '<span>No Modo por voz, o formulário faz as perguntas 1 a 4 em voz alta e você responde falando. As etapas e os prints você preenche na tela (cada caixa de texto tem um microfone). Funciona no Chrome e no Edge.</span>')

    # --- exemplos
    s = rep_re(s, r"  var EXEMPLOS = \{.*?\n  var EXEMPLOS_NOTAS = \{.*?\};\n",
               "  var EXEMPLOS = " + json.dumps(EXEMPLOS, ensure_ascii=False) + ";\n  var EXEMPLOS_NOTAS = " + json.dumps(EXEMPLOS_NOTAS, ensure_ascii=False) + ";\n")
    s = rep(s, "// Cada exemplo é um trecho de POP já aprovado", "// Cada exemplo é um trecho de IT já elaborada")
    s = rep(s, '(lista.length > 1 ? "Exemplos de POPs já publicados" : "Exemplo de POP já publicado")',
               '(lista.length > 1 ? "Exemplos de ITs já elaboradas" : "Exemplo de IT já elaborada")')
    s = rep(s, ".map(function(c){ return c.trim(); }).join(\" — \");", ".map(function(c){ return c.trim(); }).join(\" – \");")

    s = rep(s, "</style>", CSS_JS + "</style>")
    b64 = base64.b64encode(open(modelo_docx, "rb").read()).decode("ascii")
    s = s.replace("<<B64>>", b64)
    open(out_html, "w", encoding="utf-8").write(s)
    print("ok ->", out_html, len(s), "bytes")


if __name__ == "__main__":
    if len(sys.argv) != 4:
        raise SystemExit(__doc__)
    main(*sys.argv[1:])
