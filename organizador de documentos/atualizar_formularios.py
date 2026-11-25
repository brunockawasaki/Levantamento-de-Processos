"""Incorpora os formularios Elaborar POP, IT, Regimento, Ata e Politica ao painel G-Docs.

uso: python atualizar_formularios.py <formulario_POP.html> <formulario_IT.html> <formulario_Regimento.html> <formulario_Ata.html> <formulario_Politica.html> <painel-organizacao-documentos.html>
Regera CSS, marcacao e JS dos cinco formularios, sempre a partir dos arquivos de origem.
O painel e lido e gravado sempre com final de linha LF (no Windows, open() normal gravaria CRLF).
"""
import re, sys

POP_SRC, IT_SRC, RG_SRC, AT_SRC, PL_SRC, DST = sys.argv[1], sys.argv[2], sys.argv[3], sys.argv[4], sys.argv[5], sys.argv[6]
s = open(DST, encoding='utf-8', newline='').read().replace('\r\n', '\n')


def must_replace(text, old, new):
    assert text.count(old) == 1, (old[:80], text.count(old))
    return text.replace(old, new)


def split_rules(text):
    out = []; i = 0; n = len(text)
    while i < n:
        while i < n and text[i].isspace(): i += 1
        if i >= n: break
        j = text.index('{', i); head = text[i:j].strip(); depth = 1; k = j + 1
        while depth and k < n:
            if text[k] == '{': depth += 1
            elif text[k] == '}': depth -= 1
            k += 1
        out.append((head, text[j + 1:k - 1])); i = k
    return out


DARK = ':root[data-theme="dark"]'

CFG = {
    'pop': dict(src=POP_SRC, root='elab-root', view='elaborar', banner='elab-pending-banner',
                css_marker='/* ===== Elaborar POP (formulario de levantamento) ===== */',
                h1_old='<h1>G-Docs Formulário de Levantamento de Processo</h1>',
                h1_new='<h1>Elaborar POP · Levantamento de Processo</h1>',
                rename_ids=False, idprefix='', toast_css=True),
    'it': dict(src=IT_SRC, root='elit-root', view='elaborar-it', banner='elit-pending-banner',
               css_marker='/* ===== Elaborar IT (formulario de levantamento de IT) ===== */',
               h1_old='<h1>G-Docs Formulário de Levantamento de IT</h1>',
               h1_new='<h1>Elaborar IT · Levantamento de IT</h1>',
               rename_ids=True, idprefix='it_', toast_css=False),
    'rg': dict(src=RG_SRC, root='elrg-root', view='elaborar-regimento', banner='elrg-pending-banner',
               css_marker='/* ===== Elaborar Regimento (formulario de levantamento de Regimento) ===== */',
               h1_old='<h1>G-Docs Formulário de Levantamento de Regimento</h1>',
               h1_new='<h1>Elaborar Regimento · Levantamento de Regimento</h1>',
               rename_ids=True, idprefix='rg_', toast_css=False),
    'at': dict(src=AT_SRC, root='elat-root', view='elaborar-ata', banner='elat-pending-banner',
               css_marker='/* ===== Elaborar Ata (formulario de ata de reuniao) ===== */',
               h1_old='<h1>G-Docs Formulário de Ata de Reunião</h1>',
               h1_new='<h1>Elaborar Ata · Ata de Reunião</h1>',
               rename_ids=True, idprefix='at_', toast_css=False),
    'pl': dict(src=PL_SRC, root='elpl-root', view='elaborar-politica', banner='elpl-pending-banner',
               css_marker='/* ===== Elaborar Política (formulario de politica de departamento) ===== */',
               h1_old='<h1>G-Docs Formulário de Política de Departamento</h1>',
               h1_new='<h1>Elaborar Política · Política de Departamento</h1>',
               rename_ids=True, idprefix='pl_', toast_css=False),
}


def build_css(f, cfg):
    ROOT = '#' + cfg['root']
    fcss = re.search(r'<style[^>]*>(.*?)</style>', f, re.S).group(1)
    fcss = re.sub(r'/\*.*?\*/', '', fcss, flags=re.S)

    def scope_sel(sel):
        sel = sel.strip().replace('#pending-banner', '#' + cfg['banner'])
        if '.btn-theme-toggle' in sel or '.help-fab' in sel: return None
        dark = sel.startswith(DARK)
        rest = (sel[len(DARK):] if dark else sel).strip()
        if rest in ('', ':root', 'body'): core = ROOT
        elif rest == '*': core = ROOT + ' *'
        elif rest.startswith('body '): core = ROOT + ' ' + rest[5:]
        else: core = ROOT + ' ' + rest
        return dark, core

    def emit(head, body, op, oa, om):
        sels = [scope_sel(x) for x in head.split(',')]
        sels = [x for x in sels if x]
        if not sels: return
        body = body.replace('padding-top:34px;', '')
        for dk in (False, True):
            group = [c for d, c in sels if d == dk]
            if not group: continue
            if not dk:
                op.append(','.join(group) + '{' + body + '}')
            else:
                oa.append(','.join(DARK + ' ' + c for c in group) + '{' + body + '}')
                om.append(','.join(':root:not([data-theme="light"]) ' + c for c in group) + '{' + body + '}')

    plain = []; dattr = []; dmedia = []
    for head, body in split_rules(fcss):
        if head.startswith('@keyframes'):
            plain.append(head + '{' + body + '}'); continue
        if head.startswith('@media'):
            ip = []; ia = []; im = []
            for h2, b2 in split_rules(body): emit(h2, b2, ip, ia, im)
            if ip: plain.append(head + '{' + ''.join(ip) + '}')
            continue
        emit(head, body, plain, dattr, dmedia)

    scoped = ''.join(plain)
    scoped = re.sub(r'position:fixed;\s*top:0;\s*left:0;\s*right:0;', 'position:sticky;top:var(--elab-top,0);', scoped)
    assert 'position:sticky;top:var(--elab-top,0)' in scoped
    scoped = scoped.replace('z-index:1000;', 'z-index:20;')
    dark_css = ''.join(dattr) + '@media (prefers-color-scheme: dark){' + ''.join(dmedia) + '}'
    extra = ''
    if cfg['toast_css']:
        extra += ('#elab-toast{position:fixed;left:50%;bottom:24px;transform:translateX(-50%) translateY(20px);'
                  'max-width:min(520px,calc(100vw - 32px));padding:12px 18px;border-radius:10px;background:#2b2b2b;color:#fff;'
                  'font-size:13.5px;line-height:1.45;box-shadow:0 8px 24px rgba(0,0,0,.3);opacity:0;pointer-events:none;'
                  'transition:opacity .2s,transform .2s;z-index:600;}'
                  '#elab-toast.show{opacity:1;transform:translateX(-50%) translateY(0);}')
    extra += (ROOT + '{--elab-top:0px;}' + ROOT + ' h1{font-family:inherit;font-weight:700;letter-spacing:normal;text-wrap:wrap;}'
              '@media (max-width:760px){' + ROOT + '{--elab-top:61px;}}')
    return '\n' + cfg['css_marker'] + '\n' + scoped + dark_css + extra + '\n'


def build_panel(f, cfg):
    body = f[f.index('<body>') + 6:f.index('<script')]
    body = re.sub(r'\s*<img src="data:[^"]+" alt="Grupo Gadens">', '', body, count=1)
    body = re.sub(r'\s*<button type="button" class="btn-theme-toggle".*?</button>', '', body, count=1, flags=re.S)
    k = body.find('<div class="help-fab"')
    if k != -1:
        depth = 0
        for mm in re.finditer(r'<div\b|</div>', body[k:]):
            depth += 1 if mm.group(0) == '<div' else -1
            if depth == 0:
                end = k + mm.end(); break
        body = body[:k] + body[end:]
    body = body.replace('id="pending-banner"', 'id="%s"' % cfg['banner'])
    assert cfg['h1_old'] in body, cfg['h1_old']
    body = body.replace(cfg['h1_old'], cfg['h1_new'])
    if cfg['rename_ids']:
        body = re.sub(r'\bid_(nome|email|setor|setor_sigla|empresa|aprovador|data|tarefa)\b', cfg['idprefix'] + r'id_\1', body)
    assert 'help-fab' not in body and 'btn-theme-toggle' not in body and 'data:image' not in body
    return ('    <div id="%s" data-view="%s">\n' % (cfg['root'], cfg['view'])) + body.strip() + '\n    </div>'


HELPERS = '''  var elabToastTimer = null;
  function elabToast(msg){
    var t = window.document.getElementById("elab-toast");
    if(!t){
      t = window.document.createElement("div");
      t.id = "elab-toast";
      t.setAttribute("role", "status");
      window.document.body.appendChild(t);
    }
    t.textContent = msg;
    t.classList.add("show");
    if(elabToastTimer) clearTimeout(elabToastTimer);
    elabToastTimer = setTimeout(function(){ t.classList.remove("show"); }, 7000);
  }
  var alert = function(msg){ elabToast(msg); };
  var elabBlockedHint = (window.self !== window.top) ? " Este recurso é bloqueado quando o painel é aberto dentro do Claude. Use o endereço do G-Docs no GitHub Pages." : "";

  function elabCopyFallback(text){
    try{
      var ta = document.createElement("textarea");
      ta.value = text;
      ta.setAttribute("readonly", "");
      ta.style.cssText = "position:fixed;top:0;left:0;opacity:0;";
      document.body.appendChild(ta);
      ta.select();
      var ok = document.execCommand("copy");
      document.body.removeChild(ta);
      return ok;
    }catch(e){ return false; }
  }

'''


def shim(root):
    return '''  var __real = window.document;
  var __root = __real.getElementById("%s");
  var document = {
    getElementById: function(id){ return __root.querySelector('[id="' + id + '"]'); },
    querySelector: function(sel){ return __root.querySelector(sel); },
    querySelectorAll: function(sel){ return __root.querySelectorAll(sel); },
    createElement: function(t){ return __real.createElement(t); },
    addEventListener: function(){ return __real.addEventListener.apply(__real, arguments); },
    removeEventListener: function(){ return __real.removeEventListener.apply(__real, arguments); },
    execCommand: function(){ return __real.execCommand.apply(__real, arguments); },
    get activeElement(){ return __real.activeElement; },
    get body(){ return __real.body; },
    get documentElement(){ return __real.documentElement; }
  };

''' % root


def build_js(f, cfg):
    scripts = re.findall(r'<script[^>]*>(.*?)</script>', f, re.S)
    main = scripts[1]
    assert main.count('initTheme();\n') == 1
    main = main.replace('initTheme();\n', '')
    mt = re.search(r'document\.getElementById\("btn-theme-toggle"\)\.addEventListener\("click",\s*function\(\)\{.*?\}\);', main, re.S)
    assert mt
    main = main.replace(mt.group(0), '')
    assert 'btn-theme-toggle' not in main
    main = main.replace('"pending-banner"', '"%s"' % cfg['banner'])
    if cfg['rename_ids']:
        main = re.sub(r'\bid_(nome|email|setor|setor_sigla|empresa|aprovador|data|tarefa)\b', cfg['idprefix'] + r'id_\1', main)

    # isolamento do DOM do formulario + avisos na tela
    main = must_replace(main, '  "use strict";\n', '  "use strict";\n\n' + shim(cfg['root']) + HELPERS)

    main = must_replace(main, '''        try{
          await navigator.clipboard.writeText(promptText);
          btn.classList.add("copied");''', '''        var copiado = false;
        try{
          await navigator.clipboard.writeText(promptText);
          copiado = true;
        }catch(err){
          copiado = elabCopyFallback(promptText);
        }
        try{
          if(!copiado) throw new Error("copy-failed");
          btn.classList.add("copied");''')
    main = must_replace(main, 'alert("Não foi possível copiar automaticamente. Selecione o texto manualmente e copie (Ctrl+C).");',
                        'alert("Não foi possível copiar automaticamente. Selecione o texto manualmente e copie (Ctrl+C)." + elabBlockedHint);')
    main = must_replace(main, '''        recognition.onerror = function(e){
          if(e.error === "not-allowed" || e.error === "service-not-allowed"){
            alert("Permissão de microfone negada. Habilite o microfone nas configurações do navegador para usar o ditado por voz.");
          }
        };''', '''        recognition.onerror = function(e){
          if(e.error === "not-allowed" || e.error === "service-not-allowed"){
            alert("Permissão de microfone negada. Habilite o microfone nas configurações do navegador para usar o ditado por voz." + elabBlockedHint);
          }else if(e.error === "audio-capture"){
            alert("Não encontrei nenhum microfone neste computador.");
          }else if(e.error === "network"){
            alert("O ditado por voz precisa de internet. Verifique a conexão e tente de novo.");
          }else if(e.error !== "no-speech" && e.error !== "aborted"){
            alert("Não consegui usar o ditado por voz (" + e.error + ")." + elabBlockedHint);
          }
        };''')
    main = must_replace(main, '''        try{
          recognition.start();
        }catch(err){
          btn.classList.remove("listening");
        }''', '''        try{
          recognition.start();
        }catch(err){
          btn.classList.remove("listening");
          alert("Não consegui iniciar o ditado por voz." + elabBlockedHint);
        }''')
    return scripts[0], main


# ---------- gerar tudo ----------
css_parts = []; panels = {}; js_parts = {}; fflate = None
for key in ('pop', 'it', 'rg', 'at', 'pl'):
    cfg = CFG[key]
    f = open(cfg['src'], encoding='utf-8').read()
    css_parts.append(build_css(f, cfg))
    panels[key] = build_panel(f, cfg)
    ff, main = build_js(f, cfg)
    fflate = fflate or ff
    js_parts[key] = main

# CSS: da primeira marca de formulario ate o fim do <style>
a = s.index(CFG['pop']['css_marker'])
a = s.rindex('\n', 0, a)  # inicio da linha da marca
b = s.index('</style>', a)
s = s[:a] + ''.join(css_parts) + s[b:]

# painéis
a = s.index('    <div id="elab-root" data-view="elaborar">')
b = s.index('\n\n    <div class="panel" data-view="ajuda">', a)
s = s[:a] + panels['pop'] + '\n\n' + panels['it'] + '\n\n' + panels['rg'] + '\n\n' + panels['at'] + '\n\n' + panels['pl'] + s[b:]

# scripts: de <script> da biblioteca ate o ultimo </script>
a = s.rindex('<script>', 0, s.index('fflate v0.8.2'))
last = s.rindex('</script>') + len('</script>')
s = s[:a] + '<script>\n' + fflate + '\n</script>\n<script>\n' + js_parts['pop'] + '\n</script>\n<script>\n' + js_parts['it'] + '\n</script>\n<script>\n' + js_parts['rg'] + '\n</script>\n<script>\n' + js_parts['at'] + '\n</script>\n<script>\n' + js_parts['pl'] + '\n</script>' + s[last:]

# ---------- menu, rotas e abas do painel (so aplica se ainda nao existir) ----------
if 'data-nav="elaborar-ata"' not in s:
    marker = '<i class="sb-sub-label">Regimento</i></a>\n'
    s = must_replace(s, marker, marker + '      <a class="sb-sub-link" href="#/elaborar-ata" data-nav="elaborar-ata" title="Elaborar Ata">'
                     '<svg class="sb-ico" viewBox="0 0 24 24" aria-hidden="true"><path d="M6 3h12v18H6z"/><path d="M9 8h6M9 12h6M9 16h3"/>'
                     '<path d="M15 16.5l1.2 1.2 2-2.2"/></svg><i class="sb-sub-label">Ata</i></a>\n')
    s = must_replace(s, '"elaborar-regimento","ajuda"]', '"elaborar-regimento","elaborar-ata","ajuda"]')
    s = must_replace(s, 'const inElab = (v === "elaborar" || v === "elaborar-it" || v === "elaborar-regimento");',
                     'const inElab = (v === "elaborar" || v === "elaborar-it" || v === "elaborar-regimento" || v === "elaborar-ata");')
    s = must_replace(s, 'if(v === "elaborar" || v === "elaborar-it" || v === "elaborar-regimento"){',
                     'if(v === "elaborar" || v === "elaborar-it" || v === "elaborar-regimento" || v === "elaborar-ata"){')
    s = must_replace(s, '#elab-root textarea.autogrow, #elit-root textarea.autogrow, #elrg-root textarea.autogrow',
                     '#elab-root textarea.autogrow, #elit-root textarea.autogrow, #elrg-root textarea.autogrow, #elat-root textarea.autogrow')

if 'data-nav="elaborar-politica"' not in s:
    marker = '<i class="sb-sub-label">Ata</i></a>\n'
    s = must_replace(s, marker, marker + '      <a class="sb-sub-link" href="#/elaborar-politica" data-nav="elaborar-politica" title="Elaborar Política">'
                     '<svg class="sb-ico" viewBox="0 0 24 24" aria-hidden="true"><path d="M12 3l7 3v5c0 4.500-3 8-7 10-4-2-7-5.500-7-10V6z"/>'
                     '<path d="M9 12l2 2 4-4"/></svg><i class="sb-sub-label">Política</i></a>\n')
    s = must_replace(s, '"elaborar-ata","ajuda"]', '"elaborar-ata","elaborar-politica","ajuda"]')
    s = must_replace(s, 'v === "elaborar-regimento" || v === "elaborar-ata");',
                     'v === "elaborar-regimento" || v === "elaborar-ata" || v === "elaborar-politica");')
    s = must_replace(s, 'if(v === "elaborar" || v === "elaborar-it" || v === "elaborar-regimento" || v === "elaborar-ata"){',
                     'if(v === "elaborar" || v === "elaborar-it" || v === "elaborar-regimento" || v === "elaborar-ata" || v === "elaborar-politica"){')
    s = must_replace(s, '#elat-root textarea.autogrow").forEach', '#elat-root textarea.autogrow, #elpl-root textarea.autogrow").forEach')

open(DST, 'w', encoding='utf-8', newline='\n').write(s)
print('ok', len(s))
