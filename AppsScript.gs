/* =============================================================================
   EOD Dropscale — Google Apps Script (backend)
   -----------------------------------------------------------------------------
   Mesmo desenho do CS Dashboard do Vitor: a planilha é o banco de dados, este
   script é a API e o index.html (GitHub Pages) é o site.

   Abas (criadas sozinhas pela função configurar()):
     Usuarios   quem entra no site (senha guardada como hash com salt)
     Sessoes    tokens de login
     Catalogo   lista de tarefas conhecidas (alimenta as sugestões da busca)
     Registros  o que cada pessoa fez em cada dia (uma linha por tarefa + loja)
     Vistos     "visto" da chefe em cada EOD (pessoa + dia)
     Mensagens  conversa da chefe com a pessoa em cada EOD

   Para instalar: veja o README.md do repositório.
   ========================================================================== */

var VERSION = 1;
var TZ = 'America/Sao_Paulo';
var SESSAO_DIAS = 7;            // quanto tempo o login vale
var STATUS_FINAL = 'Concluído'; // tarefa longa com esse status para de ser sugerida

var ABAS = {
  Usuarios:  ['email','nome','papel','lojas','eod','salt','hash','ativo','criadoEm','ultimoAcesso'],
  Sessoes:   ['token','email','criadoEm','expiraEm'],
  Catalogo:  ['id','nome','tipo','lojaPadrao','criadoPor','criadoEm','ativo'],
  Registros: ['id','data','email','tarefaId','tarefa','tipo','loja','quantidade','minutos','status','bloqueado','bloqueioNota','obs','criadoEm','atualizadoEm'],
  Vistos:    ['data','email','vistoPor','vistoEm'],
  Mensagens: ['id','data','email','autor','autorNome','texto','criadoEm']
};

/* ---------------------------------------------------------------------------
   Rode UMA vez no editor (botão Executar com "configurar" selecionado).
   Cria as abas e mostra no registro de execução o código de primeiro acesso.
   ------------------------------------------------------------------------ */
function configurar() {
  Object.keys(ABAS).forEach(function (n) { aba_(n); });
  var props = PropertiesService.getScriptProperties();
  var codigo = props.getProperty('CODIGO_PRIMEIRO_ACESSO');
  if (!codigo) {
    codigo = Utilities.getUuid().replace(/-/g, '').slice(0, 8).toUpperCase();
    props.setProperty('CODIGO_PRIMEIRO_ACESSO', codigo);
  }
  Logger.log('Abas prontas. Código de primeiro acesso: ' + codigo);
  Logger.log('Use esse código no site, na tela "Primeiro acesso", para criar o primeiro admin.');
}

/* ======================= ENTRADA ======================= */

function doGet(e) {
  return json_({ ok: true, app: 'eod-dropscale', version: VERSION });
}

function doPost(e) {
  var d;
  try { d = JSON.parse(e.postData.contents); }
  catch (err) { return json_({ ok: false, error: 'Pedido inválido.' }); }
  try { return json_(rota_(d || {})); }
  catch (err) { return json_({ ok: false, error: String(err && err.message || err) }); }
}

function rota_(d) {
  switch (d.action) {
    case 'ping':           return { ok: true, version: VERSION, temUsuarios: lerUsuarios_().length > 0 };
    case 'primeiroAcesso': return primeiroAcesso_(d);
    case 'login':          return login_(d);
  }
  var u = sessao_(d.token);
  if (!u) return { ok: false, error: 'sessao', msg: 'Sua sessão expirou. Entre de novo.' };

  switch (d.action) {
    case 'eu':              return { ok: true, user: publico_(u) };
    case 'logout':          return logout_(d);
    case 'trocarSenha':     return trocarSenha_(u, d);
    case 'dia':             return dia_(u, d);
    case 'salvarRegistro':  return salvarRegistro_(u, d);
    case 'excluirRegistro': return excluirRegistro_(u, d);
    case 'criarTarefa':     return criarTarefa_(u, d);
    case 'mensagem':        return mensagem_(u, d);
    case 'historico':       return historico_(u, d);
  }
  if (u.papel !== 'admin') throw new Error('Só admin pode fazer isso.');
  switch (d.action) {
    case 'painel':          return painel_(u, d);
    case 'visto':           return visto_(u, d);
    case 'equipe':          return { ok: true, users: lerUsuarios_().map(publico_) };
    case 'salvarUsuario':   return salvarUsuario_(u, d);
    case 'editarTarefa':    return editarTarefa_(u, d);
  }
  throw new Error('Ação desconhecida.');
}

/* ======================= PLANILHA ======================= */

function aba_(nome) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(nome);
  if (!sh) {
    sh = ss.insertSheet(nome);
    sh.getRange(1, 1, 1, ABAS[nome].length).setValues([ABAS[nome]]).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  return sh;
}

/* Lê a aba inteira como objetos. Usa getDisplayValues (texto puro), como no
   CS Dashboard: evita datas virando epoch e horas virando serial de 1899. */
function linhas_(nome) {
  var sh = aba_(nome);
  var n = sh.getLastRow();
  if (n < 2) return [];
  var head = ABAS[nome];
  var vals = sh.getRange(2, 1, n - 1, head.length).getDisplayValues();
  var out = [];
  for (var i = 0; i < vals.length; i++) {
    var o = { _row: i + 2 };
    for (var j = 0; j < head.length; j++) o[head[j]] = vals[i][j];
    out.push(o);
  }
  return out;
}

function linhaDe_(nome, obj) {
  return ABAS[nome].map(function (k) {
    var v = obj[k];
    if (v === undefined || v === null) return '';
    if (typeof v === 'number') return v;
    return txt_(v);
  });
}

function inserir_(nome, obj) { aba_(nome).appendRow(linhaDe_(nome, obj)); }

function atualizar_(nome, row, obj) {
  aba_(nome).getRange(row, 1, 1, ABAS[nome].length).setValues([linhaDe_(nome, obj)]);
}

function apagarLinha_(nome, row) { aba_(nome).deleteRow(row); }

function comLock_(fn) {
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try { return fn(); } finally { lock.releaseLock(); }
}

/* ======================= USUÁRIOS E LOGIN ======================= */

function lerUsuarios_() {
  return linhas_('Usuarios').map(function (u) {
    u.email = emailLimpo_(u.email);
    u.ativo = simNao_(u.ativo) === 'SIM';
    u.eod = u.eod === '' ? true : simNao_(u.eod) === 'SIM';
    u.lojas = u.lojas ? u.lojas.split('|').filter(String) : [];
    return u;
  }).filter(function (u) { return u.email; });
}

function publico_(u) {
  return { email: u.email, nome: u.nome, papel: u.papel, lojas: u.lojas, eod: u.eod,
           ativo: u.ativo, criadoEm: u.criadoEm, ultimoAcesso: u.ultimoAcesso };
}

function hash_(salt, senha) {
  var b = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, salt + '::' + senha, Utilities.Charset.UTF_8);
  return b.map(function (x) { return ('0' + (x & 0xff).toString(16)).slice(-2); }).join('');
}

function primeiroAcesso_(d) {
  return comLock_(function () {
    if (lerUsuarios_().length) throw new Error('O primeiro acesso já foi feito. Entre com e-mail e senha.');
    var codigo = PropertiesService.getScriptProperties().getProperty('CODIGO_PRIMEIRO_ACESSO');
    if (!codigo || String(d.codigo || '').trim().toUpperCase() !== codigo)
      throw new Error('Código de primeiro acesso incorreto. Rode configurar() no editor para ver o código.');
    var email = emailLimpo_(d.email), nome = limpa_(d.nome, 60), senha = String(d.senha || '');
    if (!/@/.test(email)) throw new Error('E-mail inválido.');
    if (!nome) throw new Error('Digite o seu nome.');
    if (senha.length < 6) throw new Error('A senha precisa ter pelo menos 6 caracteres.');
    var salt = aleatorio_();
    inserir_('Usuarios', { email: email, nome: nome, papel: 'admin', lojas: '', eod: 'SIM', salt: salt,
                           hash: hash_(salt, senha), ativo: 'SIM', criadoEm: iso_(new Date()), ultimoAcesso: '' });
    return login_({ email: email, senha: senha });
  });
}

function login_(d) {
  var email = emailLimpo_(d.email), senha = String(d.senha || '');
  var u = lerUsuarios_().filter(function (x) { return x.email === email; })[0];
  if (!u || !u.ativo || hash_(u.salt, senha) !== u.hash) {
    Utilities.sleep(800);
    throw new Error('E-mail ou senha incorretos.');
  }
  var token = aleatorio_();
  var agora = new Date();
  var expira = new Date(agora.getTime() + SESSAO_DIAS * 864e5);
  comLock_(function () {
    limparSessoesVencidas_();
    inserir_('Sessoes', { token: token, email: email, criadoEm: iso_(agora), expiraEm: iso_(expira) });
    var sh = aba_('Usuarios');
    sh.getRange(u._row, ABAS.Usuarios.indexOf('ultimoAcesso') + 1).setValue(txt_(iso_(agora)));
  });
  return { ok: true, token: token, user: publico_(u) };
}

function logout_(d) {
  comLock_(function () {
    var s = linhas_('Sessoes').filter(function (x) { return x.token === d.token; })[0];
    if (s) apagarLinha_('Sessoes', s._row);
  });
  CacheService.getScriptCache().remove('s_' + d.token);
  return { ok: true };
}

function sessao_(token) {
  if (!/^[0-9a-f]{64}$/.test(String(token || ''))) return null;
  var cache = CacheService.getScriptCache();
  var email = cache.get('s_' + token);
  if (!email) {
    var s = linhas_('Sessoes').filter(function (x) { return x.token === token; })[0];
    if (!s || s.expiraEm < iso_(new Date())) return null;
    email = s.email;
    cache.put('s_' + token, email, 600);
  }
  var u = lerUsuarios_().filter(function (x) { return x.email === email; })[0];
  return u && u.ativo ? u : null;
}

function limparSessoesVencidas_() {
  var agora = iso_(new Date());
  var velhas = linhas_('Sessoes').filter(function (x) { return x.expiraEm < agora; });
  for (var i = velhas.length - 1; i >= 0; i--) apagarLinha_('Sessoes', velhas[i]._row);
}

function trocarSenha_(u, d) {
  if (hash_(u.salt, String(d.atual || '')) !== u.hash) throw new Error('A senha atual não confere.');
  var nova = String(d.nova || '');
  if (nova.length < 6) throw new Error('A senha nova precisa ter pelo menos 6 caracteres.');
  comLock_(function () {
    var salt = aleatorio_();
    var sh = aba_('Usuarios');
    sh.getRange(u._row, ABAS.Usuarios.indexOf('salt') + 1, 1, 2).setValues([[txt_(salt), txt_(hash_(salt, nova))]]);
  });
  return { ok: true };
}

function salvarUsuario_(admin, d) {
  return comLock_(function () {
    var todos = lerUsuarios_();
    var email = emailLimpo_(d.email);
    if (!/@/.test(email)) throw new Error('E-mail inválido.');
    var nome = limpa_(d.nome, 60);
    if (!nome) throw new Error('Digite o nome.');
    var papel = d.papel === 'admin' ? 'admin' : 'agente';
    var ativo = d.ativo !== false;
    var lojas = (d.lojas || []).map(function (l) { return limpa_(l, 40); }).filter(String).join('|');
    var existente = todos.filter(function (x) { return x.email === email; })[0];
    var original = d.original ? emailLimpo_(d.original) : '';
    if (original && original !== email) throw new Error('Não dá para trocar o e-mail. Crie a pessoa de novo com o e-mail novo e desative a antiga.');

    if (existente && !original) throw new Error('Já existe alguém com esse e-mail.');
    if (existente && existente.email === admin.email && (!ativo || papel !== 'admin'))
      throw new Error('Você não pode tirar o seu próprio acesso de admin.');

    var senha = String(d.senha || '');
    var obj;
    if (existente) {
      obj = { email: email, nome: nome, papel: papel, lojas: lojas, eod: d.eod === false ? 'NAO' : 'SIM',
              salt: existente.salt, hash: existente.hash, ativo: ativo ? 'SIM' : 'NAO',
              criadoEm: existente.criadoEm, ultimoAcesso: existente.ultimoAcesso };
      if (senha) {
        if (senha.length < 6) throw new Error('A senha precisa ter pelo menos 6 caracteres.');
        obj.salt = aleatorio_(); obj.hash = hash_(obj.salt, senha);
      }
      atualizar_('Usuarios', existente._row, obj);
      if (!ativo || senha) esquecerSessoes_(email);
    } else {
      if (senha.length < 6) throw new Error('Defina uma senha com pelo menos 6 caracteres.');
      var salt = aleatorio_();
      obj = { email: email, nome: nome, papel: papel, lojas: lojas, eod: d.eod === false ? 'NAO' : 'SIM',
              salt: salt, hash: hash_(salt, senha), ativo: ativo ? 'SIM' : 'NAO', criadoEm: iso_(new Date()), ultimoAcesso: '' };
      inserir_('Usuarios', obj);
    }
    return { ok: true, users: lerUsuarios_().map(publico_) };
  });
}

function esquecerSessoes_(email) {
  var cache = CacheService.getScriptCache();
  var ss = linhas_('Sessoes').filter(function (x) { return emailLimpo_(x.email) === email; });
  for (var i = ss.length - 1; i >= 0; i--) { cache.remove('s_' + ss[i].token); apagarLinha_('Sessoes', ss[i]._row); }
}

/* ======================= CATÁLOGO DE TAREFAS ======================= */

function lerCatalogo_() {
  return linhas_('Catalogo').map(function (t) {
    return { _row: t._row, id: t.id, nome: t.nome, tipo: t.tipo === 'longa' ? 'longa' : 'repetitiva',
             lojaPadrao: t.lojaPadrao, criadoPor: t.criadoPor, criadoEm: t.criadoEm, ativo: simNao_(t.ativo) === 'SIM' };
  });
}

function catalogoPublico_() {
  return lerCatalogo_().filter(function (t) { return t.ativo; })
    .map(function (t) { return { id: t.id, nome: t.nome, tipo: t.tipo, lojaPadrao: t.lojaPadrao }; });
}

function criarTarefa_(u, d) {
  return comLock_(function () {
    var nome = limpa_(d.nome, 80);
    if (!nome) throw new Error('Dê um nome para a tarefa.');
    var chave = chave_(nome);
    var ja = lerCatalogo_().filter(function (t) { return chave_(t.nome) === chave; })[0];
    if (ja) return { ok: true, tarefa: { id: ja.id, nome: ja.nome, tipo: ja.tipo, lojaPadrao: ja.lojaPadrao }, existia: true };
    var t = { id: 't' + Utilities.getUuid().replace(/-/g, '').slice(0, 10), nome: nome,
              tipo: d.tipo === 'longa' ? 'longa' : 'repetitiva', lojaPadrao: limpa_(d.lojaPadrao, 40),
              criadoPor: u.email, criadoEm: iso_(new Date()), ativo: 'SIM' };
    inserir_('Catalogo', t);
    return { ok: true, tarefa: { id: t.id, nome: t.nome, tipo: t.tipo, lojaPadrao: t.lojaPadrao } };
  });
}

function editarTarefa_(u, d) {
  return comLock_(function () {
    var t = lerCatalogo_().filter(function (x) { return x.id === d.id; })[0];
    if (!t) throw new Error('Tarefa não encontrada.');
    var nome = limpa_(d.nome, 80) || t.nome;
    atualizar_('Catalogo', t._row, { id: t.id, nome: nome, tipo: d.tipo === 'longa' ? 'longa' : 'repetitiva',
      lojaPadrao: limpa_(d.lojaPadrao, 40), criadoPor: t.criadoPor, criadoEm: t.criadoEm, ativo: d.ativo === false ? 'NAO' : 'SIM' });
    return { ok: true, catalogo: catalogoPublico_() };
  });
}

/* ======================= REGISTROS ======================= */

function lerRegistros_(filtro) {
  return linhas_('Registros').map(function (r) {
    r.email = emailLimpo_(r.email);
    r.quantidade = r.quantidade === '' ? null : Number(r.quantidade);
    r.minutos = r.minutos === '' ? null : Number(r.minutos);
    r.bloqueado = simNao_(r.bloqueado) === 'SIM';
    return r;
  }).filter(filtro || function () { return true; });
}

function limpaRegistro_(r) {
  return { id: r.id, data: r.data, email: r.email, tarefaId: r.tarefaId, tarefa: r.tarefa, tipo: r.tipo,
           loja: r.loja, quantidade: r.quantidade, minutos: r.minutos, status: r.status,
           bloqueado: r.bloqueado, bloqueioNota: r.bloqueioNota, obs: r.obs, atualizadoEm: r.atualizadoEm };
}

function salvarRegistro_(u, d) {
  var data = dataValida_(d.data);
  return comLock_(function () {
    var cat = lerCatalogo_().filter(function (t) { return t.id === d.tarefaId; })[0];
    if (!cat) throw new Error('Tarefa não encontrada no catálogo.');
    var agora = iso_(new Date());
    var num = function (v) { if (v === '' || v === null || v === undefined) return ''; var n = Number(v); if (!isFinite(n) || n < 0) throw new Error('Número inválido.'); return Math.round(n); };
    var obj = {
      data: data, email: u.email, tarefaId: cat.id, tarefa: cat.nome, tipo: cat.tipo,
      loja: limpa_(d.loja, 40), quantidade: num(d.quantidade), minutos: num(d.minutos),
      status: cat.tipo === 'longa' ? limpa_(d.status, 30) : '',
      bloqueado: d.bloqueado ? 'SIM' : 'NAO', bloqueioNota: d.bloqueado ? limpa_(d.bloqueioNota, 300) : '',
      obs: limpa_(d.obs, 1000), atualizadoEm: agora
    };
    if (d.id) {
      var atual = lerRegistros_(function (r) { return r.id === d.id; })[0];
      if (!atual) throw new Error('Registro não encontrado (pode ter sido apagado).');
      if (atual.email !== u.email) throw new Error('Você só pode editar o seu próprio EOD.');
      obj.id = atual.id; obj.criadoEm = atual.criadoEm;
      atualizar_('Registros', atual._row, obj);
    } else {
      obj.id = 'r' + Utilities.getUuid().replace(/-/g, '').slice(0, 14);
      obj.criadoEm = agora;
      inserir_('Registros', obj);
    }
    var salvo = lerRegistros_(function (r) { return r.id === obj.id; })[0];
    return { ok: true, registro: limpaRegistro_(salvo) };
  });
}

function excluirRegistro_(u, d) {
  return comLock_(function () {
    var r = lerRegistros_(function (x) { return x.id === d.id; })[0];
    if (!r) return { ok: true };
    if (r.email !== u.email) throw new Error('Você só pode apagar o seu próprio EOD.');
    apagarLinha_('Registros', r._row);
    return { ok: true };
  });
}

/* Dia de uma pessoa: o que já registrou, as sugestões "com base em ontem",
   o visto da chefe e a conversa. "Ontem" é o último dia com registro antes
   da data (segunda-feira puxa a sexta). Tarefa longa que não chegou em
   Concluído continua sendo sugerida até ser concluída. */
function dia_(u, d) {
  var data = dataValida_(d.data);
  var email = u.email;
  if (d.email && u.papel === 'admin') email = emailLimpo_(d.email);

  var meus = lerRegistros_(function (r) { return r.email === email && r.data <= data; });
  var hoje = meus.filter(function (r) { return r.data === data; });
  var antes = meus.filter(function (r) { return r.data < data; });

  var baseData = '';
  antes.forEach(function (r) { if (r.data > baseData) baseData = r.data; });

  var chaveReg = function (r) { return r.tarefaId + '|' + r.loja; };
  var jaHoje = {};
  hoje.forEach(function (r) { jaHoje[chaveReg(r)] = true; });

  // último estado de cada tarefa+loja antes de hoje
  var ultimo = {};
  antes.sort(function (a, b) { return a.data < b.data ? -1 : a.data > b.data ? 1 : (a.atualizadoEm < b.atualizadoEm ? -1 : 1); });
  antes.forEach(function (r) { ultimo[chaveReg(r)] = r; });

  var sugestoes = [];
  var visto = {};
  var add = function (r, motivo) {
    var k = chaveReg(r);
    if (jaHoje[k] || visto[k]) return;
    visto[k] = true;
    var s = limpaRegistro_(ultimo[k] || r);
    s.motivo = motivo;
    sugestoes.push(s);
  };
  antes.filter(function (r) { return r.data === baseData; }).forEach(function (r) { add(r, 'ontem'); });
  Object.keys(ultimo).forEach(function (k) {
    var r = ultimo[k];
    if (r.tipo === 'longa' && r.status !== STATUS_FINAL && r.data !== baseData) add(r, 'aberta');
  });

  return {
    ok: true, data: data, email: email, baseData: baseData,
    registros: hoje.map(limpaRegistro_), sugestoes: sugestoes,
    visto: lerVisto_(data, email), mensagens: lerMensagens_(data, email),
    catalogo: catalogoPublico_()
  };
}

/* ======================= VISÃO DA CHEFE ======================= */

function painel_(u, d) {
  var data = dataValida_(d.data);
  var users = lerUsuarios_().filter(function (x) { return x.ativo; }).map(publico_);
  var regs = lerRegistros_(function (r) { return r.data === data; }).map(limpaRegistro_);
  var vistos = linhas_('Vistos').filter(function (v) { return v.data === data; })
    .map(function (v) { return { email: emailLimpo_(v.email), vistoPor: v.vistoPor, vistoEm: v.vistoEm }; });
  var msgs = lerMensagens_(data, null);
  return { ok: true, data: data, users: users, registros: regs, vistos: vistos, mensagens: msgs };
}

function lerVisto_(data, email) {
  var v = linhas_('Vistos').filter(function (x) { return x.data === data && emailLimpo_(x.email) === email; })[0];
  return v ? { vistoPor: v.vistoPor, vistoEm: v.vistoEm } : null;
}

function visto_(u, d) {
  var data = dataValida_(d.data), email = emailLimpo_(d.email);
  return comLock_(function () {
    var v = linhas_('Vistos').filter(function (x) { return x.data === data && emailLimpo_(x.email) === email; })[0];
    if (d.desfazer) { if (v) apagarLinha_('Vistos', v._row); return { ok: true, visto: null }; }
    if (!v) inserir_('Vistos', { data: data, email: email, vistoPor: u.nome, vistoEm: iso_(new Date()) });
    return { ok: true, visto: lerVisto_(data, email) };
  });
}

/* ======================= CONVERSA ======================= */

function lerMensagens_(data, email) {
  return linhas_('Mensagens').filter(function (m) {
    return m.data === data && (!email || emailLimpo_(m.email) === email);
  }).map(function (m) {
    return { id: m.id, data: m.data, email: emailLimpo_(m.email), autor: emailLimpo_(m.autor),
             autorNome: m.autorNome, texto: m.texto, criadoEm: m.criadoEm };
  });
}

function mensagem_(u, d) {
  var data = dataValida_(d.data);
  var email = u.papel === 'admin' && d.email ? emailLimpo_(d.email) : u.email;
  var texto = limpa_(d.texto, 2000);
  if (!texto) throw new Error('Escreva a mensagem.');
  return comLock_(function () {
    inserir_('Mensagens', { id: 'm' + Utilities.getUuid().replace(/-/g, '').slice(0, 14), data: data, email: email,
                            autor: u.email, autorNome: u.nome, texto: texto, criadoEm: iso_(new Date()) });
    return { ok: true, mensagens: lerMensagens_(data, email) };
  });
}

/* ======================= HISTÓRICO ======================= */

function historico_(u, d) {
  var de = dataValida_(d.de), ate = dataValida_(d.ate);
  if (de > ate) { var t = de; de = ate; ate = t; }
  var email = u.papel === 'admin' ? (d.email ? emailLimpo_(d.email) : '') : u.email;
  var regs = lerRegistros_(function (r) {
    return r.data >= de && r.data <= ate && (!email || r.email === email);
  }).map(limpaRegistro_);
  var vistos = linhas_('Vistos').filter(function (v) {
    return v.data >= de && v.data <= ate && (!email || emailLimpo_(v.email) === email);
  }).map(function (v) { return { data: v.data, email: emailLimpo_(v.email), vistoPor: v.vistoPor }; });
  var msgs = linhas_('Mensagens').filter(function (m) {
    return m.data >= de && m.data <= ate && (!email || emailLimpo_(m.email) === email);
  }).map(function (m) { return { data: m.data, email: emailLimpo_(m.email), autorNome: m.autorNome, texto: m.texto, criadoEm: m.criadoEm }; });
  var users = u.papel === 'admin' ? lerUsuarios_().map(publico_) : [publico_(u)];
  return { ok: true, de: de, ate: ate, registros: regs, vistos: vistos, mensagens: msgs, users: users };
}

/* ======================= UTILITÁRIOS ======================= */

function json_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
function txt_(v) { return "'" + String(v === null || v === undefined ? '' : v); }
function aleatorio_() { return Utilities.getUuid().replace(/-/g, '') + Utilities.getUuid().replace(/-/g, ''); }
function emailLimpo_(e) { return String(e || '').trim().toLowerCase(); }
function iso_(d) { return Utilities.formatDate(d, TZ, "yyyy-MM-dd'T'HH:mm:ss"); }
function hoje_() { return Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd'); }
function limpa_(v, max) { return String(v === null || v === undefined ? '' : v).trim().slice(0, max); }
function simNao_(v) { var t = String(v === undefined || v === null ? '' : v).trim().toUpperCase(); return (v === true || t === 'TRUE' || t === 'SIM' || t === '1') ? 'SIM' : 'NAO'; }
function chave_(s) { return String(s || '').normalize('NFD').replace(/[̀-ͯ]/g, '').toLowerCase().replace(/\s+/g, ' ').trim(); }
function dataValida_(v) {
  var s = String(v || '').trim();
  if (!s) return hoje_();
  if (!/^\d{4}-\d{2}-\d{2}$/.test(s)) throw new Error('Data inválida.');
  return s;
}
