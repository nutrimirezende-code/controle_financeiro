// ══════════════════════════════════════════════════════════════════
//  CONTROLE DE CARTÕES — Apps Script independente (v2)
//  Planilha: crie uma nova planilha e vincule este script a ela
//  Implante como Web App (acesso: qualquer pessoa)
//
//  NOVIDADES v2:
//  • Link de compartilhamento somente-leitura (token separado da sessão)
//  • Backup automático diário na aba "Backups" (últimos 30 dias)
//  • Bloqueio de 15 min após 5 tentativas erradas de PIN
//  • Trocar o PIN derruba as outras sessões
//  • Email de prévia do acerto todo dia 1 e 15 (rode criarGatilhos 1x)
// ══════════════════════════════════════════════════════════════════

const ABA_DADOS = 'Dados';
const ABA_CONFIG = 'Config';
const ABA_BACKUPS = 'Backups';
const CHUNK_SIZE = 40000; // caracteres por célula (limite do Sheets é 50k)

const EMAIL_AVISO = 'michellegprezende@gmail.com';
const URL_APP = 'https://controle-financeiro-eta-hazel.vercel.app';

// ── Helpers ──────────────────────────────────────────────────────

function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function getOrCreateSheet(nome) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let aba = ss.getSheetByName(nome);
  if (!aba) {
    aba = ss.insertSheet(nome);
  }
  return aba;
}

function initPlanilha() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.setSpreadsheetLocale('pt_BR');
  ss.setSpreadsheetTimeZone('America/Sao_Paulo');

  // Aba Config: PIN e metadata
  const cfg = getOrCreateSheet(ABA_CONFIG);
  if (!cfg.getRange('A1').getValue()) {
    cfg.getRange('A1').setValue('pin');
    cfg.getRange('A2').setValue('1234');
    cfg.getRange('B1').setValue('ultimaAtualizacao');
  }

  // Aba Dados: JSON em chunks (coluna A a partir da linha 2)
  const dados = getOrCreateSheet(ABA_DADOS);
  if (!dados.getRange('A1').getValue()) {
    dados.getRange('A1').setValue('chunks');
    dados.getRange('A2').setValue('{}');
  }

  SpreadsheetApp.getUi().alert('Planilha inicializada! PIN padrão: 1234');
}

// ── Gatilhos (rode UMA VEZ manualmente: Executar > criarGatilhos) ─

function criarGatilhos() {
  // Remove gatilhos antigos desta função para não duplicar
  const triggers = ScriptApp.getProjectTriggers();
  for (let i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'emailQuinzenal') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
  // Roda todo dia entre 7h e 8h; a função só age nos dias 1 e 15
  ScriptApp.newTrigger('emailQuinzenal').timeBased().everyDays(1).atHour(7).create();
}

// ── Sessão ───────────────────────────────────────────────────────

function gerarSessao() {
  const props = PropertiesService.getScriptProperties();
  limparSessoesVencidas(props); // evita que os SESS_ encham as propriedades
  const sessao = Utilities.getUuid();
  props.setProperty('SESS_' + sessao, String(Date.now()));
  return sessao;
}

// Remove sessões com mais de 7 dias e, se ainda houver muitas, as mais antigas
function limparSessoesVencidas(props) {
  try {
    const todas = props.getProperties();
    const sess = [];
    const CORTE = Date.now() - 7 * 24 * 60 * 60 * 1000;
    for (const k in todas) {
      if (k.indexOf('SESS_') === 0) {
        const ts = Number(todas[k]) || 0;
        if (ts < CORTE) props.deleteProperty(k);
        else sess.push({ k: k, ts: ts });
      }
    }
    // Mantém no máximo as 20 sessões mais recentes
    if (sess.length > 20) {
      sess.sort(function(a, b) { return a.ts - b.ts; });
      for (let i = 0; i < sess.length - 20; i++) props.deleteProperty(sess[i].k);
    }
  } catch (e) {}
}

function validarSessao(sessao) {
  if (!sessao) return false;
  const props = PropertiesService.getScriptProperties();
  const ts = props.getProperty('SESS_' + sessao);
  if (!ts) return false;
  // Expira em 7 dias
  if (Date.now() - Number(ts) > 7 * 24 * 60 * 60 * 1000) {
    props.deleteProperty('SESS_' + sessao);
    return false;
  }
  return true;
}

// ── Token de compartilhamento (somente leitura, permanente) ──────

function obterShareToken() {
  const props = PropertiesService.getScriptProperties();
  let t = props.getProperty('SHARE_TOKEN');
  if (!t) {
    t = Utilities.getUuid().replace(/-/g, '');
    props.setProperty('SHARE_TOKEN', t);
  }
  return t;
}

// ── Chunks: salvar/ler JSON grande dividido em várias células ────

function salvarChunks(aba, json) {
  // Limpa dados antigos (linhas 2 em diante)
  const lastRow = aba.getLastRow();
  if (lastRow >= 2) {
    aba.getRange(2, 1, lastRow - 1, 1).clearContent();
  }
  // Divide em pedaços e salva cada um numa célula
  const chunks = [];
  for (let i = 0; i < json.length; i += CHUNK_SIZE) {
    chunks.push(json.substring(i, i + CHUNK_SIZE));
  }
  if (chunks.length === 0) chunks.push('{}');
  for (let i = 0; i < chunks.length; i++) {
    aba.getRange(2 + i, 1).setValue(chunks[i]);
  }
}

function lerChunks(aba) {
  const lastRow = aba.getLastRow();
  if (lastRow < 2) return '{}';
  const range = aba.getRange(2, 1, lastRow - 1, 1).getValues();
  let json = '';
  for (const row of range) {
    const val = (row[0] || '').toString();
    if (val) json += val;
  }
  return json || '{}';
}

// ── Backup automático: 1 por dia, mantém os últimos 30 ───────────

function backupDiario(jsonAtual) {
  if (!jsonAtual || jsonAtual === '{}') return;
  const aba = getOrCreateSheet(ABA_BACKUPS);
  const hojeStr = new Date().toLocaleDateString('pt-BR', { timeZone: 'America/Sao_Paulo' });
  const lastRow = aba.getLastRow();
  if (lastRow >= 1) {
    const ultimaData = (aba.getRange(lastRow, 1).getValue() || '').toString();
    if (ultimaData === hojeStr) return; // já existe backup de hoje
  }
  // Uma linha por dia: [data, chunk1, chunk2, ...]
  const linha = [hojeStr];
  for (let i = 0; i < jsonAtual.length; i += CHUNK_SIZE) {
    linha.push(jsonAtual.substring(i, i + CHUNK_SIZE));
  }
  aba.appendRow(linha);
  // Mantém só os últimos 30 dias
  const total = aba.getLastRow();
  if (total > 30) {
    aba.deleteRows(1, total - 30);
  }
}

// ── GET ──────────────────────────────────────────────────────────

function doGet(e) {
  try {
    const params = e.parameter || {};
    const action = params.action || '';

    if (action === 'load') {
      return handleLoad(params.sessao || '');
    }
    if (action === 'shared') {
      return handleShared(params.t || '');
    }

    return jsonResponse({ ok: false, erro: 'Ação desconhecida.' });
  } catch (err) {
    return jsonResponse({ ok: false, erro: err.message });
  }
}

function handleLoad(sessao) {
  if (!validarSessao(sessao)) {
    return jsonResponse({ ok: false, erro: 'Sessão inválida ou expirada.' });
  }
  const aba = getOrCreateSheet(ABA_DADOS);
  const raw = lerChunks(aba);
  let dados = {};
  try { dados = JSON.parse(raw); } catch (e) { dados = {}; }

  const cfg = getOrCreateSheet(ABA_CONFIG);
  const ultimaAtualizacao = (cfg.getRange('B2').getValue() || '').toString();

  return jsonResponse({ ok: true, dados, ultimaAtualizacao });
}

// Leitura pelo link compartilhado: só aceita o SHARE_TOKEN (nunca a sessão)
// e devolve apenas os campos que a página do Rogério mostra —
// nada de dicionário, orçamentos, acertos ou contas fixas.
function handleShared(t) {
  const props = PropertiesService.getScriptProperties();
  const tok = props.getProperty('SHARE_TOKEN');
  if (!t || !tok || t !== tok) {
    return jsonResponse({ ok: false, erro: 'Link inválido ou revogado.' });
  }
  const aba = getOrCreateSheet(ABA_DADOS);
  let dados = {};
  try { dados = JSON.parse(lerChunks(aba)); } catch (e) { dados = {}; }
  const faturas = (dados.faturas || []).map(function(f) {
    return {
      venc: f.venc, arquivo: f.arquivo,
      transacoes: (f.transacoes || []).map(function(x) {
        return {
          data: x.data, cartao: x.cartao, estabelecimento: x.estabelecimento,
          categoria: x.categoria, valor: x.valor, parcela: x.parcela, obs: x.obs
        };
      })
    };
  });
  return jsonResponse({ ok: true, dados: { faturas: faturas } });
}

// ── POST ─────────────────────────────────────────────────────────

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);

    // Mensagens do Telegram chegam neste mesmo endpoint (têm update_id)
    if (data.update_id || data.message || data.callback_query) {
      return handleTelegram(data);
    }

    const tipo = data.tipo || '';
    if (tipo === 'login')         return handleLogin(data);
    if (tipo === 'save')          return handleSave(data);
    if (tipo === 'getShareToken') return handleGetShareToken(data);
    if (tipo === 'revogarShare')  return handleRevogarShare(data);

    return jsonResponse({ ok: false, erro: 'Tipo desconhecido: ' + tipo });
  } catch (err) {
    return jsonResponse({ ok: false, erro: err.message });
  }
}

function handleLogin(data) {
  const pin = (data.pin || '').toString().trim();
  if (!pin) throw new Error('PIN não informado.');

  // Bloqueio: 5 erros seguidos = 15 minutos de espera
  const props = PropertiesService.getScriptProperties();
  const BLOQ_MS = 15 * 60 * 1000;
  let bloq = { n: 0, ts: 0 };
  try { bloq = JSON.parse(props.getProperty('LOGIN_BLOQ') || '{"n":0,"ts":0}'); } catch (e) {}
  const dentroJanela = Date.now() - bloq.ts < BLOQ_MS;
  if (bloq.n >= 5 && dentroJanela) {
    const min = Math.ceil((BLOQ_MS - (Date.now() - bloq.ts)) / 60000);
    throw new Error('Muitas tentativas erradas. Aguarde ' + min + ' minuto(s).');
  }

  const cfg = getOrCreateSheet(ABA_CONFIG);
  const pinSalvo = (cfg.getRange('A2').getValue() || '').toString().trim();

  if (pin !== pinSalvo) {
    props.setProperty('LOGIN_BLOQ', JSON.stringify({
      n: dentroJanela ? bloq.n + 1 : 1,
      ts: Date.now()
    }));
    throw new Error('PIN incorreto.');
  }

  props.deleteProperty('LOGIN_BLOQ');
  const sessao = gerarSessao();
  return jsonResponse({ ok: true, sessao });
}

function handleSave(data) {
  const sessao = (data.sessao || '').toString().trim();
  if (!validarSessao(sessao)) throw new Error('Sessão inválida ou expirada.');

  let dadosIn = data.dados;
  if (!dadosIn) throw new Error('Dados não informados.');
  if (typeof dadosIn === 'string') { try { dadosIn = JSON.parse(dadosIn); } catch (e) { throw new Error('Dados inválidos.'); } }

  // Trava + MERGE no servidor: nunca sobrescreve cegamente. Assim uma escrita do
  // bot (ou do outro aparelho) feita depois que este cliente carregou não é perdida.
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const aba = getOrCreateSheet(ABA_DADOS);
    let cloud = {};
    try { cloud = JSON.parse(lerChunks(aba)); } catch (e) { cloud = {}; }
    // Backup automático do estado ANTERIOR (1 por dia)
    try { backupDiario(JSON.stringify(cloud)); } catch (e) {}

    const merged = mesclarEstadosGs(dadosIn, cloud);
    salvarChunks(aba, JSON.stringify(merged));

    const cfg = getOrCreateSheet(ABA_CONFIG);
    cfg.getRange('B2').setValue(new Date().toLocaleString('pt-BR', { timeZone: 'America/Sao_Paulo' }));

    // Trocar PIN se solicitado — derruba as outras sessões
    if (data.novoPin) {
      cfg.getRange('A2').setValue(data.novoPin.toString().trim());
      const props = PropertiesService.getScriptProperties();
      const keys = props.getKeys();
      for (let i = 0; i < keys.length; i++) {
        if (keys[i].indexOf('SESS_') === 0 && keys[i] !== 'SESS_' + sessao) props.deleteProperty(keys[i]);
      }
    }
  } finally {
    lock.releaseLock();
  }

  return jsonResponse({ ok: true, salvo: true });
}

// Merge server-side — mesma lógica do app (mesclarEstados no index.html):
// união por id, updatedAt decide conflito, lixeira (tombstones) evita ressurreição,
// "enviado" pegajoso em extras/avulsas.
function normalizarEstadoGs(s) {
  if (!s || typeof s !== 'object') s = {};
  if (!s.faturas) s.faturas = [];
  if (!s.dicionario) s.dicionario = {};
  if (!s.orcamentos) s.orcamentos = {};
  if (!s.categorias) s.categorias = [];
  if (!s.cartoesMichelle) s.cartoesMichelle = [];
  if (!s.contasFixas) s.contasFixas = [];
  if (!s.extras) s.extras = [];
  if (!s.avulsas) s.avulsas = [];
  if (!s.acertos) s.acertos = [];
  if (!s.pendentes) s.pendentes = [];
  if (!s.assinaturas) s.assinaturas = [];
  if (!s.lixeira) s.lixeira = {};
  if (!s.updatedAt) s.updatedAt = 0;
  return s;
}

function mesclarEstadosGs(aIn, bIn) {
  var a = normalizarEstadoGs(aIn), b = normalizarEstadoGs(bIn);
  var base = (a.updatedAt || 0) > (b.updatedAt || 0) ? a : b;
  var outro = base === a ? b : a;
  var resultado = JSON.parse(JSON.stringify(base));

  var CORTE = Date.now() - 90 * 24 * 60 * 60 * 1000;
  var TIPOS = ['faturas','extras','avulsas','acertos','contasFixas','pendentes','assinaturas','dicionario','categorias','cartoesMichelle','orcamentos'];
  resultado.lixeira = {};
  for (var ti = 0; ti < TIPOS.length; ti++) {
    var tipo = TIPOS[ti];
    var mapa = Object.create(null);
    var la = (a.lixeira[tipo] || []).concat(b.lixeira[tipo] || []);
    for (var i = 0; i < la.length; i++) { var it = la[i]; if (it && it.ts > CORTE) mapa[it.id] = it; }
    var arr = []; for (var kk in mapa) arr.push(mapa[kk]);
    resultado.lixeira[tipo] = arr;
  }
  // Object.create(null) → sem cadeia de prototype, então 'constructor'/'toString'/etc.
  // não são considerados "excluídos" por herança (iguala o Set do cliente)
  function excluidos(tipo) { var s = Object.create(null); var l = resultado.lixeira[tipo] || []; for (var i = 0; i < l.length; i++) s[l[i].id] = true; return s; }

  function mesclarColecao(campo, tipoLix, enviadoWins) {
    var del = excluidos(tipoLix);
    var mapa = Object.create(null);
    var rc = resultado[campo] || [];
    for (var i = 0; i < rc.length; i++) { var item = rc[i]; if (!del[item.id]) mapa[item.id] = item; }
    var oc = outro[campo] || [];
    for (var j = 0; j < oc.length; j++) {
      var it2 = oc[j]; if (del[it2.id]) continue;
      var atual = mapa[it2.id];
      if (!atual) { mapa[it2.id] = it2; }
      else if (enviadoWins) {
        // vence o mais recente (updatedAt) por item, mas 'enviado' é pegajoso
        var enviado = atual.enviado || it2.enviado;
        if ((it2.updatedAt || 0) > (atual.updatedAt || 0)) mapa[it2.id] = it2;
        mapa[it2.id].enviado = enviado;
      }
      else if ((it2.updatedAt || 0) > (atual.updatedAt || 0)) { mapa[it2.id] = it2; }
    }
    var out = []; for (var k in mapa) out.push(mapa[k]);
    resultado[campo] = out;
  }
  mesclarColecao('faturas', 'faturas', false);
  mesclarColecao('contasFixas', 'contasFixas', false);
  mesclarColecao('acertos', 'acertos', false);
  mesclarColecao('pendentes', 'pendentes', false);
  mesclarColecao('assinaturas', 'assinaturas', false);
  mesclarColecao('extras', 'extras', true);
  mesclarColecao('avulsas', 'avulsas', true);

  var delDic = excluidos('dicionario');
  var od = outro.dicionario || {};
  for (var dk in od) if (od.hasOwnProperty(dk)) { if (!resultado.dicionario[dk] && !delDic[dk]) resultado.dicionario[dk] = od[dk]; }
  for (var dk2 in delDic) if (resultado.dicionario[dk2]) delete resultado.dicionario[dk2];

  var delCat = excluidos('categorias');
  var ocat = outro.categorias || [];
  for (var ci = 0; ci < ocat.length; ci++) { var c = ocat[ci]; if (resultado.categorias.indexOf(c) < 0 && !delCat[c]) resultado.categorias.push(c); }
  resultado.categorias = resultado.categorias.filter(function(x) { return !delCat[x]; });

  var delCart = excluidos('cartoesMichelle');
  var ocart = outro.cartoesMichelle || [];
  for (var cti = 0; cti < ocart.length; cti++) { var ct = ocart[cti]; if (resultado.cartoesMichelle.indexOf(ct) < 0 && !delCart[ct]) resultado.cartoesMichelle.push(ct); }
  resultado.cartoesMichelle = resultado.cartoesMichelle.filter(function(x) { return !delCart[x]; });

  var delOrc = excluidos('orcamentos');
  var oorc = outro.orcamentos || {};
  for (var ok in oorc) if (oorc.hasOwnProperty(ok)) { if (resultado.orcamentos[ok] == null && !delOrc[ok]) resultado.orcamentos[ok] = oorc[ok]; }
  for (var ok2 in delOrc) if (resultado.orcamentos[ok2] != null) delete resultado.orcamentos[ok2];

  resultado.updatedAt = Math.max(a.updatedAt || 0, b.updatedAt || 0);
  return resultado;
}

function handleGetShareToken(data) {
  const sessao = (data.sessao || '').toString().trim();
  if (!validarSessao(sessao)) throw new Error('Sessão inválida ou expirada.');
  return jsonResponse({ ok: true, token: obterShareToken() });
}

// Gera um token NOVO — o link antigo do Rogério para de funcionar na hora
function handleRevogarShare(data) {
  const sessao = (data.sessao || '').toString().trim();
  if (!validarSessao(sessao)) throw new Error('Sessão inválida ou expirada.');
  const props = PropertiesService.getScriptProperties();
  props.deleteProperty('SHARE_TOKEN');
  return jsonResponse({ ok: true, token: obterShareToken() });
}

// ── Destaques do mês (para o email) ──────────────────────────────

function vencParaCompetenciaGs(venc, meses) {
  if (!venc) return null;
  var p = String(venc).split('/');
  if (p.length < 3) return null;
  var mes = parseInt(p[1], 10);
  var ano = parseInt(p[2], 10);
  mes--; if (mes < 1) { mes = 12; ano--; }
  return meses[mes - 1] + '/' + ano;
}

function compKeyGs(comp, meses) {
  if (!comp) return '';
  var p = comp.split('/');
  var mi = meses.indexOf(p[0]) + 1;
  return p[1] + (mi < 10 ? '0' + mi : '' + mi);
}

// Total do mês vs média + categorias que mais subiram (versão robusta, sem
// depender de normalização de nomes de loja).
function destaquesDoMes(dados, meses, fmt) {
  var trans = [];
  var faturas = dados.faturas || [];
  for (var i = 0; i < faturas.length; i++) {
    var comp0 = vencParaCompetenciaGs(faturas[i].venc, meses);
    if (!comp0) continue;
    var ts = faturas[i].transacoes || [];
    for (var j = 0; j < ts.length; j++) {
      // inclui estornos (negativos) para o total líquido do mês
      trans.push({ comp: comp0, cat: ts[j].categoria || '??', valor: Number(ts[j].valor || 0), cartao: ts[j].cartao });
    }
  }
  if (!trans.length) return '';

  var totalPorComp = {}, pivot = {}, cardsPerComp = {};
  for (var k = 0; k < trans.length; k++) {
    var t = trans[k];
    totalPorComp[t.comp] = (totalPorComp[t.comp] || 0) + t.valor;
    if (!pivot[t.comp]) pivot[t.comp] = {};
    pivot[t.comp][t.cat] = (pivot[t.comp][t.cat] || 0) + t.valor;
    if (!cardsPerComp[t.comp]) cardsPerComp[t.comp] = {};
    cardsPerComp[t.comp][t.cartao] = true;
  }
  var comps = Object.keys(totalPorComp).sort(function(a, b) { return compKeyGs(a, meses).localeCompare(compKeyGs(b, meses)); });
  function nCards(c) { return cardsPerComp[c] ? Object.keys(cardsPerComp[c]).length : 0; }
  var fullCards = 0;
  for (var fc = 0; fc < comps.length; fc++) fullCards = Math.max(fullCards, nCards(comps[fc]));
  function completa(c) { return nCards(c) >= fullCards; }

  // destaca o mês mais recente COMPLETO (evita % enganoso por fatura de cartão faltando)
  var comp = null;
  for (var ci2 = comps.length - 1; ci2 >= 0; ci2--) { if (completa(comps[ci2])) { comp = comps[ci2]; break; } }
  if (!comp) comp = comps[comps.length - 1];
  var idx = comps.indexOf(comp);
  var anteriores = [];
  for (var pa = idx - 1; pa >= 0 && anteriores.length < 3; pa--) { if (completa(comps[pa])) anteriores.unshift(comps[pa]); }

  var linhas = [];
  var totalComp = totalPorComp[comp] || 0;

  if (anteriores.length) {
    var media = 0;
    for (var a = 0; a < anteriores.length; a++) media += totalPorComp[anteriores[a]] || 0;
    media = media / anteriores.length;
    if (media > 0) {
      var pct = Math.round((totalComp - media) / media * 100);
      if (Math.abs(pct) >= 5) linhas.push(comp + ' fechou ' + Math.abs(pct) + '% ' + (pct > 0 ? 'acima' : 'abaixo') + ' da media: ' + fmt(totalComp) + ' (media ' + fmt(media) + ')');
      else linhas.push(comp + ' fechou em linha com a media: ' + fmt(totalComp));
    }
    var aumentos = [];
    var cats = Object.keys(pivot[comp] || {});
    for (var ci = 0; ci < cats.length; ci++) {
      var cat = cats[ci], m = 0;
      for (var ai = 0; ai < anteriores.length; ai++) { var pv = pivot[anteriores[ai]]; m += (pv && pv[cat]) || 0; }
      m = m / anteriores.length;
      var delta = (pivot[comp][cat] || 0) - m;
      if (delta >= 100) aumentos.push({ cat: cat, delta: delta, atual: pivot[comp][cat], media: m });
    }
    aumentos.sort(function(x, y) { return y.delta - x.delta; });
    for (var au = 0; au < Math.min(3, aumentos.length); au++) {
      var d = aumentos[au];
      linhas.push(d.cat + ' subiu ' + fmt(d.delta) + ' vs media (' + fmt(d.media) + ' -> ' + fmt(d.atual) + ')');
    }
  } else {
    linhas.push(comp + ': ' + fmt(totalComp));
  }
  if (!linhas.length) return '';
  var out = 'DESTAQUES DE ' + comp + ':\n';
  for (var li = 0; li < linhas.length; li++) out += '  • ' + linhas[li] + '\n';
  return out + '\n';
}

// ── Email de prévia do acerto (dias 1 e 15, ~7h) ─────────────────

function emailQuinzenal() {
  const hoje = new Date();
  const dia = Number(Utilities.formatDate(hoje, 'America/Sao_Paulo', 'd'));
  if (dia !== 1 && dia !== 15) return;

  const aba = getOrCreateSheet(ABA_DADOS);
  let dados = {};
  try { dados = JSON.parse(lerChunks(aba)); } catch (e) { return; }

  const diaMin = dia === 1 ? 1 : 16;
  const diaMax = dia === 1 ? 15 : 31;
  const rotulo = dia === 1 ? '1 a 15' : '16 a 31';

  const contas = (dados.contasFixas || [])
    .filter(function(c) { return c.dia >= diaMin && c.dia <= diaMax; })
    .sort(function(a, b) { return a.dia - b.dia; });
  const extras = (dados.extras || []).filter(function(x) { return !x.enviado; });

  const fmt = function(v) {
    return 'R$ ' + Number(v || 0).toFixed(2).replace('.', ',').replace(/\B(?=(\d{3})+(?!\d))/g, '.');
  };
  const meses = ['Jan','Fev','Mar','Abr','Mai','Jun','Jul','Ago','Set','Out','Nov','Dez'];
  const mesLabel = meses[hoje.getMonth()] + '/' + hoje.getFullYear();

  let totalContas = 0, totalExtras = 0;
  let corpo = 'Prévia do acerto — ' + mesLabel + ' · Período ' + rotulo + '\n\n';

  // Destaques do mês (resumo automático dos gastos)
  try { corpo += destaquesDoMes(dados, meses, fmt); } catch (e) {}

  corpo += 'CONTAS DO PERÍODO (valores de referência):\n';
  if (!contas.length) corpo += '  (nenhuma conta cadastrada neste período)\n';
  for (let i = 0; i < contas.length; i++) {
    const c = contas[i];
    totalContas += Number(c.valor || 0);
    corpo += '  • ' + c.descricao + ' (dia ' + c.dia + ') — ' + fmt(c.valor) + '\n';
  }
  corpo += '  Subtotal: ' + fmt(totalContas) + '\n\n';

  corpo += 'EXTRAS PENDENTES (você pagou e ainda não cobrou):\n';
  if (!extras.length) corpo += '  (nenhum)\n';
  for (let i = 0; i < extras.length; i++) {
    const x = extras[i];
    totalExtras += Number(x.valor || 0);
    corpo += '  • ' + x.descricao + ' (' + (x.data || '') + ') — ' + fmt(x.valor) + '\n';
  }
  corpo += '  Subtotal: ' + fmt(totalExtras) + '\n\n';

  corpo += 'TOTAL ESTIMADO: ' + fmt(totalContas + totalExtras) + '\n\n';
  corpo += 'Entre no sistema para conferir, ajustar os valores reais e enviar pro Rogério:\n';
  corpo += URL_APP + '\n';

  MailApp.sendEmail({
    to: EMAIL_AVISO,
    subject: '📋 Acerto ' + rotulo + ' de ' + mesLabel + ' — prévia pronta para conferir',
    body: corpo
  });
}

// ── TESTE: envia o email AGORA, ignorando a trava de data ────────
// Selecione esta função no editor e clique em Executar para receber
// um email de teste no seu Gmail. Não afeta o envio automático do dia 1/15.
function enviarEmailTeste() {
  var hoje = new Date();
  var diaHoje = Number(Utilities.formatDate(hoje, 'America/Sao_Paulo', 'd'));
  var dia = diaHoje <= 15 ? 1 : 15;

  var aba = getOrCreateSheet(ABA_DADOS);
  var dados = {};
  try { dados = JSON.parse(lerChunks(aba)); } catch (e) { dados = {}; }

  var diaMin = dia === 1 ? 1 : 16;
  var diaMax = dia === 1 ? 15 : 31;
  var rotulo = dia === 1 ? '1 a 15' : '16 a 31';

  var contas = (dados.contasFixas || [])
    .filter(function(c) { return c.dia >= diaMin && c.dia <= diaMax; })
    .sort(function(a, b) { return a.dia - b.dia; });
  var extras = (dados.extras || []).filter(function(x) { return !x.enviado; });

  var fmt = function(v) {
    return 'R$ ' + Number(v || 0).toFixed(2).replace('.', ',').replace(/\B(?=(\d{3})+(?!\d))/g, '.');
  };
  var meses = ['Jan','Fev','Mar','Abr','Mai','Jun','Jul','Ago','Set','Out','Nov','Dez'];
  var mesLabel = meses[hoje.getMonth()] + '/' + hoje.getFullYear();

  var totalContas = 0, totalExtras = 0;
  var corpo = '[TESTE] Prévia do acerto — ' + mesLabel + ' · Período ' + rotulo + '\n\n';
  try { corpo += destaquesDoMes(dados, meses, fmt); } catch (e) {}

  corpo += 'CONTAS DO PERÍODO (valores de referência):\n';
  if (!contas.length) corpo += '  (nenhuma conta cadastrada neste período)\n';
  for (var i = 0; i < contas.length; i++) { totalContas += Number(contas[i].valor || 0); corpo += '  • ' + contas[i].descricao + ' (dia ' + contas[i].dia + ') — ' + fmt(contas[i].valor) + '\n'; }
  corpo += '  Subtotal: ' + fmt(totalContas) + '\n\n';

  corpo += 'EXTRAS PENDENTES (você pagou e ainda não cobrou):\n';
  if (!extras.length) corpo += '  (nenhum)\n';
  for (var j = 0; j < extras.length; j++) { totalExtras += Number(extras[j].valor || 0); corpo += '  • ' + extras[j].descricao + ' (' + (extras[j].data || '') + ') — ' + fmt(extras[j].valor) + '\n'; }
  corpo += '  Subtotal: ' + fmt(totalExtras) + '\n\n';

  corpo += 'TOTAL ESTIMADO: ' + fmt(totalContas + totalExtras) + '\n\n';
  corpo += 'Este é um email de TESTE. O envio automático continua no dia 1 e 15.\n';
  corpo += URL_APP + '\n';

  MailApp.sendEmail({ to: EMAIL_AVISO, subject: '[TESTE] Acerto ' + rotulo + ' de ' + mesLabel, body: corpo });
}

// ══════════════════════════════════════════════════════════════════
//  BOT DO TELEGRAM — inserir gastos por texto/áudio + consultas
//  Propriedades do script necessárias: BOT_TOKEN, GEMINI_KEY
//  (MEU_CHAT_ID é preenchido sozinho na 1ª mensagem que você mandar.)
//  Depois de configurar o BOT_TOKEN, rode setWebhook() UMA vez.
// ══════════════════════════════════════════════════════════════════

var CARTOES_VALIDOS = ['1895', '8470', '6418', 'NUBANK'];

function tgProp(nome) { return PropertiesService.getScriptProperties().getProperty(nome); }
function geminiModel() { return tgProp('GEMINI_MODEL') || 'gemini-2.0-flash'; }
function fmtReais(v) { return 'R$ ' + Number(v || 0).toFixed(2).replace('.', ',').replace(/\B(?=(\d{3})+(?!\d))/g, '.'); }

// URL /exec pública do web app (a mesma que o app usa). NÃO usar getUrl(): no
// editor ele devolve o /dev, que exige login e o Telegram leva 401.
var WEBHOOK_EXEC_URL = 'https://script.google.com/macros/s/AKfycbwkh0oWPmJE6JbksGgnNJxX6_Y7GTUDeGPlR0xYwQ7vpmHMu7R3LFV2t3_fcErigJFc/exec';

function setWebhook() {
  var token = tgProp('BOT_TOKEN');
  if (!token) { Logger.log('Configure BOT_TOKEN nas Propriedades do script primeiro.'); return; }
  var resp = UrlFetchApp.fetch('https://api.telegram.org/bot' + token + '/setWebhook?url=' + encodeURIComponent(WEBHOOK_EXEC_URL), { muteHttpExceptions: true });
  Logger.log(resp.getContentText());
}
function removeWebhook() {
  var token = tgProp('BOT_TOKEN'); if (!token) return;
  Logger.log(UrlFetchApp.fetch('https://api.telegram.org/bot' + token + '/deleteWebhook', { muteHttpExceptions: true }).getContentText());
}

// ── Chamadas ao Telegram ──
function tgApi(metodo, payload) {
  var token = tgProp('BOT_TOKEN'); if (!token) return null;
  try {
    var resp = UrlFetchApp.fetch('https://api.telegram.org/bot' + token + '/' + metodo, {
      method: 'post', contentType: 'application/json', payload: JSON.stringify(payload), muteHttpExceptions: true
    });
    return JSON.parse(resp.getContentText());
  } catch (e) { return null; }
}
function tgSend(chatId, texto, botoes) {
  var p = { chat_id: chatId, text: texto };
  if (botoes) p.reply_markup = { inline_keyboard: botoes };
  tgApi('sendMessage', p);
}
function tgBaixarArquivoBase64(fileId) {
  var token = tgProp('BOT_TOKEN');
  var info = tgApi('getFile', { file_id: fileId });
  if (!info || !info.ok) return null;
  var resp = UrlFetchApp.fetch('https://api.telegram.org/file/bot' + token + '/' + info.result.file_path, { muteHttpExceptions: true });
  return Utilities.base64Encode(resp.getBlob().getBytes());
}

// ── Entrada principal (chamada pelo doPost) ──
function handleTelegram(update) {
  // Anti-duplicata: o Apps Script responde 302 e o Telegram reenvia a MESMA
  // mensagem várias vezes. Ignoramos update_id já processado (guardado no cache 1h).
  if (update.update_id != null) {
    var cache = CacheService.getScriptCache();
    var ck = 'TGU_' + update.update_id;
    if (cache.get(ck)) return jsonResponse({ ok: true });
    cache.put(ck, '1', 3600);
  }

  var msg = update.message;
  var cb = update.callback_query;
  var chatId = msg ? (msg.chat && msg.chat.id) : (cb && cb.message && cb.message.chat.id);
  if (!chatId) return jsonResponse({ ok: true });

  // Trava de segurança: só o SEU chat. O vínculo exige um código secreto
  // (BOT_SECRET), então ninguém pode "sequestrar" o bot mandando mensagem antes.
  var meu = tgProp('MEU_CHAT_ID');
  if (!meu) {
    var segredo = tgProp('BOT_SECRET');
    var txt0 = (msg && msg.text ? msg.text.trim() : '');
    if (segredo && txt0.indexOf('/bind ') === 0 && txt0.substring(6).trim() === segredo) {
      PropertiesService.getScriptProperties().setProperty('MEU_CHAT_ID', String(chatId));
      tgSend(chatId, '✅ Bot conectado a você! Manda seus gastos por texto ou áudio.\n\nEx.: "gastei 50 na farmácia" · "paguei 200 de faxina pro acerto" · "quanto tá o acerto?"');
    } else if (!segredo) {
      tgSend(chatId, 'Configure a propriedade BOT_SECRET no Apps Script e depois envie: /bind SEUCODIGO');
    } else {
      tgSend(chatId, 'Para ativar, envie: /bind seguido do seu código.');
    }
    return jsonResponse({ ok: true });
  }
  if (String(chatId) !== String(meu)) return jsonResponse({ ok: true }); // ignora estranhos

  try {
    if (cb) return handleTgCallback(cb, chatId);
    if (msg.text) {
      var t = msg.text.trim();
      if (t === '/start' || t === '/id') { tgSend(chatId, 'Pronto pra usar. Seu chat ID: ' + chatId); return jsonResponse({ ok: true }); }
      processarMensagem(chatId, interpretarTexto(t));
    } else if (msg.voice || msg.audio) {
      tgSend(chatId, 'No momento eu entendo só texto 😉. Me escreve, ex.: "gastei 50 na farmácia".');
    } else {
      tgSend(chatId, 'Manda o gasto por texto, ex.: "gastei 50 na farmácia".');
    }
  } catch (err) {
    tgSend(chatId, '⚠ Deu um erro: ' + err.message);
  }
  return jsonResponse({ ok: true });
}

// ── Interpretação por REGRAS (sem IA): extrai valor, tipo e descrição ──
function interpretarTexto(texto) {
  var original = (texto || '').trim();
  var t = original.toLowerCase();

  // Valor: primeiro número (aceita 1.234,56 · 50 · 50,00 · r$ 50 · 50.00)
  var mVal = t.match(/(?:r\$\s*)?(\d{1,3}(?:\.\d{3})+(?:,\d{1,2})?|\d+(?:[.,]\d{1,2})?)/);
  var valor = null;
  if (mVal) {
    var raw = mVal[1];
    if (raw.indexOf(',') >= 0) raw = raw.replace(/\./g, '').replace(',', '.'); // 1.234,56 -> 1234.56
    valor = parseFloat(raw);
  }

  // Sem valor → é uma pergunta (consulta)
  if (valor === null || isNaN(valor) || valor <= 0) {
    return { acao: 'consulta', pergunta: original };
  }

  // Tipo: "acerto"/extra se pagou por fora (pix/dinheiro/débito) ou mencionou
  // rogério/acerto/extra/reembolso; senão é gasto no cartão de crédito
  var acao = /(acerto|rog[eé]rio|reembols|cobrar|\bextra\b|\bpix\b|dinheiro|d[eé]bito|transfer)/.test(t) ? 'acerto' : 'cartao';

  // Cartão citado
  var cartao = '';
  if (/\bnubank\b|\bnu\b/.test(t)) cartao = 'NUBANK';
  else { var mc = t.match(/\b(1895|8470|6418)\b/); if (mc) cartao = mc[1]; }

  // Descrição: remove o valor, moeda, verbos e conectivos
  var desc = original.replace(mVal[0], ' ');
  desc = desc.replace(/r\$/ig, ' ');
  desc = desc.replace(/\b(gastei|paguei|comprei|coloquei|gasto|foi|de|do|da|no|na|em|com|pro|para|pra|reais|real|cart[aã]o|nubank|acerto|rog[eé]rio|reembolso|extra|pix|dinheiro|d[eé]bito|transfer[eê]ncia)\b/ig, ' ');
  desc = desc.replace(/\s+/g, ' ').trim();
  if (!desc) desc = (acao === 'acerto' ? 'gasto do acerto' : 'gasto no cartão');
  desc = desc.charAt(0).toUpperCase() + desc.slice(1);

  return { acao: acao, valor: valor, descricao: desc, cartao: cartao, pergunta: null };
}

// ── Processa a intenção ──
function processarMensagem(chatId, intent) {
  if (!intent || intent.acao === 'desconhecido') {
    tgSend(chatId, 'Não entendi 🤔. Tenta "gastei 50 na farmácia" ou "quanto tá o acerto?".'); return;
  }
  if (intent.acao === 'consulta') { tgSend(chatId, responderConsulta(intent.pergunta || '')); return; }

  var valor = Number(intent.valor);
  if (!valor || valor <= 0 || !intent.descricao) {
    tgSend(chatId, 'Faltou o valor ou a descrição. Ex.: "50 na farmácia".'); return;
  }
  var pend = {
    acao: intent.acao, valor: valor,
    descricao: String(intent.descricao).slice(0, 80),
    cartao: (CARTOES_VALIDOS.indexOf(intent.cartao) >= 0 ? intent.cartao : '')
  };
  // id único por pendência: cada botão confirma o SEU item (não o último digitado)
  var pendId = Utilities.getUuid().replace(/-/g, '').slice(0, 8);
  PropertiesService.getScriptProperties().setProperty('TGPEND_' + chatId + '_' + pendId, JSON.stringify(pend));
  var onde = intent.acao === 'cartao'
    ? ('no cartão' + (pend.cartao ? ' final ' + pend.cartao : ''))
    : 'pro acerto com o Rogério';
  tgSend(chatId, '📝 Anotei: ' + pend.descricao + ' — ' + fmtReais(valor) + ' ' + onde + '.\n\nConfirma?',
    [[ { text: '✅ Confirmar', callback_data: 'ok:' + pendId }, { text: '❌ Cancelar', callback_data: 'no:' + pendId } ]]);
}

function handleTgCallback(cb, chatId) {
  tgApi('answerCallbackQuery', { callback_query_id: cb.id });
  var partes = String(cb.data || '').split(':');
  var acao = partes[0], pendId = partes[1] || '';
  var chave = 'TGPEND_' + chatId + '_' + pendId;
  var raw = tgProp(chave);
  if (!raw) { tgSend(chatId, 'Essa confirmação expirou. Manda o gasto de novo. 🙂'); return jsonResponse({ ok: true }); }
  if (acao !== 'ok') {
    PropertiesService.getScriptProperties().deleteProperty(chave);
    tgSend(chatId, '❌ Cancelado.');
    return jsonResponse({ ok: true });
  }
  // Grava PRIMEIRO; só apaga a pendência se o save deu certo (senão dá pra tentar de novo)
  salvarGastoBot(JSON.parse(raw));
  PropertiesService.getScriptProperties().deleteProperty(chave);
  tgSend(chatId, '✅ Salvo! Já aparece no app.');
  return jsonResponse({ ok: true });
}

// ── Grava no MESMO estado (planilha) que o app usa, com trava de concorrência ──
function salvarGastoBot(p) {
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var aba = getOrCreateSheet(ABA_DADOS);
    var dados = {};
    try { dados = JSON.parse(lerChunks(aba)); } catch (e) { dados = {}; }
    if (!dados.faturas) dados.faturas = []; // nunca deixa a nuvem sem a chave 'faturas'
    var agora = Date.now();
    var hoje = Utilities.formatDate(new Date(), 'America/Sao_Paulo', 'dd/MM/yyyy');
    var rnd = Utilities.getUuid().replace(/-/g, '').slice(0, 8);
    if (p.acao === 'cartao') {
      if (!dados.pendentes) dados.pendentes = [];
      dados.pendentes.push({
        id: 'pd_' + agora + '_tg' + rnd, data: hoje, estabelecimento: p.descricao,
        valor: p.valor, cartao: p.cartao || '', categoria: '??', conciliado: false, updatedAt: agora
      });
    } else {
      if (!dados.extras) dados.extras = [];
      dados.extras.push({ id: 'ex_' + agora + '_tg' + rnd, descricao: p.descricao, valor: p.valor, data: hoje, enviado: false });
    }
    dados.updatedAt = agora;
    salvarChunks(aba, JSON.stringify(dados));
    var cfg = getOrCreateSheet(ABA_CONFIG);
    cfg.getRange('B2').setValue(new Date().toLocaleString('pt-BR', { timeZone: 'America/Sao_Paulo' }));
  } finally {
    lock.releaseLock();
  }
}

// ── Consultas por REGRAS (sem IA): responde perguntas comuns ──
function responderConsulta(pergunta) {
  var aba = getOrCreateSheet(ABA_DADOS);
  var dados = {};
  try { dados = JSON.parse(lerChunks(aba)); } catch (e) { dados = {}; }
  var meses = ['Jan','Fev','Mar','Abr','Mai','Jun','Jul','Ago','Set','Out','Nov','Dez'];
  var f = fmtReais;
  var t = (pergunta || '').toLowerCase();

  // Acerto com o Rogério
  if (/acerto|rog[eé]rio|reembols/.test(t)) {
    var extras = (dados.extras || []).filter(function(x){ return !x.enviado; });
    if (!extras.length) return 'Nenhum extra pendente pro acerto no momento. 🙂';
    var tot = 0; extras.forEach(function(x){ tot += Number(x.valor||0); });
    var msg = 'Acerto pendente: ' + f(tot) + ' (' + extras.length + ' item' + (extras.length > 1 ? 's' : '') + ')\n';
    extras.slice(0, 12).forEach(function(x){ msg += '• ' + x.descricao + ' — ' + f(x.valor) + '\n'; });
    return msg.trim();
  }

  // Gastos de cartão anotados aguardando a fatura
  if (/anotad|aguardando|pendente/.test(t)) {
    var pend = (dados.pendentes || []).filter(function(p){ return !p.conciliado; });
    if (!pend.length) return 'Nenhum gasto de cartão anotado aguardando fatura.';
    var tp = 0; pend.forEach(function(p){ tp += Number(p.valor||0); });
    var m2 = 'Anotados aguardando a fatura: ' + f(tp) + ' (' + pend.length + ')\n';
    pend.slice(0, 12).forEach(function(p){ m2 += '• ' + p.estabelecimento + ' — ' + f(p.valor) + '\n'; });
    return m2.trim();
  }

  // Monta pivot competência × categoria
  var trans = [];
  (dados.faturas || []).forEach(function(fat){
    var comp = vencParaCompetenciaGs(fat.venc, meses);
    if (!comp) return;
    (fat.transacoes || []).forEach(function(tr){ trans.push({ comp: comp, cat: tr.categoria || '??', valor: Number(tr.valor || 0) }); });
  });
  var totalPorComp = {}, pivot = {};
  trans.forEach(function(tr){
    totalPorComp[tr.comp] = (totalPorComp[tr.comp] || 0) + tr.valor;
    (pivot[tr.comp] = pivot[tr.comp] || {})[tr.cat] = (pivot[tr.comp][tr.cat] || 0) + tr.valor;
  });
  var comps = Object.keys(totalPorComp).sort(function(a, b){ return compKeyGs(a, meses).localeCompare(compKeyGs(b, meses)); });
  if (!comps.length) return 'Ainda não tem fatura importada pra eu consultar.';
  var comp = comps[comps.length - 1];

  // Categoria específica citada na pergunta
  var todasCats = (dados.categorias || []).concat(Object.keys(pivot[comp] || {}));
  var achou = null;
  for (var i = 0; i < todasCats.length; i++) {
    var c = todasCats[i];
    if (c && c !== '??' && t.indexOf(c.toLowerCase()) >= 0) { achou = c; break; }
  }
  if (achou) {
    var v = (pivot[comp] && pivot[comp][achou]) || 0;
    var resp = achou + ' em ' + comp + ': ' + f(v);
    if (comps.length > 1) {
      var ant = comps[comps.length - 2];
      var va = (pivot[ant] && pivot[ant][achou]) || 0;
      var d = v - va;
      resp += '\n(' + ant + ' foi ' + f(va) + (Math.abs(d) >= 1 ? ' · ' + (d > 0 ? '+' : '') + f(d) : '') + ')';
    }
    return resp;
  }

  // Total do mês / quanto gastei
  if (/total|quanto|gast|m[eê]s/.test(t)) {
    var resp2 = comp + ': total ' + f(totalPorComp[comp]) + '\n';
    var ord = Object.keys(pivot[comp]).sort(function(a, b){ return pivot[comp][b] - pivot[comp][a]; });
    ord.slice(0, 6).forEach(function(c){ resp2 += '• ' + c + ': ' + f(pivot[comp][c]) + '\n'; });
    return resp2.trim();
  }

  return 'Posso te dizer, por exemplo:\n• "quanto tá o acerto?"\n• "quanto gastei em supermercado?"\n• "qual o total do mês?"\n• "quais gastos anotados?"';
}
