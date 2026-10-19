// ============================================================
// COPARENTALIDAD — Apps Script Central Multi-Cliente
// Con sistema de tokens de sesión para escritura segura
// ============================================================

var SS_ID = '1rj-nUlQ8jTurpj2wNEEAcXLXNh0sr6rpfEpxMtIQ-D8';

// TTL del token en milisegundos (2 horas)
var TOKEN_TTL = 2 * 60 * 60 * 1000;

// Sesiones simultáneas permitidas (Juan, Mariana, tablet, celulares...).
// Antes había UN solo token: cada login pateaba al otro y cada logout
// (automático a los 2 min) cerraba la sesión de todos.
var MAX_SESIONES = 12;

// ============================================================
// CONCURRENCIA — todas las escrituras pasan por un único candado
// ============================================================
// Incidente 19/10/2026: dos guardados casi simultáneos se pisaron. Cada
// request leía la config, la modificaba y la reescribía entera, y setConf
// borraba la fila antes de volver a agregarla: el que leía en ese hueco
// veía "sin cambios" y escribía su único día encima de todo.
// Ahora: un candado por script (LockService) + escritura en el lugar.
var _CANDADO = null;
function withLock(fn) {
  if (_CANDADO) return fn();                    // reentrante dentro de la misma ejecución
  var lock = LockService.getScriptLock();
  try { lock.waitLock(25000); }
  catch (e) { return { error: 'servidor ocupado, reintentá en unos segundos', code: 503, busy: true }; }
  _CANDADO = lock;
  try { return fn(); }
  finally {
    try { SpreadsheetApp.flush(); } catch (x) {}
    _CANDADO = null;
    lock.releaseLock();
  }
}

// Idempotencia: el panel manda un opId por cambio. Si un reintento llega
// cuando el original ya se aplicó, no se vuelve a aplicar ni a registrar.
function opYaAplicada(opId) {
  if (!opId) return false;
  try { return CacheService.getScriptCache().get('op_' + opId) === '1'; } catch (e) { return false; }
}
function marcarOp(opId) {
  if (!opId) return;
  try { CacheService.getScriptCache().put('op_' + opId, '1', 21600); } catch (e) {}
}

// ── Helpers de fecha ─────────────────────────────────────────
function fmtDate(val) {
  if (!val) return '';
  try { return Utilities.formatDate(val, 'America/Argentina/Buenos_Aires', 'yyyy-MM-dd'); }
  catch(e) {
    var s = String(val);
    if (/^\d{4}-\d{2}-\d{2}/.test(s)) return s.substring(0, 10);
    try { return Utilities.formatDate(new Date(s), 'America/Argentina/Buenos_Aires', 'yyyy-MM-dd'); } catch(e2) {}
    return s.substring(0, 10);
  }
}
function hoyAR() { return Utilities.formatDate(new Date(), 'America/Argentina/Buenos_Aires', 'yyyy-MM-dd'); }
function mesAR()  { return Utilities.formatDate(new Date(), 'America/Argentina/Buenos_Aires', 'yyyy-MM'); }

// ============================================================
// ENTRADA PRINCIPAL
// ============================================================
// Acciones que sólo leen: no toman el candado.
var SOLO_LECTURA = { getAll: 1, getEvents: 1, debug: 1, getResumen: 1 };

function doGet(e) {
  var p   = e.parameter || {};
  var res;
  try {
    var ss       = SpreadsheetApp.openById(SS_ID);
    var clientId = p.clientId || '';
    var mes      = p.month    || mesAR();
    var act      = p.action   || 'getAll';

    if (!clientId) return jsonp(p.callback, { error: 'clientId requerido' });

    if (SOLO_LECTURA[act]) res = rutear(ss, p, act, clientId, mes);
    else res = withLock(function() { return rutear(ss, p, act, clientId, mes); });
  } catch (err) {
    res = { error: err.toString() };
  }
  return jsonp(p.callback, res);
}

function rutear(ss, p, act, clientId, mes) {
  // ── Acciones públicas (no requieren token) ────────────────
  // Lectura y carga de datos cotidianos: sin friccion, sin PIN.
  // El PIN/token solo protege acciones sensibles (presupuesto,
  // calendario, configuracion), no la carga de gastos ni la lista.
  if (act === 'getAll')         return getAll(ss, clientId, mes);
  if (act === 'validatePin')    return validatePin(ss, clientId, p.pin || '');
  if (act === 'changePin')      return changePin(ss, clientId, p.actual || '', p.nuevo || '');
  if (act === 'debug')          return debug(ss, clientId, mes);
  if (act === 'addExp')         return addExp(ss, clientId, p);
  if (act === 'delExp')         return delExp(ss, clientId, p.id);
  if (act === 'saveShopping')   return saveShopping(ss, clientId, p.list || '[]');
  if (act === 'getEvents')      return getCalendarEvents(p.month || mes);
  if (act === 'saveEventExtra') return saveEventExtra(ss, clientId, p);
  if (act === 'getResumen')     return getResumen(ss, clientId, p.month || mes);

  // ── Acciones sensibles (requieren token válido) ───────────
  var tokenCheck = checkToken(ss, clientId, p.token || '');
  if (!tokenCheck.ok) return { error: 'sesion invalida', code: 401 };

  // quién hace el cambio, según el token (owner / ex)
  var actorMode = tokenCheck.mode ? tokenCheck.mode : 'owner';

  var opId = /^[A-Za-z0-9_-]{1,64}$/.test('' + (p.opId || '')) ? '' + p.opId : '';
  if (opId && opYaAplicada(opId)) return { ok: true, dup: true };

  var res;
  if      (act === 'clearMonth')       res = clearMonth(ss, clientId, p.month || mes);
  else if (act === 'setBudget')        res = setConf(ss, clientId, 'budget', p.amount);
  else if (act === 'setBudgetType')    res = setBudgetType(ss, clientId, p.value);
  else if (act === 'setOverride')      res = setOverride(ss, clientId, p.date, p.who, actorMode);
  else if (act === 'setOverridesBulk') res = setOverridesBulk(ss, clientId, p.changes, actorMode);
  else if (act === 'applyAgreement')   res = applyAgreement(ss, clientId, p, actorMode);
  else if (act === 'resetOverrides')   res = resetOverrides(ss, clientId, actorMode);
  else if (act === 'addRem')           res = modRems(ss, clientId, 'add', p);
  else if (act === 'delRem')           res = modRems(ss, clientId, 'del', p);
  else if (act === 'setShareShop')     res = setConf(ss, clientId, 'share_shopping', p.value === 'true' ? 'true' : 'false');
  else if (act === 'setPanelUi')       res = setPanelUi(ss, clientId, p);
  else if (act === 'logout')           res = revokeToken(ss, clientId, p.token || '');
  else                                 res = { error: 'accion desconocida: ' + act };

  if (opId && res && !res.error) marcarOp(opId);
  return res;
}

function jsonp(callback, data) {
  var out = JSON.stringify(data);
  if (callback) {
    return ContentService
      .createTextOutput(callback + '(' + out + ')')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService
    .createTextOutput(out)
    .setMimeType(ContentService.MimeType.JSON);
}

// ============================================================
// SISTEMA DE TOKENS
// ============================================================

// Genera un token aleatorio de 32 chars
function generateToken() {
  var chars = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789';
  var token = '';
  for (var i = 0; i < 32; i++) {
    token += chars.charAt(Math.floor(Math.random() * chars.length));
  }
  return token;
}

// Sesiones: mapa { token: {mode, expiresAt, createdAt} } en la clave
// 'session_tokens'. Varias sesiones conviven (owner y ex a la vez).
function leerSesiones(cfg) {
  var m = {};
  try { m = JSON.parse(cfg.session_tokens || '{}') || {}; } catch (e) { m = {}; }
  if (typeof m !== 'object' || m instanceof Array) m = {};
  return m;
}

// Crea un token nuevo para el cliente (sin pisar las otras sesiones)
function createToken(ss, clientId, mode) {
  return withLock(function() {
    var cfg  = readConfig(ss, clientId);
    var m    = leerSesiones(cfg);
    var now  = Date.now();
    var keys = Object.keys(m);
    for (var i = 0; i < keys.length; i++) {
      if (!m[keys[i]] || !(m[keys[i]].expiresAt > now)) delete m[keys[i]];   // vencidas
    }
    var token     = generateToken();
    var expiresAt = now + TOKEN_TTL;
    m[token] = { mode: mode, expiresAt: expiresAt, createdAt: now };
    keys = Object.keys(m).sort(function(a, b) { return (m[a].createdAt || 0) - (m[b].createdAt || 0); });
    while (keys.length > MAX_SESIONES) delete m[keys.shift()];               // tope: se van las más viejas
    setConf(ss, clientId, 'session_tokens', JSON.stringify(m));
    return { ok: true, token: token, mode: mode, expiresAt: expiresAt };
  });
}

// Verifica que el token sea válido y no haya expirado
function checkToken(ss, clientId, token) {
  if (!token) return { ok: false, error: 'token requerido' };
  var cfg = readConfig(ss, clientId);
  var s   = leerSesiones(cfg)[token];
  if (s) {
    if (Date.now() > s.expiresAt) return { ok: false, error: 'sesion expirada' };
    return { ok: true, mode: s.mode };
  }
  // Compatibilidad: sesión abierta con la versión anterior (un solo token)
  if (cfg.session_token) {
    try {
      var d = JSON.parse(cfg.session_token);
      if (d && d.token === token && Date.now() <= d.expiresAt) return { ok: true, mode: d.mode };
    } catch (e) {}
  }
  return { ok: false, error: 'sesion invalida' };
}

// Invalida SOLO el token de quien cierra sesión (antes cerraba la de todos)
function revokeToken(ss, clientId, token) {
  return withLock(function() {
    if (!token) return { ok: true };
    var cfg = readConfig(ss, clientId);
    var m   = leerSesiones(cfg);
    if (m[token]) { delete m[token]; setConf(ss, clientId, 'session_tokens', JSON.stringify(m)); }
    if (cfg.session_token) {
      try {
        var d = JSON.parse(cfg.session_token);
        if (d && d.token === token) setConf(ss, clientId, 'session_token', '{}');
      } catch (e) {}
    }
    return { ok: true };
  });
}

// ============================================================
// HOJAS
// ============================================================
function getGastos(ss) {
  var sh = ss.getSheetByName('gastos');
  if (!sh) {
    sh = ss.insertSheet('gastos');
    sh.appendRow(['id','date','cat','amt','desc','ts','payment','cuotas','clientId','moneda']);
  }
  return sh;
}

function getConfigSheet(ss) {
  var sh = ss.getSheetByName('config');
  if (!sh) sh = ss.insertSheet('config');
  return sh;
}

function getClientesSheet(ss) {
  var sh = ss.getSheetByName('clientes');
  if (!sh) {
    sh = ss.insertSheet('clientes');
    sh.appendRow(['clientId','name','email','createdAt','active','telegramChatId','telegramToken']);
  }
  return sh;
}

// ============================================================
// CONFIG
// ============================================================
function readConfig(ss, clientId) {
  var cs  = getConfigSheet(ss);
  var cfg = {};
  if (cs.getLastRow() < 1) return cfg;
  var data = cs.getDataRange().getValues();
  for (var i = 0; i < data.length; i++) {
    if ('' + data[i][0] === '' + clientId && data[i][1]) {
      cfg['' + data[i][1]] = data[i][2];
    }
  }
  return cfg;
}

// Escribe una clave de config EN SU LUGAR (nunca borra y re-agrega la fila:
// ese hueco era el que dejaba a otro request leyendo "vacío").
// Si quedaron filas duplicadas de la misma clave, deja sólo la vigente.
function setConf(ss, clientId, key, val) {
  return withLock(function() {
    var cs   = getConfigSheet(ss);
    var last = cs.getLastRow();
    var data = last ? cs.getRange(1, 1, last, 2).getValues() : [];
    var keep = -1, dups = [];
    for (var i = 0; i < data.length; i++) {
      if ('' + data[i][0] === '' + clientId && '' + data[i][1] === '' + key) {
        if (keep >= 0) dups.push(keep);
        keep = i;                                   // la última es la que lee readConfig
      }
    }
    if (keep >= 0) {
      cs.getRange(keep + 1, 3).setValue(val);
      for (var j = dups.length - 1; j >= 0; j--) cs.deleteRow(dups[j] + 1);   // de abajo hacia arriba
    } else {
      cs.appendRow([clientId, key, val]);
    }
    return { ok: true };
  });
}

// Tipo de presupuesto: contra que gastos se mide el tope mensual.
// Se valida contra la lista para no guardar basura desde la URL.
var BUDGET_TYPES = { efectivo: 1, digital: 1, todo: 1 };
function setBudgetType(ss, clientId, value) {
  var v = '' + (value || '');
  if (!BUDGET_TYPES[v]) return { error: 'tipo invalido: ' + v };
  setConf(ss, clientId, 'budget_type', v);
  return { ok: true, budgetType: v };
}

function ensureConf(ss, clientId, key, defVal) {
  return withLock(function() {
    var cs   = getConfigSheet(ss);
    var data = cs.getDataRange().getValues();
    for (var i = 0; i < data.length; i++) {
      if ('' + data[i][0] === '' + clientId && '' + data[i][1] === '' + key) return;
    }
    cs.appendRow([clientId, key, defVal]);
  });
}

// ============================================================
// GET ALL (público — no requiere token)
// ============================================================
function getAll(ss, clientId, month) {
  var res = {
    expenses: [], budget: 3000000, budgetType: 'efectivo', overrides: {}, cycleAnchor: null,
    reminders: [], shoppingList: [], shareShopping: true,
    panelUi: defaultPanelUi()
  };

  // Gastos
  var gs = ss.getSheetByName('gastos');
  if (gs && gs.getLastRow() > 1) {
    var ncols = gs.getLastColumn();
    var rows  = gs.getRange(2, 1, gs.getLastRow() - 1, ncols).getValues();
    for (var i = 0; i < rows.length; i++) {
      var r = rows[i];
      if (!r[0]) continue;
      var rowClient = ncols >= 9 ? '' + r[8] : '';
      if (rowClient && rowClient !== '' + clientId) continue;
      var d = fmtDate(r[1]);
      if (!month || d.substring(0, 7) === month) {
        res.expenses.push({
          id:      '' + r[0],
          date:    d,
          cat:     '' + r[2],
          amt:     +r[3],
          desc:    '' + (r[4] || ''),
          payment: (ncols > 6 && r[6]) ? '' + r[6] : 'efectivo',
          cuotas:  (ncols > 7 && r[7]) ? +r[7]     : 1,
          moneda:  (ncols > 9 && r[9]) ? '' + r[9] : 'ARS'
        });
      }
    }
  }

  // Config
  var cfg = readConfig(ss, clientId);
  if (cfg.budget)         res.budget        = +cfg.budget || 3000000;
  if (cfg.budget_type && BUDGET_TYPES['' + cfg.budget_type]) res.budgetType = '' + cfg.budget_type;
  if (cfg.overrides)      { try { res.overrides     = JSON.parse(cfg.overrides)     || {}; } catch(x){} }
  if (cfg.cycle_anchor)   { try { var _a = JSON.parse(cfg.cycle_anchor); res.cycleAnchor = (_a && _a.date) ? _a : null; } catch(x){ res.cycleAnchor = null; } }
  if (cfg.reminders)      { try { res.reminders     = JSON.parse(cfg.reminders)     || []; } catch(x){} }
  if (cfg.shopping_list)  { try { res.shoppingList  = JSON.parse(cfg.shopping_list) || []; } catch(x){} }
  res.shareShopping = (cfg.share_shopping !== undefined) ? ('' + cfg.share_shopping) === 'true' : true;
  if (cfg.panel_ui)       { try { res.panelUi = mergePanelUi(JSON.parse(cfg.panel_ui)); } catch(x){} }
  // Historial: hoja 'historial' (una fila por cambio, sin tope). Si todavía
  // no existe, se lee la clave vieja de config.
  var hs = readHistorial(ss, clientId, HIST_MAX_GETALL);
  if (hs) res.overrideHistory = hs;
  else if (cfg.override_history) { try { res.overrideHistory = JSON.parse(cfg.override_history) || []; } catch(x){ res.overrideHistory = []; } }
  else res.overrideHistory = [];
  if (cfg.event_extras)   { try { res.eventExtras = JSON.parse(cfg.event_extras) || {}; } catch(x){ res.eventExtras = {}; } }
  else res.eventExtras = {};

  return res;
}

// ============================================================
// GASTOS
// ============================================================
function addExp(ss, clientId, p) {
  var gs = getGastos(ss);
  var id = '' + Date.now();
  gs.appendRow([id, p.date || hoyAR(), p.cat || 'otros', +p.amt || 0,
                p.desc || '', Date.now(), p.payment || 'efectivo', +p.cuotas || 1, clientId,
                (p.moneda === 'USD' ? 'USD' : 'ARS')]);
  SpreadsheetApp.flush();
  return { ok: true, id: id };
}

function delExp(ss, clientId, id) {
  var gs = ss.getSheetByName('gastos');
  if (!gs || gs.getLastRow() < 2) return { ok: true };
  var data = gs.getDataRange().getValues();
  for (var i = data.length - 1; i >= 1; i--) {
    if ('' + data[i][0] === '' + id) { gs.deleteRow(i + 1); }
  }
  SpreadsheetApp.flush();
  return { ok: true };
}

function clearMonth(ss, clientId, month) {
  var gs = ss.getSheetByName('gastos');
  if (!gs || gs.getLastRow() < 2) return { ok: true };
  var data  = gs.getDataRange().getValues();
  var ncols = gs.getLastColumn();
  for (var i = data.length - 1; i >= 1; i--) {
    var rowClient = ncols >= 9 ? '' + data[i][8] : '';
    var rowDate   = fmtDate(data[i][1]).substring(0, 7);
    if (rowDate === month && (!rowClient || rowClient === '' + clientId)) {
      gs.deleteRow(i + 1);
    }
  }
  SpreadsheetApp.flush();
  return { ok: true };
}

// ============================================================
// CONFIG HELPERS
// ============================================================
function setOverride(ss, clientId, date, who, actorMode) {
  if (!/^\d{4}-\d{2}-\d{2}$/.test('' + date)) return { error: 'fecha invalida' };
  if (who !== 'papa' && who !== 'mama')     return { error: 'valor invalido: ' + who };
  var cfg = readConfig(ss, clientId);
  var ov  = {};
  try { ov = JSON.parse(cfg.overrides || '{}') || {}; } catch(x) {}
  var prev = ov[date] || '';
  ov[date] = who;
  setConf(ss, clientId, 'overrides', JSON.stringify(ov));
  registrarHistorial(ss, clientId, date, prev, who, actorMode);
  return { ok: true };
}

// Aplica varios cambios de una (feature "aplicar acuerdo desde un día")
function setOverridesBulk(ss, clientId, changesJson, actorMode) {
  var changes;
  try { changes = JSON.parse(changesJson || '[]'); } catch(e) { return { error: 'JSON invalido' }; }
  if (!Array.isArray(changes)) return { error: 'se esperaba un array' };
  var cfg = readConfig(ss, clientId);
  var ov  = {};
  try { ov = JSON.parse(cfg.overrides || '{}') || {}; } catch(x) {}
  var n = 0;
  for (var i = 0; i < changes.length; i++) {
    var c = changes[i];
    if (!c || !c.date || !c.who) continue;
    ov[c.date] = c.who;
    n++;
  }
  setConf(ss, clientId, 'overrides', JSON.stringify(ov));
  // Una sola entrada de historial por operacion masiva: escribir una por dia
  // desbordaba el tope de 200 y borraba el historico previo.
  if (n) registrarHistorial(ss, clientId, '#bulk', '', '' + n, actorMode, true);
  return { ok: true, applied: n };
}

// Aplica el acuerdo base desde una fecha, re-anclando el ciclo de 14 dias.
// - from: 'YYYY-MM-DD', primer dia del acuerdo
// - who:  'papa' | 'mama', con quien arranca ese primer dia
// - drop: JSON array de fechas cuyo cambio manual se descarta (vuelven al acuerdo)
// Nunca toca fechas anteriores a 'from': el historico queda intacto.
function applyAgreement(ss, clientId, p, actorMode) {
  var from = '' + (p.from || '');
  if (!/^\d{4}-\d{2}-\d{2}$/.test(from)) return { error: 'fecha invalida' };
  var who = (p.who === 'papa') ? 'papa' : 'mama';

  var drop = [];
  try { drop = JSON.parse(p.drop || '[]') || []; } catch (e) { drop = []; }
  if (!(drop instanceof Array)) drop = [];

  var cfg = readConfig(ss, clientId);
  var ov  = {};
  try { ov = JSON.parse(cfg.overrides || '{}') || {}; } catch (x) {}

  var removed = 0;
  for (var i = 0; i < drop.length; i++) {
    var dk = '' + drop[i];
    if (!/^\d{4}-\d{2}-\d{2}$/.test(dk)) continue;
    if (dk < from) continue;                       // blindaje: jamas hacia atras
    if (ov.hasOwnProperty(dk)) { delete ov[dk]; removed++; }
  }

  setConf(ss, clientId, 'overrides', JSON.stringify(ov));
  setConf(ss, clientId, 'cycle_anchor', JSON.stringify({ date: from, who: who }));

  // nolog='1' lo usan los envios por tandas para no duplicar el historial
  if (('' + (p.nolog || '')) !== '1') {
    registrarHistorial(ss, clientId, '#agreement', '', who, actorMode, true, from, removed);
  }
  return { ok: true, removed: removed, anchor: { date: from, who: who } };
}

function resetOverrides(ss, clientId, actorMode) {
  setConf(ss, clientId, 'overrides', '{}');
  setConf(ss, clientId, 'cycle_anchor', '');   // vuelve al esquema original
  registrarHistorial(ss, clientId, '*', '', 'reset', actorMode);
  return { ok: true };
}

// ============================================================
// HISTORIAL DE CAMBIOS DE CUSTODIA — hoja 'historial'
// ============================================================
// Una fila por cambio, sólo se agrega (append), nunca se reescribe.
// Antes era un JSON en una celda de config con tope de 200: cualquier
// carrera lo reescribía entero y por eso el panel mostraba sólo 2.
var HIST_COLS = ['ts','clientId','date','prev','who','by','at','bulk','from','removed','note'];
var HIST_MAX_GETALL = 1000;

function getHistSheet(ss) {
  var sh = ss.getSheetByName('historial');
  if (!sh) {
    sh = ss.insertSheet('historial');
    var head = sh.getRange(1, 1, 1, HIST_COLS.length);
    head.setNumberFormat('@');
    head.setValues([HIST_COLS]);
    sh.setFrozenRows(1);
  }
  return sh;
}

// Agrega filas como texto plano (sin que Sheets convierta '2026-11-18' en fecha)
function appendHistRows(sh, rows) {
  if (!rows || !rows.length) return;
  var start = sh.getLastRow() + 1;
  var need  = start + rows.length - 1;
  if (need > sh.getMaxRows()) sh.insertRowsAfter(sh.getMaxRows(), need - sh.getMaxRows());
  var rng = sh.getRange(start, 1, rows.length, HIST_COLS.length);
  rng.setNumberFormat('@');
  rng.setValues(rows);
}

function textoHist(v) {
  if (v instanceof Date) return fmtDate(v);
  return (v === null || v === undefined) ? '' : '' + v;
}

// Entradas del cliente ordenadas por momento de carga (ts).
// null = la hoja todavía no existe (getAll usa la clave vieja de config).
function readHistorial(ss, clientId, limit) {
  var sh = ss.getSheetByName('historial');
  if (!sh) return null;
  var last = sh.getLastRow();
  if (last < 2) return [];
  var rows = sh.getRange(2, 1, last - 1, HIST_COLS.length).getValues();
  var out = [];
  for (var i = 0; i < rows.length; i++) {
    var r = rows[i];
    if (textoHist(r[1]) !== '' + clientId) continue;
    var e = {
      ts:   +r[0] || 0,
      date: textoHist(r[2]),
      prev: textoHist(r[3]),
      who:  textoHist(r[4]),
      by:   textoHist(r[5]) === 'ex' ? 'ex' : 'owner',
      at:   textoHist(r[6]),
      bulk: textoHist(r[7]) === '1' ? 1 : 0
    };
    if (textoHist(r[8]) !== '')  e.from    = textoHist(r[8]);
    if (textoHist(r[9]) !== '')  e.removed = +r[9] || 0;
    if (textoHist(r[10]) !== '') e.note    = textoHist(r[10]);
    out.push(e);
  }
  out.sort(function(a, b) { return a.ts - b.ts; });
  if (limit && out.length > limit) out = out.slice(out.length - limit);
  return out;
}

// Registra una línea en el historial de cambios de custodia
function registrarHistorial(ss, clientId, date, prev, who, actorMode, bulk, from, removed, note) {
  return withLock(function() {
    appendHistRows(getHistSheet(ss), [[
      '' + Date.now(), '' + clientId, '' + date, '' + (prev || ''), '' + (who || ''),
      actorMode === 'ex' ? 'ex' : 'owner', hoyAR(), bulk ? '1' : '0',
      (from === undefined || from === null) ? '' : '' + from,
      (removed === undefined || removed === null) ? '' : '' + removed,
      note ? '' + note : ''
    ]]);
    return { ok: true };
  });
}

// ============================================================
// CALENDARIO FAMILIA (Google Calendar) — lectura pública
// ============================================================
// Lee el calendario llamado "Familia" de la cuenta que corre el script.
function getCalendarEvents(month) {
  var out = [];
  try {
    var cals = CalendarApp.getCalendarsByName('Familia');
    if (!cals || cals.length === 0) return { events: [], warn: 'no se encontro calendario Familia' };
    var cal = cals[0];
    // rango del mes (month = 'YYYY-MM')
    var parts = month.split('-');
    var y = parseInt(parts[0], 10);
    var m = parseInt(parts[1], 10) - 1;
    var start = new Date(y, m, 1, 0, 0, 0);
    var end   = new Date(y, m + 1, 1, 0, 0, 0);
    var evs = cal.getEvents(start, end);
    for (var i = 0; i < evs.length; i++) {
      var ev = evs[i];
      var s  = ev.getStartTime();
      var allDay = ev.isAllDayEvent();
      var dstr = Utilities.formatDate(s, 'America/Argentina/Buenos_Aires', 'yyyy-MM-dd');
      var baseId = ev.getId();
      // Clave única POR OCURRENCIA: los eventos recurrentes comparten getId(),
      // así que le sumamos la fecha para poder configurar cada día por separado.
      out.push({
        id:     baseId + '|' + dstr,
        baseId: baseId,
        title:  ev.getTitle(),
        date:   dstr,
        time:   allDay ? '' : Utilities.formatDate(s, 'America/Argentina/Buenos_Aires', 'HH:mm'),
        allDay: allDay ? 1 : 0
      });
    }
  } catch (e) {
    return { events: [], error: e.toString() };
  }
  return { events: out };
}

// Guarda la info extra de un evento (qué nena, quién la lleva, etc.)
// Es público: ambos padres pueden coordinar. Clave = id del evento.
function saveEventExtra(ss, clientId, p) {
  var cfg = readConfig(ss, clientId);
  var extras = {};
  try { extras = JSON.parse(cfg.event_extras || '{}') || {}; } catch(x) {}
  var key = p.eventId || '';
  if (!key) return { error: 'falta eventId' };
  extras[key] = {
    kid:           p.kid     || '',
    takenBy:       p.takenBy || '',
    takenByName:   p.takenByName || '',
    otherWith:     p.otherWith || '',
    otherWithName: p.otherWithName || '',
    note:          p.note    || ''
  };
  setConf(ss, clientId, 'event_extras', JSON.stringify(extras));
  return { ok: true };
}

function modRems(ss, clientId, mode, p) {
  var cfg  = readConfig(ss, clientId);
  var rems = [];
  try { rems = JSON.parse(cfg.reminders || '[]') || []; } catch(x) {}
  if (mode === 'add') rems.push({ time: p.time, msg: p.msg, lastFired: '' });
  if (mode === 'del') rems.splice(+p.idx, 1);
  return setConf(ss, clientId, 'reminders', JSON.stringify(rems));
}

function saveShopping(ss, clientId, listJson) {
  var arr;
  try { arr = JSON.parse(listJson); } catch(e) { return { error: 'JSON invalido: ' + e.toString() }; }
  if (!Array.isArray(arr)) return { error: 'se esperaba un array' };
  return setConf(ss, clientId, 'shopping_list', JSON.stringify(arr));
}

// ============================================================
// PANEL UI
// ============================================================
function defaultPanelUi() {
  return {
    papaName: 'Papa', mamaName: 'Mama',
    kidsTemplate: 'los chicos', kidsTemplateSingular: 'el/la chico/a',
    kidsCount: 2, country: 'AR',
    helperName: '', ownerHouse: '', exHouse: '',
    hasHelper: false
  };
}

function mergePanelUi(raw) {
  var base = defaultPanelUi();
  if (!raw || typeof raw !== 'object') return base;
  for (var k in raw) {
    if (raw.hasOwnProperty(k) && raw[k] !== undefined && raw[k] !== '') base[k] = raw[k];
  }
  return base;
}

function setPanelUi(ss, clientId, p) {
  var cfg = readConfig(ss, clientId);
  var ui  = defaultPanelUi();
  try { ui = mergePanelUi(JSON.parse(cfg.panel_ui || '{}')); } catch(x) {}
  var fields = ['papaName','mamaName','kidsTemplate','kidsTemplateSingular',
                'kidsCount','country','helperName','ownerHouse','exHouse'];
  for (var i = 0; i < fields.length; i++) {
    if (p[fields[i]] !== undefined && p[fields[i]] !== '') {
      ui[fields[i]] = fields[i] === 'kidsCount' ? +p[fields[i]] : p[fields[i]];
    }
  }
  if (p.hasHelper !== undefined) ui.hasHelper = p.hasHelper === 'true' || p.hasHelper === true;
  return setConf(ss, clientId, 'panel_ui', JSON.stringify(ui));
}

// ============================================================
// PIN Y AUTENTICACIÓN
// ============================================================
function hashPin(pin, role) {
  var s = '' + pin + '|' + role + '|pj2026';
  var h = 0;
  for (var i = 0; i < s.length; i++) h = ((h << 5) - h + s.charCodeAt(i)) | 0;
  return (h >>> 0).toString(36);
}

// validatePin: verifica PIN y devuelve token si es correcto
function validatePin(ss, clientId, pin) {
  if (!pin || ('' + pin).length !== 4) return { error: 'PIN invalido' };
  var cfg    = readConfig(ss, clientId);
  var ownerH = '' + (cfg.pin_owner || '');
  var exH    = '' + (cfg.pin_ex    || '');
  if (!ownerH && !exH) return { error: 'PINs no configurados' };
  var p = '' + pin;
  var mode = null;
  if (ownerH && hashPin(p, 'owner') === ownerH) mode = 'full';
  else if (exH && hashPin(p, 'ex')  === exH)    mode = 'ex';
  if (!mode) return { ok: false, error: 'PIN incorrecto' };
  // Crear token de sesión
  var session = createToken(ss, clientId, mode);
  return { ok: true, mode: mode, token: session.token, expiresAt: session.expiresAt };
}

// changePin: el DUEÑO cambia su PIN desde el panel (pide el PIN actual).
// El PIN nuevo sólo viaja al servidor, que guarda el hash: nunca queda en el
// código ni en el repo. El PIN del otro progenitor no se toca.
function changePin(ss, clientId, actual, nuevo) {
  actual = '' + actual; nuevo = '' + nuevo;
  if (!/^\d{4}$/.test(nuevo)) return { error: 'el PIN nuevo tiene que tener 4 dígitos' };
  var cfg    = readConfig(ss, clientId);
  var ownerH = '' + (cfg.pin_owner || '');
  if (!ownerH || hashPin(actual, 'owner') !== ownerH) return { error: 'PIN actual incorrecto' };
  if (nuevo === actual) return { error: 'el PIN nuevo es igual al actual' };
  if (cfg.pin_ex && hashPin(nuevo, 'ex') === '' + cfg.pin_ex) return { error: 'ese PIN es el del otro progenitor: elegí otro' };
  var h = hashPin(nuevo, 'owner');
  // Forzar texto si el hash pudiera leerse como número o fecha en Sheets
  setConf(ss, clientId, 'pin_owner', /^[0-9eE.+\-\/]+$/.test(h) ? "'" + h : h);
  // Cerrar las sesiones de dueño abiertas con el PIN viejo
  var m = leerSesiones(cfg), cerradas = 0;
  for (var k in m) { if (m.hasOwnProperty(k) && m[k] && m[k].mode === 'full') { delete m[k]; cerradas++; } }
  setConf(ss, clientId, 'session_tokens', JSON.stringify(m));
  if (cfg.session_token) {
    try { var d = JSON.parse(cfg.session_token); if (d && d.mode === 'full') setConf(ss, clientId, 'session_token', '{}'); } catch (e) {}
  }
  return { ok: true, sesionesCerradas: cerradas };
}

// ============================================================
// REGISTRO DE CLIENTES
// ============================================================
function registerClient(ss, p) {
  var clientId = p.clientId || '';
  var name     = p.name     || '';
  var email    = p.email    || '';
  if (!clientId || !name || !email) return { error: 'clientId, name y email requeridos' };

  var cs   = getClientesSheet(ss);
  var data = cs.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if ('' + data[i][0] === '' + clientId) return { error: 'clientId ya existe' };
  }
  cs.appendRow([clientId, name, email, hoyAR(), 'true', '', '']);

  var ownerPin = p.ownerPin || '0000';
  var exPin    = p.exPin    || '0000';
  var budget   = p.budget   || '3000000';
  var panelUi  = p.panelUi  || JSON.stringify(defaultPanelUi());

  var cs2 = getConfigSheet(ss);
  cs2.appendRow([clientId, 'budget',         budget]);
  cs2.appendRow([clientId, 'overrides',      '{}']);
  cs2.appendRow([clientId, 'reminders',      '[]']);
  cs2.appendRow([clientId, 'shopping_list',  '[]']);
  cs2.appendRow([clientId, 'share_shopping', 'true']);
  cs2.appendRow([clientId, 'panel_ui',       panelUi]);
  cs2.appendRow([clientId, 'pin_owner',      hashPin(ownerPin, 'owner')]);
  cs2.appendRow([clientId, 'pin_ex',         hashPin(exPin, 'ex')]);

  SpreadsheetApp.flush();
  return { ok: true, clientId: clientId };
}

// ============================================================
// DEBUG
// ============================================================
function debug(ss, clientId, mes) {
  var gs  = ss.getSheetByName('gastos');
  var cfg = readConfig(ss, clientId);
  return {
    clientId:      clientId,
    gastosExiste:  !!gs,
    gastosFilas:   gs ? gs.getLastRow() - 1 : 0,
    configClaves:  Object.keys(cfg),
    budget:        cfg.budget || 'no seteado',
    shoppingItems: cfg.shopping_list ? JSON.parse(cfg.shopping_list || '[]').length : 0,
    mes:           mes
  };
}

// ============================================================
// SETUP / MIGRACIÓN
// ============================================================
function runSetup() {
  var ss       = SpreadsheetApp.openById(SS_ID);
  var clientId = 'juanma';

  var gs = ss.getSheetByName('gastos');
  if (!gs) {
    gs = ss.insertSheet('gastos');
    gs.appendRow(['id','date','cat','amt','desc','ts','payment','cuotas','clientId','moneda']);
    Logger.log('gastos: creada');
  }

  ensureConf(ss, clientId, 'budget',         '3000000');
  ensureConf(ss, clientId, 'overrides',      '{}');
  ensureConf(ss, clientId, 'reminders',      '[]');
  ensureConf(ss, clientId, 'shopping_list',  '[]');
  ensureConf(ss, clientId, 'share_shopping', 'true');
  ensureConf(ss, clientId, 'panel_ui',       JSON.stringify(defaultPanelUi()));

  var cfg = readConfig(ss, clientId);
  if (!cfg.pin_owner) {
    // Instalación nueva: completá los PIN iniciales SÓLO en el editor (nunca en
    // el repo, que es público), corré runSetup y volvé a dejarlos vacíos.
    // Después, el dueño cambia el suyo desde el panel (Cambiar mi PIN).
    var PIN_DUENO_INICIAL = '', PIN_EX_INICIAL = '';
    if (PIN_DUENO_INICIAL && PIN_EX_INICIAL) {
      setConf(ss, clientId, 'pin_owner', hashPin(PIN_DUENO_INICIAL, 'owner'));
      setConf(ss, clientId, 'pin_ex',    hashPin(PIN_EX_INICIAL, 'ex'));
      Logger.log('PINs inicializados');
    } else {
      Logger.log('Faltan los PIN iniciales: completalos en runSetup (sólo en el editor).');
    }
  }

  SpreadsheetApp.flush();
  Logger.log('Setup OK — claves: ' + Object.keys(readConfig(ss, clientId)).join(', '));
}

// (19/10/2026) Se eliminó reconstruirConfig(): borraba toda la hoja 'config' y
// la reescribía con datos de junio, incluido el PIN de dueño viejo. Si hay que
// recuperar config, usar el historial de versiones del Sheet.

// Prueba manual de tokens: llamala desde otra función pasando un PIN válido,
// p. ej. function probar() { runTestToken('····'); } — y borrala después.
function runTestToken(pin) {
  var ss  = SpreadsheetApp.openById(SS_ID);
  Logger.log('=== TEST TOKENS ===');
  if (!pin) { Logger.log('Falta el PIN: ver comentario de runTestToken.'); return; }

  // Simular validatePin con PIN correcto
  var res = validatePin(ss, 'juanma', '' + pin);
  Logger.log('validatePin: ' + JSON.stringify(res));

  if (res.ok && res.token) {
    // Verificar token
    var check = checkToken(ss, 'juanma', res.token);
    Logger.log('checkToken válido: ' + JSON.stringify(check));

    // Verificar token inválido
    var bad = checkToken(ss, 'juanma', 'tokenfalso123');
    Logger.log('checkToken inválido: ' + JSON.stringify(bad));

    // Revocar token
    revokeToken(ss, 'juanma', res.token);
    var afterRevoke = checkToken(ss, 'juanma', res.token);
    Logger.log('checkToken post-revoke: ' + JSON.stringify(afterRevoke));
  }
  Logger.log('=== TEST COMPLETO ===');
}
function deleteTestClient() {
  var ss       = SpreadsheetApp.openById(SS_ID);
  var clientId = 'test001';

  // Borrar de hoja clientes
  var cs   = ss.getSheetByName('clientes');
  var data = cs.getDataRange().getValues();
  for (var i = data.length - 1; i >= 1; i--) {
    if ('' + data[i][0] === clientId) cs.deleteRow(i + 1);
  }

  // Borrar todas las filas de config
  var cfg  = ss.getSheetByName('config');
  var rows = cfg.getDataRange().getValues();
  for (var j = rows.length - 1; j >= 0; j--) {
    if ('' + rows[j][0] === clientId) cfg.deleteRow(j + 1);
  }

  SpreadsheetApp.flush();
  Logger.log('test001 eliminado OK');
}

// ============================================================
// ============================================================
// SISTEMA DE NOTIFICACIONES TELEGRAM
// Reconstruido — integrado al backend de Panel Juanma
// ============================================================
// ============================================================

// Token del bot de Telegram. IMPORTANTE: si este token se filtró,
// regeneralo en @BotFather (/revoke) y reemplazá solo esta línea.
var TG_BOT_TOKEN = 'TOKEN_SOLO_EN_EL_EDITOR';   // copia local/repo: el token real vive sólo en el editor de Apps Script

// Chat ID de Juanma (destino de las alertas del panel personal).
var TG_CHAT_ID = '15534655';

// Ciclo de custodia (mismo patrón que el panel). Necesario para que el
// servidor sepa quién tiene a las nenas sin depender del navegador.
var TG_CYCLE_START = '2026-05-08'; // 8 de mayo 2026
var TG_CYCLE = ['mama','mama','mama','papa','papa','mama','mama','papa','papa','papa','mama','mama','papa','papa'];

var TG_CLIENT_ID = 'juanma';

// ── Envío de mensajes a Telegram ─────────────────────────────
function tgSend(text) {
  if (!TG_BOT_TOKEN || !TG_CHAT_ID) return;
  var url = 'https://api.telegram.org/bot' + TG_BOT_TOKEN + '/sendMessage';
  var payload = {
    chat_id: TG_CHAT_ID,
    text: text,
    parse_mode: 'HTML'
  };
  try {
    UrlFetchApp.fetch(url, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });
  } catch (e) {
    Logger.log('Error enviando a Telegram: ' + e.toString());
  }
}

// ── Custodia del lado del servidor ───────────────────────────
function tgPad2(n) { return n < 10 ? '0' + n : '' + n; }
function tgDateKey(d) { return Utilities.formatDate(d, 'America/Argentina/Buenos_Aires', 'yyyy-MM-dd'); }

// Ancla del acuerdo vigente ({date, who}); la carga tgGetOverrides() junto con los overrides.
var TG_ANCHOR = null;

function tgGetOverrides() {
  TG_ANCHOR = null;
  try {
    var ss  = SpreadsheetApp.openById(SS_ID);
    var cfg = readConfig(ss, TG_CLIENT_ID);
    if (cfg.cycle_anchor) {
      try { var a = JSON.parse(cfg.cycle_anchor); TG_ANCHOR = (a && a.date) ? a : null; } catch (x) { TG_ANCHOR = null; }
    }
    return JSON.parse(cfg.overrides || '{}') || {};
  } catch (e) { return {}; }
}

function tgDsToDate(ds) {
  var p = ('' + ds).split('-');
  return new Date(parseInt(p[0],10), parseInt(p[1],10)-1, parseInt(p[2],10));
}

// Misma resolucion en 3 capas que getCustody() del panel:
// 1) override del dia  2) ciclo re-anclado (cycle_anchor)  3) esquema original.
function tgGetCustody(dateStr, overrides) {
  if (overrides && overrides[dateStr]) return overrides[dateStr];
  var d = tgDsToDate(dateStr);
  if (TG_ANCHOR && TG_ANCHOR.date) {
    var dd = Math.round((d.getTime() - tgDsToDate(TG_ANCHOR.date).getTime()) / 86400000);
    if (dd >= 0) {
      var slot = TG_CYCLE[((dd % 14) + 14) % 14];
      if (TG_ANCHOR.who === 'papa') return slot === 'mama' ? 'papa' : 'mama';
      return slot;
    }
  }
  var diff = Math.round((d.getTime() - tgDsToDate(TG_CYCLE_START).getTime()) / 86400000);
  if (diff < 0) return null;
  return TG_CYCLE[((diff % 14) + 14) % 14];
}

function tgWhoLabel(who) { return who === 'papa' ? 'PAPÁ' : (who === 'mama' ? 'MAMÁ' : ''); }

// ── ALERTA: aviso a la tarde (20:00) si mañana hay cambio ────
function tgAlertaTarde() {
  var ov = tgGetOverrides();
  var hoy = new Date();
  var manana = new Date(hoy.getTime()); manana.setDate(manana.getDate() + 1);
  var cHoy = tgGetCustody(tgDateKey(hoy), ov);
  var cMan = tgGetCustody(tgDateKey(manana), ov);
  if (cHoy && cMan && cHoy !== cMan) {
    var dias = ['Domingo','Lunes','Martes','Miércoles','Jueves','Viernes','Sábado'];
    tgSend('🌙 <b>Mañana hay cambio</b>\nMañana ' + dias[manana.getDay()] + ' las nenas pasan con <b>' + tgWhoLabel(cMan) + '</b>.');
  }
}

// ── ALERTA: buenos días (08:00) si hoy cambió el turno ───────
function tgAlertaManana() {
  var ov = tgGetOverrides();
  var hoy = new Date();
  var ayer = new Date(hoy.getTime()); ayer.setDate(ayer.getDate() - 1);
  var cHoy = tgGetCustody(tgDateKey(hoy), ov);
  var cAyer = tgGetCustody(tgDateKey(ayer), ov);
  if (cHoy && cAyer && cHoy !== cAyer) {
    tgSend('☀️ <b>Buenos días!</b>\nHoy las nenas están con <b>' + tgWhoLabel(cHoy) + '</b>.');
  }
}

// ── ALERTA: presupuesto (12:00) si queda 20% o menos ─────────
function tgAlertaPresupuesto() {
  var ss = SpreadsheetApp.openById(SS_ID);
  var info = checkBudgetAlert(ss, TG_CLIENT_ID);
  if (info && info.alert) {
    tgSend('💰 <b>Atención presupuesto</b>\nTe queda el <b>' + info.pct + '%</b> del presupuesto de este mes (' +
           fmtMoneyAR(info.rem) + ' de ' + fmtMoneyAR(info.budget) + ').');
  }
}

function fmtMoneyAR(n) {
  var s = Math.abs(Math.round(n)).toString(); var r = '';
  for (var i = 0; i < s.length; i++) { if (i > 0 && (s.length-i) % 3 === 0) r += '.'; r += s[i]; }
  return '$' + r;
}

// ── ALERTA: recordatorios (cada minuto) ──────────────────────
// El trigger corre cada minuto y reescribe 'reminders': mismo candado que la web.
function tgCheckRecordatorios() { return withLock(tgCheckRecordatoriosSinLock); }
function tgCheckRecordatoriosSinLock() {
  var ss = SpreadsheetApp.openById(SS_ID);
  var cfg = readConfig(ss, TG_CLIENT_ID);
  var rems = [];
  try { rems = JSON.parse(cfg.reminders || '[]') || []; } catch(x) { return; }
  if (!rems.length) return;

  var now = new Date();
  var hhmm = tgPad2(now.getHours()) + ':' + tgPad2(now.getMinutes());
  var hoy = tgDateKey(now);
  var changed = false;

  for (var i = 0; i < rems.length; i++) {
    if (rems[i].time === hhmm && rems[i].lastFired !== hoy) {
      tgSend('⏰ <b>Recordatorio</b>\n' + (rems[i].msg || ''));
      rems[i].lastFired = hoy;
      changed = true;
    }
  }
  if (changed) setConf(ss, TG_CLIENT_ID, 'reminders', JSON.stringify(rems));
}

// ============================================================
// INSTALACIÓN DE TRIGGERS — ejecutar UNA vez a mano
// ============================================================
// Cómo usar:
//   1. En el editor, elegí la función "tgSetupTriggers" en el selector
//   2. Ejecutar
//   3. Autorizá los permisos (UrlFetch + Triggers) cuando lo pida
//   4. Listo: quedan los 4 disparadores activos
// ============================================================
function tgSetupTriggers() {
  var triggers = ScriptApp.getProjectTriggers();
  var nombres  = ['tgAlertaManana','tgAlertaTarde','tgAlertaPresupuesto',
                  'tgCheckRecordatorios','tgAlertaEventosNoche','tgAlertaEventos3hs'];
  for (var i = 0; i < triggers.length; i++) {
    if (nombres.indexOf(triggers[i].getHandlerFunction()) !== -1)
      ScriptApp.deleteTrigger(triggers[i]);
  }

  ScriptApp.newTrigger('tgAlertaManana').timeBased().atHour(8).everyDays(1)
    .inTimezone('America/Argentina/Buenos_Aires').create();
  ScriptApp.newTrigger('tgAlertaTarde').timeBased().atHour(20).everyDays(1)
    .inTimezone('America/Argentina/Buenos_Aires').create();
  ScriptApp.newTrigger('tgAlertaPresupuesto').timeBased().atHour(12).everyDays(1)
    .inTimezone('America/Argentina/Buenos_Aires').create();
  ScriptApp.newTrigger('tgCheckRecordatorios').timeBased().everyMinutes(1).create();
  ScriptApp.newTrigger('tgAlertaEventosNoche').timeBased().atHour(21).everyDays(1)
    .inTimezone('America/Argentina/Buenos_Aires').create();
  // tgAlertaEventos3hs NO se registra acá: lo crea tgAlertaEventosNoche on-demand

  Logger.log('Triggers instalados OK');
  tgSend('✅ <b>Alertas de eventos activadas</b>\nVas a recibir aviso a las 21hs si hay evento al día siguiente, y otro 3hs antes si el evento tiene horario.');
}

// Función suelta para probar el envío sin esperar a un horario
function tgTestEnvio() {
  tgSend('🔔 Prueba manual: si recibís esto, Telegram está funcionando.');
}

// ============================================================
// CÁLCULO DE ALERTA DE PRESUPUESTO (recuperado del backend viejo)
// ============================================================
function checkBudgetAlert(ss, clientId) {
  var cfg        = readConfig(ss, clientId);
  var budget     = +cfg.budget || 3000000;
  var budgetType = '' + (cfg.budget_type || 'efectivo');
  var month      = mesAR();

  // Calcular gastado según tipo
  var gs   = ss.getSheetByName('gastos');
  var cash = 0; var digital = 0;
  if (gs && gs.getLastRow() > 1) {
    var ncols = gs.getLastColumn();
    var rows  = gs.getRange(2, 1, gs.getLastRow() - 1, ncols).getValues();
    for (var i = 0; i < rows.length; i++) {
      var r = rows[i];
      if (!r[0]) continue;
      var rowClient = ncols >= 9 ? '' + r[8] : '';
      if (rowClient && rowClient !== clientId) continue;
      var d = fmtDate(r[1]);
      if (d.substring(0, 7) !== month) continue;
      var pay = ncols > 6 ? '' + r[6] : 'efectivo';
      var amt = +r[3];
      if (!pay || pay === 'efectivo') cash    += amt;
      if (pay === 'billetera')        digital += amt;
      // tarjeta nunca cuenta
    }
  }

  var spent = 0;
  if (budgetType === 'efectivo') spent = cash;
  if (budgetType === 'digital')  spent = digital;
  if (budgetType === 'todo')     spent = cash + digital;

  var rem  = budget - spent;
  var pct  = budget > 0 ? Math.round((rem / budget) * 100) : 100;
  var alert = pct <= 20;

  // Datos del panel para personalizar el mensaje
  var panelUi = defaultPanelUi();
  if (cfg.panel_ui) { try { panelUi = mergePanelUi(JSON.parse(cfg.panel_ui)); } catch(x) {} }

  return {
    ok:          true,
    clientId:    clientId,
    alert:       alert,
    pct:         pct,
    rem:         rem,
    budget:      budget,
    spent:       spent,
    budgetType:  budgetType,
    month:       month,
    papaName:    panelUi.papaName || 'Papa',
    telegramChatId: getTelegramChatId(ss, clientId)
  };
}

function getTelegramChatId(ss, clientId) {
  var cs = ss.getSheetByName('clientes');
  if (!cs) return '';
  var data = cs.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if ('' + data[i][0] === clientId) return '' + (data[i][5] || '');
  }
  return '';
}
// ── Nombres de las nenas ──────────────────────────────────────
var TG_KIDS        = ['Maia', 'Clari'];
var PROP_3H_ALERTS = 'pending3hAlerts'; // clave en PropertiesService

// ── Helpers internos ─────────────────────────────────────────
function tgGetEventExtras(ss) {
  var cfg = readConfig(ss, TG_CLIENT_ID);
  try { return JSON.parse(cfg.event_extras || '{}') || {}; } catch(x) { return {}; }
}
function tgGetPanelUi(ss) {
  var cfg = readConfig(ss, TG_CLIENT_ID);
  var ui = defaultPanelUi();
  if (cfg.panel_ui) { try { ui = mergePanelUi(JSON.parse(cfg.panel_ui)); } catch(x) {} }
  return ui;
}

// ── Constructor del mensaje ───────────────────────────────────
function tgBuildEventoMsg(ev, extra, panelUi, header) {
  var papa  = panelUi.papaName || 'Papá';
  var mama  = panelUi.mamaName || 'Mamá';
  var parts = ev.date.split('-');
  var fecha = parts[2] + '/' + parts[1];

  var msg = header + '\n';
  msg += '📌 <b>' + ev.title + '</b>\n';
  msg += '📅 ' + fecha;
  if (ev.time) msg += ' · <b>' + ev.time + 'hs</b>';
  msg += '\n';

  if (extra && extra.kid) {
    if (extra.kid === 'juntas') {
      msg += '👧👧 Van las dos juntas\n';
      if (extra.takenBy) {
        var ll = extra.takenBy === 'papa' ? papa : extra.takenBy === 'mama' ? mama : (extra.takenByName || '');
        msg += '🚗 Las lleva <b>' + ll + '</b>\n';
      }
    } else {
      msg += '👧 Va <b>' + extra.kid + '</b>\n';
      if (extra.takenBy) {
        var ll2 = extra.takenBy === 'papa' ? papa : extra.takenBy === 'mama' ? mama : (extra.takenByName || '');
        msg += '🚗 La lleva <b>' + ll2 + '</b>\n';
      }
      var otherKid = '';
      for (var k = 0; k < TG_KIDS.length; k++) {
        if (TG_KIDS[k] !== extra.kid) { otherKid = TG_KIDS[k]; break; }
      }
      if (otherKid && extra.otherWith) {
        var cq = extra.otherWith === 'papa' ? papa : extra.otherWith === 'mama' ? mama : (extra.otherWithName || '');
        msg += '🏠 <b>' + otherKid + '</b> se queda con <b>' + cq + '</b>\n';
      }
    }
    if (extra.note) msg += '📝 ' + extra.note + '\n';
  } else {
    msg += '⚠️ <i>Sin info de coordinación cargada en el panel</i>\n';
  }
  return msg;
}

// ── 21:00 — aviso noche + programa trigger 3hs si hay hora ───
function tgAlertaEventosNoche() {
  var ss     = SpreadsheetApp.openById(SS_ID);
  var extras = tgGetEventExtras(ss);
  var ui     = tgGetPanelUi(ss);

  // Limpiar triggers 3hs viejos que pudieran quedar de ayer
  var viejos = ScriptApp.getProjectTriggers();
  for (var t = 0; t < viejos.length; t++) {
    if (viejos[t].getHandlerFunction() === 'tgAlertaEventos3hs')
      ScriptApp.deleteTrigger(viejos[t]);
  }

  var manana    = new Date(); manana.setDate(manana.getDate() + 1);
  var mananaStr = tgDateKey(manana);
  var mesStr    = mananaStr.substring(0, 7);

  var result = getCalendarEvents(mesStr);
  if (!result || !result.events) return;

  var pending3h = [];

  for (var i = 0; i < result.events.length; i++) {
    var ev = result.events[i];
    if (ev.date !== mananaStr) continue;

    // Aviso nocturno siempre
    tgSend(tgBuildEventoMsg(ev, extras[ev.id] || null, ui, '🌙 <b>Mañana hay evento</b>'));

    // Si tiene hora → calcular trigger 3hs antes
    if (ev.time && !ev.allDay) {
      var tp       = ev.time.split(':');
      var evDate   = new Date(manana.getFullYear(), manana.getMonth(), manana.getDate(), +tp[0], +tp[1], 0);
      var alertDate = new Date(evDate.getTime() - 3 * 60 * 60 * 1000);

      if (alertDate > new Date()) {
        pending3h.push({ ev: ev, extra: extras[ev.id] || null, alertAt: alertDate.getTime() });
        ScriptApp.newTrigger('tgAlertaEventos3hs').timeBased().at(alertDate).create();
      }
    }
  }

  // Guardar metadata para que tgAlertaEventos3hs sepa qué enviar
  PropertiesService.getScriptProperties().setProperty(PROP_3H_ALERTS, JSON.stringify(pending3h));
}

// ── One-shot — se autodestruye al ejecutarse ──────────────────
function tgAlertaEventos3hs() {
  var props   = PropertiesService.getScriptProperties();
  var pending = [];
  try { pending = JSON.parse(props.getProperty(PROP_3H_ALERTS) || '[]'); } catch(x) {}

  var ss  = SpreadsheetApp.openById(SS_ID);
  var ui  = tgGetPanelUi(ss);
  var now = Date.now();

  var remaining = [];
  for (var i = 0; i < pending.length; i++) {
    var p    = pending[i];
    var diff = Math.abs(now - p.alertAt);
    if (diff <= 25 * 60 * 1000) {          // dentro de ±25 min del tiempo programado
      tgSend(tgBuildEventoMsg(p.ev, p.extra, ui, '⏰ <b>Evento en ~3 horas</b>'));
    } else if (p.alertAt > now) {
      remaining.push(p);                   // todavía futuro, lo preservamos
    }
    // si ya pasó y no matcheó: lo descartamos silenciosamente
  }

  props.setProperty(PROP_3H_ALERTS, JSON.stringify(remaining));

  // Autodestrucción: borra todos los triggers 3hs y recrea solo los pendientes
  var triggers = ScriptApp.getProjectTriggers();
  for (var t = 0; t < triggers.length; t++) {
    if (triggers[t].getHandlerFunction() === 'tgAlertaEventos3hs')
      ScriptApp.deleteTrigger(triggers[t]);
  }
  for (var j = 0; j < remaining.length; j++) {
    ScriptApp.newTrigger('tgAlertaEventos3hs').timeBased().at(new Date(remaining[j].alertAt)).create();
  }
}

// ============================================================
// RESUMEN DEL MES (público, sólo lectura) — lo usa el informe diario
// ============================================================
function getResumen(ss, clientId, month) {
  var cfg    = readConfig(ss, clientId);
  var budget = +cfg.budget || 3000000;          // total del mes (base + arrastre, como en el panel)
  var m      = month || mesAR();
 
  var cash = 0, billetera = 0, tarjetaArs = 0, tarjetaUsd = 0;
  var gs = ss.getSheetByName('gastos');
  if (gs && gs.getLastRow() > 1) {
    var ncols = gs.getLastColumn();
    var rows  = gs.getRange(2, 1, gs.getLastRow() - 1, ncols).getValues();
    for (var i = 0; i < rows.length; i++) {
      var r = rows[i];
      if (!r[0]) continue;
      var rowClient = ncols >= 9 ? '' + r[8] : '';
      if (rowClient && rowClient !== '' + clientId) continue;
      var d = fmtDate(r[1]);
      if (d.substring(0, 7) !== m) continue;
      var pay    = (ncols > 6 && r[6]) ? ('' + r[6]).toLowerCase() : 'efectivo';
      var moneda = (ncols > 9 && r[9]) ? '' + r[9] : 'ARS';
      var amt    = +r[3] || 0;
      if (pay === 'efectivo' || pay === '')  cash      += amt;
      else if (pay === 'billetera')          billetera += amt;
      else { if (moneda === 'USD') tarjetaUsd += amt; else tarjetaArs += amt; } // tarjeta = resto
    }
  }
 
  return {
    ok: true,
    month: m,
    total: budget,                          // total del mes (3.100.000)
    gastado_efectivo: cash,                 // 2.514.612
    gastado_billetera: billetera,
    disponible_efectivo: budget - cash,     // 585.388  <-- el "dinero para existir" real
    tarjeta_ars: tarjetaArs,                // 989.483 (comprometido, no descuenta)
    tarjeta_usd: tarjetaUsd,                // 20
    total_general: cash + billetera + tarjetaArs
  };
}

// ============================================================
// RESTAURACIÓN 19/10/2026 — correr UNA sola vez desde el editor
// ============================================================
// Elegí "restaurarCustodia20261019" en el selector de funciones → Ejecutar.
// - Lo cargado DESPUÉS del incidente manda sobre lo recuperado.
// - Deja una sola fila por clave en 'config' y pasa el historial a la hoja
//   'historial'. Si se corre de nuevo, no hace nada.
// Días recuperados (reconstruidos cruzando 9 versiones del Sheet:
// 11/10 10:24 y 19/10 10:12, 10:14, 10:17, 10:19, 10:20, 10:21, 10:27, 10:29).
var RESTAURACION_20261019 = {
  "2026-08-17": "mama", "2026-08-19": "papa", "2026-08-21": "mama", "2026-08-26": "mama", "2026-08-28": "papa", "2026-08-29": "papa", "2026-08-31": "mama",
  "2026-09-01": "mama", "2026-09-02": "papa", "2026-09-03": "papa",
  "2026-10-09": "mama",
  "2026-11-18": "papa", "2026-11-19": "papa", "2026-11-20": "mama", "2026-11-21": "mama", "2026-11-22": "mama", "2026-11-24": "papa", "2026-11-26": "mama", "2026-11-28": "papa", "2026-11-29": "papa", "2026-11-30": "mama",
  "2026-12-01": "mama", "2026-12-03": "papa", "2026-12-07": "papa", "2026-12-08": "papa", "2026-12-10": "mama",
  "2027-01-06": "mama", "2027-01-07": "mama", "2027-01-11": "mama", "2027-01-12": "mama", "2027-01-13": "papa", "2027-01-14": "papa"
};

// Historial recuperado: [ts, día, antes, quién, por, fecha de carga, bulk]
var HISTORIAL_RECUPERADO_20261019 = [
  [1791670916511,"2026-10-09","mama","papa","owner","2026-10-10",0],
  [1791670917477,"2026-10-09","mama","mama","owner","2026-10-10",0],
  [1792415575798,"2026-11-18","","papa","ex","2026-10-19",0],
  [1792415665787,"2026-11-22","","mama","ex","2026-10-19",0],
  [1792415667478,"2026-11-24","","papa","ex","2026-10-19",0],
  [1792415675769,"2026-11-26","","mama","ex","2026-10-19",0],
  [1792415680864,"2026-11-29","","papa","ex","2026-10-19",0],
  [1792415687590,"2026-11-28","","papa","ex","2026-10-19",0],
  [1792415691364,"2026-11-30","","mama","ex","2026-10-19",0],
  [1792415826976,"2026-11-18","","papa","ex","2026-10-19",0],
  [1792415828266,"2026-11-19","","papa","ex","2026-10-19",0],
  [1792415941300,"2026-12-01","","mama","ex","2026-10-19",0],
  [1792415958165,"2026-12-03","","papa","ex","2026-10-19",0],
  [1792416007616,"2026-12-07","","papa","ex","2026-10-19",0],
  [1792416007723,"2026-12-08","","papa","ex","2026-10-19",0],
  [1792416079340,"2026-12-08","","papa","ex","2026-10-19",0],
  [1792416541229,"2027-01-12","","mama","owner","2026-10-19",0],
  [1792416544636,"2027-01-13","","papa","owner","2026-10-19",0],
  [1792416544703,"2027-01-14","","papa","owner","2026-10-19",0]
];
function restaurarCustodia20261019() {
  var props = PropertiesService.getScriptProperties();
  if (props.getProperty('restauracion_20261019') === 'hecha') {
    Logger.log('La restauración ya se había corrido: no hago nada.');
    return { ok: true, yaHecha: true };
  }
  var r = withLock(function() {
    var ss  = SpreadsheetApp.openById(SS_ID);
    var cid = 'juanma';
    var cfg = readConfig(ss, cid);
    var cur = {};
    try { cur = JSON.parse(cfg.overrides || '{}') || {}; } catch (e) { cur = {}; }

    var merged = {}, k, agregados = 0, conflictos = [];
    for (k in RESTAURACION_20261019) merged[k] = RESTAURACION_20261019[k];
    for (k in cur) {
      if (!cur.hasOwnProperty(k)) continue;
      if (merged.hasOwnProperty(k) && merged[k] !== cur[k]) conflictos.push(k + ' recuperado ' + merged[k] + ' / actual ' + cur[k]);
      merged[k] = cur[k];
    }
    for (k in merged) if (merged.hasOwnProperty(k) && !cur.hasOwnProperty(k)) agregados++;

    setConf(ss, cid, 'overrides', JSON.stringify(merged));             // en su lugar + borra duplicadas
    if (cfg.override_history !== undefined) setConf(ss, cid, 'override_history', '' + cfg.override_history); // sólo deduplica

    // Historial recuperado → hoja 'historial' (sin duplicar)
    var sh = getHistSheet(ss);
    var ya = {}, prev = readHistorial(ss, cid, 0) || [];
    for (var i = 0; i < prev.length; i++) ya[prev[i].ts + '|' + prev[i].date + '|' + prev[i].who] = 1;
    var filas = [];
    for (var j = 0; j < HISTORIAL_RECUPERADO_20261019.length; j++) {
      var h = HISTORIAL_RECUPERADO_20261019[j];
      if (ya[h[0] + '|' + h[1] + '|' + h[3]]) continue;
      filas.push(['' + h[0], cid, h[1], h[2], h[3], h[4], h[5], '' + h[6], '', '', 'recuperado']);
    }
    appendHistRows(sh, filas);
    registrarHistorial(ss, cid, '#bulk', '', '' + agregados, 'owner', true, undefined, undefined, 'restauracion');
    return { ok: true, total: Object.keys(merged).length, agregados: agregados,
             conflictos: conflictos, historialImportado: filas.length };
  });
  if (r && r.ok) props.setProperty('restauracion_20261019', 'hecha');
  Logger.log(JSON.stringify(r));
  return r;
}
