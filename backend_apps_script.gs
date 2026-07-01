// ============================================================
// COPARENTALIDAD — Apps Script Central Multi-Cliente
// Con sistema de tokens de sesión para escritura segura
// ============================================================

var SS_ID = '1rj-nUlQ8jTurpj2wNEEAcXLXNh0sr6rpfEpxMtIQ-D8';

// TTL del token en milisegundos (2 horas)
var TOKEN_TTL = 2 * 60 * 60 * 1000;

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
function doGet(e) {
  var p   = e.parameter || {};
  var res;
  try {
    var ss       = SpreadsheetApp.openById(SS_ID);
    var clientId = p.clientId || '';
    var mes      = p.month    || mesAR();
    var act      = p.action   || 'getAll';

    if (!clientId) return jsonp(p.callback, { error: 'clientId requerido' });

    // ── Acciones públicas (no requieren token) ────────────────
    // Lectura y carga de datos cotidianos: sin friccion, sin PIN.
    // El PIN/token solo protege acciones sensibles (presupuesto,
    // calendario, configuracion), no la carga de gastos ni la lista.
    if (act === 'getAll')           { res = getAll(ss, clientId, mes); }
    else if (act === 'validatePin') { res = validatePin(ss, clientId, p.pin || ''); }
    else if (act === 'debug')       { res = debug(ss, clientId, mes); }
    else if (act === 'addExp')      { res = addExp(ss, clientId, p); }
    else if (act === 'delExp')      { res = delExp(ss, clientId, p.id); }
    else if (act === 'saveShopping'){ res = saveShopping(ss, clientId, p.list || '[]'); }
    else if (act === 'getEvents')   { res = getCalendarEvents(p.month || mes); }
    else if (act === 'saveEventExtra') { res = saveEventExtra(ss, clientId, p); }

    // ── Acciones sensibles (requieren token válido) ───────────
    else {
      var tokenCheck = checkToken(ss, clientId, p.token || '');
      if (!tokenCheck.ok) return jsonp(p.callback, { error: 'sesion invalida', code: 401 });

      // quién hace el cambio, según el token (owner / ex)
      var actorMode = tokenCheck.mode ? tokenCheck.mode : 'owner';

      if      (act === 'clearMonth')     res = clearMonth(ss, clientId, p.month || mes);
      else if (act === 'setBudget')      res = setConf(ss, clientId, 'budget', p.amount);
      else if (act === 'setOverride')    res = setOverride(ss, clientId, p.date, p.who, actorMode);
      else if (act === 'setOverridesBulk') res = setOverridesBulk(ss, clientId, p.changes, actorMode);
      else if (act === 'resetOverrides') res = resetOverrides(ss, clientId, actorMode);
      else if (act === 'addRem')         res = modRems(ss, clientId, 'add', p);
      else if (act === 'delRem')         res = modRems(ss, clientId, 'del', p);
      else if (act === 'setShareShop')   res = setConf(ss, clientId, 'share_shopping', p.value === 'true' ? 'true' : 'false');
      else if (act === 'setPanelUi')     res = setPanelUi(ss, clientId, p);
      else if (act === 'logout')         res = revokeToken(ss, clientId, p.token || '');
      else                               res = { error: 'accion desconocida: ' + act };
    }
  } catch (err) {
    res = { error: err.toString() };
  }
  return jsonp(p.callback, res);
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

// Crea un token nuevo para el cliente y lo guarda en config
function createToken(ss, clientId, mode) {
  var token     = generateToken();
  var expiresAt = Date.now() + TOKEN_TTL;
  var tokenData = JSON.stringify({ token: token, mode: mode, expiresAt: expiresAt });
  setConf(ss, clientId, 'session_token', tokenData);
  return { ok: true, token: token, mode: mode, expiresAt: expiresAt };
}

// Verifica que el token sea válido y no haya expirado
function checkToken(ss, clientId, token) {
  if (!token) return { ok: false, error: 'token requerido' };
  var cfg = readConfig(ss, clientId);
  if (!cfg.session_token) return { ok: false, error: 'sin sesion activa' };
  var data;
  try { data = JSON.parse(cfg.session_token); } catch(e) { return { ok: false, error: 'token corrupto' }; }
  if (data.token !== token)          return { ok: false, error: 'token invalido' };
  if (Date.now() > data.expiresAt)   return { ok: false, error: 'sesion expirada' };
  return { ok: true, mode: data.mode };
}

// Invalida el token (logout)
function revokeToken(ss, clientId, token) {
  setConf(ss, clientId, 'session_token', '{}');
  return { ok: true };
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

function setConf(ss, clientId, key, val) {
  var cs   = getConfigSheet(ss);
  var data = cs.getDataRange().getValues();
  for (var i = data.length - 1; i >= 0; i--) {
    if ('' + data[i][0] === '' + clientId && '' + data[i][1] === '' + key) {
      cs.deleteRow(i + 1);
    }
  }
  cs.appendRow([clientId, key, val]);
  SpreadsheetApp.flush();
  return { ok: true };
}

function ensureConf(ss, clientId, key, defVal) {
  var cs   = getConfigSheet(ss);
  var data = cs.getDataRange().getValues();
  for (var i = 0; i < data.length; i++) {
    if ('' + data[i][0] === '' + clientId && '' + data[i][1] === '' + key) return;
  }
  cs.appendRow([clientId, key, defVal]);
}

// ============================================================
// GET ALL (público — no requiere token)
// ============================================================
function getAll(ss, clientId, month) {
  var res = {
    expenses: [], budget: 3000000, overrides: {},
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
  if (cfg.overrides)      { try { res.overrides     = JSON.parse(cfg.overrides)     || {}; } catch(x){} }
  if (cfg.reminders)      { try { res.reminders     = JSON.parse(cfg.reminders)     || []; } catch(x){} }
  if (cfg.shopping_list)  { try { res.shoppingList  = JSON.parse(cfg.shopping_list) || []; } catch(x){} }
  res.shareShopping = (cfg.share_shopping !== undefined) ? ('' + cfg.share_shopping) === 'true' : true;
  if (cfg.panel_ui)       { try { res.panelUi = mergePanelUi(JSON.parse(cfg.panel_ui)); } catch(x){} }
  if (cfg.override_history){ try { res.overrideHistory = JSON.parse(cfg.override_history) || []; } catch(x){ res.overrideHistory = []; } }
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
  for (var i = 0; i < changes.length; i++) {
    var c = changes[i];
    if (!c || !c.date || !c.who) continue;
    var prev = ov[c.date] || '';
    ov[c.date] = c.who;
    registrarHistorial(ss, clientId, c.date, prev, c.who, actorMode, true);
  }
  setConf(ss, clientId, 'overrides', JSON.stringify(ov));
  return { ok: true, applied: changes.length };
}

function resetOverrides(ss, clientId, actorMode) {
  setConf(ss, clientId, 'overrides', '{}');
  registrarHistorial(ss, clientId, '*', '', 'reset', actorMode);
  return { ok: true };
}

// Registra una línea en el historial de cambios de custodia
function registrarHistorial(ss, clientId, date, prev, who, actorMode, bulk) {
  var cfg = readConfig(ss, clientId);
  var hist = [];
  try { hist = JSON.parse(cfg.override_history || '[]') || []; } catch(x) {}
  hist.push({
    date: date,
    prev: prev,
    who: who,
    by: actorMode === 'ex' ? 'ex' : 'owner',
    at: hoyAR(),
    ts: Date.now(),
    bulk: bulk ? 1 : 0
  });
  // Limitar a las últimas 200 entradas para no inflar
  if (hist.length > 200) hist = hist.slice(hist.length - 200);
  setConf(ss, clientId, 'override_history', JSON.stringify(hist));
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
    setConf(ss, clientId, 'pin_owner', hashPin('7746', 'owner'));
    setConf(ss, clientId, 'pin_ex',    hashPin('1056', 'ex'));
    Logger.log('PINs inicializados');
  }

  SpreadsheetApp.flush();
  Logger.log('Setup OK — claves: ' + Object.keys(readConfig(ss, clientId)).join(', '));
}

function reconstruirConfig() {
  var ss = SpreadsheetApp.openById(SS_ID);
  var cs = ss.getSheetByName('config');
  if (!cs) cs = ss.insertSheet('config');
  cs.clearContents();

  var rows = [
    ['juanma', 'budget',         '848480'],
    ['juanma', 'overrides',      '{"2026-06-17":"mama","2026-06-18":"mama","2026-05-25":"mama","2026-05-27":"papa","2026-06-15":"papa","2026-06-16":"papa","2026-06-21":"papa"}'],
    ['juanma', 'reminders',      '[]'],
    ['juanma', 'pin_owner',      '1qj9jun'],
    ['juanma', 'pin_ex',         '1p8rzop'],
    ['juanma', 'panel_ui',       '{"papaName":"Juan","mamaName":"Mariana","kidsTemplate":"l@s chic@s","kidsTemplateSingular":"la nena","kidsCount":2,"country":"AR","helperName":"Maria","ownerHouse":"Prados","exHouse":"Aires","hasHelper":true}'],
    ['juanma', 'share_shopping', 'false'],
    ['juanma', 'shopping_list',  '[{"id":"li1779084202750","name":"Almohaditas Rellenas","done":false},{"id":"li1779084188878","name":"Mani Tostado sin Sal","done":false}]']
  ];

  for (var i = 0; i < rows.length; i++) cs.appendRow(rows[i]);
  SpreadsheetApp.flush();

  var data = cs.getDataRange().getValues();
  Logger.log('Filas: ' + data.length);
  for (var j = 0; j < data.length; j++) {
    Logger.log('Fila '+(j+1)+': ['+data[j][0]+'] | ['+data[j][1]+'] | ['+data[j][2]+']');
  }
}

function runTestToken() {
  var ss  = SpreadsheetApp.openById(SS_ID);
  Logger.log('=== TEST TOKENS ===');

  // Simular validatePin con PIN correcto
  var res = validatePin(ss, 'juanma', '7746');
  Logger.log('validatePin 7746: ' + JSON.stringify(res));

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
var TG_BOT_TOKEN = '8902659574:AAH3wN1Di3gAzw0YMI_jhmnoYV9s3BPMUSw';

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
function tgDateKey(d) { return d.getFullYear() + '-' + tgPad2(d.getMonth()+1) + '-' + tgPad2(d.getDate()); }

function tgGetOverrides() {
  try {
    var ss  = SpreadsheetApp.openById(SS_ID);
    var cfg = readConfig(ss, TG_CLIENT_ID);
    return JSON.parse(cfg.overrides || '{}') || {};
  } catch (e) { return {}; }
}

function tgGetCustody(dateStr, overrides) {
  if (overrides && overrides[dateStr]) return overrides[dateStr];
  var p = dateStr.split('-');
  var d = new Date(parseInt(p[0],10), parseInt(p[1],10)-1, parseInt(p[2],10));
  var sp = TG_CYCLE_START.split('-');
  var start = new Date(parseInt(sp[0],10), parseInt(sp[1],10)-1, parseInt(sp[2],10));
  var diff = Math.round((d.getTime() - start.getTime()) / 86400000);
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
function tgCheckRecordatorios() {
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
