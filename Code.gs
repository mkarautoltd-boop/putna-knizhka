/**
 * ЕЛЕКТРОННА ПЪТНА КНИЖКА – сървърна част (Google Apps Script)
 * Базата данни е тази Google таблица („Автопарк“). Листовете се създават автоматично.
 * Не редактирай ръчно колоната „DATA“ – от нея приложението чете данните.
 */

const SHEETS = {
  users:    { name: 'Потребители', cols: [['username','Потребител'],['name','Име'],['role','Роля'],['dept','Отдел'],['active','Активен']] },
  depts:    { name: 'Отдели',      cols: [['name','Отдел']] },
  cars:     { name: 'Автомобили',  cols: [['reg','Рег. номер'],['make','Марка'],['model','Модел'],['year','Година'],['fuel','Гориво'],['dept','Отдел'],['driverId','Постоянен шофьор'],['odo','Километраж'],['tankCap','Резервоар л'],['active','Активен']] },
  docsView: { name: 'Документи',   cols: [['reg','Рег. номер'],['gtp','ГТП до'],['go','Гражданска до'],['kasko','Каско до'],['vignette','Винетка до'],['others','Други']] },
  trips:    { name: 'Километраж',  cols: [['date','Дата'],['time','Час'],['carId','Автомобил'],['driver','Шофьор'],['startKm','Начален км'],['endKm','Краен км'],['km','Изминати км'],['from','От'],['to','До'],['route','Маршрут'],['purpose','Цел'],['note','Бележка'],['temp','Временно'],['by','Въвел'],['del','Изтрит']] },
  fuel:     { name: 'Зареждания',  cols: [['date','Дата'],['time','Час'],['carId','Автомобил'],['driver','Шофьор'],['km','Километраж'],['liters','Литри'],['sum','Сума €'],['ppl','Цена/л €'],['station','Бензиностанция'],['note','Бележка'],['by','Въвел'],['del','Изтрит']] },
  service:  { name: 'Сервиз',      cols: [['date','Дата'],['carId','Автомобил'],['t','Вид'],['km','Километраж'],['what','Дейност'],['cat','Категория'],['cost','Сервиз €'],['sum','Разход €'],['place','Сервиз / място'],['note','Бележка'],['by','Въвел'],['del','Изтрит']] },
  tyres:    { name: 'Гуми',        cols: [['date','Дата'],['carId','Автомобил'],['kind','Смяна'],['km','Километраж'],['by','Отбелязал'],['del','Изтрит']] },
  checks:   { name: 'Проверки',    cols: [['date','Дата'],['carId','Автомобил'],['km','Километраж'],['driver','Шофьор'],['oil','Масло'],['oilAdd','Долято масло'],['antifreeze','Антифриз'],['afAdd','Долят антифриз'],['tyresOk','Гуми'],['inflated','Надути'],['note','Бележка'],['by','Въвел'],['del','Изтрит']] },
  assign:   { name: 'Назначения',  cols: [['date','Дата'],['time','Час'],['carId','Автомобил'],['phase','Етап'],['user','Временен шофьор'],['km','Километраж'],['note','Бележка'],['by','Въвел']] },
  history:  { name: 'История',     cols: [['date','Дата'],['carId','Автомобил'],['text','Действие'],['by','Потребител']] },
  problems: { name: 'Проблеми',    cols: [['date','Дата'],['carId','Автомобил'],['cat','Категория'],['desc','Проблем'],['priority','Приоритет'],['status','Статус'],['by','Подал']] },
  tasks:    { name: 'Задачи',      cols: [['title','Задача'],['carId','Автомобил'],['type','Вид'],['due','Срок'],['status','Статус'],['assignee','Изпълнител']] },
  messages: { name: 'Съобщения',   cols: [['text','Съобщение'],['dept','Отдел'],['by','От'],['del','Изтрито']] },
  notif:    { name: 'Известия',    cols: [['count','Прочетени известия']] },
  meta:     { name: 'Настройки',   cols: [] }
};
const ENTRY_KEYS = ['trips', 'fuel', 'service', 'tyres', 'checks', 'assign', 'history'];
const T_SHEET = { trip: 'trips', fuel: 'fuel', service: 'service', expense: 'service', tyres: 'tyres', check: 'checks', session: 'assign', edit: 'history' };
const LABELS = {
  t: { trip: 'Пътен запис', fuel: 'Зареждане', service: 'Сервиз', expense: 'Друг разход', tyres: 'Смяна на гуми', check: 'Проверка', session: 'Временно управление', edit: 'Промяна' },
  role: { admin: 'Администратор', manager: 'Мениджър', driver: 'Шофьор' },
  kind: { winter: 'Летни → зимни', summer: 'Зимни → летни' },
  phase: { start: 'Поемане', end: 'Приключване' },
  status: { new: 'Нов', progress: 'В процес', resolved: 'Решен', open: 'Отворена', done: 'Изпълнена' },
  priority: { low: 'Нисък', mid: 'Среден', high: 'Висок' }
};
const TOKEN_DAYS = 90;

/* ---------- web app ---------- */
function doGet(e) {
  const t = HtmlService.createTemplateFromFile('Index');
  t.prefillUser = (e && e.parameter && e.parameter.u) || '';
  t.appUrl = ScriptApp.getService().getUrl();
  return t.evaluate()
    .setTitle('Пътна книжка')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/* ---------- sheets ---------- */
let SS_ = null;
/* Работи и когато скриптът е отделен проект: тогава създава/ползва таблица „Автопарк“ в Google Drive. */
function ss_() {
  if (SS_) return SS_;
  const active = SpreadsheetApp.getActiveSpreadsheet();
  if (active) return (SS_ = active);
  const p = PropertiesService.getScriptProperties(), id = p.getProperty('SHEET_ID');
  if (id) { try { return (SS_ = SpreadsheetApp.openById(id)); } catch (e) {} }
  SS_ = SpreadsheetApp.create('Автопарк');
  p.setProperty('SHEET_ID', SS_.getId());
  return SS_;
}
function width_(key) { return 2 + SHEETS[key].cols.length + 1; }
let SHEETMAP_ = null;
function sheetByName_(name) {
  if (!SHEETMAP_) { SHEETMAP_ = {}; ss_().getSheets().forEach(x => { SHEETMAP_[x.getName()] = x; }); }
  return SHEETMAP_[name] || null;
}
function sheet_(key, checkHeader) {
  const d = SHEETS[key];
  let s = sheetByName_(d.name);
  if (!s) {
    s = ss_().insertSheet(d.name);
    const head = ['ID', 'Обновено'].concat(d.cols.map(c => c[1])).concat(['DATA (не редактирай)']);
    s.getRange(1, 1, 1, head.length).setValues([head]).setFontWeight('bold').setBackground('#0E2748').setFontColor('#FFFFFF');
    s.setFrozenRows(1);
    s.setColumnWidth(head.length, 120);
    SHEETMAP_[d.name] = s;
  } else if (checkHeader && s.getLastColumn() < width_(key) && s.getRange(1, 1, 1, 1).getValue() === 'ID') {
    const head = ['ID', 'Обновено'].concat(d.cols.map(c => c[1])).concat(['DATA (не редактирай)']);
    s.getRange(1, 1, 1, head.length).setValues([head]).setFontWeight('bold').setBackground('#0E2748').setFontColor('#FFFFFF');
  }
  return s;
}
function readAll_(key, cache) {
  if (cache && cache[key]) return cache[key];
  const s = sheet_(key), all = s.getDataRange().getValues(), out = {};
  const w = Math.max(width_(key), all.length ? all[0].length : 0);
  if (all.length >= 2) {
    const v = all.slice(1);
    v.forEach((r, i) => {
      if (!r[0]) return;
      let j = r[w - 1];
      if (typeof j !== 'string' || j.charAt(0) !== '{') { for (let k = r.length - 1; k > 1; k--) { if (typeof r[k] === 'string' && r[k].charAt(0) === '{') { j = r[k]; break; } } }
      try { out[r[0]] = { row: i + 2, data: JSON.parse(j) }; } catch (e) {}
    });
  }
  if (cache) cache[key] = out;
  return out;
}
function lookup_(key, id, cache) {
  const b = cache && cache._base !== undefined ? cache._base : (cache ? (cache._base = cacheGet_()) : null);
  if (b && b[key] && b[key][id]) return b[key][id];
  const r = readAll_(key, cache)[id];
  return r ? r.data : null;
}
function display_(field, val, cache) {
  if (val === null || val === undefined) return '';
  if (field === 'carId') { const c = lookup_('cars', val, cache); return c ? c.reg : val; }
  if (['driver', 'by', 'driverId', 'assignee', 'user'].indexOf(field) >= 0) { const u = lookup_('users', val, cache); return u ? u.name : val; }
  if (field === 'dept') { const d = lookup_('depts', val, cache); return d ? d.name : (val || 'Всички'); }
  if (LABELS[field] && LABELS[field][val]) return LABELS[field][val];
  if (typeof val === 'boolean') return val ? 'Да' : '';
  if (typeof val === 'object') return JSON.stringify(val);
  return val;
}
function writeDoc_(key, id, data, cache, isNew) {
  const s = sheet_(key, true), all = isNew ? (cache[key] || {}) : readAll_(key, cache), w = width_(key);
  (cache._written = cache._written || []).push([key, id, data]);
  const json = JSON.stringify(data);
  if (json.length > 49000) throw new Error('Записът е твърде голям. Опитай с по-малка снимка.');
  const row = [id, new Date()].concat(SHEETS[key].cols.map(c => display_(c[0], data[c[0]], cache))).concat([json]);
  if (all[id]) s.getRange(all[id].row, 1, 1, w).setValues([row]);
  else { s.appendRow(row); all[id] = { row: isNew ? 0 : s.getLastRow() }; }
  all[id].data = data;
  if (key === 'cars') {
    const d = data.docs || {};
    writeDoc_('docsView', id, { reg: data.reg, gtp: d.gtp || '', go: d.go || '', kasko: d.kasko || '', vignette: d.vignette || '',
      others: (d.others || []).filter(o => o && o.name).map(o => o.name + (o.until ? ' до ' + o.until : '')).join('; ') }, cache);
  }
}
function merge_(a, b) {
  const out = Object.assign({}, a || {});
  Object.keys(b || {}).forEach(k => {
    const v = b[k];
    if (v && typeof v === 'object' && !Array.isArray(v) && out[k] && typeof out[k] === 'object' && !Array.isArray(out[k])) out[k] = merge_(out[k], v);
    else out[k] = v;
  });
  return out;
}
function findEntry_(eid, cache) {
  for (let i = 0; i < ENTRY_KEYS.length; i++) { const a = readAll_(ENTRY_KEYS[i], cache); if (a[eid]) return { key: ENTRY_KEYS[i], data: a[eid].data }; }
  return null;
}
function writeEntry_(eid, data, cache, isNew) {
  const key = T_SHEET[data.t] || 'history';
  writeDoc_(key, eid, data, cache, isNew);
}
/* Прехвърля записи от старата версия (лист „Пътна книжка“), ако има такъв. */
function migrateLegacy_(cache) {
  const old = sheetByName_('Пътна книжка');
  if (!old) return;
  const n = old.getLastRow(), w = old.getLastColumn();
  if (n >= 2 && w >= 3) {
    old.getRange(2, 1, n - 1, w).getValues().forEach(r => {
      if (!r[0]) return;
      try { const d = JSON.parse(r[w - 1]); if (d && d.t && !findEntry_(r[0], cache)) writeEntry_(r[0], d, cache); } catch (e) {}
    });
  }
  old.setName('Пътна книжка (архив ' + Utilities.formatDate(new Date(), 'Europe/Sofia', 'yyyy-MM-dd') + ')');
}

/* ---------- auth ---------- */
function hash_(s) {
  return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, s, Utilities.Charset.UTF_8)
    .map(b => ('0' + ((b + 256) % 256).toString(16)).slice(-2)).join('');
}
function ensureAdmin_(cache) {
  const users = readAll_('users', cache);
  if (Object.keys(users).length === 0) {
    writeDoc_('users', 'admin', { name: 'Администратор', username: 'admin', role: 'admin', dept: null, active: true, salt: 'init', hash: hash_('init:1234'), created: new Date().toISOString() }, cache);
  }
}
function props_() { return PropertiesService.getScriptProperties(); }
function auth_(token, cache) {
  if (!token) throw new Error('AUTH');
  const raw = props_().getProperty('t_' + token);
  if (!raw) throw new Error('AUTH');
  const o = JSON.parse(raw);
  if (o.exp < Date.now()) { props_().deleteProperty('t_' + token); throw new Error('AUTH'); }
  const u = lookup_('users', o.u, cache);
  if (!u || u.active === false) throw new Error('AUTH');
  return Object.assign({ id: o.u }, u);
}
function cleanTokens_() {
  const p = props_(), all = p.getProperties(), now = Date.now();
  Object.keys(all).forEach(k => { if (k.indexOf('t_') === 0) { try { if (JSON.parse(all[k]).exp < now) p.deleteProperty(k); } catch (e) { p.deleteProperty(k); } } });
}

function login(username, password) {
  const cache = {};
  ensureAdmin_(cache);
  migrateLegacy_(cache);
  const un = String(username || '').trim().toLowerCase();
  const users = readAll_('users', cache);
  const id = Object.keys(users).find(k => String(users[k].data.username || '').toLowerCase() === un);
  const u = id && users[id].data;
  if (!u || u.active === false || hash_(u.salt + ':' + password) !== u.hash) throw new Error('Грешно потребителско име или парола.');
  cleanTokens_();
  const token = Utilities.getUuid();
  props_().setProperty('t_' + token, JSON.stringify({ u: id, exp: Date.now() + TOKEN_DAYS * 864e5 }));
  return { token: token, userId: id, data: collect_(cache, id) };
}
function logout(token) { if (token) props_().deleteProperty('t_' + token); return true; }

function changePassword(token, oldPass, newPass) {
  const cache = {}, me = auth_(token, cache), full = readAll_('users', cache)[me.id].data;
  if (hash_(full.salt + ':' + oldPass) !== full.hash) throw new Error('Текущата парола е грешна.');
  if (String(newPass || '').length < 4) throw new Error('Новата парола трябва да е поне 4 символа.');
  const salt = Utilities.getUuid().slice(0, 12);
  const data = merge_(readAll_('users', cache)[me.id].data, { salt: salt, hash: hash_(salt + ':' + newPass) });
  writeDoc_('users', me.id, data, cache);
  return true;
}

/* ---------- read ---------- */
const CK_ = 'pk_data_v2';
function cacheGet_() {
  try {
    const c = CacheService.getScriptCache(), n = +c.get(CK_ + '_n');
    if (!n) return null;
    const keys = []; for (let i = 0; i < n; i++) keys.push(CK_ + '_' + i);
    const m = c.getAll(keys); let str = '';
    for (let i = 0; i < n; i++) { const part = m[CK_ + '_' + i]; if (part == null) return null; str += part; }
    return JSON.parse(str);
  } catch (e) { return null; }
}
function cachePut_(obj) {
  try {
    const str = JSON.stringify(obj), size = 30000, parts = {}; let n = 0;
    for (let i = 0; i < str.length; i += size) parts[CK_ + '_' + (n++)] = str.slice(i, i + size);
    if (n > 900) return;
    parts[CK_ + '_n'] = String(n);
    CacheService.getScriptCache().putAll(parts, 21600);
  } catch (e) {}
}
function cacheClear_() { try { CacheService.getScriptCache().remove(CK_ + '_n'); } catch (e) {} }
/* Прилага записаното към кеша, за да не се чете цялата таблица отново. */
function cachePatch_(cache) {
  const w = cache._written; if (!w || !w.length) return;
  const b = cacheGet_(); if (!b) return;
  const sect = { users: 'users', depts: 'depts', cars: 'cars', problems: 'problems', tasks: 'tasks', messages: 'messages' };
  w.forEach(([key, id, data]) => {
    if (sect[key]) { const v = Object.assign({}, data); if (key === 'users') { delete v.salt; delete v.hash; } b[sect[key]][id] = v; }
    else if (key === 'meta' && id === 'settings') b.settings = data;
    else if (key === 'notif') { b.notifAll = b.notifAll || {}; b.notifAll[id] = data; }
    else if (ENTRY_KEYS.indexOf(key) >= 0 && data.carId && data.date) {
      const docId = data.carId + '_' + String(data.date).slice(0, 7);
      if (!b.logs[docId]) b.logs[docId] = { carId: data.carId, month: String(data.date).slice(0, 7), e: {} };
      b.logs[docId].e[id] = data;
    }
  });
  cachePut_(b);
}
function base_(cache) {
  if (cache._base) return cache._base;
  const hit = cacheGet_(); if (hit) return (cache._base = hit);
  const plain = key => { const a = readAll_(key, cache), o = {}; Object.keys(a).forEach(k => o[k] = a[k].data); return o; };
  const users = plain('users');
  Object.keys(users).forEach(k => { delete users[k].salt; delete users[k].hash; });
  const logs = {};
  ENTRY_KEYS.forEach(key => {
    const ents = readAll_(key, cache);
    Object.keys(ents).forEach(id => {
      const e = ents[id].data; if (!e.carId || !e.date) return;
      const docId = e.carId + '_' + String(e.date).slice(0, 7);
      if (!logs[docId]) logs[docId] = { carId: e.carId, month: String(e.date).slice(0, 7), e: {} };
      logs[docId].e[id] = e;
    });
  });
  const meta = readAll_('meta', cache);
  const b = { users: users, depts: plain('depts'), cars: plain('cars'), problems: plain('problems'), tasks: plain('tasks'),
    messages: plain('messages'), logs: logs, settings: meta.settings ? meta.settings.data : null, notifAll: plain('notif') };
  cachePut_(b);
  return (cache._base = b);
}
function collect_(cache, meId) {
  const b = base_(cache), out = {};
  Object.keys(b).forEach(k => { if (k !== 'notifAll') out[k] = b[k]; });
  out.me = meId; out.notif = (b.notifAll || {})[meId] || {};
  return out;
}
function getAll(token) {
  const cache = {}, me = auth_(token, cache);
  return collect_(cache, me.id);
}

/* ---------- write ---------- */
function write(token, ops) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const cache = {}, me = auth_(token, cache);
    (ops || []).forEach(op => applyOp_(op, me, cache));
    SpreadsheetApp.flush();
    cachePatch_(cache);
    return { ok: true };
  } finally { lock.releaseLock(); }
}
function applyOp_(op, me, cache) {
  const parts = String(op.path || '').split('/');
  const col = parts[0], id = parts[1];
  if (!id) throw new Error('Невалиден запис.');
  const isAdmin = me.role === 'admin', isStaff = isAdmin || me.role === 'manager';

  if (col === 'logs') {
    const carId = id.slice(0, id.lastIndexOf('_'));
    const e = (op.data && op.data.e) || {};
    Object.keys(e).forEach(eid => {
      const found = (op.op === 'set' || op.isNew) ? null : findEntry_(eid, cache);
      const data = merge_(found ? found.data : {}, Object.assign({}, e[eid], { carId: carId }));
      writeEntry_(eid, data, cache, !found);
    });
    return;
  }
  const key = col;
  if (!SHEETS[key] || ENTRY_KEYS.indexOf(key) >= 0 || key === 'docsView') throw new Error('Невалиден запис.');
  if ((key === 'depts' || key === 'meta') && !isAdmin) throw new Error('Само администратор може да прави тази промяна.');
  if (key === 'users' && !isAdmin) {
    const cur = readAll_('users', cache)[id], next = merge_(cur ? cur.data : {}, op.data || {});
    const ok = me.role === 'manager' && next.role === 'driver' && next.dept === me.dept && (!cur || (cur.data.role === 'driver' && cur.data.dept === me.dept));
    if (!ok) throw new Error('Мениджърът може да добавя и редактира само шофьори от своя отдел.');
  }
  if (key === 'cars' && op.op === 'set' && !isStaff) throw new Error('Нямаш права да добавяш автомобили.');
  if (key === 'notif' && id !== me.id) throw new Error('Невалиден запис.');
  const all = readAll_(key, cache);
  if (key === 'messages') {
    if (op.op === 'set' && !isStaff) throw new Error('Само мениджър или администратор изпраща съобщения.');
    if (op.op === 'update' && !(isAdmin || (all[id] && all[id].data.by === me.id))) throw new Error('Нямаш права за това съобщение.');
  }
  if (op.op === 'set') {
    let data = op.data || {};
    if (key === 'users' && all[id] && !data.hash) data = merge_({ salt: all[id].data.salt, hash: all[id].data.hash }, data);
    if (key === 'notif') data.count = Object.keys(data.read || {}).length;
    writeDoc_(key, id, data, cache);
  } else if (op.op === 'update') {
    if (!all[id]) throw new Error('Записът не съществува.');
    writeDoc_(key, id, merge_(all[id].data, op.data || {}), cache);
  } else throw new Error('Невалидна операция.');
}

/* ---------- export / print ---------- */
function folder_() {
  const name = 'Пътна книжка – отчети';
  const it = DriveApp.getFoldersByName(name);
  return it.hasNext() ? it.next() : DriveApp.createFolder(name);
}
function exportCsv(token, filename, csv) {
  auth_(token, {});
  return folder_().createFile(filename, csv, MimeType.CSV).getUrl();
}
function makePdf(token, name, html) {
  auth_(token, {});
  const pdf = Utilities.newBlob(html, 'text/html', name + '.html').getAs('application/pdf').setName(name + '.pdf');
  return folder_().createFile(pdf).getUrl();
}

/* ---------- първоначална настройка ----------
 * Избери „setup“ от менюто с функции горе и натисни „Изпълни“ (Run).
 * Създава всички листове и администратора admin / 1234. Може да се пуска многократно. */
function setup() {
  const cache = {};
  Object.keys(SHEETS).forEach(k => sheet_(k, true));
  cacheClear_();
  ensureAdmin_(cache);
  migrateLegacy_(cache);
  Logger.log('Готово. Базата данни е тук: ' + ss_().getUrl());
}

/* Вмъква част от приложението (App1–App4.html) в Index.html. */
function include(name) {
  return HtmlService.createHtmlOutputFromFile(name).getContent();
}

/* ---------- API за външната страница (GitHub Pages) ----------
 * Страницата праща POST с {fn, args}; връщаме JSON {ok} или {err}. */
const API_FNS = { login: login, logout: logout, getAll: getAll, write: write, changePassword: changePassword, exportCsv: exportCsv, makePdf: makePdf };
function doPost(e) {
  let out;
  try {
    const req = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    const fn = API_FNS[req.fn];
    if (!fn) throw new Error('Невалидна заявка.');
    out = { ok: fn.apply(null, req.args || []) };
  } catch (err) {
    out = { err: String(err && err.message || err) };
  }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}
