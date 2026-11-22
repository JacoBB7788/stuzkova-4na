/**
 * Stužková 4.NA – spoločná appka triedy.
 * Údaje sú v Google tabuľke, ktorú si skript vytvorí pri prvom spustení funkcie setup().
 * Appku otvorí ktokoľvek s odkazom (aj bez Google účtu) a každý v nej môže písať.
 * Pri aktualizácii: vložiť celý tento súbor do Kód.gs, uložiť, Deploy → Manage deployments → New version.
 */

const SHEETS = {
  tasks:     { name: 'Úlohy',     cols: ['id', 'title', 'who', 'price', 'due', 'status', 'note'],
               labels: ['ID', 'Úloha', 'Zodpovedná osoba', 'Cena (€)', 'Termín', 'Stav', 'Poznámky / kontakt'] },
  students:  { name: 'Žiaci',     cols: ['id', 'name', 'parent1', 'parent2', 'plusOne', 'agreed', 'paid', 'alc', 'alcQty', 'soft', 'softQty', 'drinkNote', 'order'],
               labels: ['ID', 'Meno žiaka', 'Rodič 1', 'Rodič 2', '+1', 'Dohodnutá suma (€)', 'Zaplatené (€)', 'Alkohol', 'Alkohol – množstvo', 'Nealko', 'Nealko – množstvo', 'Poznámka k nápojom', 'Objednávka nápojov'] },
  catalog:   { name: 'Ponuka nápojov', cols: ['id', 'cat', 'name', 'price'],
               labels: ['ID', 'Kategória', 'Názov', 'Cena za kus bez DPH (€)'] },
  teachers:  { name: 'Učitelia',  cols: ['id', 'name', 'role', 'note', 'invited'],
               labels: ['ID', 'Meno a priezvisko', 'Funkcia', 'Poznámka', 'Pozvánka'] },
  allergies: { name: 'Alergény',  cols: ['id', 'name', 'allergen', 'severity', 'confirmed', 'note'],
               labels: ['ID', 'Meno', 'Alergén / obmedzenie', 'Závažnosť', 'Potvrdené s cateringom', 'Poznámka'] },
  expenses:  { name: 'Výdavky',   cols: ['id', 'name', 'amount', 'state', 'note'],
               labels: ['ID', 'Položka', 'Suma (€)', 'Zaplatené?', 'Poznámka'] },
  program:   { name: 'Program',   cols: ['id', 'time', 'title'],
               labels: ['ID', 'Čas', 'Bod programu'] },
  seats:     { name: 'Zasadací poriadok', cols: ['id', 'name', 'table'],
               labels: ['Miesto', 'Meno', 'Stôl'] },
};
const SETTINGS_SHEET = 'Nastavenia';
const LOG_SHEET = 'Záznam zmien';
const SETTING_KEYS = ['className', 'date', 'venue', 'fee'];
const MAX_LEN = 2000;

// Počiatočné údaje pre setup(). Tabuľka už existuje, takže tu nič netreba (mená sa do kódu nedávajú).
const SEED = { settings: {} };

/* ---------- Jednorazové nastavenie ---------- */

function setup() {
  const props = PropertiesService.getScriptProperties();
  if (props.getProperty('SHEET_ID')) {
    Logger.log('Tabuľka už existuje: ' + SpreadsheetApp.openById(props.getProperty('SHEET_ID')).getUrl());
    return;
  }
  const ss = SpreadsheetApp.create('Stužková 4.NA – dáta appky');
  Object.keys(SHEETS).forEach(function (kind) { createSheet(ss, kind); });
  const st = ss.insertSheet(SETTINGS_SHEET);
  st.getRange(1, 1, 1000, 2).setNumberFormat('@');
  st.getRange(1, 1, SETTING_KEYS.length, 2).setValues(SETTING_KEYS.map(function (k) { return [k, clean((SEED.settings || {})[k])]; }));
  const log = ss.insertSheet(LOG_SHEET);
  log.getRange(1, 1, 1, 5).setValues([['Kedy', 'Kto', 'Časť', 'Akcia', 'Údaje']]).setFontWeight('bold');
  log.setFrozenRows(1);
  const first = ss.getSheetByName('Sheet1') || ss.getSheetByName('Hárok1');
  if (first) ss.deleteSheet(first);
  props.setProperty('SHEET_ID', ss.getId());
  bump();
  Logger.log('Hotovo. Tabuľka s dátami: ' + ss.getUrl());
}

/* ---------- Web appka a API ---------- */

// Samostatná appka (ikona na ploche) volá tieto funkcie, nič iné nie je zvonka dostupné.
const API = { getData: getData, getVersion: getVersion, saveItem: saveItem, deleteItem: deleteItem, saveSettings: saveSettings };

function doPost(e) {
  let body = {};
  try { body = JSON.parse(e.postData.contents); } catch (err) { return json({ ok: false, error: 'Neplatná požiadavka.' }); }
  return json(runApi(body.fn, body.args));
}

function runApi(fn, args) {
  if (!Object.prototype.hasOwnProperty.call(API, fn)) return { ok: false, error: 'Neznáma akcia.' };
  try { return { ok: true, result: API[fn].apply(null, Array.isArray(args) ? args : []) }; }
  catch (err) { return { ok: false, error: String((err && err.message) || err) }; }
}

function json(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

function doGet(e) {
  const p = (e && e.parameter) || {};
  if (p.api) {
    let args = [];
    try { args = p.args ? JSON.parse(p.args) : []; } catch (err) { return json({ ok: false, error: 'Neplatná požiadavka.' }); }
    const out = runApi(p.api, args);
    // JSONP pre prehliadače, ktoré blokujú bežné požiadavky na iné stránky.
    if (p.callback && /^[A-Za-z_$][\w$]{0,40}$/.test(p.callback)) {
      return ContentService.createTextOutput(p.callback + '(' + JSON.stringify(out) + ');').setMimeType(ContentService.MimeType.JAVASCRIPT);
    }
    return json(out);
  }
  // Pôvodná verzia appky priamo zo skriptu, ak súbor Index v projekte existuje.
  const t = HtmlService.createTemplateFromFile('Index');
  t.boot = JSON.stringify(getData()).replace(/</g, '\\u003c');
  return t.evaluate()
    .setTitle('Stužková 4.NA')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.DEFAULT);
}

function getVersion() {
  return PropertiesService.getScriptProperties().getProperty('VERSION') || '0';
}

function getData() {
  const ss = book();
  const data = {};
  Object.keys(SHEETS).forEach(function (kind) {
    const def = SHEETS[kind];
    const sh = ss.getSheetByName(def.name);
    const n = sh.getLastRow() - 1;
    data[kind] = n > 0 ? sh.getRange(2, 1, n, def.cols.length).getDisplayValues()
      .filter(function (r) { return r[0]; })
      .map(function (r) { const o = {}; def.cols.forEach(function (c, i) { o[c] = r[i]; }); return o; }) : [];
  });
  data.settings = {};
  ss.getSheetByName(SETTINGS_SHEET).getDataRange().getDisplayValues().forEach(function (r) { if (r[0]) data.settings[r[0]] = r[1]; });
  const props = PropertiesService.getScriptProperties();
  data.updated = Number(props.getProperty('UPDATED') || 0);
  data.log = recentLog(ss, 30);
  return { data: data, version: getVersion() };
}

// Posledné zmeny pre „Čo je nové“, najnovšie prvé.
function recentLog(ss, n) {
  const sh = ss.getSheetByName(LOG_SHEET);
  const last = sh.getLastRow();
  if (last < 2) return [];
  const from = Math.max(2, last - n + 1);
  return sh.getRange(from, 1, last - from + 1, 5).getValues().reverse().map(function (r) {
    return { at: r[0] instanceof Date ? r[0].getTime() : 0, by: String(r[1]), part: String(r[2]), action: String(r[3]), item: String(r[4]).split(' | ')[0] };
  });
}

function saveItem(kind, item, by) {
  const def = SHEETS[kind];
  if (!def || !item || !item.id) throw new Error('Neplatná položka.');
  return withLock(function () {
    const sh = book().getSheetByName(def.name);
    const r = findRow(sh, item.id);
    // Pri úprave sa prepíšu len poslané polia, aby si dvaja ľudia neprepísali navzájom iné údaje toho istého riadku.
    const old = r ? sh.getRange(r, 1, 1, def.cols.length).getDisplayValues()[0] : [];
    const row = def.cols.map(function (c, i) { return Object.prototype.hasOwnProperty.call(item, c) || !r ? clean(item[c]) : old[i]; });
    if (r) sh.getRange(r, 1, 1, row.length).setValues([row]);
    else sh.appendRow(row);
    log(by, def.name, r ? 'úprava' : 'pridanie', row.slice(1).join(' | '));
  });
}

function deleteItem(kind, id, by) {
  const def = SHEETS[kind];
  if (!def || !id) throw new Error('Neplatná položka.');
  return withLock(function () {
    const sh = book().getSheetByName(def.name);
    const r = findRow(sh, id);
    if (!r) return;
    const old = sh.getRange(r, 1, 1, def.cols.length).getDisplayValues()[0];
    sh.deleteRow(r);
    log(by, def.name, 'zmazanie', old.slice(1).join(' | '));
  });
}

function saveSettings(s, by) {
  return withLock(function () {
    const sh = book().getSheetByName(SETTINGS_SHEET);
    sh.getRange(1, 1, SETTING_KEYS.length, 2).setValues(SETTING_KEYS.map(function (k) { return [k, clean(s && s[k])]; }));
    log(by, SETTINGS_SHEET, 'úprava', SETTING_KEYS.map(function (k) { return clean(s && s[k]); }).join(' | '));
  });
}

/* ---------- Pomocné funkcie ---------- */

let schemaChecked = false;
function book() {
  const id = PropertiesService.getScriptProperties().getProperty('SHEET_ID');
  if (!id) throw new Error('Najprv spusti funkciu setup().');
  const ss = SpreadsheetApp.openById(id);
  if (!schemaChecked) { ensureSchema(ss); schemaChecked = true; }
  return ss;
}

function createSheet(ss, kind) {
  const def = SHEETS[kind];
  const sh = ss.insertSheet(def.name);
  sh.getRange(1, 1, 1, def.cols.length).setValues([def.labels]).setFontWeight('bold');
  sh.getRange(1, 1, 1000, def.cols.length).setNumberFormat('@'); // všetko ako text, nech Sheets neprepisuje dátumy a čísla
  const rows = (SEED[kind] || []).map(function (it) { return def.cols.map(function (c) { return clean(it[c]); }); });
  if (rows.length) sh.getRange(2, 1, rows.length, def.cols.length).setValues(rows);
  sh.setFrozenRows(1);
  if (kind !== 'seats') sh.hideColumns(1);
}

// Po aktualizácii skriptu doplní nové hárky a stĺpce do existujúcej tabuľky, údaje nechá tak.
function ensureSchema(ss) {
  const missing = Object.keys(SHEETS).filter(function (kind) {
    const sh = ss.getSheetByName(SHEETS[kind].name);
    return !sh || sh.getLastColumn() < SHEETS[kind].cols.length;
  });
  if (!missing.length) return;
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    missing.forEach(function (kind) {
      const def = SHEETS[kind];
      const sh = ss.getSheetByName(def.name);
      if (!sh) { createSheet(ss, kind); return; }
      if (sh.getLastColumn() >= def.cols.length) return;
      sh.getRange(1, 1, 1, def.cols.length).setValues([def.labels]).setFontWeight('bold');
      sh.getRange(1, 1, Math.max(1000, sh.getMaxRows()), def.cols.length).setNumberFormat('@');
    });
  } finally { lock.releaseLock(); }
}

function withLock(fn) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try { fn(); bump(); } finally { lock.releaseLock(); }
  return getData();
}

function findRow(sh, id) {
  const n = sh.getLastRow() - 1;
  if (n < 1) return 0;
  const ids = sh.getRange(2, 1, n, 1).getDisplayValues();
  for (let i = 0; i < ids.length; i++) if (ids[i][0] === String(id)) return i + 2;
  return 0;
}

function bump() {
  const props = PropertiesService.getScriptProperties();
  props.setProperty('VERSION', String(Number(props.getProperty('VERSION') || 0) + 1));
  props.setProperty('UPDATED', String(Date.now()));
}

function log(by, part, action, details) {
  book().getSheetByName(LOG_SHEET).appendRow([new Date(), clean(by) || 'neznámy', part, action, details]);
}

function clean(v) {
  if (v === null || v === undefined) return '';
  return String(v).slice(0, MAX_LEN).replace(/^[=+\-@]/, "'$&"); // text, nie vzorec
}
