/**
 * ASML Tyre Register — Apps Script backend (Mapon / track-loc8 edition)
 * ------------------------------------------------------------
 * Lives inside a Google Sheet (Extensions → Apps Script).
 *
 *  • Stores vehicles, tyres and settings in Sheet tabs
 *  • syncFleet()  pulls every vehicle + odometer from the tracker's Mapon API
 *                 (runs hourly once installTriggers() has been run)
 *  • doGet / doPost expose a small JSON API that the web tool calls
 *
 * ONE-TIME SETUP
 *  1. Project Settings (gear icon) → Script properties → add:
 *       TRACKER_HASH   your API key from the tracker
 *       API_TOKEN      any long random string (the web tool sends it)
 *     optional:
 *       TRACKER_URL    default https://myfleet.track-loc8.com
 *  2. Run setup()         → creates the tabs (or adds any new columns)
 *  3. Run testTracker()   → Execution log shows the vehicles it can see
 *  4. Deploy → New deployment → Web app → Execute as Me, Access: Anyone
 *     → copy the URL ending in /exec into the web tool
 *  5. Run installTriggers() → hourly sync
 *
 * Settings tab:
 *   odoSource        auto | gps | can   (auto = CAN odometer when the vehicle
 *                                        reports one, otherwise GPS mileage)
 *   gpsMileageUnit   m | km             (Mapon reports GPS mileage in metres)
 */

const TRACKER_URL_DEFAULT = 'https://myfleet.track-loc8.com';
const ODO_HISTORY_KEEP = 120;

const COLS = {
  Vehicles: ['id','trackerId','name','plate','vin','imei','icon','type','brand','model','trailerAxles','hasSpare',
             'odometer','odoSource','lastActive','moveState','moveSince','group','source','odoHistory','updatedAt'],
  Tyres:    ['id','status','vehicleId','position','serial','dot','brand','model','size','newDepth','minDepth',
             'lastDepth','lastPressure','fittedDate','fittedOdo','cost','condition','notes','lastInspected',
             'inspections','lastVehicleId','lastPosition','removedDate','removedOdo','removalReason','totalKm','updatedAt'],
  Settings: ['key','value'],
  SyncLog:  ['time','result','units','added','updated','canOdometers','note'],
  CCS:      ['id','status','statusCode','issuedAt','issuedOn','expiresOn','permit','farmerId','farmerName','fieldId','plot',
             'plate','plateKey','transporterCode','transporterName','weighState','weighmentId','weighedAt','weighedDatetime',
             'weighedOn','netKg','weighbridgeId','grossWbId','tareWbId','terminal','lastChecked','updatedAt'],
  WBLog:    ['time','result','calls','newCcs','rechecked','closed','cursor','pending','note'],
};
const NUM_FIELDS  = new Set(['trailerAxles','odometer','newDepth','minDepth','lastDepth','lastPressure','fittedOdo','cost','removedOdo','totalKm','netKg']);
const JSON_FIELDS = new Set(['odoHistory','inspections']);
const BOOL_FIELDS = new Set(['hasSpare','terminal']);
const COL_FOR = { vehicles:'Vehicles', tyres:'Tyres' };
const DEFAULT_SETTINGS = { odoSource:'auto', gpsMileageUnit:'m', tyreIdPrefix:'ASML' };

// Mapon vehicle icon → tyre-register vehicle type (only used for NEW vehicles)
const ICON_TYPE = {
  tractor:'tractor',
  truck:'prime_mover', 'truck-heavy':'prime_mover', 'truck-no-trailer':'prime_mover', cistern:'prime_mover',
  car:'light', minibus:'light', bus:'light', emergency:'light',
  loader:'loader',
  excavator:'tracked', bulldozer:'tracked', paver:'tracked', roller:'tracked', crane:'tracked', 'stationary-crane':'tracked',
};
// Label hints that override the icon (your tracker tags wheel loaders as excavator/bulldozer)
const LABEL_TYPE = [
  [/\b(wheel\s*)?loader\b|\bZL\d|\bLW\d|\bCLG8|\bSDLG\b|\bL9\d\d\b/i, 'loader'],
  [/\bgrader\b|\bGR\d{3}|\b1[24]0[HKM]\b/i, 'grader'],
];
const TRAILER_DEFAULT = { prime_mover:3, tractor:1 };
// Two-word brands, so "Deutz Fahr AGROLUX 4.8" splits as brand "Deutz Fahr", model "AGROLUX 4.8"
const TWO_WORD_BRANDS = ['deutz fahr','massey ferguson','new holland','john deere','case ih','land rover','mercedes benz','tata motors'];

function guessType(u) {
  const label = String(u.label || '');
  for (const [re, t] of LABEL_TYPE) if (re.test(label)) return t;
  return ICON_TYPE[String(u.icon || '').toLowerCase()] || '';
}
function splitLabel(label) {
  const t = String(label || '').trim().replace(/\s+/g, ' ');
  if (!t) return { brand: '', model: '' };
  const lower = t.toLowerCase();
  for (const b of TWO_WORD_BRANDS) if (lower.startsWith(b + ' ') || lower === b) return { brand: t.slice(0, b.length), model: t.slice(b.length).trim() };
  const parts = t.split(' ');
  if (parts.length === 1) return { brand: '', model: t };
  return { brand: parts[0], model: parts.slice(1).join(' ') };
}

/* ============================================================
   Setup & triggers
   ============================================================ */
function setup() {
  Object.keys(COLS).forEach(name => {
    const s = sheet(name);
    const hdr = headerOf(s);
    const missing = COLS[name].filter(c => hdr.indexOf(c) < 0);
    if (missing.length) {
      s.getRange(1, hdr.length + 1, 1, missing.length).setValues([missing]).setFontWeight('bold');
      Logger.log(name + ': added column(s) ' + missing.join(', '));
    }
  });
  const s = getSettings();
  const missing = {};
  Object.keys(DEFAULT_SETTINGS).forEach(k => { if (!(k in s)) missing[k] = DEFAULT_SETTINGS[k]; });
  if (Object.keys(missing).length) setSettings(missing);
  Logger.log('Setup complete. Tabs: ' + Object.keys(COLS).join(', '));
}

function installTriggers() {
  ScriptApp.getProjectTriggers().forEach(t => { if (['syncFleet', 'syncWeighbridge'].indexOf(t.getHandlerFunction()) >= 0) ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('syncFleet').timeBased().everyHours(1).create();
  Logger.log('Hourly syncFleet trigger installed.');
  if (PropertiesService.getScriptProperties().getProperty('WB_KEY')) {
    ScriptApp.newTrigger('syncWeighbridge').timeBased().everyMinutes(30).create();
    Logger.log('syncWeighbridge trigger installed (every 30 minutes).');
  } else Logger.log('WB_KEY not set — weighbridge trigger not installed. Add it and run installTriggers again.');
}

function removeTriggers() {
  ScriptApp.getProjectTriggers().forEach(t => ScriptApp.deleteTrigger(t));
  Logger.log('All triggers removed.');
}

/* ============================================================
   Tracker (Mapon API)
   ============================================================ */
function testTracker() {
  const units = fetchUnits();
  Logger.log('Vehicles visible to this API key: ' + units.length);
  if (!units.length) return;
  const u = units[0];
  Logger.log('First vehicle, key fields: ' + JSON.stringify({
    unit_id: u.unit_id, label: u.label, number: u.number, vin: u.vin, icon: u.icon, type: u.type,
    mileage_raw: u.mileage, can_odom: u.can && u.can.odom ? u.can.odom.value : null,
    last_update: u.last_update, movement_state: u.movement_state || u.state,
  }));
  const withCan = units.filter(x => x.can && x.can.odom && toNum(x.can.odom.value) > 0).length;
  Logger.log('Vehicles reporting a CAN odometer: ' + withCan + ' of ' + units.length);
  Logger.log('--- All vehicles (number | label | icon → type | GPS km | CAN km | last update) ---');
  units.slice(0, 60).forEach(x => {
    const gps = toNum(x.mileage); const can = x.can && x.can.odom ? toNum(x.can.odom.value) : null;
    Logger.log([x.number || '', x.label || x.vehicle_title || '', (x.icon || '') + ' → ' + (guessType(x) || '?'), gps == null ? '' : Math.round(gps / 1000), can == null ? '' : Math.round(can), x.last_update || ''].join(' | '));
  });
  const counts = {}; units.forEach(x => { const t = guessType(x) || 'unset'; counts[t] = (counts[t] || 0) + 1; });
  Logger.log('Type guesses: ' + JSON.stringify(counts) + '  (unset = set by hand in the tool)');
  if (units.length > 60) Logger.log('… and ' + (units.length - 60) + ' more');
}

function syncFleet() {
  let units;
  try { units = fetchUnits(); }
  catch (err) { logSync('error', 0, 0, 0, 0, String(err.message || err)); throw err; }

  const settings = getSettings();
  const gpsUnit = settings.gpsMileageUnit === 'km' ? 'km' : 'm';
  const pref = ['gps', 'can'].indexOf(settings.odoSource) >= 0 ? settings.odoSource : 'auto';
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  let added = 0, updated = 0, canCount = 0;
  try {
    const vehicles = readAll('Vehicles');
    const byTracker = new Map(vehicles.filter(v => v.trackerId).map(v => [String(v.trackerId), v]));
    const toWrite = [];
    units.forEach(u => {
      const tid = String(u.unit_id == null ? '' : u.unit_id).trim();
      if (!tid) return;
      let v = byTracker.get(tid);
      const isNew = !v;
      if (isNew) {
        const type = guessType(u);
        v = { id: uid('v_'), trackerId: tid, type, source: 'tracker', hasSpare: true, trailerAxles: TRAILER_DEFAULT[type] || 0 };
      }
      // In this tracker "number" is the identifier operators use and "label" is make/model
      v.name = String(u.number || u.vehicle_title || u.label || ('Unit ' + tid));
      if (u.number) v.plate = String(u.number);
      if (!v.brand && !v.model && u.label) { const bm = splitLabel(u.label); v.brand = bm.brand; v.model = bm.model; }
      if (u.vin) v.vin = String(u.vin);
      if (u.icon) v.icon = String(u.icon);
      if (u.device && u.device.imei) v.imei = String(u.device.imei);
      const when = parseTrackerDate(u.last_update);
      if (when) v.lastActive = when;
      const ms = u.movement_state || u.state;
      if (ms && ms.name) { v.moveState = String(ms.name); v.moveSince = parseTrackerDate(ms.start) || ''; }

      const gps = toNum(u.mileage);
      const gpsKm = gps == null ? null : (gpsUnit === 'm' ? gps / 1000 : gps);
      const can = (u.can && u.can.odom) ? toNum(u.can.odom.value) : null;
      let odo = null, src = '';
      if (pref === 'can')      { if (can != null && can > 0) { odo = can; src = 'can'; } }
      else if (pref === 'gps') { if (gpsKm != null) { odo = gpsKm; src = 'gps'; } }
      else { if (can != null && can > 0) { odo = can; src = 'can'; } else if (gpsKm != null) { odo = gpsKm; src = 'gps'; } }
      if (odo != null) { odo = Math.round(odo); v.odometer = odo; v.odoSource = src; pushOdoHistory(v, odo, when); if (src === 'can') canCount++; }

      v.updatedAt = new Date().toISOString();
      if (isNew) added++; else updated++;
      toWrite.push(v);
    });
    upsertMany('Vehicles', toWrite);
    setSettings({ lastSync: new Date().toISOString(), lastSyncUnits: String(units.length) });
  } finally {
    lock.releaseLock();
  }
  logSync('ok', units.length, added, updated, canCount, '');
  return { devices: units.length, added, updated, canOdometers: canCount, odoField: 'mapon unit/list', odoUnit: 'km' };
}

/** Calls Mapon unit/list with CAN + device details and returns the units array. */
function fetchUnits() {
  const P = PropertiesService.getScriptProperties();
  const base = (P.getProperty('TRACKER_URL') || TRACKER_URL_DEFAULT).replace(/\/+$/, '');
  const key = P.getProperty('TRACKER_HASH');
  if (!key) throw new Error('Add TRACKER_HASH (your tracker API key) under Project Settings → Script properties.');
  const url = base + '/api/v1/unit/list.json?key=' + encodeURIComponent(key) + '&include[]=can&include[]=device';
  const res = UrlFetchApp.fetch(url, { muteHttpExceptions: true, followRedirects: true });
  const code = res.getResponseCode(), text = res.getContentText();
  let payload;
  try { payload = JSON.parse(text); }
  catch (e) { throw new Error('Tracker returned non-JSON (HTTP ' + code + '): ' + text.replace(/\s+/g, ' ').slice(0, 160)); }
  if (payload && payload.error) throw new Error('Tracker error ' + payload.error.code + ': ' + payload.error.msg);
  const units = payload && payload.data && Array.isArray(payload.data.units) ? payload.data.units : null;
  if (!units) throw new Error('Unexpected tracker response: ' + text.slice(0, 160));
  return units;
}

function pushOdoHistory(v, km, dateISO) {
  const d = (dateISO || new Date().toISOString()).slice(0, 10);
  const h = Array.isArray(v.odoHistory) ? v.odoHistory.slice() : [];
  if (h.length && h[h.length - 1].d === d) h[h.length - 1] = { d, o: km }; else h.push({ d, o: km });
  v.odoHistory = h.slice(-ODO_HISTORY_KEEP);
}

function parseTrackerDate(s) {
  if (!s) return null;
  const str = String(s).trim();
  let d = new Date(str.replace(' ', 'T'));
  if (isNaN(d.getTime())) d = new Date(str);
  return isNaN(d.getTime()) ? null : d.toISOString();
}

function logSync(result, units, added, updated, canCount, note) {
  sheet('SyncLog').appendRow([new Date().toISOString(), result, units, added, updated, canCount, note]);
}

/* ============================================================
   Web app API
   ============================================================ */
function doGet(e) {
  const p = (e && e.parameter) || {};
  const authErr = checkToken(p.token);
  if (authErr) return json({ ok: false, error: authErr });
  try {
    const action = p.action || 'all';
    if (action === 'ping') return json({ ok: true, app: 'asml-tyre-register' });
    if (action === 'sync') return json(Object.assign({ ok: true }, syncFleet()));
    if (action === 'wbsync') return json(Object.assign({ ok: true }, syncWeighbridge()));
    if (action === 'weighings') {
      const today = new Date().toISOString().slice(0, 10);
      const from = /^\d{4}-\d{2}-\d{2}$/.test(p.from || '') ? p.from : new Date(Date.now() - 6 * 86400000).toISOString().slice(0, 10);
      const to = /^\d{4}-\d{2}-\d{2}$/.test(p.to || '') ? p.to : today;
      const rows = readAll('CCS').filter(r => r.netKg != null && r.weighedOn >= from && r.weighedOn <= to)
        .map(r => ({ ccs: r.id, plate: r.plate || '', plateKey: r.plateKey || '', tr: r.transporterName || '', trc: r.transporterCode || '',
                     wb: r.weighedAt || '', dt: r.weighedDatetime || '', on: r.weighedOn || '', kg: r.netKg, farmer: r.farmerName || '',
                     plot: r.plot || '', permit: r.permit || '', issuedAt: r.issuedAt || '', status: r.statusCode || '' }));
      return json({ ok: true, from, to, rows });
    }
    if (action === 'all') {
      const settings = getSettings();
      return json({
        ok: true,
        app: 'asml-tyre-register',
        vehicles: readAll('Vehicles'),
        tyres: readAll('Tyres'),
        settings,
        lastSync: settings.lastSync || null,
        noToken: !PropertiesService.getScriptProperties().getProperty('API_TOKEN'),
      });
    }
    return json({ ok: false, error: 'Unknown action: ' + action });
  } catch (err) {
    return json({ ok: false, error: String(err.message || err) });
  }
}

function doPost(e) {
  let body;
  try { body = JSON.parse(e.postData.contents); }
  catch (err) { return json({ ok: false, error: 'Body must be JSON' }); }
  const authErr = checkToken(body.token);
  if (authErr) return json({ ok: false, error: authErr });
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
    const col = COL_FOR[body.col];
    if (body.action === 'save') {
      if (!col || !body.data || !body.data.id) return json({ ok: false, error: 'save needs col and data.id' });
      upsertMany(col, [body.data]); return json({ ok: true });
    }
    if (body.action === 'bulkSave') {
      if (!col || !Array.isArray(body.items)) return json({ ok: false, error: 'bulkSave needs col and items' });
      upsertMany(col, body.items.filter(o => o && o.id)); return json({ ok: true, count: body.items.length });
    }
    if (body.action === 'delete') {
      if (!col || !body.id) return json({ ok: false, error: 'delete needs col and id' });
      removeRow(col, body.id); return json({ ok: true });
    }
    if (body.action === 'settings') {
      setSettings(body.data || {}); return json({ ok: true });
    }
    if (body.action === 'reserveTyreIds') {
      const count = Math.max(1, Math.min(50, parseInt(body.count, 10) || 5));
      return json({ ok: true, ids: reserveTyreIds(count) });
    }
    return json({ ok: false, error: 'Unknown action: ' + body.action });
  } catch (err) {
    return json({ ok: false, error: String(err.message || err) });
  } finally {
    try { lock.releaseLock(); } catch (e2) {}
  }
}

/** Allocate the next fleet tyre IDs (PREFIX-YY-NNNN, sequence restarts each year). Called inside doPost's lock. */
function reserveTyreIds(count) {
  const S = getSettings();
  const prefix = String(S.tyreIdPrefix || 'ASML').toUpperCase().replace(/[^A-Z0-9]/g, '') || 'ASML';
  const yy = String(new Date().getFullYear()).slice(-2);
  const key = 'tyreIdSeq_' + yy;
  let seq = parseInt(S[key], 10) || 0;
  const re = new RegExp('^' + prefix + '-' + yy + '-(\\d+)$');
  readAll('Tyres').forEach(t => { const m = re.exec(String(t.serial || '').toUpperCase()); if (m) seq = Math.max(seq, parseInt(m[1], 10)); });
  const ids = [];
  for (let i = 1; i <= count; i++) ids.push(prefix + '-' + yy + '-' + String(seq + i).padStart(4, '0'));
  const upd = {}; upd[key] = String(seq + count); upd.tyreIdPrefix = prefix;
  setSettings(upd);
  return ids;
}

function checkToken(token) {
  const want = PropertiesService.getScriptProperties().getProperty('API_TOKEN');
  if (want && token !== want) return 'Invalid or missing API token';
  return null;
}

function json(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

/* ============================================================
   Sheet access
   ============================================================ */
function ss() { return SpreadsheetApp.getActiveSpreadsheet(); }

function sheet(name) {
  let s = ss().getSheetByName(name);
  if (!s) {
    s = ss().insertSheet(name);
    s.getRange(1, 1, 1, COLS[name].length).setValues([COLS[name]]).setFontWeight('bold');
    s.getRange('A:Z').setNumberFormat('@');   // keep ids, dates and JSON as plain text
    s.setFrozenRows(1);
  }
  return s;
}

function headerOf(s) {
  const lastCol = Math.max(s.getLastColumn(), 1);
  return s.getRange(1, 1, 1, lastCol).getValues()[0].map(String).filter(h => h !== '');
}

function readAll(name) {
  const s = sheet(name);
  const vals = s.getDataRange().getValues();
  if (vals.length < 2) return [];
  const hdr = vals[0].map(String);
  const rows = [];
  for (let i = 1; i < vals.length; i++) {
    const r = vals[i];
    if (r.every(c => c === '' || c === null)) continue;
    const o = {};
    hdr.forEach((h, j) => { if (!h) return; const v = fromCell(h, r[j]); if (v !== undefined) o[h] = v; });
    if (o.id !== undefined || name === 'Settings') rows.push(o);
  }
  return rows;
}

function upsertMany(name, objs) {
  if (!objs.length) return;
  const s = sheet(name);
  const hdr = headerOf(s);
  const known = new Set(COLS[name]);
  const last = s.getLastRow();
  const data = last > 1 ? s.getRange(2, 1, last - 1, hdr.length).getValues() : [];
  const idx = new Map(data.map((r, i) => [String(r[0]), i]));
  objs.forEach(o => {
    const id = String(o.id);
    const i = idx.get(id);
    if (i !== undefined) {
      const row = data[i];
      hdr.forEach((h, j) => { if (known.has(h) && h in o) row[j] = toCell(h, o[h]); });
    } else {
      const row = hdr.map(h => known.has(h) ? toCell(h, o[h]) : '');
      data.push(row); idx.set(id, data.length - 1);
    }
  });
  s.getRange(2, 1, data.length, hdr.length).setValues(data);
}

function removeRow(name, id) {
  const s = sheet(name);
  const last = s.getLastRow();
  if (last < 2) return;
  const ids = s.getRange(2, 1, last - 1, 1).getValues().map(r => String(r[0]));
  const i = ids.indexOf(String(id));
  if (i >= 0) s.deleteRow(i + 2);
}

function getSettings() {
  const o = {};
  readAll('Settings').forEach(r => { if (r.key !== undefined) o[r.key] = r.value === undefined ? '' : r.value; });
  return o;
}

function setSettings(obj) {
  const s = sheet('Settings');
  const last = s.getLastRow();
  const data = last > 1 ? s.getRange(2, 1, last - 1, 2).getValues() : [];
  const idx = new Map(data.map((r, i) => [String(r[0]), i]));
  Object.keys(obj).forEach(k => {
    const v = obj[k] == null ? '' : String(obj[k]);
    const i = idx.get(k);
    if (i !== undefined) data[i][1] = v; else { data.push([k, v]); idx.set(k, data.length - 1); }
  });
  if (data.length) s.getRange(2, 1, data.length, 2).setValues(data);
}

function fromCell(h, v) {
  if (v === '' || v === null || v === undefined) return undefined;
  if (v instanceof Date) return v.toISOString();
  if (JSON_FIELDS.has(h)) { try { return JSON.parse(v); } catch (e) { return undefined; } }
  if (NUM_FIELDS.has(h)) { const n = Number(v); return isNaN(n) ? undefined : n; }
  if (BOOL_FIELDS.has(h)) return String(v).toLowerCase() === 'true';
  return typeof v === 'string' ? v : String(v);
}

function toCell(h, v) {
  if (v === undefined || v === null) return '';
  if (JSON_FIELDS.has(h)) return JSON.stringify(v);
  if (BOOL_FIELDS.has(h)) return v ? 'TRUE' : 'FALSE';
  return v;
}

function toNum(v) { if (v === '' || v == null) return null; const n = Number(v); return isNaN(n) ? null : n; }
function uid(p) { return p + Date.now().toString(36) + Math.random().toString(36).slice(2, 7); }

/* ============================================================
   Weighbridge (Synergy Bravo CCS lookup) sync
   Script properties: WB_KEY (the Synergy API key)
   optional: WB_URL (default below), WB_START_CCS (first number to load; default 59843)
   Settings tab (created with defaults on first run):
     wbBudget            calls per run (300)
     wbGapStop           consecutive missing numbers that mark the end (30)
     wbRecentDays        "recent" = issued within this many days (7)
     wbRecentEveryHours  re-check recent open CCS this often (6)
     wbOldEveryHours     re-check older open CCS this often (24)
   ============================================================ */
const WB_URL_DEFAULT = 'https://angatasugar.synergybravo.com/api/ccslookup.php';
const WB_AUTH = {
  'header X-API-Key':             (u, k) => ({ url: u, headers: { 'X-API-Key': k } }),
  'header Authorization: Bearer': (u, k) => ({ url: u, headers: { Authorization: 'Bearer ' + k } }),
  'header api-key':               (u, k) => ({ url: u, headers: { 'api-key': k } }),
  'header X-Auth-Token':          (u, k) => ({ url: u, headers: { 'X-Auth-Token': k } }),
  'query api_key':                (u, k) => ({ url: u + '&api_key=' + encodeURIComponent(k) }),
  'query key':                    (u, k) => ({ url: u + '&key=' + encodeURIComponent(k) }),
  'query apikey':                 (u, k) => ({ url: u + '&apikey=' + encodeURIComponent(k) }),
  'query token':                  (u, k) => ({ url: u + '&token=' + encodeURIComponent(k) }),
  'no key at all':                (u, k) => ({ url: u }),
};
const WB_DEFAULTS = { wbBudget: '300', wbGapStop: '30', wbRecentDays: '7', wbRecentEveryHours: '6', wbOldEveryHours: '24' };
const WB_TERMINAL = /closed|expired|cancel|void|reject|complete/;

function wbConfig() {
  const P = PropertiesService.getScriptProperties();
  const key = P.getProperty('WB_KEY');
  if (!key) throw new Error('Add WB_KEY (the Synergy weighbridge API key) under Project Settings → Script properties.');
  return { base: P.getProperty('WB_URL') || WB_URL_DEFAULT, key, start: parseInt(P.getProperty('WB_START_CCS') || '59843', 10) };
}
function wbReq(cfg, auth, n) {
  const r = WB_AUTH[auth](cfg.base + '?ccs_number=' + n, cfg.key);
  r.muteHttpExceptions = true; r.followRedirects = true;
  return r;
}
function wbParse(txt) { try { return JSON.parse(txt); } catch (e) { return { ok: false, raw: String(txt).slice(0, 120) }; } }
function wbPlateKey(s) { return String(s || '').toUpperCase().replace(/[^A-Z0-9]/g, ''); }

/** Try each known way of sending the key against a known-good CCS number. */
function wbDetectAuth(cfg, sample) {
  for (const label of Object.keys(WB_AUTH)) {
    const r = wbReq(cfg, label, sample);
    const j = wbParse(UrlFetchApp.fetch(r.url, r).getContentText());
    if (j && j.ok === true) return label;
  }
  return null;
}

/** Fetch many CCS numbers in parallel batches of 10. Returns { number: json }. */
function wbFetchMany(cfg, auth, nums) {
  const out = {};
  for (let i = 0; i < nums.length; i += 10) {
    const batch = nums.slice(i, i + 10);
    try {
      const res = UrlFetchApp.fetchAll(batch.map(n => wbReq(cfg, auth, n)));
      res.forEach((r, k) => { out[batch[k]] = wbParse(r.getContentText()); });
    } catch (e) {
      batch.forEach(n => { out[n] = { ok: false, error: String(e.message || e) }; });
    }
  }
  return out;
}

function wbRow(n, j) {
  const w = j.weighing || {}, a = j.assignment || {}, f = j.farmer || {}, fd = j.field || {};
  const sc = String(j.status_code || j.status || '').toLowerCase();
  const today = new Date().toISOString().slice(0, 10);
  const terminal = WB_TERMINAL.test(sc) || (!!j.expires_on && String(j.expires_on) < today);
  const now = new Date().toISOString();
  return {
    id: String(n), status: j.status || '', statusCode: j.status_code || '', issuedAt: j.issued_at || '', issuedOn: j.issued_on || '',
    expiresOn: j.expires_on || '', permit: j.permit_number || '', farmerId: f.id || '', farmerName: f.name || '', fieldId: fd.id || '',
    plot: fd.plot_number || '', plate: a.vehicle_plate || '', plateKey: wbPlateKey(a.vehicle_plate), transporterCode: a.transporter_code || '',
    transporterName: a.transporter_name || '', weighState: w.state || '', weighmentId: w.weighment_id || '', weighedAt: w.weighed_at || '',
    weighedDatetime: w.weighed_datetime || '', weighedOn: w.weighed_on || '', netKg: toNum(w.net_weight), weighbridgeId: w.weighbridge_id || '',
    grossWbId: w.gross_weighbridge_id || '', tareWbId: w.tare_weighbridge_id || '', terminal, lastChecked: now, updatedAt: now,
  };
}

function syncWeighbridge() {
  const cfg = wbConfig();
  const S = getSettings();
  const missing = {}; Object.keys(WB_DEFAULTS).forEach(k => { if (!S[k]) { S[k] = WB_DEFAULTS[k]; missing[k] = WB_DEFAULTS[k]; } });
  if (Object.keys(missing).length) setSettings(missing);
  let auth = (S.wbAuth && WB_AUTH[S.wbAuth]) ? S.wbAuth : null;
  if (!auth) {
    auth = wbDetectAuth(cfg, cfg.start);
    if (!auth) { wbLog('error', 0, 0, 0, 0, '', '', 'No known way of sending the key was accepted by the CCS lookup'); throw new Error('CCS lookup: could not authenticate'); }
    setSettings({ wbAuth: auth });
  }
  const budget = Math.max(20, parseInt(S.wbBudget, 10) || 300), gapStop = parseInt(S.wbGapStop, 10) || 30;
  const lock = LockService.getScriptLock(); lock.waitLock(30000);
  let calls = 0, found = 0, rechecked = 0, closed = 0, pending = 0;
  let cursor = parseInt(S.wbCursor, 10); if (isNaN(cursor)) cursor = cfg.start - 1;
  try {
    const existing = readAll('CCS');
    const byId = new Map(existing.map(r => [String(r.id), r]));
    const toWrite = [];

    // 1) walk forward from the cursor until a run of missing numbers
    let probe = cursor, miss = 0, lastFound = cursor;
    while (calls < budget && miss < gapStop) {
      const nums = []; for (let i = 1; i <= 10; i++) nums.push(probe + i);
      const res = wbFetchMany(cfg, auth, nums); calls += nums.length;
      nums.forEach(n => {
        const j = res[n];
        if (j && j.ok === true) { const row = wbRow(n, j); if (row.terminal && row.netKg != null) closed++; toWrite.push(row); found++; lastFound = n; miss = 0; }
        else miss++;
      });
      probe += 10;
    }
    cursor = lastFound;

    // 2) re-check open CCS that are due
    const nowMs = Date.now();
    const hrs = iso => iso ? (nowMs - new Date(iso).getTime()) / 3600000 : 1e9;
    const recentCut = new Date(nowMs - (parseInt(S.wbRecentDays, 10) || 7) * 86400000).toISOString().slice(0, 10);
    const rEvery = parseFloat(S.wbRecentEveryHours) || 6, oEvery = parseFloat(S.wbOldEveryHours) || 24;
    const due = existing.filter(r => !r.terminal && ((r.issuedOn || '') >= recentCut ? hrs(r.lastChecked) >= rEvery : hrs(r.lastChecked) >= oEvery))
                        .sort((a, b) => ((a.lastChecked || '') < (b.lastChecked || '') ? -1 : 1));
    const slice = due.slice(0, Math.max(0, budget - calls)).map(r => parseInt(r.id, 10));
    pending = due.length - slice.length;
    for (let i = 0; i < slice.length; i += 10) {
      const nums = slice.slice(i, i + 10);
      const res = wbFetchMany(cfg, auth, nums); calls += nums.length;
      nums.forEach(n => {
        const j = res[n], prev = byId.get(String(n));
        if (j && j.ok === true) { const row = wbRow(n, j); if (row.terminal && row.netKg != null && !(prev && prev.terminal)) closed++; toWrite.push(row); rechecked++; }
        else if (prev) { prev.lastChecked = new Date().toISOString(); toWrite.push(prev); rechecked++; }
      });
    }
    upsertMany('CCS', toWrite);

    // 3) summary the tool can show without reading the whole tab
    const merged = new Map(byId); toWrite.forEach(r => merged.set(String(r.id), r));
    let openCount = 0; merged.forEach(r => { if (!r.terminal) openCount++; });
    setSettings({ wbCursor: String(cursor), wbLastSync: new Date().toISOString(), wbTotal: String(merged.size), wbOpen: String(openCount), wbPending: String(pending) });
  } finally {
    lock.releaseLock();
  }
  wbLog('ok', calls, found, rechecked, closed, String(cursor), String(pending), '');
  return { calls, newCcs: found, rechecked, closed, cursor, pending };
}

function wbLog(result, calls, newCcs, rechecked, closed, cursor, pending, note) {
  sheet('WBLog').appendRow([new Date().toISOString(), result, calls, newCcs, rechecked, closed, cursor, pending, note]);
}

/** Diagnostic: auth style, sequence check, newest number, daily volume, statuses and plate spellings. */
function testWeighbridge() {
  const cfg = wbConfig();
  const sample = cfg.start;
  Logger.log('--- 1. How is the key sent? ---');
  const auth = wbDetectAuth(cfg, sample);
  if (!auth) { Logger.log('None of the known key styles was accepted. Ask the vendor how the key should be sent.'); return; }
  Logger.log('Works with: ' + auth + ' (saved to Settings as wbAuth)');
  setSettings({ wbAuth: auth });
  const get = n => wbFetchMany(cfg, auth, [n])[n];
  const brief = j => j && j.ok
    ? [j.status, (j.assignment && j.assignment.vehicle_plate) || '-', 'issued ' + j.issued_on, 'weighed ' + ((j.weighing && j.weighing.weighed_on) || '-'), ((j.weighing && j.weighing.net_weight) || '-') + ' kg'].join(' | ')
    : JSON.stringify(j).slice(0, 120);
  Logger.log('--- 2. Sample and the next 5 ---');
  const first6 = wbFetchMany(cfg, auth, [0, 1, 2, 3, 4, 5].map(i => sample + i));
  Object.keys(first6).sort().forEach(n => Logger.log(n + ' → ' + brief(first6[n])));
  Logger.log('--- 3. Newest number (stepping 1000 → 100 → 10 → 1) ---');
  let lo = sample;
  [1000, 100, 10, 1].forEach(step => { let guard = 0; while (guard++ < 80) { const j = get(lo + step); if (j && j.ok) lo += step; else break; } });
  const top = get(lo), first = get(sample);
  Logger.log('Highest existing CCS: ' + lo + ' → ' + brief(top));
  Logger.log('A number that does not exist (' + (lo + 500) + ') returns: ' + JSON.stringify(get(lo + 500)).slice(0, 160));
  const days = (first.ok && top.ok) ? Math.max(1, Math.round((new Date(top.issued_on) - new Date(first.issued_on)) / 86400000)) : null;
  Logger.log((lo - sample) + ' numbers between ' + (first.issued_on || '?') + ' and ' + (top.issued_on || '?') + (days ? ' ≈ ' + Math.round((lo - sample) / days) + ' CCS per day' : ''));
  Logger.log('--- 4. Statuses and plate spellings in the latest 40 ---');
  const nums = []; for (let n = Math.max(sample, lo - 39); n <= lo; n++) nums.push(n);
  const res = wbFetchMany(cfg, auth, nums);
  const counts = {}, plates = new Set();
  nums.forEach(n => { const j = res[n]; if (j && j.ok) { const k = j.status_code || j.status; counts[k] = (counts[k] || 0) + 1; if (j.assignment && j.assignment.vehicle_plate) plates.add(j.assignment.vehicle_plate); } });
  Logger.log('Statuses: ' + JSON.stringify(counts));
  Logger.log('Plates as typed at the weighbridge: ' + [...plates].slice(0, 20).join(', '));
  Logger.log('Tip: to avoid loading everything back to ' + sample + ', set WB_START_CCS to a more recent number before the first syncWeighbridge.');
}
