/**
 * Kaizen KPIs — Google Apps Script web app.
 * Paste this into Code.gs (Extensions → Apps Script in the Kaizen KPIs Google Sheet).
 * All data lives in this spreadsheet. People log in with their name and a PIN.
 */
const APP_TITLE = 'Kaizen KPIs';
const OWNER_NAME = 'Ameliah';          // the owner's login name
const OWNER_UID = 'u_owner';
const DATA = 'Data', ACCESS = 'Access';
const CHUNK = 45000;                   // characters per cell (Sheets limit is 50,000)
const MAX_CHUNKS = 20;

/* ---------- web app ---------- */
function doGet(e) {
  if (e && e.parameter && e.parameter.app === 'kova') return HtmlService.createHtmlOutputFromFile('Kova')
    .setTitle('Kova')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  if (e && e.parameter && e.parameter.app === 'pb') return HtmlService.createHtmlOutputFromFile('PB')
    .setTitle('PB Board')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  if (e && e.parameter && e.parameter.app === 'hourly') return HtmlService.createHtmlOutputFromFile('Hourly')
    .setTitle('Hourly KPIs')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  if (e && e.parameter && (e.parameter.app === 'orexis' || e.parameter.app === 'orexishourly')) return orexis_(e.parameter.app === 'orexishourly');
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle(APP_TITLE)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/* ---------- spreadsheet menu (owner setup) ---------- */
function onOpen() {
  SpreadsheetApp.getUi().createMenu(APP_TITLE)
    .addItem('Set or reset owner PIN', 'setOwnerPin')
    .addItem('Turn on daily backups & hourly syncs', 'setupBackups')
    .addItem('Back up now', 'backupNow')
    .addToUi();
}

/* ---------- backups: a dated copy of this sheet every night, last 30 kept ---------- */
const BACKUP_FOLDER = 'Kaizen KPIs backups', BACKUP_KEEP = 30;
function backupFolder_() {
  const it = DriveApp.getFoldersByName(BACKUP_FOLDER);
  return it.hasNext() ? it.next() : DriveApp.createFolder(BACKUP_FOLDER);
}
function backupNow() {
  const ss = ss_(), folder = backupFolder_();
  const stamp = Utilities.formatDate(new Date(), 'Australia/Melbourne', 'yyyy-MM-dd HHmm');
  DriveApp.getFileById(ss.getId()).makeCopy('Kaizen KPIs backup ' + stamp, folder);
  const files = [], it = folder.getFiles();
  while (it.hasNext()) files.push(it.next());
  files.sort((a, b) => b.getDateCreated() - a.getDateCreated());
  files.slice(BACKUP_KEEP).forEach(f => f.setTrashed(true));
  return 'Backed up: ' + stamp;
}
function setupBackups() {
  ScriptApp.getProjectTriggers().filter(t => ['backupNow', 'syncHourly'].indexOf(t.getHandlerFunction()) >= 0).forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('backupNow').timeBased().everyDays(1).atHour(2).inTimezone('Australia/Melbourne').create();
  // hourly tracker -> day log at 12pm, 3:45pm and 9pm (Google runs these within ~15 minutes of the time)
  [[12, 0], [15, 45], [21, 0]].forEach(([h, m]) => ScriptApp.newTrigger('syncHourly').timeBased().everyDays(1).atHour(h).nearMinute(m).inTimezone('Australia/Melbourne').create());
  return backupNow() + ' · hourly syncs on at 12pm, 3:45pm, 9pm';
}

/* ---------- hourly tracker: add each rep's hours for today into their day log ---------- */
const H_FIELDS = ['stops', 'buyers', 'retail', 'upsells', 'upsellRetail', 'medals'];
function mondayOf_(d) { const x = new Date(d + 'T00:00:00Z'); const wd = (x.getUTCDay() + 6) % 7; x.setUTCDate(x.getUTCDate() - wd); return x.toISOString().slice(0, 10); }
function syncHourly(dateArg) {
  const date = typeof dateArg === 'string' ? dateArg : Utilities.formatDate(new Date(), 'Australia/Melbourne', 'yyyy-MM-dd');
  const lock = LockService.getScriptLock(); lock.waitLock(25000);
  try {
    const t = loadTable_(); const out = []; const nowIso = new Date().toISOString();
    Object.keys(t.map).forEach(path => {
      const seg = path.split('/');
      if (seg.length !== 4 || seg[0] !== 'reps' || seg[2] !== 'hourly' || seg[3] !== date) return;
      const doc = t.map[path]; if (doc.del || !doc.data || !doc.data.slots) return;
      const tot = {}; H_FIELDS.forEach(k => tot[k] = 0); let n = 0;
      Object.keys(doc.data.slots).forEach(k => { const x = doc.data.slots[k] || {}; if (H_FIELDS.some(f => Number(x[f]) || 0)) n++; H_FIELDS.forEach(f => tot[f] += Number(x[f]) || 0); });
      const ep = 'reps/' + seg[1] + '/entries/' + date; const cur = t.map[ep] && !t.map[ep].del ? (t.map[ep].data || {}) : {};
      const rep = t.map['reps/' + seg[1]] && t.map['reps/' + seg[1]].data || {};
      const body = Object.assign({}, cur, tot, { date, week: mondayOf_(date), month: date.slice(0, 7), hours: (Number(cur.hours) || 0) > 0 ? Number(cur.hours) : (Number(doc.data.targets && doc.data.targets.shiftHours) || 8), fromHourly: true, roadie: !!doc.data.roadie, updatedAt: nowIso });
      if (!body.campaign && rep.campaign) body.campaign = rep.campaign;
      out.push({ path: ep, data: body });
      out.push({ path, data: Object.assign({}, doc.data, { syncedAt: nowIso }) });
    });
    if (out.length) putDocs_(t, out);
    return out.length / 2;
  } finally { lock.releaseLock(); }
}
function setOwnerPin() {
  const ui = SpreadsheetApp.getUi();
  const r = ui.prompt(APP_TITLE, 'Choose a PIN (4 to 8 digits) for ' + OWNER_NAME + ':', ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() !== ui.Button.OK) return;
  const pin = r.getResponseText().trim();
  if (!/^\d{4,8}$/.test(pin)) { ui.alert('Please use 4 to 8 digits.'); return; }
  ensureSheets_();
  const acc = readAccess_();
  const row = acc.find(a => a.uid === OWNER_UID);
  if (row) { row.pinHash = hash_(pin); writeAccessRow_(row); }
  else appendAccess_({ uid: OWNER_UID, name: OWNER_NAME, pinHash: hash_(pin), role: 'owner', tokens: [], created: new Date().toISOString() });
  const t = loadTable_();
  if (!t.map['reps/' + OWNER_UID] || t.map['reps/' + OWNER_UID].del) {
    putDocs_(t, [{ path: 'reps/' + OWNER_UID, data: { name: OWNER_NAME, team: '', active: true, order: 0, joinedAt: new Date().toISOString() } }]);
  }
  if (!t.map['settings/main']) putDocs_(t, [{ path: 'settings/main', data: { teamName: 'Kaizen Enterprises' } }]);
  ui.alert('Done. Log in to the app as "' + OWNER_NAME + '" with that PIN.');
}

/* ---------- sheets ---------- */
function ss_() { return SpreadsheetApp.getActiveSpreadsheet(); }
function ensureSheets_() {
  const s = ss_();
  let d = s.getSheetByName(DATA);
  if (!d) { d = s.insertSheet(DATA); d.getRange(1, 1, 1, 4).setValues([['path', 'updated', 'deleted', 'json']]); d.setFrozenRows(1); }
  let a = s.getSheetByName(ACCESS);
  if (!a) { a = s.insertSheet(ACCESS); a.getRange(1, 1, 1, 6).setValues([['uid', 'name', 'pinHash', 'role', 'tokens', 'created']]); a.setFrozenRows(1); a.hideSheet(); }
  const p = PropertiesService.getScriptProperties();
  if (!p.getProperty('SALT')) p.setProperty('SALT', Utilities.getUuid());
}
function hash_(pin) {
  const salt = PropertiesService.getScriptProperties().getProperty('SALT') || '';
  return Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, salt + ':' + pin));
}

/* Access sheet: one row per login */
function readAccess_() {
  const sh = ss_().getSheetByName(ACCESS); if (!sh) return [];
  const v = sh.getDataRange().getValues(); const out = [];
  for (let i = 1; i < v.length; i++) {
    if (!v[i][0]) continue;
    let tokens = []; try { tokens = JSON.parse(v[i][4] || '[]'); } catch (e) {}
    out.push({ row: i + 1, uid: String(v[i][0]), name: String(v[i][1]), pinHash: String(v[i][2]), role: String(v[i][3] || 'rep'), tokens, created: v[i][5] });
  }
  return out;
}
function writeAccessRow_(a) {
  ss_().getSheetByName(ACCESS).getRange(a.row, 1, 1, 6).setValues([[a.uid, a.name, a.pinHash, a.role, JSON.stringify(a.tokens || []), a.created || '']]);
}
function appendAccess_(a) {
  ss_().getSheetByName(ACCESS).appendRow([a.uid, a.name, a.pinHash, a.role, JSON.stringify(a.tokens || []), a.created || '']);
}

/* Data sheet: one row per document, JSON split across cells */
function loadTable_() {
  const sh = ss_().getSheetByName(DATA);
  const last = sh.getLastRow(); const width = Math.max(4, sh.getLastColumn());
  const v = last > 1 ? sh.getRange(2, 1, last - 1, width).getValues() : [];
  const map = {};
  v.forEach((r, i) => {
    if (!r[0]) return;
    let json = ''; for (let c = 3; c < r.length; c++) { if (r[c] === '' || r[c] == null) break; json += r[c]; }
    let data = null; if (json) { try { data = JSON.parse(json); } catch (e) { data = null; } }
    map[String(r[0])] = { row: i + 2, ts: Number(r[1]) || 0, del: r[2] === true || r[2] === 'TRUE', data };
  });
  return { sh, map };
}
function putDocs_(t, docs) { // docs: [{path, data, del}]
  const now = Date.now();
  docs.forEach((d, k) => {
    const ts = now + k;
    const json = d.del ? '' : JSON.stringify(d.data || {});
    if (json.length > CHUNK * MAX_CHUNKS) throw new Error('TOO_LARGE');
    const cells = []; for (let i = 0; i < json.length; i += CHUNK) cells.push(json.slice(i, i + CHUNK));
    const cur = t.map[d.path];
    const width = Math.max(cells.length, cur && cur.width ? cur.width : 1);
    const row = [d.path, ts, !!d.del].concat(cells);
    if (cur) {
      // clear old overflow cells, then write
      const lastCol = t.sh.getLastColumn();
      if (lastCol > 3) t.sh.getRange(cur.row, 4, 1, lastCol - 3).clearContent();
      t.sh.getRange(cur.row, 1, 1, row.length).setValues([row]);
      t.map[d.path] = { row: cur.row, ts, del: !!d.del, data: d.del ? null : d.data };
    } else {
      t.sh.appendRow(row);
      t.map[d.path] = { row: t.sh.getLastRow(), ts, del: !!d.del, data: d.del ? null : d.data };
    }
  });
  return now + docs.length;
}

/* ---------- sessions ---------- */
function newToken_() { return Utilities.getUuid().replace(/-/g, '') + Utilities.getUuid().replace(/-/g, '').slice(0, 8); }
function auth_(token) {
  if (!token) throw new Error('AUTH');
  const acc = readAccess_();
  const a = acc.find(x => (x.tokens || []).indexOf(token) >= 0);
  if (!a) throw new Error('AUTH');
  return a;
}
function tooMany_(name) {
  const c = CacheService.getScriptCache(); const k = 'fail:' + name.toLowerCase();
  return Number(c.get(k) || 0) >= 6;
}
function noteFail_(name) {
  const c = CacheService.getScriptCache(); const k = 'fail:' + name.toLowerCase();
  c.put(k, String(Number(c.get(k) || 0) + 1), 600);
}
function issue_(a) {
  const tok = newToken_();
  a.tokens = (a.tokens || []).concat([tok]).slice(-20);
  writeAccessRow_(a);
  return { token: tok, uid: a.uid, role: a.role, name: a.name };
}

/* ---------- public API (called from the page) ---------- */
function publicInfo() {
  ensureSheets_();
  const t = loadTable_();
  const main = (t.map['settings/main'] || {}).data || {};
  const logo = (t.map['settings/logo'] || {}).data || {};
  const teams = Object.keys(t.map).filter(p => /^teams\/[^/]+$/.test(p) && !t.map[p].del)
    .map(p => ({ id: p.split('/')[1], name: (t.map[p].data || {}).name || 'Team' }));
  return { name: main.teamName || 'Kaizen Enterprises', logo: logo.data || null, teams, hasOwner: readAccess_().some(a => a.role === 'owner') };
}
function login(name, pin) {
  name = String(name || '').trim(); pin = String(pin || '').trim();
  if (!name || !pin) throw new Error('Enter your name and PIN.');
  if (tooMany_(name)) throw new Error('Too many tries. Wait 10 minutes and try again.');
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const a = readAccess_().find(x => x.name.toLowerCase() === name.toLowerCase());
    if (!a || a.pinHash !== hash_(pin)) { noteFail_(name); throw new Error('That name and PIN don’t match.'); }
    return issue_(a);
  } finally { lock.releaseLock(); }
}
function register(name, pin, team) {
  name = String(name || '').trim().replace(/\s+/g, ' '); pin = String(pin || '').trim(); team = ''; // only the owner assigns teams
  if (name.length < 2 || name.length > 40) throw new Error('Enter your first and last name.');
  if (!/^\d{4,8}$/.test(pin)) throw new Error('Your PIN needs 4 to 8 digits.');
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    ensureSheets_();
    const acc = readAccess_();
    if (!acc.some(a => a.role === 'owner')) throw new Error('The app isn’t set up yet. Ask Ameliah to finish setup.');
    if (acc.some(a => a.name.toLowerCase() === name.toLowerCase())) throw new Error('Someone with that name already has a login. Add your last name or initial.');
    const t = loadTable_();
    if (team && (!t.map['teams/' + team] || t.map['teams/' + team].del)) team = '';
    const uid = 'u_' + Utilities.getUuid().replace(/-/g, '').slice(0, 20);
    const a = { uid, name, pinHash: hash_(pin), role: 'rep', tokens: [], created: new Date().toISOString() };
    appendAccess_(a);
    putDocs_(t, [{ path: 'reps/' + uid, data: { name, team, active: true, order: Date.now(), joinedAt: new Date().toISOString() } }]);
    const fresh = readAccess_().find(x => x.uid === uid);
    return issue_(fresh);
  } finally { lock.releaseLock(); }
}
function logout(token) {
  try { const a = auth_(token); a.tokens = a.tokens.filter(x => x !== token); writeAccessRow_(a); } catch (e) {}
  return true;
}

/* permissions mirror the app's rules, enforced here */
function ctx_(a, t) {
  let role = a.role;
  // anyone who leads a team (Sales Leader and up who created one) automatically gets leader access to its members
  if (role === 'rep' && t.map['teams/' + a.uid] && !t.map['teams/' + a.uid].del) role = 'tl';
  const c = { uid: a.uid, role: role, members: {} };
  if (role === 'tl') Object.keys(t.map).forEach(p => {
    const m = /^reps\/([^/]+)$/.exec(p); if (m && !t.map[p].del && (t.map[p].data || {}).team === a.uid) c.members[m[1]] = true;
  });
  return c;
}
function canRead_(c, path) {
  if (c.role === 'owner') return true;
  const s = path.split('/'), top = s[0], id = s[1];
  if (top === 'settings' || top === 'summary' || top === 'teams' || top === 'board') return true;
  if (top === 'reps' || top === 'conf') return id === c.uid || (c.role === 'tl' && !!c.members[id]);
  return false;
}
function canWrite_(c, path) {
  if (c.role === 'owner') return true;
  const s = path.split('/'), top = s[0], id = s[1];
  if (top === 'teams') return s.length === 2 && id === c.uid;
  if (top === 'board') return s.length === 2 && (id === c.uid || (c.role === 'tl' && !!c.members[id]));
  if (top === 'reps') return id === c.uid || (c.role === 'tl' && !!c.members[id]);
  if (top === 'conf') return c.role === 'tl' && !!c.members[id];
  return false;
}
function sync(token, since) {
  const a = auth_(token); const now = Date.now();
  const t = loadTable_(); const c = ctx_(a, t); since = Number(since) || 0;
  const docs = [];
  Object.keys(t.map).forEach(p => {
    const r = t.map[p]; if (r.ts <= since) return; if (!canRead_(c, p)) return;
    docs.push(r.del ? { p, del: 1 } : { p, d: r.data });
  });
  return { now, docs, me: { uid: a.uid, role: c.role, name: a.name }, full: since === 0 };
}
function merge_(a, b) {
  const o = Object.assign({}, a || {});
  Object.keys(b || {}).forEach(k => {
    const v = b[k];
    if (v && typeof v === 'object' && !Array.isArray(v) && o[k] && typeof o[k] === 'object' && !Array.isArray(o[k])) o[k] = merge_(o[k], v);
    else o[k] = v;
  });
  return o;
}
function write(token, ops) {
  const a = auth_(token);
  const lock = LockService.getScriptLock(); lock.waitLock(25000);
  try {
    const t = loadTable_(); const c = ctx_(a, t); const out = [];
    (ops || []).forEach(op => {
      const path = String(op.path || '');
      if (!/^[A-Za-z0-9_\-.~:@+]+(\/[A-Za-z0-9_\-.~:@+]+)*$/.test(path) || path.split('/').length % 2) throw new Error('BAD_PATH');
      if (!canWrite_(c, path)) throw new Error('DENIED');
      const cur = t.map[path];
      const seg = path.split('/');
      if (c.role !== 'owner' && seg.length === 2 && seg[0] === 'reps' && op.op !== 'delete' && op.data && 'team' in op.data) {
        const was = cur && !cur.del ? (cur.data.team || '') : '';
        const now = op.data.team || '';
        if (now !== was && !(seg[1] === c.uid && now === c.uid)) throw new Error('Only Ameliah can change which team someone is on.');
      }
      if (c.role !== 'owner' && seg.length === 2 && seg[0] === 'reps' && op.data && op.data.promo && ('stage' in op.data.promo || 'promotedAt' in op.data.promo)) {
        const was = cur && !cur.del ? ((cur.data.promo || {}).stage || 'sa') : 'sa';
        if ((op.data.promo.stage || was) !== was || 'promotedAt' in op.data.promo) {
          const mine = ((t.map['reps/' + c.uid] || {}).data || {}).promo || {};
          if (['tm', 'ad', 'dir', 'md'].indexOf(mine.stage) < 0 || seg[1] === c.uid) throw new Error('Only admins and Team Managers can approve promotions.');
        }
      }
      if (op.op === 'delete') out.push({ path, del: true });
      else if (op.op === 'update') {
        if (!cur || cur.del) throw new Error('MISSING');
        out.push({ path, data: merge_(cur.data, op.data || {}) });
      } else out.push({ path, data: op.data || {} });
    });
    putDocs_(t, out);
    return { ok: true, docs: out.map(d => d.del ? { p: d.path, del: 1 } : { p: d.path, d: d.data }) };
  } finally { lock.releaseLock(); }
}

/* owner-only login management */
function adminAccess(token, action, uid, value) {
  const me = auth_(token);
  if (me.role !== 'owner') throw new Error('DENIED');
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const acc = readAccess_();
    if (action === 'list') return acc.map(x => ({ uid: x.uid, name: x.name, role: x.role }));
    const a = acc.find(x => x.uid === uid); if (!a) throw new Error('Not found');
    if (a.role === 'owner') throw new Error('The owner login can only be changed from the spreadsheet menu.');
    if (action === 'role') { if (['rep', 'tl'].indexOf(value) < 0) throw new Error('Bad role'); a.role = value; a.tokens = a.tokens; writeAccessRow_(a); }
    else if (action === 'pin') { if (!/^\d{4,8}$/.test(String(value))) throw new Error('PIN needs 4 to 8 digits'); a.pinHash = hash_(String(value)); a.tokens = []; writeAccessRow_(a); }
    else if (action === 'remove') { a.tokens = []; a.pinHash = 'removed'; a.name = a.name + ' (removed ' + new Date().toISOString().slice(0, 10) + ')'; writeAccessRow_(a); }
    else if (action === 'rename') { const n = String(value || '').trim(); if (!n) throw new Error('Name needed'); if (acc.some(x => x.uid !== uid && x.name.toLowerCase() === n.toLowerCase())) throw new Error('That name is taken'); a.name = n; writeAccessRow_(a); }
    return { ok: true };
  } finally { lock.releaseLock(); }
}

/* ---------- weekly scorecards: emailed to the owner every Saturday night ----------
   Built only from each person's data and Ameliah's coaching material:
   - Knowing your numbers: priority Stops -> Stop to buyer % -> Sale average -> Upsells, with Low / Avg / Strong bands
   - Feedback rules: 3 good things, 1 focus area, finish on a positive
   - Their own self-review: 3 good things, 2 learnings, 1 focus (daily reflections) + weekly review
   Each person's card is attached as its own printable PDF. */
const OREXIS_LOGO = 'data:image/jpeg;base64,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';
const OREXIS_LOGO_LIGHT = 'data:image/jpeg;base64,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';
/* Orexis-branded versions of the two apps: same data, same logins, Orexis look */
function orexis_(hourly) {
  let s = HtmlService.createHtmlOutputFromFile(hourly ? 'Hourly' : 'Index').getContent();
  const rep = (a, b) => { s = s.split(a).join(b); };
  rep('#7b1e2e', '#424965'); rep('#f4e6e8', '#fbe7cf'); rep('#6e1423', '#2E3A59');
  rep('<title>Kaizen KPIs</title>', '<title>Orexis KPIs</title>'); rep('<title>Kaizen Hourly</title>', '<title>Orexis Hourly</title>');
  rep("if(info.logo){", "if(window.OREXIS_LOGO)info.logo=window.OREXIS_LOGO;if(info.logo){");
  rep("info.name||'Kaizen KPIs'", "'Orexis Innovations'");
  rep("const src=S.logo&&S.logo.data;", "const src=window.OREXIS_LOGO||(S.logo&&S.logo.data);");
  rep("(name||'Kaizen KPIs')", "'Orexis Innovations'");
  rep("if(l.exists&&l.data().data)S.logo=l.data().data}catch(_){}", "if(l.exists&&l.data().data)S.logo=l.data().data}catch(_){}S.logo=window.OREXIS_LOGO||S.logo;");
  rep("h('h1',{text:'Hourly KPIs'})", "h('h1',{text:'Orexis Hourly'})");
  rep('https://ameliahcrowe-cloud.github.io/kaizen-kpis/', 'https://ameliahcrowe-cloud.github.io/kaizen-kpis/orexis/');
  rep("'All of '+((S.settings.teamName||'').trim()||'Kaizen')", "'All of Orexis'");
  rep('Synced to Kaizen KPIs', 'Synced to Orexis KPIs'); rep('Not synced to Kaizen KPIs', 'Not synced to Orexis KPIs'); rep('open Kaizen KPIs for', 'open Orexis KPIs for'); rep('numbers are in Kaizen KPIs', 'numbers are in Orexis KPIs'); rep('numbers to Kaizen KPIs', 'numbers to Orexis KPIs');
  const css = ':root{--bg:#F6F3EE!important;--line:#e3dfd7!important;--muted:#5b6275!important;--ink:#1E2740!important;--accent:#2E3A59!important;--accent-soft:#fbe7cf!important}' +
    'body{background:#F6F3EE!important}header.top,.top{background:#2E3A59!important;color:#F5B56E!important;border-radius:14px;padding:12px 14px!important;margin-bottom:10px}' +
    'header.top h1,.top h1,#teamTitle{color:#F5B56E!important;letter-spacing:.12em}#teamSub,.top .sub{color:#c9cfe0!important}' +
    'header.top button,.top button,header.top input,header.top select{background:#414A63!important;color:#fff!important;border-color:#5a6688!important}' +
    'nav.tabs button[aria-selected=true]{border-bottom-color:#F5B56E!important}#mnav button[aria-selected=true]{color:#2E3A59!important;background:#fbe7cf!important}' +
    'h2,h3{color:#2E3A59}.card{border-color:#e3dfd7}';
  s = s.replace('</style>', css + '</style>');
  s = s.replace('<head>', '<head><script>window.OREXIS=1;window.OREXIS_LOGO=' + JSON.stringify(OREXIS_LOGO) + '</script>');
  if (s.indexOf('window.OREXIS=1') < 0) s = '<script>window.OREXIS=1;window.OREXIS_LOGO=' + JSON.stringify(OREXIS_LOGO) + '</script>' + s;
  return HtmlService.createHtmlOutput(s).setTitle(hourly ? 'Orexis Hourly' : 'Orexis KPIs')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}
// Test run: only these people's cards, emailed straight to these addresses
function scorecardTest() {
  const names = ['Ameliah', 'Ben Cotter', 'Freddie Juffkins'], to = 'ameliah.crowe@gmail.com,ben@orexisinnovations.com';
  const today = Utilities.formatDate(new Date(), 'Australia/Melbourne', 'yyyy-MM-dd'), wk = (new Date(today + 'T12:00:00Z').getUTCDay() % 6 === 0) ? mondayOf_(today) : addDaysS_(mondayOf_(today), -7), tag = wk.slice(8, 10) + '/' + wk.slice(5, 7);
  const cards = scorecards_(wk).filter(c => names.indexOf(c.name.replace(/ \(test\)$/, '')) >= 0);
  const pdfs = cards.map(c => Utilities.newBlob(page_(c.html), 'text/html', 'x.html').getAs('application/pdf').setName('Scorecard WC ' + tag.replace('/', '-') + ' - ' + c.name.replace(/ \(test\)$/, '') + '.pdf'));
  MailApp.sendEmail({ to: to, subject: 'Orexis weekly scorecards · TEST · WC ' + tag, name: 'Orexis Innovations',
    htmlBody: '<div style="font-family:Helvetica,Arial,sans-serif;font-size:14px">Test run of the new weekly scorecards: ' + cards.map(c => c.name.replace(/ \(test\)$/, '')).join(', ') + '. Each one is attached as a printable PDF.</div>', attachments: pdfs });
  return 'Sent ' + cards.length + ' to ' + to;
}
const SCORE_TO = 'ameliah.crowe@gmail.com';
const SCORE_INCLUDE_TEST = true; // trial: include test people in scorecards
const BANDS = {
  stops: { low: 15, strong: 20 },   // per shift: <15 low, 15-20 avg, 20+ strong
  conv: { low: 65, strong: 75 },    // stop to buyer %: <65 low, 65-75 avg, 75+ strong
  avg: { low: 50, strong: 65 },     // sale average $: <50 low, 50-65 avg, 65+ strong
  ups: { low: 15, strong: 30 }      // buyers to upsells %: <15 low, 15-30 avg, 30%+ strong (KPI doc: 30% upsells)
};
// One mechanical piece of feedback per focus (feedback rules), from Ameliah's training docs
const FOCUS_LIB = {
  stops: { area: 'Stopping', topic: 'Stopping · Staying in the zone', why: 'Low stops = not staying in the zone. Don’t miss anyone walking past, and finish your lines on everyone you try to stop. Stop more, win more: every stop is worth about $50.',
    tip: 'Distances 6m / 3m / 1.5m: “Hey there” → “Just a one off” → “for the Paralympics”. Eye contact from hello until they stop. Less words, more silence. High tone on hello, low tone on the call to action.' },
  conv: { area: 'Stop to buyer', topic: 'Cracking the egg · Handling concerns', why: 'Stop to buyer % is dictated by your ability to deal with concerns. If it’s low, the engagement or the intro rapport isn’t there yet.',
    tip: 'Crack the egg first: giggle = gottem. Then nail “How much is it?”: “Thanks for asking, most people run away! You basically get to pick.” How much is a green light, they want to support.' },
  avg: { area: 'Sale average', topic: 'Landing impulses · Building value', why: 'A low $ average means not enough value built in the cause, or not enough confidence in the close.',
    tip: 'Land “Paralympians are self-funded” and pause. Pitch the medals low and slow, then close assumptively on the medal pack, not the minimum.' },
  ups: { area: 'Upselling', topic: 'Upselling · AARR', why: 'Every upsell is a second sale. Price is only an issue if the value isn’t there.',
    tip: '“Can I be a bit cheeky?” Then AARR: Appreciate, Anchor high (2 up, 1 down), Re-educate on what the money does, Re-close low and slow. Ask everyone.' },
  mind: { area: 'Mentality', topic: 'Bulletproof attitude · Circle of influence', why: 'The pitch is there but the results aren’t yet, so it’s a mentality piece. Let your attitude drive your sales, not your sales drive your attitude.',
    tip: 'Control the controllables: attitude, energy, work ethic, KPIs. “Ok, next one” after every interaction. Ask “Why am I here today?” 3–4 times a day.' },
  cons: { area: 'Consistency', topic: 'Creating consistency · Law of averages', why: 'Big swings day to day. Standards are the bridge between your vision and your results. The gap between your best and worst should be small.',
    tip: 'Set a daily benchmark for stops, buyers and retail and tick it off every day. Review your KPIs every hour, then reset.' }
};
// What each team needs: lowest band first, in the same priority order
const PEOPLE_NEED = { stops: 'Stopping', conv: 'Concerns', avg: 'Impulses', ups: 'Upselling', mind: 'Belief', cons: 'Consistency' };
function teamPiece_(size, leaders, newStarts, core) {
  // From: 5M model, Effective leadership, Focus training (Tarqs), Creating buy in (Shayne Way), Waaka, Tim Cimino, Ameliah's leader notes
  if (!size) return { t: 'Recruitment · Creating buy in', d: 'Get your first new start. People come in for you: tell your story, create stability, show the data.', book: 'The Ideal Team Player' };
  if (newStarts && newStarts * 2 >= size) return { t: 'Win early', d: newStarts + ' of ' + size + ' are new starts. Model then mentor: first hour just watch. Celebrate everything (first stop, first reject, first sale). Video pitch on day 5, 10 and 20. The goal is to get them back tomorrow.', book: 'Leaders Eat Last' };
  if (!leaders && size < 4) return { t: 'Build a core player', d: 'Small team, no leaders yet. Hang, train, develop: who am I hanging with, training and developing this week? Build the person and the sales will come.', book: 'The Ideal Team Player' };
  if (leaders < 2) return { t: 'Develop your next leader', d: size + ' people, ' + leaders + ' leader' + (leaders === 1 ? '' : 's') + '. 5M: Model, Mentor, Monitor (just watch), Motivate, Multiply. Spend your time on the 2–3 with goals. Show your goal stack early.', book: 'Multipliers' };
  if (size >= 8) return { t: 'Team identity', d: size + ' people, ' + leaders + ' leaders. Team nights, weekly chats, team numbers for promotions. Train in 2/4/6 week blocks, not week to week.', book: 'The Five Dysfunctions of a Team' };
  return { t: 'Touch points', d: size + ' people, ' + leaders + ' leaders' + (core ? ', ' + core + ' core' : '') + '. Split your time: key people, rising stars, bottom 10%. More one on one time, more team nights. People don’t quit businesses, they quit people.', book: 'Multipliers' };
}
// Ameliah's leader tips, one per card
const LEADER_TIPS = ['Make everyone you work with feel better.', 'Being kind beats being right.', 'Avoid opinions. Allow people to be wrong.', 'The right thing said at the wrong time is wrong.', 'Give a compliment before you give feedback.', 'More team nights, more time with just team members.', 'Show your team the elevator pitch of your goal stack early on.', 'Ask first: “Would you like some feedback?”'];
function setupScorecards() {
  ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === 'weeklyScorecards').forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('weeklyScorecards').timeBased().onWeekDay(ScriptApp.WeekDay.SATURDAY).atHour(20).inTimezone('Australia/Melbourne').create();
  return 'Scorecards will email every Saturday about 8pm';
}
function addDaysS_(d, n) { const x = new Date(d + 'T00:00:00Z'); x.setUTCDate(x.getUTCDate() + n); return x.toISOString().slice(0, 10); }
function weeklyScorecards(weekArg) {
  const today = Utilities.formatDate(new Date(), 'Australia/Melbourne', 'yyyy-MM-dd');
  const wk = typeof weekArg === 'string' ? weekArg : mondayOf_(today);
  const to = SCORE_TO || Session.getEffectiveUser().getEmail(), tag = wk.slice(8, 10) + '/' + wk.slice(5, 7);
  const cards = scorecards_(wk);
  const pdfs = cards.map(c => Utilities.newBlob(page_(c.html), 'text/html', 'x.html').getAs('application/pdf').setName('Scorecard WC ' + tag.replace('/', '-') + ' - ' + c.name + '.pdf'));
  cards.forEach((c, i) => GmailApp.createDraft(to, 'Your scorecard · WC ' + tag + ' · ' + c.name, 'Your scorecard is attached.', { htmlBody: page_(c.html), attachments: [pdfs[i]], name: 'Kaizen' }));
  MailApp.sendEmail({ to: to, subject: 'Weekly scorecards · WC ' + tag + ' (' + cards.length + ')', name: 'Kaizen scorecards',
    htmlBody: '<div style="font-family:Helvetica,Arial,sans-serif;font-size:14px">' + cards.length + ' scorecards attached as PDFs, ready to print. Each one is also in your Gmail <b>Drafts</b> to add notes and send on.</div>',
    attachments: pdfs });
  return 'Sent ' + cards.length + ' scorecards to ' + to;
}
function page_(inner) { return '<html><head><meta charset="utf-8"><style>@page{size:A4;margin:12mm}body{margin:0;font-family:Helvetica,Arial,sans-serif;color:#1d1416}</style></head><body>' + inner + '</body></html>'; }
function scorecards_(wk) {
  const t = loadTable_(), M = t.map, end = addDaysS_(wk, 6);
  const get = p => (M[p] && !M[p].del) ? M[p].data : null;
  const logo = ((get('settings/logo') || {}).data) || '';
  const coName = ((get('settings/main') || {}).teamName) || 'Kaizen Enterprises';
  const reps = Object.keys(M).filter(p => /^reps\/[^/]+$/.test(p) && !M[p].del && M[p].data && M[p].data.active !== false && (SCORE_INCLUDE_TEST || !M[p].data.test)).map(p => ({ id: p.split('/')[1], r: M[p].data }));
  const allReps = Object.keys(M).filter(p => /^reps\/[^/]+$/.test(p) && !M[p].del && M[p].data && M[p].data.active !== false).map(p => ({ id: p.split('/')[1], r: M[p].data }));
  const parentOf = x => { const tm = x.r.team || ''; if (!tm) return x.id === 'u_owner' ? null : 'u_owner'; if (tm === x.id) return null; return tm; };
  const downOf = (uid, seen) => { seen = seen || { [uid]: 1 }; let o = []; allReps.filter(x => !seen[x.id] && parentOf(x) === uid && (SCORE_INCLUDE_TEST || !x.r.test)).forEach(x => { seen[x.id] = 1; o.push(x); o = o.concat(downOf(x.id, seen)); }); return o; };
  const STG = ['sa', 'sc', 'sl', 'tl', 'tm', 'ad', 'dir', 'md'];
  const dayOf = (id, d) => get('conf/' + id + '/days/' + d) || get('reps/' + id + '/entries/' + d);
  const sum = (id, a, b) => { const s = { stops: 0, buyers: 0, retail: 0, upsells: 0, upsellRetail: 0, medals: 0, hours: 0, shifts: 0, days: [] };
    for (let d = a; d <= b; d = addDaysS_(d, 1)) { const e = dayOf(id, d); const has = e && (Number(e.stops) || Number(e.retail) || Number(e.hours));
      s.days.push({ d, retail: e ? Number(e.retail) || 0 : 0, worked: !!has, roadie: !!(e && e.roadie), hours: e ? Number(e.hours) || 0 : 0 });
      if (has) { s.shifts++; ['stops', 'buyers', 'retail', 'upsells', 'upsellRetail', 'medals', 'hours'].forEach(k => s[k] += Number(e[k]) || 0); } } return s; };
  const kp = s => { const n = Math.max(1, s.shifts); return { stops: s.stops / n, conv: s.stops ? s.buyers / s.stops * 100 : 0, avg: s.buyers ? s.retail / s.buyers : 0, ups: s.buyers ? s.upsells / s.buyers * 100 : 0, rph: s.hours ? s.retail / s.hours : 0, buyers: s.buyers / n }; };
  const band = (k, v) => v >= BANDS[k].strong ? 'STRONG' : v >= BANDS[k].low ? 'AVG' : 'LOW';
  const BC = { STRONG: '#1f7a4d', AVG: '#b86e00', LOW: '#c2410c' };
  const money = n => '$' + Math.round(n).toLocaleString('en-AU');
  const DN = ['Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat', 'Sun'];
  const ROLE = { sa: 'Sales Associate', sc: 'Sales Consultant', sl: 'Sales Leader', tl: 'Team Leader', tm: 'Team Manager', ad: 'Assistant Director', dir: 'Director', md: 'Managing Director' };
  const zone = v => v >= 150 ? 'Diamond' : v >= 125 ? 'Platinum' : v >= 100 ? 'Gold' : v >= 80 ? 'Silver' : 'Red'; // RPH zones from the KPI topic
  const esc = x => String(x == null ? '' : x).replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
  const M1 = '#7b1e2e', MUT = '#6b5b60', LINE = '#e6dcde';
  const out = [];
  reps.map(x => ({ ...x, s: sum(x.id, wk, end) })).sort((a, b) => b.s.retail - a.s.retail).forEach(({ id, r, s }) => {
    const P = sum(id, addDaysS_(wk, -28), addDaysS_(wk, -1)); const K = kp(s), U = P.shifts ? kp(P) : null;
    const rows = [['Stops per shift', 'stops', K.stops, U && U.stops, x => x.toFixed(1), s.hours ? s.stops / s.hours * 8 : K.stops], ['Stop to buyer %', 'conv', K.conv, U && U.conv, x => Math.round(x) + '%'], ['Sale average', 'avg', K.avg, U && U.avg, money], ['Buyers to upsells %', 'ups', K.ups, U && U.ups, x => Math.round(x) + '%']];
    // their own words
    const refl = []; for (let d = wk; d <= end; d = addDaysS_(d, 1)) { const x = get('reps/' + id + '/reflections/' + d); if (x) refl.push(x); }
    const own = { good: [], learn: [], focus: '' };
    refl.forEach(x => { (x.well || []).forEach(w => { if (String(w || '').trim()) own.good.push(w) }); (x.learn || []).forEach(w => { if (String(w || '').trim()) own.learn.push(w) }); if (String(x.focus || '').trim()) own.focus = x.focus; });
    const rev = get('reps/' + id + '/reviews/' + addDaysS_(wk, 7)) || get('reps/' + id + '/reviews/' + wk);
    // 3 good things, all from the data
    const good = [];
    if (s.shifts) {
      rows.filter(x => band(x[1], x[5] != null ? x[5] : x[2]) === 'STRONG').forEach(x => good.push(x[0] + ' ' + x[4](x[2]) + ', in the strong range'));
      rows.forEach(x => { if (U && x[3] && x[2] > x[3] * 1.1 && band(x[1], x[2]) !== 'STRONG') good.push(x[0] + ' up from ' + x[4](x[3]) + ' to ' + x[4](x[2])); });
      if (own.good.length && good.length < 3) good.push('Their own win: ' + own.good[0]);
      const bells = s.days.filter(d => d.worked && d.retail >= (d.roadie ? 1000 : 700)).length; if (bells) good.push(bells + ' bell day' + (bells > 1 ? 's' : '') + ' 🔔');
      const pb = r.pb || {}; ['retail', 'stops', 'buyers', 'upsells', 'medals'].forEach(k => { const x = (pb.day || {})[k]; if (x && x.date >= wk && x.date <= end) good.push('New PB: ' + (k === 'retail' ? money(x.v) + ' retail' : x.v + ' ' + k) + ' in a day (' + DN[(new Date(x.date + 'T12:00:00Z').getUTCDay() + 6) % 7] + ')'); const w = (pb.week || {})[k]; if (w && w.week === wk) good.push('New best week for ' + k + ': ' + (k === 'retail' ? money(w.v) : w.v)); });
      const best = s.days.filter(d => d.worked).sort((a, b) => b.retail - a.retail)[0]; if (best && best.retail) good.push('Best day: ' + DN[s.days.indexOf(best)] + ' ' + money(best.retail));
      if (s.medals) good.push(s.medals + ' medal' + (s.medals > 1 ? 's' : '') + ' / premium pack' + (s.medals > 1 ? 's' : ''));
    }
    const three = good.slice(0, 3);
    // 1 focus area: Planning education -> trial of elimination, then check mentality
    const hrsPer = s.shifts ? s.hours / s.shifts : 8; const stops8 = s.hours ? s.stops / s.hours * 8 : K.stops; // stops scaled to an 8hr day
    const bs = { stops: band('stops', stops8), conv: band('conv', K.conv), avg: band('avg', K.avg), ups: band('ups', K.ups) };
    let fk = null;
    if (s.shifts) {
      fk = ['stops', 'conv', 'avg', 'ups'].find(k => bs[k] === 'LOW') || null;
      const wd = s.days.filter(d => d.worked && d.hours).map(d => d.retail / d.hours); const swing = wd.length >= 3 && Math.max(...wd) > 3 * Math.max(1, Math.min(...wd));
      if (!fk && ((rev && Number(rev.motivation) && Number(rev.motivation) <= 5) || (K.rph && K.rph < 100))) fk = 'mind';
      if (!fk && swing) fk = 'cons';
      if (!fk) fk = ['stops', 'conv', 'avg', 'ups'].find(k => bs[k] === 'AVG') || null;
    }
    const ft = fk ? FOCUS_LIB[fk] : null;
    const days = s.days.filter(d => d.worked).map(d => d.retail).sort((a, b) => a - b); const bench = days.length ? Math.max(BANDS.avg.low * 5, Math.round(days[Math.floor(days.length / 2)] / 25) * 25) : 0;
    const fval = fk === 'stops' ? K.stops.toFixed(1) + ' a shift (≈' + Math.round(stops8) + ' on an 8 hr day)' : fk === 'conv' ? Math.round(K.conv) + '%' : fk === 'avg' ? money(K.avg) : fk === 'ups' ? Math.round(K.ups) + '% of buyers' : fk === 'mind' ? (rev && rev.motivation ? 'motivation ' + rev.motivation + '/10' : money(K.rph) + '/hr') : fk === 'cons' ? 'best day ' + money(days[days.length - 1] || 0) + ', lowest ' + money(days[0] || 0) : '';
    const ownF = (own.focus + ' ' + (rev && rev.plan || '')).toLowerCase(); const kw = { stops: /stop/, conv: /concern|objection|convert|close|buyer/, avg: /medal|impulse|price|average|\$av/, ups: /upsell/, mind: /mind|attitude|confiden|energy|belie/, cons: /consisten|benchmark|standard/ };
    const matches = fk && kw[fk].test(ownF);
    const learnt = (r.topics || {}); const topicDone = ft && Object.keys(learnt).some(k => learnt[k] && learnt[k].learnt && ft.topic.toLowerCase().indexOf(k.toLowerCase()) >= 0);
    // next week numbers (their formula: stops x conversion x sale average = retail)
    const hrs = s.shifts ? s.hours / s.shifts : 6;
    const base = Math.max(s.retail, P.shifts ? P.retail / 4 : 0); const target = Math.ceil(base * 1.01 / 25) * 25;
    const stopT = fk === 'stops' ? BANDS.stops.low : Math.max(BANDS.stops.low, Math.ceil(K.stops));
    const role = ROLE[(r.promo || {}).stage] || 'Sales Associate';
    const th = 'text-align:left;font-size:10px;letter-spacing:.06em;color:' + MUT + ';padding:6px 8px;border-bottom:1px solid ' + LINE;
    const td = 'padding:7px 8px;border-bottom:1px solid ' + LINE + ';font-size:13px';
    const sec = (title, body) => '<div style="margin-top:14px"><div style="font-size:11px;font-weight:700;letter-spacing:.08em;color:' + M1 + ';margin-bottom:6px">' + title + '</div>' + body + '</div>';
    const ul = a => '<ul style="margin:0;padding-left:18px;font-size:13px;line-height:1.5">' + a.map(x => '<li>' + esc(x) + '</li>').join('') + '</ul>';
    // ---- simple layout (Atlas-style): header, one line, 5 numbers, weeks, wins/challenges, 3 focuses ----
    const N = '#424965', G = '#F5B56E';
    const sec2 = t => '<div style="margin:16px 0 8px;font-size:11px;font-weight:700;letter-spacing:.14em;color:' + N + '"><span style="display:inline-block;width:18px;height:2px;background:' + G + ';vertical-align:middle;margin-right:8px"></span>' + t + '</div>';
    const wkTag = d => d.slice(8, 10) + '/' + d.slice(5, 7);
    // one-line summary
    const strong = rows.filter(x => s.shifts && band(x[1], x[5] != null ? x[5] : x[2]) === 'STRONG').map(x => x[0].toLowerCase());
    const line = !s.shifts ? 'No shifts logged this week. Let’s lock in regular shifts so you can build a rhythm.'
      : (strong.length ? 'Your ' + strong[0] + ' is in the strong range. ' : 'You made ' + s.buyers + ' sale' + (s.buyers === 1 ? '' : 's') + ' across ' + s.shifts + ' shift' + (s.shifts === 1 ? '' : 's') + '. ') + (ft ? 'Next step: ' + ft.area.toLowerCase() + '.' : 'Keep it going and chase a PB.');
    const tile = (l, v, sub) => '<td style="padding:0 14px 0 0;vertical-align:top"><div style="font-size:10px;letter-spacing:.1em;color:' + MUT + '">' + l + '</div><div style="font-size:24px;font-weight:700;color:' + N + '">' + v + '</div><div style="font-size:11px;color:' + MUT + '">' + sub + '</div></td>';
    let h = '<div style="max-width:720px;margin:0 auto;background:#fff;font-family:Helvetica,Arial,sans-serif">' +
      '<div style="background:' + N + ';color:#fff;padding:14px 20px;border-radius:6px">' +
      '<table style="width:100%;border-collapse:collapse"><tr><td style="width:110px;vertical-align:middle"><img src="' + OREXIS_LOGO + '" style="width:100px;display:block;border:0"></td><td style="font-size:11px;letter-spacing:.2em;color:' + G + ';vertical-align:middle">' + 'OREXIS INNOVATIONS · WEEKLY SCORECARD</td><td style="text-align:right;font-size:11px;color:#dfe2ec;vertical-align:middle">WC ' + wkTag(wk) + '/' + wk.slice(0, 4) + '</td></tr></table>' +
      '<div style="font-size:30px;letter-spacing:.04em;margin-top:8px">' + esc((r.name || 'Rep').toUpperCase()) + '</div>' +
      '<div style="font-size:12px;color:#dfe2ec">' + role + ' · ' + s.shifts + ' shift' + (s.shifts === 1 ? '' : 's') + ' this week</div></div>' +
      '<div style="font-size:14px;line-height:1.5;margin:14px 0">' + esc(line) + '</div>' +
      '<table style="border-collapse:collapse"><tr>' +
      tile('RETAIL', money(s.retail), s.shifts + ' shift' + (s.shifts === 1 ? '' : 's')) +
      tile('RPH', s.hours ? money(K.rph) : '—', s.hours ? zone(K.rph) : '') +
      tile('CONVERSION', s.shifts ? Math.round(K.conv) + '%' : '—', 'Standard ' + BANDS.conv.strong + '%') +
      tile('STOPS / SHIFT', s.shifts ? K.stops.toFixed(1) : '—', 'Standard ' + BANDS.stops.strong) +
      tile('SALE AVERAGE', s.shifts ? money(K.avg) : '—', 'Standard ' + money(BANDS.avg.strong)) + '</tr></table>';
    // daily retail chart (table bars so it survives the PDF converter)
    const mx = Math.max(700, ...s.days.map(d => d.retail));
    h += sec2('THE WEEK') + '<table style="width:100%;border-collapse:collapse;table-layout:fixed"><tr style="vertical-align:bottom">' + s.days.map(d => { const bh = Math.round(d.retail / mx * 90); const bell = d.worked && d.retail >= (d.roadie ? 1000 : 700);
      return '<td style="text-align:center;padding:0 6px;height:110px;vertical-align:bottom"><div style="font-size:10px;color:' + MUT + '">' + (d.worked ? money(d.retail) : '') + (bell ? ' 🔔' : '') + '</div><div style="height:' + Math.max(d.worked ? 3 : 0, bh) + 'px;background:' + (bell ? G : N) + ';border-radius:4px 4px 0 0"></div></td>'; }).join('') + '</tr><tr>' +
      s.days.map((d, i) => '<td style="text-align:center;font-size:10px;color:' + MUT + ';border-top:1px solid ' + LINE + ';padding-top:3px">' + DN[i] + '</td>').join('') + '</tr></table>';
    // tracker: standard vs this week vs their usual (last 4 weeks)
    const TR = [['Stops / shift', K.stops, U && U.stops, BANDS.stops.strong, x => x.toFixed(1)], ['Conversion', K.conv, U && U.conv, BANDS.conv.strong, x => Math.round(x) + '%'], ['Sale average', K.avg, U && U.avg, BANDS.avg.strong, money], ['Upsells', K.ups, U && U.ups, BANDS.ups.strong, x => Math.round(x) + '%'], ['RPH', K.rph, U && U.rph, 100, money]];
    const bar = (v, m, c, lab) => '<tr><td style="width:96px;white-space:nowrap;font-size:10px;color:' + MUT + ';padding:1px 6px 1px 0">' + lab + '</td><td style="padding:1px 0"><div style="height:8px;width:' + Math.max(1, Math.round(Math.min(1, (v || 0) / m) * 100)) + '%;background:' + c + ';border-radius:4px"></div></td></tr>';
    h += sec2('TRACKER · STANDARD vs YOU vs YOUR USUAL') + '<table style="width:100%;border-collapse:collapse">' + TR.map(([l, v, u, st, f]) => { const m = Math.max(st, v || 0, u || 0) * 1.05;
      return '<tr style="vertical-align:middle"><td style="width:120px;padding:6px 0;border-bottom:1px solid ' + LINE + '"><div style="font-size:12px;font-weight:700">' + l + '</div><div style="font-size:11px;color:' + MUT + '">' + (s.shifts ? f(v) : '—') + ' <span style="color:' + (s.shifts && v >= st ? '#1f7a4d' : '#c2410c') + '">' + (s.shifts ? (v >= st ? '✓' : '↓') : '') + '</span></div></td><td style="padding:6px 0;border-bottom:1px solid ' + LINE + '"><table style="width:100%;border-collapse:collapse">' +
        bar(st, m, G, 'Standard ' + f(st)) + bar(s.shifts ? v : 0, m, N, 'You ' + (s.shifts ? f(v) : '—')) + bar(u || 0, m, '#b9adb0', 'Usual ' + (u != null ? f(u) : '—')) + '</table></td></tr>'; }).join('') + '</table>';
    // last 4 weeks incl this one
    const wks = [-21, -14, -7, 0].map(o => { const a = addDaysS_(wk, o), x = sum(id, a, addDaysS_(a, 6)); return { a, x }; });
    h += sec2('YOUR WEEKS') + '<table style="width:100%;border-collapse:separate;border-spacing:6px 0"><tr>' + wks.map(({ a, x }) => '<td style="border:1px solid ' + LINE + ';border-radius:8px;padding:8px 10px;width:25%;vertical-align:top' + (a === wk ? ';border-color:' + N : '') + '"><div style="font-size:10px;letter-spacing:.1em;color:' + MUT + '">WC ' + wkTag(a) + '</div><div style="font-size:18px;font-weight:700">' + (x.shifts ? money(x.retail) : '—') + '</div><div style="font-size:11px;color:' + MUT + '">' + (x.shifts ? x.shifts + ' shift' + (x.shifts === 1 ? '' : 's') + ' · ' + x.buyers + ' supporters' : 'not rostered') + '</div></td>').join('') + '</tr></table>';
    // wins & challenges
    const ch = [];
    if (s.shifts) {
      rows.forEach(x => { const b = band(x[1], x[5] != null ? x[5] : x[2]); if (b === 'LOW') ch.push(x[0] + ' ' + x[4](x[2]) + ', under the ' + (x[1] === 'avg' ? money(BANDS.avg.strong) : x[1] === 'stops' ? BANDS.stops.strong : BANDS[x[1]].strong + '%') + ' standard.'); });
      if (K.rph && K.rph < 100) ch.push(money(K.rph) + ' an hour, under Gold ($100).');
    } else ch.push('No shifts this week.');
    const two = (a, b) => '<table style="width:100%;border-collapse:collapse"><tr style="vertical-align:top"><td style="width:50%;padding-right:12px">' + a + '</td><td>' + b + '</td></tr></table>';
    const li = a => '<ul style="margin:0;padding-left:18px;font-size:13px;line-height:1.6">' + a.map(x => '<li>' + esc(x) + '</li>').join('') + '</ul>';
    h += two(sec2('WINS') + (three.length ? li(three) : '<div style="font-size:13px;color:' + MUT + '">—</div>'), sec2('CHALLENGES') + li(ch.slice(0, 3).length ? ch.slice(0, 3) : ['Nothing under standard. Push for a PB.']));
    // 3 focuses for next week
    const fx = [];
    if (!s.shifts || s.shifts < 3) fx.push(['Lock in regular shifts', 'at least three a week.']);
    if (ft) fx.push([ft.area + ':', ft.tip.split(/(?<=[.!?])\s/)[0]]);
    fx.push([stopT + '+ stops a shift.', 'At your conversion that’s about ' + Math.ceil(stopT * Math.max(K.conv, BANDS.conv.low) / 100) + ' supporters a day.']);
    if (fx.length < 3) fx.push([money(target) + ' retail next week', '(1% up). Upsell every sign-up.']);
    h += sec2('FOCUS FOR NEXT WEEK') + '<ol style="margin:0;padding-left:20px;font-size:13px;line-height:1.6">' + fx.slice(0, 3).map(([a, b]) => '<li><b>' + esc(a) + '</b> ' + esc(b) + '</li>').join('') + '</ol>';
    // team (leaders only), kept to a few lines
    const team = downOf(id);
    if (team.length || STG.indexOf((r.promo || {}).stage) >= 2) {
      const T = { stops: 0, buyers: 0, retail: 0, upsells: 0, hours: 0, shifts: 0 }; team.forEach(m => { const x = sum(m.id, wk, end); ['stops', 'buyers', 'retail', 'upsells', 'hours', 'shifts'].forEach(k => T[k] += x[k]); });
      const TK = kp(T), ld = team.filter(m => STG.indexOf((m.r.promo || {}).stage) >= 2).length, ns = team.filter(m => m.r.joinedAt && String(m.r.joinedAt).slice(0, 10) >= addDaysS_(end, -27)).length;
      const tf = T.shifts ? ['stops', 'conv', 'avg', 'ups'].map(k => [k, k === 'stops' ? TK.stops : TK[k]]).find(([k, v]) => band(k, v) === 'LOW') : null;
      const word = !T.shifts ? 'Activity' : tf ? PEOPLE_NEED[tf[0]] : (T.hours && T.retail / T.hours < 100 ? 'Belief' : 'Recognition');
      const piece = teamPiece_(team.length, ld, ns, 0);
      const tt = (l, v, st, ok) => '<td style="border:1px solid ' + LINE + ';border-radius:8px;padding:8px;text-align:center;width:20%"><div style="font-size:10px;letter-spacing:.08em;color:' + MUT + '">' + l + '</div><div style="font-size:20px;font-weight:700;color:' + N + '">' + v + '</div><div style="font-size:10px;color:' + (T.shifts ? (ok ? '#1f7a4d' : '#c2410c') : MUT) + '">' + (T.shifts ? (ok ? '✓ ' : '↓ ') : '') + 'Standard ' + st + '</div></td>';
      const sph = T.hours ? T.stops / T.hours : 0, rphT = T.hours ? T.retail / T.hours : 0;
      const tn = ((get('teams/' + id) || {}).name) || (id === 'u_owner' || (r.promo || {}).stage === 'tm' ? coName : (r.name || '') + '’s team');
      h += sec2('YOUR TEAM · ' + esc(String(tn).toUpperCase())) + '<div style="font-size:13px;line-height:1.6">' + team.length + ' people · ' + ld + ' leaders · ' + ns + ' new starts · ' + money(T.retail) + ' team retail</div>' +
        '<table style="width:100%;border-collapse:separate;border-spacing:6px 0;margin-top:6px"><tr>' +
        tt('STOPS / HR', T.hours ? sph.toFixed(1) : '—', '3', sph >= 3) + tt('BUYERS', String(T.buyers), T.shifts ? Math.round(T.shifts * 12) + ' (12/shift)' : '12/shift', T.buyers >= T.shifts * 12) +
        tt('CONVERSION', T.stops ? Math.round(TK.conv) + '%' : '—', BANDS.conv.strong + '%', TK.conv >= BANDS.conv.strong) + tt('$ AV', T.buyers ? money(TK.avg) : '—', money(BANDS.avg.strong), TK.avg >= BANDS.avg.strong) +
        tt('RPH', T.hours ? money(rphT) : '—', '$100', rphT >= 100) + '</tr></table>' +
        two('<div style="font-size:10px;letter-spacing:.1em;color:' + MUT + ';margin-top:6px">WHAT THEY NEED</div><div style="font-size:20px;font-weight:700;color:' + N + '">' + word + '</div>', '<div style="font-size:10px;letter-spacing:.1em;color:' + MUT + ';margin-top:6px">NEXT TEAM PIECE</div><div style="font-size:13px"><b>' + piece.t + '.</b> ' + esc(piece.d) + '</div>');
    }
    h += '<div style="margin-top:22px;border-top:1px solid ' + LINE + ';padding-top:8px;font-size:10px;letter-spacing:.06em;color:' + MUT + '">STANDARDS PER SHIFT: ' + BANDS.stops.strong + ' STOPS · ' + BANDS.conv.strong + '% CONVERSION · ' + money(BANDS.avg.strong) + ' AVERAGE SALE · ' + BANDS.ups.strong + '% UPSELLS · $100+ RPH (GOLD)</div></div>';
    out.push({ name: (r.name || 'Rep') + (r.test ? ' (test)' : ''), html: h });
  });
  return out;
}
