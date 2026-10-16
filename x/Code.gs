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

/* ---------- weekly scorecards: emailed to the owner every Saturday night ---------- */
const SCORE_TO = 'ameliah.crowe@gmail.com';
const SCORE_INCLUDE_TEST = true; // trial: include 🧪 test people (e.g. Ben) in scorecards
const STD = { stops: 18, buyers: 12, conv: 70, avg: 65, upsells: 4, rph: 120 }; // company standards per shift
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
  const wrap = h => '<div style="font-family:Helvetica,Arial,sans-serif;background:#f5f1ee;padding:14px"><div style="max-width:560px;margin:0 auto">' + h + '<div style="text-align:center;color:#8a7177;font-size:11px">std = company standard per shift · usual = their last 4 weeks · 1% better every day</div></div></div>';
  const cards = scorecardsHtml_(wk, true);
  // one editable Gmail draft per person: tweak the game plan, add their email, send
  cards.forEach(c => GmailApp.createDraft(to, 'Your scorecard · WC ' + tag + ' · ' + c.name, 'Open in an HTML-capable mail app to see your scorecard.', { htmlBody: wrap(c.html), name: 'Kova' }));
  MailApp.sendEmail({ to: to, subject: 'Weekly scorecards · WC ' + tag + ' (' + cards.length + ' ready to edit in Drafts)', htmlBody: '<div style="font-family:Helvetica,Arial,sans-serif;font-size:14px;margin-bottom:10px">' + cards.length + ' scorecards are waiting in your Gmail <b>Drafts</b>. Edit the game plan, add the person\'s email and hit send.</div>' + scorecardsHtml_(wk), name: 'Kova scorecards' });
  return 'Sent to ' + to + ' + ' + cards.length + ' drafts';
}
function scorecardsHtml_(wk, wantCards) {
  const t = loadTable_(), M = t.map, end = addDaysS_(wk, 6);
  const get = p => (M[p] && !M[p].del) ? M[p].data : null;
  const reps = Object.keys(M).filter(p => /^reps\/[^/]+$/.test(p) && !M[p].del && M[p].data && M[p].data.active !== false && (SCORE_INCLUDE_TEST || !M[p].data.test)).map(p => ({ id: p.split('/')[1], r: M[p].data }));
  const dayOf = (id, d) => get('conf/' + id + '/days/' + d) || get('reps/' + id + '/entries/' + d);
  const sum = (id, a, b) => { const s = { stops: 0, buyers: 0, retail: 0, upsells: 0, hours: 0, shifts: 0, days: [] };
    for (let d = a; d <= b; d = addDaysS_(d, 1)) { const e = dayOf(id, d); const has = e && (Number(e.stops) || Number(e.retail) || Number(e.hours));
      s.days.push({ retail: e ? Number(e.retail) || 0 : 0, worked: !!has });
      if (has) { s.shifts++; ['stops','buyers','retail','upsells','hours'].forEach(k => s[k] += Number(e[k]) || 0); } } return s; };
  const kpis = s => { const n = Math.max(1, s.shifts); return [s.stops / n, s.buyers / n, s.stops ? s.buyers / s.stops * 100 : 0, s.buyers ? s.retail / s.buyers : 0, s.upsells / n, s.hours ? s.retail / s.hours : 0]; };
  const money = n => '$' + Math.round(n).toLocaleString('en-AU');
  const F = [['Stops', x => x.toFixed(0)], ['Buyers', x => x.toFixed(0)], ['Conversion', x => Math.round(x) + '%'], ['Avg sale', money], ['Upsells', x => x.toFixed(1)], ['RPH', money]];
  const STDS = [STD.stops, STD.buyers, STD.conv, STD.avg, STD.upsells, STD.rph];
  const M1 = '#6e1423', G1 = '#d4a24c', OK = '#1f7a4d', BAD = '#c2410c', MUT = '#8a7177';
  const bar = (v, max, col) => '<div style="background:#f3e3e6;border-radius:6px;height:8px"><div style="background:' + col + ';height:8px;border-radius:6px;width:' + Math.max(3, Math.min(100, max ? v / max * 100 : 0)).toFixed(0) + '%"></div></div>';
  let out = '<div style="font-family:Helvetica,Arial,sans-serif;background:#f5f1ee;padding:14px;color:#2a1a1f"><div style="max-width:560px;margin:0 auto"><div style="font-size:24px;font-weight:900;color:' + M1 + '">kova <span style="font-size:13px;font-weight:600;color:#9c2a3c">· week of ' + wk.slice(8, 10) + '/' + wk.slice(5, 7) + '</span></div><div style="height:10px"></div>';
  const cards = [];
  reps.map(x => ({ ...x, s: sum(x.id, wk, end) })).sort((a, b) => b.s.retail - a.s.retail).forEach(({ id, r, s }) => {
    const P = sum(id, addDaysS_(wk, -28), addDaysS_(wk, -1)); const K = kpis(s), U = P.shifts ? kpis(P) : null;
    const dmax = Math.max(1, ...s.days.map(x => x.retail));
    const rev = get('reps/' + id + '/reviews/' + addDaysS_(wk, 7)) || get('reps/' + id + '/reviews/' + wk);
    const gaps = s.shifts ? F.map((f, i) => [f[0], K[i] / STDS[i]]).filter(x => x[1] < 1).sort((a, b) => a[1] - b[1]) : [];
    const ROLE = { sa: 'Sales Associate', sc: 'Sales Consultant', sl: 'Sales Leader', tl: 'Team Leader', tm: 'Team Manager', ad: 'Assistant Director', dir: 'Director', md: 'Managing Director' };
    const role = ROLE[(r.promo || {}).stage] || 'Sales Associate';
    const joined = r.joinedAt ? mondayOf_(String(r.joinedAt).slice(0, 10)) : null;
    const zone = v => v >= 200 ? 'Diamond' : v >= 160 ? 'Platinum' : v >= 120 ? 'Gold' : v >= 100 ? 'Silver' : 'Red zone';
    const best = s.days.map((x, i) => [x, i]).filter(x => x[0].worked).sort((a, b) => b[0].retail - a[0].retail)[0];
    const DN = ['Mon','Tue','Wed','Thu','Fri','Sat','Sun'];
    const strong = s.shifts ? F.map((f, i) => [f[0], K[i] / STDS[i], i]).filter(x => x[1] >= 1).sort((a, b) => b[1] - a[1]) : [];
    const lower = x => x.toLowerCase();
    let summary = !s.shifts ? 'No shifts logged this week. Let\u2019s get some on the roster.' :
      (strong.length ? 'Strong ' + lower(strong[0][0]) + ' this week (' + F[strong[0][2]][1](K[strong[0][2]]) + ').' : 'Good effort getting out there.') +
      (gaps.length ? ' Next step: lift your ' + lower(gaps[0][0]) + ' so it can turn into more retail.' : ' Every standard hit, chase a PB next week.');
    const plan = [];
    const conv = K[2] / 100;
    if (s.shifts && s.shifts < 3) plan.push('<b>3+ shifts next week.</b> More time at the door = more chances.');
    gaps.slice(0, 2).forEach(g => { const i = F.findIndex(f => f[0] === g[0]);
      if (i === 0) plan.push('<b>' + STDS[0] + '+ stops a shift.</b> At your ' + Math.round(K[2]) + '% conversion that\u2019s about ' + Math.round(STDS[0] * conv) + ' buyers a day.');
      else if (i === 1) plan.push('<b>' + STDS[1] + '+ buyers a shift.</b> That\u2019s about ' + Math.ceil(STDS[1] / Math.max(conv, .3)) + ' stops at your conversion.');
      else if (i === 2) plan.push('<b>Conversion to ' + STDS[2] + '%.</b> Slow down the close and ask every time.');
      else if (i === 3) plan.push('<b>Average sale to ' + money(STDS[3]) + '.</b> Offer an upsell on every sign-up.');
      else if (i === 4) plan.push('<b>' + STDS[4] + ' upsells a shift.</b> Offer one to every buyer.');
      else plan.push('<b>RPH to ' + money(STDS[5]) + ' (Gold).</b> That\u2019s ' + money(STDS[5] * (s.hours / Math.max(1, s.shifts))) + ' on a ' + (s.hours / Math.max(1, s.shifts)).toFixed(1) + '-hour shift.'); });
    const hrsPer = s.shifts ? s.hours / s.shifts : 6;
    const base = Math.max(s.retail, P.shifts ? P.retail / 4 : 0);
    plan.push('<b>Retail target: ' + money(Math.ceil(base * 1.01 / 25) * 25) + '</b> next week (1% up on ' + (s.retail >= (P.retail / 4 || 0) ? 'this week' : 'their usual') + ').');
    if (K[0] < STDS[0] && s.shifts) plan.push('<b>' + Math.ceil(STDS[0] / Math.max(1, hrsPer)) + ' stops an hour.</b> That gets them to ' + STDS[0] + ' on a ' + hrsPer.toFixed(1) + '-hour shift.');
    if (rev && rev.plan) plan.push('<b>Their plan:</b> ' + rev.plan);
    if (rev && rev.mentor) plan.push('<b>Mentor action:</b> ' + rev.mentor + ' (book a 10-min check-in).');
    const wins = [], chal = [];
    if (s.shifts) { if (s.days.filter(x => x.worked).every(x => x.retail > 0)) wins.push('Made sales on every shift worked');
      strong.slice(0, 2).forEach(x => wins.push(x[0] + ' ' + F[x[2]][1](K[x[2]]) + ', at or above standard'));
      if (U) F.forEach((f, i) => { if (K[i] > U[i] * 1.1 && wins.length < 4 && !strong.some(z => z[2] === i)) wins.push(f[0] + ' up on their usual (' + f[1](U[i]) + ' → ' + f[1](K[i]) + ')'); });
      gaps.slice(0, 3).forEach(g => { const i = F.findIndex(f => f[0] === g[0]); chal.push(g[0] + ' ' + F[i][1](K[i]) + ', standard ' + F[i][1](STDS[i])); });
      if (s.shifts < 3) chal.push('Only ' + s.shifts + ' shift' + (s.shifts === 1 ? '' : 's') + ' this week'); }
    else chal.push('No shifts logged');
    if (!wins.length) wins.push('Showed up and logged it');
    const card = '<div style="background:#fff;border-radius:14px;margin-bottom:14px;overflow:hidden;border:1px solid #efdfe2">' +
      '<div style="background:' + M1 + ';color:#fff;padding:12px 14px"><div style="font-size:10px;letter-spacing:.14em;color:#f2c46d">WEEKLY SCORECARD · WC ' + wk.slice(8, 10) + '/' + wk.slice(5, 7) + '</div><div style="font-size:21px;font-weight:800;letter-spacing:.03em">' + (r.name || 'Rep').toUpperCase() + '</div><div style="font-size:12px;opacity:.85">' + role + (joined ? ' · joined WC ' + joined.slice(8, 10) + '/' + joined.slice(5, 7) : '') + ' · ' + s.shifts + ' shift' + (s.shifts === 1 ? '' : 's') + ' this week</div></div>' +
      '<div style="padding:12px 14px 4px;font-size:13px;line-height:1.45">' + summary + '</div>' +
      '<table style="width:100%;border-collapse:collapse;margin:4px 0 2px"><tr>' + [['SHIFTS', String(s.shifts), Math.round(s.hours * 10) / 10 + ' hrs'], ['RETAIL', money(s.retail), best ? 'best ' + DN[best[1]] + ' ' + money(best[0].retail) : '—'], ['RPH', money(K[5]), s.shifts ? zone(K[5]) : '—'], ['CONV.', Math.round(K[2]) + '%', 'std ' + STD.conv + '%'], ['STOPS', K[0].toFixed(1), 'std ' + STD.stops + '/shift']].map(([l, v, sub]) => '<td style="padding:6px 6px 6px 14px;vertical-align:top"><div style="font-size:9px;letter-spacing:.08em;color:#9c6b74">' + l + '</div><div style="font-size:19px;font-weight:900;color:' + M1 + '">' + v + '</div><div style="font-size:10px;color:' + MUT + '">' + sub + '</div></td>').join('') + '</tr></table>' +
      '<table style="width:100%;border-collapse:collapse;margin:10px 0 2px"><tr style="height:46px">' + s.days.map(x => '<td style="vertical-align:bottom;padding:0 4px"><div style="background:' + (x.worked ? M1 : '#efe3e5') + ';height:' + (x.worked ? Math.max(4, Math.round(x.retail / dmax * 40)) : 3) + 'px;border-radius:3px 3px 0 0"></div></td>').join('') + '</tr><tr>' + ['M','T','W','T','F','S','S'].map(d => '<td style="text-align:center;font-size:10px;color:' + MUT + '">' + d + '</td>').join('') + '</tr></table>' +
      '<table style="width:100%;border-collapse:collapse;margin:6px 0;font-size:12px"><tr style="color:' + MUT + ';font-size:10px"><td style="padding:2px 14px"></td><td></td><td style="text-align:right">week</td><td style="text-align:right">std</td><td style="text-align:right;padding-right:14px">usual</td></tr>' +
      F.map((f, i) => { const v = K[i], ok = s.shifts && v >= STDS[i]; const u = U ? (v >= U[i] ? '<span style="color:' + OK + '">▲</span> ' : '<span style="color:' + BAD + '">▼</span> ') + f[1](U[i]) : '—';
        return '<tr><td style="padding:3px 14px;width:22%">' + f[0] + '</td><td style="width:34%">' + bar(v, STDS[i] * 1.3, ok ? OK : BAD) + '</td><td style="text-align:right;font-weight:800">' + (s.shifts ? f[1](v) : '—') + '</td><td style="text-align:right;color:' + MUT + '">' + f[1](STDS[i]) + '</td><td style="text-align:right;padding-right:14px">' + u + '</td></tr>'; }).join('') + '</table>' +
      '<div style="padding:4px 14px 12px;font-size:12px">' +
      '<table style="width:100%;border-collapse:collapse;margin-bottom:8px"><tr style="vertical-align:top"><td style="width:50%;padding-right:8px"><div style="font-size:11px;font-weight:700;color:' + OK + ';letter-spacing:.08em">— WINS</div><ul style="margin:4px 0;padding-left:16px">' + wins.map(w => '<li style="margin:2px 0">' + w + '</li>').join('') + '</ul></td><td style="width:50%"><div style="font-size:11px;font-weight:700;color:' + BAD + ';letter-spacing:.08em">— FOCUS AREAS</div><ul style="margin:4px 0;padding-left:16px">' + (chal.length ? chal : ['Nothing under standard 🎉']).map(w => '<li style="margin:2px 0">' + w + '</li>').join('') + '</ul></td></tr></table>' +
      (rev && rev.mentor ? '<div style="margin-bottom:6px"><b style="color:' + M1 + '">Wants from mentor:</b> ' + rev.mentor + '</div>' : '') +
      '<div style="font-size:11px;font-weight:700;color:' + M1 + ';letter-spacing:.08em;margin-top:4px">— GAME PLAN FOR NEXT WEEK <span style="font-weight:400;color:' + MUT + ';letter-spacing:0">(suggested, edit before sending)</span></div><ol style="margin:4px 0 0;padding-left:18px">' + (plan.length ? plan.slice(0, 6) : ['Keep it rolling and chase a new PB.']).map(p => '<li style="margin:3px 0">' + p + '</li>').join('') + '</ol></div></div>';
    cards.push({ name: (r.name || 'Rep') + (r.test ? ' (test)' : ''), html: card }); out += card;
  });
  if (wantCards) return cards;
  return out + '<div style="text-align:center;color:' + MUT + ';font-size:11px">std = company standard per shift · usual = their last 4 weeks · 1% better every day</div></div></div>';
}
