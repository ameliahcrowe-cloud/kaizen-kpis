/**
 * Kaizen KPIs — Google Apps Script web app.
 * Paste this into Code.gs (Extensions → Apps Script in the Kaizen KPIs Google Sheet).
 * All data lives in this spreadsheet. People log in with their name and a PIN.
 */
const APP_TITLE = 'Midas KPIs';
const OWNER_NAME = 'Ameliah';          // the owner's login name
const OWNER_UID = 'u_owner';
const DATA = 'Data', ACCESS = 'Access';
const CHUNK = 45000;                   // characters per cell (Sheets limit is 50,000)
const MAX_CHUNKS = 20;

/* ---------- web app ---------- */
function doGet(e) {
  if (e && e.parameter && e.parameter.app === 'cards') return cardsJson_(e.parameter);
  if (e && e.parameter && e.parameter.app === 'kova') return HtmlService.createHtmlOutputFromFile('Kova')
    .setTitle('Kova')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  if (e && e.parameter && e.parameter.app === 'pb') return HtmlService.createHtmlOutputFromFile('PB')
    .setTitle('PB Board')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  if (e && e.parameter && e.parameter.app === 'hourly') return HtmlService.createHtmlOutputFromFile('Hourly')
    .setTitle('Midas Hourly').setFaviconUrl('https://ameliahcrowe-cloud.github.io/kaizen-kpis/icons/midas-512.png')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle(APP_TITLE).setFaviconUrl('https://ameliahcrowe-cloud.github.io/kaizen-kpis/icons/midas-512.png')
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
  // hourly tracker -> day log every hour (server side, no app needed)
  ScriptApp.newTrigger('syncHourly').timeBased().everyHours(1).create();
  return backupNow() + ' · hourly sync on every hour';
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
      const body = Object.assign({}, cur, tot, { date, week: mondayOf_(date), month: date.slice(0, 7), hours: (Number(cur.hours) || 0) > 0 ? Number(cur.hours) : (Number(doc.data.targets && doc.data.targets.shiftHours) || 8), fromHourly: true, roadie: !!doc.data.roadie, updatedAt: nowIso }); if (String(doc.data.cash || '').trim()) body.cash = Number(doc.data.cash) || 0; if (String(doc.data.eftpos || '').trim()) body.eftpos = Number(doc.data.eftpos) || 0;
      if (!body.campaign && rep.campaign) body.campaign = rep.campaign;
      out.push({ path: ep, data: body });
      out.push({ path, data: Object.assign({}, doc.data, { syncedAt: nowIso }) });
    });
    if (out.length) putDocs_(t, out);
    try { sundaySummary_(t); } catch (err) { console.error('summary ' + err); }
    return out.length / 2;
  } finally { lock.releaseLock(); }
}

/* ---------- Sunday 8am summary email to Ameliah + Ben ---------- */
function sundaySummary_(t, force) {
  const now = new Date(); const day = Utilities.formatDate(now, 'Australia/Melbourne', 'u'); const hr = Number(Utilities.formatDate(now, 'Australia/Melbourne', 'H'));
  const today = Utilities.formatDate(now, 'Australia/Melbourne', 'yyyy-MM-dd'); const wk = mondayOf_(today);
  const flag = t.map['settings/summaryMail']; const last = flag && flag.data ? flag.data.week : '';
  if (!force && (day !== '7' || hr < 8 || last === wk)) return false;
  const get = p => { const d = t.map[p]; return d && !d.del ? (d.data || {}) : null; };
  const reps = Object.keys(t.map).filter(p => /^reps\/[^/]+$/.test(p)).map(p => ({ id: p.split('/')[1], r: get(p) })).filter(x => x.r && x.r.active !== false && !x.r.test);
  const money = v => '$' + Math.round(v).toLocaleString('en-AU');
  const rows = [], missing = [], newbies = [];
  reps.forEach(({ id, r }) => {
    let rt = 0, hr2 = 0, st = 0, by = 0; const worked = [];
    for (let i = 0; i < 7; i++) { const d = addDaysS_(wk, i); const e = get('reps/' + id + '/entries/' + d); if (e && (Number(e.retail) || Number(e.stops))) { rt += Number(e.retail) || 0; hr2 += Number(e.hours) || 0; st += Number(e.stops) || 0; by += Number(e.buyers) || 0; worked.push(d); } }
    if (worked.length) rows.push({ n: r.name || id, rt, rph: hr2 ? rt / hr2 : 0, st, by, days: worked.length });
    const miss = worked.filter(d => !get('reps/' + id + '/reflections/' + d)); if (miss.length) missing.push((r.name || id) + ': ' + miss.length + ' day' + (miss.length > 1 ? 's' : ''));
    const j = String(r.joinedAt || '').slice(0, 10);
    if (j && j >= addDaysS_(today, -14)) { const sc = r.trainScores || {}; const k = Object.keys(sc).sort().pop(); newbies.push((r.name || id) + ': started ' + j + ', ' + st + ' stops this week, ' + money(rt) + (k ? ', latest scorecard ' + sc[k].total + '/50' : ', no scorecard yet')); }
  });
  rows.sort((a, b) => b.rt - a.rt);
  const tot = rows.reduce((a, x) => a + x.rt, 0), th = rows.reduce((a, x) => a + (x.rph ? x.rt / x.rph : 0), 0);
  const tr = rows.map((x, i) => '<tr><td>' + (i + 1) + '. ' + x.n + '</td><td>' + money(x.rt) + '</td><td>' + (x.rph ? money(x.rph) : '—') + '</td><td>' + x.st + '/' + x.by + '</td><td>' + x.days + '</td></tr>').join('');
  const html = '<div style="font-family:Arial,sans-serif;color:#1d2433"><h2 style="color:#151515">Midas week · WC ' + wk + '</h2><p><b>' + money(tot) + '</b> retail' + (th ? ' · <b>' + money(tot / th) + '</b> RPH' : '') + ' · ' + rows.length + ' people</p>' +
    '<table cellpadding="6" style="border-collapse:collapse;font-size:14px"><tr style="background:#151515;color:#fff"><th align="left">Person</th><th>Retail</th><th>RPH</th><th>Stops/Buyers</th><th>Days</th></tr>' + tr + '</table>' +
    '<h3>📝 Missing reflections</h3><p>' + (missing.length ? missing.join('<br>') : 'Everyone reflected on every day worked 🎉') + '</p>' +
    '<h3>🌱 New starters</h3><p>' + (newbies.length ? newbies.join('<br>') : 'No new starters this fortnight.') + '</p></div>';
  MailApp.sendEmail({ to: SCORE_TO, subject: 'Midas week · WC ' + wk + ' · ' + money(tot), htmlBody: html, name: 'Midas' });
  putDocs_(t, [{ path: 'settings/summaryMail', data: { week: wk, sentAt: now.toISOString() } }]);
  return true;
}
function sundaySummaryTest() { return sundaySummary_(loadTable_(), true); }

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
  if (!p.getProperty('SALT')) p.setProperty('SALT', '4785cac7-f05f-401a-8f7e-5521d93305c6');
}
function hash_(pin) {
  const salt = PropertiesService.getScriptProperties().getProperty('SALT') || '4785cac7-f05f-401a-8f7e-5521d93305c6';
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
  return { name: main.teamName || 'Midas', logo: logo.data || null, teams, hasOwner: readAccess_().some(a => a.role === 'owner') };
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
    if (!acc.some(a => a.role === 'owner')) throw new Error('The app isn’t set up yet. Ask Midas admin to finish setup.');
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
        if (now !== was && !(seg[1] === c.uid && now === c.uid)) throw new Error('Only Midas admin can change this.');
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
    if (action === 'add') {
      const v = value || {}, name = String(v.name || '').trim().replace(/\s+/g, ' '), pin = String(v.pin || '').trim(), role = ['rep', 'tl'].indexOf(v.role) >= 0 ? v.role : 'rep';
      if (name.length < 2) throw new Error('Enter their first and last name.');
      if (!/^\d{4,8}$/.test(pin)) throw new Error('PIN needs 4 to 8 digits');
      if (acc.some(x => x.name.toLowerCase() === name.toLowerCase())) throw new Error('That name already has a login');
      const t = loadTable_(), team = v.team && t.map['teams/' + v.team] && !t.map['teams/' + v.team].del ? v.team : '';
      const nu = 'u_' + Utilities.getUuid().replace(/-/g, '').slice(0, 20);
      appendAccess_({ uid: nu, name, pinHash: hash_(pin), role, tokens: [], created: new Date().toISOString() });
      putDocs_(t, [{ path: 'reps/' + nu, data: { name, team, active: true, order: Date.now(), joinedAt: new Date().toISOString() } }]);
      return { ok: true, uid: nu };
    }
    const a = acc.find(x => x.uid === uid); if (!a) throw new Error('Not found');
    if (a.role === 'owner' && (me.uid !== 'u_owner' || action !== 'pin')) throw new Error('Only Midas admin can reset an owner PIN. Your own PIN is changed from the spreadsheet menu.');
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
const OREXIS_LOGO = 'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAMAAAADACAIAAADdvvtQAACsJUlEQVR4nGz9d7hu6XUXCK437b2/dMI954a6t3JUlVQKpSzLsizZcpRlS8JJOEDTNA1+mPEw0IZphqZNwww8M0DPMNPdPB6SbYIHTIMAB7mtYGWpVFKV6lbVrXBzOPfEL+z0pvnjd9aqXWLOH/WcOvf7dnjf9a71W7+V1NmzZ0MIOeecszEmxqiUMsYopYqi6LqOiIwxXdfFGIuiaNvWGFOWZc5Zax1jlC9qrUMI1tqu66y1REREOWellFKqLMu2bZVSOWdrbQgB18eluq4riiLGOJlMuq7Dx5qmcc7hCvgKronbOee89yGE8XistW7bNvNPSqmqqhCCMYaIvPdE5JyTt8MftdZaa9xaKZVSwl+UUng1pVRd19ba73jTsizxv865siy7rsOLy2fwDFrrpmm01tbasiyJCEtNRH3fF0Xhva+qqm1bXARXSCkppYjIWptzxr1ijFgxIuq6rqoq770xJoSAxem6DvvS9z2u4JzDU+Gy3vsYI74ym83atm2aZjQa4WVl3bA+KSW5r1JqtVpNJhPvPRYNt8s5V1Vlcd3RaKS1Xi6XzjkiSinhciEErXXf93gZ7z0EC/+Ep8HS4K7YV8gfnqxpGnwAwgTBgiyKJOF/67rG52OM+JhzDhdJKfV9jyvjlfAmEDWlFEQTT6uUgvTgY9baqqqwMZDFnHOMEbuI9e37Hk84Ho9TSl3XQZLwIrgmngfvKCvb9z12Dg8JEUkpQTJERlNKbdvK2nrvIbL4Ygihqiq8WgjBOYev4F5YZCKC9OCNcHc8G+QVpwKLiRXA3UMIXdfhHOIWMcajo6OiKLTWRCT3kosbY7z3WPPJZAJZxGcmkwnWCg9DRMdLgy/Ixsta4+lxUWstrgUJwPZDsIiobVucOWxV27Z4bXwLC4cn67oO8icvjNfAQkB6sIs4ZHJfYwwObt/3ZVmWZQkF0LYtnn80GlVVpZSCOsRr932PjcHph8hi7cbjMT6PU1UUxXw+b9sWwo21ttbiXbBDUD9YPnkqvAj+FUrRGNO2LRHhObFbInxYMWutc265XOLY4E2h17FDWE9jjNwF/3u8cyzfbdvisrg+zt5qtRruLE4vXn82m5VlCUWANcS7VFUFpS5yj6XGFXBEIRvQeXgRderUKVwaOoCIqqrC82GhcVccGqwU/tj3Pd4ZZ05MAwQOi4jvwgLi3rgOTknXdePxGOv+HYZJZGg0GkG3YQNE0EejEQyfmC3oieGRLcsSayGGD58UcwaxxvrCnuIh5SDhxeVE4S54NrkvLEXbtqPRyDnX9z2UnKyYHAzYIxgs7BP+FU+ONcQvkKS+72X9h2ZBDi3WZLhT+AD0DQ6zWDEIQVEUk8lktVphofA62AKsg/yUZSmH/Dv+LtpRy7HACxtjoEvrusbX8EqiNgUPQd9gFXBjOe5yTPu+xxJgF51zRVFgS4AqIKmwNXgTXKSqKqwm1ANUCM4W3lkwAe4FEcEC4ZmLooDYjcdjrCmOIGRdroYzimdYrVbAELD9+ApUHRYHth66E1/EfXFZAWR4BahG7CK0CyQjpQTjhZXEKcLJwZJCSlarFWQCSkIsJnYKiwyxw6XKssQKG/6BWJRlOZvNsBqj0WgymQBC4fFk6bBWRVHguEIn4X3FqIndF7CIw6O2t7fxZLLxWCPACNxAcDGugv0WmInLyV7ixKSUgNGwPbPZbLFYwEx47/GI+AEqh80GngW4ATKD+AIkwZzjJXEEsZp45r7vIZH4CkQQZx3qEGcXF8feYGOG9gKPJ14FNkkkXmwTnh+fgYjjauPxGKcFiwvcBoUEzSRCg9PSti12GmI0XExsingeWHyIKdAexALPINpUDIIo1Nls1ve92C9gBkgbHgkLhQsKzIBcyjkRv6Hve3wSKxNCUNvb2yKGUBhVVQFdYlewHFggwbA4KMDXeD3ocAH/YnHwoPg6Dj22B0d8NBrhlSAQYtpF5+GOkBusCCQJH8P5xu5i57AKolCxlERU1zW0C37BKceH8XWcfmw23gIL7b0XR5LYjRIEKl4k/lWOGb4ITWCMWa1WzjnnHGQXziawC5SH7JksBT5WVRXkA+8C5YGFEtMjvh60OCy4yC52Vq4P1SI6Bk97rEv4BQX7yr7jkbAXkFR8JuessR84eeKhtW2L1xMXCVfEaYAM4WaCqUejEV6prmssOrwJPC4uDquPh5CjI1pNYAduIXYERg1OBAwo9rtpGuAM3BfIV5wvHCCIOPQZFDJ2Aq+Mv6SUVqsVFhRnA2+NdQeGkNcUQDb0ipumgdnCskZ2xfFfmDCcKzzVaDQSM4TjAawqPrm4IzDfzrnFYuG9h7OMNZG9xPVls7EsuKNsFmAQro/tEO8EykLUJDZddJ6oOjyegAcoAu+9hdaCphHAi+Mi5xJaEYcS7ybvKaekbVvAajyKMWY+n4tBPFZ3TG8kJioEluLvkb1rWWts23g8Xi6XsFBYJvhcYqHxVkO4I9g5MRkBSRXULM65KCGB2PivKNrE/rkIMUQHF4FpwwtCKYK2EGME5V/XtewoFA8OlayJPDB0OVQLjjfAL7YGwE6cc3AQuEjTNFgfHCGwaHgwrAAYDbHveC8ATeygtXY0GjVN0zQNznPOGeQOthIveKx7tHbOHbuFEFgIAbZzuVyCBsB3RHGJ+AP04YtDwcLSwDvFBgjeHFIOMHACv8TDgqUD3INWgPiLNQG3KbambdvlcgkUNRqNyrLEaROdPJlMqqrCS0GdQA5goMuyxAoAhRRFgV0ROAIOELwOrimnWRA93gIbKT4/TjPeVzM5iWeGe4x/lasNcTHUGwQFpmA8Hst7gSoEZMYbDY0v1CoUM5ZaYKscSCghUe3iB8hxwrEUNw1vJy8C4T6+5tbWlphJyDvEDdhCXgznTzFNKYhSMI1iUgRvArkZUtLyklhHQVHAT2KD8GExZzhVWAhcXH4ZYlu8j5hFw1wwZAVIBX/EY0fmkbEWghLw4jDocBjFCcJn4GSJKEBd4SK4nbwvjh/umJnow/bgDKQB2yuAXcA7kIPWej6fF0UhxCC2A8oeQAeMBtZQ4BfcFyxCYj5JHgZXFm5MjqgYFlHtQseIYYGxArjs+/74tXEhvAnAGtYdGliUEM63KPnM5Ld4N957hBRwDnBkRbsM3QTNlCPkVbBnZF5RFgW7CAWA94GCFXyNx8AO4cFE8wnqxzWxLoqpVZiSoTqEM4+VhS4EIMWD4SzipompW6xAURTT6RSnC4+BWwM6YFnwzDDQeE3cCOoTu4DDIPYUEARKUaI6Ygogf03TIJaCW+NhBLBDnWCJsL+ixfG/IlI4/wBV2HTD3CmWCw9PRGBkjmHG5ubmeDwW8h6iigfCojvnxuMxYPXa2tpiscDqCDCCWMhxBxwZjUYiJdhIcc3kvMpCi+zDe5JnwJXxMWgL8AuwZXB/Qgh4PIFZsMWQG6GMRXvj1IqvgQUdxpvAX+CxAQQFPmumoMDkYhtwO2ynqBDcJXIAC38UUYOg4O3wFzEi0LhYebxgYgJQRAfbbDiiN8SR4kvKL7gdDjNsseJAEMQRfxmPxzgn0Nb4O7hEzSFOLNrQf9Jaq9OnT0eOSyimqoanuWkasIsCSAXJiqrA+wztMTHbGzhkJrpKbB92TtwZrLJYa2ttXdcQHVgTMV64Mn5wLvGogHt4WjgjYF0FqMmRCiFMJhNspEgYRFOwsyAty7FD8UeIeRrRHKvVyloLYh3ocH19fbVaARgIXpHjlJkaEAiIG4k9FW8OugS7JU8rSy3Wxwy4PjCxIuvw84dQNQ5ocbGVeB5Ip3hO4FlwX7EAEPFjHgskjbg8KaW1tTXgQUj3eDzGN4EMID1yFkHPtG0rLgOeo21b7BD+CGyOmDMEHPoQagZfwbPCDoLIFgOEMyrCKsZbDDmMIIQJq6M4XiNhS5htATS4Pm4q2AtrlzmiLqjfcNRTDB+kSkwbMAcOA5YISlFSAIYOhChmeFuiOxX7xviwAOqiKBaLhQRGhvpAzp6AdGyTyB98K2GTJfSLW8jfIYh4F8gWbg2hBxSRy45GIzxSCEHdcccdgjfxUXGawBHjZjiX4jhA10ExCC07fBNBzULbTKdTAHPFvCcYUpyw6XQK/yUynysQRBzmoWRgpSC4UKeiRwPTGIJ4BEWmQbAFxhcPIEAbZwbXAZmL0BVkKDPxLVgVkgcLCJUQOHKEk4nHFl8Gv8BSQ4HhMeRFcDXYROHP8FJiJRTH2mC2DDOHsuuGswMkiu6cW61WeAtJjsAncZZE8kTpYovF3MO9B34fLogVPwLHBeuI1RefDfuNJcApJ2ZCISuBOR6xEWINLRPKgBdwzcStE9Yb0HvoxOHuw6AHVFHkdA6xm5rD13BxZaGFZxIDFDilBmst6FhwAyg+CJakaoCwJs5CEetsOPhAgzAO1logVGCKCIowMKmDNYR5tRyVA+ZABhI2SWwTjgS0pkR1BBoqzgzB8cAnRW3gMcTlhA4W6guLgPcawlZRn1hDwAPxfAWVa2HtBE7imAoSqqpKuDvDQSggDOI4nDgOkWlvPKUECMVyA3gOn0/sC3HgFwZUHNHIXC3sGu4ukB+igJWCuoYHIYADsjWbzfDA4n/ilCOCBqmFJsCLwyBaDt7h8USfCYKBCcByicR3XQddC4VfFIWYYGw/LiXJJ/jBsivm4r33EtLGO0L3y/mRvYCK0hxAdJz6o5QCRS5aEE8IUc4DPxr6DL83TZM56gw4IZZdNKh4cEopK4oLcprZdcwcupMURPwSmUSKHB5PnD8lvlhKCZ+fTCZFUYBrF5yIO+ac8XfsMVYKsIk4GU8UNZZezornDEO8Z2LGD/JKnHsE4QtMQ4sPCIwFoAalRcwpy5pGjq4AlIhKC0wSWo6g4cpi0+W7iRMKhCLBITxx4gTgkeAtAXOKIw+CBLDZoC2GVjgxlY+HEZECipcrQMEj/CcqU/xQCLcE+2QZZR+JE0qBymE6MqeUHH8FXrriPBjHWX+iD6EnBDA6zhmQFxAgBjwrUKCqKjEusnlyygXzi9+Ed8M/BQ4pi5rBvWAEQZ2JuydWD6EiUYoQRHEryrKEVuu6rus6WAqJVkIt4Wo4MFgTWEZYQOgVYm4MBxFHQuha+DsS3sH2YCm01uPxGEc8cmIGlheGxhiDI4c9I07GxUYCNkHUhGMTHIZYpBAEnpNP8C6wdxIKxBHa2NgADsPtZONgKMXzijFCF8qpSJwqGWM8BtGZqUbkYAiRI2F5fEYMcGaaLg9CReKqiBUAiE5MSMBciqEEhaOUWi6XiE4L9w09Dw0xfLzhshpOjRDsJSdvCBLjIHYrQdnM0YzEnrDsE/4iPgsAEEQZe+840RaaCecb+QKag/CRg3o4+sQZI4IWxD3+/5uwIfZusVhozjvIzCxgkcWjxFHEgcHDQHVBaDSH2ES3ac6hEIgtKTR6kC4HMdIDeloWbegkaTl2svF4UBFAaAWAIXEvhz6FeKTY0cj8pogdtJ+gZlxHcW61uA9QDzjcOM3EkVdikj5zXjZ2N6WEA5Q4guM5ERiPBAwOc4P1Upzo2TRNXdeRM0EFatBrWRmsNeKL0A2SQZs4k07UpwirMIRI8kwDFlh8KDxhGgTexS2FjhF8oziMgCvgdOEcihAAcSKjRvRQYDYYUohjhhWG0IgxCUwLQwMF5sdxU+gzuebQIqk777wT+y14VkC4oA0acLjr6+vz+fw7bixklOBW5O5Mp1PJWgocVSGmKLH6dV3PZjMI6HK5lFNI7LpLKlJgtl5WGU8IUYNWA2siay1OoiB9WKLMzG9m70yMr1h9ARZinnCsxV3AXYTml8fDB6BdIjOH0KZyxgKnNOGTmlOCzCAHHGuIK8B0CuIW2CrLK4yX5QAiHhI+NWDD0J1MKU0mE3k8UXuCEYE4lVL4HZ6dxAqJyWR1+vRpvKpgPQmdyNM75/A+gmaAmpfLJY4XeFIcMnjCMNLATM45fBIxB3HFIX9xkBoh4gLsAtheFAV8SKyX4ZBCHtQh9VxfATnQWgPiyPYortmA0Ah7K7sisWgYe+BWHGh8V8gS+RYWDZ8XNAN5xSnHU0HmBE9gg/G0spfiefScaS6BSNEiEirBk4DQW61WYpdFMRDTK2C2YNqwpGIEobcACQSQee+n06lEbSWUifMjGO41OyXWVMJvNEg0dpyyD49O3Bml1Gq1ElQknqfoBsURCWyGODjElRj4YuTsacliwU7g9SCXQkmLqAWuPoM0iGxBx+BqiEsI3y9Si6u5QfJXZidWQDeuL4dKNl4AVmKOVFIE5fULrncRp4H4JzIDrpnUFqUIZUOcWp5zHo/HyIwWs5I5wUGwJi6CfxJ6UyAzTrL48CI6NEgPX61WqGQS6RfpMcxlI7YPweq5wAuqJOesTp48KThLM30JggsyDqYEj6s5XqE5D0bkQ1IgBGSJr4Ef/JOweaKZwZfLZXFNbDMWcejSD8nWxGS/5VRfxXml6+vrgaOAmcnMIZAUnhoSiZAQdBLe2hhT17VcHCcEslXXNRjkzA6/SLzYzcjENP4XpHnkmpuhhYVYiOUFusK/RuYt5SdxdBnWhLjGwXCkEuIuOAHHDKdUNBnefZg2CGw6n89FVePzoguqqkJWbmQqUqRZra+vS+mCmHaBpRDbnhPCPVdSCjOGR4ErK7EwkXExsbLKQEKSbwSRhTvmvfdc9IP1kieRZAFITOKgNCSJmK7QgyITOH2Wy+SwhUBanjNrBXqnQa2FMK14O7wvpErivjDQgtYFNaNmVMod5fWxHwLhhSkW+cZXUAApdlmEXhCnYEH8DqsntvhVYDv4EXsioYXMSUthkGgKmCEOpkgSFP8QhGEBRTUc19YQc3fOOZxyORxiU4eLFTlACCpFEstxPwAIIUsC59UaznCD2sceAFpKWrEsOhgj4VHAC4vjLQpDMuWGtJDmalo5u23bjsdjSFXigGjJJYLAZ7B0wJuWS1CGtAKkX8xc4OS7IQQW0c/sutpB5avoY8NZ657T3bFEipkIz0VC4nNlLuCE+wkhzlwojWxSAdcpJdT4iv2iAdkB4kDeHYp2NpuJOc45i8rEfokRDJwzc+xqbG1tQSCEhpLHSlzDhqWRYhTPqaiQBnAPfpB7ClgjlAac1chRlTzIOKZBKrSsOE4SDrSEO2Dme64SgTXEOmJp4DMDTwDlQNrkUYcoB6hwKGHif8lZEkSsOQQRuKxHdCHWEQy45KYlzroXcKbYDxcLJcuIe8mZxBlDlXfm1APcHTIEcpkY+QoxFjkQLiBGVBc8fImoyFEnTsEQ7EUcg8rspRaDNFHP+ZmKHcAotYVCncmhFBdGTNL6+npm/hAiqZmAl0eRvEHcyXESE3EAHAdFD6JC2GORDAg+9g9hOMuFmFhux/kCQ1qPOO1VPOo8KIQb6nPcizi7UhhnwQ15QDFMJhORlTgg4sQuw1fSg6iz4mAWsU8HRC//Kh8Qhy4PImuI/2P9rbXI8pYPJC6ildeBgpSjKMSYZt4LrwY+XTGIDoMcB+GZmqYJXJhQcBqkGaQHZa5KIA6GamEsaECbCgc41HLg/hNzlJAe7DQqzPG4dV0D/wKxA9aIx4GXFwpLUDMYMKHn5ZQLqMoc34B9jFwaR1yLI1gVG4+QkBxTPK3lemHFWbaag1kF53yJ6hUkYbn2JXBmHLY2D1KO5FBia+Gd5UGeHXQqTimMvuJwJn7AHPZc1iLbhssKzIBPo5h/l88At2HRoCew93gezcQVngqWBP8qhhLfFVUN1CjnXCw1DXJXiqIYjUYa3CBWFmxB4khk4HRgPCINCGth3ojDNLBW1tq1tTUcR8iK4nJMw4nreFAUOokE4DQAS0ICQgggxDzn6ED/iygkjrTA5iJ9MXMaORGhpCTyj2JmIcaIw4ASNpEYGji60BkSZzWDLH34NYqdMkFFEDg8pCSPS9hclqJtW6Tm0MAdMxyuVgMuV8w63IiSK+oBD6DY4C1GZsU0BzQUo+w4CKEL2AKWkluEEBBK08wcShwTj4rtLrgyXXB327YaDHrkOmIhqmWrhho7M9OA33EOsNNhUKQROOA1RF4lN8gRRkCEVXMqnRhsCIEEivFJyxFgKCrHxUoil4oLYoT7gkq3gxQW6C0h6eXDiVlsXA0qB7s1hCZyfRxZGpSGJW5hMZ1OHVc0w+2HsMYBLQT/IHBeJX4KTnzG0mUmPBPXgYi5B4RAgEhONdYcm4JXk+C0oBmQF3hayBlxTYGYM3H0cH0YKBCBQ8cwpbRardTJkyfxJ0nvVVwjpzhnQ5YeD02DziYAIplpHkEhJfdwwD+BU85MM4qRAkmFD2ith7mI2DnNCYcdd50SBJY4DhoHWfp4YazCdDrFEmDthK3GAhVFcXR05DidEjskfpBgc/lXJO8GrmESggACgUM1jPzIMsrr4zNIvxSvggYuW9/34/E4MuGpOcEocM8GWEwJp9CAuovcPkDce1hScDRQb6IIkTKQODM/cCkIcQxKIjl6kL/hOHFAzEhGexc4O6jNEEcOjkDkzHnJg4Q8CrUj5DUxgMe/CvsnS1DXNRpEoI5YDA3cCsF0flDFbbljlzBacl7DoFoAB8gYg1RAOXaGY5yaWz7kQZaMH9QzRE4yF60zjPhCuPFJaCb8ApTgX9uKYIgRxcsTbUTMgxgukbGDZBg1yAb2g8osvEXJpW2yKYkDsbAscgB6bsYlJAi2BsBFephoTigIgzZiQnWKcYS3O2RJJEQYQjhGSTAckA9YEDggjrNwJFYq0Cxy0VPkaA5e2HFlp4C7yF1qyrJEoQKgbuYwbxqU6uVBVrUcDuligY0BNlLs96kBa0XMIQm8FesgdBE8CMu5i4aboeDdZRtkZTRHG+R0QrUYZjIlgRN6Xg6oH3QTy+yrixaR/CSwOOJAiN4SPZc54ts0Dexy4JRWKD+Jz2TOqRXoBonBNuEF4dYR0zTiOUl+iOyp4fwwyxU7eDXIgyDX4ypDSDF+94OWFIppJVksKQ2DyYcgixGF/sS/wqB4Tt5GsH1oE+EuQfwzBxTxGN573Ci8trJd6BlgfGw5OrFJaop4/o6zrobcCXHeT+AiEHG/oWilrQxAj5xOxbQ4cQQ7DIooIGqj0QhWA5+XJEmcQMfpTXhCmDYzKDaSTYJ+reu6HfRHU5x6OzxjinMXE9fi4AnhVUDiRby899gFwYjYRPFmRNMHrv34DhZUiINjlAwdJcfLcO4pTJhIkrB2uC4kBuZJD8oCBdV/xy9ikhS3VEKUgDikoDkOBZuI44UKebFWmtO1EldJIxAIKCYqx3Ci2RBvQo0JOSao03BKqBh7w6FKw2EyADWJd8KiQbmKVsaHAXeAJomTG2OMaCWTc57NZoINMmdKYYlg4vF3wFs1yBEgZpuGi2+YiBcAIPWBonrzIB8hcjoyJBhRGuK6DkgPGE7J+COOVwoGEqRxDMsi58vJGU2DCBEeInG+d4xRyiXLQZmz5pAWrCzEBcZF+Guhlyx3CBCpd1x3lzjjHXYQUCMxyUucAy8HwnD1VuQKEJEtHDvoGNzCclWJ/Ndy8x7sriRkHR4eDtPdsbJVVX1HFqzApvzavlstN7GTXRQPFEGeYlCcKu6nOFNd16EEQFCO+s+I+DzIJhPFjOWy1iKtFNQavBzRoFhV8fzF/Inrg/OAlwUGx5PEQQ8JUcbyyq+m2SZu1GW5PSDWS1AqaCUxsYErkgR/ZU7uF0wNqcrcG8VyM5fMmSHiK4m7jr/jICKjW3QmMDhIpuVymTjiTdxvEP8qsMZwXY61FhnHEk+QJwxcaIwMCtwR5sNwGMEPmu5IbaRiPhDLiJMAN56YcRF9IEcZKsRyEEPEAk8i4V41SNxxXK+NR1KD5F28ReLiCsUhufl8LnsvYVo7SJcI3NgvpSSFkVgWNDzJnOgCCcGt0SfEcOHKcYMj0VSKCXJJ3eoHNSs990GSq4dBiWDiclqoceJmSkopoc5yzqvVCs8XmcCFWytWdvgYkWOBw11X7GkL/CLO+vsOAAtvNg0yWYU5FNsqIQUsNHGyGMhMUZYC2wXVEceJh5ttrcV50BzfFQAXORdCjLL0dBMjhYivYd4yDgIDILSgWctBo2qcFsNMJi4u6A3qBEqdiMDYDY0sth4YIDNZCjsDZCnvhStAjkEfdl33qruUUppOp3hP8ASiKvBM2InM9WKem3x33LTRcw8KiF0adBUVECdchewfYGzB+dHD66P6InABCv6LYiDAVdF/wExDz0sYYURnpWWT5dAEHlgSRYRWcNyTz3POhhwe8fKEDs6cuqo4c08uPkQhapAeEzlIArXRNM1kMrGcuCMaAguVuOWosFx5EKoT+BEG2aT4JMBQ5PZk2CxcClZbcXQBmgIH23ELLOIAA/YOuE1zoQuUHL5lsRMQ/8ViAf5DBJyYGASUEcMkDgsNfCg1SN/PXLOsuGIVj4WwcBikLIqcSR6FLJPITeDGUDATnttu4gGkqiFyowLDfYBFyQnaENIscnST2PMX/tdwMWji7mbQDZAVATeisSQfQaxJHFSmCqdAHAPG+cZOYMHhWJhBhwlZRkC07xDHzEE9yEQYtBMh/hFZhB4F7sFdFMdlhXaHrRBLTdxBUSQMdly2WBSzOnPmjOIEESy9KPyCc7Dxi2RT+0EeAlan73tIodR80Wvrf9UgOVwcSBokIYGuSEwJDoFqP+gXGTgGEjmhyXGytgAIyCJ0nh+U7hMT6DiO4ro7bswLHSCazHFuv6Q0WE5Vlr2MzOvgmAp1ZDhhqOQKZcWdGy3zmZnHIVjOqMRbi2NlOQKDF8zcNFg0kBAcissUgaw9B7BDCNPpFCdc0utEgCJXl+LWhst5xfXGmRERNNyZqefE7RijOnPmDHFGgXNusVjAeCFbO3FgFT+WmyjCIkSuiBPtLYrOcyaQwGqxU8RsrIAYx638iAlW5KQiWRbnW3HppJhIEaDINSGeM+X0IAlODcpDsbtD17fg/qE0aL8/VOND0C3oFQcGWwvH4hgQcPN5oChsquFAaeaAieDZwGmj8u7yOonbXku8TxwFqW5Aci0ERVhEw/3d8RiaYz5DKFYMEkrF9RH3fpjzrwcJrIJnRCRSSsd0luOOqgUXTqPkD1sFhxDvLI6inGnBrWJxhOYP3E8Ozak8Z6LJoRePXSAnjmnH4xCsteJJ4vSkQV6wxF9FLxpOBpL3z1wOB3fBcQGJ4nYFUFFQ48UgNTtybEs0R+IGl7J8hhthZQ7zZU5RErObUsJeWu6cjAw7YRx6rozLOcPAeS5UFdRFXNogKs0O+pwYzomG+x25HRbkWCy7YICh+tGcfYYgHSCN50CvgFd5AMWVQ1ii49EkivlvMeGg3bTWGJ8DSfeDHGRoo4q7ygvsgBRm5uigA2ChYKocd/ERAcfjYpOghAO3pcGht9bi4gLuDHcwEWOET4JQQH6q5ppoUfuJh7aIahRUWwwqR/GDPxZcXp04l4OY4Ibml8iGeEyC7dSgRYll4lFWDI+NKI1AH2wn6P7ElfyRO3hC8ohdTtG1ctPVaiUuOgRLEi8N9yyE14ZEZJxePIOEMoQzE8bOc9UKce4itsY592p7gOVy6blPambXHy+gOQgPRT3EWfL70Ezg+WjQAsZw+pJgtzyg5MVECplhuKYEW24GSYb4ACgigXVyZaHFoFADpxUA34iNECEAl40XqQaDWsTMW+60QpzNCO9GmupBAwkKluCJXJOYNQg8toY4Bzdy4V8edMMU3qTl+SaCpXB3P2hiJNU8SPUkrrpMHI6V451Sgvx5jj8SRyQEvyMgAX0pES3RvoId9aAnznHSiQQoIJWQQUGySqmmaaBC5MyJKcHB9dyDN3JZuB/ku2CJYddbbp0ZBrWbWBEEC3E7bJ7nnnP4OrSXpPELzs2DRhyOu0HgwzipIFHwg+fEM4Ow8JwkhcWSosS+72ERoALFGSFmMoUaKDgPmLh/oFBlwgLkQbJbHNQBQ5FHDlQHLp8aJi5GTsCSt4Y+g+Dij9C7mTP24VRGTsAQh0B+xwp4TkEx3A1dFFvJrXNELmUclpKMiciZeOIFZM6G0ZwkGrkiLnA2jB4Uqol6qAajKsRdEjWO++Hv4/EYVUhCwkau6hIkCC1dDPJfFRfvJU7wED5XDaKAPdeoCxSQB+i534ioJYlsJO5Wqzi0In8XbkLYDdFSgi0gGXDT4JmLnpA1xGcENWeuCEuDBmpiJsThEMYV0q+YB4c1GOKSyHkdPVcia6a/JcFj6JNrzrS03DsAH3Nc/xUGLRmH/rhIW0Y+UOQUJMPVdI57ewms1ty/CGdxf38fdxqavJKbgAo7hxjkELR6Lq2CvCYO4mjm+2G/FM9LFD8lDlryimeHpRS4J3pIroz9EFgm6n14sISHxLrjwYaUXWbvCV+nQdqae23Rqhi+zO5xSgm6CiYGOAaaz3DjFVwTtXKZc9rFkRSQQBy1MIOiShFKnCWUHvecW63/s1qcwBOcBPviCoHTvMT1MUy1i8edeHYKImiolDpGdtKuS26J5RC4J+IvmAsyOwzNYLMRXsEWSpBcSIjEjfvQhFsxYzbMxCBOy5Kwg3A5CF6KRhEkTlwnFDg3StxDMNriowm8pUGUW9wxGKOea/sNVyJ77jEVmdkS3CB1SIH72Hmu2BKtmQbhdOLyvDzItYB2H7pRAgpFiIUZT1xDYrmgIrOXJDsYBk1RFSfiBe7HhUUGJoncDkuwswCvMGiWLQ8jHz7Gx9CW0vNW3G/EveGt4K5YRKmBEnUSucYPhgYBSzkZiXOZM7ecGVocCUQbjqvTa38SpzzjmoipiSinQd40ju9kMpE0Ma21FPCbwTxAxw1vDXe8FxUNScJcC/yrTEpwnPAlOyeZWYaLKCxP6hyiCohI4lTuxLFM0SuBk17ExJhBb1NsgcAsN0jwGHoDQoXjIdfX1yGOWA1gcMu93uTZEpffyLqJjcYzQBHgYYqikKF6Yh+Pz0TPWZgtNw42zKWKFaOBHxE40Vo2jwYprYqz0Ild5ZInoin2tIHKJZxuuZlBHHTzsxz19dy3QBiantsJOk4FEaISBwg42nKzVc0DwgK37xCHSDQzvug594W4zFlOuR50peh5FqTYNVgEMRDwoYBUAlfXa2ZZPSejqUEpo+MxdcMTj5pdNWi2jwWsuKt1ZobTDnIs5/O5LD70lsQigEHFMQRy7fse01HsoJeSHUzrJY7E4VKigO0wPRQKWQAHHB9Zo+HqQ5BhhsSgaG4ZIbpU8K/l7GbiGhqBL4mbxWgebAbPSDwCsURm0KhLdLgoznLQxdFwLqbinDKBRCL9ifMwh97NcrksuT275RgkLghxtAOGF5+RwDPeBV9H2Cdwnqh4ywIPhlpWZAUmeMhRBa5FFMwHY6cHNVxiLqA+46BVvLilmWNneE4slNQ1aM5hsoP205bLIyXpDHIGIEE8nvC4tGAYWBFVJL/LHgvy8BynlSRigWDYIcAOYbH8oECn574CiSs91KBaNA26XxM7emLjEnfDQCjecJRAKvcUM93QQEhIiDzAWzCjGtQVCBRD5EQw2Wg0QoqZ4Q4Y+K7oOYE1ehBFiZzYShx601w8KXIWOZQhEmA5SiWKYXjxwP0JRRmYwSw9URiyRHpQAZiY/jWDAQmJCVX8E6J42ALD5XuidIUWlr9oDhvknF/TeU8eC79bzmJEEjWAd+BpeELTRY47gmwEyQ3BwtfLwdji4RGxPFlXs8Mv1l3MiuaRbIlJIEGXOJo4THYQilccah7KN/Sl4lIVx8kqQywyBDeGeVsignqT7VScgxyYAY/cvk2SrUSpCD0GCh5pfordN+LMdjUYg+S5MS2xf+Ccm0wm0+lUlFweND+BXtTMiQgM0pzTEnloteLGWS33NYcICgBKXCyGUwc8k7ibheKklG44cWt7e9sMmvcGbtUTY5SepuI7ZG4cYQfzFoXC0ezJW+6VmZj2EHQmqk7cnMADTeXkJU69oEHcW5AEDbowicGSlwREII6c91wwpDikKr664dxILCisgwiE5pCCACZRqJPJxL+2mISYSCNm5CM3Xex55KAdpGuJfaEBsyJkmxkkHIoNErpITB5iO2KYInffokEmRuZ0CZH1uq4RWxRoLISF4agizlvkQrPACWhYPcixID8LXgcgQPwp4qREPL24f8PzJ3KdmUIU6IBfkKAksg+7FgbZZJ4Ta3B3kQMxCo47peMWKCdAFh9UiwAvUUKaWXYoW0Rw8agdj+mQw5C4sJyYCsc7Og4fEmcHRB5LEAfFMWJncU6k8lBUaeA24bgUDq4UT8qeCX7KPL8BdR2ZKdMhylSDeLPmQgBcX06L4V6+wicprqoWTa+4F0/BLS4cxyjjIOUNCC/wOABi108CCcfRMuEJSm6FJPYSyMAMiqeAHId/idzh0Q8SXiWwL95v5goB0XYppel0IrmFasB90WAGg7BEIs3gC4ZxvjCoR5Hz0XNrFVHyecDAQpOL+03SOJLVldgCfFigmOgSwTSO27UkLkmQ7CLR0OLYEoc2hx6uYDUxT4Yja2nQbxX+lxx4y3OM1IAvxvJCCgVrq8EYVCDaxOkPMB3D98VSr1Yryf4mbpUJ2BR4doLFmvrBkMTMDRwgJYYLBlB1S+y/CJgYikXkhgSATYFJP1gWGQ+buROAtS6laK1Kg8QjNehBRtyzEtoS1KXi6tI8aH3vB13DNA8KoYFrKXKTmfNE3icxBSenSHOcv+ABF5HLCwWs4OKCBsIgkJK50KLnLgjS1EZkGs9f8RytnisDBf9VPEWEmCKynEYHUKG5KNZxPx0hmoVq6bmFT9u20+kUhyTzPOQ0yBPquKtaCAFjcQyXkFvuliQGKnO2D4L/xy3DOm46EQZJnHICSh4SK7hEc7jHce8PMFcYAdH3fdPUohsAloUI4LNlcqYYs/epKEp8Nx3Pey+UUkVZwFJIcCAxh2kGVFvmnEasQsk/kBWJtooXM/RoCu4gTkSYGCSrn7kfiKBU/C/WB9npgoEsly5J+ofibvOShESDcpEYI9qY9H0PpgfSAI8SMmc4sCCVoIiw4vrGGHg2cswESiouaMSHIY5QeNAihouoAqevaC6P0YOycagDqTxU7NWLji+KQp06dUqcNFmRwEU2/aBKLXLPTRg4Yb6JlEy9wONaZxWB6sghvOp2Gg5tGmO0ViHkonCMP46PHREppUPwmbIixajomJuXDAfZOdG3wtlo5nkFk2G5sU/Qo9JdQKJIgkhE28viyu0Es8MxlDRIw90asDJi6LHBYo6JM1sE06RBBFdxSrGsVR7UBNJg9o/AbWS3EU9cUNw/Wg0Sjvu+R2VBYHISABzelmF+Kw8IGqUUOl7g7KHGRnGJnxAZwJRa3C5xd7XWKIyHrGClxBKL/gjce7ZpGq2sUrZwle9TjNS3QSlH2QSfUyTKJgbKSXdtCD53XVgum66LxujIIz/7Phhd5KRjyF3bUzbR5xQpRaX1qynuxMy9ZtbHvHaotLCuiTs3OE7QCUzryQES/1lWVnGAQuKvgYt47GCcOy5VcIkcfhJ3BBQh6AadJyTuITBoqK5KbrnqOGEmMX+rOSVZsDlWA1/JPBFQcxYbVBH+gifJgzSbyE1CoZ7FqmaO/moOXwjhIkhXD0b1yDMfpzN+h/mEtsT6oujEGDMajcqycJysZDirVymltY0hxZitdUabnFVTNyFEpQyR6r3ve991fc4ADUqRTjEtlwvUbxtOOQ0heh+NQX2j1cYaY70/TlEV6qyqSnF9Myfzwiobznk1nCGJ/UNIFdAE50kaTcJCVdzQRNQwyKfJZLK+vg6V471fLpfiMIcQEJuTLIih2AXuyaJ4TkhKCVWRiK8VXE8XuYehtMtVHMuEhAkcFs9GoLTnZE7iJDVsPw0izQJzBTD1XPEtgRHoMFhV8HY9N9cCS1RyL1E5dceHFn2i83HqkzJGi0Ky9rhhG1uEpNQxB5NS8v7VRMGctFI6JcpZGSKi7HOXTSSlc29MJo11KWzIOZPSxhKp1PdKK2uMVhRzKkcj7/uEnOJMpTs2WMbqnIPnHAM9KPBIKfV9V3AhBM6HNBAmzonOgy4zkasOxCdQA9IF2rvlBqX4LyBd4hi4uM2WI2W4+9AzhWbd2NhAbKQoCpRMCSmceQC51q+OK+l5GKDl9iaKk5YEpAOt4iEzVwAm/in/sznilrMnFDMX4t+JMsvMhoO4yszLYB0Mt2ATvSVy3Pf9qw2mcs7WGkRpOHrilTq22bx64F28MabrWtkSpXSMWZEhMkQqU0zkSYWYYu6NI6VJqcJSzpRzzCkpIkWklbU2+N5YY22plAkxGWsUkclKc5gw5RhToEzFcbRLG61DiGVZGKO9Px54CMxoOFcmMHuJZL88qGaHTwTuwHIlHqQwc6dYuHVQwKKrVqtVZBo9c1Z44h7TmkmyzGlJeCTRgkISDiGaMDff0aWw5K4rlsOIAFXGGGno7DmZUMyrGwzH8TzETeJIkC3Lgf1hrqb0mNNcMgt9iXkajtP3gA2GB1JtbGw47g5Z1yvNDYtyzjGmAUDTGO2jtSJSKaWcstIqp5xyjjHknLUyOedodIhRx1CSfvih+zY2pxTSzdu7r9y62dXRkY0UM2VlslJZaU3KGONizFYbhCiJiOKrQ8GVVtY5SG1KyRhluJeW4USF4AMRucI4TlGSpXTcUl6QihnMM1Tc3jtwZUw/KEs1zMhJrArCFAdTMqHtsbWac25Egokz3vVgvofi2eTEzdSFQXDc4ll4NcMRxsR9/lFxhdeBJMkw2sz5aJlJapEYP5hhOkRs4C3FtAkSEs8rcfwxM/kCvvQ44XVraytzE6sQpKFdzjmFcNwz0ByX5hyngHRda0xhtEUtPBFprWLygDEp50lVntla/+mP/ehP/bGPVKUvbbla9s+/cvnr33xuZ+dod+/g8Ghx/ebNZrVcLOq6D62PStvCGqV1ilEbTaTUoCjHDHrHoJOONUYdB1kLOjZS2hXa+9b7UBTOOZfScZHe0GyrwYgMGoQVBVFmJt9KbliG3co8XwKuKFS656I2bGfgXg7toLk9LCDsUeYohLC97WAEgrASSEgSgyJ2CgInrrFYloLb+8s5MYP8TKgTOFxu0DI8csONwNkm+AuUNM4PwsmMcI6zWaRLVd/36s477+y6XikKIRIpa4oYQyYyWvfhuC9Y24CxzVrrTDmGqJRJMWltGBnkkHqr8xvf9Prve8973//ed25vj+67/6TVXWwP4qp32eWoSce2773XTaeyHu2tjg4Wq/1l8+Qzz/3u733ulVeut12vtXPOZUoEzytnpTShU07OOWVSSWvlg1daF85pZXOKROht5a3T3sMbMIY7KIp8WO51T8xqALgInoVjjKMvCf920JoU2wMVgrq+zNy3OFaWczwcjxfCriwWCxp0/rfWwocASIIVzpx0CwM3TBOzXAgGiQmDHCPFyQWQP4HboCKrqpJieDEvjhMWaJDZMhQXuZREFBLXoxFHokII6uTJ04ZbtYWQjLY555yyD96nxjqXovJ9UKS0ysaalFKKMUUER1Qiyoli8I8+8sD/8Zf/3I/84PfNxmUMvTaZyIfY5W7hV0dmedQf7q3awy5kVzjjSk1bSftqvaDRKOhp26knv/HM737mK5/7/FOXLt6IKpLWWtusDCnKKhjKFJRTlXLKxy7lrKzSShnSRMpqp5UOIRRlQaS00kqT0lFQjuJ2TJAVzfFRxU2ceu7roLVGx3vFzYqNMcOx7YK7MwcTAI+wqRUPu4SiglUSExYGGUKGo7aCZCUOr7gSQfMo1szTviJ3UJB9hWqBpMK1Fs4CSk5rjRFv+F9RmYD2msNKwuJKBtVQaBT3ZhF9DICvtrZOCsFgjWvb3ljje6+U6kOttNJkc1ZEKsVImYiyUirmkFJUikLwW+trf/oXfvYX/sQv3nXvPYnIqJyi732rFFWVa5YHcXVIq8P+aKdv5n0Tlc5lYZMfd6vd8Swn8pRHbu0kTafBus6rz/7hl//T733txVeuXt892j1aZspGJWUoZq1NZVWg5A0Zo1RIFIwryiJGr40ujAu9H4/HwQfrTNstJYkxDDrqS9g1cVGwkECRY/twScC2FUWBVr0Ai+KrR47awtYkDs4DnksKSsG5yWIiBR5BOIRrxt89t5by3o/HY3TGEMiSufWg+IluMOFAwLJ4+46nqCRuwgf4IsQjboQwu3hkmnOMNJen0SBtUiJUOWd14sQ2HPgQorUuBSJFIQatdfA9EZHKRJRj8qEnIhDEmUhrtTYbv+vtb/u5n/7YD77/XeVsLdkymYJypJS0NjGElFOKjY6djfXq4KZfHPTL1CwXoT/IwU3Wxj4cUe7Xqk1djc3MtNTlpGfFiW7pvvyVbxyuuqs3d1+6fOPlS9duH+3OO3+06kPMwSeri5SS0Y4MKaNJExllstJZGWucdTGGTH6INNNgwjL+ImPbhRqx3KOO+AdUkLTbsVyFTZyPoQaZZVhTgSCJm2MIRSnrLtBePKa1tbWWp0iLCymjVYVrEGUmBO9QuQoGkniUHox4G7pjiuNCPU+OwimC/R06dGqQLoK7yPVzzmp9fZOIlKKcKQTgH6W1CiFoMjnHLjSaaGN9+sbHH9vc3Dx37uz29vZsbf3U9ok7Tm4++vAD08olCtlYN6p8IutcDNHqIme1XK6UDtGvKuNDfbRz65pNa6lrk78VOp9TaPujnFPpTqxtbNlZjtpPy83YVM6Z5XLhu3Z/77ZWpus6VWhXTW7eXrzy4rXzF17ZXzaHq27v4Gj/cH57/6hP2lVjZ4wiUlo565SmGHu4hxI7E0ANkw0lYYyZTCYIh8lnBLskzkUXZC1TtAFHEHqU3eq564P3fmNjI6Uk/f9F2gxne2Xu5UBclJgGg3lFkWhuSCK2GE8eY0RDKkgeUyrHilYEepikKzgaOJ04EkycznD8XVKuOGaMhbaQsq3M9RHWWjWZnkgxGK2MVm3XEpFWOqSUYqYUN2aTd737HT/0wx9621vf/MijD42qsTVGqRxyn/pO+76t5027ijHmPqYcynFZFqVWuu68crOmi8ZknbvV0c3YLXwfU5w4U46qdr53tTm4EWKndFEVG5P1rXJtrEs3Hp3UtF5Tb0rVzg+07/xi2c33NXWhqy3ZGJrJ+joV1WgyaYPab8vPffWpf/avfuf8hdvK5BD7ohhroyMFS0kpSjkWRVEWVd00WpvJeFw3tVIKwVgcO8OhWc+dShOHyRTHlSAZa2tr2DzQB5JtJ6SLCKIkDitOmVAcqyLOd5MQLKQT3boCl2NLvoPk+/rXVmzm1yakZ65fiJxeAuIHJLJE3AxXEYHuClwIC9Yn59y0jTFaK+190No4dwwiOTvDSgBAa61Gk7EiSjnlnFJUOaucYkrhwfvu//hPfPgD73/fO9/19ulk1vo2UUiBNOWc2z4tle+oqfuu7rpFvzpa3Dpq6s6O7ZnTd2iiRbMsJxtZjavSzQ/3FvP9FP1sbbOYnDTlZghtCgfN3rWuaYxS49HIR6rW1t14Ws5OjddOG2ebbplDb/rYLhpb6uXBFZub5cHBrZ0be4dH48nk9Mbs1Ik155xS1XMXF7/yt/4/37xyzRQjla2ymVQsbRlCp02uqiJFFSMprcFixnhcGsw06XHwUvI3BGMW3BktxihAVToiEJfeeW4JDTDkeAyD5kIDx91hAeGlX4XwLiGEzc3NxWIB8rMdDJWyg+GHNJhNrnm6CMjPljt9D9UtriOVcZpzCsSaE7culXio970rrCIVAvWdL0ob43EQtyyrnFSm7Kzr+i7nfJw44XMmZWJUFPPd58784id+8o999Mfuuf9UUYxDTzmrRH2m3mQTu1pR27QHpmtVvfD1UQirnZ2rh/vd3t6iHFXVeLq+PiudbleLwjirXRvSwXK1eeqUna7p8dZk887pbFOneHTris79we5Vik3SplNm8+Sp9VP3GLdWahtjVlpR36iY6noZ8sKVmYjaVR+zrZvmxuWXw+r2CdufMqNitP3p567+9X/4b67trZxzpUmKYiCntNKarNMxHvOcxwAlBeFUxEsXn8VwCjAyiMVbFrMSBwWN8kegnI77f9vXDlsRDlrKds0goRG3Jp44A+ycBxWxEq6hQR/7yGm1Qh4KV+k4eTJwJlrOOYbkiiLFREppTTm/mocfOJtM4v8p5RgoE+Et2f8nRceEZ0xJEVmKVQhBU1aUTm2vff+H3vdn/vSfevzRh1X2Oi26o6Vz05z71Wo/U1Jd0KFpljt9s8p1E+u6qRfL1WK56q7uHN7Y2Q8h7u3tTsfT9cnY6rA2KbtVu751qg7p0Wp2duucqkZrG9NqPPVdspMTlNrRWlfPb0ffrM1GzfxobbM1buq7VtkxaaeKkJo6d14pIqWma+tuFI5W9Wy2uXHy8dz3JpBu5reu37KlfdujDx5+6akuhdYHR6pToSitUrpre2NcpMiRTm/NcbjUDbquAqYMsw+kftJwMRC2xHD3gqF6KLgGXHFvIT7QXogoCU103FVNfDc7mAMJFipwRhjUoZDX4m1BXsUJyNxHRpCKmOCUorWOMvmey2dDX1UuhNjyFIBBalF0rshZpZwVUU7KGBMzxUA5p5wTETVtp7U2WltbqbMbW/ef2/6edz7x7ve8+bHHHxqNJyruhxBMrP0qezpcNbd8WGkqVfBxua9zTW1/+9rOzs29g1VzY+fgYD8e9n5/uTA6VkatlquLzbzva2uCI5PV5TvuOrd+cu+OOx9eq8bZrxZzX1azcr2yuixcmoyr/euvhKM9U1b7l1+5874RudLnkHzSus8u0NR6nyIlF3w1mRXVhtbG10dtWObKNVU6vTneuX1wyul7Tm2dv3FbKRuV1kqliNOTKQfrLOPT44QEx/2Kw6BCCByr6Bjg0yFtAz8fQSK4MMVgUnMa5ANJtqj0HhX/K3K5tzjGCJH23MtGHPvAvV08d5KAlYTdAUmIglHoy3Q8HI6sNdYiTSqORlXTtM6NUko5H7M4IYSuOy4xEL8ML9L3nrJSysSUrXm1FoqIYkxaa2ucIoox2f/x7/61xx+47+RIr01ylxbOHMTl4WKps7HtXm2VDn2dUts3K5Xo5s3rzXK+ODhQfX+0d3Tl+s5+k28uOlWNN8ZrJyaTslLTsTXkmiY4vX24f2vv6Kicbr50bfflW78/Pn3nW7a2inZJpvTGWae6xpMeaR2KYr3zQeWwnN+8tTM6dc/DI6N9H1sfnC5UbttmT+ucUyBtytFW5WauyspHT0apUXt4+8GH7/nROJueevbyb/+bZQhRW5coRqVMWZgyhTbmrIxNpExRUArG2JwTwqxSSqG4CxhxvVXgSho0LRim5TseeSzBEKkilbhB4lIKccQS11CLIWt54AH4a3HNEoeH8+CHcxCO3enVakWk6rodj5XW1hjTNHXO1HtPpEJI1lSZcr3qY6LgG1IUu6iUShSoPx5tqY5ziDUbU6IcUiSlApEKIVPOpE3OOueoKOVMOamYMyVt3/i68UPnbF7Nm8VR1NrSONT90XxZTad+2d7e20uhKZxK3oema/b39w4O6rqxybbe7S/zytjxyXWtRxNdzKZm80S5MXWhC4eHi5Tj9ua5O7rtW4d+7+bRjZ35Z7/29H2ve3h7NHaF7fomNtHYKrnSbti1rFrnfLffdfVyf+7Gu5NZ0lkXhmxsVwfXtT9aX5/p3NZH+7ZYr2OqJlvOlUXu+y6oZt6ow0cerc7c+8bprP2jL5y/dn1vbfvEwXx563DeRNLaqawppGJUGspG25woJtJKxxSWyyUOsUSapPEAxgvFQTpw4NY7kmiheayRAFuJnxA395CUDKFhHNcZSv258Ahu0OKIOI2auOoP3nvgDENjjoc+4+FjyIqM1Sr4RKSEbdJaxxBiSETZGJvp2BixYjM5USatSPk+pOyJVCYZeq+0NkopRTnlyMyCW9+Y2u7odjcrqO51KIrxtF3VftFUyZg+5L717eLqpVfuufuOvqkL0mOtzPrmLVIH+33Tu2p2dt7Nl91yNkrTtZHxbd3M9xYx9opi1hUtDo+KaNtFO65K5cLnv/Lkm9/86Pd891v8ahGzc1kX5dQrm1Uyaxu5XqV+GfqVyl1zNK/rvigc9fVq56qLdUitaboY8+z0vXs7V9SkPnn2bu1GKXSk3drGJtWL/ZvXCm0++sG3vP+N9/2H3/vMzaNU3nf6sA4X9w9euba3WHWKcuyJjOn6WFWjnMgUNuXj+LwQ0MYYZAsJ/YPdwrYFnkISeFSoaBrhSwyXaClONBaEkblMJXHJn2RtE+f+DhMIE9fvonsETI8gHiA5JG12XZczaWX6/hgyGs0ljgqCokmlnFOIiSgRvcp5qphyppD6ylWPPfy6Rx556LHHHp1MR66w5XSSUhiVRVXYsihyTm3bFEW5sblx7o7Ttl0kQ0WkEJJr6xxT7lvviFb1MmYytnjk0cfKQh+mUGmyqVzsHy36ZhFj552z40nsrQ5n7tgcmdHiyPdeHXax2V/dMV7bnE3cqPB1f+9odnXniM6sv3ht98uf//LjD9558uR2itnOpmSMSpqSDtqO7jhJtxdFWPk+7O3suLW18bj0e7ee/uLnHSlF2hauKst7HqU8qtZOt91enk6mMbWaCjs6uXHWjdc3qGsXe3uqVL/wZz7+za9++4VvvHC6qu4/d//O/Xc/+8Kl5y9ebvvgdaGS6ntMHYhKkbFG6+Mmh5qHJgvCNdzPL3IqludxJPDw/aAfA75bDPrDGW6DL2Rd4IxYoYUU90xhCNIPe0Ji9oXl1pYW3UVeLeKBzIWUk9EuE6GDknM2xJhfo9uOAxci3KRUTkkpGhXF6VOn3v/+9//0T/30G9/4phMbG0WpiXzKPmnn21onb7LPOWZSOSdjXdc2Wkc7Ho9j6uturtSMbBE7XYxH3fzI900yRTGZhNT7VduHPquwOqhvXL+9Umrpc12vqMsba6WbTrVan6/SxZv1QbPcXCsnthqb0aab0Sj7KnXzeamom+/7ZnXl8o1xNV4cHFg3qnOyyRXVWmFsDDloMutrtDpaLg+rUdEtj3Qs+7q5cn3n0ovXNzdPtvVye3PywovPv+197zqxuWZWN29cPto4e7euNtpkciyLNZd8PrP9xqbJ0fv3vHftXW9+4oVvX3z++ee3S/Po9z7xytV7PvXlp64uajI2pmjJhr43TsU+ukIhh7jrXp3hTVxwCOQR47HPEmOcz+eKB8ECDgO9SvMlMLrGGPRWiBEgA1cDtDoObDnnQvDwfVar496PwhQLdkZemlIqhRC9DzFqYxSplFNOlCkTqZhSzh0pIlKSO5DB7AVPFLe2Ttxz512Pv+Gxhx56eDQez2Zryhhn9bnTJ+67/947zp0ti8J3XQjzvk+kQ8pdTppCsMl3zSrFmCgnCpG8VtnkaDW51aKxxilj2q61RCEFHztNSZl+2bU600ipkXVHR0dH877vi3nX3jpa6ahOn9ysqrL16eoLNw6WB+WsHI/15qQ8ORrPysqXKqRw89bezs2DeRNXuUjFaHrqTLl+8uj2pXZ5W8/tTDlXTVdH+65w5EZqvLV959i6vRi7xcFeXoW16WTnsH5pZzHr3YmNcX2w+8S9b379W94TVP7spz+/ODj4gZ/Ymm1sdX1n8yru7bnJmi+smikd9R0n39RTP7rnkbseurvd37k9byexX3v32/75F75ya94oW/kciULy2TiXs0pEvustJxnSIKtGwhRDRCyMi+PKZc+9pEI4rnlt266qRiH4lHLTYBR1qRTFmCBhcIWKokgpt22nFEgjG0JvzHGst2mavpfewiGnpBVlUt4nYwxlUJEqpRxiSwS5UURakTJaEcXZbPrAA4/8+Ec+/JEP/9jdd945nUyVClEhPUanmEzKwVAXfWjq2DcqkzFVjmrVtCplq3QMvlu2VpsUo3bBueS7ZfTRzpfNibWRsSpFbw11dZP6rvdeUbShL2KrM+3c3Iu9r1f9URNrbV/ZuT333anNLTOZ3J4vr1zfMWQ2t2ZnTm2MR3qkirEt6naxM19cuHbj1lHXdWnZJzuZxbH71Oe/+sGvfOOjH3rX4tbLi6PF8uBa3zbrG6dTyGRM52k82j55dlvltqquzXeumLh8/MG7L1zY2dk/OHvfHY8//ubv+cCHfv8L3/q93/uD808//xf/wi8t+7SpFeWwmh9c/OYLR21694/+uFdJj/Sir914vHHu7Imts/WN81sHt+69786vf/n5m/Vj//7zz8y7rEvtjNYpGmZWJKNX/PbEOYeJm6lLvhXMARLOpWQspWStK8sCTOCwmI7tSBzmS0hegB0UucYYYky+D0TUNJ33rzaZVEpp60LKOcacU4weWNiS2do68eCDDz/40ENnz5xZn01TiHV9OJ6Mtre2H3r4/kdfd//6yGpbJUPLvtaacuopUww9ZaLoVFkGIpOtUpMQ+rZtnCUVVEw+BN+ulrHvrFKTqko+pq432i5WtZ13czc+l7tek27ale9WwfeuKtq20Ymo6w4ODo+OuqZJXRcP6nRl76jtU1mN1ja2ru3cvrGzl60+dWLtjlNba+Myh35vf3V5cXNeH6Zi1KjZ5NTZ6xcv7a8W7e1FCMl3/h/9+r/58IfeP5lsUEz7u/u+bzsVq+lmjLmoHBmlXTEqxtVkdOLkul4ePHDnjdefufjUK9cPdhbPPH3xt3/nry8PV5OybJru//J/+wfj2dq73/WOR+6/6+bVi0899+yzL177xSvmEx//8ZMbzoyjiqXKI+UKNzu9NrbbsX7uqScf2LJvvv+OL71wLZJR2SiKPgQiqsqSMliiIA6X1HYZ/hE7JXkaqLXIXAGD7ddcpd+/OpDrmEgUB03yOoSfzK9O6aMQKEYE/A30Ss6ZMnnvU07rs/Hprc3JuLrv/gde/4bH3/72tz3yyCOnT58ej0YqB8qhb+uDvdujclqUE2PjYn79YLUsJyfK6RaZiqIlbyn2NpGm2Kemj4psqZWjSN1qHv0i6xS9L7ReHO13bW2NTjnNj267wo6no7ZrfUrWzRJZ5dKkr2No532zVFpP1taVNr5dNl13OF8se7u/9Itls3fUXbmxO12bFMXk2sVbi/lhcPm+e++bTSb7R4fPvbivlR2P17vWF6PpqjerTr104cWdw/1IpKNyORnrvv71Z/7j73z24z/wruXhfGR09Iv6yCc/N2VhzKycVH2fY1c2Ko82Tlezjbd+t372/IVXls35SzeaV650dX1uc93k9tyJSXD6wrXrr/zL37rvvrsPjuZ7i5VP6m/+3f/Xv/7t3/7RD733fe9/z+kzZ7ZPj7KLG5t3uFjE5cW77t3a3dt99Nzoucu020etjcq6j95a26HTBWmtVVWVOb86flUyddKg24FktUogkxP8sjVWs/kTriVnSimSUlqpEINSmjhz+TgNMgTiIGVKpIiO9Q6kLYaUgnPlO9/x1h/4we/74He/5+6zp5xKm2fuqmbrOYSYUt/WfXuQs0+hzslPKt0sa5PKqJrkvbajmHLXNsak4Hun0nxvp6/nhmIudHJl1mXsaDbZqA9u5XDQU1jsHoaj2hhTVFaXhW+btdl0vmjIrfU52GKmfuuf/NJbHnpTOApGFW09p36uqC1cdXuv3pvvrOqD3d2Dti1297quy7uHS1uWWYcYlTXmaL578vQJN3X73apZmmaRfNtNppWyFMhevLZ35dq+90pplSmqnDPFnHL0/qH7Tv/6P/n7d59aiwfXLj3/9O6ief2b33HtyuVV2zzxXR+Ibhb9xFZTpTx1t0f+8Ouf/vz/8hv/6WvPX1XT8QP33n1460ZpSJtAWt/er9vWN96HFPqsUjkuRtUot5PUq5DO3nGqmp1oyf7vf+lPPvH2bWNDcaRvPv+tL3/tyX/+b7/x6cuLbLQJMR9Hu5S1NnRtVSK3yxLlGI9pHsQjhdBLKRmrg48p5aKolEL1eMqZUkxZJa0UUY4xxUyKrCKjUoqp14ZSBuNrFWlFUamoiHQ2UaUYg1I6Z02kiTJlUpRUDj7Frc2td737Xb/wJ/7493/o+ybjcW5X3eogk7fFiHTlY8o5hn7lrMqUu3ae2nlOrl3l6XgjZO9G1hgbVNKa0qpJXbO4fevWxZc7vzJTV7jixPbprMqcTO6Dbxf7e9eir40mTXpjc00rpclVoypqyrZQVbVsvValnYzWrLNR152vp1OnGt3V3dHu4eFuM66K3aOuqdO8rld1W9ddUVaB+qRC3zduPLnn7jsV6ZdfvtxbnWl0c29PO3206DY2T96+Pb98dZdUYZ2lHFMmYylkij5aW1x4+dpf+G/+1v/z7/3tM5Nz5+6Oz/3+p543z2+fPn3+hZcfe7svp262vtF3WaWUjW5iclvlvXcWF2/G53f3uz4u5wsfuvHEzmblpBhvTqZqfx59mBTl7dUq5LR9+oSpV6e31kpjch9JuX/xP/+zs+Mfe/D1d8e1ja23f8/3nLlv50a/W3/lhZ066VHT9doYa4yiOB6N4Ch1XZdzQvms5HFK9FRrpSg7Z5UyKamu7TIMDKlMOScKKWutFJmUuhRjijSpqtnazEdvrW2b2hqjsq7rrg99JkOkM7UAOuPxVCvd1MuUaTKb3X/f/e965xM//ZM/+573vMcWuW0Xy8OdWK+MpaRDszrQVOSstFW+X/Z1b6vS6OxD2y0ODVVtHaqNbTcaL/du6IIodKu93aPbO/tXr+9d20laz2PrSqX0+ROb66Qoe0o+t+1SqWAsndjanjer0ag02irS2hVRGWdGPkatWksh9fUhUUMxrOZ9lfzh/uHhUeOT8fthfrvdP2q8tUknW6ou9tV03EXvsrKl7rzv2ly5zdV81esuqDzdmHbJPfXCxW7RW1slT0klUrlwJuVoyOSUSZGh6ee+8Myf+nN/4W//tf/DvSfGr3/zY1/57BdOVO88tTk5vH353HRMdmR0kSgUhdm/uafC8vvf87rKmPLbe18/f6nxKmozGU3XNtdU62dr1fp0dvHlGzmkicrku0qZYm3TV8VoMtrd39853O2S/6e/9ft/Yf0n3Jqe3Pmm6pFz3/PhvfMvvVT76cWjzlkajUbBe62Qmk5dhzqK4xTVlBI4GImrhxCITEpRqZQT5axyJmO0D53WmiIFz0NZdDx16uTP/vQnPvzDH9o8MelC7Go/PzyaTp2z8ZWLF5599vnrNw6O5l3TLe67/+E3Pf7EPffeNSqKmzeuefLnzt3zute9YXNzapXxfd00te8XKjShbnpFNHIh9TZ0ftWTyVl1o8ItDpaks44xNKvg59XaprNbOUVKtVp2O1df3rt5ZXl4mJpQWL17tDhqmqP2cDE/WJ8VW1szpyqjJ1lnZdRDjzxkjVNWlbNNratkqp6Msa6Pdm3jZNt21vra+nmhUuPp8KjpKN3e8wfLtLtYhLa7cfvIZ6uV6oO3rlBE5WhMkZIq6mBWR6u+S12bXDk2RTK+Cb2+dWs3tFEpHfkrpBSpTDmprK2xSfdRkzHlV558+mf/1C//F3/8w9/7XW989/f9yDPf+FZ1cvrNp8/feefDXXvQR1PYou8Xk6o4sbExufPOtbMPdvl3C1U/+eLt3WV44vEnjm5eTLlfzfezGU+3T8e+rgK5wq3q+lbTV+uzWd92fXvQLHubP/npb9954szHf/x9hb2s1orXve3hn/jwDxz8qy8t6OBgcZR80HQsDCm9muuuORn+vvvuu3btmkQhug6cEKg/ZW0VQvChC6FPKW7ONu588N7ZbO3cubNvfecT3/3d73vLG9+kVe/7Q6uyzmueXDI+0+otb3/sYykqb65evumduvfeNxTFNKXG+47iG5xtVIw59L7d76PqV3X2TTVKR4c7WhV2NKmDz1qVZHQydbvyuZmdKFRUMaoUVefTcm8n7d14cDomM1kdHnz7S5+f37w8m5jNjS2zub6zf0RFLLJeL0/NxhsUu6mbGqPqtrOuLKbjw6Y5MRt5UsHY1oeqHNlylmJ0Zdl3vXJTG9slhVk9X62O9O7u8trNW8vau/HG7UWo62UuRpqMj722LikznlZ12xwuj9rexmTrZVcUhapsNDkrZfT40ss3+z5qpa3SZI1zGnxHzilF0mRC9taa7H1S0RVr12/Xf/Pv//q//O3TjzzwwGxtre6X1pr9eu0nPvKB8Wyz7zpqO6Xo7GNvCNbed0/6xPbGnf/psye2X/69Lz6z2t/bnm56Y1pqr+0t6j4GX9+1PfLt/I477t5f+PliaVV4+J67193tg8V8ff2uv/2P/n2YmQ+9bT6l/cnW7L3vfeMzX3v25VvX9oJX2hiVfSJXOo0uaYrquj59+tQjjzzy7W9/+y1vectyuXzssceefvrpo6OjGPNkMlVKd10w2qqkso/jkTv34Ln3ve/dP/tTP/vAAw85V07GE1MUIfgc2pCWyS9vXb/SdNWdD73RJ6VokpJTsaEuba6dHJ9co5y7ZhVyTSnq1C0PLqtupZLtXGHcJLYh1PPF7aP5wf54bctkStW0bRsbrW57772tXLdsYtbFaLxqezLFrZs3d3euzRftfJH/6HOfmaTmbY/d+8j995JSB0ufTBFUnVIYj0+0aaWTjU0xmhXV5qSYTbJWIztS2qxtzGxRqZh6n7Mma8oUlFJG6WQPFvFoGfau7eYwOjjs9ubNUb3cu3ptbfNMSDqHsFgeqNJFTcbGpl/tHRzGpFof+7Yh0tr0axuT1lPd0PWbR8umdaZUynVdMxpXpFKKyRijSFlT+OALa3OmypU5pz50ylFM5sKlW8+9dNUoV5Y2Jv/J//jp3/63/+mn/vgvvPPdT9x/6qSmSaDg4rJrunsefPzHP3b2kcfOb25Oz1+4df2Iztxxglp9550br1zevX2wumN7/JZHX3fPXfd98ekXDheHj95/39RRe2IaY902u301/hv/j3994/tufPD1W295/N58//Txh8/9zleeKijETE6lqMpE2qDlW2Xf/o63Xb502Tl36tSp3d3dra0tlIOtr6/XdZNz7ttISsdIJrc/93N/7BM/94n7H7p3a+uE1i72DcU2+93oC/JN3+x4v1gsDr/26c/WbfzYyakrZjoFbZqk+6bPMVQqbpEzZElHqxVR6mNdX3nh6bZtAsWNk3enXF556bl+dfjAww8XI9sG77Iin2PuW7/sFSlTZmdCyM1qWWl148a1/Ss755+7+LufOd+vzGq1ePPr7yBXdK1vmkUdi95XXacLp8cuReWNjRTrbuU3pltt12xsbI6Vm4zK6H1o+0KNjdIq9pRiVFpZ29S9fe65i3tXbpg+jIqN67cPF6HOhpLKt2/vLOd16oNztjDVoltlrY7mXeeTj9HnUI0mVWVMGbpck9586eJLTedH40mhixSTD0qpTDkrRSmhODo7V4UYFakYo9JKKeUKZ1KKQTlj8QVrS+/7z3356U9/5VfOnj35Qx9450/9+Ife8fbX9X1d2Cr23fjc+hPrb9yo3O9//itfe+HqN8+/tLG+trEx2pxN9g4Wewedv2/6wqVbi1W/XNW3DxblqRMntjc6ahc7zQOTjQt192u//ft29d4zp9ceuXftwYfPvfXhe6+3V2/MfU8qGWvA/cegdfHzP/fzJ0+e/LVf+7W9vb1XXnnl5MmTV65cef/733/jxvXnn3/B+5hyUqS7rv2u973pb/7f/+radC31KfmUc52bncOrzz391c+dmG5MquposXdzf+/yjVu3Xr7ktPnG5tqorEwZo6oLU2gaF6NToxNbqixSNv2qK0cm6qIh98rVWy+df/pg/3D9xB3LZTj/7Le/6/3veODtb/d25Ir1etloq1ZNR9qV0ypqFcoZmX59Su3NK/PrO0e7zfMXb19ahUJVY6NuH3ZHK7V70O3OF0dtmq9COTXT9erU2vSu02sqp8XBQd8GnYI1abo5UapMbRzNxk3TZqurolSqH40niPSvbazZ6/uL0k1Np6/evH1QL8zELuer+VGjaKzNKKRG5yJRsaiPTFn5VLWhI5NCyofL1ZTUue2TSttvPHWp89mYomuaPjXT2Wwy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';
const OREXIS_LOGO_LIGHT = 'data:image/jpeg;base64,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';
/* Midas-branded versions of the two apps: same data, same logins, Midas look */
function orexis_(hourly) {
  let s = HtmlService.createHtmlOutputFromFile(hourly ? 'Hourly' : 'Index').getContent();
  const rep = (a, b) => { s = s.split(a).join(b); };
  rep('#7b1e2e', '#424965'); rep('#f4e6e8', '#efe0cf'); rep('#6e1423', '#151515');
  rep('<title>Kaizen KPIs</title>', '<title>Midas KPIs</title>'); rep('<title>Kaizen Hourly</title>', '<title>Midas Hourly</title>');
  rep("if(info.logo){", "if(window.OREXIS_LOGO)info.logo=window.OREXIS_LOGO;if(info.logo){");
  rep("info.name||'Kaizen KPIs'", "'Midas'");
  rep("const src=S.logo&&S.logo.data;", "const src=window.OREXIS_LOGO||(S.logo&&S.logo.data);");
  rep("(name||'Kaizen KPIs')", "'Midas'");
  rep("if(l.exists&&l.data().data)S.logo=l.data().data}catch(_){}", "if(l.exists&&l.data().data)S.logo=l.data().data}catch(_){}S.logo=window.OREXIS_LOGO||S.logo;");
  rep("h('h1',{text:'Hourly KPIs'})", "h('h1',{text:'Midas Hourly'})");
  rep('https://ameliahcrowe-cloud.github.io/kaizen-kpis/', 'https://ameliahcrowe-cloud.github.io/kaizen-kpis/orexis/');
  rep("'All of '+((S.settings.teamName||'').trim()||'Kaizen')", "'All of Midas'");
  rep('Synced to Kaizen KPIs', 'Synced to Midas KPIs'); rep('Not synced to Kaizen KPIs', 'Not synced to Midas KPIs'); rep('open Kaizen KPIs for', 'open Midas KPIs for'); rep('numbers are in Kaizen KPIs', 'numbers are in Midas KPIs'); rep('numbers to Kaizen KPIs', 'numbers to Midas KPIs');
  // no Kaizen anywhere in the Midas apps, except Ameliah's own team label
  rep("S.settings.teamName||''", "'Midas'");
  rep('Kaizen', 'Midas'); rep('Midas Enterprises', 'Midas');
  // Kaizen test people never appear in Midas
  rep('applyDocs(r.docs);', 'applyDocs((r.docs||[]).filter(d=>!/^(reps|conf|teams)\\/(u_0902e9eb82174b99b8e2|u_32e3e5910f7b422aa08f)(\\/|$)/.test(d.p)));');
  const css = ':root{--bg:#F6F3EE!important;--line:#e3dfd7!important;--muted:#5b6275!important;--ink:#141414!important;--accent:#151515!important;--accent-soft:#efe0cf!important}' +
    'body{background:#F6F3EE!important}header.top,.top{background:#151515!important;color:#D4A373!important;border-radius:14px;padding:12px 14px!important;margin-bottom:10px}' +
    'header.top h1,.top h1,#teamTitle{color:#D4A373!important;letter-spacing:.12em}#teamSub,.top .sub{color:#c9cfe0!important}' +
    'header.top button,.top button,header.top input,header.top select{background:#414A63!important;color:#fff!important;border-color:#5a6688!important}' +
    'nav.tabs button[aria-selected=true]{border-bottom-color:#D4A373!important}#mnav button[aria-selected=true]{color:#151515!important;background:#efe0cf!important}' +
    'h2,h3{color:#151515}.card{border-color:#e3dfd7}';
  s = s.replace('</style>', css + '</style>');
  s = s.replace('<head>', '<head><script>window.OREXIS=1;window.OREXIS_LOGO=' + JSON.stringify(OREXIS_LOGO) + '</script>');
  if (s.indexOf('window.OREXIS=1') < 0) s = '<script>window.OREXIS=1;window.OREXIS_LOGO=' + JSON.stringify(OREXIS_LOGO) + '</script>' + s;
  return HtmlService.createHtmlOutput(s).setTitle(hourly ? 'Midas Hourly' : 'Midas KPIs')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}
// Test run: only these people's cards, emailed straight to these addresses
function scorecardTest() {
  const names = ['Ameliah', 'Ben Cotter', 'Freddie Juffkins'], to = 'ameliah.crowe@gmail.com';
  const today = Utilities.formatDate(new Date(), 'Australia/Melbourne', 'yyyy-MM-dd'), wk = (new Date(today + 'T12:00:00Z').getUTCDay() % 6 === 0) ? mondayOf_(today) : addDaysS_(mondayOf_(today), -7), tag = wk.slice(8, 10) + '/' + wk.slice(5, 7);
  const cards = scorecards_(wk).filter(c => names.indexOf(c.name.replace(/ \(test\)$/, '')) >= 0);
  const pdfs = cards.map(c => Utilities.newBlob(page_(c.html), 'text/html', 'x.html').getAs('application/pdf').setName('Scorecard WC ' + tag.replace('/', '-') + ' - ' + c.name.replace(/ \(test\)$/, '') + '.pdf'));
  MailApp.sendEmail({ to: to, subject: 'Midas weekly scorecards · TEST · WC ' + tag, name: 'Midas',
    htmlBody: '<div style="font-family:Helvetica,Arial,sans-serif;font-size:14px">Test run of the new weekly scorecards: ' + cards.map(c => c.name.replace(/ \(test\)$/, '')).join(', ') + '. Each one is attached as a printable PDF.</div>', attachments: pdfs });
  return 'Sent ' + cards.length + ' to ' + to;
}
const SCORE_TO = 'ameliah.crowe@gmail.com';
const STD_DEFAULT = { bell: 700, gong: 1000, stops: 20, conv: 75, avg: 65, ups: 30, rph: 120, silver: 100, gold: 120, platinum: 160, diamond: 200, buyersDay: 12, stopsHr: 3, buyersHr: 2 };
let STD = Object.assign({}, STD_DEFAULT);
function setting_(t, name) { const d = t.map['settings/' + name]; return d && !d.del && d.data ? d.data : {}; }
function loadStd_(t) { const x = setting_(t, 'standards'); STD = Object.assign({}, STD_DEFAULT); Object.keys(STD_DEFAULT).forEach(k => { const v = Number(x[k]); if (isFinite(v) && v > 0) STD[k] = v; });
  BANDS.stops.strong = STD.stops; BANDS.conv.strong = STD.conv; BANDS.avg.strong = STD.avg; BANDS.ups.strong = STD.ups; }
const CARDS_KEY = 'kI3F32ZWxoTf2hSY-lQ6-c-3Vy3Y-UCo';
const SCORE_INCLUDE_TEST = true; // trial: include test people in scorecards
const BANDS = {
  stops: { low: 15, strong: 20 },   // per shift: <15 low, 15-20 avg, 20+ strong
  conv: { low: 70, strong: 80 },    // stop to buyer %: manual 70 / 80
  avg: { low: 60, strong: 70 },     // sale average $: manual $60 / $70
  ups: { low: 15, strong: 25 }      // upsell %: manual 15 / 25
};
// One mechanical piece of feedback per focus (feedback rules), from Ameliah's training docs
const FOCUS_LIB = {
  stops: { area: 'Stopping', topic: 'Stopping · Staying in the zone', why: 'Low stops = not staying in the zone. Don’t miss anyone walking past, and finish your lines on everyone you try to stop. Stop more, win more: every stop is worth about $50.',
    tip: 'SEEC, then Distance Theory 6m / 4m / 2m: “Hey!” → “Just a once off!” → “For the Paralympics!”. Eye contact from hello until they stop. Less words, more silence. High tone on hello, low tone on the call to action.' },
  conv: { area: 'Stop to buyer', topic: 'Cracking the egg · Handling concerns', why: 'Stop to buyer % is dictated by your ability to deal with concerns. If it’s low, the engagement or the intro rapport isn’t there yet.',
    tip: 'Crack the egg first: giggle = gottem. Then nail “How much is it?”: “Thanks for asking, most people run away! You basically get to pick.” How much is a green light, they want to support.' },
  avg: { area: 'Sale average', topic: 'Landing impulses · Building value', why: 'A low $ average means not enough value built in the cause, or not enough confidence in the close.',
    tip: 'Land “Paralympians are self-funded” and pause. Pitch the medals low and slow, then close assumptively on the medal pack, not the minimum.' },
  ups: { area: 'Upselling', topic: 'Upselling · A.A.A.R.R.', why: 'Every upsell is a second sale. Price is only an issue if the value isn’t there.',
    tip: '“Can I be a bit cheeky?” Then A.A.A.R.R.: Appreciate, Ask, Anchor high (the anchor isn’t the ask), Re-impulse, Re-close. Ask everyone.' },
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
  cards.forEach((c, i) => GmailApp.createDraft(to, 'Your scorecard · WC ' + tag + ' · ' + c.name, 'Your scorecard is attached.', { htmlBody: page_(c.html), attachments: [pdfs[i]], name: 'Midas' }));
  MailApp.sendEmail({ to: to, subject: 'Weekly scorecards · WC ' + tag + ' (' + cards.length + ')', name: 'Midas scorecards',
    htmlBody: '<div style="font-family:Helvetica,Arial,sans-serif;font-size:14px">' + cards.length + ' scorecards attached as PDFs, ready to print. Each one is also in your Gmail <b>Drafts</b> to add notes and send on.</div>',
    attachments: pdfs });
  return 'Sent ' + cards.length + ' scorecards to ' + to;
}
function page_(inner) { return '<html><head><meta charset="utf-8"><style>@page{size:A4;margin:12mm}body{margin:0;font-family:Helvetica,Arial,sans-serif;color:#1d1416}</style></head><body>' + inner + '</body></html>'; }
function scorecards_(wk) {
  const t = loadTable_(), M = t.map, end = addDaysS_(wk, 6); loadStd_(t);
  const get = p => (M[p] && !M[p].del) ? M[p].data : null;
  const logo = ((get('settings/logo') || {}).data) || '';
  const coName = ((get('settings/main') || {}).teamName) || 'Midas';
  const reps = Object.keys(M).filter(p => /^reps\/[^/]+$/.test(p) && !M[p].del && M[p].data && M[p].data.active !== false && (SCORE_INCLUDE_TEST || !M[p].data.test) && ['u_0902e9eb82174b99b8e2', 'u_32e3e5910f7b422aa08f'].indexOf(p.split('/')[1]) < 0).map(p => ({ id: p.split('/')[1], r: M[p].data }));
  const allReps = Object.keys(M).filter(p => /^reps\/[^/]+$/.test(p) && !M[p].del && M[p].data && M[p].data.active !== false && ['u_0902e9eb82174b99b8e2', 'u_32e3e5910f7b422aa08f'].indexOf(p.split('/')[1]) < 0).map(p => ({ id: p.split('/')[1], r: M[p].data }));
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
  const zone = v => v >= STD.diamond ? 'Diamond' : v >= STD.platinum ? 'Platinum' : v >= STD.gold ? 'Gold' : v >= STD.silver ? 'Silver' : 'Red'; // RPH zones from the KPI topic
  const esc = x => String(x == null ? '' : x).replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
  const M1 = '#7b1e2e', MUT = '#6b5b60', LINE = '#e6dcde';
  const out = [];
  { // ---- company card ----
    const N = '#424965', G = '#D4A373';
    const sec2 = t => '<div style="margin:16px 0 8px;font-size:11px;font-weight:700;letter-spacing:.14em;color:' + N + '"><span style="display:inline-block;width:18px;height:2px;background:' + G + ';vertical-align:middle;margin-right:8px"></span>' + t + '</div>';
    const tile = (l, v, sub) => '<td style="padding:0 14px 0 0;vertical-align:top"><div style="font-size:10px;letter-spacing:.1em;color:' + MUT + '">' + l + '</div><div style="font-size:24px;font-weight:700;color:' + N + '">' + v + '</div><div style="font-size:11px;color:' + MUT + '">' + (sub || '') + '</div></td>';
    const all = reps.map(x => ({ ...x, s: sum(x.id, wk, end) }));
    const tot = (list) => { const T = { stops: 0, buyers: 0, retail: 0, upsells: 0, upsellRetail: 0, medals: 0, hours: 0, shifts: 0, days: DN.map(() => 0) }; list.forEach(x => { ['stops', 'buyers', 'retail', 'upsells', 'upsellRetail', 'medals', 'hours', 'shifts'].forEach(k => T[k] += x.s[k]); x.s.days.forEach((d, i) => T.days[i] += d.retail); }); return T; };
    const T = tot(all); const K = kp(T);
    const prevT = tot(reps.map(x => ({ ...x, s: sum(x.id, addDaysS_(wk, -7), addDaysS_(wk, -1)) })));
    const worked = all.filter(x => x.s.shifts);
    const bells = all.reduce((n, x) => n + x.s.days.filter(d => d.worked && d.retail >= STD.bell && d.retail < STD.gong).length, 0), gongs = all.reduce((n, x) => n + x.s.days.filter(d => d.worked && d.retail >= STD.gong).length, 0);
    const diff = prevT.retail ? Math.round((T.retail - prevT.retail) / prevT.retail * 100) : null;
    let h = '<div style="max-width:720px;margin:0 auto;background:#fff;font-family:Helvetica,Arial,sans-serif">' +
      '<div style="background:' + N + ';color:#fff;padding:14px 20px;border-radius:6px"><table style="width:100%;border-collapse:collapse"><tr><td style="width:110px;vertical-align:middle"><img src="' + OREXIS_LOGO + '" style="height:44px;display:block;border:0"></td><td style="font-size:11px;letter-spacing:.2em;color:' + G + ';vertical-align:middle;text-align:right">COMPANY SCORECARD · WC ' + wk.slice(8, 10) + '/' + wk.slice(5, 7) + '</td></tr></table>' +
      '<div style="font-size:30px;letter-spacing:.04em;margin-top:8px">' + esc(String(coName).toUpperCase()) + '</div><div style="font-size:12px;color:#dfe2ec">' + worked.length + ' people worked · ' + T.shifts + ' shifts · ' + Math.round(T.hours) + ' hours</div></div>' +
      '<div style="font-size:14px;line-height:1.5;margin:14px 0">' + money(T.retail) + ' retail this week' + (diff == null ? '' : ', ' + (diff >= 0 ? 'up ' : 'down ') + Math.abs(diff) + '% on last week (' + money(prevT.retail) + ')') + '. ' + gongs + ' gong day' + (gongs === 1 ? '' : 's') + ' 🥁 and ' + bells + ' bell day' + (bells === 1 ? '' : 's') + ' 🔔.</div>' +
      '<table style="border-collapse:collapse"><tr>' + tile('RETAIL', money(T.retail), T.shifts + ' shifts') + tile('RPH', T.hours ? money(K.rph) : '—', T.hours ? zone(K.rph) : '') + tile('CONVERSION', T.stops ? Math.round(K.conv) + '%' : '—', 'Standard ' + BANDS.conv.strong + '%') + tile('STOPS / SHIFT', T.shifts ? K.stops.toFixed(1) : '—', 'Standard ' + BANDS.stops.strong) + tile('SALE AVERAGE', T.buyers ? money(K.avg) : '—', 'Standard ' + money(BANDS.avg.strong)) + '</tr></table>' +
      '<table style="border-collapse:collapse;margin-top:10px"><tr>' + tile('STOPS', String(T.stops)) + tile('BUYERS', String(T.buyers)) + tile('UPSELLS', String(T.upsells), T.buyers ? Math.round(T.upsells / T.buyers * 100) + '% of buyers' : '') + tile('MEDALS', String(T.medals)) + '</tr></table>';
    const mx = Math.max(1, ...T.days);
    h += sec2('THE WEEK') + '<table style="width:100%;border-collapse:collapse;table-layout:fixed"><tr style="vertical-align:bottom">' + T.days.map(v => '<td style="text-align:center;padding:0 6px;height:110px;vertical-align:bottom"><div style="font-size:10px;color:' + MUT + '">' + (v ? money(v) : '') + '</div><div style="height:' + Math.max(2, Math.round(v / mx * 90)) + 'px;background:' + (v ? N : LINE) + ';border-radius:3px 3px 0 0"></div></td>').join('') + '</tr><tr>' + DN.map(d => '<td style="text-align:center;font-size:10px;color:' + MUT + ';border-top:1px solid ' + LINE + ';padding-top:3px">' + d + '</td>').join('') + '</tr></table>';
    const wks = [-21, -14, -7, 0].map(o => { const a = addDaysS_(wk, o); return { a, x: tot(reps.map(r => ({ s: sum(r.id, a, addDaysS_(a, 6)) }))) }; });
    h += sec2('LAST 4 WEEKS') + '<table style="width:100%;border-collapse:separate;border-spacing:6px 0"><tr>' + wks.map(({ a, x }) => '<td style="border:1px solid ' + (a === wk ? G : LINE) + ';border-radius:8px;padding:8px 10px;width:25%"><div style="font-size:10px;color:' + MUT + '">WC ' + a.slice(8, 10) + '/' + a.slice(5, 7) + '</div><div style="font-size:18px;font-weight:700;color:' + N + '">' + money(x.retail) + '</div><div style="font-size:11px;color:' + MUT + '">' + x.shifts + ' shifts · ' + (x.hours ? money(x.retail / x.hours) + ' RPH' : '—') + '</div></td>').join('') + '</tr></table>';
    const th = 'text-align:left;font-size:10px;letter-spacing:.06em;color:' + MUT + ';padding:6px 8px;border-bottom:1px solid ' + LINE, td = 'padding:6px 8px;border-bottom:1px solid ' + LINE + ';font-size:12px';
    h += sec2('EVERYONE') + '<table style="width:100%;border-collapse:collapse"><tr>' + ['#', 'Name', 'Shifts', 'Retail', 'RPH', 'Conv', '$ AV', 'Stops/shift'].map(x => '<th style="' + th + '">' + x + '</th>').join('') + '</tr>' +
      worked.sort((a, b) => b.s.retail - a.s.retail).map((x, i) => { const k = kp(x.s); return '<tr><td style="' + td + '">' + (i + 1) + '</td><td style="' + td + ';font-weight:700">' + esc(x.r.name || 'Rep') + '</td><td style="' + td + '">' + x.s.shifts + '</td><td style="' + td + '">' + money(x.s.retail) + '</td><td style="' + td + '">' + (x.s.hours ? money(k.rph) : '—') + '</td><td style="' + td + '">' + Math.round(k.conv) + '%</td><td style="' + td + '">' + (x.s.buyers ? money(k.avg) : '—') + '</td><td style="' + td + '">' + k.stops.toFixed(1) + '</td></tr>'; }).join('') + '</table>';
    h += '<div style="margin-top:22px;border-top:1px solid ' + LINE + ';padding-top:8px;font-size:10px;letter-spacing:.06em;color:' + MUT + '">STANDARDS PER SHIFT: ' + BANDS.stops.strong + ' STOPS · ' + BANDS.conv.strong + '% CONVERSION · ' + money(BANDS.avg.strong) + ' SALE AVERAGE · ' + money(STD.rph) + ' RPH</div></div>';
    out.push({ name: 'Company - ' + coName, html: h });
  }
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
      const gongs = s.days.filter(d => d.worked && d.retail >= STD.gong).length, bells = s.days.filter(d => d.worked && d.retail >= STD.bell && d.retail < STD.gong).length; if (gongs) good.push(gongs + ' gong day' + (gongs > 1 ? 's' : '') + ' 🥁 ($' + STD.gong + '+)'); if (bells) good.push(bells + ' bell day' + (bells > 1 ? 's' : '') + ' 🔔 ($' + STD.bell + '+)');
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
      if (!fk && ((rev && Number(rev.motivation) && Number(rev.motivation) <= 5) || (K.rph && K.rph < STD.rph))) fk = 'mind';
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
    const N = '#424965', G = '#D4A373';
    const sec2 = t => '<div style="margin:16px 0 8px;font-size:11px;font-weight:700;letter-spacing:.14em;color:' + N + '"><span style="display:inline-block;width:18px;height:2px;background:' + G + ';vertical-align:middle;margin-right:8px"></span>' + t + '</div>';
    const wkTag = d => d.slice(8, 10) + '/' + d.slice(5, 7);
    // one-line summary
    const strong = rows.filter(x => s.shifts && band(x[1], x[5] != null ? x[5] : x[2]) === 'STRONG').map(x => x[0].toLowerCase());
    const line = !s.shifts ? 'No shifts logged this week. Let’s lock in regular shifts so you can build a rhythm.'
      : (strong.length ? 'Your ' + strong[0] + ' is in the strong range. ' : 'You made ' + s.buyers + ' sale' + (s.buyers === 1 ? '' : 's') + ' across ' + s.shifts + ' shift' + (s.shifts === 1 ? '' : 's') + '. ') + (ft ? 'Next step: ' + ft.area.toLowerCase() + '.' : 'Keep it going and chase a PB.');
    const tile = (l, v, sub) => '<td style="padding:0 14px 0 0;vertical-align:top"><div style="font-size:10px;letter-spacing:.1em;color:' + MUT + '">' + l + '</div><div style="font-size:24px;font-weight:700;color:' + N + '">' + v + '</div><div style="font-size:11px;color:' + MUT + '">' + sub + '</div></td>';
    let h = '<div style="max-width:720px;margin:0 auto;background:#fff;font-family:Helvetica,Arial,sans-serif">' +
      '<div style="background:' + N + ';color:#fff;padding:14px 20px;border-radius:6px">' +
      '<table style="width:100%;border-collapse:collapse"><tr><td style="width:110px;vertical-align:middle"><img src="' + OREXIS_LOGO + '" style="height:44px;display:block;border:0"></td><td style="font-size:11px;letter-spacing:.2em;color:' + G + ';vertical-align:middle">' + 'OREXIS INNOVATIONS · WEEKLY SCORECARD</td><td style="text-align:right;font-size:11px;color:#dfe2ec;vertical-align:middle">WC ' + wkTag(wk) + '/' + wk.slice(0, 4) + '</td></tr></table>' +
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
    const mx = Math.max(STD.bell, ...s.days.map(d => d.retail));
    h += sec2('THE WEEK') + '<table style="width:100%;border-collapse:collapse;table-layout:fixed"><tr style="vertical-align:bottom">' + s.days.map(d => { const bh = Math.round(d.retail / mx * 90); const gong = d.worked && d.retail >= STD.gong, bell = d.worked && d.retail >= STD.bell;
      return '<td style="text-align:center;padding:0 6px;height:110px;vertical-align:bottom"><div style="font-size:10px;color:' + MUT + '">' + (d.worked ? money(d.retail) : '') + (gong ? ' 🥁' : bell ? ' 🔔' : '') + '</div><div style="height:' + Math.max(d.worked ? 3 : 0, bh) + 'px;background:' + (gong ? '#C8862A' : bell ? G : N) + ';border-radius:4px 4px 0 0"></div></td>'; }).join('') + '</tr><tr>' +
      s.days.map((d, i) => '<td style="text-align:center;font-size:10px;color:' + MUT + ';border-top:1px solid ' + LINE + ';padding-top:3px">' + DN[i] + '</td>').join('') + '</tr></table>';
    // tracker: standard vs this week vs their usual (last 4 weeks)
    const TR = [['Stops / shift', K.stops, U && U.stops, BANDS.stops.strong, x => x.toFixed(1)], ['Conversion', K.conv, U && U.conv, BANDS.conv.strong, x => Math.round(x) + '%'], ['Sale average', K.avg, U && U.avg, BANDS.avg.strong, money], ['Upsells', K.ups, U && U.ups, BANDS.ups.strong, x => Math.round(x) + '%'], ['RPH', K.rph, U && U.rph, STD.rph, money]];
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
      if (K.rph && K.rph < STD.rph) ch.push(money(K.rph) + ' an hour, under the ' + money(STD.rph) + ' standard.');
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
      const word = !T.shifts ? 'Activity' : tf ? PEOPLE_NEED[tf[0]] : (T.hours && T.retail / T.hours < STD.rph ? 'Belief' : 'Recognition');
      const piece = teamPiece_(team.length, ld, ns, 0);
      const tt = (l, v, st, ok) => '<td style="border:1px solid ' + LINE + ';border-radius:8px;padding:8px;text-align:center;width:25%"><div style="font-size:10px;letter-spacing:.08em;color:' + MUT + '">' + l + '</div><div style="font-size:20px;font-weight:700;color:' + N + '">' + v + '</div><div style="font-size:10px;color:' + (T.shifts ? (ok ? '#1f7a4d' : '#c2410c') : MUT) + '">' + (T.shifts ? (ok ? '✓ ' : '↓ ') : '') + 'Standard ' + st + '</div></td>';
      const bph = T.hours ? T.buyers / T.hours : 0, spd = T.shifts ? T.stops / T.shifts : 0, bpd = T.shifts ? T.buyers / T.shifts : 0, sph = T.hours ? T.stops / T.hours : 0, rphT = T.hours ? T.retail / T.hours : 0;
      const tn = r.team === id ? coName : ((get('teams/' + id) || {}).name) || (r.name ? String(r.name).trim().split(/\s+/)[0] + '’s team' : coName);
      h += sec2('YOUR TEAM · ' + esc(String(tn).toUpperCase())) + '<div style="font-size:13px;line-height:1.6">' + team.length + ' people · ' + ld + ' leaders · ' + ns + ' new starts · ' + money(T.retail) + ' team retail</div>' +
        '<table style="width:100%;border-collapse:separate;border-spacing:6px 6px;margin-top:6px"><tr>' +
        tt('STOPS / HR', T.hours ? sph.toFixed(1) : '—', String(STD.stopsHr), sph >= STD.stopsHr) + tt('BUYERS / HR', T.hours ? bph.toFixed(1) : '—', String(STD.buyersHr), bph >= STD.buyersHr) +
        tt('STOPS / DAY', T.shifts ? spd.toFixed(1) : '—', String(BANDS.stops.strong), spd >= BANDS.stops.strong) + tt('BUYERS / DAY', T.shifts ? bpd.toFixed(1) : '—', String(STD.buyersDay), bpd >= STD.buyersDay) + '</tr></table>' +
        '<table style="width:100%;border-collapse:separate;border-spacing:6px 0"><tr>' +
        tt('CONVERSION', T.stops ? Math.round(TK.conv) + '%' : '—', BANDS.conv.strong + '%', TK.conv >= BANDS.conv.strong) + tt('$ AV', T.buyers ? money(TK.avg) : '—', money(BANDS.avg.strong), TK.avg >= BANDS.avg.strong) +
        tt('RPH', T.hours ? money(rphT) : '—', money(STD.rph), rphT >= STD.rph) + '</tr></table>' +
        two('<div style="font-size:10px;letter-spacing:.1em;color:' + MUT + ';margin-top:6px">WHAT THEY NEED</div><div style="font-size:20px;font-weight:700;color:' + N + '">' + word + '</div>', '<div style="font-size:10px;letter-spacing:.1em;color:' + MUT + ';margin-top:6px">NEXT TEAM PIECE</div><div style="font-size:13px"><b>' + piece.t + '.</b> ' + esc(piece.d) + '</div>');
    }
    h += '<div style="margin-top:22px;border-top:1px solid ' + LINE + ';padding-top:8px;font-size:10px;letter-spacing:.06em;color:' + MUT + '">STANDARDS PER SHIFT: ' + BANDS.stops.strong + ' STOPS · ' + BANDS.conv.strong + '% CONVERSION · ' + money(BANDS.avg.strong) + ' AVERAGE SALE · ' + BANDS.ups.strong + '% UPSELLS · ' + money(STD.rph) + '+ RPH</div></div>';
    out.push({ name: (r.name || 'Rep') + (r.test ? ' (test)' : ''), html: h });
  });
  return out;
}

// One-off: load the starting data (Ameliah, Ben, Freddie) copied from Kaizen
function importMidasData() {
  const csv = UrlFetchApp.fetch('https://ameliahcrowe-cloud.github.io/kaizen-kpis/x/oData.csv?' + Date.now()).getContentText();
  const rows = Utilities.parseCsv(csv);
  ensureSheets_();
  const sh = ss_().getSheetByName(DATA);
  sh.getRange(1, 1, rows.length, 4).setValues(rows.map((r, i) => i ? [r[0], r[1], r[2] === 'TRUE', r[3]] : r.slice(0, 4)));
  return 'Imported ' + (rows.length - 1) + ' rows';
}

/* ---------- printable PDFs (rendered by the Sunday GitHub job, emailed from here) ---------- */
function cardsWeek_() { const today = Utilities.formatDate(new Date(), 'Australia/Sydney', 'yyyy-MM-dd'); return (new Date(today + 'T12:00:00Z').getUTCDay() % 6 === 0) ? mondayOf_(today) : addDaysS_(mondayOf_(today), -7); }
function cardsJson_(p) {
  if (p.key !== CARDS_KEY) return ContentService.createTextOutput('{"error":"key"}').setMimeType(ContentService.MimeType.JSON);
  const sc = setting_(loadTable_(), 'scorecards');
  if (p.check) return ContentService.createTextOutput(JSON.stringify({ test: !!sc.testAt && !sc.testDone || (!!sc.testAt && sc.testAt > (sc.testDone || '')) })).setMimeType(ContentService.MimeType.JSON);
  if (!p.test && sc.enabled === false) return ContentService.createTextOutput(JSON.stringify({ cards: [], off: true })).setMimeType(ContentService.MimeType.JSON);
  const wk = p.wk || cardsWeek_(), tag = wk.slice(8, 10) + '/' + wk.slice(5, 7);
  const cards = scorecards_(wk).filter(c => !/\(test\)$/.test(c.name)).map(c => ({ name: c.name, html: page_(c.html) }));
  return ContentService.createTextOutput(JSON.stringify({ wk: wk, tag: tag, cards: cards })).setMimeType(ContentService.MimeType.JSON);
}
function doPost(e) {
  let b = {}; try { b = JSON.parse(e.postData.contents); } catch (x) { }
  if (b.key !== CARDS_KEY) return ContentService.createTextOutput('bad key');
  const t0 = loadTable_(), sc = setting_(t0, 'scorecards'), tag = b.tag || '';
  const to = b.test ? (sc.testTo || 'ameliah.crowe@gmail.com') : (sc.to || SCORE_TO);
  if (b.test) putDocs_(t0, [{ path: 'settings/scorecards', data: Object.assign({}, sc, { testDone: new Date().toISOString() }) }]);
  const at = (b.pdfs || []).map(f => Utilities.newBlob(Utilities.base64Decode(f.b64), 'application/pdf', 'Scorecard WC ' + tag.replace('/', '-') + ' - ' + f.name + '.pdf'));
  MailApp.sendEmail({ to: to, subject: 'Midas weekly scorecards · WC ' + tag, name: 'Midas',
    htmlBody: '<div style="font-family:Helvetica,Arial,sans-serif;font-size:14px">Weekly scorecards for WC ' + tag + ', attached as printable PDFs: ' + (b.pdfs || []).map(f => f.name).join(', ') + '.</div>', attachments: at });
  return ContentService.createTextOutput('sent ' + at.length + ' to ' + to);
}

/* ---------- support messages to Midas admin ---------- */
function supportMessage(token, msg) {
  const a = auth_(token); msg = String(msg || '').trim().slice(0, 3000);
  if (!msg) throw new Error('Write a message first');
  const esc = x => String(x).replace(/&/g, '&amp;').replace(/</g, '&lt;');
  MailApp.sendEmail({ to: 'ameliah.crowe@gmail.com', subject: 'Midas support · ' + a.name, name: 'Midas Support',
    htmlBody: '<div style="font-family:Helvetica,Arial,sans-serif;font-size:14px"><b>From:</b> ' + esc(a.name) + '<br><b>Sent:</b> ' + Utilities.formatDate(new Date(), 'Australia/Sydney', 'EEE d MMM, h:mm a') + '<br><br>' + esc(msg).replace(/\n/g, '<br>') + '</div>' });
  return { ok: true };
}
