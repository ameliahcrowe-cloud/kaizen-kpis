/**
 * Kaizen KPIs — Google Apps Script web app.
 * Paste this into Code.gs (Extensions → Apps Script in the Kaizen KPIs Google Sheet).
 * All data lives in this spreadsheet. People log in with their name and a PIN.
 */
const APP_TITLE = 'Saint KPIs';
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
    .setTitle('Saint Hourly').setFaviconUrl('https://ameliahcrowe-cloud.github.io/kaizen-kpis/icons/saint-512.png')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle(APP_TITLE).setFaviconUrl('https://ameliahcrowe-cloud.github.io/kaizen-kpis/icons/saint-512.png')
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
function hrsSoFar_(d, date) { const shift = Number(d.targets && d.targets.shiftHours) || 8; const tz = 'Australia/Melbourne'; if (date !== Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd')) return shift; const st = String(d.start || '').match(/(\d+):(\d+)/); if (!st) return shift; const nowS = Utilities.formatDate(new Date(), tz, 'HH:mm').split(':'); const el = (Number(nowS[0]) * 60 + Number(nowS[1]) - (Number(st[1]) * 60 + Number(st[2]))) / 60; return Math.max(0.5, Math.min(shift, Math.ceil(el * 2) / 2)); }
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
      const body = Object.assign({}, cur, tot, { date, week: mondayOf_(date), month: date.slice(0, 7), hours: ((Number(cur.hours) || 0) > 0 && !cur.fromHourly) ? Number(cur.hours) : hrsSoFar_(doc.data, date), fromHourly: true, roadie: !!doc.data.roadie, updatedAt: nowIso }); if (String(doc.data.cash || '').trim()) body.cash = Number(doc.data.cash) || 0; if (String(doc.data.eftpos || '').trim()) body.eftpos = Number(doc.data.eftpos) || 0;
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
  const html = '<div style="font-family:Arial,sans-serif;color:#161616"><h2 style="color:#111111">Saint week · WC ' + wk + '</h2><p><b>' + money(tot) + '</b> retail' + (th ? ' · <b>' + money(tot / th) + '</b> RPH' : '') + ' · ' + rows.length + ' people</p>' +
    '<table cellpadding="6" style="border-collapse:collapse;font-size:14px"><tr style="background:#111111;color:#fff"><th align="left">Person</th><th>Retail</th><th>RPH</th><th>Stops/Buyers</th><th>Days</th></tr>' + tr + '</table>' +
    '<h3>📝 Missing reflections</h3><p>' + (missing.length ? missing.join('<br>') : 'Everyone reflected on every day worked 🎉') + '</p>' +
    '<h3>🌱 New starters</h3><p>' + (newbies.length ? newbies.join('<br>') : 'No new starters this fortnight.') + '</p></div>';
  MailApp.sendEmail({ to: SCORE_TO, subject: 'Saint week · WC ' + wk + ' · ' + money(tot), htmlBody: html, name: 'Saint' });
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
  return { name: main.teamName || 'Saint', logo: logo.data || null, teams, hasOwner: readAccess_().some(a => a.role === 'owner') };
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
    if (!acc.some(a => a.role === 'owner')) throw new Error('The app isn’t set up yet. Ask Saint admin to finish setup.');
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
  if (top === 'teams') return s.length === 2 && (id === c.uid || (id === 'company' && c.role === 'tl'));
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
        if (now !== was && !(seg[1] === c.uid && now === c.uid)) throw new Error('Only Saint admin can change this.');
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
    if (a.role === 'owner' && (me.uid !== 'u_owner' || action !== 'pin')) throw new Error('Only Saint admin can reset an owner PIN. Your own PIN is changed from the spreadsheet menu.');
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
const OREXIS_LOGO = 'data:image/png;base64,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';
const OREXIS_LOGO_LIGHT = 'data:image/jpeg;base64,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';
/* Saint-branded versions of the two apps: same data, same logins, Saint look */
function orexis_(hourly) {
  let s = HtmlService.createHtmlOutputFromFile(hourly ? 'Hourly' : 'Index').getContent();
  const rep = (a, b) => { s = s.split(a).join(b); };
  rep('#7b1e2e', '#3a3a3a'); rep('#f4e6e8', '#fbe1d8'); rep('#6e1423', '#111111');
  rep('<title>Kaizen KPIs</title>', '<title>Saint KPIs</title>'); rep('<title>Kaizen Hourly</title>', '<title>Saint Hourly</title>');
  rep("if(info.logo){", "if(window.OREXIS_LOGO)info.logo=window.OREXIS_LOGO;if(info.logo){");
  rep("info.name||'Kaizen KPIs'", "'Saint'");
  rep("const src=S.logo&&S.logo.data;", "const src=window.OREXIS_LOGO||(S.logo&&S.logo.data);");
  rep("(name||'Kaizen KPIs')", "'Saint'");
  rep("if(l.exists&&l.data().data)S.logo=l.data().data}catch(_){}", "if(l.exists&&l.data().data)S.logo=l.data().data}catch(_){}S.logo=window.OREXIS_LOGO||S.logo;");
  rep("h('h1',{text:'Hourly KPIs'})", "h('h1',{text:'Saint Hourly'})");
  rep('https://ameliahcrowe-cloud.github.io/kaizen-kpis/', 'https://ameliahcrowe-cloud.github.io/kaizen-kpis/orexis/');
  rep("'All of '+((S.settings.teamName||'').trim()||'Kaizen')", "'All of Saint'");
  rep('Synced to Kaizen KPIs', 'Synced to Saint KPIs'); rep('Not synced to Kaizen KPIs', 'Not synced to Saint KPIs'); rep('open Kaizen KPIs for', 'open Saint KPIs for'); rep('numbers are in Kaizen KPIs', 'numbers are in Saint KPIs'); rep('numbers to Kaizen KPIs', 'numbers to Saint KPIs');
  // no Kaizen anywhere in the Saint apps, except Ameliah's own team label
  rep("S.settings.teamName||''", "'Saint'");
  rep('Kaizen', 'Saint'); rep('Saint Enterprises', 'Saint');
  // Kaizen test people never appear in Saint
  rep('applyDocs(r.docs);', 'applyDocs((r.docs||[]).filter(d=>!/^(reps|conf|teams)\\/(u_0902e9eb82174b99b8e2|u_32e3e5910f7b422aa08f)(\\/|$)/.test(d.p)));');
  const css = ':root{--bg:#fbe1d8!important;--line:#e3dfd7!important;--muted:#3a3a3a!important;--ink:#0b0b0b!important;--accent:#111111!important;--accent-soft:#fbe1d8!important}' +
    'body{background:#fbe1d8!important}header.top,.top{background:#111111!important;color:#E8552E!important;border-radius:14px;padding:12px 14px!important;margin-bottom:10px}' +
    'header.top h1,.top h1,#teamTitle{color:#E8552E!important;letter-spacing:.12em}#teamSub,.top .sub{color:#e9e6e4!important}' +
    'header.top button,.top button,header.top input,header.top select{background:#3a3a3a!important;color:#fff!important;border-color:#3a3a3a!important}' +
    'nav.tabs button[aria-selected=true]{border-bottom-color:#E8552E!important}#mnav button[aria-selected=true]{color:#111111!important;background:#fbe1d8!important}' +
    'h2,h3{color:#111111}.card{border-color:#e3dfd7}';
  s = s.replace('</style>', css + '</style>');
  s = s.replace('<head>', '<head><script>window.OREXIS=1;window.OREXIS_LOGO=' + JSON.stringify(OREXIS_LOGO) + '</script>');
  if (s.indexOf('window.OREXIS=1') < 0) s = '<script>window.OREXIS=1;window.OREXIS_LOGO=' + JSON.stringify(OREXIS_LOGO) + '</script>' + s;
  return HtmlService.createHtmlOutput(s).setTitle(hourly ? 'Saint Hourly' : 'Saint KPIs')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}
// Test run: only these people's cards, emailed straight to these addresses
function scorecardTest() {
  const names = ['Ameliah', 'Ben Cotter', 'Freddie Juffkins'], to = 'ameliah.crowe@gmail.com';
  const today = Utilities.formatDate(new Date(), 'Australia/Melbourne', 'yyyy-MM-dd'), wk = (new Date(today + 'T12:00:00Z').getUTCDay() % 6 === 0) ? mondayOf_(today) : addDaysS_(mondayOf_(today), -7), tag = wk.slice(8, 10) + '/' + wk.slice(5, 7);
  const cards = scorecards_(wk).filter(c => names.indexOf(c.name.replace(/ \(test\)$/, '')) >= 0);
  const pdfs = cards.map(c => Utilities.newBlob(page_(c.html), 'text/html', 'x.html').getAs('application/pdf').setName('Scorecard WC ' + tag.replace('/', '-') + ' - ' + c.name.replace(/ \(test\)$/, '') + '.pdf'));
  MailApp.sendEmail({ to: to, subject: 'Saint weekly scorecards · TEST · WC ' + tag, name: 'Saint',
    htmlBody: '<div style="font-family:Helvetica,Arial,sans-serif;font-size:14px">Test run of the new weekly scorecards: ' + cards.map(c => c.name.replace(/ \(test\)$/, '')).join(', ') + '. Each one is attached as a printable PDF.</div>', attachments: pdfs });
  return 'Sent ' + cards.length + ' to ' + to;
}
const SCORE_TO = 'ameliah.crowe@gmail.com';
const STD_DEFAULT = { bell: 700, gong: 1000, stops: 20, conv: 75, avg: 65, ups: 30, rph: 120, silver: 100, gold: 120, platinum: 160, diamond: 200, buyersDay: 12, stopsHr: 3, buyersHr: 2 };
let STD = Object.assign({}, STD_DEFAULT);
function setting_(t, name) { const d = t.map['settings/' + name]; return d && !d.del && d.data ? d.data : {}; }
function loadStd_(t) { const x = setting_(t, 'standards'); STD = Object.assign({}, STD_DEFAULT); Object.keys(STD_DEFAULT).forEach(k => { const v = Number(x[k]); if (isFinite(v) && v > 0) STD[k] = v; });
  BANDS.stops.strong = STD.stops; BANDS.conv.strong = STD.conv; BANDS.avg.strong = STD.avg; BANDS.ups.strong = STD.ups; }
const CARDS_KEY = 'jjpxF5vQLDu8pzQcREH1mW79SzTmAP7x';
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
  cards.forEach((c, i) => GmailApp.createDraft(to, 'Your scorecard · WC ' + tag + ' · ' + c.name, 'Your scorecard is attached.', { htmlBody: page_(c.html), attachments: [pdfs[i]], name: 'Saint' }));
  MailApp.sendEmail({ to: to, subject: 'Weekly scorecards · WC ' + tag + ' (' + cards.length + ')', name: 'Saint scorecards',
    htmlBody: '<div style="font-family:Helvetica,Arial,sans-serif;font-size:14px">' + cards.length + ' scorecards attached as PDFs, ready to print. Each one is also in your Gmail <b>Drafts</b> to add notes and send on.</div>',
    attachments: pdfs });
  return 'Sent ' + cards.length + ' scorecards to ' + to;
}
function page_(inner) { return '<html><head><meta charset="utf-8"><style>@page{size:A4;margin:12mm}body{margin:0;font-family:Helvetica,Arial,sans-serif;color:#1d1416}</style></head><body>' + inner + '</body></html>'; }
function scorecards_(wk) {
  const t = loadTable_(), M = t.map, end = addDaysS_(wk, 6); loadStd_(t);
  const get = p => (M[p] && !M[p].del) ? M[p].data : null;
  const logo = ((get('settings/logo') || {}).data) || '';
  const coName = ((get('settings/main') || {}).teamName) || 'Saint';
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
  const BC = { STRONG: '#C1272D', AVG: '#C1272D', LOW: '#c2410c' };
  const money = n => '$' + Math.round(n).toLocaleString('en-AU');
  const DN = ['Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat', 'Sun'];
  const ROLE = { sa: 'Sales Associate', sc: 'Sales Consultant', sl: 'Sales Leader', tl: 'Team Leader', tm: 'Team Manager', ad: 'Assistant Director', dir: 'Director', md: 'Managing Director' };
  const zone = v => v >= STD.diamond ? 'Diamond' : v >= STD.platinum ? 'Platinum' : v >= STD.gold ? 'Gold' : v >= STD.silver ? 'Silver' : 'Red'; // RPH zones from the KPI topic
  const esc = x => String(x == null ? '' : x).replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
  const M1 = '#7b1e2e', MUT = '#6b5b60', LINE = '#e6dcde';
  const out = [];
  { // ---- company card ----
    const N = '#3a3a3a', G = '#E8552E';
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
      '<div style="font-size:30px;letter-spacing:.04em;margin-top:8px">' + esc(String(coName).toUpperCase()) + '</div><div style="font-size:12px;color:#e9e6e4">' + worked.length + ' people worked · ' + T.shifts + ' shifts · ' + Math.round(T.hours) + ' hours</div></div>' +
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
    const N = '#3a3a3a', G = '#E8552E';
    const sec2 = t => '<div style="margin:16px 0 8px;font-size:11px;font-weight:700;letter-spacing:.14em;color:' + N + '"><span style="display:inline-block;width:18px;height:2px;background:' + G + ';vertical-align:middle;margin-right:8px"></span>' + t + '</div>';
    const wkTag = d => d.slice(8, 10) + '/' + d.slice(5, 7);
    // one-line summary
    const strong = rows.filter(x => s.shifts && band(x[1], x[5] != null ? x[5] : x[2]) === 'STRONG').map(x => x[0].toLowerCase());
    const line = !s.shifts ? 'No shifts logged this week. Let’s lock in regular shifts so you can build a rhythm.'
      : (strong.length ? 'Your ' + strong[0] + ' is in the strong range. ' : 'You made ' + s.buyers + ' sale' + (s.buyers === 1 ? '' : 's') + ' across ' + s.shifts + ' shift' + (s.shifts === 1 ? '' : 's') + '. ') + (ft ? 'Next step: ' + ft.area.toLowerCase() + '.' : 'Keep it going and chase a PB.');
    const tile = (l, v, sub) => '<td style="padding:0 14px 0 0;vertical-align:top"><div style="font-size:10px;letter-spacing:.1em;color:' + MUT + '">' + l + '</div><div style="font-size:24px;font-weight:700;color:' + N + '">' + v + '</div><div style="font-size:11px;color:' + MUT + '">' + sub + '</div></td>';
    let h = '<div style="max-width:720px;margin:0 auto;background:#fff;font-family:Helvetica,Arial,sans-serif">' +
      '<div style="background:' + N + ';color:#fff;padding:14px 20px;border-radius:6px">' +
      '<table style="width:100%;border-collapse:collapse"><tr><td style="width:110px;vertical-align:middle"><img src="' + OREXIS_LOGO + '" style="height:44px;display:block;border:0"></td><td style="font-size:11px;letter-spacing:.2em;color:' + G + ';vertical-align:middle">' + 'OREXIS INNOVATIONS · WEEKLY SCORECARD</td><td style="text-align:right;font-size:11px;color:#e9e6e4;vertical-align:middle">WC ' + wkTag(wk) + '/' + wk.slice(0, 4) + '</td></tr></table>' +
      '<div style="font-size:30px;letter-spacing:.04em;margin-top:8px">' + esc((r.name || 'Rep').toUpperCase()) + '</div>' +
      '<div style="font-size:12px;color:#e9e6e4">' + role + ' · ' + s.shifts + ' shift' + (s.shifts === 1 ? '' : 's') + ' this week</div></div>' +
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
      return '<td style="text-align:center;padding:0 6px;height:110px;vertical-align:bottom"><div style="font-size:10px;color:' + MUT + '">' + (d.worked ? money(d.retail) : '') + (gong ? ' 🥁' : bell ? ' 🔔' : '') + '</div><div style="height:' + Math.max(d.worked ? 3 : 0, bh) + 'px;background:' + (gong ? '#E8552E' : bell ? G : N) + ';border-radius:4px 4px 0 0"></div></td>'; }).join('') + '</tr><tr>' +
      s.days.map((d, i) => '<td style="text-align:center;font-size:10px;color:' + MUT + ';border-top:1px solid ' + LINE + ';padding-top:3px">' + DN[i] + '</td>').join('') + '</tr></table>';
    // tracker: standard vs this week vs their usual (last 4 weeks)
    const TR = [['Stops / shift', K.stops, U && U.stops, BANDS.stops.strong, x => x.toFixed(1)], ['Conversion', K.conv, U && U.conv, BANDS.conv.strong, x => Math.round(x) + '%'], ['Sale average', K.avg, U && U.avg, BANDS.avg.strong, money], ['Upsells', K.ups, U && U.ups, BANDS.ups.strong, x => Math.round(x) + '%'], ['RPH', K.rph, U && U.rph, STD.rph, money]];
    const bar = (v, m, c, lab) => '<tr><td style="width:96px;white-space:nowrap;font-size:10px;color:' + MUT + ';padding:1px 6px 1px 0">' + lab + '</td><td style="padding:1px 0"><div style="height:8px;width:' + Math.max(1, Math.round(Math.min(1, (v || 0) / m) * 100)) + '%;background:' + c + ';border-radius:4px"></div></td></tr>';
    h += sec2('TRACKER · STANDARD vs YOU vs YOUR USUAL') + '<table style="width:100%;border-collapse:collapse">' + TR.map(([l, v, u, st, f]) => { const m = Math.max(st, v || 0, u || 0) * 1.05;
      return '<tr style="vertical-align:middle"><td style="width:120px;padding:6px 0;border-bottom:1px solid ' + LINE + '"><div style="font-size:12px;font-weight:700">' + l + '</div><div style="font-size:11px;color:' + MUT + '">' + (s.shifts ? f(v) : '—') + ' <span style="color:' + (s.shifts && v >= st ? '#C1272D' : '#c2410c') + '">' + (s.shifts ? (v >= st ? '✓' : '↓') : '') + '</span></div></td><td style="padding:6px 0;border-bottom:1px solid ' + LINE + '"><table style="width:100%;border-collapse:collapse">' +
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
      const tt = (l, v, st, ok) => '<td style="border:1px solid ' + LINE + ';border-radius:8px;padding:8px;text-align:center;width:25%"><div style="font-size:10px;letter-spacing:.08em;color:' + MUT + '">' + l + '</div><div style="font-size:20px;font-weight:700;color:' + N + '">' + v + '</div><div style="font-size:10px;color:' + (T.shifts ? (ok ? '#C1272D' : '#c2410c') : MUT) + '">' + (T.shifts ? (ok ? '✓ ' : '↓ ') : '') + 'Standard ' + st + '</div></td>';
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
function importSaintData() {
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
  if (b.api) return api_(b);
  if (b.key !== CARDS_KEY) return ContentService.createTextOutput('bad key');
  const t0 = loadTable_(), sc = setting_(t0, 'scorecards'), tag = b.tag || '';
  const to = b.test ? (sc.testTo || 'ameliah.crowe@gmail.com') : (sc.to || SCORE_TO);
  if (b.test) putDocs_(t0, [{ path: 'settings/scorecards', data: Object.assign({}, sc, { testDone: new Date().toISOString() }) }]);
  const at = (b.pdfs || []).map(f => Utilities.newBlob(Utilities.base64Decode(f.b64), 'application/pdf', 'Scorecard WC ' + tag.replace('/', '-') + ' - ' + f.name + '.pdf'));
  MailApp.sendEmail({ to: to, subject: 'Saint weekly scorecards · WC ' + tag, name: 'Saint',
    htmlBody: '<div style="font-family:Helvetica,Arial,sans-serif;font-size:14px">Weekly scorecards for WC ' + tag + ', attached as printable PDFs: ' + (b.pdfs || []).map(f => f.name).join(', ') + '.</div>', attachments: at });
  return ContentService.createTextOutput('sent ' + at.length + ' to ' + to);
}

/* ---------- support messages to Saint admin ---------- */
function supportMessage(token, msg) {
  const a = auth_(token); msg = String(msg || '').trim().slice(0, 3000);
  if (!msg) throw new Error('Write a message first');
  const esc = x => String(x).replace(/&/g, '&amp;').replace(/</g, '&lt;');
  MailApp.sendEmail({ to: 'ameliah.crowe@gmail.com', subject: 'Saint support · ' + a.name, name: 'Saint Support',
    htmlBody: '<div style="font-family:Helvetica,Arial,sans-serif;font-size:14px"><b>From:</b> ' + esc(a.name) + '<br><b>Sent:</b> ' + Utilities.formatDate(new Date(), 'Australia/Sydney', 'EEE d MMM, h:mm a') + '<br><br>' + esc(msg).replace(/\n/g, '<br>') + '</div>' });
  return { ok: true };
}

function api_(b) {
  const F = { sync: sync, write: write, login: login, logout: logout, register: register, publicInfo: publicInfo, adminAccess: adminAccess, supportMessage: supportMessage };
  let out;
  try { if (!F[b.api]) throw new Error('BAD_API'); out = { ok: true, r: F[b.api].apply(null, b.args || []) }; }
  catch (err) { out = { ok: false, e: String(err && err.message || err) }; }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}
