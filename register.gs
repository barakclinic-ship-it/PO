/**
 * PO — self-registration by link
 * Paste this into a new Google Sheet: Extensions → Apps Script, then Deploy → New deployment → Web app
 * (Execute as: Me · Who has access: Anyone). Paste the web app URL into PO → Settings.
 *
 * The sheet only holds open links: the face crops of a photo and what students submit.
 * Each submission is erased once PO picks it up, and everything for a link is erased when it expires.
 * @OnlyCurrentDoc
 */
const MAX_DAYS = 14, MAX_CROP = 45000, MAX_FACES = 80;

function doGet(e) { return handle_((e && e.parameter) || {}); }
function doPost(e) { let b = {}; try { b = JSON.parse(e.postData.contents); } catch (x) {} return handle_(b); }
function out_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }

function handle_(b) {
  try {
    const P = PropertiesService.getScriptProperties(), a = String(b.a || '');
    if (a === 'ping') return out_({ ok: true, po: 1 });
    if (a === 'init') {
      const k = String(b.k || ''); if (k.length < 24) return out_({ ok: false, err: 'key' });
      const cur = P.getProperty('ADMIN');
      if (!cur) { P.setProperty('ADMIN', k); return out_({ ok: true }); }
      return out_(cur === k ? { ok: true } : { ok: false, err: 'locked' });
    }
    if (a === 'invite') return out_(getInvite_(tok_(b.t)));
    if (a === 'submit') return out_(withLock_(() => submit_(b)));
    const adm = P.getProperty('ADMIN');
    if (!adm || String(b.k || '') !== adm) return out_({ ok: false, err: 'auth' });
    if (a === 'create') return out_(withLock_(() => create_(b)));
    if (a === 'pull') return out_(pull_());
    if (a === 'ack') return out_(withLock_(() => ack_(b.ids || [])));
    if (a === 'close') return out_(withLock_(() => close_(tok_(b.t))));
    return out_({ ok: false, err: 'action' });
  } catch (err) { return out_({ ok: false, err: String(err) }); }
}

function tok_(t) { t = String(t || ''); if (!/^[A-Za-z0-9]{16,40}$/.test(t)) throw 'token'; return t; }
function withLock_(f) { const L = LockService.getScriptLock(); L.waitLock(15000); try { return f(); } finally { L.releaseLock(); } }
function sh_(n, head) {
  const ss = SpreadsheetApp.getActive(); let s = ss.getSheetByName(n);
  if (!s) { s = ss.insertSheet(n); s.appendRow(head); s.setFrozenRows(1); }
  return s;
}
const INV = () => sh_('links', ['token', 'created', 'expires', 'meta']);
const FACES = () => sh_('faces', ['token', 'i', 'crop']);
const SUBS = () => sh_('submissions', ['id', 'token', 'i', 'data', 'ts', 'fetched']);
function rows_(s) { const n = s.getLastRow(); return n < 2 ? [] : s.getRange(2, 1, n - 1, s.getLastColumn()).getValues(); }
function rewrite_(s, keep) { // replace all data rows with `keep`
  const n = s.getLastRow(); if (n > 1) s.getRange(2, 1, n - 1, s.getLastColumn()).clearContent();
  if (keep.length) s.getRange(2, 1, keep.length, keep[0].length).setValues(keep);
}
function cleanup_() {
  const now = Date.now(), inv = rows_(INV()), live = new Set(inv.filter(r => r[2] > now).map(r => r[0]));
  if (live.size === inv.length) return;
  rewrite_(INV(), inv.filter(r => live.has(r[0])));
  rewrite_(FACES(), rows_(FACES()).filter(r => live.has(r[0])));
  rewrite_(SUBS(), rows_(SUBS()).filter(r => live.has(r[1]) || (r[5] === '' && r[3] !== ''))); // never drop data PO hasn't picked up
}
function inv_(t) { const r = rows_(INV()).find(x => x[0] === t); return r && r[2] > Date.now() ? { meta: JSON.parse(r[3]), exp: r[2] } : null; }

function create_(b) {
  cleanup_();
  const t = tok_(b.t), days = Math.min(MAX_DAYS, Math.max(1, +b.days || 7)), m = b.meta || {};
  const meta = { title: String(m.title || '').slice(0, 120), fields: (m.fields || []).map(String).slice(0, 20), lang: String(m.lang || 'he').slice(0, 5) };
  const faces = (b.faces || []).slice(0, MAX_FACES).filter(f => typeof f.crop === 'string' && /^data:image\/jpeg;base64,/.test(f.crop) && f.crop.length < MAX_CROP);
  if (!faces.length) return { ok: false, err: 'faces' };
  if (rows_(INV()).some(r => r[0] === t)) return { ok: false, err: 'exists' };
  const exp = Date.now() + days * 864e5;
  INV().appendRow([t, Date.now(), exp, JSON.stringify(meta)]);
  const fs = FACES(), at = fs.getLastRow() + 1;
  fs.getRange(at, 1, faces.length, 3).setValues(faces.map(f => [t, Math.floor(+f.i), f.crop]));
  return { ok: true, exp };
}
function getInvite_(t) {
  const v = inv_(t); if (!v) return { ok: false, err: 'expired' };
  const taken = new Set(rows_(SUBS()).filter(r => r[1] === t).map(r => r[2]));
  const faces = rows_(FACES()).filter(r => r[0] === t && !taken.has(r[1])).map(r => ({ i: r[1], crop: r[2] }));
  return { ok: true, meta: v.meta, exp: v.exp, faces };
}
function clean_(s, n) { return String(s == null ? '' : s).replace(/[\u0000-\u001f]/g, ' ').trim().slice(0, n); }
function submit_(b) {
  const t = tok_(b.t), v = inv_(t); if (!v) return { ok: false, err: 'expired' };
  const i = Math.floor(+b.i);
  if (!rows_(FACES()).some(r => r[0] === t && r[1] === i)) return { ok: false, err: 'taken' };
  if (rows_(SUBS()).some(r => r[1] === t && r[2] === i)) return { ok: false, err: 'taken' };
  if (b.consent !== true) return { ok: false, err: 'consent' };
  const first = clean_(b.first, 60), last = clean_(b.last, 60);
  if (!first && !last) return { ok: false, err: 'name' };
  const extra = {}; (v.meta.fields || []).forEach(f => { const x = clean_((b.extra || {})[f], 200); if (x) extra[f] = x; });
  SUBS().appendRow([Utilities.getUuid(), t, i, JSON.stringify({ first, last, extra }), Date.now(), '']);
  return { ok: true };
}
function pull_() {
  cleanup_();
  const subs = rows_(SUBS()).filter(r => r[5] === '' && r[3] !== '').map(r => { const d = JSON.parse(r[3]); return { id: r[0], t: r[1], i: r[2], first: d.first, last: d.last, extra: d.extra || {}, ts: r[4] }; });
  return { ok: true, subs };
}
function ack_(ids) { // PO has it: erase the personal details and the face crop, keep only "face i is taken"
  ids = new Set(ids.map(String)); const gone = new Set();
  const subs = rows_(SUBS()).map(r => { if (ids.has(String(r[0]))) { gone.add(r[1] + '|' + r[2]); return [r[0], r[1], r[2], '', r[4], 1]; } return r; });
  rewrite_(SUBS(), subs);
  if (gone.size) rewrite_(FACES(), rows_(FACES()).filter(r => !gone.has(r[0] + '|' + r[1])));
  return { ok: true };
}
function close_(t) {
  rewrite_(INV(), rows_(INV()).filter(r => r[0] !== t));
  rewrite_(FACES(), rows_(FACES()).filter(r => r[0] !== t));
  rewrite_(SUBS(), rows_(SUBS()).filter(r => r[1] !== t || (r[5] === '' && r[3] !== '')));
  return { ok: true };
}
