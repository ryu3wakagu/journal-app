// JOURNALアプリ（GitHub版）の裏側
// 保存、AIの整理、スプレッドシートの読み書き、写真、iPhoneへの通知を受け持ちます。
//
// スクリプトプロパティ（プロジェクトの設定 → スクリプト プロパティ）
//   ANTHROPIC_API_KEY  … AIの整理に使うキー（本人が貼る）
//   APP_KEY / VAPID_*  … 自動で作られます。触らなくて大丈夫です
// 記録はすべて「JOURNAL」スプレッドシートの「日記」「やること」などのタブに入ります

const TZ = 'Asia/Tokyo';
const APP_URL = 'https://ryu3wakagu.github.io/journal-app/';
const CHUNK = 45000, CHUNKS = 10;          // 1つの記録を最大10マスに分けて保存
const SHEET_NAMES = { days: '日記', ideas: 'IDEAの記録', books: '本の記録', todos: 'やること', summaries: 'まとめの記録' };
const JOURNAL_ID = '1BsAlrZ7ktB2DPbDPpM1V8P_PekPf319675x7GPtnztY';
const PHOTO_FOLDER = '1avqj35Zpg7NB2CN5DpsqD3GbytBewWs7';

const P = () => PropertiesService.getScriptProperties();
const prop = (k) => P().getProperty(k);

// ================= 入口 =================
function doGet(e) { return handle_(e.parameter || {}); }
function doPost(e) {
  let body = {};
  try { body = JSON.parse(e.postData.contents); } catch (x) {}
  return handle_(body);
}
function out_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }

function handle_(q) {
  try {
    const a = q.action;
    // 通知の中身（iPhoneの通知係が取りに来る。登録済みの宛先だけに答える）
    if (a === 'push_latest') return out_(pushLatest_(q.endpoint));
    if (a === 'login') return out_(login_(q.pass));
    const key = prop('APP_KEY');
    if (!key || q.key !== key) return out_({ error: 'auth' });
    switch (a) {
      case 'ping': return out_({ ok: true });
      case 'db': return out_(dbOp_(q));
      case 'ai': return out_(ai_(q));
      case 'sheet_append': return out_(sheetAppend_(q));
      case 'sheet_get': return out_(sheetGet_(q));
      case 'upload': return out_(upload_(q));
      case 'photo': return out_(photo_(q));
      case 'vapid': return out_({ publicKey: vapid_().pub });
      case 'subscribe': return out_(subscribe_(q));
      case 'test_push': return out_(testPush_());
      default: return out_({ error: 'unknown action' });
    }
  } catch (err) {
    return out_({ error: String(err && err.message || err) });
  }
}

// 合言葉。最初に入れたものがそのまま合言葉になる
function login_(pass) {
  pass = String(pass || '');
  if (pass.length < 4) return { ok: false, msg: '4文字以上にしてください' };
  const h = sha_(pass + '|journal');
  const key = prop('APP_KEY');
  if (!key) { P().setProperty('APP_KEY', h); return { ok: true, key: h, first: true }; }
  return key === h ? { ok: true, key: h } : { ok: false, msg: '合言葉が違います' };
}
function sha_(s) {
  return Utilities.base64EncodeWebSafe(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, s, Utilities.Charset.UTF_8)).replace(/=+$/, '');
}

// ================= 保存（コレクションごとに1つのタブ） =================
// 列：A ID ／ B 日付 ／ C 中身の見出し ／ D 更新日時 ／ E〜 中身（JSON）
let SS_ = null;
const cache_ = {};
function ss_() { return SS_ || (SS_ = SpreadsheetApp.openById(JOURNAL_ID)); }
function tab_(coll) {
  const name = SHEET_NAMES[coll] || ('app_' + coll);
  let sh = ss_().getSheetByName(name);
  if (!sh) {
    sh = ss_().insertSheet(name);
    sh.getRange(1, 1, 1, 4).setValues([['ID', '日付', '中身', '更新日時']]);
    sh.getRange('A:N').setNumberFormat('@');
    sh.setFrozenRows(1);
  }
  return sh;
}
function rows_(coll) {
  if (cache_[coll]) return cache_[coll];
  const sh = tab_(coll), n = sh.getLastRow();
  const list = [];
  if (n > 1) {
    const vals = sh.getRange(2, 1, n - 1, 4 + CHUNKS).getValues();
    vals.forEach(function (r, i) {
      if (r[0] === '' || r[0] === null) return;
      let json = '';
      for (let k = 4; k < 4 + CHUNKS; k++) { const v = String(r[k] || ''); json += v.charAt(0) === '~' ? v.slice(1) : v; }
      let data = {};
      try { data = JSON.parse(json || '{}'); } catch (x) {}
      list.push({ row: i + 2, id: String(r[0]), data: data });
    });
  }
  return (cache_[coll] = { sh: sh, list: list });
}
function headline_(coll, d) {
  let t = '';
  if (coll === 'days') t = (d.entries || []).map(function (e) { return e.text; }).join('\n');
  else t = d.title || d.text || d.story || d.core || '';
  return String(t).slice(0, 2000);
}
function writeRow_(coll, id, data) {
  const c = rows_(coll);
  const json = JSON.stringify(data);
  if (json.length > CHUNK * CHUNKS) throw new Error('too large');
  // 先頭に「~」を付けて、=や+で始まっても式にならないようにする
  const safe = function (t) { return /^[=+\-@]/.test(t) ? ' ' + t : t; };
  const cells = [id, String(data.date || data.start || ''), safe(headline_(coll, data)), new Date().toISOString()];
  for (let k = 0; k < CHUNKS; k++) { const part = json.slice(k * CHUNK, (k + 1) * CHUNK); cells.push(part ? '~' + part : ''); }
  const hit = c.list.find(function (x) { return x.id === id; });
  if (hit) { c.sh.getRange(hit.row, 1, 1, cells.length).setValues([cells]); hit.data = data; }
  else {
    const row = c.sh.getLastRow() + 1;
    c.sh.getRange(row, 1, 1, 4 + CHUNKS).setNumberFormat('@');
    c.sh.getRange(row, 1, 1, cells.length).setValues([cells]);
    c.list.push({ row: row, id: id, data: data });
  }
}
function getDoc_(coll, id) {
  const hit = rows_(coll).list.find(function (x) { return x.id === id; });
  return hit ? hit.data : null;
}
function setDoc_(coll, id, data) { writeRow_(coll, id, data); }
function delDoc_(coll, id) {
  const c = rows_(coll), hit = c.list.find(function (x) { return x.id === id; });
  if (!hit) return;
  c.sh.getRange(hit.row, 1, 1, 4 + CHUNKS).clearContent();
  c.list = c.list.filter(function (x) { return x !== hit; });
}
const OPS = {
  '==': function (a, b) { return a === b; }, '!=': function (a, b) { return a !== b; },
  '>': function (a, b) { return a > b; }, '>=': function (a, b) { return a >= b; },
  '<': function (a, b) { return a < b; }, '<=': function (a, b) { return a <= b; }
};
function query_(coll, where, orderBy, dir, limit) {
  let list = rows_(coll).list.slice();
  (where || []).forEach(function (w) {
    const f = OPS[w[1]]; if (!f) return;
    list = list.filter(function (x) { return f(x.data[w[0]], w[2]); });
  });
  if (orderBy) {
    const s = dir === 'desc' ? -1 : 1;
    list.sort(function (a, b) {
      const va = a.data[orderBy], vb = b.data[orderBy];
      if (va === vb) return 0;
      if (va === undefined || va === null) return 1;
      if (vb === undefined || vb === null) return -1;
      return va < vb ? -s : s;
    });
  }
  if (limit) list = list.slice(0, limit);
  return list.map(function (x) { return { id: x.id, data: x.data }; });
}
function dbOp_(q) {
  const lock = LockService.getScriptLock();
  const write = ['set', 'update', 'delete', 'add', 'batch'].indexOf(q.op) >= 0;
  if (write) lock.waitLock(20000);
  try {
    switch (q.op) {
      case 'get': { const d = getDoc_(q.coll, q.id); return { exists: !!d, data: d }; }
      case 'set': setDoc_(q.coll, q.id, q.data || {}); return { ok: true };
      case 'update': {
        const cur = getDoc_(q.coll, q.id);
        if (!cur) return { error: 'not-found' };
        setDoc_(q.coll, q.id, Object.assign(cur, q.data || {})); return { ok: true };
      }
      case 'delete': delDoc_(q.coll, q.id); return { ok: true };
      case 'add': {
        const id = Utilities.getUuid().replace(/-/g, '').slice(0, 20);
        setDoc_(q.coll, id, q.data || {}); return { ok: true, id: id };
      }
      case 'query': return { docs: query_(q.coll, q.where, q.orderBy, q.dir, q.limit) };
      case 'batch': {
        (q.writes || []).forEach(function (w) {
          if (w.op === 'delete') delDoc_(w.coll, w.id);
          else if (w.op === 'update') { const cur = getDoc_(w.coll, w.id) || {}; setDoc_(w.coll, w.id, Object.assign(cur, w.data || {})); }
          else setDoc_(w.coll, w.id, w.data || {});
        });
        return { ok: true };
      }
    }
    return { error: 'bad op' };
  } finally { if (write) lock.releaseLock(); }
}

// ================= AI =================
function ai_(q) {
  const key = prop('ANTHROPIC_API_KEY');
  if (!key) return { error: 'no-key' };
  const model = q.tier === 'quick' ? (prop('MODEL_QUICK') || 'claude-haiku-5-5') : (prop('MODEL_MAIN') || 'claude-sonnet-5-5');
  let messages = q.messages;
  if (!messages) {
    const content = [];
    (q.images || []).forEach(function (im) {
      content.push({ type: 'image', source: { type: 'base64', media_type: im.mime || 'image/jpeg', data: im.data } });
    });
    content.push({ type: 'text', text: String(q.prompt || '') });
    messages = [{ role: 'user', content: content }];
  }
  const res = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', {
    method: 'post', contentType: 'application/json', muteHttpExceptions: true,
    headers: { 'x-api-key': key, 'anthropic-version': '2023-06-01' },
    payload: JSON.stringify({ model: model, max_tokens: q.maxTokens || 8000, messages: messages })
  });
  const code = res.getResponseCode();
  const j = JSON.parse(res.getContentText() || '{}');
  if (code === 429 || code === 529) return { error: 'rate_limited' };
  if (code >= 300) return { error: (j.error && j.error.message) || ('http ' + code) };
  const text = (j.content || []).filter(function (c) { return c.type === 'text'; }).map(function (c) { return c.text; }).join('');
  return { text: text, truncated: j.stop_reason === 'max_tokens' };
}
function aiJson_(prompt, tier) {
  const r = ai_({ prompt: prompt, tier: tier });
  if (r.error) throw new Error(r.error);
  const t = r.text, s = t.indexOf('{'), e = t.lastIndexOf('}');
  return JSON.parse(t.slice(s, e + 1));
}

// ================= JOURNALスプレッドシート（LOG、控え、まとめ） =================
function journal_() { return ss_(); }
function sheetAppend_(q) {
  const m = String(q.range || '').match(/^'?([^'!]+)'?!/);
  const sh = journal_().getSheetByName(m ? m[1] : q.range);
  if (!sh) return { error: 'no-sheet' };
  const rows = q.values || [];
  if (rows.length) sh.getRange(sh.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows.map(function (r) { return r.map(function (v) { return v === null || v === undefined ? '' : v; }); }));
  return { ok: true };
}
function sheetGet_(q) {
  const m = String(q.range || '').match(/^'?([^'!]+)'?!?(.*)$/);
  const sh = journal_().getSheetByName(m[1]);
  if (!sh) return { values: [] };
  const rg = m[2] ? sh.getRange(m[2]) : sh.getDataRange();
  return { values: rg.getDisplayValues() };
}

// ================= 写真（ドライブの「JOURNALアプリの写真」フォルダ） =================
function folder_() { return DriveApp.getFolderById(PHOTO_FOLDER); }
function upload_(q) {
  const blob = Utilities.newBlob(Utilities.base64Decode(q.data), q.mime || 'image/jpeg', q.name || ('page-' + Date.now() + '.jpg'));
  const f = folder_().createFile(blob);
  return { id: f.getId() };
}
function photo_(q) {
  const b = DriveApp.getFileById(q.id).getBlob();
  return { mime: b.getContentType(), data: Utilities.base64Encode(b.getBytes()) };
}

// ================= iPhoneへの通知 =================
// 宛先は「通知先」タブに保存。送るときは中身なしで鳴らし、iPhone側が push_latest で中身を取りに来る
function subscribe_(q) {
  const sub = q.sub || {};
  if (!sub.endpoint) return { error: 'no endpoint' };
  setDoc_('push', sha_(sub.endpoint), { endpoint: sub.endpoint, at: new Date().toISOString(), last: null });
  return { ok: true };
}
function pushLatest_(endpoint) {
  if (!endpoint) return { title: 'JOURNAL', body: '' };
  const d = getDoc_('push', sha_(endpoint));
  if (!d) return { title: 'JOURNAL', body: '' };
  return d.last || { title: 'JOURNAL', body: 'お知らせがあります', url: APP_URL };
}
function notify_(title, body, path) {
  const msg = { title: title, body: String(body || '').slice(0, 180), url: APP_URL + (path || ''), at: new Date().toISOString() };
  rows_('push').list.forEach(function (x) {
    const d = x.data;
    d.last = msg;
    setDoc_('push', x.id, d);
    const code = sendPush_(d.endpoint);
    if (code === 404 || code === 410) delDoc_('push', x.id);   // 解除された宛先は消す
  });
}
function testPush_() {
  notify_('JOURNAL', 'iPhoneへのお知らせのテストです。届いていれば設定は完了です。');
  return { ok: true, count: rows_('push').list.length };
}
function sendPush_(endpoint) {
  const v = vapid_();
  const aud = endpoint.match(/^https:\/\/[^/]+/)[0];
  const jwt = jwtEs256_({ typ: 'JWT', alg: 'ES256' }, { aud: aud, exp: Math.floor(Date.now() / 1000) + 12 * 3600, sub: APP_URL }, v.d);
  const res = UrlFetchApp.fetch(endpoint, {
    method: 'post', muteHttpExceptions: true, payload: '',
    headers: { TTL: '86400', Urgency: 'normal', Authorization: 'vapid t=' + jwt + ', k=' + v.pub }
  });
  return res.getResponseCode();
}

// ---------- VAPID（P-256のECDSA署名）を自前で計算 ----------
const EC = {
  p: BigInt('0xffffffff00000001000000000000000000000000ffffffffffffffffffffffff'),
  n: BigInt('0xffffffff00000000ffffffffffffffffbce6faada7179e84f3b9cac2fc632551'),
  gx: BigInt('0x6b17d1f2e12c4247f8bce6e563a440f277037d812deb33a0f4a13945d898c296'),
  gy: BigInt('0x4fe342e2fe1a7f9b8ee7eb4a7c0f9e162bce33576b315ececbb6406837bf51f5')
};
function mod_(a, m) { const r = a % m; return r >= 0n ? r : r + m; }
function inv_(a, m) {
  let t = 0n, nt = 1n, r = m, nr = mod_(a, m);
  while (nr !== 0n) { const q = r / nr; [t, nt] = [nt, t - q * nt]; [r, nr] = [nr, r - q * nr]; }
  return mod_(t, m);
}
// ヤコビ座標
function dbl_(P1) {
  const p = EC.p; const [X, Y, Z] = P1;
  if (Y === 0n || Z === 0n) return [0n, 1n, 0n];
  const YY = mod_(Y * Y, p), S = mod_(4n * X * YY, p), ZZ = mod_(Z * Z, p);
  const M = mod_(3n * (X - ZZ) * (X + ZZ), p);
  const X3 = mod_(M * M - 2n * S, p);
  const Y3 = mod_(M * (S - X3) - 8n * YY * YY, p);
  const Z3 = mod_(2n * Y * Z, p);
  return [X3, Y3, Z3];
}
function add_(P1, P2) {
  const p = EC.p;
  if (P1[2] === 0n) return P2; if (P2[2] === 0n) return P1;
  const [X1, Y1, Z1] = P1, [X2, Y2, Z2] = P2;
  const Z1Z1 = mod_(Z1 * Z1, p), Z2Z2 = mod_(Z2 * Z2, p);
  const U1 = mod_(X1 * Z2Z2, p), U2 = mod_(X2 * Z1Z1, p);
  const S1 = mod_(Y1 * Z2 * Z2Z2, p), S2 = mod_(Y2 * Z1 * Z1Z1, p);
  if (U1 === U2) return S1 === S2 ? dbl_(P1) : [0n, 1n, 0n];
  const H = mod_(U2 - U1, p), R = mod_(S2 - S1, p);
  const HH = mod_(H * H, p), HHH = mod_(H * HH, p), V = mod_(U1 * HH, p);
  const X3 = mod_(R * R - HHH - 2n * V, p);
  const Y3 = mod_(R * (V - X3) - S1 * HHH, p);
  const Z3 = mod_(H * Z1 * Z2, p);
  return [X3, Y3, Z3];
}
function mul_(k) {
  let R = [0n, 1n, 0n], Q = [EC.gx, EC.gy, 1n];
  while (k > 0n) { if (k & 1n) R = add_(R, Q); Q = dbl_(Q); k >>= 1n; }
  const zi = inv_(R[2], EC.p), zi2 = mod_(zi * zi, EC.p);
  return [mod_(R[0] * zi2, EC.p), mod_(R[1] * zi2 * zi, EC.p)];
}
function toBig_(bytes) { let x = 0n; bytes.forEach(function (b) { x = (x << 8n) + BigInt(b & 255); }); return x; }
function toBytes32_(x) { const out = []; for (let i = 31; i >= 0; i--) out.push(Number((x >> BigInt(i * 8)) & 255n)); return out; }
function signedBytes_(arr) { return arr.map(function (b) { return b > 127 ? b - 256 : b; }); }
function b64u_(bytes) { return Utilities.base64EncodeWebSafe(signedBytes_(bytes)).replace(/=+$/, ''); }
function randomScalar_() {
  let x = 0n;
  while (x === 0n || x >= EC.n) {
    const hex = (Utilities.getUuid() + Utilities.getUuid() + Utilities.getUuid()).replace(/-/g, '');
    const h = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, hex + Date.now());
    x = toBig_(h);
  }
  return x;
}
function vapid_() {
  let d = prop('VAPID_D'), pub = prop('VAPID_PUBLIC');
  if (!d || !pub) {
    const k = randomScalar_();
    const pt = mul_(k);
    d = k.toString(16);
    pub = b64u_([4].concat(toBytes32_(pt[0]), toBytes32_(pt[1])));
    P().setProperties({ VAPID_D: d, VAPID_PUBLIC: pub });
  }
  return { d: BigInt('0x' + d), pub: pub };
}
function jwtEs256_(header, claims, d) {
  const enc = function (o) { return Utilities.base64EncodeWebSafe(JSON.stringify(o), Utilities.Charset.UTF_8).replace(/=+$/, ''); };
  const input = enc(header) + '.' + enc(claims);
  const z = toBig_(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, input, Utilities.Charset.UTF_8));
  let r = 0n, s = 0n;
  while (r === 0n || s === 0n) {
    const k = randomScalar_();
    r = mod_(mul_(k)[0], EC.n);
    s = mod_(inv_(k, EC.n) * (z + r * d), EC.n);
  }
  return input + '.' + b64u_(toBytes32_(r).concat(toBytes32_(s)));
}

// ================= 決まった時間の仕事（1時間ごと） =================
function tick() {
  const now = new Date();
  const h = Number(Utilities.formatDate(now, TZ, 'H'));
  const today = Utilities.formatDate(now, TZ, 'yyyy-MM-dd');
  const dow = Number(Utilities.formatDate(now, TZ, 'u'));     // 1=月曜
  const dom = Number(Utilities.formatDate(now, TZ, 'd'));
  const st = (getDoc_('settings', 'main') || {}).prefs || {};
  const morning = st.nMorning === undefined ? 6 : Number(st.nMorning);
  const night = st.nNight === undefined ? 21 : Number(st.nNight);
  const once = function (key, fn) { const k = 'done_' + key; if (prop(k)) return; fn(); P().setProperty(k, '1'); };

  if (h === morning) {
    if (st.nMorningOn !== false) once('morning_' + today, function () { morning_(today); });
    if (dow === 1 && st.nWeekOn !== false) once('week_' + today, function () { const d = ensureSummary_('w'); if (d) notify_('週のまとめができました', periodLabel_(d) + '　' + d.story.slice(0, 80), '#sum'); });
    if (dom === 1 && st.nMonthOn !== false) once('month_' + today, function () { const d = ensureSummary_('m'); if (d) notify_('月のまとめができました', periodLabel_(d) + '　' + d.story.slice(0, 80), '#sum'); });
  }
  if (h === night && st.nNightOn !== false) {
    once('night_' + today, function () {
      const d = getDoc_('days', today);
      if (!d || !(d.entries || []).length) notify_('今日のこと、まだ書いていません', '一言だけでも残しておきましょう。');
    });
  }
}
function morning_(today) {
  const md = today.slice(5), y = Number(today.slice(0, 4));
  const hits = query_('days', [['md', '==', md]], 'year', 'desc', 20).map(function (x) { return x.data; })
    .filter(function (d) { return d.year < y && (d.entries || []).length; });
  if (!hits.length) return;
  const d = hits[0];
  const voice = (d.entries || []).some(function (e) { return e.src === 'sheet'; });
  const text = voice && d.ai && d.ai.summary ? d.ai.summary : (d.entries || []).map(function (e) { return e.text; }).join(' ');
  notify_((y - d.year) + '年前の今日', text.slice(0, 120));
}
function setupTriggers() {
  ScriptApp.getProjectTriggers().forEach(function (t) { if (t.getHandlerFunction() === 'tick') ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('tick').timeBased().everyHours(1).create();
  vapid_();
  return 'ok';
}

// ================= 週と月のまとめ =================
function ymd_(d) { return Utilities.formatDate(d, TZ, 'yyyy-MM-dd'); }
function periodLabel_(x) {
  const md = function (s) { return Number(s.slice(5, 7)) + '月' + Number(s.slice(8)) + '日'; };
  return x.kind === 'w' ? md(x.start) + '〜' + md(x.end) : Number(x.start.slice(0, 4)) + '年' + Number(x.start.slice(5, 7)) + '月';
}
function ensureSummary_(k) {
  const t = new Date(Utilities.formatDate(new Date(), TZ, "yyyy-MM-dd'T'00:00:00"));
  let start, end;
  if (k === 'w') {
    const w = (t.getDay() + 6) % 7;
    const mon = new Date(t); mon.setDate(t.getDate() - w);
    const s = new Date(mon); s.setDate(mon.getDate() - 7);
    const e = new Date(mon); e.setDate(mon.getDate() - 1);
    start = ymd_(s); end = ymd_(e);
  } else {
    start = ymd_(new Date(t.getFullYear(), t.getMonth() - 1, 1)); end = ymd_(new Date(t.getFullYear(), t.getMonth(), 0));
  }
  const id = k + '-' + start;
  const have = getDoc_('summaries', id);
  if (have) return null;
  return makeSummary_(k, start, end);
}
function makeSummary_(k, start, end) {
  const ds = query_('days', [['date', '>=', start], ['date', '<=', end]], 'date', 'asc', 40).map(function (x) { return x.data; })
    .filter(function (d) { return (d.entries || []).length; });
  if (!ds.length) return null;
  const per = k === 'w' ? 2500 : 1200;
  const body = ds.map(function (d) {
    const voice = (d.entries || []).some(function (e) { return e.src === 'sheet'; });
    const t = voice && d.ai && d.ai.summary ? d.ai.summary : (d.entries || []).map(function (e) { return e.text; }).join('\n');
    return '■' + d.date + (typeof d.mood === 'number' ? '　調子' + d.mood : '') + '\n' + t.slice(0, per) + (d.ai && d.ai.hayato ? '\n（隼人）' + d.ai.hayato : '');
  }).join('\n\n').slice(0, 45000);
  const ideas = query_('ideas', [['date', '>=', start], ['date', '<=', end]], null, null, 50).map(function (x) { return x.data; });
  const books = query_('books', [['date', '>=', start], ['date', '<=', end]], null, null, 50).map(function (x) { return x.data; });
  const moods = ds.map(function (d) { return d.mood; }).filter(function (m) { return typeof m === 'number'; });
  const label = periodLabel_({ kind: k, start: start, end: end });
  const r = aiJson_('あなたは日記の振り返りを手伝う役です。以下は' + label + 'の日記です。音声入力や手書きを文字にしたものを含み、言い間違いが混ざっています。\n'
    + '書いた本人が読み返して、この' + (k === 'w' ? '1週間' : '1か月') + 'の流れと自分の変化が分かるように、次の項目を日本語で作ってください。\n'
    + 'story: 出来事、感じていたこと、考えていたことが流れで分かる文章。' + (k === 'w' ? '350〜450字' : '500〜650字') + '。本人の目線で常体。\n'
    + 'good: よかったこと。3つまで。番号付きで1行に1つ。\n'
    + 'themes: この期間に繰り返し出てきたテーマ。3つまで。番号付きで1行に1つ。\n'
    + 'hayato: 息子の隼人に関すること。本人の言葉を残して1〜3文。なければ空文字。\n'
    + 'questions: ' + (k === 'w' ? '来週の自分への問いを1つ' : '来月の自分への問いを3つ、番号付きで1行に1つ') + '。\n'
    + '書かれていない内容は足さないこと。数字は算用数字。文中で中黒（・）は使わないこと。\n'
    + '返答はJSONのみ: {"story":"","good":"","themes":"","hayato":"","questions":""}\n\n' + body);
  const doc = {
    kind: k, start: start, end: end, partial: false, count: ds.length,
    avgMood: moods.length ? Math.round(moods.reduce(function (a, b) { return a + b; }, 0) / moods.length * 10) / 10 : null,
    story: String(r.story || ''), good: String(r.good || ''), themes: String(r.themes || ''), hayato: String(r.hayato || ''), questions: String(r.questions || ''),
    ideas: ideas.map(function (x) { return (x.title || x.core || '').slice(0, 40) + (x.next ? '（次の一歩：' + x.next + '）' : ''); }).filter(String).join('\n'),
    books: books.map(function (x) { return (x.title || '書名なし') + (x.action ? '（行動：' + x.action + '）' : ''); }).join('\n'),
    at: new Date().toISOString(), line: true
  };
  setDoc_('summaries', k + '-' + start, doc);
  // LINEに届けるため、JOURNALの「まとめ」にも1行
  try {
    const L = [(k === 'w' ? '【週のまとめ】' : '【月のまとめ】') + label, '記録 ' + doc.count + '日' + (doc.avgMood ? '　平均の調子 ' + doc.avgMood : ''), '', doc.story];
    [['よかったこと', doc.good], ['テーマ', doc.themes], ['IDEA', doc.ideas], ['読んだ本', doc.books], ['隼人', doc.hayato], [k === 'w' ? '来週への問い' : '来月への問い', doc.questions]]
      .forEach(function (p) { if (p[1]) L.push('', '■' + p[0], p[1]); });
    sheetAppend_({ range: 'まとめ!A1', values: [[k === 'w' ? '週' : '月', label, start, end, L.join('\n'), new Date().toISOString(), '']] });
  } catch (x) {}
  return doc;
}

// ================= 最初に1回だけ実行 =================
// タブとトリガーと通知の鍵を用意し、写真とスプレッドシートの許可を取る
function setup() {
  ['days', 'ideas', 'books', 'todos', 'summaries', 'meta', 'settings', 'push'].forEach(tab_);
  setupTriggers();
  folder_();
  return 'ok';
}
