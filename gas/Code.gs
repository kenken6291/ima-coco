/**
 * 今CoCo（いまココ） バックエンド — Google Apps Script
 * GitHub Pages + GAS + Google Drive + Google スプレッドシート + Gemini API
 *
 * スクリプトプロパティ:
 *   GEMINI_API_KEY        (必須) Google AI Studio の API キー
 *   GEMINI_MODEL          (任意) 既定: gemini-3.8-flash
 *   APP_URL               (任意) 既定: https://kenken6291.github.io/ima-coco/  （メール本文のリンク）
 *   SPREADSHEET_ID        (setup() が自動設定)
 *   DRIVE_FOLDER_ID       (setup() が自動設定)
 *   PEPPER                (setup() が自動設定)
 *   MODERATION_FAIL_OPEN  (任意) "true" で Gemini 障害時も投稿を通す（既定: 通さない）
 *   TRASH_MEDIA_ON_EXPIRE (任意) "false" で期限切れ時のDriveファイル削除を無効化
 *
 * 初回: setup() を一度実行 → ウェブアプリとしてデプロイ
 */

const CONFIG = {
  SHEET_EVENTS: 'events',
  SHEET_USERS: 'users',
  SHEET_REQUESTS: 'requests',
  GEMINI_MODEL_DEFAULT: 'gemini-3.8-flash',
  APP_URL_DEFAULT: 'https://kenken6291.github.io/ima-coco/',
  MAX_ACTIVE_POSTS_PER_USER: 5,
  MAX_MEDIA_BYTES: 20 * 1024 * 1024,
  MAX_EXPIRE_HOURS: 72,
  TOKEN_TTL_DAYS: 90,
  LOGIN_MAX_FAIL: 5,
  LOGIN_LOCK_SEC: 900,          // 5回失敗で15分ロック
  TEMP_PASSWORD_TTL_HOURS: 72,  // 仮パスワードの有効期限
  FORGOT_INTERVAL_SEC: 180,     // 再発行メールの連続送信を防ぐ間隔
  EVENTS_CACHE_SEC: 20,
  HASH_ROUNDS: 100,
};

const EVENT_HEADERS = [
  'id', 'created_at', 'expires_at', 'title', 'category', 'icon', 'summary', 'description',
  'lat', 'lng', 'media_url', 'media_type', 'capacity', 'fee', 'conditions', 'user_id', 'status',
  'hashtags', 'media_file_id'
];
const USER_HEADERS = [
  'user_id', 'nickname', 'email', 'pass_hash', 'salt', 'must_change',
  'token', 'token_expires', 'created_at', 'last_login', 'temp_issued_at'
];
const REQUEST_HEADERS = ['id', 'event_id', 'user_id', 'nickname', 'message', 'created_at'];

const CATEGORIES = {
  camp: { label: 'キャンプ', icon: '⛺' },
  bbq: { label: 'BBQ', icon: '🍖' },
  fes: { label: 'フェス', icon: '🎪' },
  karaoke: { label: 'カラオケ', icon: '🎤' },
  boardgame: { label: 'ボドゲ', icon: '🎲' },
  pingpong: { label: '卓球', icon: '🏓' },
  running: { label: 'ランニング', icon: '🏃' },
  campingcar: { label: 'キャンピングカー', icon: '🚐' },
  cafe: { label: 'カフェ', icon: '☕' },
  sports: { label: 'スポーツ', icon: '⚽' },
  other: { label: 'その他', icon: '📍' },
};

/* =========================================================
 * 初期セットアップ
 * ========================================================= */
function setup() {
  const props = PropertiesService.getScriptProperties();
  let ss;
  const sid = props.getProperty('SPREADSHEET_ID');
  if (sid) {
    ss = SpreadsheetApp.openById(sid);
  } else {
    ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) throw new Error('スプレッドシートに紐づいたGASで実行するか、SPREADSHEET_ID を設定してください');
    props.setProperty('SPREADSHEET_ID', ss.getId());
  }
  ensureSheet_(ss, CONFIG.SHEET_EVENTS, EVENT_HEADERS);
  ensureSheet_(ss, CONFIG.SHEET_USERS, USER_HEADERS);
  ensureSheet_(ss, CONFIG.SHEET_REQUESTS, REQUEST_HEADERS);

  if (!props.getProperty('DRIVE_FOLDER_ID')) {
    const folder = DriveApp.createFolder('今CoCo_media');
    props.setProperty('DRIVE_FOLDER_ID', folder.getId());
  }
  if (!props.getProperty('PEPPER')) {
    props.setProperty('PEPPER', Utilities.getUuid() + Utilities.getUuid());
  }
  setupTriggers();
  Logger.log('セットアップ完了');
  Logger.log('SPREADSHEET_ID: ' + props.getProperty('SPREADSHEET_ID'));
  Logger.log('DRIVE_FOLDER_ID: ' + props.getProperty('DRIVE_FOLDER_ID'));
  Logger.log('GEMINI_API_KEY: ' + (props.getProperty('GEMINI_API_KEY') ? '設定済み' : '未設定（スクリプトプロパティに追加してください）'));
  Logger.log('メール送信の残り回数（本日）: ' + MailApp.getRemainingDailyQuota());
}

/** シートが無ければ作成、足りない列は末尾に追加（既存データはそのまま） */
function ensureSheet_(ss, name, headers) {
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  const lastCol = Math.max(sh.getLastColumn(), 1);
  const cur = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(String);
  if (cur.join('') === '') {
    sh.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
    sh.setFrozenRows(1);
  } else {
    const missing = headers.filter(h => cur.indexOf(h) < 0);
    if (missing.length) sh.getRange(1, lastCol + 1, 1, missing.length).setValues([missing]).setFontWeight('bold');
  }
  // 自動変換（日付化・数値化）を防ぐため書式なしテキストに
  sh.getRange(1, 1, sh.getMaxRows(), sh.getLastColumn()).setNumberFormat('@');
  return sh;
}

function setupTriggers() {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'expireEvents')
    .forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('expireEvents').timeBased().everyMinutes(10).create();
}

/** APIキー動作確認用（エディタから実行） */
function testGemini() {
  const r = geminiReview_({ title: 'テスト', description: '河原でBBQしてます。あと2人どうぞ！', category: 'bbq', capacity: 'あと2人', fee: '割り勘', conditions: '誰でも' });
  Logger.log(JSON.stringify(r, null, 2));
}

/** メール送信の動作確認用（エディタから実行。自分宛てに届きます） */
function testMail() {
  const me = Session.getActiveUser().getEmail();
  sendTempPasswordMail_(me, 'テスト', 'Abc12345xy', false);
  Logger.log('送信しました: ' + me + ' / 本日の残り: ' + MailApp.getRemainingDailyQuota());
}

/* =========================================================
 * Web API
 * ========================================================= */
function doGet(e) {
  try {
    const p = (e && e.parameter) || {};
    const action = p.action || 'getEvents';
    switch (action) {
      case 'getEvents':
        return json_({ ok: true, events: getEvents_(), serverTime: new Date().toISOString() });
      case 'ping':
        return json_({ ok: true, app: '今CoCo', time: new Date().toISOString() });
      default:
        return json_({ ok: false, error: '不明なactionです: ' + action });
    }
  } catch (err) {
    return json_({ ok: false, error: errMsg_(err) });
  }
}

function doPost(e) {
  try {
    const body = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    switch (body.action) {
      // 会員
      case 'register': return json_(register_(body));
      case 'login': return json_(login_(body));
      case 'forgotPassword': return json_(forgotPassword_(body));
      case 'changePassword': return json_(changePassword_(body));
      case 'updateNickname': return json_(updateNickname_(body));
      case 'me': return json_(me_(body));
      case 'logout': return json_(logout_(body));
      // 投稿
      case 'createEvent': return json_(createEvent_(body));
      case 'updateEvent': return json_(updateEvent_(body));
      case 'deleteEvent': return json_(deleteEvent_(body));
      case 'joinRequest': return json_(joinRequest_(body));
      case 'getMyData': return json_(getMyData_(body));
      default: return json_({ ok: false, error: '不明なactionです' });
    }
  } catch (err) {
    const res = { ok: false, error: errMsg_(err) };
    if (err && err.code) res.code = err.code;
    return json_(res);
  }
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
function errMsg_(err) { return String((err && err.message) || err); }
function apiError_(msg, code) { const e = new Error(msg); if (code) e.code = code; return e; }

/* =========================================================
 * シート共通
 * ========================================================= */
function ss_() {
  const id = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
  return id ? SpreadsheetApp.openById(id) : SpreadsheetApp.getActiveSpreadsheet();
}
function sheet_(name) {
  const sh = ss_().getSheetByName(name);
  if (!sh) throw new Error('シート「' + name + '」がありません。setup() を実行してください');
  return sh;
}
function headers_(sh) {
  return sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(String);
}
function readAll_(sh) {
  const values = sh.getDataRange().getValues();
  const headers = (values.shift() || []).map(String);
  return values.map((row, i) => {
    const o = { _row: i + 2 };
    headers.forEach((h, j) => { o[h] = row[j]; });
    return o;
  });
}
/** ヘッダー名に合わせて1行追加（列の並びが変わっても安全） */
function appendObj_(sh, obj) {
  const h = headers_(sh);
  sh.appendRow(h.map(k => (obj[k] === undefined || obj[k] === null) ? '' : obj[k]));
}
/** 指定行の複数フィールドを更新 */
function setFields_(sh, row, obj) {
  const h = headers_(sh);
  Object.keys(obj).forEach(k => {
    const c = h.indexOf(k);
    if (c < 0) throw new Error('列 ' + k + ' がありません。setup() を実行してください');
    sh.getRange(row, c + 1).setValue(obj[k]);
  });
}
function iso_(v) {
  if (v instanceof Date) return v.toISOString();
  return String(v || '');
}
function ms_(v) {
  if (v instanceof Date) return v.getTime();
  const t = Date.parse(String(v || ''));
  return isNaN(t) ? 0 : t;
}
function bool_(v) { return String(v).toUpperCase() === 'TRUE'; }
function clean_(v, max) {
  return String(v == null ? '' : v)
    .replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F\u007F]/g, '')
    .trim()
    .slice(0, max);
}
function withLock_(fn) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try { return fn(); } finally { lock.releaseLock(); }
}
function clearEventsCache_() {
  CacheService.getScriptCache().remove('events_v1');
}
function appUrl_() {
  return PropertiesService.getScriptProperties().getProperty('APP_URL') || CONFIG.APP_URL_DEFAULT;
}

/* =========================================================
 * 会員認証
 * ========================================================= */
function hash_(password, salt) {
  const pepper = PropertiesService.getScriptProperties().getProperty('PEPPER') || '';
  let h = salt + ':' + password + ':' + pepper;
  for (let i = 0; i < CONFIG.HASH_ROUNDS; i++) {
    const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, h, Utilities.Charset.UTF_8);
    h = bytes.map(b => ('0' + (b & 255).toString(16)).slice(-2)).join('');
  }
  return h;
}
function newToken_() {
  return (Utilities.getUuid() + Utilities.getUuid()).replace(/-/g, '');
}
function keyOf_(s) {
  return Utilities.base64EncodeWebSafe(String(s).toLowerCase());
}
function tempPassword_() {
  const chars = 'ABCDEFGHJKLMNPQRSTUVWXYZabcdefghijkmnpqrstuvwxyz23456789';
  let p = '';
  while (!/[A-Za-z]/.test(p) || !/\d/.test(p)) {
    p = '';
    for (let i = 0; i < 10; i++) p += chars.charAt(Math.floor(Math.random() * chars.length));
  }
  return p;
}
function normEmail_(v) { return clean_(v, 254).toLowerCase(); }
function validateEmail_(email) {
  if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) throw apiError_('メールアドレスの形式が正しくありません');
}
function validateNickname_(nick) {
  if (nick.length < 2 || nick.length > 20) throw apiError_('ニックネームは2〜20文字で入力してください');
  if (/[<>"'`\\]/.test(nick)) throw apiError_('ニックネームに使えない記号が含まれています');
}
function validatePassword_(pw) {
  if (pw.length < 8 || pw.length > 64) throw apiError_('パスワードは8〜64文字にしてください');
  if (!/[A-Za-z]/.test(pw) || !/\d/.test(pw)) throw apiError_('パスワードは英字と数字を両方含めてください');
}
function publicUser_(u) {
  return { user_id: String(u.user_id), nickname: String(u.nickname), email: String(u.email || '') };
}
function clearTokenCache_(token) {
  if (token) CacheService.getScriptCache().remove('tok_' + token);
}

function sendTempPasswordMail_(email, nickname, temp, isReset) {
  const subject = isReset
    ? '【今CoCo】仮パスワード再発行のお知らせ'
    : '【今CoCo】会員登録ありがとうございます（仮パスワードのお知らせ）';
  const body = [
    nickname + ' さん',
    '',
    isReset ? '仮パスワードを再発行しました。' : '今CoCo（いまココ）へのご登録ありがとうございます。',
    '',
    '　仮パスワード：' + temp,
    '',
    '次の手順でログインしてください。',
    '1. ' + appUrl_() + ' を開く',
    '2. 右上の「ログイン／登録」→「ログイン」',
    '3. メールアドレスと上の仮パスワードを入力',
    '4. 新しいパスワードを決めて保存',
    '',
    '※ 仮パスワードの有効期限は' + CONFIG.TEMP_PASSWORD_TTL_HOURS + '時間です。',
    '※ お心当たりのない場合は、このメールを破棄してください。',
    '',
    '――',
    '今CoCo（いまココ）',
    appUrl_(),
  ].join('\n');
  MailApp.sendEmail({ to: email, subject: subject, body: body, name: '今CoCo' });
}

/** 新規登録：ニックネーム＋メール → 仮パスワードをメール送信 */
function register_(b) {
  const nickname = clean_(b.nickname, 20);
  const email = normEmail_(b.email);
  validateNickname_(nickname);
  validateEmail_(email);
  if (b.agree !== true) throw apiError_('利用上の注意への同意が必要です');
  if (MailApp.getRemainingDailyQuota() < 1) throw apiError_('本日のメール送信数が上限に達しました。明日もう一度お試しください');

  const temp = tempPassword_();
  const created = withLock_(() => {
    const sh = sheet_(CONFIG.SHEET_USERS);
    const users = readAll_(sh);
    if (users.some(u => String(u.email).toLowerCase() === email)) {
      throw apiError_('このメールアドレスは登録済みです。「パスワードを忘れた方」から仮パスワードを再発行できます');
    }
    if (users.some(u => String(u.nickname).toLowerCase() === nickname.toLowerCase())) {
      throw apiError_('そのニックネームは使われています。別の名前にしてください');
    }
    const salt = Utilities.getUuid();
    const now = new Date().toISOString();
    const user = {
      user_id: Utilities.getUuid(), nickname: nickname, email: email,
      pass_hash: hash_(temp, salt), salt: salt, must_change: 'TRUE',
      token: '', token_expires: '', created_at: now, last_login: '', temp_issued_at: now,
    };
    appendObj_(sh, user);
    return user;
  });

  try {
    sendTempPasswordMail_(email, nickname, temp, false);
  } catch (err) {
    // メール送信に失敗したら登録を取り消す
    withLock_(() => {
      const sh = sheet_(CONFIG.SHEET_USERS);
      const u = readAll_(sh).find(x => String(x.user_id) === created.user_id);
      if (u) sh.deleteRow(u._row);
    });
    throw apiError_('メールを送信できませんでした。アドレスを確認してもう一度お試しください');
  }
  return { ok: true, email: email };
}

/** ログイン：メール＋パスワード（仮パスワードなら mustChange: true） */
function login_(b) {
  const email = normEmail_(b.email);
  const password = String(b.password || '');
  if (!email || !password) throw apiError_('メールアドレスとパスワードを入力してください');

  const cache = CacheService.getScriptCache();
  const failKey = 'fail_' + keyOf_(email);
  const fails = Number(cache.get(failKey) || 0);
  if (fails >= CONFIG.LOGIN_MAX_FAIL) throw apiError_('ログインに5回失敗したため、15分間ロックしています。時間をおいてお試しください');

  return withLock_(() => {
    const sh = sheet_(CONFIG.SHEET_USERS);
    const u = readAll_(sh).find(x => String(x.email).toLowerCase() === email);
    if (!u || hash_(password, String(u.salt)) !== String(u.pass_hash)) {
      cache.put(failKey, String(fails + 1), CONFIG.LOGIN_LOCK_SEC);
      const left = Math.max(0, CONFIG.LOGIN_MAX_FAIL - fails - 1);
      throw apiError_(left > 0
        ? 'メールアドレスかパスワードが違います（あと' + left + '回でロック）'
        : 'ログインに5回失敗したため、15分間ロックしました');
    }
    const mustChange = bool_(u.must_change);
    if (mustChange && ms_(u.temp_issued_at) + CONFIG.TEMP_PASSWORD_TTL_HOURS * 3600000 < Date.now()) {
      throw apiError_('仮パスワードの有効期限が切れています。「パスワードを忘れた方」から再発行してください');
    }
    cache.remove(failKey);

    const now = new Date();
    let token = String(u.token || '');
    if (!token || ms_(u.token_expires) < now.getTime()) token = newToken_();
    setFields_(sh, u._row, {
      token: token,
      token_expires: new Date(now.getTime() + CONFIG.TOKEN_TTL_DAYS * 86400000).toISOString(),
      last_login: now.toISOString(),
    });
    clearTokenCache_(token);
    return { ok: true, user: publicUser_(u), token: token, mustChange: mustChange };
  });
}

/** パスワード忘れ：仮パスワードを再発行（登録の有無は回答しない） */
function forgotPassword_(b) {
  const email = normEmail_(b.email);
  validateEmail_(email);
  const cache = CacheService.getScriptCache();
  const rk = 'forgot_' + keyOf_(email);
  if (cache.get(rk)) throw apiError_('少し前に送信しました。届かない場合は3分ほど待ってから再度お試しください');
  if (MailApp.getRemainingDailyQuota() < 1) throw apiError_('本日のメール送信数が上限に達しました。明日もう一度お試しください');

  const temp = tempPassword_();
  const target = withLock_(() => {
    const sh = sheet_(CONFIG.SHEET_USERS);
    const u = readAll_(sh).find(x => String(x.email).toLowerCase() === email);
    if (!u) return null;
    const salt = Utilities.getUuid();
    clearTokenCache_(String(u.token || ''));
    setFields_(sh, u._row, {
      pass_hash: hash_(temp, salt), salt: salt, must_change: 'TRUE',
      temp_issued_at: new Date().toISOString(),
      token: '', token_expires: '', // 他の端末のログインも解除
    });
    return { email: String(u.email), nickname: String(u.nickname) };
  });

  cache.put(rk, '1', CONFIG.FORGOT_INTERVAL_SEC);
  cache.remove('fail_' + keyOf_(email));
  if (target) {
    try { sendTempPasswordMail_(target.email, target.nickname, temp, true); }
    catch (err) { throw apiError_('メールを送信できませんでした。時間をおいてお試しください'); }
  }
  return { ok: true };
}

/**
 * トークンから会員を取得
 * opts.allowMustChange: 仮パスワードのままでも通す（パスワード変更・me用）
 */
function authUser_(token, opts) {
  token = String(token || '');
  if (!token) throw apiError_('ログインが必要です', 'AUTH');
  const cache = CacheService.getScriptCache();
  let user = null;
  const hit = cache.get('tok_' + token);
  if (hit) {
    user = JSON.parse(hit);
  } else {
    const u = readAll_(sheet_(CONFIG.SHEET_USERS)).find(x => String(x.token) === token);
    if (!u || ms_(u.token_expires) < Date.now()) throw apiError_('ログインの有効期限が切れました。もう一度ログインしてください', 'AUTH');
    user = publicUser_(u);
    user.must_change = bool_(u.must_change);
    cache.put('tok_' + token, JSON.stringify(user), 600);
  }
  if (user.must_change && !(opts && opts.allowMustChange)) {
    throw apiError_('はじめに新しいパスワードを設定してください', 'MUST_CHANGE');
  }
  return user;
}

function me_(b) {
  const user = authUser_(b.token, { allowMustChange: true });
  return { ok: true, user: { user_id: user.user_id, nickname: user.nickname, email: user.email }, mustChange: !!user.must_change };
}

/** パスワード変更（初回の仮パスワード更新もこれ） */
function changePassword_(b) {
  const token = String(b.token || '');
  const user = authUser_(token, { allowMustChange: true });
  const current = String(b.currentPassword || '');
  const next = String(b.newPassword || '');
  validatePassword_(next);
  if (current === next) throw apiError_('今と違うパスワードにしてください');

  return withLock_(() => {
    const sh = sheet_(CONFIG.SHEET_USERS);
    const u = readAll_(sh).find(x => String(x.user_id) === user.user_id);
    if (!u) throw apiError_('会員情報が見つかりません', 'AUTH');
    if (hash_(current, String(u.salt)) !== String(u.pass_hash)) throw apiError_('現在のパスワード（仮パスワード）が違います');
    const salt = Utilities.getUuid();
    setFields_(sh, u._row, { pass_hash: hash_(next, salt), salt: salt, must_change: 'FALSE', temp_issued_at: '' });
    clearTokenCache_(token);
    return { ok: true, user: publicUser_(u) };
  });
}

function updateNickname_(b) {
  const token = String(b.token || '');
  const user = authUser_(token);
  const nickname = clean_(b.nickname, 20);
  validateNickname_(nickname);
  return withLock_(() => {
    const sh = sheet_(CONFIG.SHEET_USERS);
    const users = readAll_(sh);
    if (users.some(x => String(x.user_id) !== user.user_id && String(x.nickname).toLowerCase() === nickname.toLowerCase())) {
      throw apiError_('そのニックネームは使われています。別の名前にしてください');
    }
    const u = users.find(x => String(x.user_id) === user.user_id);
    setFields_(sh, u._row, { nickname: nickname });
    clearTokenCache_(token);
    clearEventsCache_();
    return { ok: true, user: { user_id: user.user_id, nickname: nickname, email: user.email } };
  });
}

function logout_(b) {
  const token = String(b.token || '');
  if (!token) return { ok: true };
  clearTokenCache_(token);
  return withLock_(() => {
    const sh = sheet_(CONFIG.SHEET_USERS);
    const u = readAll_(sh).find(x => String(x.token) === token);
    if (u) setFields_(sh, u._row, { token: '', token_expires: '' });
    return { ok: true };
  });
}

function userMap_() {
  const map = {};
  readAll_(sheet_(CONFIG.SHEET_USERS)).forEach(u => { map[String(u.user_id)] = { nickname: String(u.nickname) }; });
  return map;
}

/* =========================================================
 * イベント取得
 * ========================================================= */
function getEvents_() {
  const cache = CacheService.getScriptCache();
  const cached = cache.get('events_v1');
  if (cached) return JSON.parse(cached);

  const now = Date.now();
  const users = userMap_();
  const counts = requestCounts_();
  const list = readAll_(sheet_(CONFIG.SHEET_EVENTS))
    .filter(r => r.id && String(r.status) === 'active' && ms_(r.expires_at) > now)
    .map(r => publicEvent_(r, users, counts));

  try { cache.put('events_v1', JSON.stringify(list), CONFIG.EVENTS_CACHE_SEC); } catch (_) { /* 100KB超はキャッシュしない */ }
  return list;
}

function requestCounts_() {
  const counts = {};
  readAll_(sheet_(CONFIG.SHEET_REQUESTS)).forEach(r => {
    const k = String(r.event_id);
    counts[k] = (counts[k] || 0) + 1;
  });
  return counts;
}

function publicEvent_(r, users, counts) {
  const uid = String(r.user_id);
  return {
    id: String(r.id),
    created_at: iso_(r.created_at),
    expires_at: iso_(r.expires_at),
    title: String(r.title),
    category: String(r.category),
    icon: String(r.icon),
    summary: String(r.summary),
    description: String(r.description),
    lat: Number(r.lat),
    lng: Number(r.lng),
    media_url: String(r.media_url || ''),
    media_type: String(r.media_type || ''),
    capacity: String(r.capacity || ''),
    fee: String(r.fee || ''),
    conditions: String(r.conditions || ''),
    user_id: uid,
    nickname: (users && users[uid] && users[uid].nickname) || '名無しさん',
    status: String(r.status),
    hashtags: String(r.hashtags || '').split(/\s+/).filter(Boolean),
    request_count: (counts && counts[String(r.id)]) || 0,
  };
}

/* =========================================================
 * 投稿作成（Geminiチェック → Driveアップロード → シート追記）
 * ========================================================= */
/** 投稿内容の入力チェック（新規・修正で共通） */
function parseEventInput_(b) {
  const now = Date.now();
  const p = {
    title: clean_(b.title, 40),
    description: clean_(b.description, 500),
    category: CATEGORIES[b.category] ? String(b.category) : 'other',
    capacity: clean_(b.capacity, 20) || '何人でも',
    fee: clean_(b.fee, 20) || '無料',
    conditions: clean_(b.conditions, 100) || '誰でも歓迎',
    lat: Number(b.lat),
    lng: Number(b.lng),
    expMs: Date.parse(String(b.expires_at || '')),
  };
  if (!p.title) throw apiError_('タイトルを入力してください');
  if (!isFinite(p.lat) || !isFinite(p.lng) || Math.abs(p.lat) > 90 || Math.abs(p.lng) > 180) throw apiError_('場所が正しくありません');
  if (isNaN(p.expMs) || p.expMs <= now + 5 * 60000) throw apiError_('終了時刻は5分以上先にしてください');
  if (p.expMs > now + CONFIG.MAX_EXPIRE_HOURS * 3600000) throw apiError_('終了時刻は' + CONFIG.MAX_EXPIRE_HOURS + '時間以内にしてください');
  return p;
}

function mediaInput_(b) {
  const media = (b.media && b.media.data) ? b.media : null;
  if (media) {
    const approxBytes = Math.floor(String(media.data).length * 3 / 4);
    if (approxBytes > CONFIG.MAX_MEDIA_BYTES) throw apiError_('ファイルが大きすぎます（20MBまで）');
  }
  return media;
}

function createEvent_(b) {
  const user = authUser_(b.token);
  const p = parseEventInput_(b);
  const { title, description, category, capacity, fee, conditions, lat, lng, expMs } = p;
  const now = Date.now();

  // 同時公開数の上限
  const mine = readAll_(sheet_(CONFIG.SHEET_EVENTS))
    .filter(r => String(r.user_id) === user.user_id && String(r.status) === 'active' && ms_(r.expires_at) > now);
  if (mine.length >= CONFIG.MAX_ACTIVE_POSTS_PER_USER) {
    throw apiError_('同時に公開できるのは' + CONFIG.MAX_ACTIVE_POSTS_PER_USER + '件までです。終わった投稿を削除してください');
  }

  const media = mediaInput_(b);

  // 1) Gemini モデレーション & 生成
  const ai = geminiReview_({ title, description, category, capacity, fee, conditions, media });
  if (!ai.safe) {
    return { ok: false, moderation: true, error: 'この内容は投稿できません：' + (ai.reason || 'ガイドラインに抵触する可能性があります') };
  }

  // 2) Drive アップロード
  const id = Utilities.getUuid();
  let up = { url: '', type: '', id: '' };
  if (media) up = uploadMedia_(media, id);

  // 3) シート追記（失敗時はアップロードを取り消し）
  const row = {
    id: id,
    created_at: new Date(now).toISOString(),
    expires_at: new Date(expMs).toISOString(),
    title: title,
    category: category,
    icon: ai.emoji || CATEGORIES[category].icon,
    summary: ai.summary || title.slice(0, 15),
    description: description,
    lat: lat.toFixed(6),
    lng: lng.toFixed(6),
    media_url: up.url,
    media_type: up.type,
    capacity: capacity,
    fee: fee,
    conditions: conditions,
    user_id: user.user_id,
    status: 'active',
    hashtags: (ai.hashtags || []).join(' '),
    media_file_id: up.id,
  };
  try {
    withLock_(() => { appendObj_(sheet_(CONFIG.SHEET_EVENTS), row); });
  } catch (err) {
    if (up.id) { try { DriveApp.getFileById(up.id).setTrashed(true); } catch (_) {} }
    throw apiError_('保存に失敗しました。もう一度お試しください');
  }
  clearEventsCache_();
  return { ok: true, event: publicEvent_(row, { [user.user_id]: { nickname: user.nickname } }, {}) };
}

/* =========================================================
 * Gemini API
 * ========================================================= */
function geminiReview_(p) {
  const props = PropertiesService.getScriptProperties();
  const key = props.getProperty('GEMINI_API_KEY');
  if (!key) throw apiError_('サーバー設定エラー：GEMINI_API_KEY が未設定です');
  const model = props.getProperty('GEMINI_MODEL') || CONFIG.GEMINI_MODEL_DEFAULT;
  const failOpen = props.getProperty('MODERATION_FAIL_OPEN') === 'true';
  const fallback = { safe: true, reason: '', emoji: '', summary: '', hashtags: [] };

  const post = {
    タイトル: p.title, カテゴリ: (CATEGORIES[p.category] || {}).label,
    本文: p.description, 募集人数: p.capacity, 参加費: p.fee, 参加条件: p.conditions,
  };
  const prompt = [
    'あなたは、今いる場所で遊んでいる様子を地図で共有し、仲間を募る公開アプリ「今CoCo」のモデレーター兼コピーライターです。',
    '次の投稿（添付画像がある場合は画像も）を確認してください。',
    '',
    '【投稿不可（safe=false）とする内容】',
    '暴力・脅迫・危険行為の誘い／違法行為（薬物・違法賭博・無許可営業など）／性的な内容や性的目的の勧誘／',
    '差別・ヘイト・誹謗中傷／他人の個人情報や顔写真の晒し／詐欺・マルチ商法・宗教や商品の強引な勧誘／',
    '未成年への飲酒・喫煙の誘い／金銭目的の出会い勧誘／スパム・意味のない文字列。',
    '成人の飲み会、アウトドア、スポーツ、趣味の集まりなど通常の遊びは safe=true です。迷う場合は safe=true にしてください。',
    '',
    '【safe=true の場合に生成するもの】',
    'emoji: 投稿内容に最も合う絵文字1つ',
    'summary: 地図ピンに表示するワクワクするキャッチコピー（日本語15文字以内、句読点控えめ）',
    'hashtags: 内容に合うハッシュタグを2〜4個（#から始める、日本語可、空白なし）',
    'reason は safe=false の場合のみ、投稿者に伝える短い理由（30文字以内）。',
    '',
    '【投稿】',
    JSON.stringify(post),
  ].join('\n');

  const parts = [{ text: prompt }];
  if (p.media && p.media.data && /^image\//.test(String(p.media.mimeType))) {
    parts.push({ inline_data: { mime_type: String(p.media.mimeType), data: String(p.media.data) } });
  }

  const payload = {
    contents: [{ role: 'user', parts: parts }],
    generationConfig: {
      temperature: 0.4,
      responseMimeType: 'application/json',
      responseSchema: {
        type: 'OBJECT',
        properties: {
          safe: { type: 'BOOLEAN' },
          reason: { type: 'STRING' },
          emoji: { type: 'STRING' },
          summary: { type: 'STRING' },
          hashtags: { type: 'ARRAY', items: { type: 'STRING' } },
        },
        required: ['safe', 'emoji', 'summary', 'hashtags'],
      },
    },
  };

  let res;
  try {
    res = UrlFetchApp.fetch('https://generativelanguage.googleapis.com/v1beta/models/' + encodeURIComponent(model) + ':generateContent', {
      method: 'post',
      contentType: 'application/json',
      headers: { 'x-goog-api-key': key },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true,
    });
  } catch (err) {
    if (failOpen) return fallback;
    throw apiError_('AIチェックに接続できませんでした。少し待って再投稿してください');
  }

  const code = res.getResponseCode();
  let data = {};
  try { data = JSON.parse(res.getContentText() || '{}'); } catch (_) {}
  if (code !== 200) {
    console.error('Gemini error', code, res.getContentText());
    if (failOpen) return fallback;
    if (code === 429) throw apiError_('AIチェックが混み合っています。1分ほど待って再投稿してください');
    throw apiError_('AIチェックに失敗しました（' + code + '）。少し待って再投稿してください');
  }

  if (data.promptFeedback && data.promptFeedback.blockReason) {
    return { safe: false, reason: 'AIの安全フィルタにより投稿できません' };
  }
  const cand = data.candidates && data.candidates[0];
  if (!cand || cand.finishReason === 'SAFETY' || cand.finishReason === 'PROHIBITED_CONTENT') {
    return { safe: false, reason: 'AIの安全フィルタにより投稿できません' };
  }
  const text = ((cand.content && cand.content.parts) || []).map(x => x.text || '').join('');
  let r;
  try {
    r = JSON.parse(text.replace(/```json|```/g, '').trim());
  } catch (_) {
    if (failOpen) return fallback;
    throw apiError_('AIの応答を読み取れませんでした。もう一度お試しください');
  }

  let emoji = clean_(r.emoji, 8);
  if (!emoji || /[A-Za-z0-9]/.test(emoji)) emoji = '';
  const hashtags = (Array.isArray(r.hashtags) ? r.hashtags : [])
    .map(t => clean_(t, 20).replace(/\s+/g, ''))
    .filter(Boolean)
    .map(t => (t.charAt(0) === '#' || t.charAt(0) === '＃') ? '#' + t.slice(1) : '#' + t)
    .slice(0, 4);

  return {
    safe: r.safe !== false,
    reason: clean_(r.reason, 60),
    emoji: emoji,
    summary: clean_(r.summary, 15),
    hashtags: hashtags,
  };
}

/* =========================================================
 * Google Drive アップロード
 * ========================================================= */
function uploadMedia_(m, eventId) {
  const mime = String(m.mimeType || '').toLowerCase();
  const isImage = /^image\/(jpeg|png|webp|gif)$/.test(mime);
  const isVideo = /^video\/(mp4|webm|quicktime)$/.test(mime);
  if (!isImage && !isVideo) throw apiError_('対応していないファイル形式です（JPEG/PNG/WebP/GIF・MP4/WebM/MOV）');

  const bytes = Utilities.base64Decode(String(m.data));
  if (bytes.length > CONFIG.MAX_MEDIA_BYTES) throw apiError_('ファイルが大きすぎます（20MBまで）');

  const extMap = {
    'image/jpeg': 'jpg', 'image/png': 'png', 'image/webp': 'webp', 'image/gif': 'gif',
    'video/mp4': 'mp4', 'video/webm': 'webm', 'video/quicktime': 'mov',
  };
  const blob = Utilities.newBlob(bytes, mime, eventId + '.' + extMap[mime]);
  const folderId = PropertiesService.getScriptProperties().getProperty('DRIVE_FOLDER_ID');
  if (!folderId) throw apiError_('サーバー設定エラー：DRIVE_FOLDER_ID が未設定です');

  const file = DriveApp.getFolderById(folderId).createFile(blob);
  try {
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  } catch (err) {
    file.setTrashed(true);
    throw apiError_('ファイルの公開設定に失敗しました（Driveの共有制限を確認してください）');
  }
  const fid = file.getId();
  return {
    id: fid,
    type: isImage ? 'image' : 'video',
    // 画像は直リンク、動画はDriveプレーヤー（iframe）で再生
    url: isImage ? 'https://lh3.googleusercontent.com/d/' + fid : 'https://drive.google.com/file/d/' + fid + '/preview',
  };
}

function trashMedia_(fileId) {
  if (!fileId) return;
  try { DriveApp.getFileById(String(fileId)).setTrashed(true); } catch (_) {}
}

/* =========================================================
 * 投稿の修正（本人のみ・公開中のみ）
 *   media: 新しい写真/動画（差し替え）
 *   removeMedia: true で写真/動画を外す
 *   どちらも無ければ今の写真/動画をそのまま使う
 * ========================================================= */
function updateEvent_(b) {
  const user = authUser_(b.token);
  const id = String(b.id || '');
  const findEv_ = () => readAll_(sheet_(CONFIG.SHEET_EVENTS)).find(r => String(r.id) === id);
  const checkOwner_ = ev => {
    if (!ev) throw apiError_('投稿が見つかりません');
    if (String(ev.user_id) !== user.user_id) throw apiError_('自分の投稿だけ修正できます');
    if (String(ev.status) !== 'active' || ms_(ev.expires_at) <= Date.now()) throw apiError_('終了した投稿は修正できません');
  };
  checkOwner_(findEv_());

  const p = parseEventInput_(b);
  const media = mediaInput_(b);
  const removeMedia = b.removeMedia === true && !media;

  // 1) Gemini で修正後の内容をチェック（キャッチコピー・絵文字・タグも作り直し）
  const ai = geminiReview_({
    title: p.title, description: p.description, category: p.category,
    capacity: p.capacity, fee: p.fee, conditions: p.conditions, media: media,
  });
  if (!ai.safe) {
    return { ok: false, moderation: true, error: 'この内容には修正できません：' + (ai.reason || 'ガイドラインに抵触する可能性があります') };
  }

  // 2) 新しい写真/動画のアップロード
  let up = null;
  if (media) up = uploadMedia_(media, id + '_' + Date.now());

  // 3) シート更新
  let oldFileId = '';
  let saved;
  try {
    saved = withLock_(() => {
      const sh = sheet_(CONFIG.SHEET_EVENTS);
      const ev = findEv_();
      checkOwner_(ev);
      oldFileId = String(ev.media_file_id || '');
      const fields = {
        title: p.title,
        description: p.description,
        category: p.category,
        icon: ai.emoji || CATEGORIES[p.category].icon,
        summary: ai.summary || p.title.slice(0, 15),
        lat: p.lat.toFixed(6),
        lng: p.lng.toFixed(6),
        expires_at: new Date(p.expMs).toISOString(),
        capacity: p.capacity,
        fee: p.fee,
        conditions: p.conditions,
        hashtags: (ai.hashtags || []).join(' '),
      };
      if (up) {
        fields.media_url = up.url; fields.media_type = up.type; fields.media_file_id = up.id;
      } else if (removeMedia) {
        fields.media_url = ''; fields.media_type = ''; fields.media_file_id = '';
      }
      setFields_(sh, ev._row, fields);
      return Object.assign({}, ev, fields);
    });
  } catch (err) {
    if (up) trashMedia_(up.id);
    throw err;
  }

  // 差し替え・削除した古いファイルはゴミ箱へ
  if ((up || removeMedia) && oldFileId) trashMedia_(oldFileId);
  clearEventsCache_();
  return { ok: true, event: publicEvent_(saved, { [user.user_id]: { nickname: user.nickname } }, requestCounts_()) };
}

/* =========================================================
 * 投稿削除（早期終了）
 * ========================================================= */
function deleteEvent_(b) {
  const user = authUser_(b.token);
  const id = String(b.id || '');
  const fileId = withLock_(() => {
    const sh = sheet_(CONFIG.SHEET_EVENTS);
    const ev = readAll_(sh).find(r => String(r.id) === id);
    if (!ev) throw apiError_('投稿が見つかりません');
    if (String(ev.user_id) !== user.user_id) throw apiError_('自分の投稿だけ削除できます');
    if (String(ev.status) !== 'active') throw apiError_('この投稿はすでに終了しています');
    setFields_(sh, ev._row, { status: 'cancelled' });
    return String(ev.media_file_id || '');
  });
  trashMedia_(fileId);
  clearEventsCache_();
  return { ok: true };
}

/* =========================================================
 * 参加リクエスト
 * ========================================================= */
function joinRequest_(b) {
  const user = authUser_(b.token);
  const eventId = String(b.id || '');
  const message = clean_(b.message, 140) || '参加したいです！';
  return withLock_(() => {
    const ev = readAll_(sheet_(CONFIG.SHEET_EVENTS)).find(r => String(r.id) === eventId);
    if (!ev || String(ev.status) !== 'active' || ms_(ev.expires_at) <= Date.now()) throw apiError_('この募集は終了しています');
    if (String(ev.user_id) === user.user_id) throw apiError_('自分の投稿には参加リクエストできません');
    const sh = sheet_(CONFIG.SHEET_REQUESTS);
    const dup = readAll_(sh).some(r => String(r.event_id) === eventId && String(r.user_id) === user.user_id);
    if (dup) throw apiError_('すでに参加リクエスト済みです');
    appendObj_(sh, {
      id: Utilities.getUuid(), event_id: eventId, user_id: user.user_id,
      nickname: user.nickname, message: message, created_at: new Date().toISOString(),
    });
    clearEventsCache_();
    return { ok: true };
  });
}

/* =========================================================
 * マイページ
 * ========================================================= */
function getMyData_(b) {
  const user = authUser_(b.token);
  const now = Date.now();
  const events = readAll_(sheet_(CONFIG.SHEET_EVENTS));
  const requests = readAll_(sheet_(CONFIG.SHEET_REQUESTS));
  const counts = {};
  requests.forEach(r => { counts[String(r.event_id)] = (counts[String(r.event_id)] || 0) + 1; });

  const myEvents = events
    .filter(r => String(r.user_id) === user.user_id && String(r.status) === 'active' && ms_(r.expires_at) > now)
    .map(r => {
      const ev = publicEvent_(r, { [user.user_id]: { nickname: user.nickname } }, counts);
      ev.requests = requests
        .filter(q => String(q.event_id) === ev.id)
        .map(q => ({ nickname: String(q.nickname), message: String(q.message), created_at: iso_(q.created_at) }));
      return ev;
    });

  const evMap = {};
  events.forEach(r => { evMap[String(r.id)] = r; });
  const sent = requests
    .filter(q => String(q.user_id) === user.user_id)
    .map(q => {
      const ev = evMap[String(q.event_id)];
      const alive = ev && String(ev.status) === 'active' && ms_(ev.expires_at) > now;
      return {
        event_id: String(q.event_id),
        title: ev ? String(ev.title) : '（削除された投稿）',
        icon: ev ? String(ev.icon) : '📍',
        active: !!alive,
        created_at: iso_(q.created_at),
      };
    })
    .sort((a, b2) => (a.created_at < b2.created_at ? 1 : -1))
    .slice(0, 30);

  return {
    ok: true,
    user: { user_id: user.user_id, nickname: user.nickname, email: user.email },
    myEvents: myEvents,
    sentRequests: sent,
  };
}

/* =========================================================
 * 自動消滅（時間主導トリガー：10分おき）
 * ========================================================= */
function expireEvents() {
  const trash = PropertiesService.getScriptProperties().getProperty('TRASH_MEDIA_ON_EXPIRE') !== 'false';
  const toTrash = [];
  let changed = 0;
  withLock_(() => {
    const sh = sheet_(CONFIG.SHEET_EVENTS);
    const values = sh.getDataRange().getValues();
    if (values.length < 2) return;
    const h = values[0].map(String);
    const cStatus = h.indexOf('status'), cExp = h.indexOf('expires_at'), cFile = h.indexOf('media_file_id');
    const now = Date.now();
    for (let i = 1; i < values.length; i++) {
      if (String(values[i][cStatus]) === 'active' && ms_(values[i][cExp]) <= now) {
        sh.getRange(i + 1, cStatus + 1).setValue('expired');
        changed++;
        if (trash && cFile >= 0 && values[i][cFile]) toTrash.push(values[i][cFile]);
      }
    }
  });
  toTrash.forEach(trashMedia_);
  if (changed) clearEventsCache_();
  Logger.log('expired: ' + changed);
}
