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
 * 初回・列追加時: setup() を実行 → ウェブアプリとしてデプロイ（新しいバージョン）
 */

const CONFIG = {
  SHEET_EVENTS: 'events',
  SHEET_USERS: 'users',
  SHEET_PARTICIPANTS: 'participants',
  GEMINI_MODEL_DEFAULT: 'gemini-3.8-flash',
  APP_URL_DEFAULT: 'https://kenken6291.github.io/ima-coco/',
  MAX_ACTIVE_POSTS_PER_USER: 5,
  MAX_MEDIA_BYTES: 20 * 1024 * 1024,
  MAX_DAYS_AHEAD: 31,           // 何日先のイベントまで登録できるか
  MAX_EVENT_DAYS: 14,           // 1イベントの最長開催期間
  MAX_PEOPLE: 20,               // 参加表明1件あたりの最大人数
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
  'hashtags', 'media_file_id',
  'start_at', 'end_at', 'address', 'pref', 'area', 'sns', 'join_method'
];
const USER_HEADERS = [
  'user_id', 'nickname', 'email', 'pass_hash', 'salt', 'must_change',
  'token', 'token_expires', 'created_at', 'last_login', 'temp_issued_at'
];
const PARTICIPANT_HEADERS = [
  'id', 'event_id', 'user_id', 'nickname', 'people', 'message', 'added_by', 'created_at', 'updated_at'
];

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
  goukon: { label: '合コン', icon: '💘' },
  nomikai: { label: '飲み会', icon: '🍻' },
  dance: { label: 'ダンス', icon: '💃' },
  offkai: { label: 'オフ会', icon: '🙌' },
  other: { label: 'その他', icon: '📍' },
};

const PREFS = ['',
  '北海道', '青森県', '岩手県', '宮城県', '秋田県', '山形県', '福島県',
  '茨城県', '栃木県', '群馬県', '埼玉県', '千葉県', '東京都', '神奈川県',
  '新潟県', '富山県', '石川県', '福井県', '山梨県', '長野県', '岐阜県', '静岡県', '愛知県',
  '三重県', '滋賀県', '京都府', '大阪府', '兵庫県', '奈良県', '和歌山県',
  '鳥取県', '島根県', '岡山県', '広島県', '山口県',
  '徳島県', '香川県', '愛媛県', '高知県',
  '福岡県', '佐賀県', '長崎県', '熊本県', '大分県', '宮崎県', '鹿児島県', '沖縄県'];

function areaOfPrefCode_(n) {
  if (n === 1) return '北海道';
  if (n <= 7) return '東北';
  if (n <= 14) return '関東';
  if (n <= 23) return '中部';
  if (n <= 30) return '近畿';
  if (n <= 35) return '中国';
  if (n <= 39) return '四国';
  if (n <= 47) return '九州・沖縄';
  return 'その他';
}

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
  ensureSheet_(ss, CONFIG.SHEET_PARTICIPANTS, PARTICIPANT_HEADERS);

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
      case 'geocode':
        return json_({ ok: true, results: geocode_(p.q) });
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
      // 参加表明
      case 'getParticipants': return json_(getParticipants_(body));
      case 'saveParticipant': return json_(saveParticipant_(body));
      case 'deleteParticipant': return json_(deleteParticipant_(body));
      // マイページ
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
    // 自分の参加表明の表示名もそろえる
    const ps = sheet_(CONFIG.SHEET_PARTICIPANTS);
    readAll_(ps).filter(r => String(r.user_id) === user.user_id).forEach(r => setFields_(ps, r._row, { nickname: nickname }));
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
  const stats = participantStats_();
  const list = readAll_(sheet_(CONFIG.SHEET_EVENTS))
    .filter(r => r.id && String(r.status) === 'active' && ms_(r.expires_at) > now)
    .map(r => publicEvent_(r, users, stats));

  try { cache.put('events_v1', JSON.stringify(list), CONFIG.EVENTS_CACHE_SEC); } catch (_) { /* 100KB超はキャッシュしない */ }
  return list;
}

/** イベントごとの参加表明数 { event_id: { n: 件数, people: 人数合計 } } */
function participantStats_(rows) {
  const stats = {};
  (rows || readAll_(sheet_(CONFIG.SHEET_PARTICIPANTS))).forEach(r => {
    const k = String(r.event_id);
    if (!stats[k]) stats[k] = { n: 0, people: 0 };
    stats[k].n += 1;
    stats[k].people += Number(r.people) || 1;
  });
  return stats;
}

function parseJson_(v, fallback) {
  try { return v ? JSON.parse(String(v)) : fallback; } catch (_) { return fallback; }
}

function publicEvent_(r, users, stats) {
  const uid = String(r.user_id);
  const st = (stats && stats[String(r.id)]) || { n: 0, people: 0 };
  return {
    id: String(r.id),
    created_at: iso_(r.created_at),
    expires_at: iso_(r.expires_at),
    start_at: iso_(r.start_at) || iso_(r.created_at),
    end_at: iso_(r.end_at) || iso_(r.expires_at),
    title: String(r.title),
    category: String(r.category),
    icon: String(r.icon),
    summary: String(r.summary),
    description: String(r.description),
    lat: Number(r.lat),
    lng: Number(r.lng),
    address: String(r.address || ''),
    pref: String(r.pref || ''),
    area: String(r.area || '') || 'その他',
    media_url: String(r.media_url || ''),
    media_type: String(r.media_type || ''),
    capacity: String(r.capacity || ''),
    fee: String(r.fee || ''),
    conditions: String(r.conditions || ''),
    sns: parseJson_(r.sns, []),
    join_method: String(r.join_method || '') === 'direct' ? 'direct' : 'app',
    user_id: uid,
    nickname: (users && users[uid] && users[uid].nickname) || '名無しさん',
    status: String(r.status),
    hashtags: String(r.hashtags || '').split(/\s+/).filter(Boolean),
    participant_n: st.n,
    participant_count: st.people,
  };
}

/* =========================================================
 * 住所 ⇔ 緯度経度（国土地理院API）
 * ========================================================= */
/** 住所検索（クライアントから直接呼べない場合の予備） */
function geocode_(q) {
  q = clean_(q, 100);
  if (!q) throw apiError_('住所を入力してください');
  const res = UrlFetchApp.fetch('https://msearch.gsi.go.jp/address-search/AddressSearch?q=' + encodeURIComponent(q), { muteHttpExceptions: true });
  if (res.getResponseCode() !== 200) throw apiError_('住所検索に失敗しました');
  const arr = parseJson_(res.getContentText(), []) || [];
  return arr.slice(0, 5).map(x => ({
    title: String((x.properties && x.properties.title) || ''),
    lat: Number(x.geometry.coordinates[1]),
    lng: Number(x.geometry.coordinates[0]),
  }));
}

/** 緯度経度 → 都道府県・エリア・町名 */
function reverseGeo_(lat, lng) {
  const out = { pref: '', area: 'その他', town: '' };
  const ck = 'rg_' + lat.toFixed(3) + '_' + lng.toFixed(3);
  const cache = CacheService.getScriptCache();
  const hit = cache.get(ck);
  if (hit) return JSON.parse(hit);
  try {
    const res = UrlFetchApp.fetch('https://mreversegeocoder.gsi.go.jp/reverse-geocoder/LonLatToAddress?lat=' + lat + '&lon=' + lng, { muteHttpExceptions: true });
    if (res.getResponseCode() !== 200) return out;
    const r = (parseJson_(res.getContentText(), {}) || {}).results;
    if (r && r.muniCd) {
      const code = Number(String(r.muniCd).padStart(5, '0').slice(0, 2));
      if (code >= 1 && code <= 47) { out.pref = PREFS[code]; out.area = areaOfPrefCode_(code); }
      out.town = String(r.lv01Nm || '').replace(/^[-－]$/, '');
    }
    cache.put(ck, JSON.stringify(out), 21600);
  } catch (_) {}
  return out;
}

/* =========================================================
 * 連絡先SNS（任意）
 * ========================================================= */
function normalizeSns_(s) {
  s = s || {};
  const out = [];
  const handle = (v, re) => String(v).replace(re, '').replace(/^@/, '').split(/[/?#]/)[0];

  const x = clean_(s.x, 200);
  if (x) {
    const h = handle(x, /^https?:\/\/(www\.)?(x|twitter)\.com\//i);
    if (!/^[A-Za-z0-9_]{1,15}$/.test(h)) throw apiError_('XのユーザーIDが正しくありません（英数字と_、15文字まで）');
    out.push({ type: 'x', label: '@' + h, url: 'https://x.com/' + h });
  }
  const ig = clean_(s.instagram, 200);
  if (ig) {
    const h = handle(ig, /^https?:\/\/(www\.)?instagram\.com\//i);
    if (!/^[A-Za-z0-9_.]{1,30}$/.test(h)) throw apiError_('InstagramのユーザーIDが正しくありません');
    out.push({ type: 'instagram', label: '@' + h, url: 'https://www.instagram.com/' + h + '/' });
  }
  const line = clean_(s.line, 200);
  if (line) {
    if (!/^https:\/\/(line\.me|lin\.ee|page\.line\.me|liff\.line\.me)\/[^\s<>"]*$/i.test(line)) throw apiError_('LINEは https://line.me/… などのURLを入力してください');
    out.push({ type: 'line', label: 'LINE', url: line });
  }
  const fb = clean_(s.facebook, 200);
  if (fb) {
    if (!/^https:\/\/(www\.|m\.)?(facebook\.com|fb\.me)\/[^\s<>"]+$/i.test(fb)) throw apiError_('Facebookは https://www.facebook.com/… のURLを入力してください');
    out.push({ type: 'facebook', label: 'Facebook', url: fb });
  }
  const other = clean_(s.other, 200);
  if (other) {
    if (!/^https:\/\/[^\s<>"]+\.[^\s<>"]+$/i.test(other)) throw apiError_('その他の連絡先は https:// から始まるURLを入力してください');
    const host = (other.match(/^https:\/\/([^/?#]+)/i) || [])[1] || 'リンク';
    out.push({ type: 'other', label: host, url: other });
  }
  return out;
}

/* =========================================================
 * 投稿内容のチェック（新規・修正で共通）
 * ========================================================= */
function parseEventInput_(b) {
  const now = Date.now();
  const DAY = 86400000;
  const p = {
    title: clean_(b.title, 40),
    description: clean_(b.description, 500),
    category: CATEGORIES[b.category] ? String(b.category) : 'other',
    capacity: clean_(b.capacity, 20) || '何人でも',
    fee: clean_(b.fee, 20) || '無料',
    conditions: clean_(b.conditions, 100) || '誰でも歓迎',
    lat: Number(b.lat),
    lng: Number(b.lng),
    address: clean_(b.address, 100),
    startMs: Date.parse(String(b.start_at || '')),
    endMs: Date.parse(String(b.end_at || '')),
    expMs: Date.parse(String(b.expires_at || '')),
    sns: normalizeSns_(b.sns),
    joinMethod: b.join_method === 'direct' ? 'direct' : 'app',
  };
  if (!p.title) throw apiError_('タイトルを入力してください');
  if (!isFinite(p.lat) || !isFinite(p.lng) || Math.abs(p.lat) > 90 || Math.abs(p.lng) > 180) throw apiError_('場所が正しくありません');
  if (isNaN(p.startMs) || isNaN(p.endMs)) throw apiError_('開催日時（開始・終了）を入力してください');
  if (p.endMs <= p.startMs) throw apiError_('終了日時は開始日時より後にしてください');
  if (p.endMs - p.startMs > CONFIG.MAX_EVENT_DAYS * DAY) throw apiError_('開催期間は' + CONFIG.MAX_EVENT_DAYS + '日以内にしてください');
  if (p.endMs <= now) throw apiError_('終了日時が過ぎています');
  if (p.startMs > now + CONFIG.MAX_DAYS_AHEAD * DAY) throw apiError_(CONFIG.MAX_DAYS_AHEAD + '日以内に始まるイベントにしてください');
  if (isNaN(p.expMs) || p.expMs <= now + 5 * 60000) throw apiError_('表示期限は5分以上先にしてください');
  if (p.expMs > now + (CONFIG.MAX_DAYS_AHEAD + CONFIG.MAX_EVENT_DAYS) * DAY) throw apiError_('表示期限が長すぎます');
  if (p.startMs > now && p.expMs < p.startMs) throw apiError_('表示期限が開催開始より前になっています');
  if (p.joinMethod === 'direct' && !p.sns.length) throw apiError_('「主催者に直接申し込み」にする場合は、連絡先SNSを1つ以上入力してください');
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

/** シートに書く項目（新規・修正で共通） */
function eventFields_(p, ai, geo) {
  return {
    title: p.title,
    description: p.description,
    category: p.category,
    icon: ai.emoji || CATEGORIES[p.category].icon,
    summary: ai.summary || p.title.slice(0, 15),
    lat: p.lat.toFixed(6),
    lng: p.lng.toFixed(6),
    start_at: new Date(p.startMs).toISOString(),
    end_at: new Date(p.endMs).toISOString(),
    expires_at: new Date(p.expMs).toISOString(),
    capacity: p.capacity,
    fee: p.fee,
    conditions: p.conditions,
    hashtags: (ai.hashtags || []).join(' '),
    address: p.address || [geo.pref, geo.town].filter(Boolean).join(' '),
    pref: geo.pref,
    area: geo.area,
    sns: JSON.stringify(p.sns || []),
    join_method: p.joinMethod,
  };
}

function reviewInput_(p, media) {
  return {
    title: p.title, description: p.description, category: p.category,
    capacity: p.capacity, fee: p.fee, conditions: p.conditions,
    address: p.address, sns: (p.sns || []).map(x => x.url).join(' '),
    media: media,
  };
}

/* =========================================================
 * 投稿作成（Geminiチェック → Driveアップロード → シート追記）
 * ========================================================= */
function createEvent_(b) {
  const user = authUser_(b.token);
  const p = parseEventInput_(b);
  const now = Date.now();

  // 同時公開数の上限
  const mine = readAll_(sheet_(CONFIG.SHEET_EVENTS))
    .filter(r => String(r.user_id) === user.user_id && String(r.status) === 'active' && ms_(r.expires_at) > now);
  if (mine.length >= CONFIG.MAX_ACTIVE_POSTS_PER_USER) {
    throw apiError_('同時に公開できるのは' + CONFIG.MAX_ACTIVE_POSTS_PER_USER + '件までです。終わった投稿を削除してください');
  }

  const media = mediaInput_(b);

  // 1) Gemini モデレーション & 生成
  const ai = geminiReview_(reviewInput_(p, media));
  if (!ai.safe) {
    return { ok: false, moderation: true, error: 'この内容は投稿できません：' + (ai.reason || 'ガイドラインに抵触する可能性があります') };
  }

  // 2) エリア判定・Drive アップロード
  const geo = reverseGeo_(p.lat, p.lng);
  const id = Utilities.getUuid();
  let up = { url: '', type: '', id: '' };
  if (media) up = uploadMedia_(media, id);

  // 3) シート追記（失敗時はアップロードを取り消し）
  const row = Object.assign({
    id: id,
    created_at: new Date(now).toISOString(),
    user_id: user.user_id,
    status: 'active',
    media_url: up.url,
    media_type: up.type,
    media_file_id: up.id,
  }, eventFields_(p, ai, geo));
  try {
    withLock_(() => { appendObj_(sheet_(CONFIG.SHEET_EVENTS), row); });
  } catch (err) {
    if (up.id) trashMedia_(up.id);
    throw apiError_('保存に失敗しました。もう一度お試しください');
  }
  clearEventsCache_();
  return { ok: true, event: publicEvent_(row, { [user.user_id]: { nickname: user.nickname } }, {}) };
}

/* =========================================================
 * 投稿の修正（本人のみ・公開中のみ）
 *   media: 新しい写真/動画（差し替え） / removeMedia: true で外す
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

  const ai = geminiReview_(reviewInput_(p, media));
  if (!ai.safe) {
    return { ok: false, moderation: true, error: 'この内容には修正できません：' + (ai.reason || 'ガイドラインに抵触する可能性があります') };
  }

  const geo = reverseGeo_(p.lat, p.lng);
  let up = null;
  if (media) up = uploadMedia_(media, id + '_' + Date.now());

  let oldFileId = '';
  let saved;
  try {
    saved = withLock_(() => {
      const sh = sheet_(CONFIG.SHEET_EVENTS);
      const ev = findEv_();
      checkOwner_(ev);
      oldFileId = String(ev.media_file_id || '');
      const fields = eventFields_(p, ai, geo);
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

  if ((up || removeMedia) && oldFileId) trashMedia_(oldFileId);
  clearEventsCache_();
  return { ok: true, event: publicEvent_(saved, { [user.user_id]: { nickname: user.nickname } }, participantStats_()) };
}

/* =========================================================
 * 投稿削除
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
 * 参加表明
 *   参加者リストは「主催者」と「参加表明した人」だけが見られる
 *   閲覧者には人数だけを公開（getEvents の participant_count）
 * ========================================================= */
function partPublic_(r, user) {
  return {
    id: String(r.id),
    nickname: String(r.nickname),
    people: Number(r.people) || 1,
    message: String(r.message || ''),
    added_by: String(r.added_by || 'self'),
    created_at: iso_(r.created_at),
    mine: String(r.user_id) === user.user_id,
  };
}

function participantsView_(user, ev, allRows) {
  const rows = allRows
    .filter(r => String(r.event_id) === String(ev.id))
    .sort((a, b) => (iso_(a.created_at) < iso_(b.created_at) ? -1 : 1));
  const isOwner = String(ev.user_id) === user.user_id;
  const meRow = rows.find(r => String(r.user_id) === user.user_id);
  const allowed = isOwner || !!meRow;
  return {
    ok: true,
    event_id: String(ev.id),
    isOwner: isOwner,
    joined: !!meRow,
    me: meRow ? partPublic_(meRow, user) : null,
    list: allowed ? rows.map(r => partPublic_(r, user)) : null,
    n: rows.length,
    count: rows.reduce((a, r) => a + (Number(r.people) || 1), 0),
  };
}

function findEvent_(eventId) {
  const ev = readAll_(sheet_(CONFIG.SHEET_EVENTS)).find(r => String(r.id) === String(eventId));
  if (!ev) throw apiError_('イベントが見つかりません');
  return ev;
}

function getParticipants_(b) {
  const user = authUser_(b.token);
  const ev = findEvent_(b.event_id);
  return participantsView_(user, ev, readAll_(sheet_(CONFIG.SHEET_PARTICIPANTS)));
}

/**
 * 参加表明の登録・更新
 *   pid なし・asOwner なし → 自分の参加表明（あれば更新）
 *   pid なし・asOwner      → 主催者が参加者を追加（会員以外もOK）
 *   pid あり               → 主催者は全員、本人は自分の分を編集
 */
function saveParticipant_(b) {
  const user = authUser_(b.token);
  const eventId = String(b.event_id || '');
  const people = Math.max(1, Math.min(CONFIG.MAX_PEOPLE, Math.floor(Number(b.people) || 1)));
  const message = clean_(b.message, 140);
  return withLock_(() => {
    const ev = findEvent_(eventId);
    if (String(ev.status) !== 'active' || ms_(ev.expires_at) <= Date.now()) throw apiError_('このイベントは掲載が終了しています');
    const isOwner = String(ev.user_id) === user.user_id;
    const sh = sheet_(CONFIG.SHEET_PARTICIPANTS);
    const rows = readAll_(sh);
    const now = new Date().toISOString();

    if (b.pid) {
      const r = rows.find(x => String(x.id) === String(b.pid) && String(x.event_id) === eventId);
      if (!r) throw apiError_('参加者が見つかりません');
      const own = String(r.user_id) === user.user_id;
      if (!isOwner && !own) throw apiError_('この参加者は編集できません');
      const f = { people: people, message: message, updated_at: now };
      if (isOwner) {
        const nn = clean_(b.nickname, 20);
        if (nn) f.nickname = nn;
      }
      setFields_(sh, r._row, f);
    } else if (b.asOwner) {
      if (!isOwner) throw apiError_('参加者を追加できるのは主催者だけです');
      const nn = clean_(b.nickname, 20);
      if (!nn) throw apiError_('参加者の名前を入力してください');
      appendObj_(sh, {
        id: Utilities.getUuid(), event_id: eventId, user_id: '', nickname: nn,
        people: people, message: message, added_by: 'owner', created_at: now, updated_at: now,
      });
    } else {
      if (isOwner) throw apiError_('主催者は自分のイベントに参加表明できません（参加者の追加を使ってください）');
      const mineNow = rows.find(x => String(x.event_id) === eventId && String(x.user_id) === user.user_id);
      if (String(ev.join_method) === 'direct' && !mineNow) throw apiError_('このイベントの参加表明は、主催者に直接申し込んでください');
      if (ms_(ev.end_at || ev.expires_at) <= Date.now()) throw apiError_('このイベントは終了しています');
      const mine = rows.find(x => String(x.event_id) === eventId && String(x.user_id) === user.user_id);
      if (mine) {
        setFields_(sh, mine._row, { people: people, message: message, nickname: user.nickname, updated_at: now });
      } else {
        appendObj_(sh, {
          id: Utilities.getUuid(), event_id: eventId, user_id: user.user_id, nickname: user.nickname,
          people: people, message: message, added_by: 'self', created_at: now, updated_at: now,
        });
      }
    }
    clearEventsCache_();
    return participantsView_(user, ev, readAll_(sh));
  });
}

/** 参加表明の削除（主催者は全員、本人は自分の分＝取り消し） */
function deleteParticipant_(b) {
  const user = authUser_(b.token);
  const eventId = String(b.event_id || '');
  return withLock_(() => {
    const ev = findEvent_(eventId);
    const isOwner = String(ev.user_id) === user.user_id;
    const sh = sheet_(CONFIG.SHEET_PARTICIPANTS);
    const r = readAll_(sh).find(x => String(x.id) === String(b.pid) && String(x.event_id) === eventId);
    if (!r) throw apiError_('参加者が見つかりません');
    if (!isOwner && String(r.user_id) !== user.user_id) throw apiError_('この参加者は削除できません');
    sh.deleteRow(r._row);
    clearEventsCache_();
    return participantsView_(user, ev, readAll_(sh));
  });
}

/* =========================================================
 * マイページ
 * ========================================================= */
function getMyData_(b) {
  const user = authUser_(b.token);
  const now = Date.now();
  const events = readAll_(sheet_(CONFIG.SHEET_EVENTS));
  const parts = readAll_(sheet_(CONFIG.SHEET_PARTICIPANTS));
  const stats = participantStats_(parts);
  const me = { [user.user_id]: { nickname: user.nickname } };

  const myEvents = events
    .filter(r => String(r.user_id) === user.user_id && String(r.status) === 'active' && ms_(r.expires_at) > now)
    .map(r => {
      const ev = publicEvent_(r, me, stats);
      ev.participants = participantsView_(user, r, parts).list || [];
      return ev;
    })
    .sort((a, b2) => (a.start_at < b2.start_at ? -1 : 1));

  const evMap = {};
  events.forEach(r => { evMap[String(r.id)] = r; });
  const joined = parts
    .filter(q => String(q.user_id) === user.user_id)
    .map(q => {
      const ev = evMap[String(q.event_id)];
      const alive = ev && String(ev.status) === 'active' && ms_(ev.expires_at) > now;
      return {
        event_id: String(q.event_id),
        title: ev ? String(ev.title) : '（削除されたイベント）',
        icon: ev ? String(ev.icon) : '📍',
        start_at: ev ? (iso_(ev.start_at) || iso_(ev.created_at)) : '',
        active: !!alive,
        people: Number(q.people) || 1,
      };
    })
    .sort((a, b2) => (a.start_at < b2.start_at ? -1 : 1))
    .slice(0, 50);

  return {
    ok: true,
    user: { user_id: user.user_id, nickname: user.nickname, email: user.email },
    myEvents: myEvents,
    joinedEvents: joined,
  };
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
    場所: p.address || '', 連絡先URL: p.sns || '',
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
