// LINE日報ボット（Google Apps Script）
// 入口は2つ（どちらも doPost）: (1) LIFFフォームからのJSON {action,...}  (2) LINEのWebhook {events}
// 画面(index.html)はGitHub Pagesで配信。GASは保存・一覧のAPIと通知だけ。
// 保存先は先頭シート1枚。人ごとのタブはFILTER数式で自動生成（生データは触らない）。
// ponytail: GASのdoPostはHTTPヘッダーが読めずLINE署名検証ができない。Webhook URLを公開しないことで代替。
const TZ = 'Asia/Tokyo';
const PROPS = PropertiesService.getScriptProperties();
const TOKEN = PROPS.getProperty('CHANNEL_ACCESS_TOKEN');   // Messaging APIチャネルのアクセストークン
const RECEIVER = PROPS.getProperty('RECEIVER_ID');         // 受信者のuserId
const LOGIN_CHANNEL_ID = PROPS.getProperty('LOGIN_CHANNEL_ID'); // LINE LoginチャネルID（IDトークン検証用）
const HEADER = ['名前', 'userId', '送信日時', '対象日', '本文'];
const ss = () => SpreadsheetApp.getActiveSpreadsheet();
const sheet = () => ss().getSheets()[0];
const fmt = (d, p) => Utilities.formatDate(d, TZ, p);

// ---------- 入口 ----------
function doGet() { return ContentService.createTextOutput('日報API'); }

function doPost(e) {
  const body = JSON.parse(e.postData.contents || '{}');
  if (body.events) { body.events.forEach(handle); return ContentService.createTextOutput('ok'); } // LINE Webhook
  let out;
  try {
    if (body.action === 'list') out = { reports: myReports(body.idToken) };
    else if (body.action === 'save') out = saveReport(body.idToken, body.date, body.text, body.overwrite);
    else throw new Error('不明な操作です');
  } catch (err) { out = { error: err.message }; }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}

// ---------- LIFFフォーム用API ----------
// LIFFから渡されたIDトークンをLINEに照会して本人確認（信頼境界）
function verify(idToken) {
  const res = UrlFetchApp.fetch('https://api.line.me/oauth2/v2.1/verify', {
    method: 'post', payload: { id_token: idToken, client_id: LOGIN_CHANNEL_ID }, muteHttpExceptions: true,
  });
  if (res.getResponseCode() !== 200) throw new Error('本人確認に失敗しました。画面を開き直してください');
  const p = JSON.parse(res.getContentText());
  return { uid: p.sub, name: p.name || p.sub };
}

function myReports(idToken) {
  const { uid } = verify(idToken);
  return rowsOf(uid).sort((a, b) => b[2] - a[2]) // 送信の新しい順
    .map(r => ({ date: fmt(r[3], 'yyyy-MM-dd'), sentAt: fmt(r[2], 'MM/dd HH:mm'), text: r[4] }));
}

function saveReport(idToken, date, text, overwrite) {
  const { uid, name } = verify(idToken);
  text = String(text || '').trim();
  if (!/^\d{4}-\d{2}-\d{2}$/.test(date) || !text) throw new Error('日付と本文を入れてください');
  const [y, m, d] = date.split('-').map(Number);
  const row = findRow(uid, date);
  if (row && !overwrite) return { dup: true };
  save(name, uid, new Date(), new Date(y, m - 1, d), text, row);
  return { ok: true };
}

// ---------- トークで送る方式 ----------
function handle(ev) {
  if (ev.type !== 'message' || ev.message.type !== 'text' || ev.source.type !== 'user') return;
  const uid = ev.source.userId;
  const text = ev.message.text.trim();
  if (text === 'ID') return reply(ev.replyToken, uid);
  if (text === '履歴') return reply(ev.replyToken, history(uid));
  // TODO 本番切替時: 受信者自身の送信を弾くなら次の行を戻す
  // if (uid === RECEIVER) return reply(ev.replyToken, 'このアカウントは受信者として登録されています');

  const sentAt = new Date(ev.timestamp);
  const forDate = targetDate(text, sentAt);
  save(profileName(uid), uid, sentAt, forDate, text);
  reply(ev.replyToken, `${fmt(forDate, 'M/d')}分の日報を受け取りました`);
}

// 1行目が「10/15」のような日付ならその日、なければ送信日。年は送信年、未来になるなら前年。
function targetDate(text, sentAt) {
  const m = text.match(/^(\d{1,2})\/(\d{1,2})/);
  if (!m) return new Date(sentAt.getFullYear(), sentAt.getMonth(), sentAt.getDate());
  const d = new Date(sentAt.getFullYear(), m[1] - 1, m[2]);
  if (d > sentAt) d.setFullYear(d.getFullYear() - 1);
  return d;
}

function history(uid) {
  const rows = rowsOf(uid).slice(-10);
  if (!rows.length) return 'まだ日報はありません';
  const out = rows.map(r => `■ ${fmt(r[3], 'yyyy/MM/dd')}（送信 ${fmt(r[2], 'MM/dd HH:mm')}）\n${r[4]}`).join('\n\n');
  return out.length > 4900 ? '…' + out.slice(-4900) : out; // LINE1通の上限5000字
}

// ---------- 共通: 保存・検索・通知 ----------
// row を渡すとその行を上書き（同じ日の再送信）
function save(name, uid, sentAt, forDate, text, row) {
  const s = sheet();
  if (s.getLastRow() === 0) s.appendRow(HEADER);
  if (row) s.getRange(row, 1, 1, 5).setValues([[name, uid, sentAt, forDate, text]]);
  else s.appendRow([name, uid, sentAt, forDate, text]);
  ensurePersonTab(name, uid);
  if (RECEIVER) push(RECEIVER, `【日報${row ? '・更新' : ''}】${name}\n対象日: ${fmt(forDate, 'yyyy/MM/dd')}\n送信: ${fmt(sentAt, 'yyyy/MM/dd HH:mm')}\n\n${text}`);
}

const isRow = r => r[2] instanceof Date && r[3] instanceof Date;
// ponytail: 全行スキャン。数万行を超えたらuserId列のインデックス化を検討
function rowsOf(uid) { return sheet().getDataRange().getValues().filter(r => r[1] === uid && isRow(r)); }
function findRow(uid, date) {
  const i = sheet().getDataRange().getValues().findIndex(r => r[1] === uid && isRow(r) && fmt(r[3], 'yyyy-MM-dd') === date);
  return i < 0 ? null : i + 1;
}

// 人ごとのタブ: 生データをFILTERで映すだけ。初めての人が送った時に1回だけ作る
function ensurePersonTab(name, uid) {
  if (ss().getSheets().some(s => s.getRange('A1').getFormula().includes(uid))) return;
  const src = sheet().getName();
  let title = (name || uid).slice(0, 50);
  if (ss().getSheetByName(title)) title += ' ' + uid.slice(-4); // 同名の人がいた時
  const tab = ss().insertSheet(title, ss().getNumSheets());
  tab.getRange('A1').setFormula(`={'${src}'!A1:E1; IFERROR(FILTER('${src}'!A2:E, '${src}'!B2:B="${uid}"), "")}`);
}

function profileName(uid) {
  const res = UrlFetchApp.fetch('https://api.line.me/v2/bot/profile/' + uid, {
    headers: { Authorization: 'Bearer ' + TOKEN }, muteHttpExceptions: true,
  });
  return res.getResponseCode() === 200 ? JSON.parse(res.getContentText()).displayName : uid;
}

function reply(replyToken, text) { lineApi('reply', { replyToken, messages: [{ type: 'text', text }] }); }
function push(to, text) { lineApi('push', { to, messages: [{ type: 'text', text: text.slice(0, 5000) }] }); }
function lineApi(path, payload) {
  UrlFetchApp.fetch('https://api.line.me/v2/bot/message/' + path, {
    method: 'post', contentType: 'application/json',
    headers: { Authorization: 'Bearer ' + TOKEN },
    payload: JSON.stringify(payload), muteHttpExceptions: true,
  });
}

// エディタで選んで「実行」→ ログに ok と出れば日付判定は正常
function test_targetDate() {
  const sent = new Date(2026, 0, 2, 9, 0); // 2026/1/2 9:00 に送信
  const eq = (a, b) => { if (fmt(a, 'yyyy/MM/dd') !== b) throw new Error(`${fmt(a, 'yyyy/MM/dd')} !== ${b}`); };
  eq(targetDate('今日の作業です', sent), '2026/01/02');
  eq(targetDate('1/1\n昨日の分', sent), '2026/01/01');
  eq(targetDate('12/31\n年末の分', sent), '2025/12/31');
  Logger.log('ok');
}
