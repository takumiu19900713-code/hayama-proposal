// ============================================================
// ハヤマGS 洗車管理システム - Google Apps Script
// このコードをGoogleスプレッドシートのApps Scriptに貼り付けてください
// ============================================================

const SHEET_ID = '1nl9agpXOhxzUInvM2ui1IBA7ozaxZsdF';
const LOG_SHEET  = '依頼ログ';
const MASTER_SHEET = '設定マスタ';

// LINE WORKS Bot設定（③で取得したものを入れる）
const LW_BOT_SECRET  = 'ここにBotのSecretを入れる';
const LW_BOT_ID      = 'ここにBot IDを入れる';
const LW_ACCESS_TOKEN = 'ここにアクセストークンを入れる';

// ============================================================
// GET リクエスト処理（Webアプリからの読み取り）
// ============================================================
function doGet(e) {
  const action = e.parameter.action;
  let result;

  try {
    if (action === 'ping') {
      result = {status:'ok', message:'接続成功'};

    } else if (action === 'getToday') {
      const date = e.parameter.date || todayStr();
      result = {status:'ok', requests: getTodayRequests(date)};

    } else if (action === 'getMonthly') {
      const ym = e.parameter.ym || currentYM();
      result = {status:'ok', ...getMonthlyData(ym)};

    } else {
      result = {status:'error', message:'不明なアクション'};
    }
  } catch(err) {
    result = {status:'error', message: err.message};
  }

  return jsonResponse(result);
}

// ============================================================
// POST リクエスト処理（受入・完了の書き込み）
// ============================================================
function doPost(e) {
  const params = JSON.parse(e.postData.contents);
  const action = params.action;
  let result;

  try {
    if (action === 'accept') {
      acceptRequest(params.id, params.acceptedAt, params.estimateAt);
      // スカイグループへLINE WORKS通知
      notifyToSky(params.id, 'accept', params.acceptedAt, params.estimateAt);
      result = {status:'ok'};

    } else if (action === 'complete') {
      completeRequest(params.id, params.doneAt);
      // スカイグループへ完了通知
      notifyToSky(params.id, 'complete', params.doneAt, '');
      result = {status:'ok'};

    } else if (action === 'newRequest') {
      // LINE WORKSボットから新規依頼が来たとき
      const id = addNewRequest(params);
      result = {status:'ok', id};

    } else {
      result = {status:'error', message:'不明なアクション'};
    }
  } catch(err) {
    result = {status:'error', message: err.message};
  }

  return jsonResponse(result);
}

// ============================================================
// データ取得関数
// ============================================================

// 当日の依頼一覧を取得
function getTodayRequests(dateStr) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(LOG_SHEET);
  const data = sheet.getDataRange().getValues();
  const headers = data[0];

  const COL = {};
  headers.forEach((h, i) => COL[h] = i);

  const today = normalizeDate(dateStr);
  const results = [];

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const rowDate = normalizeDate(String(row[COL['日付']]));
    if (rowDate !== today) continue;

    const co = String(row[COL['会社コード']]);
    const status = String(row[COL['ステータス']]);
    const acceptedAt = String(row[COL['受入時刻']] || '');

    results.push({
      id:          String(row[COL['ID']]),
      co:          co,
      requester:   String(row[COL['依頼者名']] || ''),
      requestedAt: String(row[COL['依頼時刻']] || ''),
      status:      status,
      acceptedAt:  acceptedAt,
      estimateAt:  String(row[COL['完了目安']] || ''),
      doneAt:      String(row[COL['完了時刻']] || ''),
      // 受入済みなら経過時間計算用タイムスタンプ
      startTime:   (status === '作業中' && acceptedAt)
                     ? parseTimeToTimestamp(today, acceptedAt)
                     : null,
    });
  }

  // 新着→作業中→完了の順に並べる
  const order = {'待機中':0,'作業中':1,'完了':2};
  results.sort((a,b)=>(order[a.status]||0)-(order[b.status]||0));
  return results;
}

// 月次集計データを取得
function getMonthlyData(ym) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(LOG_SHEET);
  const masterSheet = ss.getSheetByName(MASTER_SHEET);
  const data = sheet.getDataRange().getValues();
  const headers = data[0];

  const COL = {};
  headers.forEach((h, i) => COL[h] = i);

  // 設定マスタから会社情報取得
  const masterData = masterSheet.getDataRange().getValues();
  const companies = {};
  for (let i = 1; i < masterData.length; i++) {
    if (!masterData[i][0]) continue;
    companies[masterData[i][0]] = {
      co: masterData[i][0],
      name: masterData[i][1],
      quota: Number(masterData[i][3]) || 30,
      count: 0,
    };
  }

  // 日別集計用
  const dailyMap = {};

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const dateRaw = String(row[COL['日付']]);
    if (!dateRaw || dateRaw === 'undefined') continue;

    const rowYM = dateRaw.substring(0, 7).replace('/', '-');
    if (rowYM !== ym) continue;

    const co = String(row[COL['会社コード']]);
    if (companies[co]) companies[co].count++;

    // 日別集計
    const dateKey = dateRaw;
    if (!dailyMap[dateKey]) dailyMap[dateKey] = {};
    dailyMap[dateKey][co] = (dailyMap[dateKey][co] || 0) + 1;
  }

  // 日別データを配列化・降順ソート
  const daily = Object.entries(dailyMap)
    .sort(([a],[b])=>b.localeCompare(a))
    .slice(0, 14)
    .map(([date, counts]) => ({
      date: formatDateLabel(date),
      counts
    }));

  return {
    companies: Object.values(companies),
    daily
  };
}

// ============================================================
// 書き込み関数
// ============================================================

// 新規依頼を追加（LINE WORKSボットから）
function addNewRequest(params) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(LOG_SHEET);

  const lastRow = sheet.getLastRow();
  const newId = lastRow; // 行番号をIDとして使用
  const now = new Date();

  sheet.appendRow([
    newId,
    formatDate(now),     // 日付
    formatTime(now),     // 依頼時刻
    coName(params.co),   // 会社名
    params.co,           // 会社コード
    params.requester || '',
    '待機中',            // ステータス
    '',                  // 受入時刻
    '',                  // 完了目安
    '',                  // 完了時刻
    '',                  // 作業時間
    '',                  // 備考
  ]);

  // ハヤマGSスタッフへ新着通知
  notifyToGS(newId, params.co, params.requester);

  return String(newId);
}

// 受入処理
function acceptRequest(id, acceptedAt, estimateAt) {
  const row = findRowById(id);
  if (!row) throw new Error('ID:' + id + ' が見つかりません');
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(LOG_SHEET);
  const COL = getColMap(sheet);
  sheet.getRange(row, COL['ステータス']+1).setValue('作業中');
  sheet.getRange(row, COL['受入時刻']+1).setValue(acceptedAt);
  if (estimateAt) sheet.getRange(row, COL['完了目安']+1).setValue(estimateAt);
}

// 完了処理
function completeRequest(id, doneAt) {
  const row = findRowById(id);
  if (!row) throw new Error('ID:' + id + ' が見つかりません');
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(LOG_SHEET);
  const COL = getColMap(sheet);
  const acceptedAt = sheet.getRange(row, COL['受入時刻']+1).getValue();
  sheet.getRange(row, COL['ステータス']+1).setValue('完了');
  sheet.getRange(row, COL['完了時刻']+1).setValue(doneAt);
}

// ============================================================
// LINE WORKS 通知
// ============================================================

// スカイグループの該当会社チャンネルへ通知
function notifyToSky(id, type, time, estimate) {
  const row = findRowById(id);
  if (!row) return;
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(LOG_SHEET);
  const COL = getColMap(sheet);
  const co = sheet.getRange(row, COL['会社コード']+1).getValue();

  // 設定マスタからチャンネルIDを取得
  const channelId = getChannelId(co);
  if (!channelId || channelId.startsWith('ch_')) return; // 未設定

  let message = '';
  if (type === 'accept') {
    message = `【受入完了】\n${coName(co)}の車を受け入れました。\n完了目安：${estimate || '未定'}`;
  } else if (type === 'complete') {
    message = `【洗車完了】\n${coName(co)}の洗車が完了しました。\n完了時刻：${time}\nお引き取りをお願いします。`;
  }

  sendLWMessage(channelId, message);
}

// ハヤマGSスタッフへ新着通知
function notifyToGS(id, co, requester) {
  const gsChannelId = getChannelId('GS');
  if (!gsChannelId || gsChannelId.startsWith('ch_')) return;
  const message = `【新着依頼】\n${coName(co)}\n依頼者：${requester||'不明'}\n管理画面で受入してください。`;
  sendLWMessage(gsChannelId, message);
}

// LINE WORKS メッセージ送信
function sendLWMessage(channelId, message) {
  try {
    const url = `https://www.worksapis.com/v1.0/bots/${LW_BOT_ID}/channels/${channelId}/messages`;
    const payload = JSON.stringify({
      content: {type:'text', text: message}
    });
    UrlFetchApp.fetch(url, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'Authorization': 'Bearer ' + LW_ACCESS_TOKEN,
      },
      payload,
      muteHttpExceptions: true,
    });
  } catch(e) {
    console.log('LW通知エラー:', e.message);
  }
}

// ============================================================
// ユーティリティ
// ============================================================

function findRowById(id) {
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(LOG_SHEET);
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(id)) return i + 1;
  }
  return null;
}

function getColMap(sheet) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const map = {};
  headers.forEach((h, i) => map[h] = i);
  return map;
}

function getChannelId(co) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(MASTER_SHEET);
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === co) return String(data[i][2]);
  }
  return '';
}

function coName(co) {
  const names = {BMW:'BMW正規販売店',JLR:'ジャガーランドローバー',POR:'ポルシェ',VOL:'ボルボ'};
  return names[co] || co;
}

function formatDate(d) {
  return `${d.getFullYear()}/${String(d.getMonth()+1).padStart(2,'0')}/${String(d.getDate()).padStart(2,'0')}`;
}

function formatTime(d) {
  return `${String(d.getHours()).padStart(2,'0')}:${String(d.getMinutes()).padStart(2,'0')}`;
}

function normalizeDate(s) {
  return s.replace(/\//g,'-').substring(0,10);
}

function todayStr() {
  return formatDate(new Date());
}

function currentYM() {
  const d = new Date();
  return `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,'0')}`;
}

function parseTimeToTimestamp(dateStr, timeStr) {
  if (!timeStr || timeStr === 'null') return null;
  const d = new Date(`${dateStr}T${timeStr}:00`);
  return isNaN(d.getTime()) ? null : d.getTime();
}

function formatDateLabel(dateStr) {
  const d = new Date(dateStr.replace(/\//g,'-'));
  const wd = ['日','月','火','水','木','金','土'];
  return `${d.getMonth()+1}/${d.getDate()}（${wd[d.getDay()]}）`;
}

function jsonResponse(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}
