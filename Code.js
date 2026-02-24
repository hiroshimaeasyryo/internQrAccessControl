const STAFF_SHEET_NAME = 'スタッフDB';
const TIMESTAMP_SHEET_NAME = '打刻記録';

/**
 * メイン画面表示 / API エンドポイント
 */
function doGet(e) {
  const action = e.parameter.action;

  // API 呼び出しの場合 (action パラメータがある場合)
  if (action) {
    try {
      let result;
      switch (action) {
        case 'getStaffList':
          result = getStaffList();
          break;
        case 'verifyStaff':
          result = verifyStaff(e.parameter.uuid, e.parameter.birthdate);
          break;
        case 'recordTimestamp':
          // JSON 文字列としてパース
          const payload = JSON.parse(e.parameter.payload);
          result = recordTimestamp(payload);
          break;
        case 'clearStaffCache':
          result = clearStaffCache();
          break;
        default:
          throw new Error('Unknown action: ' + action);
      }
      const output = JSON.stringify(result);
      // JSONP 対応: callback があれば関数呼び出し形式にする
      const callback = e.parameter.callback;
      if (callback) {
        return ContentService.createTextOutput(callback + '(' + output + ')')
          .setMimeType(ContentService.MimeType.JAVASCRIPT);
      }
      return ContentService.createTextOutput(output)
        .setMimeType(ContentService.MimeType.JSON);
    } catch (err) {
      const errorOutput = JSON.stringify({ ok: false, message: err.message });
      const callback = e.parameter.callback;
      if (callback) {
        return ContentService.createTextOutput(callback + '(' + errorOutput + ')')
          .setMimeType(ContentService.MimeType.JAVASCRIPT);
      }
      return ContentService.createTextOutput(errorOutput)
        .setMimeType(ContentService.MimeType.JSON);
    }
  }

  // 通常の Web 表示 (後方互換性のため残す)
  const tmpl = HtmlService.createTemplateFromFile('index');
  tmpl.placeId = e.parameter.place || '';
  return tmpl.evaluate()
    .setTitle('KANPAI Hütte 入退室記録')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no')
    .setFaviconUrl('https://drive.google.com/uc?id=1YkdqM2adcpxtVM-nA8uVGGGPi2WYPkRu&.png');
}

/**
 * テンプレート内でHTMLファイルをインクルードするためのヘルパー
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * スタッフ一覧を取得（プルダウン用）
 * キャッシュを利用して高速化する
 * return: [{ uuid, name, birthdate, img }, ...]
 */
function getStaffList() {
  const cacheKey = 'staff_list_cache';
  const cache = CacheService.getScriptCache();
  const cachedData = cache.get(cacheKey);

  if (cachedData) {
    console.log('Using cached staff list');
    return JSON.parse(cachedData);
  }

  console.log('Cache miss. Fetching from Spreadsheet');
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(STAFF_SHEET_NAME);
  if (!sheet) {
    throw new Error('staffs シートが見つかりません');
  }

  const values = sheet.getDataRange().getValues();
  const header = values[0];
  const uuidIndex = header.indexOf('uuid');
  const nameIndex = header.indexOf('name');
  const birthIndex = header.indexOf('birthdate');
  const imgIndex = header.indexOf('img');

  if (uuidIndex === -1 || nameIndex === -1 || birthIndex === -1) {
    throw new Error('staffs シートに uuid / name / birthdate 列がありません');
  }

  const result = [];
  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    if (!row[uuidIndex] || !row[nameIndex]) continue;
    result.push({
      uuid: String(row[uuidIndex]),
      name: String(row[nameIndex]),
      birthdate: row[birthIndex] ? Utilities.formatDate(new Date(row[birthIndex]), Session.getScriptTimeZone(), 'yyyy-MM-dd') : '',
      img: row[imgIndex] ? String(row[imgIndex]) : ''
    });
  }

  // キャッシュに保存（有効期限は6時間 = 21600秒）
  try {
    cache.put(cacheKey, JSON.stringify(result), 21600);
  } catch (e) {
    console.error('Failed to put cache:', e);
  }

  return result;
}

/**
 * スタッフ一覧のキャッシュを明示的にクリアする
 */
function clearStaffCache() {
  const cache = CacheService.getScriptCache();
  cache.remove('staff_list_cache');
  return { ok: true, message: 'キャッシュをクリアしました' };
}

/**
 * 生年月日とスタッフUUIDを検証
 * @param {string} uuid
 * @param {string} birthdateStr - 'YYYY-MM-DD'
 * @returns {{ok: boolean, name?: string, message?: string}}
 */
function verifyStaff(uuid, birthdateStr) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(STAFF_SHEET_NAME);
  if (!sheet) {
    return { ok: false, message: 'staffs シートが見つかりません' };
  }

  const values = sheet.getDataRange().getValues();
  const header = values[0];
  const uuidIndex = header.indexOf('uuid');
  const nameIndex = header.indexOf('name');
  const birthIndex = header.indexOf('birthdate');

  if (uuidIndex === -1 || nameIndex === -1 || birthIndex === -1) {
    return { ok: false, message: 'staffs シートに uuid / name / birthdate 列がありません' };
  }

  // 入力された birthdateStr を Dateとして扱う
  const normalizedInput = birthdateStr.trim();

  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    if (String(row[uuidIndex]) !== String(uuid)) continue;

    const name = String(row[nameIndex]);
    const cell = row[birthIndex];
    let cellStr = '';

    if (cell instanceof Date) {
      cellStr = Utilities.formatDate(cell, Session.getScriptTimeZone(), 'yyyy-MM-dd');
    } else {
      cellStr = String(cell);
    }

    if (cellStr === normalizedInput) {
      return { ok: true, name };
    } else {
      return { ok: false, message: '生年月日が一致しません' };
    }
  }

  return { ok: false, message: '該当のスタッフが見つかりません' };
}

/**
 * 打刻を記録
 * @param {object} payload
 *   payload = {
 *     uuid: string,
 *     name: string,
 *     placeId: string,
 *     type: string, // 'in' | 'out' など
 *     qrValue: string, // 実際に読み取ったQRの文字列（ログに残したければ）
 *   }
 */
function recordTimestamp(payload) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(TIMESTAMP_SHEET_NAME);
  if (!sheet) {
    throw new Error('timestamps シートが見つかりません');
  }

  const now = new Date();
  const userAgent = Session.getActiveUser().getEmail() || 'unknown';

  sheet.appendRow([
    now,
    payload.uuid || '',
    payload.name || '',
    payload.placeId || '',
    payload.type || '',
    userAgent
  ]);

  // 今月の出勤回数を計算（今回の分も含む）
  let attendanceCount = 0;
  if (payload.type === 'in') {
    attendanceCount = getMonthlyAttendanceCount(payload.uuid);
  }

  // メッセージ生成
  let message = '';
  if (payload.type === 'in') {
    message = `おかえり！ (今月${attendanceCount}回目の出勤)`;
  } else if (payload.type === 'out') {
    message = 'お疲れ！またね！';
  } else {
    message = '打刻完了';
  }

  return {
    ok: true,
    timestamp: Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss'),
    message: message,
    attendanceCount: attendanceCount
  };
}

/**
 * 指定UUIDの今月の出勤回数を取得
 * @param {string} uuid 
 * @return {number}
 */
function getMonthlyAttendanceCount(uuid) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(TIMESTAMP_SHEET_NAME);
  if (!sheet) return 0;

  const values = sheet.getDataRange().getValues();
  // ヘッダー行を除く
  // 列定義: 0:timestamp, 1:uuid, 4:type (in/out)

  if (values.length <= 1) return 0;

  const now = new Date();
  const currentYear = now.getFullYear();
  const currentMonth = now.getMonth(); // 0-indexed

  let count = 0;
  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    const ts = new Date(row[0]);
    const rowUuid = String(row[1]);
    const type = String(row[4]);

    if (rowUuid === uuid && type === 'in') {
      if (ts.getFullYear() === currentYear && ts.getMonth() === currentMonth) {
        count++;
      }
    }
  }
  return count;
}
