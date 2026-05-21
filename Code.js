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
 * 打刻を記録（セッション型）
 *
 * 入室（type='in'）: 新しい行を追加。退出時間・稼働時間は空のまま。
 * 退出（type='out'）: 同一UUIDで最新の「退出時間が空」の行を探して更新。
 *                    対応する入室行が見つからない場合は退出のみ記録。
 *
 * 打刻記録シートの列順:
 *   A:入室時間 / B:UUID / C:氏名 / D:拠点 / E:退出時間 / F:稼働時間(分) / G:行動指針ID / H:行動指針テキスト
 *
 * @param {object} payload
 *   payload = {
 *     uuid: string,
 *     name: string,
 *     placeId: string,       // 'tokyo' | 'niigata' | 'nagoya' | 'fukuoka'
 *     type: string,          // 'in' | 'out'
 *     guidelineId: string,   // 選択した行動指針のID（1〜7）。退室時は空。
 *     guidelineText: string, // 選択した行動指針の文言。退室時は空。
 *   }
 */
function recordTimestamp(payload) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(TIMESTAMP_SHEET_NAME);
  if (!sheet) {
    throw new Error('「' + TIMESTAMP_SHEET_NAME + '」シートが見つかりません');
  }

  // シートが空ならヘッダー行を自動作成
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(['入室時間', 'UUID', '氏名', '拠点', '退出時間', '稼働時間(分)', '行動指針ID', '行動指針テキスト']);
  }

  const now = new Date();
  const tz = Session.getScriptTimeZone();

  if (payload.type === 'in') {
    // ── 入室: 新しい行を追加（退出時間・稼働時間は空） ──
    sheet.appendRow([
      now,                          // A: 入室時間
      payload.uuid || '',           // B: UUID
      payload.name || '',           // C: 氏名
      payload.placeId || '',        // D: 拠点
      '',                           // E: 退出時間（未入力）
      '',                           // F: 稼働時間（未入力）
      payload.guidelineId || '',    // G: 行動指針ID
      payload.guidelineText || ''   // H: 行動指針テキスト
    ]);

    return {
      ok: true,
      timestamp: Utilities.formatDate(now, tz, 'yyyy-MM-dd HH:mm:ss'),
      message: 'おかえり！'
    };

  } else if (payload.type === 'out') {
    // ── 退室: 最新の未退出行（退出時間が空）を探して更新 ──
    const data = sheet.getDataRange().getValues();

    for (let i = data.length - 1; i >= 1; i--) {
      const row = data[i];
      const rowUuid  = String(row[1]); // B列: UUID
      const checkOut = row[4];         // E列: 退出時間

      if (rowUuid === String(payload.uuid) && !checkOut) {
        // 対応する入室行を発見 → 退出時間・稼働時間を記録
        const checkIn       = new Date(row[0]);
        const minutesWorked = Math.round((now - checkIn) / 60000);
        const rowNumber     = i + 1; // スプレッドシートは1始まり

        sheet.getRange(rowNumber, 5).setValue(now);           // E: 退出時間
        sheet.getRange(rowNumber, 6).setValue(minutesWorked); // F: 稼働時間(分)

        return {
          ok: true,
          timestamp: Utilities.formatDate(now, tz, 'yyyy-MM-dd HH:mm:ss'),
          minutesWorked: minutesWorked,
          message: 'お疲れ！またね！'
        };
      }
    }

    // 対応する入室行が見つからなかった場合: 退出のみ記録
    sheet.appendRow([
      '',                           // A: 入室時間（不明）
      payload.uuid || '',           // B: UUID
      payload.name || '',           // C: 氏名
      payload.placeId || '',        // D: 拠点
      now,                          // E: 退出時間
      '',                           // F: 稼働時間（不明）
      '',                           // G: 行動指針ID
      ''                            // H: 行動指針テキスト
    ]);

    return {
      ok: true,
      timestamp: Utilities.formatDate(now, tz, 'yyyy-MM-dd HH:mm:ss'),
      message: 'お疲れ！またね！'
    };

  } else {
    // フォールバック（未知のtype）
    sheet.appendRow([
      now, payload.uuid || '', payload.name || '', payload.placeId || '',
      '', '', payload.guidelineId || '', payload.guidelineText || ''
    ]);
    return {
      ok: true,
      timestamp: Utilities.formatDate(now, tz, 'yyyy-MM-dd HH:mm:ss'),
      message: '打刻完了'
    };
  }
}

