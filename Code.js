const TIMESTAMP_SHEET_NAME = '打刻記録';

// Firestore（AccountForm のスタッフDB）
const FIRESTORE_COLLECTION = 'staff';
const INTERN_ATTRIBUTE = 'インターン学生';
// 拠点ID → 支部名（フロントの PLACE_NAMES と対応）
const PLACE_TO_BRANCH = { tokyo: '東京', niigata: '新潟', nagoya: '名古屋', fukuoka: '福岡' };

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
          result = getStaffList(e.parameter.place);
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
          result = clearStaffCache(e.parameter.place);
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

/* ============================================================
 * Firestore アクセス（サービスアカウント + REST）
 * 認証情報はスクリプトプロパティで管理する（コード直書き禁止）:
 *   FIRESTORE_PROJECT_ID  例: amazing-centaur-468005-b0
 *   FIRESTORE_DATABASE_ID 例: accountform
 *   SA_CLIENT_EMAIL       サービスアカウントの client_email
 *   SA_PRIVATE_KEY        サービスアカウントの private_key（\n を含む全文）
 * ============================================================ */

/**
 * Firestore 接続設定をスクリプトプロパティから取得
 */
function getFirestoreConfig_() {
  const props = PropertiesService.getScriptProperties();
  const projectId = props.getProperty('FIRESTORE_PROJECT_ID');
  const databaseId = props.getProperty('FIRESTORE_DATABASE_ID') || '(default)';
  if (!projectId) {
    throw new Error('FIRESTORE_PROJECT_ID が未設定です（スクリプトプロパティ）');
  }
  return { projectId: projectId, databaseId: databaseId };
}

/**
 * Base64 URL-safe エンコード（パディング除去）
 */
function base64UrlEncode_(input) {
  return Utilities.base64EncodeWebSafe(input).replace(/=+$/, '');
}

/**
 * サービスアカウントで Firestore 用の OAuth2 アクセストークンを取得する。
 * JWT(RS256) を自前で署名し、jwt-bearer グラントで交換する。
 * 取得したトークンは約55分キャッシュする。
 */
function getFirestoreToken_() {
  const cache = CacheService.getScriptCache();
  const cached = cache.get('fs_access_token');
  if (cached) return cached;

  const props = PropertiesService.getScriptProperties();
  let clientEmail = props.getProperty('SA_CLIENT_EMAIL');
  let privateKey = props.getProperty('SA_PRIVATE_KEY');
  if (!privateKey) {
    throw new Error('サービスアカウント認証情報(SA_PRIVATE_KEY)が未設定です');
  }

  privateKey = privateKey.trim();

  // サービスアカウントJSON全体が貼られていた場合は private_key を取り出す。
  // その JSON の client_email は秘密鍵と必ず対になるため、iss 不一致による
  // "Invalid JWT Signature" を避けるべく、別途設定の SA_CLIENT_EMAIL より優先する。
  if (privateKey.charAt(0) === '{') {
    try {
      const obj = JSON.parse(privateKey);
      if (obj && obj.private_key) {
        privateKey = obj.private_key;
        if (obj.client_email) clientEmail = obj.client_email;
      }
    } catch (e) {
      // JSON でなければそのまま続行
    }
  }

  if (!clientEmail) {
    throw new Error('サービスアカウント認証情報(SA_CLIENT_EMAIL)が未設定です');
  }

  // 貼り付けゆれを正規化する:
  //  - JSON からコピーした際に付く両端の二重引用符を除去
  //  - エスケープされた改行(\r\n / \n)および CRLF を実改行へ変換
  if (privateKey.charAt(0) === '"' && privateKey.charAt(privateKey.length - 1) === '"') {
    privateKey = privateKey.slice(1, -1);
  }
  privateKey = privateKey
    .replace(/\\r\\n/g, '\n')
    .replace(/\\n/g, '\n')
    .replace(/\r\n/g, '\n');

  const now = Math.floor(Date.now() / 1000);
  const header = { alg: 'RS256', typ: 'JWT' };
  const claim = {
    iss: clientEmail,
    scope: 'https://www.googleapis.com/auth/datastore',
    aud: 'https://oauth2.googleapis.com/token',
    iat: now,
    exp: now + 3600
  };
  const toSign = base64UrlEncode_(JSON.stringify(header)) + '.' + base64UrlEncode_(JSON.stringify(claim));
  const signature = Utilities.computeRsaSha256Signature(toSign, privateKey);
  const jwt = toSign + '.' + base64UrlEncode_(signature);

  const res = UrlFetchApp.fetch('https://oauth2.googleapis.com/token', {
    method: 'post',
    contentType: 'application/x-www-form-urlencoded',
    payload: {
      grant_type: 'urn:ietf:params:oauth:grant-type:jwt-bearer',
      assertion: jwt
    },
    muteHttpExceptions: true
  });
  const body = JSON.parse(res.getContentText());
  if (res.getResponseCode() !== 200 || !body.access_token) {
    throw new Error('Firestoreアクセストークン取得に失敗: ' + res.getContentText());
  }
  cache.put('fs_access_token', body.access_token, 3300);
  return body.access_token;
}

/**
 * Firestore REST の型付き値（{stringValue:...} など）を素の値へ変換する
 */
function fsValue_(field) {
  if (field == null) return null;
  if (field.stringValue !== undefined) return field.stringValue;
  if (field.booleanValue !== undefined) return field.booleanValue;
  if (field.integerValue !== undefined) return Number(field.integerValue);
  if (field.doubleValue !== undefined) return field.doubleValue;
  if (field.timestampValue !== undefined) return field.timestampValue;
  if (field.nullValue !== undefined) return null;
  if (field.mapValue !== undefined) {
    const out = {};
    const f = (field.mapValue && field.mapValue.fields) || {};
    Object.keys(f).forEach(function (k) { out[k] = fsValue_(f[k]); });
    return out;
  }
  if (field.arrayValue !== undefined) {
    const vals = (field.arrayValue && field.arrayValue.values) || [];
    return vals.map(fsValue_);
  }
  return null;
}

/**
 * 拠点ID(place) を支部名へ変換する。未知/空なら null（=支部で絞らない）
 */
function placeToBranch_(place) {
  if (!place) return null;
  return PLACE_TO_BRANCH[place] || null;
}

/**
 * 生年月日を YYYY-MM-DD 文字列へ正規化する。
 * Firestore では文字列(YYYY-MM-DD)前提だが、timestamp 等で入っていても対応する。
 */
function normalizeBirthdate_(val) {
  if (!val) return '';
  if (typeof val === 'string') {
    // 時刻付きISO/タイムスタンプ(Firestore timestampValue 等)は、UTCの日付を
    // そのまま拾うと1日ズレるため、スクリプトTZ(Asia/Tokyo)で日付化する。
    if (val.indexOf('T') >= 0) {
      const d = new Date(val);
      if (!isNaN(d.getTime())) {
        return Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd');
      }
    }
    // 区切りを '-' に寄せ、先頭の YYYY-MM-DD を取り出す
    const s = val.replace(/\//g, '-');
    const m = s.match(/^\d{4}-\d{2}-\d{2}/);
    if (m) return m[0];
    return String(val);
  }
  try {
    return Utilities.formatDate(new Date(val), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  } catch (e) {
    return String(val);
  }
}

/**
 * Firestore staff コレクションを属性=インターン学生（＋支部）で絞り込んで取得する。
 * 戻り値は各ドキュメントの fields オブジェクトの配列。
 * @param {string|null} branch 支部名。null なら支部で絞らない。
 */
function queryStaffDocs_(branch) {
  const conf = getFirestoreConfig_();
  const token = getFirestoreToken_();
  const url = 'https://firestore.googleapis.com/v1/projects/' + conf.projectId +
    '/databases/' + conf.databaseId + '/documents:runQuery';

  const attributeFilter = {
    fieldFilter: {
      field: { fieldPath: 'attribute' },
      op: 'EQUAL',
      value: { stringValue: INTERN_ATTRIBUTE }
    }
  };

  let where;
  if (branch) {
    where = {
      compositeFilter: {
        op: 'AND',
        filters: [
          attributeFilter,
          {
            fieldFilter: {
              field: { fieldPath: 'branch' },
              op: 'EQUAL',
              value: { stringValue: branch }
            }
          }
        ]
      }
    };
  } else {
    where = attributeFilter;
  }

  const body = {
    structuredQuery: {
      from: [{ collectionId: FIRESTORE_COLLECTION }],
      where: where
    }
  };

  const res = UrlFetchApp.fetch(url, {
    method: 'post',
    contentType: 'application/json',
    headers: { Authorization: 'Bearer ' + token },
    payload: JSON.stringify(body),
    muteHttpExceptions: true
  });
  if (res.getResponseCode() !== 200) {
    throw new Error('Firestore runQuery 失敗 (' + res.getResponseCode() + '): ' + res.getContentText());
  }

  const rows = JSON.parse(res.getContentText());
  const docs = [];
  rows.forEach(function (row) {
    if (row.document && row.document.fields) {
      docs.push(row.document.fields);
    }
  });
  return docs;
}

/**
 * スタッフ一覧を取得（プルダウン用）
 * Firestore から 属性=インターン学生・現役（退職/卒業フラグなし）を取得し、
 * place 指定時はその支部のみに絞る。キャッシュで高速化する。
 * return: [{ uuid, name, birthdate, img }, ...]
 */
function getStaffList(place) {
  const branch = placeToBranch_(place);
  const cacheKey = 'staff_list_cache_' + (branch || 'all');
  const cache = CacheService.getScriptCache();
  const cachedData = cache.get(cacheKey);

  if (cachedData) {
    console.log('Using cached staff list: ' + cacheKey);
    return JSON.parse(cachedData);
  }

  console.log('Cache miss. Fetching from Firestore (branch=' + (branch || 'all') + ')');
  const fieldsList = queryStaffDocs_(branch);

  const result = [];
  fieldsList.forEach(function (f) {
    // 退職/卒業フラグが立っている人は除外（欠落/false/0 は現役扱い）
    if (fsValue_(f.graduatedFlag) === true) return;
    const uuid = fsValue_(f.staffId);
    const name = fsValue_(f.fullName);
    if (!uuid || !name) return;
    result.push({
      uuid: String(uuid),
      name: String(name),
      birthdate: normalizeBirthdate_(fsValue_(f.birthDate)),
      img: f.photoUrl ? String(fsValue_(f.photoUrl) || '') : ''
    });
  });

  // キャッシュに保存（有効期限は6時間 = 21600秒）
  try {
    cache.put(cacheKey, JSON.stringify(result), 21600);
  } catch (e) {
    console.error('Failed to put cache:', e);
  }

  return result;
}

/**
 * スタッフ一覧のキャッシュを明示的にクリアする（拠点別）
 */
function clearStaffCache(place) {
  const cache = CacheService.getScriptCache();
  const branch = placeToBranch_(place);
  const cacheKey = 'staff_list_cache_' + (branch || 'all');
  cache.remove(cacheKey);
  return { ok: true, message: 'キャッシュをクリアしました' };
}

/**
 * 生年月日とスタッフUUIDを検証
 * @param {string} uuid
 * @param {string} birthdateStr - 'YYYY-MM-DD'
 * @returns {{ok: boolean, name?: string, message?: string}}
 */
function verifyStaff(uuid, birthdateStr) {
  if (!uuid) return { ok: false, message: 'スタッフが指定されていません' };

  const conf = getFirestoreConfig_();
  const token = getFirestoreToken_();
  const url = 'https://firestore.googleapis.com/v1/projects/' + conf.projectId +
    '/databases/' + conf.databaseId + '/documents/' + FIRESTORE_COLLECTION + '/' + encodeURIComponent(uuid);

  const res = UrlFetchApp.fetch(url, {
    method: 'get',
    headers: { Authorization: 'Bearer ' + token },
    muteHttpExceptions: true
  });
  const code = res.getResponseCode();
  if (code === 404) {
    return { ok: false, message: '該当のスタッフが見つかりません' };
  }
  if (code !== 200) {
    return { ok: false, message: 'スタッフ情報の取得に失敗しました (' + code + ')' };
  }

  const doc = JSON.parse(res.getContentText());
  const fields = doc.fields || {};
  const name = String(fsValue_(fields.fullName) || '');
  const cellStr = normalizeBirthdate_(fsValue_(fields.birthDate));
  const normalizedInput = String(birthdateStr || '').trim();

  if (cellStr && cellStr === normalizedInput) {
    return { ok: true, name: name };
  }
  return { ok: false, message: '生年月日が一致しません' };
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
 *     guidelineId: string, // 選択した行動指針のID（1〜7）
 *     guidelineText: string, // 選択した行動指針の文言
 *   }
 *
 * 打刻記録シートの列順:
 *   timestamp / uuid / name / placeId / type / userAgent / guidelineId / guidelineText
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
    userAgent,
    payload.guidelineId || '',
    payload.guidelineText || ''
  ]);

  // メッセージ生成
  let message = '';
  if (payload.type === 'in') {
    message = 'おかえり！';
  } else if (payload.type === 'out') {
    message = 'お疲れ！またね！';
  } else {
    message = '打刻完了';
  }

  return {
    ok: true,
    timestamp: Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss'),
    message: message
  };
}
