/**
 * ネクストエンジン マスタ情報同期スクリプト (GAS)
 * 目的: 店舗マスタとモール/カートマスタをネクストエンジンAPIから取得し、
 * Googleスプレッドシートにシートを分けて自動で書き込む。
 * 実行トリガー: 毎週1回
 * * 変更履歴:
 * - 認証の仕組みを統合し、APIコール時にトークンを自動更新するロジックを追加。
 */

// ====================================================================
// 0. トークン管理ヘルパー関数
// ====================================================================

/**
 * スクリプトプロパティから設定値を取得するヘルパー関数
 * @param {string} key - 取得するキー名
 */
function getProperty(key) {
  return PropertiesService.getScriptProperties().getProperty(key);
}

/**
 * トークンをスクリプトプロパティに保存する共通関数
 * @param {string} accessToken 
 * @param {string} refreshToken 
 */
function saveTokens(accessToken, refreshToken) {
  const props = PropertiesService.getScriptProperties();
  const data = {};

  if (accessToken) data['NEXT_ENGINE_ACCESS_TOKEN'] = accessToken;
  if (refreshToken) data['NEXT_ENGINE_REFRESH_TOKEN'] = refreshToken;

  props.setProperties(data);
  Logger.log('スクリプトプロパティに新しいアクセストークンとリフレッシュトークンを保存しました。');
}

// ====================================================================
// 1. 環境設定の定義
// ====================================================================

/**
 * スクリプトプロパティから必要な設定値を取得する。
 * @return {Object} 設定値オブジェクト
 */
function getAppConfig() {
  const props = PropertiesService.getScriptProperties();

  // 必須項目: スクリプトプロパティに設定してください
  const config = {
    // 認証情報
    accessToken: props.getProperty('NEXT_ENGINE_ACCESS_TOKEN'),
    refreshToken: props.getProperty('NEXT_ENGINE_REFRESH_TOKEN'),
    // スプレッドシート情報
    SPREADSHEET_ID: props.getProperty('SPREADSHEET_ID'),
    SHEET_NAME_SHOP: props.getProperty('SHEET_NAME_SHOP'), // 例: 店舗マスタ
    SHEET_NAME_MALL: props.getProperty('SHEET_NAME_MALL'), // 例: モールマスタ
    SHEET_NAME_CANCEL: props.getProperty('SHEET_NAME_CANCEL'), // 例: 受注キャンセル区分
  };

  // 必須設定のチェック
  if (!config.accessToken || !config.refreshToken) {
    Logger.log('エラー: アクセストークン(NEXT_ENGINE_ACCESS_TOKEN)とリフレッシュトークン(NEXT_ENGINE_REFRESH_TOKEN)の両方が必要です。認証スクリプトで取得し、プロパティに設定してください。');
    throw new Error('設定エラー: 認証トークン情報が必要です。');
  }

  // スプレッドシート関連の必須設定のチェック
  if (!config.SPREADSHEET_ID || !config.SHEET_NAME_SHOP || !config.SHEET_NAME_MALL || !config.SHEET_NAME_CANCEL) {
    Logger.log('エラー: スプレッドシート情報が不足しています。');
    throw new Error('設定エラー: SPREADSHEET_ID, SHEET_NAME_SHOP, SHEET_NAME_MALL, SHEET_NAME_CANCEL が必要です。');
  }

  return config;
}


// ====================================================================
// 2. データ取得共通処理
// ====================================================================

/**
 * ネクストエンジンAPIを実行し、データを取得する共通関数。
 * ★APIコール時にトークンを渡し、レスポンスに新しいトークンが含まれていれば自動更新します。
 * @param {string} endpoint - APIのエンドポイント (例: 'api_v1_master_shop/search')
 * @param {string} fields - 取得するフィールドのカンマ区切り文字列
 * @param {string} accessToken - アクセストークン
 * @param {string} refreshToken - リフレッシュトークン
 * @return {Array<Object>|null} 取得したデータ配列、またはエラーの場合はnull
 */
function nextEngineApiSearch(endpoint, fields, accessToken, refreshToken) {
  const url = `https://api.next-engine.org/${endpoint}`;

  const payload = {
    'access_token': accessToken,
    'refresh_token': refreshToken, // トークン自動更新のため毎回付与
    'wait_flag': '1' // 処理を待機
  };

  // ★改修ポイント: fieldsが指定されている場合のみpayloadに追加
  // /infoエンドポイント（fieldsが不要）にも対応できるようにする
  if (fields) {
    payload['fields'] = fields;
  }

  const options = {
    'method': 'post',
    'payload': payload,
    'muteHttpExceptions': true
  };

  Logger.log(`APIコール開始: ${endpoint} (フィールド数: ${fields ? fields.split(',').length : 'N/A - info endpoint'})`);

  try {
    const response = UrlFetchApp.fetch(url, options);
    const json = JSON.parse(response.getContentText());

    // ★重要: APIレスポンスに新しいトークンが含まれていれば保存し直す
    if (json.access_token || json.refresh_token) {
      // 新しいリフレッシュトークンが返ってこない可能性も考慮し、
      // 返ってきたトークンを優先し、返ってこなければ現在のトークンを維持して保存
      const newAccessToken = json.access_token || accessToken;
      const newRefreshToken = json.refresh_token || refreshToken;
      saveTokens(newAccessToken, newRefreshToken);
    }

    if (json.result === 'success') {
      Logger.log(`取得成功。件数: ${json.count}`);
      return json.data || []; // データがない場合も空配列を返す
    } else {
      Logger.log('=== APIエラー発生 ===');
      Logger.log(`Result: ${json.result}`);
      Logger.log(`Code: ${json.code}`);
      Logger.log(`Message: ${json.message}`);

      // トークンエラーを含む認証エラーの場合は、認証フローの再実行を促す。
      if (json.code === 'C005' || json.code === '002003' || json.code === '002004') {
        Logger.log('致命的な認証エラーが発生しました。リフレッシュトークンも無効化された可能性があります。認証スクリプトを再実行してください。');
      }

      return null;
    }

  } catch (e) {
    Logger.log('例外エラー: ' + e.toString());
    return null;
  }
}

/**
 * JSONデータ配列をスプレッドシート書き込み用の2次元配列に変換する。
 * ヘッダーはフィールド名ではなく、日本語の項目名を使用する。
 * @param {Array<Object>} data - 取得したJSONデータ配列
 * @param {Object} headerMap - {フィールド名: 項目名} のマップ
 * @return {Array<Array<any>>} 2次元配列 ([項目名(ヘッダー)行, データ行, ...])
 */
function jsonToSheetArray(data, headerMap) {
  if (!data || data.length === 0 || !headerMap || Object.keys(headerMap).length === 0) {
    return null;
  }

  // APIフィールド名 (JSONキー) の配列
  const apiFields = Object.keys(headerMap);

  // 日本語項目名 (スプレッドシートのヘッダー) の配列
  const japaneseHeaders = apiFields.map(field => headerMap[field]);

  const result = [japaneseHeaders]; // 1行目は日本語ヘッダー

  data.forEach(item => {
    // APIフィールド名に対応する値をJSONデータから取得して行を構築
    const row = apiFields.map(field => item[field.trim()]);
    result.push(row);
  });

  return result;
}


// ====================================================================
// 3. スプレッドシートへの書き込み処理
// ====================================================================

/**
 * 指定されたシートにデータを書き込む共通関数。
 * 既存のデータをクリアしてから書き込む。
 * @param {Array<Array<any>>} sheetArray - 書き込む2次元配列データ
 * @param {string} sheetName - 書き込み対象のシート名
 * @param {string} spreadsheetId - スプレッドシートID
 */
function writeToSheet(sheetArray, sheetName, spreadsheetId) {
  if (!sheetArray || sheetArray.length === 0) {
    Logger.log(`警告: ${sheetName} に書き込むデータがありません。処理をスキップします。`);
    return;
  }

  try {
    const ss = SpreadsheetApp.openById(spreadsheetId);
    let sheet = ss.getSheetByName(sheetName);

    if (!sheet) {
      // シートが存在しない場合は新規作成
      sheet = ss.insertSheet(sheetName);
      Logger.log(`シート「${sheetName}」を新規作成しました。`);
    } else {
      // 既存のデータをクリア (ヘッダー行はデータに含まれているため全クリアでOK)
      sheet.clearContents();
    }

    // 書き込み範囲を設定
    const numRows = sheetArray.length;
    const numCols = sheetArray[0].length;
    const range = sheet.getRange(1, 1, numRows, numCols);

    // データ書き込み
    range.setValues(sheetArray);

    // ヘッダー行を固定・装飾 (オプション)
    sheet.setFrozenRows(1);
    sheet.getRange('A1:' + sheet.getRange(1, numCols).getA1Notation()).setBackground('#f0f0f0').setFontWeight('bold');

    Logger.log(`「${sheetName}」に ${numRows - 1} 件のデータを書き込みました。`);

  } catch (e) {
    Logger.log(`スプレッドシート書き込み中にエラーが発生しました (${sheetName}): ` + e.toString());
  }
}

// ====================================================================
// 4. メイン実行関数 (トリガー設定用)
// ====================================================================

/**
 * 週間トリガーに設定するメイン関数。
 * 店舗マスタとモール/カートマスタの取得とスプレッドシートへの書き込みを実行する。
 */
function mainMasterSync() {
  Logger.log('--- マスタ情報同期処理 開始 ---');

  // 1. 設定値の取得
  const config = getAppConfig();
  const token = config.accessToken;
  const refreshToken = config.refreshToken; // リフレッシュトークンを取得

  // --- 店舗マスタの処理 ---
  const shopHeaderMap = {
    'shop_id': '店舗ID',
    'shop_name': '店舗名',
    'shop_kana': '店舗名カナ',
    'shop_abbreviated_name': '店舗略名',
    'shop_handling_goods_name': '取扱商品名',
    'shop_close_date': '閉店日',
    'shop_note': '備考',
    'shop_mall_id': 'モールID',
    'shop_authorization_type_id': 'オーソリ区分ID',
    'shop_authorization_type_name': 'オーソリ区分名',
    'shop_tax_id': '税区分ID',
    'shop_tax_name': '税区分名',
    'shop_currency_unit_id': '通貨単位区分ID',
    'shop_currency_unit_name': '通貨単位区分名',
    'shop_tax_calculation_sequence_id': '税計算順序',
    'shop_type_id': '後払い.com サイトID',
    'shop_deleted_flag': '削除フラグ',
    'shop_creation_date': '作成日',
    'shop_last_modified_date': '最終更新日',
    'shop_last_modified_null_safe_date': '最終更新日(Null Safe)',
    'shop_creator_id': '作成担当者ID',
    'shop_creator_name': '作成担当者名',
    'shop_last_modified_by_id': '最終更新者ID',
    'shop_last_modified_by_null_safe_id': '最終更新者ID(Null Safe)',
    'shop_last_modified_by_name': '最終更新者名',
    'shop_last_modified_by_null_safe_name': '最終更新者名(Null Safe)'
  };

  const shopFields = Object.keys(shopHeaderMap).join(',');

  const shopData = nextEngineApiSearch('api_v1_master_shop/search', shopFields, token, refreshToken);

  if (shopData) {
    const shopArray = jsonToSheetArray(shopData, shopHeaderMap);
    writeToSheet(shopArray, config.SHEET_NAME_SHOP, config.SPREADSHEET_ID);
  }

  // --- モール/カートマスタの処理 ---
  const mallHeaderMap = {
    'mall_id': 'モール/カートID',
    'mall_name': 'モール/カート名',
    'mall_kana': 'モール/カートカナ',
    'mall_note': '備考',
    'mall_country_id': '国ID',
    'mall_deleted_flag': '削除フラグ'
  };

  const mallFields = Object.keys(mallHeaderMap).join(',');

  const mallData = nextEngineApiSearch('api_v1_master_shop/search', mallFields, token, refreshToken);

  if (mallData) {
    const mallArray = jsonToSheetArray(mallData, mallHeaderMap);
    writeToSheet(mallArray, config.SHEET_NAME_MALL, config.SPREADSHEET_ID);
  }

  // --- 受注キャンセル区分マスタの処理 ---
  syncCancelTypeMaster(config, token, refreshToken);

  Logger.log('--- マスタ情報同期処理 完了 ---');
}