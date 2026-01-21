/**
 * ネクストエンジンAPI連携 共通ライブラリ
 * 
 * 目的: GetStoreMasterAcquisitionプロジェクト内で使用する共通関数を集約
 * 作成日: 2026-01-21
 * 
 * 含まれる関数:
 * - トークン管理
 * - 設定管理
 * - API呼び出し
 * - データ変換
 * - スプレッドシート操作
 */

// ====================================================================
// トークン管理
// ====================================================================

/**
 * スクリプトプロパティから設定値を取得するヘルパー関数
 * @param {string} key - 取得するキー名
 * @return {string|null} プロパティの値
 */
function getProperty(key) {
    return PropertiesService.getScriptProperties().getProperty(key);
}

/**
 * トークンをスクリプトプロパティに保存する共通関数
 * @param {string} accessToken - アクセストークン
 * @param {string} refreshToken - リフレッシュトークン
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
// 設定管理
// ====================================================================

/**
 * スクリプトプロパティから必要な設定値を取得する
 * @return {Object} 設定値オブジェクト
 * @throws {Error} 設定値が不足している場合
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
        SHEET_NAME_PAYMENT: props.getProperty('SHEET_NAME_PAYMENT'), // 例: 支払区分情報
    };

    // 必須設定のチェック
    if (!config.accessToken || !config.refreshToken) {
        Logger.log('エラー: アクセストークン(NEXT_ENGINE_ACCESS_TOKEN)とリフレッシュトークン(NEXT_ENGINE_REFRESH_TOKEN)の両方が必要です。認証スクリプトで取得し、プロパティに設定してください。');
        throw new Error('設定エラー: 認証トークン情報が必要です。');
    }

    // スプレッドシート関連の必須設定のチェック
    if (!config.SPREADSHEET_ID || !config.SHEET_NAME_SHOP || !config.SHEET_NAME_MALL || !config.SHEET_NAME_CANCEL || !config.SHEET_NAME_PAYMENT) {
        Logger.log('エラー: スプレッドシート情報が不足しています。');
        throw new Error('設定エラー: SPREADSHEET_ID, SHEET_NAME_SHOP, SHEET_NAME_MALL, SHEET_NAME_CANCEL, SHEET_NAME_PAYMENT が必要です。');
    }

    return config;
}

// ====================================================================
// API呼び出し
// ====================================================================

/**
 * ネクストエンジンAPIを実行し、データを取得する共通関数
 * APIコール時にトークンを渡し、レスポンスに新しいトークンが含まれていれば自動更新します
 * 
 * @param {string} endpoint - APIのエンドポイント (例: 'api_v1_master_shop/search')
 * @param {string|null} fields - 取得するフィールドのカンマ区切り文字列（nullの場合は/infoエンドポイント）
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

    // fieldsが指定されている場合のみpayloadに追加
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

            // トークンエラーを含む認証エラーの場合は、認証フローの再実行を促す
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

// ====================================================================
// データ変換
// ====================================================================

/**
 * JSONデータ配列をスプレッドシート書き込み用の2次元配列に変換する
 * ヘッダーはフィールド名ではなく、日本語の項目名を使用する
 * 
 * @param {Array<Object>} data - 取得したJSONデータ配列
 * @param {Object} headerMap - {フィールド名: 項目名} のマップ
 * @return {Array<Array<any>>|null} 2次元配列 ([項目名(ヘッダー)行, データ行, ...])
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
// スプレッドシート操作
// ====================================================================

/**
 * 指定されたシートにデータを書き込む共通関数
 * 既存のデータをクリアしてから書き込む
 * 
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
