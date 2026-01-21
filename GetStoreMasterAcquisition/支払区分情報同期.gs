/**
 * 支払区分情報同期.gs
 * 目的: ネクストエンジンから支払区分情報を取得し、スプレッドシートに書き込む
 * 
 * @param {Object} config - アプリ設定オブジェクト
 * @param {string} token - アクセストークン
 * @param {string} refreshToken - リフレッシュトークン
 */
function syncPaymentMethodMaster(config, token, refreshToken) {
    Logger.log('--- 支払区分情報同期処理開始 ---');

    const headerMap = {
        'payment_method_id': '支払区分',
        'payment_method_name': '支払名',
        'payment_method_deleted_flag': '非表示フラグ'
    };

    // /info エンドポイントはfieldsパラメータ不要のためnullを指定
    const fields = null;

    // 共通ライブラリの関数を使用してデータ取得
    const data = nextEngineApiSearch('api_v1_system_paymentmethod/info', fields, token, refreshToken);

    if (data) {
        // 共通ライブラリの関数を使用してデータ変換・書き込み
        // SHEET_NAME_PAYMENT は共通ライブラリの getAppConfig で定義済み
        const array = jsonToSheetArray(data, headerMap);
        writeToSheet(array, config.SHEET_NAME_PAYMENT, config.SPREADSHEET_ID);
    } else {
        Logger.log('支払区分情報の取得に失敗したか、データがありませんでした。');
    }
}

function testSyncPayment() {
    const config = getAppConfig();
    // アクセストークンなどはプロパティから取得
    syncPaymentMethodMaster(config, config.accessToken, config.refreshToken);
}