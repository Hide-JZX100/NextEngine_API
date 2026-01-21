/**
 * モールマスタ同期.gs
 * 目的: ネクストエンジンからモール/カートマスタを取得し、スプレッドシートに書き込む
 * 
 * @param {Object} config - アプリ設定オブジェクト
 * @param {string} token - アクセストークン
 * @param {string} refreshToken - リフレッシュトークン
 */
function syncMallMaster(config, token, refreshToken) {
    Logger.log('--- モールマスタ処理開始 ---');

    const headerMap = {
        'mall_id': 'モール/カートID',
        'mall_name': 'モール/カート名',
        'mall_kana': 'モール/カートカナ',
        'mall_note': '備考',
        'mall_country_id': '国ID',
        'mall_deleted_flag': '削除フラグ'
    };

    const fields = Object.keys(headerMap).join(',');

    // 共通ライブラリの関数を使用してデータ取得
    // 注: モールマスタのAPIエンドポイントも店舗と同じ api_v1_master_shop/search を使用する仕様
    const data = nextEngineApiSearch('api_v1_master_shop/search', fields, token, refreshToken);

    if (data) {
        // 共通ライブラリの関数を使用してデータ変換・書き込み
        const array = jsonToSheetArray(data, headerMap);
        writeToSheet(array, config.SHEET_NAME_MALL, config.SPREADSHEET_ID);
    } else {
        Logger.log('モールマスタの取得に失敗したか、データがありませんでした。');
    }
}
