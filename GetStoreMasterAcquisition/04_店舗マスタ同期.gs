/**
 * 店舗マスタ同期.gs
 * 目的: ネクストエンジンから店舗マスタを取得し、スプレッドシートに書き込む
 * 
 * @param {Object} config - アプリ設定オブジェクト
 * @param {string} token - アクセストークン
 * @param {string} refreshToken - リフレッシュトークン
 */
function syncShopMaster(config, token, refreshToken) {
    Logger.log('--- 店舗マスタ処理開始 ---');

    const headerMap = {
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

    const fields = Object.keys(headerMap).join(',');

    // 共通ライブラリの関数を使用してデータ取得
    const data = nextEngineApiSearch('api_v1_master_shop/search', fields, token, refreshToken);

    if (data) {
        // 共通ライブラリの関数を使用してデータ変換・書き込み
        const array = jsonToSheetArray(data, headerMap);
        writeToSheet(array, config.SHEET_NAME_SHOP, config.SPREADSHEET_ID);
    } else {
        Logger.log('店舗マスタの取得に失敗したか、データがありませんでした。');
    }
}
