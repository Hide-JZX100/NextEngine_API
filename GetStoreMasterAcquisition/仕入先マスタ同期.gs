/**
 * 仕入先マスタ同期.gs
 * 目的: ネクストエンジンから仕入先マスタを取得し、スプレッドシートに書き込む
 * 
 * @param {Object} config - アプリ設定オブジェクト
 * @param {string} token - アクセストークン
 * @param {string} refreshToken - リフレッシュトークン
 */
function syncSupplierMaster(config, token, refreshToken) {
    Logger.log('--- 仕入先マスタ処理開始 ---');

    // 定義ファイルに基づき全てのフィールドを定義
    const headerMap = {
        'supplier_id': '仕入先コード',
        'supplier_name': '仕入先名',
        'supplier_kana': '仕入先カナ',
        'supplier_zip_code': '郵便番号',
        'supplier_address1': '住所１',
        'supplier_address2': '住所２',
        'supplier_tel': '電話番号',
        'supplier_fax': 'FAX番号',
        'supplier_mail_address': 'メールアドレス',
        'supplier_post': '部署名',
        'supplier_pic_name': '担当者名',
        'supplier_pic_kana': '担当者カナ',
        'supplier_post_zip_code': '部署郵便番号',
        'supplier_post_address1': '部署住所１',
        'supplier_post_address2': '部署住所２',
        'supplier_post_tel': '部署電話番号',
        'supplier_post_fax': '部署FAX',
        'supplier_post_mail_address': '部署メールアドレス',
        'supplier_order_id': '発注区分',
        'supplier_order_name': '発注名',
        'supplier_condition_quantity': '数量条件',
        'supplier_condition_money': '金額条件',
        'supplier_order_condition_id': '発注条件区分',
        'supplier_order_condition_name': '発注条件名',
        'supplier_order_pending_date': '発注保留日',
        'supplier_pay_out_id': '支払方法区分',
        'supplier_pay_out_name': '支払方法名',
        'supplier_closing_day': '締日',
        'supplier_usance_date': '支払サイト日',
        'supplier_order_invalid_flag': '発注無効フラグ',
        'supplier_order_forbid_date': '発注禁止日',
        'supplier_note': '備考',
        'supplier_wholesale_retail_ratio': '掛率',
        'supplier_deleted_flag': '削除フラグ',
        'supplier_creation_date': '作成日',
        'supplier_last_modified_date': '最終更新日',
        'supplier_last_modified_null_safe_date': '最終更新日(Null Safe)',
        'supplier_creator_id': '作成担当者ID',
        'supplier_creator_name': '作成担当者名',
        'supplier_last_modified_by_id': '最終更新者ID',
        'supplier_last_modified_by_null_safe_id': '最終更新者ID(Null Safe)',
        'supplier_last_modified_by_name': '最終更新者名',
        'supplier_last_modified_by_null_safe_name': '最終更新者名(Null Safe)'
    };

    const fields = Object.keys(headerMap).join(',');

    // 共通ライブラリの関数を使用してデータ取得
    const data = nextEngineApiSearch('api_v1_master_supplier/search', fields, token, refreshToken);

    if (data) {
        // 共通ライブラリの関数を使用してデータ変換・書き込み
        const array = jsonToSheetArray(data, headerMap);

        // A列(index 0: 仕入先コード)のみゼロ落ち防止のためテキスト形式で書き込む指定を追加
        writeToSheet(array, config.SHEET_NAME_SUPPLIER, config.SPREADSHEET_ID, [0]);
    } else {
        Logger.log('仕入先マスタの取得に失敗したか、データがありませんでした。');
    }
}
