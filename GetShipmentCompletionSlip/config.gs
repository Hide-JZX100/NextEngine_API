/**
 * 設定ファイル
 * プロジェクト全体の定数や設定を管理します。
 */
const CONFIG = {
    // 開発モード設定 (スクリプトプロパティ 'IS_DEV_MODE' が 'false' でない限り、デフォルトtrue)
    // スクリプトプロパティに 'IS_DEV_MODE' : 'false' を設定すると本番モード（昨日分取得）になります。
    IS_DEV_MODE: (PropertiesService.getScriptProperties().getProperty('IS_DEV_MODE') !== 'false'),

    // 開発用固定日付 (本番環境では昨日の日付になりますが、開発中は特定の過去日を指定)
    DEV_TARGET_DATE: '2025/10/03',

    // API設定
    API: {
        BASE_URL: 'https://api.next-engine.org',
        ENDPOINT_SEARCH: '/api_v1_receiveorder_base/search',
    },

    // スプレッドシート設定
    SHEET: {
        PROPERTY_KEY_ID: 'SPREADSHEET_ID',   // スプレッドシートIDを保存しているプロパティキー
        PROPERTY_KEY_NAME: 'SHEET_NAME'      // シート(タブ)名を保存しているプロパティキー
    },

    // 取得するフィールド一覧 (CSVサンプルと受注伝票検索.txtに基づく)
    FIELDS: [
        'receive_order_id',                           // 伝票番号
        'receive_order_shop_cut_form_id',             // 受注番号
        'receive_order_send_plan_date',               // 出荷予定日
        'receive_order_date',                         // 受注日
        'receive_order_purchaser_name',               // 購入者名
        'receive_order_payment_method_name',          // 支払名 (CSV等では支払方法)
        'receive_order_total_amount',                 // 総合計
        'receive_order_goods_amount',                 // 商品計
        'receive_order_tax_amount',                   // 税金
        'receive_order_delivery_fee_amount',          // 発送代
        'receive_order_charge_amount',                // 手数料
        'receive_order_other_amount',                 // 他費用
        'receive_order_point_amount',                 // ポイント数
        'receive_order_delivery_cut_form_id',         // 発送伝票番号
        'receive_order_payment_method_id',            // 支払区分
        'receive_order_include_to_order_id',          // 同梱先伝票番号
        'receive_order_multi_delivery_parent_order_id', // 複数配送親伝票番号
        'receive_order_divide_from_order_id',         // 分割元伝票番号
        'receive_order_copy_from_order_id',           // 複写元伝票番号
        'receive_order_multi_delivery_parent_flag',   // 複数配送親フラグ
        'receive_order_cancel_type_id',               // 受注キャンセル区分
        'receive_order_cancel_type_name',             // 受注キャンセル名
        'receive_order_cancel_date',                  // 受注キャンセル日
        'receive_order_order_status_id',              // 受注状態区分
        'receive_order_order_status_name'             // 受注状態名
    ]
};
