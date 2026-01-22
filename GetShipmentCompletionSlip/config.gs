/**
 * 設定ファイル
 * プロジェクト全体の定数や設定を管理します。
 */

/**
 * 開発モード判定ヘルパー関数
 * スクリプトプロパティ 'IS_DEV_MODE' が 'false' でない限りtrue
 * 
 * @return {boolean} 開発モードかどうか
 */
function isDevelopmentMode() {
    return PropertiesService.getScriptProperties().getProperty('IS_DEV_MODE') !== 'false';
}

const CONFIG = {
    // 開発モード設定
    DEV_MODE: {
        // スクリプトプロパティ 'IS_DEV_MODE' が 'false' でない限りtrue
        // 本番モードにするには、スクリプトプロパティに 'IS_DEV_MODE' : 'false' を設定
        ENABLED: isDevelopmentMode(),
        // 開発用固定日付(本番では昨日の日付を使用)
        TARGET_DATE: '2025/10/03'
    },

    // データフィルタリング条件
    FILTER: {
        // 受注状態区分: 出荷確定済み
        ORDER_STATUS_ID: '50',
        // 受注キャンセル区分: 分割・統合によりキャンセル
        CANCEL_TYPE_ID: '3'
    },

    // API設定
    API: {
        BASE_URL: 'https://api.next-engine.org',
        ENDPOINT_SEARCH: '/api_v1_receiveorder_base/search',
        LIMIT: 1000,                    // 1回あたりの最大取得件数
        WAIT_MS: 500                    // API制限回避の待機時間(ミリ秒)
    },

    // スプレッドシート設定
    SHEET: {
        PROPERTY_KEY_ID: 'SPREADSHEET_ID',     // スプレッドシートIDを保存しているプロパティキー
        PROPERTY_KEY_NAME: 'SHEET_NAME'        // シート(タブ)名を保存しているプロパティキー
    },

    // 取得するフィールド一覧
    // api: ネクストエンジンAPIのフィールド名
    // header: スプレッドシートのヘッダー名
    FIELDS: [
        { api: 'receive_order_id', header: '伝票番号' },
        { api: 'receive_order_shop_cut_form_id', header: '受注番号' },
        { api: 'receive_order_send_plan_date', header: '出荷予定日' },
        { api: 'receive_order_date', header: '受注日' },
        { api: 'receive_order_purchaser_name', header: '購入者名' },
        { api: 'receive_order_payment_method_name', header: '支払方法' },
        { api: 'receive_order_total_amount', header: '総合計' },
        { api: 'receive_order_goods_amount', header: '商品計' },
        { api: 'receive_order_tax_amount', header: '税金' },
        { api: 'receive_order_delivery_fee_amount', header: '発送代' },
        { api: 'receive_order_charge_amount', header: '手数料' },
        { api: 'receive_order_other_amount', header: '他費用' },
        { api: 'receive_order_point_amount', header: 'ポイント数' },
        { api: 'receive_order_delivery_cut_form_id', header: '発送伝票番号' },
        { api: 'receive_order_payment_method_id', header: '支払区分' },
        { api: 'receive_order_include_to_order_id', header: '同梱先伝票番号' },
        { api: 'receive_order_multi_delivery_parent_order_id', header: '複数配送親伝票番号' },
        { api: 'receive_order_divide_from_order_id', header: '分割元伝票番号' },
        { api: 'receive_order_copy_from_order_id', header: '複写元伝票番号' },
        { api: 'receive_order_multi_delivery_parent_flag', header: '複数配送親フラグ' },
        { api: 'receive_order_cancel_type_id', header: '受注キャンセル区分' },
        { api: 'receive_order_cancel_type_name', header: '受注キャンセル名' },
        { api: 'receive_order_cancel_date', header: '受注キャンセル日' },
        { api: 'receive_order_order_status_id', header: '受注状態区分' },
        { api: 'receive_order_order_status_name', header: '受注状態名' }
    ]
};