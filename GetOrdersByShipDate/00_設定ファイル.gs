/**
 * @file 00_設定ファイル.gs
 * @description Phase 1：基盤構築。NextEngine API連携に必要な定数定義および設定情報の読み込み処理を管理します。
 */

/**
 * 定数定義（変更しない固定値）
 */
const NE_API_BASE_URL = 'https://api.next-engine.org';
const NE_ENDPOINT_ORDER_ROW = '/api_v1_receiveorder_row/search';
const LIMIT = 1000;
const SHOP_CODE_COLUMN = 1; // 店舗情報シートの店舗コード列番号
const SHOP_NAME_COLUMN = 2; // 店舗情報シートの店舗名列番号

/**
 * スクリプトプロパティ（設定が必要なもの）
 * 
 * キー名                  | 必須 | 説明
 * ------------------------|------|---------------------------
 * ACCESS_TOKEN            | 必須 | NEAuthライブラリが管理
 * REFRESH_TOKEN           | 必須 | NEAuthライブラリが管理
 * TARGET_SPREADSHEET_ID   | 必須 | データ書き込み先スプレッドシートID
 * MASTER_SPREADSHEET_ID   | 必須 | 店舗情報マスタのスプレッドシートID
 * NOTIFICATION_EMAIL      | 必須 | エラー通知先メールアドレス
 * SHEET_NAME_ORDERS       | 任意 | 省略時は '受注情報'
 * SHEET_NAME_CANCEL       | 任意 | 省略時は 'キャンセル情報'
 * SHEET_NAME_SHOP         | 任意 | 省略時は '店舗情報'
 */

/**
 * スクリプトプロパティから設定値を読み込む
 * @return {Object} 設定情報のオブジェクト
 * @throws {Error} 必須項目(TARGET_SPREADSHEET_ID, MASTER_SPREADSHEET_ID, NOTIFICATION_EMAIL)が未設定の場合
 */
function getConfig() {
  const props = PropertiesService.getScriptProperties().getProperties();
  
  const requiredKeys = ['TARGET_SPREADSHEET_ID', 'MASTER_SPREADSHEET_ID', 'NOTIFICATION_EMAIL'];
  const missingKeys = requiredKeys.filter(key => !props[key]);
  
  if (missingKeys.length > 0) {
    throw new Error('以下の必須スクリプトプロパティが未設定です: ' + missingKeys.join(', '));
  }
  
  return {
    targetSpreadsheetId: props['TARGET_SPREADSHEET_ID'],
    masterSpreadsheetId: props['MASTER_SPREADSHEET_ID'],
    notificationEmail: props['NOTIFICATION_EMAIL'],
    sheetNameOrders: props['SHEET_NAME_ORDERS'] || '受注情報',
    sheetNameCancel: props['SHEET_NAME_CANCEL'] || 'キャンセル情報',
    sheetNameShop: props['SHEET_NAME_SHOP'] || '店舗情報'
  };
}

/**
 * 受注情報フィールド定義
 */
const ORDER_FIELDS = [
  'receive_order_send_plan_date',
  'receive_order_row_goods_id',
  'receive_order_row_goods_name',
  'receive_order_row_quantity',
  'receive_order_row_sub_total_price',
  'receive_order_row_unit_price',
  'receive_order_row_receive_order_id',
  'receive_order_row_shop_cut_form_id',
  'receive_order_date',
  'receive_order_payment_method_name',
  'receive_order_total_amount',
  'receive_order_purchaser_name',
  'receive_order_purchaser_kana',
  'receive_order_purchaser_tel',
  'receive_order_purchaser_zip_code',
  'receive_order_purchaser_address1',
  'receive_order_purchaser_address2',
  'receive_order_purchaser_mail_address',
  'receive_order_consignee_name',
  'receive_order_consignee_kana',
  'receive_order_consignee_tel',
  'receive_order_consignee_zip_code',
  'receive_order_consignee_address1',
  'receive_order_consignee_address2',
  'receive_order_shop_id',
  'receive_order_picking_instruct',
  'receive_order_delivery_name',
  'receive_order_delivery_cut_form_id'
].join(',');

/**
 * キャンセル情報フィールド定義
 */
const CANCEL_FIELDS = [
  'receive_order_row_receive_order_id',
  'receive_order_cancel_date',
  'receive_order_row_goods_id',
  'receive_order_row_goods_name',
  'receive_order_row_quantity',
  'receive_order_row_unit_price',
  'receive_order_delivery_fee_amount',
  'receive_order_date',
  'receive_order_shop_id'
].join(',');

/**
 * 受注情報ヘッダー定義
 */
const ORDER_HEADERS = [
  '出荷予定日','ロケーションコード','商品コード','商品名','受注数','小計','売単価',
  '伝票番号','受注番号','受注日','支払方法','総合計',
  '購入者名','購入者カナ','購入者電話番号','購入者郵便番号','購入者（住所1+住所2）',
  '購入者メールアドレス','送り先名','送り先カナ','送り先電話番号','送り先郵便番号',
  '送り先（住所1+住所2）','店舗コード','ピッキング指示','発送方法','発送伝票番号'
];

/**
 * キャンセル情報ヘッダー定義
 */
const CANCEL_HEADERS = [
  '店舗名','伝票番号','キャンセル日','商品コード（伝票）','商品名（伝票）',
  '受注数','小計','発送代','受注日','店舗コード'
];
