/**
 * 店舗マスタをAPIから取得してログに出力する関数
 * * 目的：店舗マスタ(api_v1_master_shop/search)の動作確認
 * 処理：指定されたフィールドの情報を取得し、共通関数を使ってログ出力する
 */
function fetchAndLogShopMaster() {
  // 1. 取得するフィールドの定義 (仕様書より抜粋)
  var fields = 'shop_id,shop_name,shop_kana,shop_abbreviated_name,' +
               'shop_handling_goods_name,shop_close_date,shop_note,' +
               'shop_mall_id,shop_authorization_type_id,shop_authorization_type_name,' +
               'shop_tax_id,shop_tax_name,shop_currency_unit_id,' +
               'shop_currency_unit_name,shop_tax_calculation_sequence_id,' +
               'shop_type_id,shop_deleted_flag,shop_creation_date,' +
               'shop_last_modified_date,shop_last_modified_null_safe_date,' +
               'shop_creator_id,shop_creator_name,shop_last_modified_by_id,' +
               'shop_last_modified_by_null_safe_id,shop_last_modified_by_name,' +
               'shop_last_modified_by_null_safe_name';

  // 2. 共通関数（テスト共通関数.gs）を呼び出して処理を実行
  testApiCall('api_v1_master_shop/search', fields, '店舗マスタ');
}