/**
 * モール/カートマスタをAPIから取得してログに出力する関数
 * * 目的：モールマスタ(api_v1_master_mall/search)の動作確認
 * 処理：指定されたフィールドの情報を取得し、共通関数を使ってログ出力する
 */
function fetchAndLogMallMaster() {
  // 1. 取得するフィールドの定義 (モール/カート用)
  var fields = 'mall_id,mall_name,mall_kana,mall_note,mall_country_id,mall_deleted_flag';

  // 2. 共通関数でAPIを呼び出す
  // 注: モールマスタの取得エンドポイントは店舗マスタと同じapi_v1_master_shop/searchを使用する仕様
  testApiCall('api_v1_master_shop/search', fields, 'モールマスタ');
}