/**
 * 店舗マスタをAPIから取得してログに出力する関数
 * * 目的：店舗マスタ(api_v1_master_shop/search)の動作確認
 * 処理：指定されたフィールドの情報を取得し、Logger.logで出力する
 * 作成日：2025-11-23
 */
function fetchAndLogShopMaster() {
  // 1. スクリプトプロパティからアクセストークンを取得
  // ※既存の認証システムに合わせて、トークンが期限切れの場合はリフレッシュする処理が必要であれば適宜呼び出してください
  var props = PropertiesService.getScriptProperties();
  var accessToken = props.getProperty('NEXT_ENGINE_ACCESS_TOKEN'); // プロパティ名は既存のものに合わせてください

  if (!accessToken) {
    Logger.log("エラー: アクセストークンが取得できませんでした。");
    return;
  }

  // 2. 取得するフィールドの定義 (仕様書より抜粋)
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

  // 3. APIリクエストの設定
  var url = 'https://api.next-engine.org/api_v1_master_shop/search';
  var payload = {
    'access_token': accessToken,
    'fields': fields,
    'wait_flag': '1' // 待機フラグ (在庫取得時の経験から推奨)
  };

  var options = {
    'method': 'post',
    'payload': payload,
    'muteHttpExceptions': true // エラー時もレスポンスを確認できるようにする
  };

  // 4. API実行とログ出力
  try {
    var response = UrlFetchApp.fetch(url, options);
    var json = JSON.parse(response.getContentText());

    if (json.result === 'success') {
      Logger.log('=== 取得成功 ===');
      Logger.log('取得件数: ' + json.count);
      
      // データの中身を少しだけ確認（最初の1件だけ表示）
      if (json.data && json.data.length > 0) {
        Logger.log('先頭の店舗データ: ' + JSON.stringify(json.data[0]));
      } else {
        Logger.log('データが存在しません。');
      }
      
    } else {
      Logger.log('=== エラー発生 ===');
      Logger.log('Result: ' + json.result);
      Logger.log('Code: ' + json.code);
      Logger.log('Message: ' + json.message);
    }

  } catch (e) {
    Logger.log('例外エラー: ' + e.toString());
  }
}