/**
 * モール/カートマスタをAPIから取得してログに出力する関数
 * * 目的：モールマスタ(api_v1_master_mall/search)の動作確認
 * 処理：指定されたフィールドの情報を取得し、Logger.logで出力する
 * 作成日：2025-xx-xx
 */
function fetchAndLogMallMaster() {
  // 1. スクリプトプロパティからアクセストークンを取得
  var props = PropertiesService.getScriptProperties();
  // 修正: キー名を 'NEXT_ENGINE_ACCESS_TOKEN' に変更
  var accessToken = props.getProperty('NEXT_ENGINE_ACCESS_TOKEN');

  if (!accessToken) {
    Logger.log("エラー: アクセストークン(NEXT_ENGINE_ACCESS_TOKEN)が取得できませんでした。");
    return;
  }

  // 2. 取得するフィールドの定義 (モール/カート用)
  var fields = 'mall_id,mall_name,mall_kana,mall_note,mall_country_id,mall_deleted_flag';

  // 3. APIリクエストの設定
  // 店舗マスタが shop なので、モールは mall と推測されます
  var url = 'https://api.next-engine.org/api_v1_master_mall/search';
  var payload = {
    'access_token': accessToken,
    'fields': fields,
    'wait_flag': '1'
  };

  var options = {
    'method': 'post',
    'payload': payload,
    'muteHttpExceptions': true
  };

  // 4. API実行とログ出力
  try {
    var response = UrlFetchApp.fetch(url, options);
    var json = JSON.parse(response.getContentText());

    if (json.result === 'success') {
      Logger.log('=== モール取得成功 ===');
      Logger.log('取得件数: ' + json.count);
      
      // データの中身を確認
      if (json.data && json.data.length > 0) {
        Logger.log('先頭のモールデータ: ' + JSON.stringify(json.data[0]));
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