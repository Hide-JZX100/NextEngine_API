/**
 * 受注キャンセル区分マスタをAPIから取得してログに出力する関数
 * 
 * 目的：受注キャンセル区分マスタ(api_v1_system_canceltype/info)の動作確認
 * 処理：APIから情報を取得し、Logger.logで出力する
 * 作成日：2026-01-19
 */
function fetchAndLogCancelTypeMaster() {
    // 1. スクリプトプロパティからアクセストークンとリフレッシュトークンを取得
    var props = PropertiesService.getScriptProperties();
    var accessToken = props.getProperty('NEXT_ENGINE_ACCESS_TOKEN');
    var refreshToken = props.getProperty('NEXT_ENGINE_REFRESH_TOKEN');

    if (!accessToken || !refreshToken) {
        Logger.log("エラー: アクセストークン(NEXT_ENGINE_ACCESS_TOKEN)またはリフレッシュトークン(NEXT_ENGINE_REFRESH_TOKEN)が取得できませんでした。");
        return;
    }

    // 2. APIリクエストの設定
    // /infoエンドポイントはfieldsパラメータが不要
    var url = 'https://api.next-engine.org/api_v1_system_canceltype/info';
    var payload = {
        'access_token': accessToken,
        'refresh_token': refreshToken,
        'wait_flag': '1'
    };

    var options = {
        'method': 'post',
        'payload': payload,
        'muteHttpExceptions': true // エラー時もレスポンスを確認できるようにする
    };

    // 3. API実行とログ出力
    try {
        Logger.log('=== 受注キャンセル区分マスタ取得開始 ===');
        Logger.log('URL: ' + url);

        var response = UrlFetchApp.fetch(url, options);
        var json = JSON.parse(response.getContentText());

        if (json.result === 'success') {
            Logger.log('=== 取得成功 ===');
            Logger.log('取得件数: ' + json.count);

            // データの中身を確認（全件表示）
            if (json.data && json.data.length > 0) {
                Logger.log('');
                Logger.log('--- 受注キャンセル区分一覧 ---');
                json.data.forEach(function (item, index) {
                    Logger.log((index + 1) + '. ID: ' + item.cancel_type_id +
                        ' | 名称: ' + item.cancel_type_name +
                        ' | 削除フラグ: ' + (item.cancel_type_deleted_flag || '0'));
                });
            } else {
                Logger.log('データが存在しません。');
            }

            // トークンが更新された場合の情報
            if (json.access_token || json.refresh_token) {
                Logger.log('');
                Logger.log('--- トークン情報 ---');
                Logger.log('新しいアクセストークンが返却されました');
                Logger.log('Access Token有効期限: ' + json.access_token_end_date);
                Logger.log('Refresh Token有効期限: ' + json.refresh_token_end_date);
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

/**
 * ヘッダーマップを確認するテスト用関数
 * 実際のスプレッドシート書き込み時に使用するヘッダーマップを確認
 */
function showCancelTypeHeaderMap() {
    var headerMap = {
        'cancel_type_id': '受注キャンセル区分ID',
        'cancel_type_name': '受注キャンセル名',
        'cancel_type_deleted_flag': '非表示フラグ'
    };

    Logger.log('=== 受注キャンセル区分 ヘッダーマップ ===');
    Logger.log(JSON.stringify(headerMap, null, 2));
    Logger.log('');
    Logger.log('フィールド数: ' + Object.keys(headerMap).length);
}
