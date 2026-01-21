/**
 * テスト用共通関数
 * 
 * 目的: 各マスタのAPIテスト用スクリプトで使用する共通処理を集約
 * 作成日: 2026-01-21
 * 
 * 含まれる関数:
 * - testApiCall: API呼び出しとログ出力の共通処理
 */

/**
 * ネクストエンジンAPIをテスト用に呼び出す共通関数
 * APIから取得したデータをログに出力する
 * 
 * @param {string} endpoint - APIエンドポイント（例: 'api_v1_master_shop/search'）
 * @param {string|null} fields - 取得フィールド（nullの場合は/infoエンドポイント）
 * @param {string} logPrefix - ログ接頭辞（例: '店舗マスタ'）
 */
function testApiCall(endpoint, fields, logPrefix) {
    // 1. スクリプトプロパティからトークンを取得
    const props = PropertiesService.getScriptProperties();
    const accessToken = props.getProperty('NEXT_ENGINE_ACCESS_TOKEN');
    const refreshToken = props.getProperty('NEXT_ENGINE_REFRESH_TOKEN');

    if (!accessToken || !refreshToken) {
        Logger.log(`エラー: アクセストークン(NEXT_ENGINE_ACCESS_TOKEN)またはリフレッシュトークン(NEXT_ENGINE_REFRESH_TOKEN)が取得できませんでした。`);
        return;
    }

    // 2. APIリクエストの設定
    const url = `https://api.next-engine.org/${endpoint}`;
    const payload = {
        'access_token': accessToken,
        'refresh_token': refreshToken,
        'wait_flag': '1'
    };

    // fieldsが指定されている場合のみpayloadに追加
    if (fields) {
        payload['fields'] = fields;
    }

    const options = {
        'method': 'post',
        'payload': payload,
        'muteHttpExceptions': true
    };

    // 3. API実行とログ出力
    try {
        Logger.log(`=== ${logPrefix}取得開始 ===`);
        Logger.log(`URL: ${url}`);
        if (fields) {
            Logger.log(`フィールド数: ${fields.split(',').length}`);
        } else {
            Logger.log(`エンドポイントタイプ: /info（fieldsパラメータ不要）`);
        }

        const response = UrlFetchApp.fetch(url, options);
        const json = JSON.parse(response.getContentText());

        if (json.result === 'success') {
            Logger.log('=== 取得成功 ===');
            Logger.log(`取得件数: ${json.count}`);

            // データの中身を確認
            if (json.data && json.data.length > 0) {
                // /infoエンドポイントの場合は全件表示、それ以外は先頭1件のみ
                if (!fields) {
                    // 全件表示（受注キャンセル区分など）
                    Logger.log('');
                    Logger.log(`--- ${logPrefix}一覧 ---`);
                    json.data.forEach((item, index) => {
                        Logger.log(`${index + 1}. ${JSON.stringify(item)}`);
                    });
                } else {
                    // 先頭1件のみ表示
                    Logger.log(`先頭の${logPrefix}データ: ${JSON.stringify(json.data[0])}`);
                }
            } else {
                Logger.log('データが存在しません。');
            }

            // トークンが更新された場合の情報
            if (json.access_token || json.refresh_token) {
                Logger.log('');
                Logger.log('--- トークン情報 ---');
                Logger.log('新しいアクセストークンが返却されました');
                if (json.access_token_end_date) {
                    Logger.log(`Access Token有効期限: ${json.access_token_end_date}`);
                }
                if (json.refresh_token_end_date) {
                    Logger.log(`Refresh Token有効期限: ${json.refresh_token_end_date}`);
                }
            }

        } else {
            Logger.log('=== エラー発生 ===');
            Logger.log(`Result: ${json.result}`);
            Logger.log(`Code: ${json.code}`);
            Logger.log(`Message: ${json.message}`);
        }

    } catch (e) {
        Logger.log('例外エラー: ' + e.toString());
    }
}
