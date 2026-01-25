/**
 * テスト共通関数.gs
 * 
 * 【ファイルの役割】
 * このファイルは、ネクストエンジンAPIの動作確認を行うテスト用スクリプト（店舗マスタ取得.gsなど）
 * から共通して呼び出されるユーティリティ関数を提供します。
 * 本番用の同期処理とは独立しており、API呼び出しの結果をConsoleログに出力して確認するために使用します。
 * 
 * 【特徴】
 * - エラーハンドリング結果をログに出力
 * - 取得データの表示形式を自動調整（検索系は先頭1件、情報系は全件）
 * - トークン情報のログ出力機能
 * 
 * 作成日: 2026-01-21
 */

/**
 * ネクストエンジンAPIテスト実行用 共通関数
 * 指定されたエンドポイントに対してAPIリクエストを実行し、結果を詳細なログとして出力します。
 * 
 * 【機能詳細】
 * 1. スクリプトプロパティから認証トークンを自動取得
 * 2. fieldsパラメータの有無に応じてペイロードを構築
 * 3. API実行と結果のJSONパース
 * 4. 実行結果に応じたログ出力
 *    - 成功時: 件数、データのサンプル（または全件）、トークン更新情報
 *    - 失敗時: エラーコードとメッセージ
 * 
 * @param {string} endpoint - APIエンドポイント（例: 'api_v1_master_shop/search'）
 * @param {string|null} fields - 取得するフィールド名のカンマ区切り文字列。/infoエンドポイント等の場合はnullを指定する。
 * @param {string} logPrefix - ログ出力時に使用する機能名の接頭辞（例: '店舗マスタ' -> '店舗マスタ取得開始'）
 * @return {void} 戻り値はありません。結果はすべてLogger.logで出力されます。
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
