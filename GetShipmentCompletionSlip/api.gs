/**
 * API連携処理ファイル (GetShipmentCompletionSlip)
 * 
 * 【目的】
 * ネクストエンジンAPIとの通信処理を一元管理するファイルです。
 * データの検索、取得、ページネーション処理、およびAPI制限のハンドリングを行います。
 * ネクストエンジンAPIのエンドポイントやパラメータ構築のロジックをここに集約しています。
 * 
 * 【含まれる関数】
 * - searchCompletedSlips(targetDate): 指定された日付(出荷予定日)の伝票を検索し、全件取得するメイン関数
 * - getAccessToken_(): スクリプトプロパティから認証用トークンを取得する内部ヘルパー関数
 * - testSearchApi(): API経由でのデータ取得動作を確認するための開発用テスト関数
 * 
 * 【依存関係】
 * - config.gs: APIエンドポイントやフィールド設定を参照
 * - NE_認証ライブラリ使用必須関数.gs: 認証トークンの管理（本ファイルでは直接呼び出さないが、トークンが保存されていることが前提）
 */

/**
 * 出荷完了伝票を検索して取得する
 * 対象日(出荷予定日)を指定して全件取得する(ページネーション対応)
 * 
 * @param {string} targetDate - 'YYYY-MM-DD' 形式の対象日
 * @return {Array<Object>} 取得した伝票データの配列
 */
function searchCompletedSlips(targetDate) {
    console.log(`=== API検索開始: 対象日 ${targetDate} ===`);

    const accessToken = getAccessToken_();
    if (!accessToken) {
        throw new Error('アクセストークンが取得できませんでした。認証を行ってください。');
    }

    let allData = [];
    let offset = 0;
    let hasMore = true;
    const LIMIT = CONFIG.API.LIMIT; // APIの1回あたりの最大取得件数

    while (hasMore) {
        console.log(`取得中... Offset: ${offset}`);

        // パラメータ構築
        // 出荷予定日が対象日であるデータを検索
        // wait_flag=1: 処理待ち等の制御用(通常指定推奨)
        const params = {
            'access_token': accessToken,
            'wait_flag': '1',
            'fields': CONFIG.FIELDS.map(f => f.api).join(','),
            'receive_order_send_plan_date-eq': targetDate,
            'limit': LIMIT.toString(),
            'offset': offset.toString()
        };

        const url = CONFIG.API.BASE_URL + CONFIG.API.ENDPOINT_SEARCH;
        const options = {
            'method': 'post',
            'payload': params,
            'muteHttpExceptions': true
        };

        const response = UrlFetchApp.fetch(url, options);
        const responseCode = response.getResponseCode();
        const responseText = response.getContentText();

        if (responseCode !== 200) {
            // トークン期限切れの可能性などを考慮してエラーハンドリング
            console.error('API Error Response:', responseText);
            throw new Error(`APIリクエスト失敗 (Code: ${responseCode}): ${responseText}`);
        }

        const json = JSON.parse(responseText);

        if (json.result !== 'success') {
            throw new Error(`APIエラー: ${json.message}`);
        }

        const data = json.data;
        if (!data || data.length === 0) {
            hasMore = false;
            console.log('件数: 0 (終了)');
        } else {
            allData = allData.concat(data);
            console.log(`取得件数: ${data.length} (累計: ${allData.length})`);

            if (data.length < LIMIT) {
                hasMore = false; // 上限未満ならこれ以上データはない
            } else {
                offset += LIMIT;
                // API制限回避のために少し待機
                Utilities.sleep(CONFIG.API.WAIT_MS);
            }
        }
    }

    console.log(`=== API検索終了: 合計 ${allData.length} 件 ===`);
    return allData;
}

/**
 * スクリプトプロパティからアクセストークンを取得する
 * 
 * @return {string|null} アクセストークン
 */
function getAccessToken_() {
    const props = PropertiesService.getScriptProperties();
    return props.getProperty('ACCESS_TOKEN');
}

/**
 * 【開発用】API検索テスト
 */
function testSearchApi() {
    try {
        const targetDate = getTargetDate();
        const results = searchCompletedSlips(targetDate);

        if (results.length > 0) {
            console.log('先頭データのサンプル:', JSON.stringify(results[0], null, 2));
        } else {
            console.log('データが見つかりませんでした。');
        }

    } catch (e) {
        console.error('テスト失敗:', e.message);
    }
}
