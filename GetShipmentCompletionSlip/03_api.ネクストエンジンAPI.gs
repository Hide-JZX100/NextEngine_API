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
 * 日付範囲(出荷予定日)を指定して全件取得する(ページネーション対応)
 * 
 * @param {string} startDate - 'YYYY-MM-DD' 形式の開始日
 * @param {string} endDate - 'YYYY-MM-DD' 形式の終了日
 * @return {Array<Object>} 取得した伝票データの配列
 */
function searchCompletedSlips(startDate, endDate) {
    console.log(`=== API検索開始: ${startDate} ～ ${endDate} ===`);

    let allData = [];
    let offset = 0;
    let hasMore = true;
    const LIMIT = CONFIG.API.LIMIT; // APIの1回あたりの最大取得件数

    while (hasMore) {
        console.log(`取得中... Offset: ${offset}`);

        // アクセストークンとリフレッシュトークンを取得
        const props = PropertiesService.getScriptProperties();
        const accessToken = props.getProperty('ACCESS_TOKEN');
        const refreshToken = props.getProperty('REFRESH_TOKEN');

        if (!accessToken || !refreshToken) {
            throw new Error('トークンが見つかりません。認証を行ってください。');
        }

        // パラメータ構築
        // 出荷予定日が指定範囲内であるデータを検索
        // wait_flag=1: 処理待ち等の制御用(通常指定推奨)
        const params = {
            'access_token': accessToken,
            'refresh_token': refreshToken,
            'wait_flag': '1',
            'fields': CONFIG.FIELDS.map(f => f.api).join(','),
            'receive_order_send_plan_date-gte': startDate,
            'receive_order_send_plan_date-lte': endDate,
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

        // トークンが更新された場合は保存(ネクストエンジンAPI仕様)
        if (json.access_token && json.refresh_token) {
            const props = PropertiesService.getScriptProperties();
            props.setProperties({
                'ACCESS_TOKEN': json.access_token,
                'REFRESH_TOKEN': json.refresh_token,
                'TOKEN_UPDATED_AT': new Date().getTime().toString()
            });
            console.log('APIトークンを更新しました');
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
 * 【開発用】API検索テスト
 */
function testSearchApi() {
    try {
        const dateRange = getTargetDateRange(1); // バッチ1でテスト
        const results = searchCompletedSlips(dateRange.startDate, dateRange.endDate);

        if (results.length > 0) {
            console.log('先頭データのサンプル:', JSON.stringify(results[0], null, 2));
        } else {
            console.log('データが見つかりませんでした。');
        }

    } catch (e) {
        console.error('テスト失敗:', e.message);
    }
}

/**
 * 【フェーズ1検証用】指定日付範囲の伝票件数・取得ログ検証
 * 
 * 【目的】
 * 本番トリガーに影響を与えず、特定の日付範囲（例: 2026年5月分）における
 * APIデータ取得時のoffset推移および日ごとの件数を集計してログに出力します。
 * スプレッドシートへの書き込みは行わず、console.logへの出力のみを行います。
 * 
 * @param {string} [startDate] - 取得開始日 (YYYY-MM-DD)。省略時は '2026-05-01'
 * @param {string} [endDate] - 取得終了日 (YYYY-MM-DD)。省略時は '2026-05-31'
 */
function testVerifyDataByRange(startDate, endDate) {
    startDate = startDate || '2026-05-01';
    endDate = endDate || '2026-05-31';

    console.log(`=== データ取得検証開始: ${startDate} ～ ${endDate} ===`);

    let allDataCount = 0;
    let offset = 0;
    let hasMore = true;
    const LIMIT = CONFIG.API.LIMIT; // 設定から1000件を使用
    const dateMap = {};

    try {
        while (hasMore) {
            console.log(`検証データ取得中... Offset: ${offset}`);

            // アクセストークンとリフレッシュトークンを取得
            const props = PropertiesService.getScriptProperties();
            const accessToken = props.getProperty('ACCESS_TOKEN');
            const refreshToken = props.getProperty('REFRESH_TOKEN');

            if (!accessToken || !refreshToken) {
                throw new Error('トークンが見つかりません。認証を行ってください。');
            }

            // 検証に必要な最小限のフィールドのみ取得（高速化のため）
            const params = {
                'access_token': accessToken,
                'refresh_token': refreshToken,
                'wait_flag': '1',
                'fields': 'receive_order_id,receive_order_send_plan_date,receive_order_shop_id',
                'receive_order_send_plan_date-gte': startDate,
                'receive_order_send_plan_date-lte': endDate,
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
                console.error(`❌ Offset ${offset} でAPIリクエスト失敗 (HTTP ${responseCode}):`, responseText);
                break;
            }

            const json = JSON.parse(responseText);

            if (json.result !== 'success') {
                console.error(`❌ Offset ${offset} でAPIエラー発生:`, json.message);
                break;
            }

            // トークンが更新された場合は保存
            if (json.access_token && json.refresh_token) {
                props.setProperties({
                    'ACCESS_TOKEN': json.access_token,
                    'REFRESH_TOKEN': json.refresh_token,
                    'TOKEN_UPDATED_AT': new Date().getTime().toString()
                });
                console.log('APIトークンを更新しました');
            }

            const data = json.data;
            if (!data || data.length === 0) {
                hasMore = false;
                console.log(`取得データが0件に達したためループを終了します（最終 Offset: ${offset}）`);
            } else {
                allDataCount += data.length;
                console.log(`Offset ${offset} より ${data.length} 件取得完了 (累計: ${allDataCount} 件)`);

                // 日付ごとの件数集計
                data.forEach(item => {
                    const date = item.receive_order_send_plan_date || '日付不明';
                    dateMap[date] = (dateMap[date] || 0) + 1;
                });

                if (data.length < LIMIT) {
                    hasMore = false; // 上限未満ならこれ以上データはない
                } else {
                    offset += LIMIT;
                    // API制限回避のために待機
                    Utilities.sleep(CONFIG.API.WAIT_MS);
                }
            }
        }

        console.log('=== 検証結果集計 ===');
        console.log(`対象期間: ${startDate} ～ ${endDate}`);
        console.log(`総取得件数: ${allDataCount} 件`);
        console.log('日付ごとの件数内訳（出荷予定日基準）:');

        // 日付昇順で並び替えて出力
        const sortedDates = Object.keys(dateMap).sort();
        sortedDates.forEach(date => {
            console.log(`- ${date}: ${dateMap[date]} 件`);
        });

    } catch (e) {
        console.error(`❌ 検証中に例外エラーが発生しました: ${e.message}`);
    }
}

