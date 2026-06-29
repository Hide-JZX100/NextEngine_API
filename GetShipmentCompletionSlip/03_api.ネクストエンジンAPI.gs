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
    let lastTotalCount = null; // SRE的アプローチ: 総件数確認用
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
        const params = {
            'access_token': accessToken,
            'refresh_token': refreshToken,
            'wait_flag': '1',
            'fields': CONFIG.FIELDS.map(f => f.api).join(','),
            'receive_order_send_plan_date-gte': startDate,
            'receive_order_send_plan_date-lte': endDate,
            'receive_order_cancel_type_id-in': '0,3', // API側で不要なキャンセル伝票(自己都合キャンセル等)を除外
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
        const totalCount = json.count ? parseInt(json.count, 10) : null;
        if (totalCount !== null) {
            lastTotalCount = totalCount;
        }

        if (!data || data.length === 0) {
            hasMore = false;
            console.log('取得データが0件に達したためループを終了します。');
        } else {
            allData = allData.concat(data);
            console.log(`取得件数: ${data.length} (累計: ${allData.length} / 総件数: ${totalCount !== null ? totalCount : '不明'})`);

            // 終了判定の修正:
            // APIの count フィールドはシステム負荷などにより信頼できない（実際より少ない値が返る）ことがあるため、
            // 早期終了を防ぐために総件数による終了判定は行わず、データが空になるまで取得を続けます。
            
            // offsetには実際の取得件数を加算して歯抜け(スキップ)を防止
            offset += data.length;
            // API制限回避のために少し待機
            Utilities.sleep(CONFIG.API.WAIT_MS);
        }
    }

    // SRE的アプローチ: データ不整合検知（件数不足警告）
    // APIが返した総件数(lastTotalCount)よりも、実際に取得できた件数が少ない場合のみ警告を出します
    if (lastTotalCount !== null && allData.length < lastTotalCount) {
        console.warn(`⚠️ 警告: API上の総件数 (${lastTotalCount} 件) よりも実際に取得できた件数 (${allData.length} 件) が少ないです。データの欠落が発生している可能性があります。`);
    } else {
        console.log(`✅ データ整合性チェック: OK（API総件数: ${lastTotalCount || '不明'} 件に対し、${allData.length} 件取得完了）`);
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
 * 【フェーズ1検証用】指定日付範囲の伝票件数・取得ログ検証（詳細集計版）
 * 
 * 【目的】
 * APIからの全件取得ログに加え、店舗別・日付別の件数、および
 * フィルタリング（出荷確定50/同梱キャンセル3）による除外状況を集計出力します。
 * 
 * @param {string} [startDate] - 取得開始日 (YYYY-MM-DD)。省略時は '2026-05-01'
 * @param {string} [endDate] - 取得終了日 (YYYY-MM-DD)。省略時は '2026-05-31'
 */
function testVerifyDataByRange(startDate, endDate) {
    startDate = startDate || '2026-05-01';
    endDate = endDate || '2026-05-31';

    console.log(`=== 詳細データ検証開始: ${startDate} ～ ${endDate} ===`);

    let allDataCount = 0;
    let offset = 0;
    let hasMore = true;
    const LIMIT = CONFIG.API.LIMIT;

    // 集計用オブジェクト
    const rawDateMap = {};       // フィルタ前 日付別件数
    const filteredDateMap = {};  // フィルタ後 日付別件数
    const shopMap = {};          // 店舗別 累計件数（フィルタ前/後）
    const excludedStatusMap = {}; // 除外されたデータのステータス分布 (7-9日, 19-20日対象)

    const targetShortDays = ['07', '08', '09', '19', '20']; // 報告のあった日付（日）

    try {
        while (hasMore) {
            console.log(`検証データ取得中... Offset: ${offset}`);

            const props = PropertiesService.getScriptProperties();
            const accessToken = props.getProperty('ACCESS_TOKEN');
            const refreshToken = props.getProperty('REFRESH_TOKEN');

            if (!accessToken || !refreshToken) {
                throw new Error('トークンが見つかりません。認証を行ってください。');
            }

            // フィルタ条件判定に必要なフィールドも追加
            const params = {
                'access_token': accessToken,
                'refresh_token': refreshToken,
                'wait_flag': '1',
                'fields': 'receive_order_id,receive_order_send_plan_date,receive_order_shop_id,receive_order_order_status_id,receive_order_cancel_type_id',
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
            }

            const data = json.data;
            if (!data || data.length === 0) {
                hasMore = false;
            } else {
                allDataCount += data.length;

                data.forEach(item => {
                    const fullDate = item.receive_order_send_plan_date || '日付不明';
                    const shopId = item.receive_order_shop_id || '店舗不明';
                    const statusId = item.receive_order_order_status_id;
                    const cancelTypeId = item.receive_order_cancel_type_id;

                    // 1. フィルタ前集計
                    rawDateMap[fullDate] = (rawDateMap[fullDate] || 0) + 1;
                    if (!shopMap[shopId]) {
                        shopMap[shopId] = { raw: 0, filtered: 0 };
                    }
                    shopMap[shopId].raw++;

                    // 2. フィルタ条件判定
                    const isPass = cancelTypeId === CONFIG.FILTER.CANCEL_TYPE_ID_INTEGRATION ||
                        (statusId === CONFIG.FILTER.ORDER_STATUS_ID && cancelTypeId === CONFIG.FILTER.CANCEL_TYPE_ID_VALID);

                    if (isPass) {
                        filteredDateMap[fullDate] = (filteredDateMap[fullDate] || 0) + 1;
                        shopMap[shopId].filtered++;
                    } else {
                        // 3. 除外データのステータス調査 (特に7-9日, 19-20日)
                        const dayStr = fullDate.split(' ')[0].split('-')[2]; // YYYY-MM-DD から DD を抽出
                        if (targetShortDays.indexOf(dayStr) !== -1) {
                            const statusKey = `状態:${statusId}/キャンセル:${cancelTypeId}`;
                            excludedStatusMap[statusKey] = (excludedStatusMap[statusKey] || 0) + 1;
                        }
                    }
                });

                if (data.length < LIMIT) {
                    hasMore = false;
                } else {
                    offset += LIMIT;
                    Utilities.sleep(CONFIG.API.WAIT_MS);
                }
            }
        }

        console.log('=== 【検証結果集集計】 ===');
        console.log(`対象期間: ${startDate} ～ ${endDate}`);
        console.log(`API総取得（フィルタ前）: ${allDataCount} 件`);

        console.log('\n--- 1. 日付別の件数比較（フィルタ前 vs フィルタ後） ---');
        const sortedDates = Object.keys(rawDateMap).sort();
        sortedDates.forEach(date => {
            const rawCount = rawDateMap[date] || 0;
            const filteredCount = filteredDateMap[date] || 0;
            const diff = rawCount - filteredCount;
            console.log(`- ${date}: フィルタ前 ${rawCount} 件 -> フィルタ後 ${filteredCount} 件 (除外: ${diff} 件)`);
        });

        console.log('\n--- 2. 店舗別の件数内訳（フィルタ前 vs フィルタ後） ---');
        const sortedShops = Object.keys(shopMap).sort((a, b) => a - b);
        sortedShops.forEach(shopId => {
            const counts = shopMap[shopId];
            console.log(`- 店舗コード ${shopId}: フィルタ前 ${counts.raw} 件 -> フィルタ後 ${counts.filtered} 件`);
        });

        console.log('\n--- 3. 報告のあった日付（7-9日, 19-20日）で【除外された】データのステータス内訳 ---');
        const sortedStatusKeys = Object.keys(excludedStatusMap).sort();
        if (sortedStatusKeys.length === 0) {
            console.log('対象日に除外されたデータはありませんでした。');
        } else {
            sortedStatusKeys.forEach(key => {
                console.log(`- ${key}: ${excludedStatusMap[key]} 件`);
            });
        }

    } catch (e) {
        console.error(`❌ 検証中に例外エラーが発生しました: ${e.message}`);
    }
}

