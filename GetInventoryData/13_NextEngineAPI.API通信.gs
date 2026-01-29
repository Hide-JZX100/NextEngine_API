/**
 * =============================================================================
 * NextEngineAPI.gs - API通信・リトライ管理
 * =============================================================================

 主要な処理を実行する関数

 getBatchStockData(goodsCodeList, tokens)
 在庫マスタAPI（／api_v1_master_stock／search）を呼び出し、
 複数の商品について在庫情報（在庫数、引当数など）をまとめて取得します。
 これもstock_goods_id-inパラメータを利用して、効率的な一括検索を行います。

 updateStoredTokens(accessToken, refreshToken)
 ネクストエンジンAPIから新しいトークンが返された際に、
 スクリプトプロパティを更新し、トークン情報を保存します。

 --- 内部処理・統計関連関数 ---
 @see getBatchStockDataWithRetry     - APIへ接続し、リトライ処理を実行する中心的な内部関数です。

*/

// ============================================================================
// APIコール基本関数
// ============================================================================

/**
 * 在庫マスタデータ取得（API呼び出し）
 * @param {Array} goodsCodeList - 商品コードリスト
 * @param {Object} tokens - 認証トークン
 * @param {number} batchNumber - バッチ番号(ログ用)
 * @return {Map} 在庫データマップ (key: stock_goods_id, value: stockData)
 */
function getBatchStockData(goodsCodeList, tokens, batchNumber) {
    const url = `${NE_API_URL}/api_v1_master_stock/search`;

    const goodsIdCondition = goodsCodeList.join(',');

    const payload = {
        'access_token': tokens.accessToken,
        'refresh_token': tokens.refreshToken,
        'stock_goods_id-in': goodsIdCondition,
        'fields': 'stock_goods_id,stock_quantity,stock_allocation_quantity,stock_defective_quantity,stock_remaining_order_quantity,stock_out_quantity,stock_free_quantity,stock_advance_order_quantity,stock_advance_order_allocation_quantity,stock_advance_order_free_quantity',
        'limit': MAX_ITEMS_PER_CALL.toString()
    };

    const options = {
        'method': 'POST',
        'headers': {
            'Content-Type': 'application/x-www-form-urlencoded'
        },
        'payload': Object.keys(payload).map(key =>
            encodeURIComponent(key) + '=' + encodeURIComponent(payload[key])
        ).join('&')
    };

    const stockDataMap = new Map();

    try {
        const response = UrlFetchApp.fetch(url, options);
        const responseText = response.getContentText();
        const responseData = JSON.parse(responseText);

        if (responseData.access_token && responseData.refresh_token) {
            updateStoredTokens(responseData.access_token, responseData.refresh_token);
            tokens.accessToken = responseData.access_token;
            tokens.refreshToken = responseData.refresh_token;
        }

        if (responseData.result === 'success' && responseData.data) {
            responseData.data.forEach(stockData => {
                stockDataMap.set(stockData.stock_goods_id, stockData);
            });
            logWithLevel(LOG_LEVEL.DETAILED, `  API応答: ${responseData.data.length}件取得`);
        } else {
            logAPIErrorDetail(
                '在庫マスタAPI',
                {
                    goodsCodeCount: goodsCodeList.length,
                    firstCode: goodsCodeList[0],
                    lastCode: goodsCodeList[goodsCodeList.length - 1]
                },
                responseData,
                new Error(responseData.message || 'API呼び出しに失敗しました')
            );

            logError(`  在庫マスタAPI エラー: ${responseData.message || 'Unknown error'}`);
        }

        return stockDataMap;

    } catch (error) {
        logAPIErrorDetail(
            '在庫マスタAPI（通信エラー）',
            {
                goodsCodeCount: goodsCodeList.length,
                firstCode: goodsCodeList[0],
                lastCode: goodsCodeList[goodsCodeList.length - 1]
            },
            null,
            error
        );

        logError(`在庫マスタ一括取得エラー: ${error.message}`);
        return stockDataMap;
    }
}

/**
 * トークン更新処理
 */
function updateStoredTokens(accessToken, refreshToken) {
    const properties = PropertiesService.getScriptProperties();
    properties.setProperties({
        'ACCESS_TOKEN': accessToken,
        'REFRESH_TOKEN': refreshToken,
        'TOKEN_UPDATED_AT': new Date().getTime().toString()
    });
    console.log('トークンを更新しました');
}

// ============================================================================
// リトライ機能付きラッパー
// ============================================================================

/**
 * リトライ処理付き在庫マスタデータ取得
 * 
 * 【変更内容】
 * - 既存のgetBatchStockData関数をラップ
 * - エクスポネンシャルバックオフでリトライ
 * - リトライ統計を記録
 * 
 * @param {Array} goodsCodeList - 商品コードリスト
 * @param {Object} tokens - 認証トークン
 * @param {number} batchNumber - バッチ番号
 * @param {number} maxRetries - 最大リトライ回数
 * @return {Map} 在庫データマップ
 */
function getBatchStockDataWithRetry(goodsCodeList, tokens, batchNumber, maxRetries = RETRY_CONFIG.MAX_RETRIES) {
    // リトライ機能が無効の場合は既存関数をそのまま呼び出し
    if (!RETRY_CONFIG.ENABLE_RETRY) {
        return getBatchStockData(goodsCodeList, tokens, batchNumber);
    }

    let lastError = null;

    for (let attempt = 1; attempt <= maxRetries; attempt++) {
        try {
            // リトライ回数を記録
            recordRetryAttempt(batchNumber, attempt);

            if (attempt > 1) {
                logWithLevel(LOG_LEVEL.SUMMARY, `  リトライ ${attempt}/${maxRetries}回目...`);
            }

            // ★ 既存の関数を呼び出し
            const stockDataMap = getBatchStockData(goodsCodeList, tokens, batchNumber);

            // 成功したらデータを返す
            if (attempt > 1) {
                logWithLevel(LOG_LEVEL.SUMMARY, `  ✓ リトライ成功（${attempt}回目の試行で成功）`);
            }

            return stockDataMap;

        } catch (error) {
            lastError = error;

            // エラーの種類を判定してリトライすべきか判断
            const errorMessage = error.message.toLowerCase();

            // リトライすべきでないエラー（認証・権限系）
            if (
                errorMessage.includes('認証') ||
                errorMessage.includes('auth') ||
                errorMessage.includes('permission') ||
                errorMessage.includes('権限') ||
                errorMessage.includes('invalid') ||
                errorMessage.includes('token')
            ) {
                logError(`  即座に失敗: リトライ不可能なエラー - ${error.message}`);
                throw error; // リトライせずに即座にスロー
            }

            logError(`  ✗ API接続エラー（試行 ${attempt}/${maxRetries}）: ${error.message}`);

            // 最後の試行でなければ、待機してからリトライ
            if (attempt < maxRetries) {
                // エクスポネンシャルバックオフ（指数バックオフ）
                // 1秒、2秒、4秒...
                const waitSeconds = Math.pow(2, attempt - 1);
                logWithLevel(LOG_LEVEL.SUMMARY, `  → ${waitSeconds}秒後にリトライします...`);
                Utilities.sleep(waitSeconds * 1000);
            }
        }
    }

    // すべてのリトライが失敗した場合
    const errorMessage = `API接続失敗（${maxRetries}回試行）: ${lastError.message}`;
    logError(`  ✗✗✗ ${errorMessage}`);

    // 詳細なエラー情報を記録
    logAPIErrorDetail(
        '在庫マスタAPI（リトライ失敗）',
        {
            goodsCodeCount: goodsCodeList.length,
            firstCode: goodsCodeList[0],
            lastCode: goodsCodeList[goodsCodeList.length - 1],
            totalAttempts: maxRetries
        },
        null,
        lastError
    );

    throw new Error(errorMessage);
}