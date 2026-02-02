/**
 * =============================================================================
 * InventoryLogic.gs - ビジネスロジック
 * =============================================================================

 --- 内部処理・統計関連関数 ---
 @see getBatchInventoryDataWithRetry - バッチ単位で在庫データを取得・整形する内部関数です。

 */


/**
 * Old  バッチ単位で在庫情報を取得（リトライ対応版）
 * 
 * 【変更内容】
 * - getBatchStockData → getBatchStockDataWithRetry に変更
 * - その他の処理は既存のまま維持
 */

/**
 * バッチ単位で在庫情報を取得・整形
 * 
 * @param {Array} goodsCodeList - 商品コードリスト
 * @param {Object} tokens - 認証トークン
 * @param {number} batchNumber - バッチ番号
 * @return {Map} 在庫データマップ (key: 商品コード, value: InventoryDataオブジェクト)
 */
function getBatchInventoryDataWithRetry(goodsCodeList, tokens, batchNumber) {
    const inventoryDataMap = new Map();

    try {
        logWithLevel(LOG_LEVEL.DETAILED, `  在庫マスタ一括検索: ${goodsCodeList.length}件`);

        const codeMapping = new Map();
        for (const code of goodsCodeList) {
            codeMapping.set(code.toLowerCase(), code);
        }

        // ★★★ ここだけ変更: リトライ機能付き関数に置き換え ★★★
        const stockDataMap = getBatchStockDataWithRetry(goodsCodeList, tokens, batchNumber);

        logWithLevel(LOG_LEVEL.DETAILED, `  在庫マスタ取得完了: ${stockDataMap.size}件`);

        if (stockDataMap.size === 0) {
            logWithLevel(LOG_LEVEL.SUMMARY, '  在庫データが見つかりませんでした');

            logAPIErrorDetail(
                '在庫マスタAPI',
                {
                    goodsCodeCount: goodsCodeList.length,
                    firstCode: goodsCodeList[0],
                    lastCode: goodsCodeList[goodsCodeList.length - 1]
                },
                { message: 'データが1件も取得できませんでした' },
                new Error('API応答にデータが含まれていません')
            );

            return inventoryDataMap;
        }

        for (const [goodsCode, stockData] of stockDataMap) {
            const originalCode = codeMapping.get(goodsCode.toLowerCase());

            if (!originalCode) {
                logErrorDetail(goodsCode, 'コードマッピングエラー', '元のコードが見つかりません', {
                    'バッチ番号': batchNumber,
                    'API返却コード': goodsCode,
                    'マッピング数': codeMapping.size,
                    '要求コード数': goodsCodeList.length
                });
                continue;
            }

            const inventoryData = {
                goods_id: stockData.stock_goods_id,
                goods_name: '',
                stock_quantity: parseInt(stockData.stock_quantity) || 0,
                stock_allocated_quantity: parseInt(stockData.stock_allocation_quantity) || 0,
                stock_free_quantity: parseInt(stockData.stock_free_quantity) || 0,
                stock_defective_quantity: parseInt(stockData.stock_defective_quantity) || 0,
                stock_advance_order_quantity: parseInt(stockData.stock_advance_order_quantity) || 0,
                stock_advance_order_allocation_quantity: parseInt(stockData.stock_advance_order_allocation_quantity) || 0,
                stock_advance_order_free_quantity: parseInt(stockData.stock_advance_order_free_quantity) || 0,
                stock_remaining_order_quantity: parseInt(stockData.stock_remaining_order_quantity) || 0,
                stock_out_quantity: parseInt(stockData.stock_out_quantity) || 0
            };

            inventoryDataMap.set(originalCode, inventoryData);
        }

        logWithLevel(LOG_LEVEL.DETAILED, `  在庫情報構築完了: ${inventoryDataMap.size}件`);
        return inventoryDataMap;

    } catch (error) {
        logError(`在庫情報取得エラー: ${error.message}`);

        logAPIErrorDetail(
            '在庫情報構築処理',
            {
                goodsCodeCount: goodsCodeList.length,
                firstCode: goodsCodeList[0],
                lastCode: goodsCodeList[goodsCodeList.length - 1]
            },
            null,
            error
        );

        // エラー時は空または部分的なMapを返す（呼び出し元で処理続行可能にするため）
        return inventoryDataMap;
    }
}