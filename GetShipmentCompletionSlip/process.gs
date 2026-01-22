/**
 * 取得したデータのフィルタリングと加工を行う
 */

/**
 * 伝票データのフィルタリングとマッピングを行う
 * 
 * @param {Array<Object>} rawData - APIから取得した生の伝票データ配列
 * @return {Array<Array>} スプレッドシート出力用の2次元配列
 */
function filterAndMapSlips(rawData) {
    console.log(`=== データ処理開始: 元データ ${rawData.length} 件 ===`);

    // 1. フィルタリング
    const filteredData = rawData.filter(item => {
        const statusId = item.receive_order_order_status_id;
        const cancelTypeId = item.receive_order_cancel_type_id;
        return statusId === CONFIG.FILTER.ORDER_STATUS_ID || cancelTypeId === CONFIG.FILTER.CANCEL_TYPE_ID;
    });

    console.log(`フィルタリング後: ${filteredData.length} 件`);

    // 2. マッピング (CONFIG.FIELDSの定義順に従ってデータを抽出)
    const mappedData = filteredData.map(item => {
        return CONFIG.FIELDS.map(field => item[field.api]);
    });

    console.log(`=== データ処理終了 ===`);
    return mappedData;
}

/**
 * 【開発用】データ処理テスト
 */
function testProcessData() {
    try {
        const targetDate = getTargetDate();
        const rawResults = searchCompletedSlips(targetDate);

        if (rawResults.length === 0) {
            console.log('テストデータがないため、処理をスキップします。');
            return;
        }

        const processedResults = filterAndMapSlips(rawResults);

        if (processedResults.length > 0) {
            console.log('処理後データ1件目:', processedResults[0]);
        } else {
            console.log('フィルタリングの結果、データが0件になりました。');
        }

    } catch (e) {
        console.error('処理テスト失敗:', e.message);
    }
}
