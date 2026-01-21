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
    // 条件: 受注状態区分が '50' (出荷確定済み) または 受注キャンセル区分が '3' (分割・統合によりキャンセル)
    const filteredData = rawData.filter(item => {
        const statusId = item.receive_order_order_status_id;
        const cancelTypeId = item.receive_order_cancel_type_id;
        return statusId === '50' || cancelTypeId === '3';
    });

    console.log(`フィルタリング後: ${filteredData.length} 件`);

    // 2. マッピング (CSVサンプルのヘッダー順に合わせる)
    // ヘッダー: 伝票番号,受注番号,出荷予定日,受注日,購入者名,支払方法,総合計,商品計,税金,発送代,手数料,他費用,ポイント数,発送伝票番号,支払区分,同梱先伝票番号,複数配送親伝票番号,分割元伝票番号,複写元伝票番号,複数配送親フラグ,受注キャンセル区分,受注キャンセル名,受注キャンセル日,受注状態区分,受注状態名
    const mappedData = filteredData.map(item => {
        return [
            item.receive_order_id,
            item.receive_order_shop_cut_form_id,
            item.receive_order_send_plan_date,
            item.receive_order_date,
            item.receive_order_purchaser_name,
            item.receive_order_payment_method_name, // 支払方法 (APIフィールドは支払名)
            item.receive_order_total_amount,
            item.receive_order_goods_amount,
            item.receive_order_tax_amount,
            item.receive_order_delivery_fee_amount,
            item.receive_order_charge_amount,
            item.receive_order_other_amount,
            item.receive_order_point_amount,
            item.receive_order_delivery_cut_form_id,
            item.receive_order_payment_method_id,   // 支払区分
            item.receive_order_include_to_order_id,
            item.receive_order_multi_delivery_parent_order_id,
            item.receive_order_divide_from_order_id,
            item.receive_order_copy_from_order_id,
            item.receive_order_multi_delivery_parent_flag,
            item.receive_order_cancel_type_id,
            item.receive_order_cancel_type_name,
            item.receive_order_cancel_date,
            item.receive_order_order_status_id,
            item.receive_order_order_status_name
        ];
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
