/**
 * ネクストエンジンAPI接続のウォームアップ（受注明細検索版）
 *
 * 【目的】
 * 本番の updateShippingData 実行前に、同一エンドポイント・同一フィールドで
 * 1件だけ取得することで、ネクストエンジン側の3テーブル結合キャッシュを
 * 事前に温め、初回実行時の応答遅延を軽減する。
 *
 * 【動作】
 * - 出荷予定日を当日に絞り込み、limit=1 で受注明細検索APIを呼び出す
 * - 本番（fetchAllShippingData）と同一のフィールド・フィルタ条件を使用する
 * - トークンが更新された場合はスクリプトプロパティを更新する
 * - エラーが発生しても本番処理を阻害しないようログ出力に留める
 *
 * 【トリガー設定】
 * 本番実行（7:56）の約20分前、7:35 に実行すること
 */
function warmUpNextEngineConnection() {
    console.log('=== ウォームアップ開始 ===');

    try {
        const properties = PropertiesService.getScriptProperties();
        const accessToken = properties.getProperty('ACCESS_TOKEN');
        const refreshToken = properties.getProperty('REFRESH_TOKEN');

        if (!accessToken || !refreshToken) {
            console.warn('トークンが見つからないため、ウォームアップをスキップします。');
            return;
        }

        const NE_API_URL = properties.getProperty('NE_API_URL') || 'https://api.next-engine.org';
        const url = `${NE_API_URL}/api_v1_receiveorder_row/search`;

        // 当日の日付を取得（フィルタ条件）
        const today = new Date();
        const year = today.getFullYear();
        const month = String(today.getMonth() + 1).padStart(2, '0');
        const day = String(today.getDate()).padStart(2, '0');
        const formattedDate = `${year}-${month}-${day}`;

        console.log(`ウォームアップ対象日: ${formattedDate}`);

        // 本番（fetchAllShippingData）と同一のフィールド・フィルタ条件
        const payload = {
            'access_token': accessToken,
            'refresh_token': refreshToken,
            'limit': '1',
            // 当日分のみに絞り込む
            'receive_order_send_plan_date-gte': `${formattedDate} 00:00:00`,
            'receive_order_send_plan_date-lte': `${formattedDate} 23:59:59`,
            // キャンセル行を除外（本番と同条件）
            'receive_order_row_cancel_flag-eq': '0',
            // 本番と同一フィールド
            'fields': [
                'receive_order_row_receive_order_id',
                'receive_order_row_goods_id',
                'receive_order_row_goods_name',
                'receive_order_row_quantity',
                'receive_order_row_stock_allocation_quantity',
                'goods_length',
                'goods_width',
                'goods_height',
                'goods_weight',
                'receive_order_date',
                'receive_order_send_plan_date',
                'receive_order_delivery_id',
                'receive_order_delivery_name',
                'receive_order_order_status_id',
                'receive_order_consignee_address1'
            ].join(',')
        };

        const options = {
            'method': 'POST',
            'headers': { 'Content-Type': 'application/x-www-form-urlencoded' },
            'payload': Object.keys(payload).map(key =>
                encodeURIComponent(key) + '=' + encodeURIComponent(payload[key])
            ).join('&'),
            'muteHttpExceptions': true
        };

        const startTime = new Date();
        const response = UrlFetchApp.fetch(url, options);
        const duration = (new Date() - startTime) / 1000;

        console.log(`応答時間: ${duration.toFixed(2)}秒`);

        const responseData = JSON.parse(response.getContentText());

        if (responseData.result === 'success') {
            console.log(`ウォームアップ成功: ${responseData.count}件取得`);

            // トークンが更新された場合は保存
            if (responseData.access_token && responseData.refresh_token) {
                properties.setProperties({
                    'ACCESS_TOKEN': responseData.access_token,
                    'REFRESH_TOKEN': responseData.refresh_token,
                    'TOKEN_UPDATED_AT': new Date().getTime().toString()
                });
                console.log('トークンを更新しました');
            }
        } else {
            console.warn('ウォームアップAPIからエラーが返されました:', responseData.message);
        }

    } catch (error) {
        console.error('ウォームアップ中に予期せぬエラーが発生しました:', error.message);
    }

    console.log('=== ウォームアップ終了 ===');
}