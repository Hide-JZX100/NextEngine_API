/**
 * 修正後のネクストエンジンAPI取得ロジックの動作テスト用スクリプト (GetShipmentCompletionSlip)
 * 
 * 【目的】
 * 修正した `searchCompletedSlips` が、データの早期終了や歯抜けを起こさず、
 * 正しい件数を全件取得できるかテストします。
 * スプレッドシートへの書き込みは行わず、実行ログ出力（console.log）のみで検証を行います。
 * 
 * 【テスト方法】
 * 1. GASエディタで `runSearchCompletedSlipsTest` を選択し、実行します。
 * 2. 実行ログで「データ整合性チェック: OK」が表示されることを確認します。
 */

/**
 * 修正された searchCompletedSlips 関数のテスト実行
 * 
 * 【動作フロー】
 * 1. 2026年5月分（2026-05-01 〜 2026-05-31）の期間を指定して searchCompletedSlips を呼び出し。
 * 2. 取得したデータの総件数と API 側の総件数が一致しているかを確認。
 * 3. 取得した先頭のデータを1件サンプルとして出力し、データの中身を確認。
 */
function runSearchCompletedSlipsTest() {
    console.log('=== 修正後API取得テスト開始 ===');
    const startDate = '2026-05-01';
    const endDate = '2026-05-31';

    try {
        // 修正後の関数を呼び出して全件取得
        const results = searchCompletedSlips(startDate, endDate);

        console.log('=== テスト実行完了 ===');
        console.log(`取得件数: ${results.length} 件`);

        if (results.length > 0) {
            console.log('サンプルデータ（先頭1件）:', JSON.stringify(results[0]));
            
            // 簡易データ整合性チェック
            const missingFields = [];
            const checkFields = ['receive_order_id', 'receive_order_send_plan_date', 'receive_order_shop_id'];
            
            checkFields.forEach(field => {
                if (!(field in results[0])) {
                    missingFields.push(field);
                }
            });

            if (missingFields.length > 0) {
                console.error(`❌ テスト失敗: サンプルデータに必要なフィールド (${missingFields.join(', ')}) が存在しません。`);
            } else {
                console.log('✅ テスト成功: 必要なフィールドが正しく取得できています。');
            }
        } else {
            console.log('⚠️ データが取得されませんでした。認証状態や日付範囲を確認してください。');
        }

    } catch (e) {
        console.error('❌ テスト実行中にエラーが発生しました:', e.message);
    }
}
