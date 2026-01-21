/**
 * メイン処理
 * プロジェクトのエントリーポイント
 */

/**
 * メイン関数
 * 定期実行（タイムトリガー）の対象となる関数
 */
function main() {
    const startTime = new Date();
    console.log(`=== 出荷済み伝票取得処理開始 ${startTime.toLocaleString('ja-JP')} ===`);

    try {
        // 1. 対象日の取得
        const targetDate = getTargetDate();
        console.log(`対象日: ${targetDate}`);

        // 2. APIからデータ取得
        // 指定された対象日の「出荷予定日」を持つ伝票を全件取得
        const rawData = searchCompletedSlips(targetDate);

        if (rawData.length === 0) {
            console.log('APIからの取得データが0件でした。処理を終了します。');
            return;
        }

        // 3. データ処理 (フィルタリング & マッピング)
        // フィルタ条件: 受注状態区分=50 OR 受注キャンセル区分=3
        const processedData = filterAndMapSlips(rawData);

        if (processedData.length === 0) {
            console.log('フィルタリングの結果、対象データがありませんでした。処理を終了します。');
            return;
        }

        console.log(`書込対象データ: ${processedData.length} 件`);

        // 4. スプレッドシートへ出力
        writeToSpreadsheet(processedData);

        console.log('=== 全処理が正常に完了しました ===');

    } catch (e) {
        console.error(`❌ エラーが発生しました: ${e.message}`);
        // 必要に応じて管理者へのメール通知などを実装
    } finally {
        const endTime = new Date();
        const duration = (endTime - startTime) / 1000;
        console.log(`処理時間: ${duration}秒`);
    }
}
