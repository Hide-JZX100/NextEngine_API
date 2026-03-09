/**
 * メイン処理ファイル (GetShipmentCompletionSlip)
 * 
 * 【目的】
 * 本プロジェクトのエントリーポイントとなるファイルです。
 * 分割された各モジュール（検索・加工・出力）を統合し、正しい順序で実行制御を行います。
 * 定期実行トリガー（毎日実行など）はこのファイルの `main()` 関数に対して設定します。
 * 
 * 【処理フロー】
 * 1. 検索対象日付の決定（開発モード/本番モードの判定）
 * 2. ネクストエンジンAPIからの出荷完了伝票検索・取得
 * 3. 取得データのフィルタリング（ステータス50/キャンセル3）およびマッピング
 * 4. Googleスプレッドシートへのデータ追記
 * 
 * 【含まれる関数】
 * - main(): 全体の処理フローを制御するメイン関数（トリガー実行対象）
 * 
 * 【依存関係】
 * - utils.gs: 日付取得
 * - api.gs: データ検索
 * - process.gs: データ加工
 * - sheet.gs: シート出力
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
