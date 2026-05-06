/**
 * メイン処理ファイル (GetShipmentCompletionSlip)
 * 
 * 【目的】
 * 本プロジェクトのエントリーポイントとなるファイルです。
 * 分割された各モジュール（検索・加工・出力）を統合し、正しい順序で実行制御を行います。
 * 大容量データの取得に対応するため、ウォームアップおよび3段階のバッチ分割実行をサポートします。
 * 通常、定期実行トリガーは `warmupAndScheduleMain()` に対して設定します。
 * 
 * 【処理フロー】
 * 1. [ウォームアップ] APIへの軽量リクエストによる接続確認とトークン更新。
 * 2. [トリガー予約] 指定時間後にメインバッチ（executeBatch1）を予約。
 * 3. [メイン処理] 検索対象日付の決定、APIからのデータ取得。
 * 4. [データ加工] フィルタリング（受注状態区分:50 / 受注キャンセル区分:3）とマッピング。
 * 5. [出力] Googleスプレッドシートへの追記。
 * 6. [連鎖実行] 次のバッチが存在する場合、5分後の実行を自動的にスケジューリング。
 * 
 * 【含まれる関数】
 * - main(batchNumber): 全体の処理フローを制御する中核関数
 * - warmupAndScheduleMain(batchNumber): コールドスタート回避と実行予約を行うエントリーポイント
 * - scheduleNextBatch(nextBatchNumber): 次のバッチ実行トリガーを作成
 * - executeBatch1/2/3(): トリガーから直接呼び出されるラッパー関数
 * 
 * 【依存関係】
 * - utils.ユーティリティ.gs: 日付取得
 * - api.ネクストエンジンAPI.gs: データ検索
 * - process.フィルタとマッピング.gs: データ加工
 * - sheet.スプレッドシート書込.gs: シート出力
 */

/**
 * メイン関数
 *
 * @param {number} batchNumber - バッチ番号（1, 2, 3）省略時は1
 */
function main(batchNumber) {
    // トリガーから実行された場合、第1引数はイベントオブジェクトになる
    // 数値でない場合はデフォルト値1を使用
    if (typeof batchNumber !== 'number' || batchNumber < 1 || batchNumber > 3) {
        batchNumber = 1;
    }

    const startTime = new Date();
    console.log(`=== 出荷済み伝票取得処理開始 (バッチ${batchNumber}) ${startTime.toLocaleString('ja-JP')} ===`);

    try {
        // 1. 対象日付範囲の取得
        const dateRange = getTargetDateRange(batchNumber);
        console.log(`対象範囲: ${dateRange.startDate} ～ ${dateRange.endDate}`);

        // 2. APIからデータ取得
        // 指定された日付範囲の「出荷予定日」を持つ伝票を全件取得
        const rawData = searchCompletedSlips(dateRange.startDate, dateRange.endDate);

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

        // 次のバッチをスケジュール
        if (batchNumber < 3) {
            scheduleNextBatch(batchNumber + 1);
        }

    } catch (e) {
        console.error(`❌ エラーが発生しました: ${e.message}`);
        // 必要に応じて管理者へのメール通知などを実装
    } finally {
        const endTime = new Date();
        const duration = (endTime - startTime) / 1000;
        console.log(`処理時間: ${duration}秒`);
    }
}

/**
 * ウォームアップ関数
 * コールドスタートを回避し、メイン処理の実行速度を向上させる
 * 
 * 【実行フロー】
 * 1. この関数をトリガーで実行
 * 2. APIに軽量リクエストを送信してサーバー側キャッシュを準備
 * 3. 2分後にmain関数を実行するトリガーを登録
 * 
 * @param {number} batchNumber - バッチ番号（1, 2, 3）省略時は1
 */
function warmupAndScheduleMain(batchNumber) {
    // トリガーから実行された場合、第1引数はイベントオブジェクトになる
    // 数値でない場合はデフォルト値1を使用
    if (typeof batchNumber !== 'number' || batchNumber < 1 || batchNumber > 3) {
        batchNumber = 1;
    }

    console.log(`=== ウォームアップ開始 (バッチ${batchNumber}用) ===`);

    try {
        // バッチ1の場合、事前クリーンアップを実行
        if (batchNumber === 1) {
            cleanupOldTriggersAndData();
        }

        // 1. 対象日付範囲の取得
        const dateRange = getTargetDateRange(batchNumber);

        // 2. スクリプトプロパティからトークン取得
        const props = PropertiesService.getScriptProperties();
        const accessToken = props.getProperty('ACCESS_TOKEN');
        const refreshToken = props.getProperty('REFRESH_TOKEN');

        if (!accessToken || !refreshToken) {
            throw new Error('トークンが見つかりません。認証を行ってください。');
        }

        // 3. メイン処理と同じエンドポイントに最小限のリクエスト
        const params = {
            'access_token': accessToken,
            'refresh_token': refreshToken,
            'wait_flag': '1',
            'fields': 'receive_order_id',  // 最小限のフィールド
            'receive_order_send_plan_date-gte': dateRange.startDate,
            'limit': '1',  // 1件のみ
            'offset': '0'
        };

        const url = CONFIG.API.BASE_URL + CONFIG.API.ENDPOINT_SEARCH;
        const options = {
            'method': 'post',
            'payload': params,
            'muteHttpExceptions': true
        };

        console.log('ウォームアップリクエスト送信中...');
        const response = UrlFetchApp.fetch(url, options);
        const json = JSON.parse(response.getContentText());

        if (json.result !== 'success') {
            throw new Error(`ウォームアップ失敗: ${json.message}`);
        }

        // 4. トークン更新
        if (json.access_token && json.refresh_token) {
            props.setProperties({
                'ACCESS_TOKEN': json.access_token,
                'REFRESH_TOKEN': json.refresh_token,
                'TOKEN_UPDATED_AT': new Date().getTime().toString()
            });
            console.log('トークンを更新しました');
        }

        console.log('✅ ウォームアップ完了');

        // 5. 2分後にmain関数を実行するトリガーを作成
        const triggerTime = new Date(new Date().getTime() + 2 * 60 * 1000); // 2分後
        ScriptApp.newTrigger('executeBatch' + batchNumber)
            .timeBased()
            .at(triggerTime)
            .create();

        console.log(`${triggerTime.toLocaleString('ja-JP')} にバッチ${batchNumber}を実行するトリガーを作成しました`);

    } catch (error) {
        console.error('❌ ウォームアップエラー:', error.message);
        throw error;
    }
}

/**
 * 次のバッチを5分後に実行するトリガーを作成
 * 
 * @param {number} nextBatchNumber - 次のバッチ番号（2 or 3）
 */
function scheduleNextBatch(nextBatchNumber) {
    if (nextBatchNumber < 2 || nextBatchNumber > 3) {
        console.log('次のバッチはありません。処理完了');
        return;
    }

    const intervalMinutes = CONFIG.BATCH.INTERVAL_MINUTES;
    const triggerTime = new Date(new Date().getTime() + intervalMinutes * 60 * 1000);

    ScriptApp.newTrigger(`executeBatch${nextBatchNumber}`)
        .timeBased()
        .at(triggerTime)
        .create();

    console.log(`${triggerTime.toLocaleString('ja-JP')} にバッチ${nextBatchNumber}を実行するトリガーを作成しました`);
}

/**
 * バッチ1実行関数（トリガーから呼ばれる）
 */
function executeBatch1() {
    deleteMyTrigger('executeBatch1'); // 自分のトリガーを削除
    main(1);
}

/**
 * バッチ2実行関数（トリガーから呼ばれる）
 */
function executeBatch2() {
    deleteMyTrigger('executeBatch2'); // 自分のトリガーを削除
    main(2);
}

/**
 * バッチ3実行関数（トリガーから呼ばれる）
 */
function executeBatch3() {
    deleteMyTrigger('executeBatch3'); // 自分のトリガーを削除
    main(3);
}