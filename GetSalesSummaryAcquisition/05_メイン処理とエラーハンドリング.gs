/**
 * @file 05_メイン処理とエラーハンドリング.gs
 * @description ネクストエンジン受注明細取得スクリプト - Phase 5: メイン処理とエラーハンドリング。
 * これまでのPhaseで実装した機能を統合し、メイン処理として実行します。
 * リトライ処理、エラーハンドリング、メール通知機能を提供します。
 * 
 * ### 実装内容
 * 1. メイン処理関数 (全フェーズの統合)
 * 2. リトライ処理 (API呼び出し失敗時)
 * 3. エラーハンドリング
 * 4. 実行時間測定と統計情報
 * 5. メール通知機能 (成功時・失敗時)
 * 
 * ### 実行フロー
 * 1. スクリプトプロパティ読み込み -> 検索期間計算
 * 2. 店舗マスタ読み込み -> NE API呼び出し
 * 3. キャンセル除外 -> データ整形
 * 4. スプレッドシート書き込み -> 実行結果通知
 * 
 * ### 依存関係
 * - **01_基盤構築.gs**: getScriptConfig(), getSearchDateRange(), logMessage()
 * - **02_店舗マスタ連携.gs**: getShopMapWithCache()
 * - **03_ネクストエンジンAPI接続.gs**: searchReceiveOrderRowAll()
 * - **04_データ整形・書き込み.gs**: convertAllToSpreadsheetData(), writeToSpreadsheet()
 * - **外部ライブラリ**: NEAuth (v5)
 * 
 * ### 注意事項
 * - **トリガー**: 運用時は `mainWithRetry()` または `integratedDailyUpdate()` を設定してください。
 * - **リトライ**: `RETRY_COUNT` プロパティに従い、失敗のたびに待機時間（5s, 10s...）を増やして再試行します。
 * - **通知**: `NOTIFY_ON_SUCCESS` が `true` の場合、成功時もメール送信されます（1日100通の上限に注意）。
 * - **トリガーチェーン**: `integratedDailyUpdate()` は mainWithRetry() 完了後、sleepせずに dailyUpdate() 起動用のワンタイムトリガーを作成して終了します。両処理は別々の実行枠で動くため6分制限の影響を受けません。
 * 
 * @version 1.1
 * @date 2026-08-13
 * @see mainWithRetry - リトライ付きメイン処理
 * @see testPhase5 - 統合テスト
 * @see integratedDailyUpdate - 統合実行関数 (NE取得 + 売れ数転記トリガー予約)
 */

// =============================================================================
// メイン処理
// =============================================================================

/**
 * メイン処理: ネクストエンジン受注明細取得→スプレッドシート書き込み
 * 
 * @details
 * システムの「メイン・プログラム」に相当する関数です。
 * これまでに構築した Phase 1 から Phase 4 までの各モジュール（GOSUBに相当）を
 * 順番に実行し、データの流れ（パイプライン）を制御します。
 * 
 * 1. 設定のロード（環境変数チェック）
 * 2. 期間計算（実行日から7日間を算出）
 * 3. マスタ準備（店舗名マップの生成）
 * 4. データ抽出（API通信とページネーション）
 * 5. データ変換（オブジェクトから配列への整形）
 * 6. 出力（スプレッドシートへの一括書き込み）
 * 
 * 全体の開始から終了までの時間を計測し、処理の成否と統計情報をオブジェクトとして返却します。
 * 
 * @return {Object} 実行結果オブジェクト {success, statistics, error}
 */
function main() {
  console.log('╔════════════════════════════════════════════════════════════╗');
  console.log('║  ネクストエンジン受注明細取得 - メイン処理開始           ║');
  console.log('╚════════════════════════════════════════════════════════════╝');
  console.log('');

  const overallStartTime = new Date();

  try {
    // ========================================
    // Phase 1: 設定読み込み
    // ========================================
    console.log('【Phase 1】設定読み込み');
    const config = getScriptConfig();
    logMessage(`ログレベル: ${config.logLevel} ${getLogLevelName(config.logLevel)}`);
    console.log('');

    // ========================================
    // Phase 1: 検索期間計算
    // ========================================
    console.log('【Phase 1】検索期間計算');
    const dateRange = getSearchDateRange();
    console.log(`検索期間: ${dateRange.startDateStr} ～ ${dateRange.endDateStr}`);
    console.log('');

    // ========================================
    // Phase 2: 店舗マスタ読み込み
    // ========================================
    console.log('【Phase 2】店舗マスタ読み込み');
    const shopMap = getShopMapWithCache();
    console.log(`店舗マスタ: ${shopMap.size}件`);
    console.log('');

    // ========================================
    // Phase 3: ネクストエンジンAPI呼び出し
    // ========================================
    console.log('【Phase 3】ネクストエンジンAPI呼び出し');
    const orderData = searchReceiveOrderRowAll(
      dateRange.startDateStr,
      dateRange.endDateStr
    );
    console.log(`取得件数: ${orderData.length}件 (API側でキャンセル除外済み)`);
    console.log('');

    // ========================================
    // Phase 4: データ整形
    // ========================================
    console.log('【Phase 4】データ整形');
    const spreadsheetData = convertAllToSpreadsheetData(orderData, shopMap);
    console.log(`整形完了: ${spreadsheetData.length}件`);
    console.log('');

    // ========================================
    // Phase 4: スプレッドシート書き込み
    // ========================================
    console.log('【Phase 4】スプレッドシート書き込み');
    writeToSpreadsheet(spreadsheetData);
    console.log('');

    // ========================================
    // 実行結果サマリー
    // ========================================
    const overallEndTime = new Date();
    const totalElapsedTime = (overallEndTime - overallStartTime) / 1000;

    const statistics = {
      startTime: overallStartTime,
      endTime: overallEndTime,
      elapsedTime: totalElapsedTime,
      searchPeriod: {
        start: dateRange.startDateStr,
        end: dateRange.endDateStr
      },
      dataCount: {
        fetched: orderData.length,        // API取得件数(キャンセル除外済み)
        written: spreadsheetData.length   // スプレッドシート書き込み件数
      },
      shopMasterCount: shopMap.size
    };

    console.log('╔════════════════════════════════════════════════════════════╗');
    console.log('║  ✅ メイン処理完了                                       ║');
    console.log('╚════════════════════════════════════════════════════════════╝');
    console.log('');
    console.log('【実行結果サマリー】');
    console.log(`- 実行時間: ${totalElapsedTime.toFixed(2)}秒`);
    console.log(`- 検索期間: ${dateRange.startDateStr} ～ ${dateRange.endDateStr}`);
    console.log(`- 取得件数: ${orderData.length}件 (API側でキャンセル除外済み)`);
    console.log(`- 書き込み完了: ${spreadsheetData.length}件`);
    console.log('');

    // 成功時のメール通知
    if (config.notifyOnSuccess) {
      sendSuccessNotification(statistics);
    }

    return {
      success: true,
      statistics,
      error: null
    };

  } catch (error) {
    const overallEndTime = new Date();
    const totalElapsedTime = (overallEndTime - overallStartTime) / 1000;

    console.error('');
    console.error('╔════════════════════════════════════════════════════════════╗');
    console.error('║  ❌ メイン処理エラー                                     ║');
    console.error('╚════════════════════════════════════════════════════════════╝');
    console.error('');
    console.error('エラー内容:', error.message);
    console.error('実行時間:', totalElapsedTime.toFixed(2), '秒');
    console.error('');

    // エラー時のメール通知
    sendErrorNotification(error, totalElapsedTime);

    return {
      success: false,
      statistics: null,
      error: error.message
    };
  }
}

// =============================================================================
// リトライ処理
// =============================================================================

/**
 * リトライ付きメイン処理
 * 
 * @details
 * インターネット通信を伴う処理では、一時的なネットワークエラーやAPI側の過負荷により
 * 処理が失敗することがあります。この関数は `main()` をラップし、失敗時に自動で
 * 立て直しを図る「レジリエンス（回復力）」を提供します。
 * 
 * **【設計のポイント：エクスポネンシャル・バックオフ】**
 * 失敗するたびに `Utilities.sleep()` で待機時間を増やします（例: 5秒, 10秒...）。
 * これは、APIサーバーが混雑している場合に、すぐに再試行して負荷をかけ続けるのではなく、
 * 「少し時間を置いてから再度接続する」というマナーに則った設計です。
 * 
 * BASICでの「ON ERROR GOTO」による再試行を、より現代的なループ構造で実現しています。
 * 
 * @return {Object} 最終的な実行結果オブジェクト。全回数失敗した場合は最後の例外を保持します。
 * @see getScriptConfig - RETRY_COUNT設定を使用
 */
function mainWithRetry() {
  console.log('╔════════════════════════════════════════════════════════════╗');
  console.log('║  ネクストエンジン受注明細取得 - リトライ付き実行         ║');
  console.log('╚════════════════════════════════════════════════════════════╝');
  console.log('');

  const config = getScriptConfig();
  const maxRetries = config.retryCount;

  console.log(`最大リトライ回数: ${maxRetries}回`);
  console.log('');

  let lastError = null;

  for (let attempt = 1; attempt <= maxRetries; attempt++) {
    console.log(`--- 実行試行 ${attempt}/${maxRetries} ---`);
    console.log('');

    const result = main();

    if (result.success) {
      console.log('');
      console.log(`✅ ${attempt}回目の試行で成功しました`);
      return result;
    }

    lastError = result.error;

    if (attempt < maxRetries) {
      const waitTime = attempt * 5; // リトライ間隔を徐々に増やす
      console.log('');
      console.log(`⚠️ ${attempt}回目の試行が失敗しました`);
      console.log(`${waitTime}秒後に再試行します...`);
      console.log('');

      Utilities.sleep(waitTime * 1000);
    }
  }

  console.error('');
  console.error(`❌ ${maxRetries}回のリトライすべてが失敗しました`);
  console.error('最終エラー:', lastError);

  return {
    success: false,
    statistics: null,
    error: `${maxRetries}回のリトライすべてが失敗: ${lastError}`
  };
}

// =============================================================================
// メール通知
// =============================================================================

/**
 * 成功通知メールを送信
 * 
 * @details
 * 処理が完了したことを管理者に報告します。
 * 単に「終わった」だけでなく、`statistics` オブジェクトから「何件処理したか」「何秒かかったか」を
 * 抽出し、人間が読みやすいレポート形式に整形して送信します。
 * 
 * 毎日実行されるプログラムにおいて、スプレッドシートを開かずに
 * 「正しく動いているか」を把握するための生存確認（ヘルスチェック）として機能します。
 * 
 * @param {Object} statistics - 実行統計情報
 * @see MailApp
 */
function sendSuccessNotification(statistics) {
  try {
    const config = getScriptConfig();
    const recipient = config.notificationEmail;

    const subject = '✅ [成功] ネクストエンジン受注明細取得完了';

    const body = `
ネクストエンジン受注明細取得が正常に完了しました。

【実行結果】
- 実行時間: ${statistics.elapsedTime.toFixed(2)}秒
- 開始時刻: ${statistics.startTime.toLocaleString('ja-JP')}
- 終了時刻: ${statistics.endTime.toLocaleString('ja-JP')}

【検索条件】
- 検索期間: ${statistics.searchPeriod.start} ～ ${statistics.searchPeriod.end}

【取得データ】
- API取得件数: ${statistics.dataCount.fetched}件 (キャンセル除外済み)
- スプレッドシート書き込み: ${statistics.dataCount.written}件

【店舗マスタ】
- 読み込み件数: ${statistics.shopMasterCount}件

---
このメールは自動送信されています。
スクリプトID: ${ScriptApp.getScriptId()}
    `;

    MailApp.sendEmail(recipient, subject, body);

    logMessage(`成功通知メールを送信しました: ${recipient}`);

  } catch (error) {
    console.error('⚠️ 成功通知メール送信エラー:', error.message);
    // メール送信失敗は致命的エラーではないのでログのみ
  }
}

/**
 * エラー通知メールを送信
 * 
 * @details
 * プログラムが異常終了した際に、迅速に復旧作業へ移るためのアラートメールを送信します。
 * 
 * メールの本文には、発生したエラーメッセージに加えて「GASエディタへの直接リンク」や
 * 「チェックリスト」を記載し、受け取った人が即座にデバッグを開始できるように配慮しています。
 * これはシステムのダウンタイムを最小限にするための重要な設計です。
 * 
 * @param {Error} error - エラーオブジェクト
 * @param {number} elapsedTime - 実行時間(秒)
 * @see ScriptApp.getScriptId
 */
function sendErrorNotification(error, elapsedTime) {
  try {
    const config = getScriptConfig();
    const recipient = config.notificationEmail;

    const subject = '❌ [エラー] ネクストエンジン受注明細取得失敗';

    const body = `
ネクストエンジン受注明細取得でエラーが発生しました。

【エラー内容】
${error.message}

【実行情報】
- 実行時間: ${elapsedTime.toFixed(2)}秒
- 発生時刻: ${new Date().toLocaleString('ja-JP')}

【対処方法】
1. GASエディタを開く
   URL: https://script.google.com/home/projects/${ScriptApp.getScriptId()}/edit

2. 実行ログを確認
   左メニュー「実行数」から詳細なログを確認できます

3. 以下の項目を確認してください
   - スクリプトプロパティが正しく設定されているか
   - ACCESS_TOKEN / REFRESH_TOKEN が有効か
   - スプレッドシートへのアクセス権限があるか
   - ネクストエンジンAPIが正常に動作しているか

4. 問題が解決しない場合
   - 手動で main() または testPhase1() ～ testPhase5() を実行してエラー箇所を特定

---
このメールは自動送信されています。
スクリプトID: ${ScriptApp.getScriptId()}
    `;

    MailApp.sendEmail(recipient, subject, body);

    console.log(`エラー通知メールを送信しました: ${recipient}`);

  } catch (mailError) {
    console.error('⚠️ エラー通知メール送信エラー:', mailError.message);
    // メール送信失敗は致命的エラーではないのでログのみ
  }
}

// =============================================================================
// 実行統計情報
// =============================================================================

/**
 * 実行統計情報を表示
 * 
 * @details
 * `main()` の実行結果を受け取り、開発者がデバッグや性能評価を行うための
 * 詳細データをコンソールに出力します。
 * 
 * 特に「処理速度（件数/秒）」を算出しており、データ量が増加した際の
 * ボトルネック調査に役立ちます。
 * 
 * @param {Object} result - main()の戻り値
 */
function displayStatistics(result) {
  console.log('');
  console.log('╔════════════════════════════════════════════════════════════╗');
  console.log('║  実行統計情報                                             ║');
  console.log('╚════════════════════════════════════════════════════════════╝');
  console.log('');

  if (!result.success) {
    console.error('❌ 実行失敗');
    console.error('エラー:', result.error);
    return;
  }

  const stats = result.statistics;

  console.log('【実行時間】');
  console.log(`  開始: ${stats.startTime.toLocaleString('ja-JP')}`);
  console.log(`  終了: ${stats.endTime.toLocaleString('ja-JP')}`);
  console.log(`  所要時間: ${stats.elapsedTime.toFixed(2)}秒`);
  console.log('');

  console.log('【検索条件】');
  console.log(`  期間: ${stats.searchPeriod.start} ～ ${stats.searchPeriod.end}`);
  console.log('');

  console.log('【データ処理】');
  console.log(`  API取得: ${stats.dataCount.fetched}件 (キャンセル除外済み)`);
  console.log(`  書き込み完了: ${stats.dataCount.written}件`);
  console.log('');

  console.log('【店舗マスタ】');
  console.log(`  店舗数: ${stats.shopMasterCount}件`);
  console.log('');

  // パフォーマンス指標
  const recordsPerSecond = (stats.dataCount.written / stats.elapsedTime).toFixed(2);
  console.log('【パフォーマンス】');
  console.log(`  処理速度: ${recordsPerSecond}件/秒`);
  console.log('');
}

// =============================================================================
// 統合実行関数(ネクストエンジン取得 + 売れ数転記)
// =============================================================================

/**
 * 統合メイン処理: ネクストエンジン受注明細取得 → 売れ数転記トリガー予約（トリガーチェーン）
 * 
 * @details
 * 本システムと、別のシステム（売れ数転記スクリプト）を非同期に連携させるマスタ・プログラムです。
 * 
 * 【トリガーチェーン方式】
 * `mainWithRetry()` を実行し、成功した場合にスクリプトプロパティ `DAILY_UPDATE_DELAY_SECONDS`
 * に設定された遅延秒数後に `dailyUpdate()` 起動用のワンタイムトリガーを動的作成します。
 * これにより、`Utilities.sleep()` による長時間の待機を排除し、GASの6分実行制限を回避します。
 * 
 * 運用時の時限トリガー（8:00, 12:00, 17:00など）には、この関数を登録します。
 * 
 * @return {Object} NE取得結果および dailyUpdate 予約状況を含む統合レポート
 */
function integratedDailyUpdate() {
  console.log('╔════════════════════════════════════════════════════════════╗');
  console.log('║  統合処理開始: NE取得 → 売れ数転記(トリガーチェーン)       ║');
  console.log('╚════════════════════════════════════════════════════════════╝');
  console.log('');

  const overallStartTime = new Date();
  const results = {
    neExecution: null,
    dailyUpdate: null,
    startTime: overallStartTime,
    endTime: null,
    totalElapsedTime: 0
  };

  try {
    // ========================================
    // Phase 1: ネクストエンジン受注明細取得
    // ========================================
    console.log('【Phase 1】ネクストエンジン受注明細取得');
    console.log('');

    results.neExecution = mainWithRetry();

    console.log('');
    console.log('✅ ネクストエンジン受注明細取得完了');
    console.log('');

    // ========================================
    // Phase 2: 売れ数転記トリガー予約
    // ========================================
    if (results.neExecution && results.neExecution.success) {
      console.log('【Phase 2】売れ数転記トリガー予約');
      console.log('');

      const delaySeconds = getDailyUpdateDelaySeconds();
      scheduleDailyUpdateTrigger(delaySeconds);

      results.dailyUpdate = 'scheduled';
      console.log('');
      console.log(`✅ 売れ数転記トリガー予約完了 (${delaySeconds}秒後に起動)`);
      console.log('');
    } else {
      console.warn('⚠️ ネクストエンジン受注明細取得が失敗したため、売れ数転記トリガー予約をスキップします。');
      results.dailyUpdate = 'skipped';
    }

    // ========================================
    // 統合結果サマリー
    // ========================================
    results.endTime = new Date();
    results.totalElapsedTime = (results.endTime - overallStartTime) / 1000;

    displayIntegratedSummary(results);

    return results;

  } catch (error) {
    console.error('');
    console.error('╔════════════════════════════════════════════════════════════╗');
    console.error('║  ❌ 統合処理エラー                                       ║');
    console.error('╚════════════════════════════════════════════════════════════╝');
    console.error('');
    console.error('エラー内容:', error.message);

    // エラー通知
    const subject = '❌ [エラー] 統合処理(NE取得+売れ数転記予約)失敗';
    const body = `
統合処理でエラーが発生しました。

【エラー内容】
${error.message}

【実行状況】
- ネクストエンジン取得: ${results.neExecution && results.neExecution.success ? '成功' : '失敗またはエラー'}
- 売れ数転記予約: ${results.dailyUpdate || '未実行'}

【発生時刻】
${new Date().toLocaleString('ja-JP')}

【対処方法】
GASエディタで実行ログを確認してください。
URL: https://script.google.com/home/projects/${ScriptApp.getScriptId()}/edit
    `;

    MailApp.sendEmail(Session.getActiveUser().getEmail(), subject, body);

    throw error;
  }
}

/**
 * スクリプトプロパティから dailyUpdate 起動までの遅延秒数を取得
 *
 * @details
 * ハードコーディング・フォールバックデフォルトは禁止のため、
 * スクリプトプロパティ 'DAILY_UPDATE_DELAY_SECONDS' が未設定、
 * または不正な値の場合は例外を送出する。
 *
 * @return {number} 遅延秒数
 * @throws {Error} プロパティ未設定または不正値の場合
 */
function getDailyUpdateDelaySeconds() {
  const propertyName = 'DAILY_UPDATE_DELAY_SECONDS';
  const val = PropertiesService.getScriptProperties().getProperty(propertyName);

  if (val === null || val === undefined || val.trim() === '') {
    throw new Error(`スクリプトプロパティ '${propertyName}' が設定されていません。GASエディタの「プロジェクトの設定」から登録してください。`);
  }

  const delaySeconds = Number(val);
  if (isNaN(delaySeconds) || delaySeconds < 0) {
    throw new Error(`スクリプトプロパティ '${propertyName}' の値が不正です (入力値: '${val}')。正の数値を指定してください。`);
  }

  return delaySeconds;
}

/**
 * dailyUpdate() 起動用のワンタイムトリガーを作成
 *
 * @details
 * 重複起動を防ぐため、既存の dailyUpdate 向けトリガーを
 * 07_トリガー作成スクリプト.gs の deleteTriggersForFunction() で
 * 先に削除してから、指定秒数後に発火するワンタイムトリガーを作成する。
 *
 * @param {number} delaySeconds - 発火までの遅延秒数
 */
function scheduleDailyUpdateTrigger(delaySeconds) {
  const targetFunctionName = 'dailyUpdate';

  // 既存トリガー削除（07_トリガー作成スクリプト.gs に定義済みの deleteTriggersForFunction を利用）
  deleteTriggersForFunction(targetFunctionName);

  const delayMs = delaySeconds * 1000;
  ScriptApp.newTrigger(targetFunctionName)
    .timeBased()
    .after(delayMs)
    .create();

  console.log(`トリガー作成完了: 関数 '${targetFunctionName}' を ${delaySeconds}秒後に実行します。`);
}

/**
 * 統合処理結果サマリーを表示
 * 
 * @details
 * `integratedDailyUpdate` 内の2つの大きなタスクが、それぞれどのような状態（成功/失敗/予約）
 * で終わったかを一目で確認するためのサマリーをログ出力します。
 * 
 * @param {Object} results - 統合実行の結果オブジェクト
 */
function displayIntegratedSummary(results) {
  console.log('╔════════════════════════════════════════════════════════════╗');
  console.log('║  統合処理結果サマリー                                     ║');
  console.log('╚════════════════════════════════════════════════════════════╝');
  console.log('');

  console.log('【実行時間】');
  console.log(`開始: ${results.startTime.toLocaleString('ja-JP')}`);
  console.log(`終了: ${results.endTime.toLocaleString('ja-JP')}`);
  console.log(`総所要時間: ${results.totalElapsedTime.toFixed(2)}秒`);
  console.log('');

  console.log('【ネクストエンジン受注明細取得】');
  if (results.neExecution && results.neExecution.success) {
    const stats = results.neExecution.statistics;
    console.log(`ステータス: 成功`);
    console.log(`取得件数: ${stats.dataCount.fetched}件`);
    console.log(`書き込み完了: ${stats.dataCount.written}件`);
    console.log(`実行時間: ${stats.elapsedTime.toFixed(2)}秒`);
  } else {
    console.log('ステータス: 失敗またはエラー');
  }
  console.log('');

  console.log('【売れ数転記処理】');
  if (results.dailyUpdate === 'scheduled') {
    console.log('ステータス: 起動予約済み（非同期実行）');
  } else if (results.dailyUpdate === 'skipped') {
    console.log('ステータス: スキップ（NE取得失敗のため）');
  } else if (results.dailyUpdate) {
    console.log('ステータス: 完了(詳細は上記参照)');
  } else {
    console.log('ステータス: 未実行またはエラー');
  }
  console.log('');

  console.log('✅ 統合処理がすべて完了しました');
}

/**
 * メイン処理テスト(ドライラン)
 * 
 * @details
 * `main()` 関数を単独で実行します。
 * 実際のAPI呼び出しと書き込みが発生するため、開発中の最終確認用として使用します。
 * @return {Object} 実行結果
 */
function testMain() {
  console.log('=== メイン処理テスト ===');
  console.log('⚠️ 実際にAPIを呼び出し、スプレッドシートを更新します!');
  console.log('');

  const result = main();

  console.log('');
  displayStatistics(result);

  return result;
}

/**
 * リトライ処理テスト
 * 
 * @details
 * リトライロジックが組み込まれた `mainWithRetry()` をテストします。
 * 万が一の通信エラー時にも自動復旧することを確認するための入り口です。
 * @return {Object} 実行結果
 */
function testMainWithRetry() {
  console.log('=== リトライ付きメイン処理テスト ===');
  console.log('⚠️ 実際にAPIを呼び出し、スプレッドシートを更新します!');
  console.log('');

  const result = mainWithRetry();

  console.log('');
  displayStatistics(result);

  return result;
}

/**
 * メール通知テスト
 * 
 * @details
 * 実際にメールが送信されるか、HTMLタグや変数が正しく展開されているかを確認します。
 * 本番運用前に、自分のメールアドレスに通知が届くことを確認するための関数です。
 * 誤字脱字やリンク切れのチェックに使用します。
 */
function testEmailNotification() {
  console.log('=== メール通知テスト ===');
  console.log('');

  const config = getScriptConfig();
  console.log(`通知先: ${config.notificationEmail}`);
  console.log('');

  // テスト用統計情報
  const testStatistics = {
    startTime: new Date(Date.now() - 30000),
    endTime: new Date(),
    elapsedTime: 30.5,
    searchPeriod: {
      start: '2025-11-17 00:00:00',
      end: '2025-11-24 23:59:59'
    },
    dataCount: {
      raw: 6500,
      cancelled: 150,
      valid: 6350,
      written: 6350
    },
    shopMasterCount: 25
  };

  try {
    console.log('【1】成功通知メール送信テスト');
    sendSuccessNotification(testStatistics);
    console.log('✅ 成功通知メール送信完了');
    console.log('');

    console.log('【2】エラー通知メール送信テスト');
    const testError = new Error('これはテスト用のエラーメッセージです');
    sendErrorNotification(testError, 15.3);
    console.log('✅ エラー通知メール送信完了');
    console.log('');

    console.log('✅ メール通知テスト完了!');
    console.log('');
    console.log('【確認事項】');
    console.log('- メールが届いているか確認してください');
    console.log('- 件名と本文が正しいか確認してください');

  } catch (error) {
    console.error('❌ メール通知テストエラー:', error.message);
    throw error;
  }
}

// =============================================================================
// Phase 5 統合テスト
// =============================================================================

/**
 * Phase 5 統合テスト
 * 
 * @details
 * メイン処理と通知系の統合テストです。
 * これが完了すれば、全ての単体パーツが組み合わさって一つのシステムとして完成したことになります。
 * 
 * @throws {Error} 通知設定等に不備がある場合
 */
function testPhase5() {
  console.log('╔════════════════════════════════════════════════════════════╗');
  console.log('║  Phase 5: メイン処理とエラーハンドリング - 統合テスト    ║');
  console.log('╚════════════════════════════════════════════════════════════╝');
  console.log('');

  try {
    // 1. メール通知テスト
    console.log('【1】メール通知テスト');
    testEmailNotification();
    console.log('');

    // 2. メイン処理テスト(実行確認)
    console.log('【2】メイン処理テスト');
    console.log('⚠️ 実際にAPIを呼び出し、スプレッドシートを更新します!');
    console.log('実行しますか? (手動でtestMain()を実行してください)');
    console.log('');

    console.log('╔════════════════════════════════════════════════════════════╗');
    console.log('║  ✅ Phase 5 統合テスト: 完了                             ║');
    console.log('╚════════════════════════════════════════════════════════════╝');
    console.log('');
    console.log('【次のステップ】');
    console.log('1. testMain() を実行してメイン処理テスト');
    console.log('2. 本番運用の準備');
    console.log('   - トリガー設定(毎日自動実行)');
    console.log('   - ログレベル調整(LOG_LEVEL=3推奨)');
    console.log('   - 通知設定確認');

  } catch (error) {
    console.error('');
    console.error('╔════════════════════════════════════════════════════════════╗');
    console.error('║  ❌ Phase 5 統合テスト: エラー発生                       ║');
    console.error('╚════════════════════════════════════════════════════════════╝');
    console.error('');
    console.error('エラー内容:', error.message);

    throw error;
  }
}

// =============================================================================
// 全体統合テスト
// =============================================================================

/**
 * 全Phase統合テスト
 * 
 * @details
 * システム開発の最終工程です。Phase 1 から Phase 5 までの全てのテスト関数を
 * 順番にオートメーションで実行します。
 * 全てにチェックが入れば、納品・運用開始が可能な状態です。
 * 
 * @note 実際の通信が発生するため、実行には数分間の時間を要します。
 */
function testAllPhases() {
  console.log('╔════════════════════════════════════════════════════════════╗');
  console.log('║  全Phase統合テスト                                        ║');
  console.log('╚════════════════════════════════════════════════════════════╝');
  console.log('');
  console.log('⚠️ すべてのPhaseのテストを順次実行します');
  console.log('⚠️ 実際のAPIを呼び出し、スプレッドシートを更新します');
  console.log('');

  const overallStartTime = new Date();

  try {
    // Phase 1
    console.log('【Phase 1】基盤構築');
    testPhase1();
    console.log('');
    console.log('✅ Phase 1 完了');
    console.log('');
    console.log('─────────────────────────────────────────────────────────');
    console.log('');

    // Phase 2
    console.log('【Phase 2】店舗マスタ連携');
    testPhase2();
    console.log('');
    console.log('✅ Phase 2 完了');
    console.log('');
    console.log('─────────────────────────────────────────────────────────');
    console.log('');

    // Phase 3
    console.log('【Phase 3】ネクストエンジンAPI接続');
    testPhase3();
    console.log('');
    console.log('✅ Phase 3 完了');
    console.log('');
    console.log('─────────────────────────────────────────────────────────');
    console.log('');

    // Phase 4
    console.log('【Phase 4】データ整形・書き込み');
    testPhase4();
    console.log('');
    console.log('✅ Phase 4 完了');
    console.log('');
    console.log('─────────────────────────────────────────────────────────');
    console.log('');

    // Phase 5
    console.log('【Phase 5】メイン処理とエラーハンドリング');
    testPhase5();
    console.log('');
    console.log('✅ Phase 5 完了');
    console.log('');

    const overallEndTime = new Date();
    const totalTime = (overallEndTime - overallStartTime) / 1000;

    console.log('╔════════════════════════════════════════════════════════════╗');
    console.log('║  ✅ 全Phase統合テスト: すべて成功!                       ║');
    console.log('╚════════════════════════════════════════════════════════════╝');
    console.log('');
    console.log(`総実行時間: ${totalTime.toFixed(2)}秒`);
    console.log('');
    console.log('【開発完了!】');
    console.log('すべてのフェーズのテストが完了しました。');
    console.log('本番運用の準備に進んでください。');

  } catch (error) {
    console.error('');
    console.error('╔════════════════════════════════════════════════════════════╗');
    console.error('║  ❌ 全Phase統合テスト: エラー発生                        ║');
    console.error('╚════════════════════════════════════════════════════════════╝');
    console.error('');
    console.error('エラー内容:', error.message);

    throw error;
  }
}

/**
 * トリガーチェーンヘルパー関数の単体テスト
 * 
 * @details
 * - getDailyUpdateDelaySeconds() の未設定時・設定時の動作検証
 * - scheduleDailyUpdateTrigger() によるワンタイムトリガー作成検証
 */
function testTriggerChainHelpers() {
  console.log('=== トリガーチェーンヘルパー関数 テスト開始 ===');
  const propKey = 'DAILY_UPDATE_DELAY_SECONDS';
  const props = PropertiesService.getScriptProperties();
  const originalValue = props.getProperty(propKey);

  try {
    // 1. 未設定時エラーテスト
    console.log('\n--- 1. 未設定時エラーテスト ---');
    props.deleteProperty(propKey);
    try {
      getDailyUpdateDelaySeconds();
      console.error('❌ FAIL: 未設定時エラーがスローされませんでした。');
    } catch (e) {
      console.log('✅ PASS: 未設定時エラーキャッチ:', e.message);
    }

    // 2. 不正値エラーテスト
    console.log('\n--- 2. 不正値エラーテスト ---');
    props.setProperty(propKey, 'invalid_number');
    try {
      getDailyUpdateDelaySeconds();
      console.error('❌ FAIL: 不正値エラーがスローされませんでした。');
    } catch (e) {
      console.log('✅ PASS: 不正値エラーキャッチ:', e.message);
    }

    // 3. 正常取得テスト
    console.log('\n--- 3. 正常取得テスト ---');
    props.setProperty(propKey, '30');
    const delay = getDailyUpdateDelaySeconds();
    if (delay === 30) {
      console.log('✅ PASS: 遅延秒数取得成功:', delay);
    } else {
      console.error(`❌ FAIL: 取得値が一致しません。期待値: 30, 実際: ${delay}`);
    }

    // 4. ワンタイムトリガー作成テスト
    console.log('\n--- 4. ワンタイムトリガー作成テスト ---');
    scheduleDailyUpdateTrigger(30);
    const triggers = ScriptApp.getProjectTriggers();
    const dailyUpdateTriggers = triggers.filter(t => t.getHandlerFunction() === 'dailyUpdate');
    if (dailyUpdateTriggers.length === 1) {
      console.log('✅ PASS: dailyUpdate のワンタイムトリガーが1件正常に作成されました。');
    } else {
      console.error(`❌ FAIL: トリガー件数が不正です。期待値: 1, 実際: ${dailyUpdateTriggers.length}`);
    }

    // テスト作成したトリガーを後掃除
    deleteTriggersForFunction('dailyUpdate');
    console.log('🧹 テスト用トリガーを削除しました。');

  } finally {
    // リストア
    if (originalValue !== null) {
      props.setProperty(propKey, originalValue);
    } else {
      props.deleteProperty(propKey);
    }
    console.log('\n=== テスト完了 ===');
  }
}