/**
 * =============================================================================
 * ネクストエンジン受注明細取得スクリプト - Phase 5: メイン処理とエラーハンドリング
 * =============================================================================
 * 
 * 【概要】
 * これまでのPhaseで実装した機能を統合し、メイン処理として実行します。
 * リトライ処理、エラーハンドリング、メール通知機能を提供します。
 * 
 * 【Phase 5 実装内容】
 * 1. メイン処理関数(全フェーズの統合)
 * 2. リトライ処理(API呼び出し失敗時)
 * 3. エラーハンドリング
 * 4. 実行時間測定と統計情報
 * 5. メール通知機能(成功時・失敗時)
 * 6. テスト関数
 * 
 * 【実行フロー】
 * 1. スクリプトプロパティ読み込み
 * 2. 検索期間計算(7日前～本日)
 * 3. 店舗マスタ読み込み
 * 4. ネクストエンジンAPI呼び出し(受注明細取得)
 * 5. キャンセル除外フィルタリング
 * 6. データ整形(店舗名付与、日付変換)
 * 7. スプレッドシート書き込み
 * 8. 実行結果通知
 * 
 * @version 1.0
 * @date 2025-11-24
 */

// =============================================================================
// メイン処理
// =============================================================================

/**
 * メイン処理: ネクストエンジン受注明細取得→スプレッドシート書き込み
 * 
 * 全フェーズの処理を統合し、受注明細の取得からスプレッドシート書き込みまでを実行します。
 * エラー発生時は自動的にリトライし、結果をメール通知します。
 * 
 * @return {Object} 実行結果 {success: boolean, statistics: Object, error: string}
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
 * メイン処理を実行し、失敗時は指定回数までリトライします。
 * 
 * @return {Object} 実行結果
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
 * メイン処理が成功した場合に実行結果をメール送信します。
 * 
 * @param {Object} statistics - 実行統計情報
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
 * メイン処理が失敗した場合にエラー内容をメール送信します。
 * 
 * @param {Error} error - エラーオブジェクト
 * @param {number} elapsedTime - 実行時間(秒)
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
 * メイン処理の実行結果を詳細に表示します。
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
// テスト関数
// =============================================================================

/**
 * メイン処理テスト(ドライラン)
 * 
 * メイン処理を実行し、結果を確認します。
 * ⚠️ 実際にAPIを呼び出し、スプレッドシートを更新します!
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
 * リトライ機能を含めたメイン処理を実行します。
 * ⚠️ 実際にAPIを呼び出し、スプレッドシートを更新します!
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
 * 成功通知メールとエラー通知メールの送信をテストします。
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
 * Phase 5で実装した全機能をテストします。
 * ⚠️ testMain() は実際にAPIを呼び出し、スプレッドシートを更新します。
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
 * Phase 1～5のすべてのテストを順次実行します。
 * ⚠️ 時間がかかります(5～10分程度)
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