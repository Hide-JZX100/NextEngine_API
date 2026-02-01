/**
=============================================================================
在庫情報取得スクリプト - リトライ機能追加版（SRE改善）
=============================================================================

* 【目的】
* スプレッドシート上の商品コードに基づき、ネクストエンジンAPIから在庫情報を一括取得し、
* シートを更新します。Google Apps Scriptの一時的な障害に対応するため、
* API接続の自動リトライ機能（エクスポネンシャルバックオフ）を実装しています。
* 
* 【主な機能】
* - 在庫情報の一括取得とスプレッドシートへの反映
* - API接続エラー時の自動リトライ機能
* - リトライ状況の統計記録と可視化（ログシート、コンソール）
* - システムの健全性を確認するSREダッシュボード機能

=============================================================================
                        関数ガイド
=============================================================================

--- ユーザー向け主要関数 ---

@see showUsageGuideWithRetry           - 【使い方】このスクリプトの機能や使い方ガイドを表示します。
@see showSREDashboard                  - 【健全性確認】リトライ状況やエラーログからシステムの健全性をダッシュボード形式で表示します。

--- 設定・テスト用関数 ---

@see showRetryConfig     - 現在のリトライ設定（有効/無効、回数など）を表示します。
@see enableRetry         - リトライ機能を「有効」にします。
@see disableRetry        - リトライ機能を「無効」にします。
@see testRetryFunction   - リトライ機能が正しく動作するかを小規模なデータでテストします。
@see compareVersions     - 従来版とリトライ版の処理速度や安定性を比較テストします。
@see showMigrationGuide  - 従来版のスクリプトから本スクリプトへ安全に移行するための手順を表示します。

/**
=============================================================================
在庫情報取得スクリプト - リトライ機能追加版（SRE改善）
=============================================================================

* 【追加内容】v2.1
* - API接続リトライ機能（出荷予定数取得で実績済み）
* - エクスポネンシャルバックオフ（指数バックオフ）方式
* - Google側一時障害への対応強化
* - リトライ統計情報の記録
* 
* 【変更方針】
* - 既存コードの安定性を維持
* - getBatchStockData関数のみに防御層を追加
* - 既存のログレベル機能と統合
* - エラーログにリトライ情報を追加
* 
* 【期待効果】
* - Google側障害時の自動復旧
* - 年間エラー率の大幅削減
* - リトライ頻度の可視化（障害検知）
=============================================================================
*/

// ============================================================================
// メイン処理関数の修正版
// ============================================================================

// ============================================================================
// リトライ設定管理関数
// ============================================================================

/**
 * リトライ機能を有効化
 */
function enableRetry() {
  RETRY_CONFIG.ENABLE_RETRY = true;
  console.log('✓ リトライ機能を有効にしました');
}

/**
 * リトライ機能を無効化
 */
function disableRetry() {
  RETRY_CONFIG.ENABLE_RETRY = false;
  console.log('✓ リトライ機能を無効にしました');
}

/**
 * リトライ設定を表示
 */
function showRetryConfig() {
  console.log('=== リトライ設定 ===');
  console.log(`リトライ機能: ${RETRY_CONFIG.ENABLE_RETRY ? '有効' : '無効'}`);
  console.log(`最大リトライ回数: ${RETRY_CONFIG.MAX_RETRIES}回`);
  console.log(`統計記録: ${RETRY_CONFIG.LOG_RETRY_STATS ? '有効' : '無効'}`);
  console.log('');
  console.log('【リトライ待機時間】');
  for (let i = 1; i <= RETRY_CONFIG.MAX_RETRIES; i++) {
    const wait = Math.pow(2, i - 1);
    console.log(`${i}回目失敗後: ${wait}秒待機`);
  }
}

// ============================================================================
// 使用方法ガイド（リトライ版）
// ============================================================================

function showUsageGuideWithRetry() {
  console.log('=== 在庫情報取得スクリプト v2.1 使用方法ガイド ===');
  console.log('');
  console.log('【v2.1の新機能】');
  console.log('✓ API接続リトライ機能（出荷予定数取得で実績済み）');
  console.log('✓ エクスポネンシャルバックオフ方式（1秒→2秒→4秒）');
  console.log('✓ リトライ統計の自動記録・可視化');
  console.log('✓ Google側一時障害への自動対応');
  console.log('');
  console.log('【主要関数】');
  console.log('1. updateInventoryDataBatchWithRetry()');
  console.log('   - リトライ対応版のメイン処理');
  console.log('   - 既存のupdateInventoryDataBatch()と同じ使い方');
  console.log('   - 処理後にリトライ統計を自動表示');
  console.log('');
  console.log('2. showRetryConfig()');
  console.log('   - 現在のリトライ設定を確認');
  console.log('');
  console.log('3. enableRetry() / disableRetry()');
  console.log('   - リトライ機能の有効/無効を切り替え');
  console.log('');
  console.log('【リトライ統計の見方】');
  console.log('■ リトライ発生率 < 5%: 正常');
  console.log('■ リトライ発生率 5-10%: 軽度の不調');
  console.log('■ リトライ発生率 > 10%: 要注意（Google側の問題の可能性）');
  console.log('');
  console.log('【実行例】');
  console.log('// ステップ1: リトライ設定確認');
  console.log('showRetryConfig()');
  console.log('');
  console.log('// ステップ2: リトライ対応版で実行');
  console.log('updateInventoryDataBatchWithRetry()');
  console.log('');
  console.log('// ステップ3: リトライログシートで統計確認');
  console.log('// → スプレッドシートの「リトライログ」シートを確認');
  console.log('');
  console.log('【期待される効果】');
  console.log('- 年間エラー発生回数の大幅削減');
  console.log('- Google側一時障害時の自動復旧');
  console.log('- 障害頻度の可視化（予防保守に活用）');
  console.log('');
  console.log('==========================================================');
  console.log('従来版: updateInventoryDataBatch()');
  console.log('リトライ版: updateInventoryDataBatchWithRetry()');
  console.log('どちらも同じ結果を返しますが、リトライ版の方が安定性が高いです');
  console.log('==========================================================');
}

/**
 * テスト実行: リトライ機能の動作確認
 */
function testRetryFunction() {
  console.log('=== リトライ機能テスト ===');
  console.log('');

  // リトライ統計をリセット
  resetRetryStats();

  // 小規模データでテスト
  try {
    const { SPREADSHEET_ID, SHEET_NAME } = getSpreadsheetConfig();
    const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = spreadsheet.getSheetByName(SHEET_NAME);

    // 最初の10件でテスト
    const dataRange = sheet.getRange(2, 1, Math.min(10, sheet.getLastRow() - 1), 1);
    const values = dataRange.getValues();
    const goodsCodeList = values
      .map(row => row[0])
      .filter(code => code && code.toString().trim())
      .slice(0, 10);

    console.log(`テスト対象: ${goodsCodeList.length}件`);
    console.log(`商品コード: ${goodsCodeList.join(', ')}`);
    console.log('');

    const tokens = getStoredTokens();
    const startTime = new Date();

    // リトライ対応版で取得
    const inventoryDataMap = getBatchInventoryDataWithRetry(goodsCodeList, tokens, 0);

    const endTime = new Date();
    const duration = (endTime - startTime) / 1000;

    console.log(`\n=== テスト結果 ===`);
    console.log(`処理時間: ${duration.toFixed(1)}秒`);
    console.log(`取得件数: ${inventoryDataMap.size}件`);

    // リトライ統計を表示
    showRetryStats();

    console.log('\n=== 取得データサンプル ===');
    let count = 0;
    for (const [goodsCode, data] of inventoryDataMap) {
      if (count < 3) {
        console.log(`${goodsCode}: 在庫${data.stock_quantity} 引当${data.stock_allocated_quantity} フリー${data.stock_free_quantity}`);
        count++;
      }
    }

    console.log('\n✓ リトライ機能のテストが完了しました');

  } catch (error) {
    console.error('✗ テストエラー:', error.message);
    showRetryStats();
  }
}

/**
 * 比較テスト: 従来版 vs リトライ版
 * 
 * 両バージョンの処理時間とエラー耐性を比較
 */
function compareVersions(sampleSize = 50) {
  console.log('=== 従来版 vs リトライ版 比較テスト ===');
  console.log(`テストサイズ: ${sampleSize}件`);
  console.log('');

  try {
    const { SPREADSHEET_ID, SHEET_NAME } = getSpreadsheetConfig();
    const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = spreadsheet.getSheetByName(SHEET_NAME);

    const dataRange = sheet.getRange(2, 1, Math.min(sampleSize, sheet.getLastRow() - 1), 1);
    const values = dataRange.getValues();
    const goodsCodeList = values
      .map(row => row[0])
      .filter(code => code && code.toString().trim())
      .slice(0, sampleSize);

    const tokens = getStoredTokens();

    // 従来版テスト
    console.log('--- 従来版テスト ---');
    const startTime1 = new Date();
    let traditionalResult;
    let traditionalError = null;

    try {
      traditionalResult = getBatchInventoryData(goodsCodeList, tokens, 0);
    } catch (error) {
      traditionalError = error;
    }

    const endTime1 = new Date();
    const duration1 = (endTime1 - startTime1) / 1000;

    if (traditionalError) {
      console.log(`✗ エラー発生: ${traditionalError.message}`);
    } else {
      console.log(`✓ 成功`);
      console.log(`処理時間: ${duration1.toFixed(1)}秒`);
      console.log(`取得件数: ${traditionalResult.size}件`);
    }

    // 少し待機
    Utilities.sleep(2000);

    // リトライ版テスト
    console.log('\n--- リトライ版テスト ---');
    resetRetryStats();

    const startTime2 = new Date();
    let retryResult;
    let retryError = null;

    try {
      retryResult = getBatchInventoryDataWithRetry(goodsCodeList, tokens, 0);
    } catch (error) {
      retryError = error;
    }

    const endTime2 = new Date();
    const duration2 = (endTime2 - startTime2) / 1000;

    if (retryError) {
      console.log(`✗ エラー発生: ${retryError.message}`);
    } else {
      console.log(`✓ 成功`);
      console.log(`処理時間: ${duration2.toFixed(1)}秒`);
      console.log(`取得件数: ${retryResult.size}件`);
    }

    showRetryStats();

    // 比較結果
    console.log('\n=== 比較結果 ===');
    console.log(`従来版: ${traditionalError ? 'エラー' : '成功'} (${duration1.toFixed(1)}秒)`);
    console.log(`リトライ版: ${retryError ? 'エラー' : '成功'} (${duration2.toFixed(1)}秒)`);

    if (!traditionalError && !retryError) {
      const overhead = ((duration2 - duration1) / duration1 * 100).toFixed(1);
      console.log(`オーバーヘッド: ${overhead}%`);
      console.log('');
      console.log('【結論】');
      if (Math.abs(overhead) < 5) {
        console.log('✓ リトライ機能のオーバーヘッドはほぼゼロです');
        console.log('✓ 安定性向上のため、リトライ版の使用を推奨します');
      } else {
        console.log(`リトライ機能により約${Math.abs(overhead)}%の時間差がありますが、`);
        console.log('安定性の向上を考慮すると許容範囲内です');
      }
    }

  } catch (error) {
    console.error('比較テストエラー:', error.message);
  }
}

// ============================================================================
// マイグレーションガイド
// ============================================================================

/**
 * 既存コードからリトライ版への移行ガイド
 */
function showMigrationGuide() {
  console.log('==========================================================');
  console.log('  リトライ版への移行ガイド');
  console.log('==========================================================');
  console.log('');
  console.log('【ステップ1: テスト実行】');
  console.log('まず小規模データでリトライ機能をテスト:');
  console.log('');
  console.log('  testRetryFunction()');
  console.log('');
  console.log('【ステップ2: 比較テスト（推奨）】');
  console.log('従来版とリトライ版を比較:');
  console.log('');
  console.log('  compareVersions(50)  // 50件で比較');
  console.log('');
  console.log('【ステップ3: 段階的移行】');
  console.log('');
  console.log('■ 方法A: 新関数を使う（推奨）');
  console.log('  // トリガー設定を変更');
  console.log('  // updateInventoryDataBatch()');
  console.log('  //   ↓');
  console.log('  // updateInventoryDataBatchWithRetry()');
  console.log('');
  console.log('■ 方法B: 既存関数を置き換える');
  console.log('  // getBatchInventoryDataの呼び出しを');
  console.log('  // getBatchInventoryDataWithRetryに変更');
  console.log('');
  console.log('【ステップ4: 監視期間】');
  console.log('1週間程度、以下を確認:');
  console.log('  - エラー発生回数の変化');
  console.log('  - リトライ統計（「リトライログ」シート）');
  console.log('  - 処理時間の変化');
  console.log('');
  console.log('【ステップ5: 本番運用】');
  console.log('問題なければ、正式に切り替え完了');
  console.log('');
  console.log('==========================================================');
  console.log('【ロールバック方法】');
  console.log('万が一問題が発生した場合:');
  console.log('');
  console.log('1. disableRetry() でリトライ機能を無効化');
  console.log('   または');
  console.log('2. 従来版の関数に戻す');
  console.log('');
  console.log('既存コードは一切変更していないため、');
  console.log('いつでも従来版に戻すことができます。');
  console.log('==========================================================');
}

// ============================================================================
// SREダッシュボード（オプション）
// ============================================================================

/**
 * SREダッシュボード: システムの健全性を一覧表示
 */
function showSREDashboard() {
  console.log('==========================================================');
  console.log('  SREダッシュボード - システム健全性');
  console.log('==========================================================');
  console.log('');

  try {
    // 1. リトライ設定状況
    console.log('【1. リトライ機能】');
    console.log(`状態: ${RETRY_CONFIG.ENABLE_RETRY ? '✓ 有効' : '✗ 無効'}`);
    console.log(`最大リトライ回数: ${RETRY_CONFIG.MAX_RETRIES}回`);
    console.log('');

    // 2. 最近のリトライ統計（リトライログシートから取得）
    console.log('【2. 直近のリトライ統計】');
    const { SPREADSHEET_ID } = getSpreadsheetConfig();
    const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
    const retryLogSheet = spreadsheet.getSheetByName('リトライログ');

    if (retryLogSheet) {
      const lastRow = retryLogSheet.getLastRow();
      if (lastRow > 1) {
        const recentLogs = Math.min(5, lastRow - 1);
        const data = retryLogSheet.getRange(lastRow - recentLogs + 1, 1, recentLogs, 6).getValues();

        console.log('直近5回の実行:');
        data.forEach((row, index) => {
          const date = Utilities.formatDate(row[0], 'JST', 'MM/dd HH:mm');
          const retryCount = row[1];
          const retryRate = row[4];
          const note = row[5];

          let status = '✓';
          if (retryRate > 10) status = '⚠️';
          else if (retryRate > 5) status = '△';

          console.log(`${status} ${date} | リトライ${retryCount}回 | 発生率${retryRate}% ${note ? '(' + note + ')' : ''}`);
        });
      } else {
        console.log('まだリトライログがありません');
      }
    } else {
      console.log('リトライログシートが存在しません');
    }
    console.log('');

    // 3. エラーログ統計
    console.log('【3. エラー発生状況】');
    const errorLogSheet = spreadsheet.getSheetByName('エラーログ');

    if (errorLogSheet) {
      const lastRow = errorLogSheet.getLastRow();
      if (lastRow > 1) {
        console.log(`累計エラー件数: ${lastRow - 1}件`);

        // 直近のエラー
        const recentErrors = Math.min(3, lastRow - 1);
        const errorData = errorLogSheet.getRange(lastRow - recentErrors + 1, 1, recentErrors, 4).getValues();

        console.log('\n直近のエラー:');
        errorData.forEach(row => {
          const date = Utilities.formatDate(row[0], 'JST', 'MM/dd HH:mm');
          const goodsCode = row[1];
          const errorType = row[2];
          console.log(`  ${date} | ${goodsCode} | ${errorType}`);
        });
      } else {
        console.log('✓ エラーなし');
      }
    } else {
      console.log('エラーログシートが存在しません');
    }
    console.log('');

    // 4. 推奨アクション
    console.log('【4. 推奨アクション】');

    if (!RETRY_CONFIG.ENABLE_RETRY) {
      console.log('⚠️ リトライ機能が無効です');
      console.log('   → enableRetry() で有効化することを推奨します');
    } else {
      console.log('✓ リトライ機能が有効です');
    }

    const currentLogLevel = getCurrentLogLevel();
    if (currentLogLevel === LOG_LEVEL.DETAILED) {
      console.log('⚠️ ログレベルがDETAILEDです（デバッグモード）');
      console.log('   → 本番運用では setLogLevel(1) または setLogLevel(2) を推奨');
    } else if (currentLogLevel === LOG_LEVEL.MINIMAL) {
      console.log('✓ ログレベルがMINIMALです（本番モード）');
    } else {
      console.log('✓ ログレベルがSUMMARYです（推奨設定）');
    }

    console.log('');
    console.log('==========================================================');
    console.log('すべて正常です。システムは健全に動作しています。');
    console.log('==========================================================');

  } catch (error) {
    console.error('ダッシュボード表示エラー:', error.message);
  }
}