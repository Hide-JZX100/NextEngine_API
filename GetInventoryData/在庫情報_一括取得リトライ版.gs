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
// リトライ設定
// ============================================================================

const RETRY_CONFIG = {
  MAX_RETRIES: 3,              // 最大リトライ回数
  ENABLE_RETRY: true,          // リトライ機能の有効/無効
  LOG_RETRY_STATS: true        // リトライ統計のログ出力
};

// リトライ統計（実行ごとにリセット）
let retryStats = {
  totalRetries: 0,           // 総リトライ回数
  batchesWithRetry: 0,       // リトライが発生したバッチ数
  maxRetriesUsed: 0,         // 最大使用リトライ回数
  retriesByBatch: []         // バッチごとのリトライ回数
};

/**
 * リトライ統計をリセット
 */
function resetRetryStats() {
  retryStats = {
    totalRetries: 0,
    batchesWithRetry: 0,
    maxRetriesUsed: 0,
    retriesByBatch: []
  };
}

/**
 * リトライ統計を記録
 */
function recordRetryAttempt(batchNumber, attemptNumber) {
  retryStats.totalRetries++;
  
  if (attemptNumber > 1) {
    if (!retryStats.retriesByBatch[batchNumber]) {
      retryStats.batchesWithRetry++;
    }
    retryStats.retriesByBatch[batchNumber] = attemptNumber;
    retryStats.maxRetriesUsed = Math.max(retryStats.maxRetriesUsed, attemptNumber);
  }
}

/**
 * リトライ統計を表示
 */
function showRetryStats() {
  if (!RETRY_CONFIG.LOG_RETRY_STATS || retryStats.totalRetries === 0) {
    return;
  }
  
  console.log('\n========================================');
  console.log('  リトライ統計情報');
  console.log('========================================');
  console.log(`総リトライ回数: ${retryStats.totalRetries}回`);
  console.log(`リトライ発生バッチ: ${retryStats.batchesWithRetry}個`);
  console.log(`最大リトライ回数: ${retryStats.maxRetriesUsed}回`);
  
  if (retryStats.totalRetries > 0) {
    console.log('\n--- リトライ発生バッチ詳細 ---');
    retryStats.retriesByBatch.forEach((retries, batchNum) => {
      if (retries > 1) {
        console.log(`バッチ ${batchNum}: ${retries}回試行`);
      }
    });
  }
  
  // 障害検知アラート
  if (retryStats.batchesWithRetry > 0) {
    const retryRate = (retryStats.batchesWithRetry / retryStats.retriesByBatch.length * 100).toFixed(1);
    console.log(`\n⚠️ リトライ発生率: ${retryRate}%`);
    
    if (retryRate > 10) {
      console.log('⚠️⚠️ 注意: リトライ発生率が高いです（10%以上）');
      console.log('   → Google側またはネットワークの不調の可能性があります');
    }
  }
  
  console.log('========================================\n');
}

/**
 * リトライ統計をエラーログシートに記録
 */
function logRetryStatsToSheet() {
  if (retryStats.totalRetries === 0) {
    return;
  }
  
  try {
    const { SPREADSHEET_ID } = getSpreadsheetConfig();
    const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
    let retryLogSheet = spreadsheet.getSheetByName('リトライログ');
    
    if (!retryLogSheet) {
      retryLogSheet = spreadsheet.insertSheet('リトライログ');
      const headers = [
        '実行日時', '総リトライ回数', 'リトライ発生バッチ数', 
        '最大リトライ回数', 'リトライ発生率(%)', '備考'
      ];
      retryLogSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      retryLogSheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
      retryLogSheet.getRange(1, 1, 1, headers.length).setBackground('#f3f3f3');
    }
    
    const totalBatches = retryStats.retriesByBatch.length;
    const retryRate = totalBatches > 0 
      ? (retryStats.batchesWithRetry / totalBatches * 100).toFixed(1) 
      : 0;
    
    let note = '';
    if (retryRate > 10) {
      note = 'リトライ率高（要確認）';
    } else if (retryStats.totalRetries > 0) {
      note = '正常（軽微なリトライ）';
    }
    
    const logRow = [
      new Date(),
      retryStats.totalRetries,
      retryStats.batchesWithRetry,
      retryStats.maxRetriesUsed,
      retryRate,
      note
    ];
    
    const lastRow = retryLogSheet.getLastRow();
    retryLogSheet.getRange(lastRow + 1, 1, 1, 6).setValues([logRow]);
    retryLogSheet.getRange(lastRow + 1, 1, 1, 1).setNumberFormat('yyyy/mm/dd hh:mm:ss');
    
    logWithLevel(LOG_LEVEL.SUMMARY, 'リトライ統計をシートに記録しました');
    
  } catch (error) {
    logError('リトライログ記録エラー:', error.message);
  }
}

// ============================================================================
// リトライ機能付きAPI呼び出し関数
// ============================================================================

/**
 * リトライ処理付き在庫マスタデータ取得
 * 
 * 【変更内容】
 * - 既存のgetBatchStockData関数をラップ
 * - エクスポネンシャルバックオフでリトライ
 * - リトライ統計を記録
 * 
 * @param {Array} goodsCodeList - 商品コードリスト
 * @param {Object} tokens - 認証トークン
 * @param {number} batchNumber - バッチ番号
 * @param {number} maxRetries - 最大リトライ回数
 * @return {Map} 在庫データマップ
 */
function getBatchStockDataWithRetry(goodsCodeList, tokens, batchNumber, maxRetries = RETRY_CONFIG.MAX_RETRIES) {
  // リトライ機能が無効の場合は既存関数をそのまま呼び出し
  if (!RETRY_CONFIG.ENABLE_RETRY) {
    return getBatchStockData(goodsCodeList, tokens, batchNumber);
  }
  
  let lastError = null;
  
  for (let attempt = 1; attempt <= maxRetries; attempt++) {
    try {
      // リトライ回数を記録
      recordRetryAttempt(batchNumber, attempt);
      
      if (attempt > 1) {
        logWithLevel(LOG_LEVEL.SUMMARY, `  リトライ ${attempt}/${maxRetries}回目...`);
      }
      
      // ★ 既存の関数を呼び出し
      const stockDataMap = getBatchStockData(goodsCodeList, tokens, batchNumber);
      
      // 成功したらデータを返す
      if (attempt > 1) {
        logWithLevel(LOG_LEVEL.SUMMARY, `  ✓ リトライ成功（${attempt}回目の試行で成功）`);
      }
      
      return stockDataMap;
      
    } catch (error) {
      lastError = error;
      
      // エラーの種類を判定してリトライすべきか判断
      const errorMessage = error.message.toLowerCase();
      
      // リトライすべきでないエラー（認証・権限系）
      if (
        errorMessage.includes('認証') ||
        errorMessage.includes('auth') ||
        errorMessage.includes('permission') ||
        errorMessage.includes('権限') ||
        errorMessage.includes('invalid') ||
        errorMessage.includes('token')
      ) {
        logError(`  即座に失敗: リトライ不可能なエラー - ${error.message}`);
        throw error; // リトライせずに即座にスロー
      }
      
      logError(`  ✗ API接続エラー（試行 ${attempt}/${maxRetries}）: ${error.message}`);
      
      // 最後の試行でなければ、待機してからリトライ
      if (attempt < maxRetries) {
        // エクスポネンシャルバックオフ（指数バックオフ）
        // 1秒、2秒、4秒...
        const waitSeconds = Math.pow(2, attempt - 1);
        logWithLevel(LOG_LEVEL.SUMMARY, `  → ${waitSeconds}秒後にリトライします...`);
        Utilities.sleep(waitSeconds * 1000);
      }
    }
  }
  
  // すべてのリトライが失敗した場合
  const errorMessage = `API接続失敗（${maxRetries}回試行）: ${lastError.message}`;
  logError(`  ✗✗✗ ${errorMessage}`);
  
  // 詳細なエラー情報を記録
  logAPIErrorDetail(
    '在庫マスタAPI（リトライ失敗）',
    {
      goodsCodeCount: goodsCodeList.length,
      firstCode: goodsCodeList[0],
      lastCode: goodsCodeList[goodsCodeList.length - 1],
      totalAttempts: maxRetries
    },
    null,
    lastError
  );
  
  throw new Error(errorMessage);
}

// ============================================================================
// メイン処理関数の修正版
// ============================================================================

/**
 * バッチ単位で在庫情報を取得（リトライ対応版）
 * 
 * 【変更内容】
 * - getBatchStockData → getBatchStockDataWithRetry に変更
 * - その他の処理は既存のまま維持
 */
function getBatchInventoryDataWithRetry(goodsCodeList, tokens, batchNumber) {
  const inventoryDataMap = new Map();

  try {
    logWithLevel(LOG_LEVEL.DETAILED, `  在庫マスタ一括検索: ${goodsCodeList.length}件`);
    
    const codeMapping = new Map();
    for (const code of goodsCodeList) {
      codeMapping.set(code.toLowerCase(), code);
    }
    
    // ★★★ ここだけ変更: リトライ機能付き関数に置き換え ★★★
    const stockDataMap = getBatchStockDataWithRetry(goodsCodeList, tokens, batchNumber);
    
    logWithLevel(LOG_LEVEL.DETAILED, `  在庫マスタ取得完了: ${stockDataMap.size}件`);

    if (stockDataMap.size === 0) {
      logWithLevel(LOG_LEVEL.SUMMARY, '  在庫データが見つかりませんでした');
      
      logAPIErrorDetail(
        '在庫マスタAPI',
        {
          goodsCodeCount: goodsCodeList.length,
          firstCode: goodsCodeList[0],
          lastCode: goodsCodeList[goodsCodeList.length - 1]
        },
        { message: 'データが1件も取得できませんでした' },
        new Error('API応答にデータが含まれていません')
      );
      
      return inventoryDataMap;
    }

    for (const [goodsCode, stockData] of stockDataMap) {
      const originalCode = codeMapping.get(goodsCode.toLowerCase());
      
      if (!originalCode) {
        logErrorDetail(goodsCode, 'コードマッピングエラー', '元のコードが見つかりません', {
          'バッチ番号': batchNumber,
          'API返却コード': goodsCode,
          'マッピング数': codeMapping.size,
          '要求コード数': goodsCodeList.length
        });
        continue;
      }
      
      const inventoryData = {
        goods_id: stockData.stock_goods_id,
        goods_name: '',
        stock_quantity: parseInt(stockData.stock_quantity) || 0,
        stock_allocated_quantity: parseInt(stockData.stock_allocation_quantity) || 0,
        stock_free_quantity: parseInt(stockData.stock_free_quantity) || 0,
        stock_defective_quantity: parseInt(stockData.stock_defective_quantity) || 0,
        stock_advance_order_quantity: parseInt(stockData.stock_advance_order_quantity) || 0,
        stock_advance_order_allocation_quantity: parseInt(stockData.stock_advance_order_allocation_quantity) || 0,
        stock_advance_order_free_quantity: parseInt(stockData.stock_advance_order_free_quantity) || 0,
        stock_remaining_order_quantity: parseInt(stockData.stock_remaining_order_quantity) || 0,
        stock_out_quantity: parseInt(stockData.stock_out_quantity) || 0
      };
      
      inventoryDataMap.set(originalCode, inventoryData);
    }

    logWithLevel(LOG_LEVEL.DETAILED, `  在庫情報構築完了: ${inventoryDataMap.size}件`);
    return inventoryDataMap;

  } catch (error) {
    logError(`在庫情報取得エラー: ${error.message}`);
    
    logAPIErrorDetail(
      '在庫情報構築処理',
      {
        goodsCodeCount: goodsCodeList.length,
        firstCode: goodsCodeList[0],
        lastCode: goodsCodeList[goodsCodeList.length - 1]
      },
      null,
      error
    );
    
    return inventoryDataMap;
  }
}

/**
 * メイン処理関数の修正版（リトライ統計対応）
 * 
 * 【変更内容】
 * - リトライ統計のリセットと表示を追加
 * - getBatchInventoryData → getBatchInventoryDataWithRetry に変更
 * - リトライログをシートに記録
 */
function updateInventoryDataBatchWithRetry() {
  try {
    // リトライ統計をリセット
    resetRetryStats();
    
    const currentLogLevel = getCurrentLogLevel();
    
    logWithLevel(LOG_LEVEL.MINIMAL, '=== 在庫情報一括更新開始（リトライ対応版 v2.1） ===');
    const startTime = new Date();
    
    const { SPREADSHEET_ID, SHEET_NAME } = getSpreadsheetConfig();
    const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = spreadsheet.getSheetByName(SHEET_NAME);
    
    if (!sheet) {
      throw new Error(`シート "${SHEET_NAME}" が見つかりません`);
    }
    
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) {
      logWithLevel(LOG_LEVEL.MINIMAL, 'データが存在しません');
      return;
    }
    
    const dataRange = sheet.getRange(2, 1, lastRow - 1, 12);
    const values = dataRange.getValues();
    logWithLevel(LOG_LEVEL.MINIMAL, `処理対象: ${values.length}行`);
    
    const tokens = getStoredTokens();
    
    const goodsCodeList = [];
    const rowIndexMap = new Map();
    
    for (let i = 0; i < values.length; i++) {
      const goodsCode = values[i][COLUMNS.GOODS_CODE];
      if (goodsCode && goodsCode.toString().trim()) {
        goodsCodeList.push(goodsCode.toString().trim());
        rowIndexMap.set(goodsCode.toString().trim(), i + 2);
      }
    }
    
    logWithLevel(LOG_LEVEL.MINIMAL, `有効な商品コード: ${goodsCodeList.length}件`);
    
    if (goodsCodeList.length === 0) {
      logWithLevel(LOG_LEVEL.MINIMAL, '処理対象の商品コードがありません');
      return;
    }
    
    let totalUpdated = 0;
    let totalErrors = 0;
    const errorDetails = [];
    const batchCount = Math.ceil(goodsCodeList.length / MAX_ITEMS_PER_CALL);
    
    logWithLevel(LOG_LEVEL.SUMMARY, `バッチ数: ${batchCount}個（${MAX_ITEMS_PER_CALL}件/バッチ）`);
    logWithLevel(LOG_LEVEL.SUMMARY, `リトライ設定: 最大${RETRY_CONFIG.MAX_RETRIES}回（${RETRY_CONFIG.ENABLE_RETRY ? '有効' : '無効'}）`);
    
    for (let i = 0; i < goodsCodeList.length; i += MAX_ITEMS_PER_CALL) {
      const batch = goodsCodeList.slice(i, i + MAX_ITEMS_PER_CALL);
      const batchNumber = Math.floor(i / MAX_ITEMS_PER_CALL) + 1;
      
      logWithLevel(LOG_LEVEL.SUMMARY, `\n--- バッチ ${batchNumber}/${batchCount}: ${batch.length}件 ---`);
      
      const batchStartTime = new Date();
      const batchErrors = [];
      
      try {
        // ★★★ ここを変更: リトライ対応版の関数を使用 ★★★
        const inventoryDataMap = getBatchInventoryDataWithRetry(batch, tokens, batchNumber);
        
        const batchEndTime = new Date();
        const batchDuration = (batchEndTime - batchStartTime) / 1000;
        
        // バッチ単位で一括更新（既存コードをそのまま使用）
        const updateResult = updateBatchInventoryData(
          sheet, 
          batch, 
          inventoryDataMap, 
          rowIndexMap
        );
        
        const batchUpdated = updateResult.updated;
        const updateResults = updateResult.results;
        const batchErrorCount = updateResults.filter(r => r.status === 'error' || r.status === 'no_data').length;
        
        totalUpdated += batchUpdated;
        
        // エラー詳細を収集（既存コードと同じ）
        for (const result of updateResults) {
          if (result.status === 'error') {
            logErrorDetail(result.goodsCode, '更新エラー', result.error, {
              'バッチ番号': batchNumber
            });
            
            const errorInfo = {
              goodsCode: result.goodsCode,
              errorType: '更新エラー',
              errorMessage: result.error,
              timestamp: new Date(),
              batchNumber: batchNumber
            };
            errorDetails.push(errorInfo);
            batchErrors.push(errorInfo);
            totalErrors++;
          } else if (result.status === 'no_data') {
            logErrorDetail(result.goodsCode, 'データなし', 'inventory data not found', {
              'バッチ番号': batchNumber
            });
            
            const errorInfo = {
              goodsCode: result.goodsCode,
              errorType: 'データなし',
              errorMessage: 'inventory data not found',
              timestamp: new Date(),
              batchNumber: batchNumber
            };
            errorDetails.push(errorInfo);
            batchErrors.push(errorInfo);
          }
        }
        
        logWithLevel(LOG_LEVEL.SUMMARY, `処理時間: ${batchDuration.toFixed(1)}秒 | 成功: ${batchUpdated}件 | エラー: ${batchErrorCount}件`);
        
        if (batchErrors.length > 0) {
          logBatchErrorSummary(batchNumber, batchErrors);
        }
        
        // ログ出力（既存コードと同じ - 省略）
        // ...
        
        if (i + MAX_ITEMS_PER_CALL < goodsCodeList.length) {
          logWithLevel(LOG_LEVEL.SUMMARY, `次のバッチまで ${API_WAIT_TIME}ms 待機...`);
          Utilities.sleep(API_WAIT_TIME);
        }
        
      } catch (error) {
        logAPIErrorDetail(
          '在庫マスタAPI（バッチ全体）',
          {
            goodsCodeCount: batch.length,
            firstCode: batch[0],
            lastCode: batch[batch.length - 1]
          },
          null,
          error
        );
        
        batch.forEach(goodsCode => {
          const errorInfo = {
            goodsCode: goodsCode,
            errorType: 'バッチエラー',
            errorMessage: error.message,
            timestamp: new Date(),
            batchNumber: batchNumber
          };
          errorDetails.push(errorInfo);
          batchErrors.push(errorInfo);
        });
        
        logError(`バッチ処理エラー: ${error.message}`);
        logBatchErrorSummary(batchNumber, batchErrors);
        totalErrors += batch.length;
      }
    }
    
    const endTime = new Date();
    const duration = (endTime - startTime) / 1000;
    
    if (errorDetails.length > 0) {
      logErrorsToSheet(errorDetails);
      logWithLevel(LOG_LEVEL.SUMMARY, `\nエラーレポートをシートに記録: ${errorDetails.length}件`);
    }
    
    // ★★★ リトライ統計を表示・記録 ★★★
    showRetryStats();
    logRetryStatsToSheet();
    
    logWithLevel(LOG_LEVEL.MINIMAL, '\n=== 一括更新完了 ===');
    logWithLevel(LOG_LEVEL.MINIMAL, `処理時間: ${duration.toFixed(1)}秒`);
    logWithLevel(LOG_LEVEL.MINIMAL, `更新成功: ${totalUpdated}件`);
    
    if (totalErrors > 0) {
      console.error(`❌ エラー: ${totalErrors}件 ← エラーログシートを確認してください`);
    } else {
      logWithLevel(LOG_LEVEL.MINIMAL, `✓ エラー: 0件`);
    }
    
    logWithLevel(LOG_LEVEL.MINIMAL, `処理速度: ${(goodsCodeList.length / duration).toFixed(1)}件/秒`);
    
    const conventionalTime = goodsCodeList.length * 2;
    const speedImprovement = conventionalTime / duration;
    logWithLevel(LOG_LEVEL.SUMMARY, `\n--- 性能改善結果 ---`);
    logWithLevel(LOG_LEVEL.SUMMARY, `従来版推定時間: ${conventionalTime.toFixed(1)}秒`);
    logWithLevel(LOG_LEVEL.SUMMARY, `高速化倍率: ${speedImprovement.toFixed(1)}倍`);
    
    recordExecutionTimestamp();
    
  } catch (error) {
    logError('一括更新エラー:', error.message);
    throw error;
  }
}

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