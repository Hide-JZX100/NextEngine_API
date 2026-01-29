/**
使用方法
showUsageGuide関数を実行してユーザーガイドを確認してください。

主要な処理を実行する関数
updateInventoryDataBatch()
このスクリプトのメインとなる関数です。
スプレッドシート上の全商品コードを取得し、
MAX_ITEMS_PER_CALL（デフォルト1000件）ごとにバッチに分割して、
getBatchInventoryData関数で在庫情報を一括取得します。
取得したデータでスプレッドシートを更新し、
処理が完了したら合計時間や更新件数、エラー件数などをレポートとして出力します。

getBatchInventoryData(goodsCodeList, tokens)
バッチ処理の中核を担う関数です。
複数の商品コードに対応する在庫情報を一括で取得し、
商品コードをキーとした在庫情報のマップ（Map）を返します。
内部ではgetBatchStockDataを呼び出し、在庫マスタAPIから直接データを取得します。

API呼び出しを担当する関数
fetchInventoryWithSingleAPI(goodsCodes, tokens)
単一API版の性能比較テストで使用される関数です。
商品名を含まない在庫情報のみを、在庫マスタAPIへの1回の呼び出しで取得します。
この関数は、APIコールの回数を減らすことで処理速度を向上させるという、
このスクリプトの主要なコンセプトを示しています。

fetchInventoryWithDoubleAPI(goodsCodes, tokens)
二重API版の性能比較テストで使用される関数です。
商品マスタAPIで商品IDと商品名を取得し、
次に在庫マスタAPIで在庫情報を取得するという、
APIを2回呼び出す従来の方式をシミュレートします。
この関数とfetchInventoryWithSingleAPIを比較することで、
APIコール数削減による性能向上が評価されます。

getBatchGoodsData(goodsCodeList, tokens)
商品マスタAPI（／api_v1_master_goods／search）を呼び出し、
複数の商品について基本情報（商品ID、商品名など）をまとめて取得します。
goods_id-inというパラメータを使って、複数の商品を一度に検索する点が特徴です。

テストおよびユーティリティ関数
compareAPIVersions(sampleSize)
fetchInventoryWithDoubleAPIとfetchInventoryWithSingleAPIを実際に実行し、
処理時間を比較するテスト関数です。
APIコールの回数を減らすことによる高速化効果を、具体的な数値で示します。

testSingleAPIFunction(maxItems)
単一API版のfetchInventoryWithSingleAPIが正しく動作するかを確認するためのテスト関数です。
小規模な件数で実行し、取得したデータの内容をログに出力します。

testBatchProcessing(maxItems)
メインのupdateInventoryDataBatch関数と同様のバッチ処理を、
小規模な件数でテストする関数です。
実際の更新処理を行う前に、全体の流れとgetBatchInventoryData関数の動作を確認するのに役立ちます。

comparePerformance(sampleSize)
このスクリプトの「一括版」と、架空の「従来版（1件ずつAPIを叩く）」の推定処理時間を比較する関数です。
高速化の倍率を計算し、大幅な時間短縮効果を数値で示します。

updateRowWithInventoryData(sheet, rowIndex, inventoryData)
取得した在庫情報（inventoryData）を基に、
スプレッドシートの指定された行（rowIndex）の在庫関連の列を更新します。

logErrorsToSheet(errorDetails)
処理中に発生したエラーの詳細を、スプレッドシート上の「エラーログ」シートに記録します。
これにより、どの商品でどのような問題が発生したかを後から確認できます。

showCurrentProperties()
現在のスクリプトプロパティの設定内容をログに出力します。

debugSpecificProducts()
大文字、小文字の表記ゆれがある場合に在庫情報が取得できない場合があるので、テスト関数を作成してその原因の特定を行う。

recordExecutionTimestamp()
実行完了日時を指定されたシートに記録する関数
シート名はスクリプトプロパティに保存するので、任意のシート名を設定してください。
また、実行完了日時はそのシートのA1セルに記録するようにしていますので、
A1セルには他の情報を入力しないようにしてください。

showUsageGuide()
スクリプトの主要な機能、使用方法、そして期待される効果について説明します。
=============================================================================
*/

/**
 * バッチ単位で在庫データを一括更新する関数（最適化版）
 */
function updateBatchInventoryData(sheet, batch, inventoryDataMap, rowIndexMap) {
  const updateData = [];
  const results = [];

  // ステップ1: 行番号でソートして連続した範囲を特定
  const sortedItems = [];

  for (const goodsCode of batch) {
    const inventoryData = inventoryDataMap.get(goodsCode);
    const rowIndex = rowIndexMap.get(goodsCode);

    if (inventoryData && rowIndex) {
      sortedItems.push({
        goodsCode: goodsCode,
        rowIndex: rowIndex,
        inventoryData: inventoryData
      });
    } else {
      results.push({
        goodsCode: goodsCode,
        status: 'no_data'
      });
    }
  }

  // 行番号でソート
  sortedItems.sort((a, b) => a.rowIndex - b.rowIndex);

  // ステップ2: 連続した範囲をグループ化
  const rangeGroups = [];
  let currentGroup = null;

  for (const item of sortedItems) {
    if (!currentGroup || item.rowIndex !== currentGroup.endRow + 1) {
      if (currentGroup) {
        rangeGroups.push(currentGroup);
      }
      currentGroup = {
        startRow: item.rowIndex,
        endRow: item.rowIndex,
        items: [item]
      };
    } else {
      currentGroup.endRow = item.rowIndex;
      currentGroup.items.push(item);
    }
  }

  if (currentGroup) {
    rangeGroups.push(currentGroup);
  }

  // ステップ3: 各グループごとに一括更新
  let totalUpdated = 0;

  for (const group of rangeGroups) {
    try {
      const updateValues = group.items.map(item => [
        item.inventoryData.stock_quantity || 0,
        item.inventoryData.stock_allocated_quantity || 0,
        item.inventoryData.stock_free_quantity || 0,
        item.inventoryData.stock_advance_order_quantity || 0,
        item.inventoryData.stock_advance_order_allocation_quantity || 0,
        item.inventoryData.stock_advance_order_free_quantity || 0,
        item.inventoryData.stock_defective_quantity || 0,
        item.inventoryData.stock_remaining_order_quantity || 0,
        item.inventoryData.stock_out_quantity || 0
      ]);

      const range = sheet.getRange(
        group.startRow,
        COLUMNS.STOCK_QTY + 1,
        updateValues.length,
        9
      );
      range.setValues(updateValues);

      totalUpdated += group.items.length;

      for (const item of group.items) {
        results.push({
          goodsCode: item.goodsCode,
          status: 'success',
          stock: item.inventoryData.stock_quantity
        });
      }

    } catch (error) {
      for (const item of group.items) {
        results.push({
          goodsCode: item.goodsCode,
          status: 'error',
          error: error.message
        });
      }

      logError(`グループ更新エラー (行 ${group.startRow}-${group.endRow}): ${error.message}`);
    }
  }

  return {
    updated: totalUpdated,
    results: results
  };
}

// ============================================================================
// メイン処理関数
// ============================================================================
function updateInventoryDataBatch() {
  try {
    const currentLogLevel = getCurrentLogLevel();

    logWithLevel(LOG_LEVEL.MINIMAL, '=== 在庫情報一括更新開始 ===');
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

    for (let i = 0; i < goodsCodeList.length; i += MAX_ITEMS_PER_CALL) {
      const batch = goodsCodeList.slice(i, i + MAX_ITEMS_PER_CALL);
      const batchNumber = Math.floor(i / MAX_ITEMS_PER_CALL) + 1;

      logWithLevel(LOG_LEVEL.SUMMARY, `\n--- バッチ ${batchNumber}/${batchCount}: ${batch.length}件 ---`);

      const batchStartTime = new Date();
      const batchErrors = [];

      try {
        const inventoryDataMap = getBatchInventoryData(batch, tokens, batchNumber);

        const batchEndTime = new Date();
        const batchDuration = (batchEndTime - batchStartTime) / 1000;

        // ========================================================      
        // 【新コード - これに置き換え】
        // バッチ単位で一括更新
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

        // エラー詳細を収集
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

        // ========================================================
        // ★★★ ここまで変更 ★★★
        // ========================================================

        // ↓↓↓ 以降は既存のコードをそのまま維持 ↓↓↓

        logWithLevel(LOG_LEVEL.SUMMARY, `処理時間: ${batchDuration.toFixed(1)}秒 | 成功: ${batchUpdated}件 | エラー: ${batchErrorCount}件`);

        if (batchErrors.length > 0) {
          logBatchErrorSummary(batchNumber, batchErrors);
        }

        if (currentLogLevel >= LOG_LEVEL.SUMMARY && updateResults.length > 0) {
          const displayCount = Math.min(3, updateResults.length);

          logWithLevel(LOG_LEVEL.SUMMARY, '【最初の3件】');
          for (let j = 0; j < displayCount; j++) {
            const result = updateResults[j];
            if (result.status === 'success') {
              logWithLevel(LOG_LEVEL.SUMMARY, ` ✓ ${result.goodsCode}: 在庫${result.stock}`);
            } else if (result.status === 'error') {
              logWithLevel(LOG_LEVEL.SUMMARY, ` ✗ ${result.goodsCode}: ${result.error}`);
            } else {
              logWithLevel(LOG_LEVEL.SUMMARY, ` - ${result.goodsCode}: データなし`);
            }
          }

          if (updateResults.length > 6) {
            logWithLevel(LOG_LEVEL.SUMMARY, ` ... 中間 ${updateResults.length - 6}件 省略 ...`);
          }

          if (updateResults.length > 3) {
            logWithLevel(LOG_LEVEL.SUMMARY, '【最後の3件】');
            const startIdx = Math.max(displayCount, updateResults.length - 3);
            for (let j = startIdx; j < updateResults.length; j++) {
              const result = updateResults[j];
              if (result.status === 'success') {
                logWithLevel(LOG_LEVEL.SUMMARY, ` ✓ ${result.goodsCode}: 在庫${result.stock}`);
              } else if (result.status === 'error') {
                logWithLevel(LOG_LEVEL.SUMMARY, ` ✗ ${result.goodsCode}: ${result.error}`);
              } else {
                logWithLevel(LOG_LEVEL.SUMMARY, ` - ${result.goodsCode}: データなし`);
              }
            }
          }
        }

        if (currentLogLevel >= LOG_LEVEL.DETAILED) {
          logWithLevel(LOG_LEVEL.DETAILED, '\n【全件詳細】');
          for (const result of updateResults) {
            if (result.status === 'success') {
              logWithLevel(LOG_LEVEL.DETAILED, ` ✓ ${result.goodsCode}: 在庫${result.stock}`);
            } else if (result.status === 'error') {
              logWithLevel(LOG_LEVEL.DETAILED, ` ✗ ${result.goodsCode}: ${result.error}`);
            } else {
              logWithLevel(LOG_LEVEL.DETAILED, ` - ${result.goodsCode}: データなし`);
            }
          }
        }

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

function getBatchInventoryData(goodsCodeList, tokens, batchNumber) {
  const inventoryDataMap = new Map();

  try {
    logWithLevel(LOG_LEVEL.DETAILED, `  在庫マスタ一括検索: ${goodsCodeList.length}件`);

    const codeMapping = new Map();
    for (const code of goodsCodeList) {
      codeMapping.set(code.toLowerCase(), code);
    }

    const stockDataMap = getBatchStockData(goodsCodeList, tokens, batchNumber);
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

// ユーティリティ関数(旧表示)

function updateRowWithInventoryData(sheet, rowIndex, inventoryData) {
  const updateValues = [
    inventoryData.stock_quantity || 0,
    inventoryData.stock_allocated_quantity || 0,
    inventoryData.stock_free_quantity || 0,
    inventoryData.stock_advance_order_quantity || 0,
    inventoryData.stock_advance_order_allocation_quantity || 0,
    inventoryData.stock_advance_order_free_quantity || 0,
    inventoryData.stock_defective_quantity || 0,
    inventoryData.stock_remaining_order_quantity || 0,
    inventoryData.stock_out_quantity || 0
  ];

  const range = sheet.getRange(rowIndex, COLUMNS.STOCK_QTY + 1, 1, updateValues.length);
  range.setValues([updateValues]);
}

function logErrorsToSheet(errorDetails) {
  try {
    const { SPREADSHEET_ID } = getSpreadsheetConfig();
    const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
    let errorSheet = spreadsheet.getSheetByName('エラーログ');

    if (!errorSheet) {
      errorSheet = spreadsheet.insertSheet('エラーログ');
      const headers = [
        '発生日時', '商品コード', 'エラー種別',
        'エラー内容', 'バッチ番号', '処理日時'
      ];
      errorSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      errorSheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
    }

    const errorRows = errorDetails.map(error => [
      error.timestamp,
      error.goodsCode,
      error.errorType,
      error.errorMessage,
      error.batchNumber,
      new Date()
    ]);

    if (errorRows.length > 0) {
      const lastRow = errorSheet.getLastRow();
      const range = errorSheet.getRange(lastRow + 1, 1, errorRows.length, 6);
      range.setValues(errorRows);

      errorSheet.getRange(lastRow + 1, 1, errorRows.length, 1)
        .setNumberFormat('yyyy/mm/dd hh:mm:ss');
      errorSheet.getRange(lastRow + 1, 6, errorRows.length, 1)
        .setNumberFormat('yyyy/mm/dd hh:mm:ss');
    }

    console.log(`エラーログに${errorRows.length}件を記録しました`);

  } catch (error) {
    console.error('エラーログ記録中にエラーが発生:', error.message);
  }
}

function recordExecutionTimestamp() {
  try {
    const properties = PropertiesService.getScriptProperties();
    const spreadsheetId = properties.getProperty('SPREADSHEET_ID');
    const sheetName = properties.getProperty('LOG_SHEET_NAME');

    if (!spreadsheetId || !sheetName) {
      throw new Error('スクリプトプロパティ SPREADSHEET_ID または LOG_SHEET_NAME が設定されていません。');
    }

    const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    const sheet = spreadsheet.getSheetByName(sheetName);

    if (!sheet) {
      console.error(`シート "${sheetName}" が見つかりません。日時の記録をスキップします。`);
      return;
    }

    sheet.getRange(1, 1).setValue(
      Utilities.formatDate(new Date(), 'JST', 'MM月dd日HH時mm分ss秒')
    );
    console.log(`実行日時をシート "${sheetName}" のA1セルに記録しました。`);

  } catch (error) {
    console.error('実行日時の記録中にエラーが発生しました:', error.message);
  }
}

// ============================================================================
// テスト・管理用関数
// ============================================================================

function showUsageGuide() {
  console.log('=== 在庫情報取得スクリプト v2.0 使用方法ガイド ===');
  console.log('');
  console.log('【主要関数】');
  console.log('1. updateInventoryDataBatch()');
  console.log('   - 全商品の在庫情報を一括処理で更新');
  console.log('   - 1000件ずつのバッチで自動分割処理');
  console.log('   - ログレベルに応じた出力（デフォルト: SUMMARY）');
  console.log('');
  console.log('【ログレベル管理】');
  console.log('2. showCurrentLogLevel()');
  console.log('   - 現在のログレベルを確認');
  console.log('');
  console.log('3. setLogLevel(レベル)');
  console.log('   - ログレベルを変更');
  console.log('   - setLogLevel(1): MINIMAL（本番運用、最速）');
  console.log('   - setLogLevel(2): SUMMARY（デフォルト、推奨）');
  console.log('   - setLogLevel(3): DETAILED（デバッグ用）');
  console.log('');
  console.log('【期待される効果】');
  console.log('- 処理時間: 20-25秒（3000件の場合）');
  console.log('- APIコール数: 大幅削減（1000件を3回で処理）');
  console.log('- ログ出力: 約99%削減（可読性向上）');
  console.log('- エラー時: 自動的に詳細情報を出力');
  console.log('');
  console.log('【推奨実行手順】');
  console.log('1. showCurrentLogLevel() でログレベル確認');
  console.log('2. 必要に応じて setLogLevel(1) で本番モードに変更');
  console.log('3. updateInventoryDataBatch() で全件更新');
  console.log('4. エラーがあれば「エラーログ」シートを確認');
  console.log('');
  console.log('【v2.0の改善点】');
  console.log('✓ ログ出力の最適化（実行時間5-9秒短縮）');
  console.log('✓ エラー発生時の詳細情報自動出力');
  console.log('✓ ログレベル切り替え機能');
  console.log('✓ 最初3件+最後3件表示方式');
}

function testBatchProcessing(maxItems = 10) {
  try {
    console.log(`=== バッチ処理テスト（最大${maxItems}件） ===`);

    const { SPREADSHEET_ID, SHEET_NAME } = getSpreadsheetConfig();
    const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = spreadsheet.getSheetByName(SHEET_NAME);
    const lastRow = sheet.getLastRow();

    if (lastRow <= 1) {
      console.log('テスト用データが存在しません');
      return;
    }

    const dataRange = sheet.getRange(2, 1, Math.min(maxItems, lastRow - 1), 1);
    const values = dataRange.getValues();
    const goodsCodeList = values
      .map(row => row[0])
      .filter(code => code && code.toString().trim())
      .slice(0, maxItems);

    console.log(`テスト対象商品コード: ${goodsCodeList.join(', ')}`);

    const tokens = getStoredTokens();

    const startTime = new Date();
    const inventoryDataMap = getBatchInventoryData(goodsCodeList, tokens, 0);
    const endTime = new Date();
    const duration = (endTime - startTime) / 1000;

    console.log(`\n=== テスト結果 ===`);
    console.log(`処理時間: ${duration.toFixed(1)}秒`);
    console.log(`取得件数: ${inventoryDataMap.size}件`);
    console.log(`処理速度: ${(goodsCodeList.length / duration).toFixed(1)}件/秒`);

    for (const [goodsCode, data] of inventoryDataMap) {
      console.log(`${goodsCode}: 在庫${data.stock_quantity} 引当${data.stock_allocated_quantity} フリー${data.stock_free_quantity}`);
    }

  } catch (error) {
    console.error('バッチテストエラー:', error.message);
    throw error;
  }
}

function showCurrentProperties() {
  const properties = PropertiesService.getScriptProperties();
  console.log('=== 現在のスクリプトプロパティ設定 ===');
  console.log(`SPREADSHEET_ID: ${properties.getProperty('SPREADSHEET_ID') || '未設定'}`);
  console.log(`SHEET_NAME: ${properties.getProperty('SHEET_NAME') || '未設定'}`);
  console.log(`LOG_SHEET_NAME: ${properties.getProperty('LOG_SHEET_NAME') || '未設定'}`);
  console.log(`LOG_LEVEL: ${properties.getProperty('LOG_LEVEL') || '未設定（デフォルト: 2=SUMMARY）'}`);
  console.log('');
  console.log('認証情報:');
  console.log(`ACCESS_TOKEN: ${properties.getProperty('ACCESS_TOKEN') ? '設定済み' : '未設定'}`);
  console.log(`REFRESH_TOKEN: ${properties.getProperty('REFRESH_TOKEN') ? '設定済み' : '未設定'}`);
}

/**
 * クイックスタートガイド
 * 初めて使用する際に実行してください
*/
function quickStartGuide() {
  console.log('==========================================================');
  console.log('   在庫情報取得スクリプト v2.0 クイックスタートガイド');
  console.log('==========================================================');
  console.log('');
  console.log('【ステップ1: 設定確認】');
  console.log('showCurrentProperties() を実行して設定を確認');
  console.log('');
  console.log('【ステップ2: ログレベル設定】');
  console.log('■ 本番運用の場合（最速）');
  console.log('  setLogLevel(1)');
  console.log('');
  console.log('■ 通常運用の場合（推奨）');
  console.log('  setLogLevel(2)  ← デフォルト、これを推奨');
  console.log('');
  console.log('■ トラブルシューティングの場合');
  console.log('  setLogLevel(3)');
  console.log('');
  console.log('【ステップ3: テスト実行（推奨）】');
  console.log('testBatchProcessing(10)');
  console.log('→ 10件でテスト実行して動作確認');
  console.log('');
  console.log('【ステップ4: 本番実行】');
  console.log('updateInventoryDataBatch()');
  console.log('→ 全件を一括処理で更新');
  console.log('');
  console.log('【ステップ5: 結果確認】');
  console.log('■ 正常終了の場合');
  console.log('  → ログに「✓ エラー: 0件」と表示');
  console.log('');
  console.log('■ エラーがあった場合');
  console.log('  → 「エラーログ」シートを確認');
  console.log('  → ログに詳細なエラー情報が自動出力されます');
  console.log('');
  console.log('【v2.0の主な改善点】');
  console.log('✓ 実行時間: 28-32秒 → 20-25秒（約25%高速化）');
  console.log('✓ ログ出力: 3000行 → 約20行（99%削減）');
  console.log('✓ エラー追跡: 詳細情報を自動出力');
  console.log('✓ 運用モード: 本番/開発の切り替え可能');
  console.log('');
  console.log('==========================================================');
  console.log('詳細は showUsageGuide() を実行してください');
  console.log('==========================================================');
}