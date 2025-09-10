/**
 * ネクストエンジン在庫情報取得スクリプト（一括処理版）
 * 
 * 【目的】
 * 商品コードを配列で渡し、一度のAPIコールで複数商品の在庫情報を効率的に取得
 * 
 * 【主な改善点】
 * 1. 商品マスタAPIで複数商品を一度に検索（最大100件）
 * 2. 在庫マスタAPIで複数商品の在庫を一度に取得（最大100件）
 * 3. バッチ処理による大幅な高速化
 * 4. APIコール数の削減によるレート制限回避
 * 
 * 【注意事項】
 * - 認証スクリプトで事前にトークンを取得済みである必要があります
 * - 一度に処理できる商品数は最大100件です
 * - 大量データの場合は自動的にバッチ分割します
 */

// スプレッドシートの設定（既存と同じ）
const SPREADSHEET_ID = '   ';
const SHEET_NAME = '   ';

// 列のマッピング（既存と同じ）
const COLUMNS = {
  GOODS_CODE: 0,    // A列: 商品コード
  GOODS_NAME: 1,    // B列: 商品名
  STOCK_QTY: 2,     // C列: 在庫数
  ALLOCATED_QTY: 3, // D列: 引当数
  FREE_QTY: 4,      // E列: フリー在庫数
  RESERVE_QTY: 5,   // F列: 予約在庫数
  RESERVE_ALLOCATED_QTY: 6, // G列: 予約引当数
  RESERVE_FREE_QTY: 7,      // H列: 予約フリー在庫数
  DEFECTIVE_QTY: 8,         // I列: 不良在庫数
  ORDER_REMAINING_QTY: 9,   // J列: 発注残数
  SHORTAGE_QTY: 10,         // K列: 欠品数
  JAN_CODE: 11      // L列: JANコード
};

// バッチ処理設定
const BATCH_SIZE = 100;  // 一度に処理する商品数（APIの制限に合わせて調整可能）
const API_WAIT_TIME = 500; // APIコール間の待機時間（ミリ秒）- 大幅に短縮

/**
 * メイン関数：一括処理による在庫情報更新
 */
function updateInventoryDataBatch() {
  try {
    console.log('=== 在庫情報一括更新開始 ===');
    const startTime = new Date();
    
    // スプレッドシートを取得
    const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = spreadsheet.getSheetByName(SHEET_NAME);
    
    if (!sheet) {
      throw new Error(`シート "${SHEET_NAME}" が見つかりません`);
    }
    
    // データ範囲を取得
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) {
      console.log('データが存在しません');
      return;
    }
    
    const dataRange = sheet.getRange(2, 1, lastRow - 1, 12);
    const values = dataRange.getValues();
    console.log(`処理対象: ${values.length}行`);
    
    // トークンを取得
    const tokens = getStoredTokens();
    
    // 商品コードのリストを作成（空でないもののみ）
    const goodsCodeList = [];
    const rowIndexMap = new Map(); // 商品コード → 行インデックスのマッピング
    
    for (let i = 0; i < values.length; i++) {
      const goodsCode = values[i][COLUMNS.GOODS_CODE];
      if (goodsCode && goodsCode.toString().trim()) {
        goodsCodeList.push(goodsCode.toString().trim());
        rowIndexMap.set(goodsCode.toString().trim(), i + 2); // 実際の行番号（1ベース）
      }
    }
    
    console.log(`有効な商品コード: ${goodsCodeList.length}件`);
    
    if (goodsCodeList.length === 0) {
      console.log('処理対象の商品コードがありません');
      return;
    }
    
    // バッチ処理で在庫情報を取得・更新
    let totalUpdated = 0;
    let totalErrors = 0;
    
    for (let i = 0; i < goodsCodeList.length; i += BATCH_SIZE) {
      const batch = goodsCodeList.slice(i, i + BATCH_SIZE);
      console.log(`\n--- バッチ ${Math.floor(i / BATCH_SIZE) + 1}: ${batch.length}件 ---`);
      
      try {
        // バッチで在庫情報を取得
        const inventoryDataMap = await getBatchInventoryData(batch, tokens);
        
        // スプレッドシートを更新
        for (const goodsCode of batch) {
          const inventoryData = inventoryDataMap.get(goodsCode);
          const rowIndex = rowIndexMap.get(goodsCode);
          
          if (inventoryData && rowIndex) {
            try {
              updateRowWithInventoryData(sheet, rowIndex, inventoryData);
              totalUpdated++;
              console.log(`  ✓ ${goodsCode}: 更新完了`);
            } catch (error) {
              console.error(`  ✗ ${goodsCode}: 更新エラー - ${error.message}`);
              totalErrors++;
            }
          } else {
            console.log(`  - ${goodsCode}: データなし`);
          }
        }
        
        // バッチ間の待機（APIレート制限対策）
        if (i + BATCH_SIZE < goodsCodeList.length) {
          console.log(`次のバッチまで ${API_WAIT_TIME}ms 待機...`);
          Utilities.sleep(API_WAIT_TIME);
        }
        
      } catch (error) {
        console.error(`バッチ処理エラー:`, error.message);
        totalErrors += batch.length;
      }
    }
    
    const endTime = new Date();
    const duration = (endTime - startTime) / 1000;
    
    console.log('\n=== 一括更新完了 ===');
    console.log(`処理時間: ${duration.toFixed(1)}秒`);
    console.log(`更新成功: ${totalUpdated}件`);
    console.log(`エラー: ${totalErrors}件`);
    console.log(`処理速度: ${(goodsCodeList.length / duration).toFixed(1)}件/秒`);
    
    // 従来版との比較情報を表示
    const conventionalTime = goodsCodeList.length * 2; // 従来版の推定時間（2秒/件）
    const speedImprovement = conventionalTime / duration;
    console.log(`\n--- 性能改善結果 ---`);
    console.log(`従来版推定時間: ${conventionalTime.toFixed(1)}秒`);
    console.log(`高速化倍率: ${speedImprovement.toFixed(1)}倍`);
    
  } catch (error) {
    console.error('一括更新エラー:', error.message);
    throw error;
  }
}

/**
 * バッチで在庫情報を取得
 * @param {string[]} goodsCodeList - 商品コードの配列
 * @param {Object} tokens - アクセストークンとリフレッシュトークン
 * @returns {Map<string, Object>} 商品コード → 在庫情報のマップ
 */
async function getBatchInventoryData(goodsCodeList, tokens) {
  const inventoryDataMap = new Map();
  
  try {
    console.log(`  商品マスタ一括検索: ${goodsCodeList.length}件`);
    
    // ステップ1: 商品マスタAPIで複数商品を一括検索
    const goodsDataMap = await getBatchGoodsData(goodsCodeList, tokens);
    console.log(`  商品マスタ取得完了: ${goodsDataMap.size}件`);
    
    if (goodsDataMap.size === 0) {
      console.log('  商品が見つかりませんでした');
      return inventoryDataMap;
    }
    
    // ステップ2: 在庫マスタAPIで複数商品の在庫を一括取得
    console.log(`  在庫マスタ一括検索: ${goodsDataMap.size}件`);
    const stockDataMap = await getBatchStockData(Array.from(goodsDataMap.keys()), tokens);
    console.log(`  在庫マスタ取得完了: ${stockDataMap.size}件`);
    
    // ステップ3: 商品情報と在庫情報を結合
    for (const [goodsCode, goodsData] of goodsDataMap) {
      const stockData = stockDataMap.get(goodsCode);
      
      const completeInventoryData = {
        goods_id: goodsData.goods_id,
        goods_name: goodsData.goods_name,
        stock_quantity: stockData ? parseInt(stockData.stock_quantity) || 0 : parseInt(goodsData.stock_quantity) || 0,
        stock_allocated_quantity: stockData ? parseInt(stockData.stock_allocation_quantity) || 0 : 0,
        stock_free_quantity: stockData ? parseInt(stockData.stock_free_quantity) || 0 : 0,
        stock_defective_quantity: stockData ? parseInt(stockData.stock_defective_quantity) || 0 : 0,
        stock_advance_order_quantity: stockData ? parseInt(stockData.stock_advance_order_quantity) || 0 : 0,
        stock_advance_order_allocation_quantity: stockData ? parseInt(stockData.stock_advance_order_allocation_quantity) || 0 : 0,
        stock_advance_order_free_quantity: stockData ? parseInt(stockData.stock_advance_order_free_quantity) || 0 : 0,
        stock_remaining_order_quantity: stockData ? parseInt(stockData.stock_remaining_order_quantity) || 0 : 0,
        stock_out_quantity: stockData ? parseInt(stockData.stock_out_quantity) || 0 : 0
      };
      
      inventoryDataMap.set(goodsCode, completeInventoryData);
    }
    
    console.log(`  結合完了: ${inventoryDataMap.size}件`);
    return inventoryDataMap;
    
  } catch (error) {
    console.error(`バッチ在庫取得エラー:`, error.message);
    return inventoryDataMap;
  }
}

/**
 * 複数商品の基本情報を一括取得
 * @param {string[]} goodsCodeList - 商品コードの配列
 * @param {Object} tokens - トークン情報
 * @returns {Map<string, Object>} 商品コード → 商品情報のマップ
 */
async function getBatchGoodsData(goodsCodeList, tokens) {
  const url = `${NE_API_URL}/api_v1_master_goods/search`;
  
  // 複数の商品IDを検索条件に設定
  const goodsIdCondition = goodsCodeList.join(',');
  
  const payload = {
    'access_token': tokens.accessToken,
    'refresh_token': tokens.refreshToken,
    'goods_id-in': goodsIdCondition, // IN条件で複数商品を一括検索
    'fields': 'goods_id,goods_name,stock_quantity',
    'limit': BATCH_SIZE.toString() // 取得件数制限
  };
  
  const options = {
    'method': 'POST',
    'headers': {
      'Content-Type': 'application/x-www-form-urlencoded'
    },
    'payload': Object.keys(payload).map(key =>
      encodeURIComponent(key) + '=' + encodeURIComponent(payload[key])
    ).join('&')
  };
  
  const goodsDataMap = new Map();
  
  try {
    const response = UrlFetchApp.fetch(url, options);
    const responseText = response.getContentText();
    const responseData = JSON.parse(responseText);
    
    // トークンが更新された場合は保存
    if (responseData.access_token && responseData.refresh_token) {
      updateStoredTokens(responseData.access_token, responseData.refresh_token);
      // トークンを更新
      tokens.accessToken = responseData.access_token;
      tokens.refreshToken = responseData.refresh_token;
    }
    
    if (responseData.result === 'success' && responseData.data) {
      responseData.data.forEach(goodsData => {
        goodsDataMap.set(goodsData.goods_id, {
          goods_id: goodsData.goods_id,
          goods_name: goodsData.goods_name,
          stock_quantity: goodsData.stock_quantity
        });
      });
      
      console.log(`    API応答: ${responseData.data.length}件取得`);
    } else {
      console.error(`    商品マスタAPI エラー:`, responseData.message || 'Unknown error');
    }
    
    return goodsDataMap;
    
  } catch (error) {
    console.error(`商品マスタ一括取得エラー:`, error.message);
    return goodsDataMap;
  }
}

/**
 * 複数商品の在庫情報を一括取得
 * @param {string[]} goodsCodeList - 商品コードの配列
 * @param {Object} tokens - トークン情報
 * @returns {Map<string, Object>} 商品コード → 在庫情報のマップ
 */
async function getBatchStockData(goodsCodeList, tokens) {
  const url = `${NE_API_URL}/api_v1_master_stock/search`;
  
  // 複数の商品IDを検索条件に設定
  const goodsIdCondition = goodsCodeList.join(',');
  
  const payload = {
    'access_token': tokens.accessToken,
    'refresh_token': tokens.refreshToken,
    'stock_goods_id-in': goodsIdCondition, // IN条件で複数商品の在庫を一括検索
    'fields': 'stock_goods_id,stock_quantity,stock_allocation_quantity,stock_defective_quantity,stock_remaining_order_quantity,stock_out_quantity,stock_free_quantity,stock_advance_order_quantity,stock_advance_order_allocation_quantity,stock_advance_order_free_quantity',
    'limit': BATCH_SIZE.toString()
  };
  
  const options = {
    'method': 'POST',
    'headers': {
      'Content-Type': 'application/x-www-form-urlencoded'
    },
    'payload': Object.keys(payload).map(key =>
      encodeURIComponent(key) + '=' + encodeURIComponent(payload[key])
    ).join('&')
  };
  
  const stockDataMap = new Map();
  
  try {
    const response = UrlFetchApp.fetch(url, options);
    const responseText = response.getContentText();
    const responseData = JSON.parse(responseText);
    
    // トークンが更新された場合は保存
    if (responseData.access_token && responseData.refresh_token) {
      updateStoredTokens(responseData.access_token, responseData.refresh_token);
      // トークンを更新
      tokens.accessToken = responseData.access_token;
      tokens.refreshToken = responseData.refresh_token;
    }
    
    if (responseData.result === 'success' && responseData.data) {
      responseData.data.forEach(stockData => {
        stockDataMap.set(stockData.stock_goods_id, stockData);
      });
      
      console.log(`    API応答: ${responseData.data.length}件取得`);
    } else {
      console.error(`    在庫マスタAPI エラー:`, responseData.message || 'Unknown error');
    }
    
    return stockDataMap;
    
  } catch (error) {
    console.error(`在庫マスタ一括取得エラー:`, error.message);
    return stockDataMap;
  }
}

/**
 * 保存されたトークンを取得（既存関数）
 */
function getStoredTokens() {
  const properties = PropertiesService.getScriptProperties();
  const accessToken = properties.getProperty('ACCESS_TOKEN');
  const refreshToken = properties.getProperty('REFRESH_TOKEN');
  
  if (!accessToken || !refreshToken) {
    throw new Error('アクセストークンが見つかりません。先に認証を完了してください。');
  }
  
  return {
    accessToken,
    refreshToken
  };
}

/**
 * スプレッドシートの行を在庫データで更新（既存関数）
 */
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

/**
 * トークンを更新保存（既存関数）
 */
function updateStoredTokens(accessToken, refreshToken) {
  const properties = PropertiesService.getScriptProperties();
  properties.setProperties({
    'ACCESS_TOKEN': accessToken,
    'REFRESH_TOKEN': refreshToken,
    'TOKEN_UPDATED_AT': new Date().getTime().toString()
  });
  console.log('  トークンを更新しました');
}

/**
 * テスト用：小規模バッチでの動作確認
 * @param {number} maxItems - テスト対象の最大商品数（デフォルト: 10）
 */
function testBatchProcessing(maxItems = 10) {
  try {
    console.log(`=== バッチ処理テスト（最大${maxItems}件） ===`);
    
    // スプレッドシートから商品コードを取得
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
    
    // バッチで在庫情報を取得
    const startTime = new Date();
    const inventoryDataMap = getBatchInventoryData(goodsCodeList, tokens);
    const endTime = new Date();
    const duration = (endTime - startTime) / 1000;
    
    console.log(`\n=== テスト結果 ===`);
    console.log(`処理時間: ${duration.toFixed(1)}秒`);
    console.log(`取得件数: ${inventoryDataMap.size}件`);
    console.log(`処理速度: ${(goodsCodeList.length / duration).toFixed(1)}件/秒`);
    
    // 取得したデータの内容を表示
    for (const [goodsCode, data] of inventoryDataMap) {
      console.log(`${goodsCode}: 在庫${data.stock_quantity} 引当${data.stock_allocated_quantity} フリー${data.stock_free_quantity}`);
    }
    
  } catch (error) {
    console.error('バッチテストエラー:', error.message);
    throw error;
  }
}

/**
 * パフォーマンス比較用：従来版と一括版の処理時間を比較
 * @param {number} sampleSize - 比較対象のサンプル数（デフォルト: 10）
 */
function comparePerformance(sampleSize = 10) {
  console.log(`=== パフォーマンス比較テスト（${sampleSize}件） ===`);
  
  // スプレッドシートから商品コードを取得
  const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = spreadsheet.getSheetByName(SHEET_NAME);
  const lastRow = sheet.getLastRow();
  
  if (lastRow <= 1) {
    console.log('テスト用データが存在しません');
    return;
  }
  
  const dataRange = sheet.getRange(2, 1, Math.min(sampleSize, lastRow - 1), 1);
  const values = dataRange.getValues();
  const goodsCodeList = values
    .map(row => row[0])
    .filter(code => code && code.toString().trim())
    .slice(0, sampleSize);
  
  console.log(`比較対象商品コード: ${goodsCodeList.join(', ')}`);
  
  const tokens = getStoredTokens();
  
  // 従来版の推定時間（実際には実行しない）
  const conventionalEstimatedTime = goodsCodeList.length * 2; // 2秒/件
  
  // 一括版の実際の処理時間
  console.log('\n一括版実行中...');
  const startTime = new Date();
  const inventoryDataMap = getBatchInventoryData(goodsCodeList, tokens);
  const endTime = new Date();
  const batchTime = (endTime - startTime) / 1000;
  
  // 結果比較
  const speedImprovement = conventionalEstimatedTime / batchTime;
  
  console.log('\n=== 性能比較結果 ===');
  console.log(`従来版推定時間: ${conventionalEstimatedTime.toFixed(1)}秒（${sampleSize} × 2秒/件）`);
  console.log(`一括版実際時間: ${batchTime.toFixed(1)}秒`);
  console.log(`高速化倍率: ${speedImprovement.toFixed(1)}倍`);
  console.log(`取得成功率: ${(inventoryDataMap.size / goodsCodeList.length * 100).toFixed(1)}%`);
  
  // 数千件での推定効果
  const estimatedFor1000 = {
    conventional: 1000 * 2 / 60, // 分
    batch: 1000 / goodsCodeList.length * batchTime / 60 // 分
  };
  
  console.log('\n=== 1000件処理時の推定時間 ===');
  console.log(`従来版: ${estimatedFor1000.conventional.toFixed(1)}分`);
  console.log(`一括版: ${estimatedFor1000.batch.toFixed(1)}分`);
  console.log(`時間短縮: ${(estimatedFor1000.conventional - estimatedFor1000.batch).toFixed(1)}分`);
}

/**
 * 使用方法ガイド
 */
function showBatchUsageGuide() {
  console.log('=== 一括処理版 使用方法ガイド ===');
  console.log('');
  console.log('【主要関数】');
  console.log('1. updateInventoryDataBatch()');
  console.log('   - 全商品の在庫情報を一括処理で更新');
  console.log('   - 100件ずつのバッチで自動分割処理');
  console.log('   - 従来版より大幅に高速化');
  console.log('');
  console.log('2. testBatchProcessing(件数)');
  console.log('   - 小規模テスト用（デフォルト10件）');
  console.log('   - 例: testBatchProcessing(5)');
  console.log('');
  console.log('3. comparePerformance(件数)');
  console.log('   - 従来版との性能比較テスト');
  console.log('   - 例: comparePerformance(20)');
  console.log('');
  console.log('【期待される改善効果】');
  console.log('- APIコール数: 1/50～1/100に削減');
  console.log('- 処理速度: 10～50倍の高速化');
  console.log('- 実行時間制限: 数千件でも制限内で完了');
  console.log('- API制限: レート制限に引っかかりにくい');
  console.log('');
  console.log('【推奨実行手順】');
  console.log('1. testBatchProcessing(10) で動作確認');
  console.log('2. comparePerformance(20) で性能確認');
  console.log('3. updateInventoryDataBatch() で全件更新');
  console.log('');
  console.log('【設定変更可能項目】');
  console.log('- BATCH_SIZE: バッチサイズ（現在100件）');
  console.log('- API_WAIT_TIME: API間隔（現在500ms）');
}