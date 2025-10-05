/**
 * 在庫情報取得_改良版.gs
 * 既存システムと完全同等のAPI使用 + 高速化機能
 */

// スプレッドシートの設定（既存コードと同じ）
const SPREADSHEET_ID_改良版 = '1noQTPM0EMlyBNDdX4JDPZcBvh-3RT1VtWzNDA85SIkM';

/**
 * メイン実行関数（改良版）
 */
function main_改良版() {
  try {
    console.log("=== 在庫情報取得（改良版）開始 ===");
    
    // 実行開始時間を記録
    const startTime = new Date().getTime();
    const maxExecutionTime = 5.5 * 60 * 1000; // 5.5分
    
    // スプレッドシート準備
    const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID_改良版);
    const sheet = spreadsheet.getActiveSheet();
    
    // 前回の処理継続位置を取得
    const properties = PropertiesService.getScriptProperties();
    const lastProcessedRow = parseInt(properties.getProperty('lastProcessedRow_改良版') || '2');
    
    // データ範囲を取得
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) {
      console.log('データが存在しません');
      return;
    }
    
    const dataRange = sheet.getRange(lastProcessedRow, 1, lastRow - lastProcessedRow + 1, 12);
    const values = dataRange.getValues();
    
    console.log(`処理対象: ${values.length}行（${lastProcessedRow}行目から）`);
    
    if (values.length === 0) {
      console.log("処理対象の商品がありません。処理完了。");
      properties.deleteProperty('lastProcessedRow_改良版');
      return;
    }
    
    // トークン取得（既存方式）
    const tokens = getStoredTokens_改良版();
    console.log("トークン取得成功");
    
    // バッチ処理で在庫情報を取得・更新
    let processedCount = 0;
    const batchSize = 20; // バッチサイズ（書き込み頻度調整）
    const updateBatch = [];
    
    for (let i = 0; i < values.length; i++) {
      // 実行時間チェック
      const currentTime = new Date().getTime();
      if (currentTime - startTime > maxExecutionTime) {
        console.log("実行時間制限に近づいたため処理を中断します");
        break;
      }
      
      const row = values[i];
      const goodsCode = row[0]; // A列: 商品コード
      const actualRowNumber = lastProcessedRow + i;
      
      if (!goodsCode) {
        console.log(`${actualRowNumber}行目: 商品コードが空のためスキップ`);
        continue;
      }
      
      try {
        console.log(`${actualRowNumber}行目: ${goodsCode} の在庫情報を取得中...`);
        
        // 既存システムと同じ方法で在庫情報を取得
        const inventoryData = getInventoryByGoodsCode_改良版(goodsCode, tokens);
        
        if (inventoryData) {
          updateBatch.push({
            rowNumber: actualRowNumber,
            data: inventoryData
          });
          
          console.log(`${actualRowNumber}行目: ${goodsCode} 取得完了`);
        } else {
          console.log(`${actualRowNumber}行目: ${goodsCode} の在庫情報が見つかりません`);
        }
        
        // バッチ更新
        if (updateBatch.length >= batchSize || i === values.length - 1) {
          updateRowsBatch_改良版(sheet, updateBatch);
          processedCount += updateBatch.length;
          updateBatch.length = 0; // 配列クリア
          
          console.log(`バッチ更新完了: ${processedCount}件`);
        }
        
        // API制限対策の待機（0.3秒）
        if (i < values.length - 1) {
          Utilities.sleep(300);
        }
        
      } catch (error) {
        console.error(`${actualRowNumber}行目: ${goodsCode} のエラー:`, error.message);
        // エラー時は少し長めに待機
        Utilities.sleep(500);
        continue;
      }
    }
    
    // 次回処理開始位置を保存
    const nextRow = lastProcessedRow + processedCount;
    if (processedCount > 0) {
      if (nextRow <= lastRow) {
        properties.setProperty('lastProcessedRow_改良版', nextRow.toString());
        console.log(`次回処理開始行を保存: ${nextRow}行目`);
      } else {
        properties.deleteProperty('lastProcessedRow_改良版');
        console.log("全件処理完了");
      }
    }
    
    console.log(`=== 処理完了: ${processedCount}件処理 ===`);
    
  } catch (error) {
    console.error("メイン処理でエラー:", error.message);
    throw error;
  }
}

/**
 * 保存されたトークンを取得（既存コードと同じ）
 */
function getStoredTokens_改良版() {
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
 * 商品コードから完全な在庫情報を取得（既存コードベース）
 */
function getInventoryByGoodsCode_改良版(goodsCode, tokens) {
  try {
    // ステップ1: 商品マスタAPIで基本情報を取得
    const goodsData = searchGoodsWithStock_改良版(goodsCode, tokens);
    if (!goodsData) {
      return null;
    }
    
    // ステップ2: 在庫マスタAPIで詳細在庫情報を取得
    const stockDetails = getStockByGoodsId_改良版(goodsCode, tokens);
    
    let completeInventoryData;
    if (stockDetails) {
      // 商品情報と詳細在庫情報を結合
      completeInventoryData = {
        goods_id: goodsData.goods_id,
        goods_name: goodsData.goods_name,
        stock_quantity: parseInt(stockDetails.stock_quantity) || parseInt(goodsData.stock_quantity) || 0,
        stock_allocated_quantity: parseInt(stockDetails.stock_allocation_quantity) || 0,
        stock_free_quantity: parseInt(stockDetails.stock_free_quantity) || 0,
        stock_defective_quantity: parseInt(stockDetails.stock_defective_quantity) || 0,
        stock_advance_order_quantity: parseInt(stockDetails.stock_advance_order_quantity) || 0,
        stock_advance_order_allocation_quantity: parseInt(stockDetails.stock_advance_order_allocation_quantity) || 0,
        stock_advance_order_free_quantity: parseInt(stockDetails.stock_advance_order_free_quantity) || 0,
        stock_remaining_order_quantity: parseInt(stockDetails.stock_remaining_order_quantity) || 0,
        stock_out_quantity: parseInt(stockDetails.stock_out_quantity) || 0
      };
    } else {
      // 詳細情報が取得できない場合は基本情報のみ使用
      completeInventoryData = {
        goods_id: goodsData.goods_id,
        goods_name: goodsData.goods_name,
        stock_quantity: parseInt(goodsData.stock_quantity) || 0,
        stock_allocated_quantity: 0,
        stock_free_quantity: 0,
        stock_defective_quantity: 0,
        stock_advance_order_quantity: 0,
        stock_advance_order_allocation_quantity: 0,
        stock_advance_order_free_quantity: 0,
        stock_remaining_order_quantity: 0,
        stock_out_quantity: 0
      };
    }
    
    return completeInventoryData;
    
  } catch (error) {
    console.error(`商品コード ${goodsCode} の在庫取得エラー:`, error.message);
    return null;
  }
}

/**
 * 商品コードで商品マスタを検索（既存コードと同じ）
 */
function searchGoodsWithStock_改良版(goodsCode, tokens) {
  const url = 'https://api.next-engine.org/api_v1_master_goods/search';
  
  const payload = {
    'access_token': tokens.accessToken,
    'refresh_token': tokens.refreshToken,
    'goods_id-eq': goodsCode,
    'fields': 'goods_id,goods_name,stock_quantity'
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
  
  try {
    const response = UrlFetchApp.fetch(url, options);
    const responseText = response.getContentText();
    const responseData = JSON.parse(responseText);
    
    // トークンが更新された場合は保存
    if (responseData.access_token && responseData.refresh_token) {
      updateStoredTokens_改良版(responseData.access_token, responseData.refresh_token);
    }
    
    if (responseData.result === 'success') {
      if (responseData.data && responseData.data.length > 0) {
        const goodsData = responseData.data[0];
        return {
          goods_id: goodsData.goods_id,
          goods_name: goodsData.goods_name,
          stock_quantity: goodsData.stock_quantity
        };
      } else {
        return null;
      }
    } else {
      console.error(`商品検索エラー:`, responseData.message);
      return null;
    }
  } catch (error) {
    console.error('商品マスタAPI呼び出しエラー:', error.toString());
    return null;
  }
}

/**
 * 商品IDから詳細在庫情報を取得（既存コードと同じ）
 */
function getStockByGoodsId_改良版(goodsId, tokens) {
  const url = 'https://api.next-engine.org/api_v1_master_stock/search';
  
  const payload = {
    'access_token': tokens.accessToken,
    'refresh_token': tokens.refreshToken,
    'stock_goods_id-eq': goodsId,
    'fields': 'stock_goods_id,stock_quantity,stock_allocation_quantity,stock_defective_quantity,stock_remaining_order_quantity,stock_out_quantity,stock_free_quantity,stock_advance_order_quantity,stock_advance_order_allocation_quantity,stock_advance_order_free_quantity'
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
  
  try {
    const response = UrlFetchApp.fetch(url, options);
    const responseText = response.getContentText();
    const responseData = JSON.parse(responseText);
    
    // トークンが更新された場合は保存
    if (responseData.access_token && responseData.refresh_token) {
      updateStoredTokens_改良版(responseData.access_token, responseData.refresh_token);
    }
    
    if (responseData.result === 'success' && responseData.data && responseData.data.length > 0) {
      return responseData.data[0];
    } else {
      return null;
    }
  } catch (error) {
    console.error('在庫マスタAPI呼び出しエラー:', error.toString());
    return null;
  }
}

/**
 * 複数行をバッチで更新
 */
function updateRowsBatch_改良版(sheet, updateBatch) {
  for (const update of updateBatch) {
    updateRowWithInventoryData_改良版(sheet, update.rowNumber, update.data);
  }
  
  // スプレッドシートの変更を強制実行
  SpreadsheetApp.flush();
}

/**
 * スプレッドシートの行を在庫データで更新（既存コードと同じ）
 */
function updateRowWithInventoryData_改良版(sheet, rowIndex, inventoryData) {
  // 在庫情報の列を更新（C列からK列まで）
  const updateValues = [
    inventoryData.stock_quantity || 0, // C列: 在庫数
    inventoryData.stock_allocated_quantity || 0, // D列: 引当数
    inventoryData.stock_free_quantity || 0, // E列: フリー在庫数
    inventoryData.stock_advance_order_quantity || 0, // F列: 予約在庫数
    inventoryData.stock_advance_order_allocation_quantity || 0, // G列: 予約引当数
    inventoryData.stock_advance_order_free_quantity || 0, // H列: 予約フリー在庫数
    inventoryData.stock_defective_quantity || 0, // I列: 不良在庫数
    inventoryData.stock_remaining_order_quantity || 0, // J列: 発注残数
    inventoryData.stock_out_quantity || 0 // K列: 欠品数
  ];
  
  // C列からK列まで更新
  const range = sheet.getRange(rowIndex, 3, 1, updateValues.length);
  range.setValues([updateValues]);
}

/**
 * トークンを更新保存（既存コードと同じ）
 */
function updateStoredTokens_改良版(accessToken, refreshToken) {
  const properties = PropertiesService.getScriptProperties();
  properties.setProperties({
    'ACCESS_TOKEN': accessToken,
    'REFRESH_TOKEN': refreshToken,
    'TOKEN_UPDATED_AT': new Date().getTime().toString()
  });
}

/**
 * 処理状況をリセット
 */
function resetProgress_改良版() {
  const properties = PropertiesService.getScriptProperties();
  properties.deleteProperty('lastProcessedRow_改良版');
  console.log("改良版の処理状況をリセットしました。");
}

/**
 * 現在の処理状況を確認
 */
function checkProgress_改良版() {
  const properties = PropertiesService.getScriptProperties();
  const lastRow = properties.getProperty('lastProcessedRow_改良版');
  
  if (lastRow) {
    console.log(`改良版の現在の処理位置: ${lastRow}行目から開始`);
  } else {
    console.log("改良版の処理位置は保存されていません（初回実行または完了済み）");
  }
}