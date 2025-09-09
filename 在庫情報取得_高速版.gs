/**
 * 在庫情報取得_高速版.gs
 * バッチ処理・分割実行・待機時間最適化を実装した高速版
 * 既存システムと完全独立
 */

/**
 * メイン実行関数
 * 分割実行で大量データを効率的に処理
 */
function main_高速版() {
  try {
    console.log("=== 在庫情報取得（高速版）開始 ===");
    
    // 実行開始時間を記録
    const startTime = new Date().getTime();
    const maxExecutionTime = 5.5 * 60 * 1000; // 5.5分（安全マージン）
    
    // トークン取得（独自実装）
    const tokens = getTokensDirectly();
    console.log("トークン取得成功");
    
    // スプレッドシート準備
    const spreadsheet = SpreadsheetApp.openById("1noQTPM0EMlyBNDdX4JDPZcBvh-3RT1VtWzNDA85SIkM");
    const sheet = spreadsheet.getActiveSheet();
    
    // 前回の処理継続位置を取得
    const properties = PropertiesService.getScriptProperties();
    const lastProcessedIndex = parseInt(properties.getProperty('lastProcessedIndex_高速版') || '0');
    const isFirstRun = lastProcessedIndex === 0;
    
    if (isFirstRun) {
      // 初回実行時はヘッダーを設定
      setupHeaders(sheet);
    }
    
    console.log(`前回処理位置: ${lastProcessedIndex}から開始`);
    
    // 商品リスト取得
    const products = getProductListDirect(tokens, lastProcessedIndex);
    if (products.length === 0) {
      console.log("処理対象の商品がありません。処理完了または初期化します。");
      properties.deleteProperty('lastProcessedIndex_高速版');
      return;
    }
    
    console.log(`取得した商品数: ${products.length}`);
    
    // バッチ処理で在庫情報を取得
    let processedCount = 0;
    const batchSize = 50; // バッチサイズ
    const stockDataBatch = [];
    
    for (let i = 0; i < products.length; i++) {
      // 実行時間チェック
      const currentTime = new Date().getTime();
      if (currentTime - startTime > maxExecutionTime) {
        console.log("実行時間制限に近づいたため処理を中断します");
        break;
      }
      
      const product = products[i];
      
      try {
        // 在庫情報取得
        const stockInfo = getStockInfoDirect(tokens, product.goods_id);
        if (stockInfo) {
          stockDataBatch.push({
            商品ID: product.goods_id,
            商品名: product.goods_name,
            ...stockInfo
          });
          
          console.log(`処理中: ${i + 1}/${products.length} - ${product.goods_name}`);
        }
        
        // バッチ書き込み
        if (stockDataBatch.length >= batchSize || i === products.length - 1) {
          writeBatchToSheet(sheet, stockDataBatch, isFirstRun && processedCount === 0);
          processedCount += stockDataBatch.length;
          stockDataBatch.length = 0; // 配列クリア
          
          console.log(`バッチ書き込み完了: ${processedCount}件`);
        }
        
        // API制限対策の待機（0.3秒）
        if (i < products.length - 1) {
          Utilities.sleep(300);
        }
        
      } catch (error) {
        console.error(`商品ID ${product.goods_id} の処理でエラー:`, error.message);
        // エラー時は少し長めに待機
        Utilities.sleep(500);
        continue;
      }
    }
    
    // 次回処理開始位置を保存
    const nextIndex = lastProcessedIndex + processedCount;
    if (processedCount > 0) {
      properties.setProperty('lastProcessedIndex_高速版', nextIndex.toString());
      console.log(`次回処理開始位置を保存: ${nextIndex}`);
    }
    
    console.log(`=== 処理完了: ${processedCount}件処理 ===`);
    
    // 全件処理完了判定
    if (products.length < 1000) { // 取得件数が制限値未満なら完了
      properties.deleteProperty('lastProcessedIndex_高速版');
      console.log("全件処理完了");
    }
    
  } catch (error) {
    console.error("メイン処理でエラー:", error.message);
    throw error;
  }
}

/**
 * トークンを直接取得（既存関数を使わない）
 */
function getTokensDirectly() {
  const properties = PropertiesService.getScriptProperties();
  const accessToken = properties.getProperty('ACCESS_TOKEN');
  const refreshToken = properties.getProperty('REFRESH_TOKEN');
  
  if (!accessToken || !refreshToken) {
    console.error('アクセストークンが見つかりません。');
    console.error('以下の手順で認証を完了してください:');
    console.error('1. generateAuthUrl() を実行');
    console.error('2. 出力されたURLをブラウザで開いて認証');
    throw new Error('認証が必要です。generateAuthUrl()を実行してください。');
  }
  
  return {
    accessToken,
    refreshToken
  };
}

/**
 * 商品リストを取得（直接実装）
 */
function getProductListDirect(tokens, offset = 0) {
  const url = "https://api.next-engine.org/api_v1_receiveorder/receiveOrderListForStock";
  
  const params = {
    access_token: tokens.accessToken,
    refresh_token: tokens.refreshToken,
    offset: offset,
    limit: 1000,
    fields: "goods_id,goods_name"
  };
  
  const options = {
    method: "POST",
    headers: {
      "Content-Type": "application/x-www-form-urlencoded"
    },
    payload: Object.keys(params).map(key => 
      encodeURIComponent(key) + "=" + encodeURIComponent(params[key])
    ).join("&")
  };
  
  try {
    const response = UrlFetchApp.fetch(url, options);
    const data = JSON.parse(response.getContentText());
    
    // トークン更新処理
    if (data.access_token && data.refresh_token) {
      updateTokensDirect(data.access_token, data.refresh_token);
    }
    
    if (data.result === "success" && data.data) {
      return data.data;
    } else {
      console.error("商品リスト取得エラー:", data.message || "不明なエラー");
      return [];
    }
  } catch (error) {
    console.error("商品リスト取得でネットワークエラー:", error.message);
    return [];
  }
}

/**
 * 在庫情報を取得（既存コードと同じAPI使用）
 */
function getStockInfoDirect(tokens, goodsId) {
  const url = "https://api.next-engine.org/api_v1_master_stock/search";
  
  const params = {
    access_token: tokens.accessToken,
    refresh_token: tokens.refreshToken,
    'stock_goods_id-eq': goodsId,
    fields: "stock_goods_id,stock_quantity,stock_allocation_quantity,stock_defective_quantity,stock_remaining_order_quantity,stock_out_quantity,stock_free_quantity,stock_advance_order_quantity,stock_advance_order_allocation_quantity,stock_advance_order_free_quantity"
  };
  
  const options = {
    method: "POST",
    headers: {
      "Content-Type": "application/x-www-form-urlencoded"
    },
    payload: Object.keys(params).map(key => 
      encodeURIComponent(key) + "=" + encodeURIComponent(params[key])
    ).join("&")
  };
  
  try {
    const response = UrlFetchApp.fetch(url, options);
    const data = JSON.parse(response.getContentText());
    
    // トークン更新処理
    if (data.access_token && data.refresh_token) {
      updateTokensDirect(data.access_token, data.refresh_token);
    }
    
    if (data.result === "success" && data.data && data.data.length > 0) {
      const stockData = data.data[0];
      return {
        在庫数: parseInt(stockData.stock_quantity) || 0,
        引当数: parseInt(stockData.stock_allocation_quantity) || 0,
        フリー在庫数: parseInt(stockData.stock_free_quantity) || 0,
        発注残数: parseInt(stockData.stock_remaining_order_quantity) || 0,
        欠品数: parseInt(stockData.stock_out_quantity) || 0
      };
    } else {
      console.warn(`商品ID ${goodsId} の在庫情報が見つかりません`);
      return null;
    }
  } catch (error) {
    console.error(`商品ID ${goodsId} の在庫情報取得でエラー:`, error.message);
    throw error;
  }
}

/**
 * トークンを更新保存（直接実装）
 */
function updateTokensDirect(accessToken, refreshToken) {
  const properties = PropertiesService.getScriptProperties();
  properties.setProperties({
    'ACCESS_TOKEN': accessToken,
    'REFRESH_TOKEN': refreshToken,
    'TOKEN_UPDATED_AT': new Date().getTime().toString()
  });
  console.log('トークンを更新しました');
}

/**
 * ヘッダー行を設定
 */
function setupHeaders(sheet) {
  const headers = [
    "商品ID", "商品名", "在庫数", "引当数", "フリー在庫数", "発注残数", "欠品数"
  ];
  
  sheet.clear();
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length).setFontWeight("bold");
  
  console.log("ヘッダー行を設定しました");
}

/**
 * バッチでシートに書き込み
 */
function writeBatchToSheet(sheet, stockDataBatch, isFirstWrite = false) {
  if (stockDataBatch.length === 0) return;
  
  // 書き込み用の二次元配列を作成
  const writeData = stockDataBatch.map(item => [
    item.商品ID,
    item.商品名,
    item.在庫数,
    item.引当数,
    item.フリー在庫数,
    item.発注残数,
    item.欠品数
  ]);
  
  // 書き込み開始行を決定
  const startRow = isFirstWrite ? 2 : sheet.getLastRow() + 1;
  
  // 一括書き込み
  const range = sheet.getRange(startRow, 1, writeData.length, writeData[0].length);
  range.setValues(writeData);
  
  // スプレッドシートの変更を強制実行
  SpreadsheetApp.flush();
}

/**
 * 処理状況をリセットする関数（手動実行用）
 */
function resetProgress_高速版() {
  const properties = PropertiesService.getScriptProperties();
  properties.deleteProperty('lastProcessedIndex_高速版');
  console.log("高速版の処理状況をリセットしました。次回実行時は最初から開始します。");
}

/**
 * 現在の処理状況を確認する関数（手動実行用）
 */
function checkProgress_高速版() {
  const properties = PropertiesService.getScriptProperties();
  const lastIndex = properties.getProperty('lastProcessedIndex_高速版');
  
  if (lastIndex) {
    console.log(`高速版の現在の処理位置: ${lastIndex}`);
  } else {
    console.log("高速版の処理位置は保存されていません（初回実行または完了済み）");
  }
}