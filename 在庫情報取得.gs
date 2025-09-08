/**
 * ネクストエンジン在庫情報取得スクリプト（完全版）
 * 
 * 【目的】
 * スプレッドシートの商品コードに対応する詳細在庫情報をネクストエンジンAPIから取得し更新
 * 
 * 【機能】
 * 1. スプレッドシートから商品コードを読み取り
 * 2. ネクストエンジン商品マスタAPIで基本情報を取得
 * 3. ネクストエンジン在庫マスタAPIで詳細在庫情報を取得
 * 4. スプレッドシートの在庫情報を詳細データで更新
 * 
 * 【注意事項】
 * - 認証スクリプトで事前にトークンを取得済みである必要があります
 * - API制限を考慮して適切な間隔でリクエストを送信します
 * - 大量データの場合は時間がかかる可能性があります
 */

// ネクストエンジンAPIのエンドポイントは認証.gsで定義済み

// スプレッドシートの設定
const SPREADSHEET_ID = '    ';
const SHEET_NAME = 'GAS'; // 必要に応じて変更してください

// 列のマッピング（0ベース）
const COLUMNS = {
  GOODS_CODE: 0,     // A列: 商品コード
  GOODS_NAME: 1,     // B列: 商品名  
  STOCK_QTY: 2,      // C列: 在庫数
  ALLOCATED_QTY: 3,  // D列: 引当数
  FREE_QTY: 4,       // E列: フリー在庫数
  RESERVE_QTY: 5,    // F列: 予約在庫数
  RESERVE_ALLOCATED_QTY: 6, // G列: 予約引当数
  RESERVE_FREE_QTY: 7,      // H列: 予約フリー在庫数
  DEFECTIVE_QTY: 8,  // I列: 不良在庫数
  ORDER_REMAINING_QTY: 9,   // J列: 発注残数
  SHORTAGE_QTY: 10,  // K列: 欠品数
  JAN_CODE: 11       // L列: JANコード
};

/**
 * メイン関数：在庫情報を更新
 */
function updateInventoryData() {
  try {
    console.log('=== 在庫情報更新開始 ===');
    
    // スプレッドシートを取得
    const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = spreadsheet.getSheetByName(SHEET_NAME);
    
    if (!sheet) {
      throw new Error(`シート "${SHEET_NAME}" が見つかりません`);
    }
    
    // データ範囲を取得（ヘッダー行を除く）
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) {
      console.log('データが存在しません');
      return;
    }
    
    const dataRange = sheet.getRange(2, 1, lastRow - 1, 12); // 2行目から最終行まで、A列からL列まで
    const values = dataRange.getValues();
    
    console.log(`処理対象: ${values.length}行`);
    
    // トークンを取得
    const tokens = getStoredTokens();
    
    // 各行の在庫情報を更新
    let updateCount = 0;
    let errorCount = 0;
    
    for (let i = 0; i < values.length; i++) {
      const row = values[i];
      const goodsCode = row[COLUMNS.GOODS_CODE];
      
      if (!goodsCode) {
        console.log(`${i + 2}行目: 商品コードが空のためスキップ`);
        continue;
      }
      
      try {
        console.log(`${i + 2}行目: ${goodsCode} の在庫情報を取得中...`);
        
        // 在庫情報を取得
        const inventoryData = getInventoryByGoodsCode(goodsCode, tokens);
        
        if (inventoryData) {
          // スプレッドシートの行を更新
          updateRowWithInventoryData(sheet, i + 2, inventoryData);
          updateCount++;
          console.log(`${i + 2}行目: ${goodsCode} 更新完了`);
        } else {
          console.log(`${i + 2}行目: ${goodsCode} の在庫情報が見つかりません`);
        }
        
        // API制限を考慮して少し待機（2秒）- 2つのAPIを呼び出すため少し長めに設定
        Utilities.sleep(2000);
        
      } catch (error) {
        console.error(`${i + 2}行目: ${goodsCode} のエラー:`, error.message);
        errorCount++;
      }
    }
    
    console.log('=== 在庫情報更新完了 ===');
    console.log(`更新成功: ${updateCount}件`);
    console.log(`エラー: ${errorCount}件`);
    
  } catch (error) {
    console.error('在庫情報更新エラー:', error.message);
    throw error;
  }
}

/**
 * 保存されたトークンを取得
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
 * 商品コードから完全な在庫情報を取得（完全版）
 * @param {string} goodsCode - 商品コード
 * @param {Object} tokens - アクセストークンとリフレッシュトークン
 * @returns {Object|null} 完全な在庫情報
 */
function getInventoryByGoodsCode(goodsCode, tokens) {
  try {
    console.log('在庫情報取得開始...');
    console.log(`対象商品コード: ${goodsCode}`);
    
    // ステップ1: 商品マスタAPIで基本情報を取得
    console.log('商品マスタAPI呼び出し中...');
    const goodsData = searchGoodsWithStock(goodsCode, tokens);
    
    if (!goodsData) {
      console.log('商品が見つかりませんでした');
      return null;
    }
    
    console.log('商品基本情報取得完了');
    console.log(`商品名: ${goodsData.goods_name}`);
    console.log(`基本在庫数: ${goodsData.stock_quantity}`);
    
    // ステップ2: 在庫マスタAPIで詳細在庫情報を取得
    console.log('在庫マスタAPI呼び出し中...');
    const stockDetails = getStockByGoodsId(goodsCode, tokens);
    
    let completeInventoryData;
    
    if (stockDetails) {
      console.log('詳細在庫情報取得完了');
      
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
      
      console.log('完全な在庫情報を構築しました:');
      console.log(`- 在庫数: ${completeInventoryData.stock_quantity}`);
      console.log(`- 引当数: ${completeInventoryData.stock_allocated_quantity}`);
      console.log(`- フリー在庫数: ${completeInventoryData.stock_free_quantity}`);
      console.log(`- 不良在庫数: ${completeInventoryData.stock_defective_quantity}`);
      console.log(`- 予約在庫数: ${completeInventoryData.stock_advance_order_quantity}`);
      console.log(`- 予約引当数: ${completeInventoryData.stock_advance_order_allocation_quantity}`);
      console.log(`- 予約フリー在庫数: ${completeInventoryData.stock_advance_order_free_quantity}`);
      console.log(`- 発注残数: ${completeInventoryData.stock_remaining_order_quantity}`);
      console.log(`- 欠品数: ${completeInventoryData.stock_out_quantity}`);
      
    } else {
      console.log('詳細在庫情報が取得できませんでした。基本情報のみで構築します。');
      
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
    
    console.log('在庫情報取得完了');
    return completeInventoryData;
    
  } catch (error) {
    console.error(`商品コード ${goodsCode} の在庫取得エラー:`, error.message);
    console.error('エラー詳細:', error.stack);
    return null;
  }
}

/**
 * 商品コードで商品マスタを検索し在庫情報も取得（修正版）
 * @param {string} goodsCode - 商品コード
 * @param {Object} tokens - トークン情報
 * @returns {Object|null} 商品情報と在庫情報
 */
function searchGoodsWithStock(goodsCode, tokens) {
  const url = `${NE_API_URL}/api_v1_master_goods/search`;
  
  const payload = {
    'access_token': tokens.accessToken,
    'refresh_token': tokens.refreshToken,
    'goods_id-eq': goodsCode, // goods_idで検索
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
    
    console.log('商品マスタAPI応答:', responseData.result);
    
    // トークンが更新された場合は保存
    if (responseData.access_token && responseData.refresh_token) {
      updateStoredTokens(responseData.access_token, responseData.refresh_token);
    }
    
    if (responseData.result === 'success') {
      if (responseData.data && responseData.data.length > 0) {
        const goodsData = responseData.data[0];
        console.log('取得した商品データ:', goodsData);
        
        return {
          goods_id: goodsData.goods_id,
          goods_name: goodsData.goods_name,
          stock_quantity: goodsData.stock_quantity
        };
      } else {
        console.log(`商品コード ${goodsCode} が見つかりません`);
        return null;
      }
    } else {
      console.error(`商品検索エラー:`, JSON.stringify(responseData));
      if (responseData.message) {
        console.error('エラーメッセージ:', responseData.message);
      }
      return null;
    }
  } catch (error) {
    console.error('商品マスタAPI呼び出しエラー:', error.toString());
    return null;
  }
}

/**
 * 商品IDから詳細在庫情報を取得（完全版）
 * @param {string} goodsId - 商品ID
 * @param {Object} tokens - トークン情報
 * @returns {Object|null} 詳細在庫情報
 */
function getStockByGoodsId(goodsId, tokens) {
  const url = `${NE_API_URL}/api_v1_master_stock/search`;
  
  const payload = {
    'access_token': tokens.accessToken,
    'refresh_token': tokens.refreshToken,
    'stock_goods_id-eq': goodsId, // 正しいフィールド名
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
    
    console.log('在庫マスタAPI応答:', responseData.result);
    
    // トークンが更新された場合は保存
    if (responseData.access_token && responseData.refresh_token) {
      updateStoredTokens(responseData.access_token, responseData.refresh_token);
    }
    
    if (responseData.result === 'success' && responseData.data && responseData.data.length > 0) {
      const stockData = responseData.data[0];
      console.log('取得した詳細在庫データ:', stockData);
      
      // 在庫マスタAPIから取得したデータをそのまま返す
      return stockData;
      
    } else {
      console.log(`商品ID ${goodsId} の在庫情報が見つかりません`);
      console.log('API応答詳細:', JSON.stringify(responseData, null, 2));
      return null;
    }
  } catch (error) {
    console.error('在庫マスタAPI呼び出しエラー:', error.toString());
    return null;
  }
}

/**
 * スプレッドシートの行を在庫データで更新（完全版）
 * @param {Sheet} sheet - シートオブジェクト
 * @param {number} rowIndex - 更新する行番号（1ベース）
 * @param {Object} inventoryData - 在庫データ
 */
function updateRowWithInventoryData(sheet, rowIndex, inventoryData) {
  // 在庫情報の列を更新（C列からK列まで）
  const updateValues = [
    inventoryData.stock_quantity || 0,                            // C列: 在庫数
    inventoryData.stock_allocated_quantity || 0,                  // D列: 引当数  
    inventoryData.stock_free_quantity || 0,                       // E列: フリー在庫数
    inventoryData.stock_advance_order_quantity || 0,              // F列: 予約在庫数
    inventoryData.stock_advance_order_allocation_quantity || 0,   // G列: 予約引当数
    inventoryData.stock_advance_order_free_quantity || 0,         // H列: 予約フリー在庫数
    inventoryData.stock_defective_quantity || 0,                  // I列: 不良在庫数
    inventoryData.stock_remaining_order_quantity || 0,            // J列: 発注残数
    inventoryData.stock_out_quantity || 0                         // K列: 欠品数
  ];
  
  // C列からK列まで更新
  const range = sheet.getRange(rowIndex, COLUMNS.STOCK_QTY + 1, 1, updateValues.length);
  range.setValues([updateValues]);
  
  console.log('更新データ:', updateValues);
  console.log('在庫管理スクリプト: データ取得の不具合で最新版を実行したログとなります。');
  console.log('');
  console.log(`取得した在庫データ: { goods_id: '${inventoryData.goods_id}', goods_name: '${inventoryData.goods_name}', stock_quantity: ${inventoryData.stock_quantity}, stock_allocated_quantity: ${inventoryData.stock_allocated_quantity}, stock_free_quantity: ${inventoryData.stock_free_quantity}, stock_defective_quantity: ${inventoryData.stock_defective_quantity}, stock_advance_order_quantity: ${inventoryData.stock_advance_order_quantity}, stock_advance_order_allocation_quantity: ${inventoryData.stock_advance_order_allocation_quantity}, stock_advance_order_free_quantity: ${inventoryData.stock_advance_order_free_quantity}, stock_remaining_order_quantity: ${inventoryData.stock_remaining_order_quantity}, stock_out_quantity: ${inventoryData.stock_out_quantity} }`);
}

/**
 * トークンを更新保存
 * @param {string} accessToken - 新しいアクセストークン
 * @param {string} refreshToken - 新しいリフレッシュトークン
 */
function updateStoredTokens(accessToken, refreshToken) {
  const properties = PropertiesService.getScriptProperties();
  properties.setProperties({
    'ACCESS_TOKEN': accessToken,
    'REFRESH_TOKEN': refreshToken,
    'TOKEN_UPDATED_AT': new Date().getTime().toString()
  });
  console.log('トークンを更新しました');
}

/**
 * 特定の商品コードのみ更新（テスト用・完全版）
 * @param {string} goodsCode - 更新したい商品コード
 */
function updateSingleProduct(goodsCode) {
  try {
    console.log('=== バージョン確認: 完全版が実行されています ===');
    console.log(`=== 単品更新開始: ${goodsCode} ===`);
    
    const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
    console.log('スプレッドシート取得成功');
    
    const sheet = spreadsheet.getSheetByName(SHEET_NAME);
    console.log('シート取得結果:', sheet ? 'success' : 'null');
    
    if (!sheet) {
      throw new Error(`シート "${SHEET_NAME}" が見つかりません`);
    }
    
    // 商品コードを検索
    const dataRange = sheet.getRange(2, 1, sheet.getLastRow() - 1, 12);
    console.log('データ範囲取得成功');
    
    const values = dataRange.getValues();
    console.log('データ取得成功、行数:', values.length);
    
    let targetRowIndex = -1;
    for (let i = 0; i < values.length; i++) {
      if (values[i][COLUMNS.GOODS_CODE] === goodsCode) {
        targetRowIndex = i + 2; // 実際の行番号（1ベース）
        break;
      }
    }
    
    if (targetRowIndex === -1) {
      throw new Error(`商品コード ${goodsCode} がスプレッドシートに見つかりません`);
    }
    
    console.log('商品コード発見、行番号:', targetRowIndex);
    
    // トークンを取得
    const tokens = getStoredTokens();
    console.log('トークン取得成功');
    
    // 在庫情報を取得
    console.log('getInventoryByGoodsCode 関数を呼び出します');
    const inventoryData = getInventoryByGoodsCode(goodsCode, tokens);
    console.log('getInventoryByGoodsCode 完了、結果:', inventoryData ? 'データあり' : 'データなし');
    console.log('在庫情報取得結果:', inventoryData ? 'success' : 'null');
    
    if (inventoryData) {
      console.log('在庫データ更新開始...');
      updateRowWithInventoryData(sheet, targetRowIndex, inventoryData);
      console.log(`商品コード ${goodsCode} の更新が完了しました`);
    } else {
      console.log(`商品コード ${goodsCode} の在庫情報が見つかりませんでした`);
    }
    
  } catch (error) {
    console.error('単品更新エラー:', error.message);
    console.error('エラー発生箇所の詳細:', error.stack);
    throw error;
  }
}

/**
 * 在庫マスタAPIのフィールド確認用テスト関数
 */
function testStockMasterFields(goodsCode = "dcmcoverg-s-S") {
  try {
    console.log('=== 在庫マスタAPI フィールド確認テスト ===');
    console.log(`対象商品コード: ${goodsCode}`);
    
    const tokens = getStoredTokens();
    const url = `${NE_API_URL}/api_v1_master_stock/search`;
    
    const payload = {
      'access_token': tokens.accessToken,
      'refresh_token': tokens.refreshToken,
      'stock_goods_id-eq': goodsCode,
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
    
    const response = UrlFetchApp.fetch(url, options);
    const responseText = response.getContentText();
    const responseData = JSON.parse(responseText);
    
    console.log('API応答結果:', responseData.result);
    
    if (responseData.result === 'success' && responseData.data && responseData.data.length > 0) {
      console.log('取得できたフィールド一覧:');
      const stockData = responseData.data[0];
      Object.keys(stockData).forEach(key => {
        console.log(`- ${key}: ${stockData[key]}`);
      });
    } else {
      console.log('データが取得できませんでした');
      console.log('応答詳細:', JSON.stringify(responseData, null, 2));
    }
    
  } catch (error) {
    console.error('テストエラー:', error.message);
  }
}

/**
 * 在庫情報の手動リセット（テスト用）
 * すべての在庫数値を0にリセット
 */
function resetAllInventoryData() {
  try {
    console.log('=== 在庫情報リセット開始 ===');
    
    const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = spreadsheet.getSheetByName(SHEET_NAME);
    
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) {
      console.log('データが存在しません');
      return;
    }
    
    // C列からK列までを0でクリア
    const range = sheet.getRange(2, COLUMNS.STOCK_QTY + 1, lastRow - 1, 9);
    const resetValues = Array(lastRow - 1).fill(Array(9).fill(0));
    range.setValues(resetValues);
    
    console.log(`${lastRow - 1}行の在庫情報をリセットしました`);
    
  } catch (error) {
    console.error('リセットエラー:', error.message);
    throw error;
  }
}

/**
 * スプレッドシートのシート名を確認
 */
function checkSheetNames() {
  try {
    console.log('使用しているスプレッドシートID:', SPREADSHEET_ID);
    const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheets = spreadsheet.getSheets();
    
    console.log('=== スプレッドシート内のシート名一覧 ===');
    for (let i = 0; i < sheets.length; i++) {
      const sheetName = sheets[i].getName();
      console.log(`シート${i + 1}: "${sheetName}" (文字数: ${sheetName.length})`);
    }
    console.log('');
    console.log('現在のSHEET_NAME設定:', `"${SHEET_NAME}"`);
    console.log('上記のシート名のいずれかと完全に一致するようにSHEET_NAMEを設定してください');
    
  } catch (error) {
    console.error('シート名確認エラー:', error.message);
    console.error('スプレッドシートIDが正しいか確認してください');
  }
}

/**
 * スクリプト使用方法ガイド
 */
function showUsageGuide() {
  console.log('=== 在庫情報取得スクリプト使用方法（完全版） ===');
  console.log('');
  console.log('【前提条件】');
  console.log('- 認証スクリプトでトークンが取得済みであること');
  console.log('- スプレッドシートに商品コードが入力済みであること');
  console.log('');
  console.log('【主要関数】');
  console.log('1. updateInventoryData()');
  console.log('   - 全商品の詳細在庫情報を更新');
  console.log('   - 処理時間: 商品数 × 約4秒（2つのAPIを呼び出すため）');
  console.log('');
  console.log('2. updateSingleProduct("商品コード")');
  console.log('   - 特定商品の詳細在庫情報のみ更新（テスト用）');
  console.log('   - 例: updateSingleProduct("dcmcoverg-s-S")');
  console.log('');
  console.log('3. testStockMasterFields("商品コード")');
  console.log('   - 在庫マスタAPIの動作確認（デバッグ用）');
  console.log('');
  console.log('4. resetAllInventoryData()');
  console.log('   - 全在庫数値を0にリセット（テスト用）');
  console.log('');
  console.log('【更新される在庫情報】');
  console.log('- C列: 在庫数');
  console.log('- D列: 引当数');
  console.log('- E列: フリー在庫数');
  console.log('- F列: 予約在庫数');
  console.log('- G列: 予約引当数');
  console.log('- H列: 予約フリー在庫数');
  console.log('- I列: 不良在庫数');
  console.log('- J列: 発注残