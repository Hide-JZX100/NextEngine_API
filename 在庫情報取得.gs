/**
 * ネクストエンジン在庫情報取得スクリプト
 * 
 * 【目的】
 * スプレッドシートの商品コードに対応する在庫情報をネクストエンジンAPIから取得し更新
 * 
 * 【機能】
 * 1. スプレッドシートから商品コードを読み取り
 * 2. ネクストエンジンAPIで在庫情報を取得
 * 3. スプレッドシートの在庫情報を更新
 * 
 * 【注意事項】
 * - 認証スクリプトで事前にトークンを取得済みである必要があります
 * - API制限を考慮して適切な間隔でリクエストを送信します
 * - 大量データの場合は時間がかかる可能性があります
 */

// ネクストエンジンAPIのエンドポイントは認証.gsで定義済み

// スプレッドシートの設定
const SPREADSHEET_ID = '1noQTPM0EMlyBNDdX4JDPZcBvh-3RT1VtWzNDA85SIkM';
const SHEET_NAME = 'シート1'; // 必要に応じて変更してください

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
        
        // API制限を考慮して少し待機（1秒）
        Utilities.sleep(1000);
        
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
 * 商品コードから在庫情報を取得
 * @param {string} goodsCode - 商品コード
 * @param {Object} tokens - アクセストークンとリフレッシュトークン
 * @returns {Object|null} 在庫情報
 */
async function getInventoryByGoodsCode(goodsCode, tokens) {
  try {
    // 商品マスタから商品IDを取得
    const goodsData = await searchGoodsByCode(goodsCode, tokens);
    
    if (!goodsData || goodsData.length === 0) {
      console.log(`商品コード ${goodsCode} が見つかりません`);
      return null;
    }
    
    const goodsId = goodsData[0].goods_id;
    
    // 在庫情報を取得
    const inventoryData = await getStockByGoodsId(goodsId, tokens);
    
    return inventoryData;
    
  } catch (error) {
    console.error(`商品コード ${goodsCode} の在庫取得エラー:`, error.message);
    return null;
  }
}

/**
 * 商品コードで商品マスタを検索
 * @param {string} goodsCode - 商品コード
 * @param {Object} tokens - トークン情報
 * @returns {Array} 商品マスタ情報
 */
async function searchGoodsByCode(goodsCode, tokens) {
  const url = `${NE_API_URL}/api_v1_master_goods/search`;
  
  const payload = {
    'access_token': tokens.accessToken,
    'refresh_token': tokens.refreshToken,
    'goods_code-eq': goodsCode, // 商品コードで完全一致検索
    'fields': 'goods_id,goods_code,goods_name'
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
  
  // トークンが更新された場合は保存
  if (responseData.access_token && responseData.refresh_token) {
    updateStoredTokens(responseData.access_token, responseData.refresh_token);
  }
  
  if (responseData.result === 'success') {
    return responseData.data;
  } else {
    throw new Error(`商品検索エラー: ${JSON.stringify(responseData)}`);
  }
}

/**
 * 商品IDから在庫情報を取得
 * @param {string} goodsId - 商品ID
 * @param {Object} tokens - トークン情報
 * @returns {Object} 在庫情報
 */
async function getStockByGoodsId(goodsId, tokens) {
  const url = `${NE_API_URL}/api_v1_master_stock/search`;
  
  const payload = {
    'access_token': tokens.accessToken,
    'refresh_token': tokens.refreshToken,
    'goods_id-eq': goodsId, // 商品IDで検索
    'fields': 'stock_quantity,allocated_quantity,free_stock_quantity,reserve_stock_quantity,reserve_allocated_quantity,reserve_free_stock_quantity,defective_stock_quantity,order_remaining_quantity,shortage_quantity'
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
  
  // トークンが更新された場合は保存
  if (responseData.access_token && responseData.refresh_token) {
    updateStoredTokens(responseData.access_token, responseData.refresh_token);
  }
  
  if (responseData.result === 'success' && responseData.data.length > 0) {
    return responseData.data[0]; // 最初のレコードを返す
  } else {
    console.log(`商品ID ${goodsId} の在庫情報が見つかりません`);
    return null;
  }
}

/**
 * スプレッドシートの行を在庫データで更新
 * @param {Sheet} sheet - シートオブジェクト
 * @param {number} rowIndex - 更新する行番号（1ベース）
 * @param {Object} inventoryData - 在庫データ
 */
function updateRowWithInventoryData(sheet, rowIndex, inventoryData) {
  // 在庫情報の列を更新（C列からK列まで）
  const updateValues = [
    inventoryData.stock_quantity || 0,           // C列: 在庫数
    inventoryData.allocated_quantity || 0,       // D列: 引当数
    inventoryData.free_stock_quantity || 0,      // E列: フリー在庫数
    inventoryData.reserve_stock_quantity || 0,   // F列: 予約在庫数
    inventoryData.reserve_allocated_quantity || 0, // G列: 予約引当数
    inventoryData.reserve_free_stock_quantity || 0, // H列: 予約フリー在庫数
    inventoryData.defective_stock_quantity || 0, // I列: 不良在庫数
    inventoryData.order_remaining_quantity || 0, // J列: 発注残数
    inventoryData.shortage_quantity || 0         // K列: 欠品数
  ];
  
  // C列からK列まで更新
  const range = sheet.getRange(rowIndex, COLUMNS.STOCK_QTY + 1, 1, updateValues.length);
  range.setValues([updateValues]);
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
 * 特定の商品コードのみ更新（テスト用）
 * @param {string} goodsCode - 更新したい商品コード
 */
function updateSingleProduct(goodsCode) {
  try {
    console.log(`=== 単品更新開始: ${goodsCode} ===`);
    
    const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = spreadsheet.getSheetByName(SHEET_NAME);
    
    // 商品コードを検索
    const dataRange = sheet.getRange(2, 1, sheet.getLastRow() - 1, 12);
    const values = dataRange.getValues();
    
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
    
    // トークンを取得
    const tokens = getStoredTokens();
    
    // 在庫情報を取得
    const inventoryData = getInventoryByGoodsCode(goodsCode, tokens);
    
    if (inventoryData) {
      updateRowWithInventoryData(sheet, targetRowIndex, inventoryData);
      console.log(`商品コード ${goodsCode} の更新が完了しました`);
    } else {
      console.log(`商品コード ${goodsCode} の在庫情報が見つかりませんでした`);
    }
    
  } catch (error) {
    console.error('単品更新エラー:', error.message);
    throw error;
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
    const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheets = spreadsheet.getSheets();
    
    console.log('=== スプレッドシート内のシート名一覧 ===');
    for (let i = 0; i < sheets.length; i++) {
      console.log(`シート${i + 1}: "${sheets[i].getName()}"`);
    }
    console.log('');
    console.log('上記のシート名のいずれかをSHEET_NAMEに設定してください');
    
  } catch (error) {
    console.error('シート名確認エラー:', error.message);
  }
}

/**
 * スクリプト使用方法ガイド
 */
function showUsageGuide() {
  console.log('=== 在庫情報取得スクリプト使用方法 ===');
  console.log('');
  console.log('【前提条件】');
  console.log('- 認証スクリプトでトークンが取得済みであること');
  console.log('- スプレッドシートに商品コードが入力済みであること');
  console.log('');
  console.log('【主要関数】');
  console.log('1. updateInventoryData()');
  console.log('   - 全商品の在庫情報を更新');
  console.log('   - 処理時間: 商品数 × 約2秒');
  console.log('');
  console.log('2. updateSingleProduct("商品コード")');
  console.log('   - 特定商品の在庫情報のみ更新（テスト用）');
  console.log('   - 例: updateSingleProduct("ABC123")');
  console.log('');
  console.log('3. resetAllInventoryData()');
  console.log('   - 全在庫数値を0にリセット（テスト用）');
  console.log('');
  console.log('【注意事項】');
  console.log('- APIレート制限のため各商品間に1秒の待機時間があります');
  console.log('- 大量の商品がある場合は処理時間が長くなります');
  console.log('- エラーが発生した商品はスキップされます');
  console.log('');
  console.log('【実行推奨順序】');
  console.log('1. まず updateSingleProduct() でテスト');
  console.log('2. 問題なければ updateInventoryData() で全更新');
}