/**
 * 単一API調査・最適化版スクリプト
 * 
 * 【目的】
 * 在庫マスタAPIのみでの処理可能性を調査し、APIコール数をさらに削減
 * 
 * 【調査ポイント】
 * 1. 在庫マスタAPIで商品名も取得可能か？
 * 2. 商品マスタAPIの情報が本当に必要か？
 * 3. 単一APIでの処理による性能改善効果
 */

/**
 * 在庫マスタAPIで取得可能なフィールドを調査
 * @param {string[]} sampleGoodsCodeList - サンプル商品コードリスト
 * @param {Object} tokens - トークン情報
 */
function investigateStockApiFields(sampleGoodsCodeList, tokens) {
  const url = `${NE_API_URL}/api_v1_master_stock/search`;
  
  // サンプル商品コード（最初の5件程度）
  const sampleCodes = sampleGoodsCodeList.slice(0, 5);
  const goodsIdCondition = sampleCodes.join(',');
  
  const payload = {
    'access_token': tokens.accessToken,
    'refresh_token': tokens.refreshToken,
    'stock_goods_id-in': goodsIdCondition,
    'fields': '', // 空にして全フィールドを取得
    'limit': '5'
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
  
  console.log('=== 在庫マスタAPI フィールド調査 ===');
  console.log(`調査対象商品: ${sampleCodes.join(', ')}`);
  
  try {
    const response = UrlFetchApp.fetch(url, options);
    const responseText = response.getContentText();
    const responseData = JSON.parse(responseText);
    
    if (responseData.result === 'success' && responseData.data && responseData.data.length > 0) {
      const firstRecord = responseData.data[0];
      
      console.log('\n--- 利用可能フィールド一覧 ---');
      Object.keys(firstRecord).forEach(field => {
        console.log(`${field}: ${firstRecord[field]}`);
      });
      
      // 商品名関連のフィールドがあるかチェック
      const hasGoodsName = Object.keys(firstRecord).some(field => 
        field.toLowerCase().includes('name') || field.toLowerCase().includes('goods_name')
      );
      
      console.log('\n--- 調査結果 ---');
      console.log(`商品名フィールドの存在: ${hasGoodsName ? '有り' : '無し'}`);
      console.log(`取得可能フィールド数: ${Object.keys(firstRecord).length}`);
      
      return {
        success: true,
        hasGoodsName,
        availableFields: Object.keys(firstRecord),
        sampleData: firstRecord
      };
      
    } else {
      console.error('在庫マスタAPI調査失敗:', responseData.message);
      return { success: false, error: responseData.message };
    }
    
  } catch (error) {
    console.error('在庫マスタAPI調査エラー:', error.message);
    return { success: false, error: error.message };
  }
}

/**
 * 在庫マスタAPI単体でのバッチ処理テスト
 * @param {string[]} goodsCodeList - 商品コードの配列
 * @param {Object} tokens - トークン情報
 * @returns {Map<string, Object>} 商品コード → 在庫情報のマップ
 */
function getSingleApiInventoryData(goodsCodeList, tokens) {
  const url = `${NE_API_URL}/api_v1_master_stock/search`;
  const properties = PropertiesService.getScriptProperties();
  const batchSize = parseInt(properties.getProperty('BATCH_SIZE')) || 100;
  
  const goodsIdCondition = goodsCodeList.join(',');
  
  const payload = {
    'access_token': tokens.accessToken,
    'refresh_token': tokens.refreshToken,
    'stock_goods_id-in': goodsIdCondition,
    // 必要な在庫情報フィールドを明示的に指定
    'fields': 'stock_goods_id,stock_quantity,stock_allocation_quantity,stock_defective_quantity,stock_remaining_order_quantity,stock_out_quantity,stock_free_quantity,stock_advance_order_quantity,stock_advance_order_allocation_quantity,stock_advance_order_free_quantity',
    'limit': batchSize.toString()
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
  
  const inventoryDataMap = new Map();
  
  try {
    console.log(`  在庫マスタAPI単体呼び出し: ${goodsCodeList.length}件`);
    
    const response = UrlFetchApp.fetch(url, options);
    const responseText = response.getContentText();
    const responseData = JSON.parse(responseText);
    
    console.log(`    API応答: result=${responseData.result}, count=${responseData.count || 0}`);
    
    // トークンが更新された場合は保存
    if (responseData.access_token && responseData.refresh_token) {
      updateStoredTokens(responseData.access_token, responseData.refresh_token);
      tokens.accessToken = responseData.access_token;
      tokens.refreshToken = responseData.refresh_token;
    }
    
    if (responseData.result === 'success' && responseData.data) {
      responseData.data.forEach(stockData => {
        const completeInventoryData = {
          goods_id: stockData.stock_goods_id,
          goods_name: stockData.stock_goods_id, // 商品名が取得できない場合は商品IDを使用
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
        
        inventoryDataMap.set(stockData.stock_goods_id, completeInventoryData);
      });
      
      console.log(`    取得完了: ${responseData.data.length}件`);
    } else {
      console.error(`    在庫マスタAPI エラー:`, responseData.message || 'Unknown error');
    }
    
    return inventoryDataMap;
    
  } catch (error) {
    console.error(`在庫マスタ単体取得エラー:`, error.message);
    return inventoryDataMap;
  }
}

/**
 * 単一API版メイン関数：在庫マスタAPIのみで在庫情報更新
 */
function updateInventoryDataSingleApi() {
  try {
    console.log('=== 単一API版在庫情報更新開始 ===');
    const startTime = new Date();
    
    // スクリプトプロパティから設定値を取得
    const properties = PropertiesService.getScriptProperties();
    const spreadsheetId = properties.getProperty('SPREADSHEET_ID');
    const sheetName = properties.getProperty('SHEET_NAME');
    const batchSize = parseInt(properties.getProperty('BATCH_SIZE')) || 100;
    const apiWaitTime = parseInt(properties.getProperty('API_WAIT_TIME')) || 500;
    
    if (!spreadsheetId || !sheetName) {
      throw new Error('SPREADSHEET_IDまたはSHEET_NAMEがスクリプトプロパティに設定されていません。');
    }
    
    // スプレッドシートを取得
    const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    const sheet = spreadsheet.getSheetByName(sheetName);
    
    if (!sheet) {
      throw new Error(`シート "${sheetName}" が見つかりません`);
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
    const rowIndexMap = new Map();
    
    for (let i = 0; i < values.length; i++) {
      const goodsCode = values[i][COLUMNS.GOODS_CODE];
      if (goodsCode && goodsCode.toString().trim()) {
        goodsCodeList.push(goodsCode.toString().trim());
        rowIndexMap.set(goodsCode.toString().trim(), i + 2);
      }
    }
    
    console.log(`有効な商品コード: ${goodsCodeList.length}件`);
    
    if (goodsCodeList.length === 0) {
      console.log('処理対象の商品コードがありません');
      return;
    }
    
    // 単一API版バッチ処理で在庫情報を取得・更新
    let totalUpdated = 0;
    let totalErrors = 0;
    let totalApiCalls = 0;
    
    for (let i = 0; i < goodsCodeList.length; i += batchSize) {
      const batch = goodsCodeList.slice(i, i + batchSize);
      console.log(`\n--- バッチ ${Math.floor(i / batchSize) + 1}: ${batch.length}件 ---`);
      
      try {
        // 単一APIで在庫情報を取得
        const inventoryDataMap = getSingleApiInventoryData(batch, tokens);
        totalApiCalls++; // 1回のAPIコールのみ
        
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
        
        // バッチ間の待機
        if (i + batchSize < goodsCodeList.length) {
          console.log(`次のバッチまで ${apiWaitTime}ms 待機...`);
          Utilities.sleep(apiWaitTime);
        }
        
      } catch (error) {
        console.error(`バッチ処理エラー:`, error.message);
        totalErrors += batch.length;
      }
    }
    
    const endTime = new Date();
    const duration = (endTime - startTime) / 1000;
    
    console.log('\n=== 単一API版更新完了 ===');
    console.log(`処理時間: ${duration.toFixed(1)}秒`);
    console.log(`APIコール数: ${totalApiCalls}回`);
    console.log(`更新成功: ${totalUpdated}件`);
    console.log(`エラー: ${totalErrors}件`);
    console.log(`処理速度: ${(goodsCodeList.length / duration).toFixed(1)}件/秒`);
    
    // 従来版・二重API版との比較
    const conventionalTime = goodsCodeList.length * 2;
    const dualApiCalls = Math.ceil(goodsCodeList.length / batchSize) * 2;
    const singleApiImprovement = dualApiCalls / totalApiCalls;
    
    console.log(`\n--- API最適化結果 ---`);
    console.log(`従来版APIコール数: ${goodsCodeList.length * 2}回`);
    console.log(`二重API版コール数: ${dualApiCalls}回`);
    console.log(`単一API版コール数: ${totalApiCalls}回`);
    console.log(`二重API版からの削減率: ${((dualApiCalls - totalApiCalls) / dualApiCalls * 100).toFixed(1)}%`);
    console.log(`APIコール削減倍率: ${singleApiImprovement.toFixed(1)}倍`);
    
  } catch (error) {
    console.error('単一API版更新エラー:', error.message);
    throw error;
  }
}

/**
 * 段階的フィールド調査：在庫マスタAPIで商品名フィールドの存在を確認
 * @param {string[]} sampleGoodsCodeList - サンプル商品コードリスト
 * @param {Object} tokens - トークン情報
 */
function investigateStockApiStepByStep(sampleGoodsCodeList, tokens) {
  const url = `${NE_API_URL}/api_v1_master_stock/search`;
  const sampleCodes = sampleGoodsCodeList.slice(0, 3);
  const goodsIdCondition = sampleCodes.join(',');
  
  console.log('=== 段階的在庫マスタAPI フィールド調査 ===');
  console.log(`調査対象商品: ${sampleCodes.join(', ')}`);
  
  // 段階1: 基本的な在庫フィールドのみで試行
  const basicFields = 'stock_goods_id,stock_quantity,stock_allocation_quantity';
  console.log('\n--- 段階1: 基本フィールド調査 ---');
  
  const basicPayload = {
    'access_token': tokens.accessToken,
    'refresh_token': tokens.refreshToken,
    'stock_goods_id-in': goodsIdCondition,
    'fields': basicFields,
    'limit': '3'
  };
  
  const options = {
    'method': 'POST',
    'headers': {
      'Content-Type': 'application/x-www-form-urlencoded'
    },
    'payload': Object.keys(basicPayload).map(key =>
      encodeURIComponent(key) + '=' + encodeURIComponent(basicPayload[key])
    ).join('&')
  };
  
  try {
    let response = UrlFetchApp.fetch(url, options);
    let responseText = response.getContentText();
    let responseData = JSON.parse(responseText);
    
    if (responseData.result === 'success' && responseData.data && responseData.data.length > 0) {
      console.log('✓ 基本フィールドでの取得成功');
      console.log(`取得件数: ${responseData.data.length}件`);
      
      // 段階2: 商品名フィールドを追加して試行
      console.log('\n--- 段階2: 商品名フィールド追加調査 ---');
      
      const candidateNameFields = [
        'stock_goods_name',
        'goods_name', 
        'name',
        'product_name',
        'item_name'
      ];
      
      let successfulNameField = null;
      
      for (const nameField of candidateNameFields) {
        console.log(`  ${nameField} フィールドを試行中...`);
        
        const extendedPayload = {
          'access_token': tokens.accessToken,
          'refresh_token': tokens.refreshToken,
          'stock_goods_id-in': goodsIdCondition,
          'fields': `${basicFields},${nameField}`,
          'limit': '3'
        };
        
        const extendedOptions = {
          'method': 'POST',
          'headers': {
            'Content-Type': 'application/x-www-form-urlencoded'
          },
          'payload': Object.keys(extendedPayload).map(key =>
            encodeURIComponent(key) + '=' + encodeURIComponent(extendedPayload[key])
          ).join('&')
        };
        
        try {
          response = UrlFetchApp.fetch(url, extendedOptions);
          responseText = response.getContentText();
          responseData = JSON.parse(responseText);
          
          if (responseData.result === 'success') {
            console.log(`    ✓ ${nameField}: 成功`);
            successfulNameField = nameField;
            
            // 実際のデータサンプルを表示
            if (responseData.data && responseData.data.length > 0) {
              const sample = responseData.data[0];
              console.log(`    サンプルデータ: ${JSON.stringify(sample)}`);
            }
            break;
          } else {
            console.log(`    ✗ ${nameField}: ${responseData.message}`);
          }
        } catch (error) {
          console.log(`    ✗ ${nameField}: エラー - ${error.message}`);
        }
        
        // API負荷軽減のため少し待機
        Utilities.sleep(200);
      }
      
      console.log('\n--- 調査結果まとめ ---');
      if (successfulNameField) {
        console.log(`✓ 商品名フィールド発見: ${successfulNameField}`);
        console.log('→ 単一API処理が可能です');
        return {
          success: true,
          hasGoodsName: true,
          goodsNameField: successfulNameField,
          recommendedFields: `${basicFields},${successfulNameField}`
        };
      } else {
        console.log('✗ 商品名フィールドが見つかりませんでした');
        console.log('→ 在庫数値のみの更新なら単一API処理可能');
        console.log('→ 商品名が必要な場合は二重API処理が必要');
        return {
          success: true,
          hasGoodsName: false,
          recommendedFields: basicFields
        };
      }
      
    } else {
      console.error('基本フィールドでの取得に失敗:', responseData.message);
      return { success: false, error: responseData.message };
    }
    
  } catch (error) {
    console.error('段階的調査エラー:', error.message);
    return { success: false, error: error.message };
  }
}

/**
 * 改良版APIフィールド調査テスト関数
 */
function testApiFieldInvestigationImproved() {
  try {
    console.log('=== 改良版APIフィールド調査テスト開始 ===');
    
    // スクリプトプロパティから設定値を取得
    const properties = PropertiesService.getScriptProperties();
    const spreadsheetId = properties.getProperty('SPREADSHEET_ID');
    const sheetName = properties.getProperty('SHEET_NAME');
    
    if (!spreadsheetId || !sheetName) {
      throw new Error('SPREADSHEET_IDまたはSHEET_NAMEがスクリプトプロパティに設定されていません。');
    }
    
    // スプレッドシートから最初の5件の商品コードを取得
    const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    const sheet = spreadsheet.getSheetByName(sheetName);
    const lastRow = sheet.getLastRow();
    
    if (lastRow <= 1) {
      console.log('テスト用データが存在しません');
      return;
    }
    
    const dataRange = sheet.getRange(2, 1, Math.min(5, lastRow - 1), 1);
    const values = dataRange.getValues();
    const sampleGoodsCodeList = values
      .map(row => row[0])
      .filter(code => code && code.toString().trim());
    
    const tokens = getStoredTokens();
    
    // 段階的フィールド調査実行
    const investigationResult = investigateStockApiStepByStep(sampleGoodsCodeList, tokens);
    
    if (investigationResult.success) {
      console.log('\n=== 最終調査結果 ===');
      
      if (investigationResult.hasGoodsName) {
        console.log('✅ 単一API処理が推奨されます');
        console.log(`推奨フィールド設定: ${investigationResult.recommendedFields}`);
        console.log('次のステップ: compareApiVersions(10) で性能比較');
      } else {
        console.log('⚠️ 商品名が取得できません');
        console.log('選択肢:');
        console.log('1. 商品名不要なら単一API処理可能');
        console.log('2. 商品名必要なら現在の二重API処理継続');
        console.log(`在庫のみフィールド設定: ${investigationResult.recommendedFields}`);
      }
    } else {
      console.log('❌ 調査に失敗しました:', investigationResult.error);
    }
    
  } catch (error) {
    console.error('改良版APIフィールド調査エラー:', error.message);
  }
}

/**
 * 二重API版の処理（比較用）- 別ファイルの関数を再実装
 * @param {string[]} goodsCodeList - 商品コードの配列
 * @param {Object} tokens - トークン情報
 * @returns {Map<string, Object>} 商品コード → 在庫情報のマップ
 */
function getDualApiInventoryData(goodsCodeList, tokens) {
  const inventoryDataMap = new Map();
  
  try {
    console.log(`  二重API処理: ${goodsCodeList.length}件`);
    
    // ステップ1: 商品マスタAPIで複数商品を一括検索
    const goodsDataMap = getBatchGoodsDataForComparison(goodsCodeList, tokens);
    console.log(`  商品マスタ取得完了: ${goodsDataMap.size}件`);
    
    if (goodsDataMap.size === 0) {
      console.log('  商品が見つかりませんでした');
      return inventoryDataMap;
    }
    
    // ステップ2: 在庫マスタAPIで複数商品の在庫を一括取得
    const stockDataMap = getBatchStockDataForComparison(Array.from(goodsDataMap.keys()), tokens);
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
    console.error(`二重API取得エラー:`, error.message);
    return inventoryDataMap;
  }
}

/**
 * 商品マスタAPI呼び出し（比較用）
 */
function getBatchGoodsDataForComparison(goodsCodeList, tokens) {
  const url = `${NE_API_URL}/api_v1_master_goods/search`;
  const properties = PropertiesService.getScriptProperties();
  const batchSize = parseInt(properties.getProperty('BATCH_SIZE')) || 100;
  const goodsIdCondition = goodsCodeList.join(',');
  
  const payload = {
    'access_token': tokens.accessToken,
    'refresh_token': tokens.refreshToken,
    'goods_id-in': goodsIdCondition,
    'fields': 'goods_id,goods_name,stock_quantity',
    'limit': batchSize.toString()
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
    
    if (responseData.access_token && responseData.refresh_token) {
      updateStoredTokens(responseData.access_token, responseData.refresh_token);
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
    }
    
    return goodsDataMap;
    
  } catch (error) {
    console.error(`商品マスタ比較用取得エラー:`, error.message);
    return goodsDataMap;
  }
}

/**
 * 在庫マスタAPI呼び出し（比較用）
 */
function getBatchStockDataForComparison(goodsCodeList, tokens) {
  const url = `${NE_API_URL}/api_v1_master_stock/search`;
  const properties = PropertiesService.getScriptProperties();
  const batchSize = parseInt(properties.getProperty('BATCH_SIZE')) || 100;
  const goodsIdCondition = goodsCodeList.join(',');
  
  const payload = {
    'access_token': tokens.accessToken,
    'refresh_token': tokens.refreshToken,
    'stock_goods_id-in': goodsIdCondition,
    'fields': 'stock_goods_id,stock_quantity,stock_allocation_quantity,stock_defective_quantity,stock_remaining_order_quantity,stock_out_quantity,stock_free_quantity,stock_advance_order_quantity,stock_advance_order_allocation_quantity,stock_advance_order_free_quantity',
    'limit': batchSize.toString()
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
    
    if (responseData.access_token && responseData.refresh_token) {
      updateStoredTokens(responseData.access_token, responseData.refresh_token);
      tokens.accessToken = responseData.access_token;
      tokens.refreshToken = responseData.refresh_token;
    }
    
    if (responseData.result === 'success' && responseData.data) {
      responseData.data.forEach(stockData => {
        stockDataMap.set(stockData.stock_goods_id, stockData);
      });
    }
    
    return stockDataMap;
    
  } catch (error) {
    console.error(`在庫マスタ比較用取得エラー:`, error.message);
    return stockDataMap;
  }
}

/**
 * 性能比較テスト：二重API版 vs 単一API版
 */
function compareApiVersions(sampleSize = 10) {
  console.log(`=== API版本比較テスト（${sampleSize}件） ===`);
  
  try {
    // スクリプトプロパティから設定値を取得
    const properties = PropertiesService.getScriptProperties();
    const spreadsheetId = properties.getProperty('SPREADSHEET_ID');
    const sheetName = properties.getProperty('SHEET_NAME');
    
    if (!spreadsheetId || !sheetName) {
      throw new Error('設定が不完全です');
    }
    
    // サンプル商品コードを取得
    const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    const sheet = spreadsheet.getSheetByName(sheetName);
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
    
    // 二重API版テスト
    console.log('\n--- 二重API版実行 ---');
    const dualApiStartTime = new Date();
    const dualApiResult = getDualApiInventoryData(goodsCodeList, tokens);
    const dualApiEndTime = new Date();
    const dualApiDuration = (dualApiEndTime - dualApiStartTime) / 1000;
    
    // 単一API版テスト
    console.log('\n--- 単一API版実行 ---');
    const singleApiStartTime = new Date();
    const singleApiResult = getSingleApiInventoryData(goodsCodeList, tokens);
    const singleApiEndTime = new Date();
    const singleApiDuration = (singleApiEndTime - singleApiStartTime) / 1000;
    
    // 結果比較
    console.log('\n=== 比較結果 ===');
    console.log(`二重API版時間: ${dualApiDuration.toFixed(1)}秒`);
    console.log(`単一API版時間: ${singleApiDuration.toFixed(1)}秒`);
    
    if (dualApiDuration > 0) {
      const timeReduction = ((dualApiDuration - singleApiDuration) / dualApiDuration * 100);
      console.log(`時間短縮効果: ${timeReduction.toFixed(1)}%`);
    }
    
    console.log(`二重API版取得件数: ${dualApiResult.size}件`);
    console.log(`単一API版取得件数: ${singleApiResult.size}件`);
    
    if (dualApiResult.size > 0) {
      console.log(`取得率比較: ${(singleApiResult.size / dualApiResult.size * 100).toFixed(1)}%`);
    }
    
    // APIコール数比較
    const dualApiCalls = 2; // 商品マスタ + 在庫マスタ
    const singleApiCalls = 1; // 在庫マスタのみ
    console.log(`APIコール数削減: ${dualApiCalls}回 → ${singleApiCalls}回（${((dualApiCalls - singleApiCalls) / dualApiCalls * 100).toFixed(0)}%削減）`);
    
    // 大量データでの推定効果
    console.log('\n=== 3106件での推定効果 ===');
    const estimatedDualTime = (dualApiDuration / sampleSize) * 3106;
    const estimatedSingleTime = (singleApiDuration / sampleSize) * 3106;
    console.log(`二重API版推定時間: ${estimatedDualTime.toFixed(1)}秒`);
    console.log(`単一API版推定時間: ${estimatedSingleTime.toFixed(1)}秒`);
    console.log(`推定時間短縮: ${(estimatedDualTime - estimatedSingleTime).toFixed(1)}秒`);
    
  } catch (error) {
    console.error('比較テストエラー:', error.message);
  }
}

/**
 * 単一API版使用方法ガイド
 */
function showSingleApiUsageGuide() {
  console.log('=== 単一API版使用方法ガイド ===');
  console.log('');
  console.log('【事前調査】');
  console.log('1. testApiFieldInvestigation()');
  console.log('   - 在庫マスタAPIで商品名が取得可能かを調査');
  console.log('   - 単一API処理の可否を判定');
  console.log('');
  console.log('【性能テスト】');
  console.log('2. compareApiVersions(件数)');
  console.log('   - 二重API版と単一API版の性能比較');
  console.log('   - 例: compareApiVersions(20)');
  console.log('');
  console.log('【実行】');
  console.log('3. updateInventoryDataSingleApi()');
  console.log('   - 在庫マスタAPIのみで全件更新');
  console.log('   - APIコール数を約50%削減');
  console.log('');
  console.log('【期待効果】');
  console.log('- APIサーバー負荷: 50%削減');
  console.log('- 処理時間: さらなる短縮');
  console.log('- 安定性: APIエラーリスクの軽減');
}