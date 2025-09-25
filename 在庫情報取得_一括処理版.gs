/**
=============================================================================
ネクストエンジン在庫情報取得スクリプト（一括処理版 + 単一API版テスト）
=============================================================================
* 【目的】
* 商品コードを配列で渡し、一度のAPIコールで複数商品の在庫情報を効率的に取得

* 【主な改善点】
* 1. 商品マスタAPIで複数商品を一度に検索（最大1000件）←商品マスタAPIは使わず
     在庫マスタAPIだけに修正
* 2. 在庫マスタAPIで複数商品の在庫を一度に取得（最大1000件）
     const MAX_ITEMS_PER_CALL = 1000; で定義
* 3. バッチ処理による大幅な高速化
* 4. APIコール数の削減によるレート制限回避
* 5. 単一API版での性能比較テスト追加

* 【注意事項】
* - 認証スクリプトで事前にトークンを取得済みである必要があります
* - 一度に処理できる商品数は最大1000件です
* - 大量データの場合は自動的にバッチ分割します


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

getBatchStockData(goodsCodeList, tokens)
在庫マスタAPI（／api_v1_master_stock／search）を呼び出し、
複数の商品について在庫情報（在庫数、引当数など）をまとめて取得します。
これもstock_goods_id-inパラメータを利用して、効率的な一括検索を行います。

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

getSpreadsheetConfig()
スクリプトプロパティからスプレッドシートのIDとシート名を取得します。
設定がなければエラーを発生させ、スクリプトの実行に必要な情報が揃っているかを確認します。

getStoredTokens()
スクリプトプロパティに保存されているアクセストークンとリフレッシュトークンを取得します。
API呼び出しのたびに認証情報を取得する手間を省くためのユーティリティ関数です。

updateStoredTokens(accessToken, refreshToken)
ネクストエンジンAPIから新しいトークンが返された際に、
スクリプトプロパティを更新し、トークン情報を保存します。

updateRowWithInventoryData(sheet, rowIndex, inventoryData)
取得した在庫情報（inventoryData）を基に、
スプレッドシートの指定された行（rowIndex）の在庫関連の列を更新します。

logErrorsToSheet(errorDetails)
処理中に発生したエラーの詳細を、スプレッドシート上の「エラーログ」シートに記録します。
これにより、どの商品でどのような問題が発生したかを後から確認できます。

showCurrentProperties()
現在のスクリプトプロパティの設定内容をログに出力します。

showUsageGuide()
スクリプトの主要な機能、使用方法、そして期待される効果について説明します。

=============================================================================
*/

// ファイルトップに（既に定義されているなら上書きしない）
if (typeof NE_API_URL === 'undefined') {
  const NE_API_URL = PropertiesService.getScriptProperties().getProperty('NE_API_URL') || 'https://api.next-engine.org';
}

/**
 * スプレッドシート設定を取得
 */
function getSpreadsheetConfig() {
  const properties = PropertiesService.getScriptProperties();
  const SPREADSHEET_ID = properties.getProperty('SPREADSHEET_ID');
  const SHEET_NAME = properties.getProperty('SHEET_NAME');

  if (!SPREADSHEET_ID || !SHEET_NAME) {
    throw new Error('スプレッドシート設定が不完全です。スクリプトプロパティにSPREADSHEET_IDとSHEET_NAMEを設定してください。');
  }

  return {
    SPREADSHEET_ID,
    SHEET_NAME
  };
}

// 列のマッピング（既存と同じ）
const COLUMNS = {
  GOODS_CODE: 0,        // A列: 商品コード
  GOODS_NAME: 1,        // B列: 商品名
  STOCK_QTY: 2,        // C列: 在庫数
  ALLOCATED_QTY: 3,    // D列: 引当数
  FREE_QTY: 4,         // E列: フリー在庫数
  RESERVE_QTY: 5,      // F列: 予約在庫数
  RESERVE_ALLOCATED_QTY: 6,  // G列: 予約引当数
  RESERVE_FREE_QTY: 7, // H列: 予約フリー在庫数
  DEFECTIVE_QTY: 8,    // I列: 不良在庫数
  ORDER_REMAINING_QTY: 9,    // J列: 発注残数
  SHORTAGE_QTY: 10,    // K列: 欠品数
  JAN_CODE: 11         // L列: JANコード
};

// 設定値
const MAX_ITEMS_PER_CALL = 1000;  // 1回のAPIコールで処理する最大件数（上限1000件）
const API_WAIT_TIME = 500;        // APIコール間の待機時間（ミリ秒）

/**
 * API版本比較テスト：二重API版 vs 単一API版
 * @param {number} sampleSize - テスト対象のサンプル数（デフォルト: 10）
 */
/**
 * 修正版 compareAPIVersions関数
 */
function compareAPIVersions(sampleSize = 10) {
  console.log(`=== API版本比較テスト（${sampleSize}件） ===`);
  
  // スプレッドシートから商品コードを取得（ハードコーディング修正）
  const { SPREADSHEET_ID, SHEET_NAME } = getSpreadsheetConfig();
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

  // ★トークンを1回だけ取得して両方で共有
  const tokens = getStoredTokens();

  // 二重API版実行
  console.log('\n--- 二重API版実行 ---');
  const doubleAPIStartTime = new Date();
  // ★tokensオブジェクトをコピーして渡す
  const doubleAPIResult = fetchInventoryWithDoubleAPI(goodsCodeList, {
    accessToken: tokens.accessToken,
    refreshToken: tokens.refreshToken
  });
  const doubleAPIEndTime = new Date();
  const doubleAPITime = (doubleAPIEndTime - doubleAPIStartTime) / 1000;

  // 単一API版実行
  console.log('\n--- 単一API版実行 ---');
  const singleAPIStartTime = new Date();
  // ★tokensオブジェクトをコピーして渡す
  const singleAPIResult = fetchInventoryWithSingleAPI(goodsCodeList, {
    accessToken: tokens.accessToken,
    refreshToken: tokens.refreshToken
  });
  const singleAPIEndTime = new Date();
  const singleAPITime = (singleAPIEndTime - singleAPIStartTime) / 1000;

  // 結果比較
  console.log('\n=== 比較結果 ===');
  console.log(`二重API版時間: ${doubleAPITime.toFixed(1)}秒`);
  console.log(`単一API版時間: ${singleAPITime.toFixed(1)}秒`);
  console.log(`時間短縮効果: ${doubleAPITime > 0 ? ((doubleAPITime - singleAPITime) / doubleAPITime * 100).toFixed(1) : 0}%`);
  console.log(`二重API版取得件数: ${doubleAPIResult.size}件`);
  console.log(`単一API版取得件数: ${singleAPIResult.size}件`);
  console.log(`取得率比較: ${doubleAPIResult.size > 0 ? (singleAPIResult.size / doubleAPIResult.size * 100).toFixed(1) : 0}%`);
  console.log(`APIコール数削減: 2回 → 1回（50%削減）`);

  // 大量データでの推定効果（3106件での推定）
  const totalItems = 3106;
  if (doubleAPITime > 0 && singleAPITime > 0) {
    const doubleAPIEstimated = totalItems / goodsCodeList.length * doubleAPITime;
    const singleAPIEstimated = totalItems / goodsCodeList.length * singleAPITime;
    
    console.log(`\n=== ${totalItems}件での推定効果 ===`);
    console.log(`二重API版推定時間: ${doubleAPIEstimated.toFixed(1)}秒`);
    console.log(`単一API版推定時間: ${singleAPIEstimated.toFixed(1)}秒`);
    console.log(`推定時間短縮: ${(doubleAPIEstimated - singleAPIEstimated).toFixed(1)}秒`);
  }
}

/**
 * 単一API版のテスト用関数
 * @param {number} maxItems - テスト件数（デフォルト: 5）
 */
function testSingleAPIFunction(maxItems = 5) {
  try {
    console.log(`=== 単一API版テスト（${maxItems}件） ===`);
    
    // スプレッドシートから商品コードを取得
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

    // トークンを取得
    const tokens = getStoredTokens();

    // 単一API版を実行
    const startTime = new Date();
    const result = fetchInventoryWithSingleAPI(goodsCodeList, tokens);
    const endTime = new Date();
    const duration = (endTime - startTime) / 1000;

    console.log('\n=== テスト結果 ===');
    console.log(`処理時間: ${duration.toFixed(1)}秒`);
    console.log(`取得件数: ${result.size}件`);
    console.log(`処理速度: ${(goodsCodeList.length / duration).toFixed(1)}件/秒`);

    // 取得したデータの内容を表示
    console.log('\n--- 取得データ詳細 ---');
    for (const [goodsCode, data] of result) {
      console.log(`${goodsCode}: 在庫${data.stock_quantity} 引当${data.stock_allocated_quantity} フリー${data.stock_free_quantity}`);
    }

  } catch (error) {
    console.error('単一APIテストエラー:', error.message);
    throw error;
  }
}

// 二重API版: 商品マスタ + 在庫マスタの2回のAPI呼び出し
function fetchInventoryWithDoubleAPI(goodsCodes, tokens) {
  try {
    console.log(`  二重API処理: ${goodsCodes.length}件`);
    
    // ステップ1: 商品マスタAPIで商品情報を取得
    const goodsDataMap = getBatchGoodsData(goodsCodes, tokens);
    
    if (goodsDataMap.size === 0) {
      return new Map();
    }
    
    // ステップ2: 在庫マスタAPIで在庫情報を取得
    const stockDataMap = getBatchStockData(Array.from(goodsDataMap.keys()), tokens);
    
    // ステップ3: 結合処理
    const inventoryDataMap = new Map();
    for (const [goodsCode, goodsData] of goodsDataMap) {
      const stockData = stockDataMap.get(goodsCode);
      const completeInventoryData = {
        goods_id: goodsData.goods_id,
        goods_name: goodsData.goods_name,
        stock_quantity: stockData ? parseInt(stockData.stock_quantity) || 0 : parseInt(goodsData.stock_quantity) || 0,
        stock_allocated_quantity: stockData ? parseInt(stockData.stock_allocation_quantity) || 0 : 0,
        stock_free_quantity: stockData ? parseInt(stockData.stock_free_quantity) || 0 : 0
      };
      inventoryDataMap.set(goodsCode, completeInventoryData);
    }
    
    console.log(`  結合完了: ${inventoryDataMap.size}件`);
    return inventoryDataMap;
    
  } catch (error) {
    console.error(`  二重API処理エラー: ${error.message}`);
    return new Map();
  }
}

// 単一API版: 在庫マスタAPIのみで取得（在庫情報のみ）
function fetchInventoryWithSingleAPI(goodsCodes, tokens) {
  try {
    console.log(`  在庫マスタAPI単体呼び出し: ${goodsCodes ? goodsCodes.length : 'undefined'}件`);
    
    if (!goodsCodes || !Array.isArray(goodsCodes)) {
      console.error(`  エラー: goodsCodesが無効です`);
      return new Map();
    }
    
    // 在庫マスタAPIを使用（既存のgetBatchStockDataと同じエンドポイント）
    const url = `${NE_API_URL}/api_v1_master_stock/search`;
    const goodsIdCondition = goodsCodes.join(',');
    
    const payload = {
      'access_token': tokens.accessToken,
      'refresh_token': tokens.refreshToken,
      'stock_goods_id-in': goodsIdCondition, // 在庫マスタAPIの正しいフィールド名
      'fields': 'stock_goods_id,stock_quantity,stock_allocation_quantity,stock_free_quantity,stock_defective_quantity,stock_remaining_order_quantity,stock_out_quantity,stock_advance_order_quantity,stock_advance_order_allocation_quantity,stock_advance_order_free_quantity',
      'limit': MAX_ITEMS_PER_CALL.toString()
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
      tokens.accessToken = responseData.access_token;
      tokens.refreshToken = responseData.refresh_token;
    }

    const inventoryDataMap = new Map();
    
    if (responseData.result === 'success' && responseData.data && Array.isArray(responseData.data)) {
      console.log(`    API応答: ${responseData.data.length}件取得`);
      
      responseData.data.forEach(stockData => {
        inventoryDataMap.set(stockData.stock_goods_id, {
          goods_id: stockData.stock_goods_id,
          goods_name: '', // 単一API版では商品名は取得しない
          stock_quantity: parseInt(stockData.stock_quantity) || 0,
          stock_allocated_quantity: parseInt(stockData.stock_allocation_quantity) || 0,
          stock_free_quantity: parseInt(stockData.stock_free_quantity) || 0,
          stock_defective_quantity: parseInt(stockData.stock_defective_quantity) || 0,
          stock_advance_order_quantity: parseInt(stockData.stock_advance_order_quantity) || 0,
          stock_advance_order_allocation_quantity: parseInt(stockData.stock_advance_order_allocation_quantity) || 0,
          stock_advance_order_free_quantity: parseInt(stockData.stock_advance_order_free_quantity) || 0,
          stock_remaining_order_quantity: parseInt(stockData.stock_remaining_order_quantity) || 0,
          stock_out_quantity: parseInt(stockData.stock_out_quantity) || 0
        });
      });
    } else {
      // エラー詳細をログ出力
      console.log(`    API応答詳細:`);
      console.log(`      result: ${responseData.result || 'undefined'}`);
      console.log(`      data: ${responseData.data ? 'exists' : 'undefined'}`);
      console.log(`      count: ${responseData.count || 'undefined'}`);
      if (responseData.message) {
        console.error(`    在庫マスタAPI エラー: ${responseData.message}`);
      }
      if (responseData.data && !Array.isArray(responseData.data)) {
        console.error(`    データ形式エラー: dataが配列ではありません`);
      }
    }
    
    return inventoryDataMap;
    
  } catch (error) {
    console.error(`  単一API処理エラー: ${error.message}`);
    return new Map();
  }
}

/**
 * メイン関数：一括処理による在庫情報更新
 */
function updateInventoryDataBatch() {

  try {
    console.log('=== 在庫情報一括更新開始 ===');
    const startTime = new Date();
    // スプレッドシートを取得
    const { SPREADSHEET_ID, SHEET_NAME } = getSpreadsheetConfig();
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
    const errorDetails = []; // ★エラー詳細を収集
    
    for (let i = 0; i < goodsCodeList.length; i += MAX_ITEMS_PER_CALL) {
      const batch = goodsCodeList.slice(i, i + MAX_ITEMS_PER_CALL);
      console.log(`\n--- バッチ ${Math.floor(i / MAX_ITEMS_PER_CALL) + 1}: ${batch.length}件 ---`);
      try {
        // バッチで在庫情報を取得
        const inventoryDataMap = getBatchInventoryData(batch, tokens);
        // スプレッドシートを更新
        for (const goodsCode of batch) {
          const inventoryData = inventoryDataMap.get(goodsCode);
          const rowIndex = rowIndexMap.get(goodsCode);
          if (inventoryData && rowIndex) {
            try {
              updateRowWithInventoryData(sheet, rowIndex, inventoryData);
              totalUpdated++;
              console.log(` ✓ ${goodsCode}: 更新完了`);
            } catch (error) {
              // ★個別更新エラーの詳細を記録
              const errorInfo = {
                goodsCode: goodsCode,
                errorType: '更新エラー',
                errorMessage: error.message,
                timestamp: new Date(),
                batchNumber: Math.floor(i / MAX_ITEMS_PER_CALL) + 1
              };
              errorDetails.push(errorInfo);
              console.error(` ✗ ${goodsCode}: 更新エラー - ${error.message}`);
              totalErrors++;
            }
          } else {
            // ★データなしの場合も記録
            const errorInfo = {
              goodsCode: goodsCode,
              errorType: 'データなし',
              errorMessage: inventoryData ? 'rowIndex not found' : 'inventory data not found',
              timestamp: new Date(),
              batchNumber: Math.floor(i / MAX_ITEMS_PER_CALL) + 1
            };
            errorDetails.push(errorInfo);
            console.log(` - ${goodsCode}: データなし`);
          }
        }
        // バッチ間の待機（APIレート制限対策）
        if (i + MAX_ITEMS_PER_CALL < goodsCodeList.length) {
          console.log(`次のバッチまで ${API_WAIT_TIME}ms 待機...`);
          Utilities.sleep(API_WAIT_TIME);
        }
      } catch (error) {
        // ★バッチ全体のエラーを記録
        batch.forEach(goodsCode => {
          const errorInfo = {
            goodsCode: goodsCode,
            errorType: 'バッチエラー',
            errorMessage: error.message,
            timestamp: new Date(),
            batchNumber: Math.floor(i / MAX_ITEMS_PER_CALL) + 1
          };
          errorDetails.push(errorInfo);
        });
        console.error(`バッチ処理エラー:`, error.message);
        totalErrors += batch.length;
      }
    }
    
    const endTime = new Date();
    const duration = (endTime - startTime) / 1000;
    
    // ★エラーレポートの作成
    if (errorDetails.length > 0) {
      logErrorsToSheet(errorDetails);
      console.log(`\n--- エラーレポート ---`);
      console.log(`エラーレポートをシートに記録しました: ${errorDetails.length}件`);
    }
    
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
/**
 * バッチで在庫情報を取得（在庫APIのみ版）
 * @param {string[]} goodsCodeList - 商品コードの配列
 * @param {Object} tokens - アクセストークンとリフレッシュトークン
 * @returns {Map<string, Object>} 商品コード → 在庫情報のマップ
 */
function getBatchInventoryData(goodsCodeList, tokens) {
  const inventoryDataMap = new Map();

  try {
    console.log(`  在庫マスタ一括検索: ${goodsCodeList.length}件`);
    
    // 在庫マスタAPIのみで在庫情報を直接取得
    const stockDataMap = getBatchStockData(goodsCodeList, tokens);
    console.log(`  在庫マスタ取得完了: ${stockDataMap.size}件`);

    if (stockDataMap.size === 0) {
      console.log('  在庫データが見つかりませんでした');
      return inventoryDataMap;
    }

    // 在庫情報のみでデータを構築（商品名は空文字）
    for (const [goodsCode, stockData] of stockDataMap) {
      const inventoryData = {
        goods_id: stockData.stock_goods_id,
        goods_name: '', // 商品名は更新しない（空文字で統一）
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
      inventoryDataMap.set(goodsCode, inventoryData);
    }

    console.log(`  在庫情報構築完了: ${inventoryDataMap.size}件`);
    return inventoryDataMap;

  } catch (error) {
    console.error(`在庫情報取得エラー:`, error.message);
    return inventoryDataMap;
  }
}

/**
 * 複数商品の基本情報を一括取得
 * @param {string[]} goodsCodeList - 商品コードの配列
 * @param {Object} tokens - トークン情報
 * @returns {Map<string, Object>} 商品コード → 商品情報のマップ
 */
function getBatchGoodsData(goodsCodeList, tokens) {
  const url = `${NE_API_URL}/api_v1_master_goods/search`;
  
  // 複数の商品IDを検索条件に設定
  const goodsIdCondition = goodsCodeList.join(',');
  
  const payload = {
    'access_token': tokens.accessToken,
    'refresh_token': tokens.refreshToken,
    'goods_id-in': goodsIdCondition, // IN条件で複数商品を一括検索
    'fields': 'goods_id,goods_name,stock_quantity',
    'limit': MAX_ITEMS_PER_CALL.toString() // 取得件数制限
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
      console.log(`  API応答: ${responseData.data.length}件取得`);
    } else {
      console.error(`  商品マスタAPI エラー:`, responseData.message || 'Unknown error');
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
function getBatchStockData(goodsCodeList, tokens) {
  const url = `${NE_API_URL}/api_v1_master_stock/search`;
  
  // 複数の商品IDを検索条件に設定
  const goodsIdCondition = goodsCodeList.join(',');
  
  const payload = {
    'access_token': tokens.accessToken,
    'refresh_token': tokens.refreshToken,
    'stock_goods_id-in': goodsIdCondition, // IN条件で複数商品の在庫を一括検索
    'fields': 'stock_goods_id,stock_quantity,stock_allocation_quantity,stock_defective_quantity,stock_remaining_order_quantity,stock_out_quantity,stock_free_quantity,stock_advance_order_quantity,stock_advance_order_allocation_quantity,stock_advance_order_free_quantity',
    'limit': MAX_ITEMS_PER_CALL.toString()
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
      console.log(`  API応答: ${responseData.data.length}件取得`);
    } else {
      console.error(`  在庫マスタAPI エラー:`, responseData.message || 'Unknown error');
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
  console.log('トークンを更新しました');
}

/**
 * テスト用：小規模バッチでの動作確認
 * @param {number} maxItems - テスト対象の最大商品数（デフォルト: 10）
 */
function testBatchProcessing(maxItems = 10) {

  try {
    console.log(`=== バッチ処理テスト（最大${maxItems}件） ===`);
    
    // スプレッドシートから商品コードを取得
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
  const { SPREADSHEET_ID, SHEET_NAME } = getSpreadsheetConfig();
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
 * 現在のスクリプトプロパティ設定を表示
 */
function showCurrentProperties() {
  const properties = PropertiesService.getScriptProperties();
  console.log('=== 現在のスクリプトプロパティ設定 ===');
  console.log(`SPREADSHEET_ID: ${properties.getProperty('SPREADSHEET_ID') || '未設定'}`);
  console.log(`SHEET_NAME: ${properties.getProperty('SHEET_NAME') || '未設定'}`);
  console.log(`BATCH_SIZE: ${properties.getProperty('BATCH_SIZE') || '未設定'}`);
  console.log(`API_WAIT_TIME: ${properties.getProperty('API_WAIT_TIME') || '未設定'}`);
  console.log('');
  console.log('認証情報:');
  console.log(`ACCESS_TOKEN: ${properties.getProperty('ACCESS_TOKEN') ? '設定済み' : '未設定'}`);
  console.log(`REFRESH_TOKEN: ${properties.getProperty('REFRESH_TOKEN') ? '設定済み' : '未設定'}`);
}

/**
 * エラー詳細をスプレッドシートに記録
 * @param {Array} errorDetails - エラー詳細の配列
 */
function logErrorsToSheet(errorDetails) {
  try {
    const { SPREADSHEET_ID, SHEET_NAME } = getSpreadsheetConfig();
    const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
    let errorSheet = spreadsheet.getSheetByName('エラーログ');
    
    // エラーログシートが存在しない場合は作成
    if (!errorSheet) {
      errorSheet = spreadsheet.insertSheet('エラーログ');
      // ヘッダー行を設定
      const headers = [
        '発生日時', '商品コード', 'エラー種別', 
        'エラー内容', 'バッチ番号', '処理日時'
      ];
      errorSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      errorSheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
    }
    
    // エラーデータを準備
    const errorRows = errorDetails.map(error => [
      error.timestamp,
      error.goodsCode,
      error.errorType,
      error.errorMessage,
      error.batchNumber,
      new Date()
    ]);
    
    // データを追加
    if (errorRows.length > 0) {
      const lastRow = errorSheet.getLastRow();
      const range = errorSheet.getRange(lastRow + 1, 1, errorRows.length, 6);
      range.setValues(errorRows);
      
      // 日時列のフォーマット設定
      errorSheet.getRange(lastRow + 1, 1, errorRows.length, 1)
                .setNumberFormat('yyyy/mm/dd hh:mm:ss');
      errorSheet.getRange(lastRow + 1, 6, errorRows.length, 1)
                .setNumberFormat('yyyy/mm/dd hh:mm:ss');
    }
    
    console.log(`エラーログに${errorRows.length}件を記録しました`);
    
  } catch (error) {
    console.error('エラーログ記録中にエラーが発生:', error.message);
    // エラーログの記録に失敗してもメイン処理は継続
  }
}

/**
 * 使用方法ガイド
 */
function showUsageGuide() {
  console.log('=== 在庫情報取得スクリプト 使用方法ガイド ===');
  console.log('');
  console.log('【主要関数】');
  console.log('1. compareAPIVersions(件数)');
  console.log('   - 二重API版 vs 単一API版の性能比較');
  console.log('   - 例: compareAPIVersions(10)');
  console.log('');
  console.log('2. updateInventoryDataBatch()');
  console.log('   - 全商品の在庫情報を一括処理で更新');
  console.log('   - 1000件ずつのバッチで自動分割処理');
  console.log('   - 従来版より大幅に高速化');
  console.log('');
  console.log('3. testBatchProcessing(件数)');
  console.log('   - 小規模テスト用（デフォルト10件）');
  console.log('   - 例: testBatchProcessing(5)');
  console.log('');
  console.log('4. comparePerformance(件数)');
  console.log('   - 従来版との性能比較テスト');
  console.log('   - 例: comparePerformance(20)');
  console.log('');
  console.log('【期待される改善効果】');
  console.log('- APIコール数: 大幅削減（1000件を3回で処理）');
  console.log('- 処理速度: 10～50倍の高速化');
  console.log('- 実行時間制限: 数千件でも制限内で完了');
  console.log('- API制限: レート制限に引っかかりにくい');
  console.log('');
  console.log('【推奨実行手順】');
  console.log('1. compareAPIVersions(10) でAPI版本比較');
  console.log('2. testBatchProcessing(10) で動作確認');
  console.log('3. comparePerformance(20) で性能確認');
  console.log('4. updateInventoryDataBatch() で全件更新');
  console.log('');
  console.log('【設定変更可能項目】');
  console.log('- MAX_ITEMS_PER_CALL: バッチサイズ（現在1000件）');
  console.log('- API_WAIT_TIME: API間隔（現在500ms）');
}