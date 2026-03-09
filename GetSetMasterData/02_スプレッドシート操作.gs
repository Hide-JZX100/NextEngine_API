/**
 * =============================================================================
 * スプレッドシート操作
 * =============================================================================
 * セット商品マスタデータをスプレッドシートに書き込む関数群
 * 
 * 【主な機能】
 * - ヘッダー行の作成・確認
 * - データの全件上書き(ヘッダーは保持)
 * - バッチ書き込みによる高速化
 * =============================================================================
 */

/**
 * ヘッダー定義
 * スプレッドシートの列順序とAPIフィールドの対応を定義
 * 
 * @return {Array} ヘッダー情報の配列
 */
function getHeaderDefinition() {
  return [
    { header: 'set_syohin_code', field: 'set_goods_id' },
    { header: 'set_syohin_name', field: 'set_goods_name' },
    { header: 'set_baika_tnk', field: 'set_goods_selling_price' },
    { header: 'syohin_code', field: 'set_goods_detail_goods_id' },
    { header: 'suryo', field: 'set_goods_detail_quantity' }
  ];
}

/**
 * ヘッダー行を取得
 * 
 * @return {Array} ヘッダー行の配列
 */
function getHeaderRow() {
  const headers = getHeaderDefinition();
  return headers.map(h => h.header);
}

/**
 * シートを初期化
 * ヘッダー行のみの状態にする
 * 
 * @param {Sheet} sheet - 対象シート
 */
function initializeSheet(sheet) {
  // 全データをクリア
  sheet.clear();
  
  // ヘッダー行を設定
  const headerRow = getHeaderRow();
  sheet.getRange(1, 1, 1, headerRow.length).setValues([headerRow]);
  
  // ヘッダー行の書式設定
  const headerRange = sheet.getRange(1, 1, 1, headerRow.length);
  headerRange.setFontWeight('bold');
  headerRange.setBackground('#f3f3f3');
  headerRange.setHorizontalAlignment('center');
  
  console.log('シートを初期化しました(ヘッダー行のみ)');
}

/**
 * ヘッダーが正しく設定されているか確認
 * 
 * @param {Sheet} sheet - 対象シート
 * @return {boolean} ヘッダーが正しければtrue
 */
function validateHeader(sheet) {
  const lastRow = sheet.getLastRow();
  
  // シートが空の場合
  if (lastRow === 0) {
    console.log('シートが空です。ヘッダーを初期化します。');
    initializeSheet(sheet);
    return true;
  }
  
  // ヘッダー行を取得
  const expectedHeaders = getHeaderRow();
  const actualHeaders = sheet.getRange(1, 1, 1, expectedHeaders.length).getValues()[0];
  
  // ヘッダーの比較
  for (let i = 0; i < expectedHeaders.length; i++) {
    if (actualHeaders[i] !== expectedHeaders[i]) {
      console.log('ヘッダーが不一致です。');
      console.log('期待値:', expectedHeaders);
      console.log('実際の値:', actualHeaders);
      return false;
    }
  }
  
  console.log('ヘッダーの検証OK');
  return true;
}

/**
 * APIデータをスプレッドシート行形式に変換
 * 
 * @param {Array} apiData - APIから取得したデータ配列
 * @return {Array} スプレッドシート行データの2次元配列
 */
function convertApiDataToRows(apiData) {
  const headers = getHeaderDefinition();
  const rows = [];
  
  for (const record of apiData) {
    const row = headers.map(h => {
      const value = record[h.field];
      // 未定義の場合は空文字列
      return value !== undefined ? value : '';
    });
    rows.push(row);
  }
  
  return rows;
}

/**
 * データをスプレッドシートに書き込み
 * 既存データをクリアして全件上書き(ヘッダーは保持)
 * 
 * @param {Array} apiData - APIから取得したデータ配列
 * @return {Object} 書き込み結果 {success: boolean, rowCount: number, message: string}
 */
function writeDataToSheet(apiData) {
  try {
    console.log('=== データ書き込み開始 ===');
    
    // 設定取得
    const config = getSpreadsheetConfig();
    const spreadsheet = SpreadsheetApp.openById(config.spreadsheetId);
    const sheet = spreadsheet.getSheetByName(config.sheetName);
    
    if (!sheet) {
      throw new Error(`シート "${config.sheetName}" が見つかりません。`);
    }
    
    // ヘッダーの確認
    if (!validateHeader(sheet)) {
      console.log('ヘッダーを再初期化します。');
      initializeSheet(sheet);
    }
    
    // データが空の場合
    if (!apiData || apiData.length === 0) {
      console.log('書き込むデータがありません。');
      // データ行のみクリア(ヘッダーは保持)
      const lastRow = sheet.getLastRow();
      if (lastRow > 1) {
        sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).clear();
      }
      return {
        success: true,
        rowCount: 0,
        message: 'データが空のため、既存データをクリアしました。'
      };
    }
    
    // データ変換
    const rows = convertApiDataToRows(apiData);
    console.log('変換後の行数:', rows.length);
    
    // 既存データのクリア(ヘッダーは保持)
    const lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).clear();
      console.log('既存データをクリアしました');
    }
    
    // データ書き込み(バッチ処理で高速化)
    const startRow = 2; // ヘッダーの次の行から
    const numRows = rows.length;
    const numCols = rows[0].length;
    
    sheet.getRange(startRow, 1, numRows, numCols).setValues(rows);
    
    console.log('✅ データ書き込み完了');
    console.log('書き込み行数:', numRows);
    
    return {
      success: true,
      rowCount: numRows,
      message: `${numRows}行のデータを書き込みました。`
    };
    
  } catch (error) {
    console.error('❌ データ書き込みエラー:', error.message);
    return {
      success: false,
      rowCount: 0,
      message: error.message
    };
  }
}

/**
 * スプレッドシート書き込みのテスト
 * サンプルデータで動作確認
 */
function testWriteDataToSheet() {
  console.log('=== スプレッドシート書き込みテスト ===');
  
  // サンプルデータ(APIレスポンス形式)
  const sampleData = [
    {
      "set_goods_id": "0002-JN3400-85-DB",
      "set_goods_name": "８５−ＪＮ３４００／ＤＢ",
      "set_goods_selling_price": "0",
      "set_goods_detail_goods_id": "85-JN3400-FF-DB",
      "set_goods_detail_quantity": "1"
    },
    {
      "set_goods_id": "0002-JN3400-85-DB",
      "set_goods_name": "８５−ＪＮ３４００／ＤＢ",
      "set_goods_selling_price": "0",
      "set_goods_detail_goods_id": "85-JN3400-PA-DB-dm",
      "set_goods_detail_quantity": "1"
    },
    {
      "set_goods_id": "0002-JN3400-85-DB",
      "set_goods_name": "８５−ＪＮ３４００／ＤＢ",
      "set_goods_selling_price": "0",
      "set_goods_detail_goods_id": "85-JN34All-PA",
      "set_goods_detail_quantity": "1"
    }
  ];
  
  console.log('サンプルデータ件数:', sampleData.length);
  
  // 書き込み実行
  const result = writeDataToSheet(sampleData);
  
  console.log('');
  console.log('=== 書き込み結果 ===');
  console.log('成功:', result.success);
  console.log('行数:', result.rowCount);
  console.log('メッセージ:', result.message);
  
  if (result.success) {
    console.log('');
    console.log('✅ スプレッドシート書き込みテスト成功!');
    console.log('スプレッドシートを確認してください。');
  } else {
    throw new Error('書き込みテスト失敗: ' + result.message);
  }
}