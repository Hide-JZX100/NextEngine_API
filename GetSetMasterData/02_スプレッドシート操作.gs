/**
 * @fileoverview スプレッドシート操作モジュール
 * 
 * ネクストエンジンから取得したセット商品マスタデータを、Google スプレッドシートへ
 * 出力するための操作を担当します。
 * 
 * 主な機能:
 * - ヘッダー行（列名）の動的生成と整合性チェック
 * - シートの初期化および書式設定
 * - APIレスポンスからスプレッドシート用2次元配列へのデータ変換
 * - 大量データの高速バッチ書き込み（全件上書き）
 */

/**
 * ヘッダー定義
 * 
 * スプレッドシートの列名（ラベル）と、APIレスポンスのフィールド名のマッピングを定義します。
 * 列の並び順はこの配列の順序に従います。
 * 
 * @return {Array<{header: string, field: string}>} ヘッダー情報の配列
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
 * ヘッダー定義から、スプレッドシートの1行目に書き込むためのラベル配列を生成します。
 * 
 * @return {string[]} ヘッダーラベルの配列
 */
function getHeaderRow() {
  const headers = getHeaderDefinition();
  return headers.map(h => h.header);
}

/**
 * シートを初期化
 * 
 * 指定されたシートの全セルをクリアし、ヘッダー行を再作成します。
 * ヘッダーには背景色、太字、中央揃えの書式を適用します。
 * 
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - 操作対象のシートオブジェクト
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
 * シートの1行目を読み込み、定義されたヘッダーと一致するか検証します。
 * シートが空、またはヘッダーが異なる場合は必要に応じて初期化を促します。
 * 
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - 検証対象のシートオブジェクト
 * @return {boolean} ヘッダーが定義通りであれば true、不一致であれば false
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
 * ネクストエンジン API から取得したオブジェクトの配列を、
 * `setValues()` で使用可能な 2次元配列（行・列）に変換します。
 * 
 * @param {Object[]} apiData - APIから取得したレコードの配列
 * @return {any[][]} スプレッドシート書き込み用の2次元配列
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
 * 
 * 既存のデータ行（2行目以降）を一度削除し、新しいデータを一括で書き込みます。
 * シートが見つからない場合や書き込みに失敗した場合は、エラー情報を返します。
 * 
 * @param {Object[]} apiData - APIから取得したデータ配列
 * @return {{success: boolean, rowCount: number, message: string}} 処理結果オブジェクト
 * @throws {Error} スプレッドシートIDやシート名が不正な場合
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
 * 
 * 固定のサンプルデータ（セット商品3明細分）を使用して、
 * シートへの書き込み、ヘッダーの生成、書式設定が正常に動作するかを検証します。
 * @throws {Error} 書き込み処理が失敗した場合
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