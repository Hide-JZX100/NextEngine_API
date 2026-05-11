/**
 * @fileoverview スプレッドシート書き込みテスト（ステップ4-2）
 * 
 * 【目的】ネクストエンジンAPIから取得し、変換したデータをスプレッドシートに書き込む。
 * 
 * 【主な機能】
 * - シートの接続確認およびヘッダー行の自動生成
 * - 少量データ（3件）による書き込み位置・形式の検証
 * - 大量データ（一括）書き込みのパフォーマンス計測
 * 
 * 【前提条件】
 * - スクリプトプロパティ `SPREADSHEET_ID` および `SHEET_NAME` が設定されていること。
 */

/**
 * 💡 スプレッドシートの接続確認とヘッダー行の保証（ステップ4-2-1）
 * 
 * 実際にデータを書き込む前に、スプレッドシートと目的のシートが正しく準備されているかを確認します。
 * 
 * 機能:
 * 1. スクリプトプロパティから ID とシート名を取得し、接続を確認。
 * 2. 指定されたシートが存在しない場合は新規作成。
 * 3. 1行目に規定の14項目（出荷予定日, 伝票番号等）のヘッダーがあるか確認。
 * 4. ヘッダーがない場合、自動作成し、太字・背景色・中央揃えのスタイルを適用。
 * 
 * 重要な処理: 
 * データ書き込みの基盤となるシートを確実に用意し、データの整合性を保つためのヘッダー構造を保証します。
 * 
 * @returns {GoogleAppsScript.Spreadsheet.Sheet} 準備が整ったシートオブジェクト
 * 
 * 【テストの目的】
 * 1. スプレッドシートが正しく開けるか確認
 * 2. 指定したシートが存在するか確認
 * 3. ヘッダー行が存在するか確認、なければ作成
 */
function testPrepareSheet() {
  try {
    console.log('=== スプレッドシート準備テスト開始 ===');
    console.log('');
    
    // スクリプトプロパティから設定を取得
    const properties = PropertiesService.getScriptProperties();
    const spreadsheetId = properties.getProperty('SPREADSHEET_ID');
    const sheetName = properties.getProperty('SHEET_NAME');
    
    if (!spreadsheetId || !sheetName) {
      throw new Error('スクリプトプロパティが設定されていません。SPREADSHEET_ID と SHEET_NAME を設定してください。');
    }
    
    console.log(`スプレッドシートID: ${spreadsheetId}`);
    console.log(`シート名: ${sheetName}`);
    console.log('');
    
    // スプレッドシートを開く
    const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    console.log(`スプレッドシート名: ${spreadsheet.getName()}`);
    console.log('スプレッドシートを開きました');
    console.log('');
    
    // シートを取得または作成
    let sheet = spreadsheet.getSheetByName(sheetName);
    
    if (!sheet) {
      console.log(`シート「${sheetName}」が存在しないため、新規作成します`);
      sheet = spreadsheet.insertSheet(sheetName);
      console.log('シートを作成しました');
    } else {
      console.log(`シート「${sheetName}」を取得しました`);
    }
    console.log('');
    
    // ヘッダー行の定義（Shipping_piece.csvの日本語ヘッダー）
    const headers = [
      '出荷予定日',
      '伝票番号',
      '商品コード',
      '商品名',
      '受注数',
      '引当数',
      '奥行き（cm）',
      '幅（cm）',
      '高さ（cm）',
      '発送方法コード',
      '発送方法',
      '重さ（g）',
      '受注状態区分',
      '送り先住所1'
    ];
    
    // 既存のヘッダー行を確認
    const existingHeaders = sheet.getRange(1, 1, 1, headers.length).getValues()[0];
    const hasHeaders = existingHeaders.some(cell => cell !== '');
    
    if (!hasHeaders) {
      console.log('ヘッダー行が存在しないため、作成します');
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      
      // ヘッダー行のスタイルを設定
      const headerRange = sheet.getRange(1, 1, 1, headers.length);
      headerRange.setFontWeight('bold');
      headerRange.setBackground('#f3f3f3');
      headerRange.setHorizontalAlignment('center');
      
      console.log('ヘッダー行を作成しました');
    } else {
      console.log('ヘッダー行は既に存在します');
      console.log('既存のヘッダー:', existingHeaders.join(', '));
    }
    console.log('');
    
    // シートの状態を確認
    const lastRow = sheet.getLastRow();
    const lastColumn = sheet.getLastColumn();
    
    console.log('=== シートの状態 ===');
    console.log(`最終行: ${lastRow}`);
    console.log(`最終列: ${lastColumn}`);
    console.log(`データ行数: ${Math.max(0, lastRow - 1)}行`);
    console.log('');
    
    console.log('=== スプレッドシート準備完了 ===');
    console.log('');
    console.log('=== 次のステップ ===');
    console.log('1. スプレッドシートを開いて、ヘッダー行が正しく作成されているか確認してください');
    console.log('2. 確認できたら、次は少量データ（3件）の書き込みテストに進みます');
    
    return sheet;
    
  } catch (error) {
    console.error('スプレッドシート準備エラー:', error.message);
    console.error('エラー詳細:', error);
    throw error;
  }
}

/**
 * 💡 少量のデータで書き込みロジックを検証（ステップ4-2-2）
 * 
 * APIから取得したデータの最初の3件を抽出し、スプレッドシートの最終行の次から追記できるかを確認します。
 * 
 * 機能:
 * 1. `fetchAllShippingData` からデータを取得。
 * 2. 最初の3件のみを `convertApiDataToSheetRow` で変換。
 * 3. `sheet.getLastRow()` を使用して、既存データの末尾から追記。
 * 
 * 重要な処理: 
 * 少数データでのテストを通じて、書き込みの開始位置やデータ形式に関するミスがないかを素早く検証します。
 * 
 * 【テストの目的】
 * 1. データが正しく書き込まれるか確認
 * 2. 書き込み位置が正しいか確認
 * 3. データ形式が正しいか確認
 */
function testWriteThreeRows() {
  try {
    console.log('=== 3件データ書き込みテスト開始 ===');
    console.log('');
    
    // スクリプトプロパティから設定を取得
    const properties = PropertiesService.getScriptProperties();
    const spreadsheetId = properties.getProperty('SPREADSHEET_ID');
    const sheetName = properties.getProperty('SHEET_NAME');
    
    // スプレッドシートとシートを取得
    const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    let sheet = spreadsheet.getSheetByName(sheetName);
    
    if (!sheet) {
      console.log('シートが存在しないため、準備します');
      sheet = testPrepareSheet();
    }
    
    console.log('データ取得中...');
    
    // APIから全データを取得
    const apiData = fetchAllShippingData('2025-10-03', '2025-10-05');
    console.log(`取得データ件数: ${apiData.length}件`);
    console.log('');
    
    // 最初の3件だけ変換
    console.log('データ変換中（3件のみ）...');
    const convertedData = [];
    for (let i = 0; i < Math.min(3, apiData.length); i++) {
      const row = convertApiDataToSheetRow(apiData[i]);
      convertedData.push(row);
    }
    console.log(`変換完了: ${convertedData.length}件`);
    console.log('');
    
    // 既存のデータ行数を確認
    const lastRow = sheet.getLastRow();
    console.log(`現在のデータ行数: ${Math.max(0, lastRow - 1)}行`);
    
    // 書き込み開始行（ヘッダーの次の行から）
    const startRow = lastRow + 1;
    console.log(`書き込み開始行: ${startRow}行目`);
    console.log('');
    
    // データを書き込む
    console.log('スプレッドシートに書き込み中...');
    const range = sheet.getRange(startRow, 1, convertedData.length, convertedData[0].length);
    range.setValues(convertedData);
    
    console.log('');
    console.log('=== 3件データ書き込み成功 ===');
    console.log(`書き込み位置: ${startRow}行目 ～ ${startRow + convertedData.length - 1}行目`);
    console.log(`書き込み件数: ${convertedData.length}件`);
    console.log('');
    
    // 書き込んだデータの確認
    console.log('--- 書き込んだデータ ---');
    for (let i = 0; i < convertedData.length; i++) {
      console.log(`${i + 1}件目:`);
      console.log(`  出荷予定日: ${convertedData[i][0]}`);
      console.log(`  伝票番号: ${convertedData[i][1]}`);
      console.log(`  商品コード: ${convertedData[i][2]}`);
      console.log(`  商品名: ${convertedData[i][3].substring(0, 50)}...`);
    }
    
    console.log('');
    console.log('=== 次のステップ ===');
    console.log('1. スプレッドシートを開いて、3件のデータが正しく書き込まれているか確認してください');
    console.log('2. データの位置、形式、内容が正しいか確認してください');
    console.log('3. 問題なければ、全132件の書き込みに進みます');
    
  } catch (error) {
    console.error('3件データ書き込みエラー:', error.message);
    console.error('エラー詳細:', error);
    throw error;
  }
}

/**
 * 💡 全データの一括書き込みとパフォーマンス計測（ステップ4-2-3）
 * 
 * 大量データ（テスト時は132件程度を想定）を一括で書き込み、処理速度と正確性を検証します。
 * 
 * 機能:
 * 1. 処理開始時刻を記録。
 * 2. `fetchAllShippingData` で全データを取得・変換。
 * 3. ヘッダーを除く既存データを一括クリア（2行目以降）。
 * 4. `range.setValues()` を使用し、2行目から変換済みデータを一括書き込み。
 * 5. 処理時間、書き込み件数、シートの最終行数をログ出力。
 * 
 * 重要な処理: 
 * GASの制限（実行時間）を意識し、一括書き込み（setValues）によるパフォーマンスの良さと、
 * 全データが過不足なく書き込まれているかを確認します。
 * 
 * 【テストの目的】
 * 1. 全132件を一括で書き込む
 * 2. パフォーマンスを確認
 * 3. 書き込み完了後のシート状態を確認
 */
function testWriteAllRows() {
  try {
    console.log('=== 全件データ書き込み開始 ===');
    console.log('');
    
    const startTime = new Date();
    
    // スクリプトプロパティから設定を取得
    const properties = PropertiesService.getScriptProperties();
    const spreadsheetId = properties.getProperty('SPREADSHEET_ID');
    const sheetName = properties.getProperty('SHEET_NAME');
    
    // スプレッドシートとシートを取得
    const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    let sheet = spreadsheet.getSheetByName(sheetName);
    
    if (!sheet) {
      console.log('シートが存在しないため、準備します');
      sheet = testPrepareSheet();
    }
    
    console.log('データ取得中...');
    
    // APIから全データを取得
    const apiData = fetchAllShippingData('2025-10-03', '2025-10-05');
    console.log(`取得データ件数: ${apiData.length}件`);
    console.log('');
    
    // 全データを変換
    console.log('データ変換中...');
    const convertedData = [];
    for (let i = 0; i < apiData.length; i++) {
      const row = convertApiDataToSheetRow(apiData[i]);
      convertedData.push(row);
      
      // 進捗表示（20件ごと）
      if ((i + 1) % 20 === 0 || i === apiData.length - 1) {
        console.log(`変換完了: ${i + 1}/${apiData.length}件`);
      }
    }
    console.log('');
    
    // 既存のデータをクリア（ヘッダー行は残す）
    const lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      console.log(`既存データをクリア中... (${lastRow - 1}行)`);
      sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).clearContent();
      console.log('既存データをクリアしました');
      console.log('');
    }
    
    // 書き込み開始行（ヘッダーの次の行から）
    const startRow = 2;
    console.log(`書き込み開始行: ${startRow}行目`);
    console.log('');
    
    // データを一括書き込み
    console.log('スプレッドシートに一括書き込み中...');
    const range = sheet.getRange(startRow, 1, convertedData.length, convertedData[0].length);
    range.setValues(convertedData);
    
    const endTime = new Date();
    const elapsedTime = (endTime - startTime) / 1000; // 秒に変換
    
    console.log('');
    console.log('=== 全件データ書き込み成功 ===');
    console.log(`書き込み位置: ${startRow}行目 ～ ${startRow + convertedData.length - 1}行目`);
    console.log(`書き込み件数: ${convertedData.length}件`);
    console.log(`処理時間: ${elapsedTime.toFixed(2)}秒`);
    console.log('');
    
    // シートの最終状態
    console.log('=== シートの最終状態 ===');
    console.log(`総行数: ${sheet.getLastRow()}行（ヘッダー含む）`);
    console.log(`データ行数: ${sheet.getLastRow() - 1}行`);
    console.log('');
    
    console.log('=== 次のステップ ===');
    console.log('1. スプレッドシートを開いて、全データが正しく書き込まれているか確認してください');
    console.log('2. データの件数、内容が正しいか確認してください');
    console.log('3. これでステップ4-2（スプレッドシート書き込み）は完了です！');
    
  } catch (error) {
    console.error('全件データ書き込みエラー:', error.message);
    console.error('エラー詳細:', error);
    throw error;
  }
}