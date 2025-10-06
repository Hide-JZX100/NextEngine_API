/**
=============================================================================
本番用: ネクストエンジン出荷予定データ更新
=============================================================================

* 【目的】
* ネクストエンジンから出荷予定データを取得し、スプレッドシートを更新する
* 
* 【機能】
* - 指定期間の出荷予定データを取得
* - スプレッドシートに一括書き込み
* - 実行ログの記録
* 
* 【使用方法】
* updateShippingData() - デフォルト（本日含む未来3日分）
* updateShippingData('2025-10-03', '2025-10-05') - 期間指定
* 
=============================================================================
*/

/**
 * 本番用: 出荷予定データを更新
 * 
 * @param {string} startDate - 開始日（YYYY-MM-DD形式）省略時は本日
 * @param {string} endDate - 終了日（YYYY-MM-DD形式）省略時は本日+2日
 */
function updateShippingData(startDate, endDate) {
  try {
    // ========================================
    // 1. 初期設定
    // ========================================
    console.log('');
    console.log('========================================');
    console.log('  出荷予定データ更新開始');
    console.log('========================================');
    console.log('');
    
    const startTime = new Date();
    
    // 日付が指定されていない場合は、本日から3日分をデフォルトとする
    if (!startDate) {
      const today = new Date();
      startDate = Utilities.formatDate(today, 'Asia/Tokyo', 'yyyy-MM-dd');
    }
    
    if (!endDate) {
      const today = new Date();
      const threeDaysLater = new Date(today.getTime() + (2 * 24 * 60 * 60 * 1000)); // 本日+2日
      endDate = Utilities.formatDate(threeDaysLater, 'Asia/Tokyo', 'yyyy-MM-dd');
    }
    
    console.log(`取得期間: ${startDate} ～ ${endDate}`);
    console.log('');
    
    // ========================================
    // 2. データ取得
    // ========================================
    console.log('--- ステップ1: データ取得 ---');
    console.log('ネクストエンジンAPIからデータを取得中...');
    
    const apiData = fetchAllShippingData(startDate, endDate);
    
    console.log(`✓ 取得完了: ${apiData.length}件`);
    console.log('');
    
    // ========================================
    // 3. データ変換
    // ========================================
    console.log('--- ステップ2: データ変換 ---');
    console.log('スプレッドシート形式に変換中...');
    
    const convertedData = [];
    for (let i = 0; i < apiData.length; i++) {
      const row = convertApiDataToSheetRow(apiData[i]);
      convertedData.push(row);
      
      // 進捗表示（50件ごと）
      if ((i + 1) % 50 === 0 || i === apiData.length - 1) {
        console.log(`  変換中: ${i + 1}/${apiData.length}件`);
      }
    }
    
    console.log(`✓ 変換完了: ${convertedData.length}件`);
    console.log('');
    
    // ========================================
    // 4. スプレッドシート書き込み
    // ========================================
    console.log('--- ステップ3: スプレッドシート書き込み ---');
    
    // スクリプトプロパティから設定を取得
    const properties = PropertiesService.getScriptProperties();
    const spreadsheetId = properties.getProperty('SPREADSHEET_ID');
    const sheetName = properties.getProperty('SHEET_NAME');
    
    if (!spreadsheetId || !sheetName) {
      throw new Error('スクリプトプロパティが設定されていません');
    }
    
    // スプレッドシートとシートを取得
    const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    let sheet = spreadsheet.getSheetByName(sheetName);
    
    // シートが存在しない場合は作成
    if (!sheet) {
      console.log(`シート「${sheetName}」が存在しないため、作成します`);
      sheet = spreadsheet.insertSheet(sheetName);
      
      // ヘッダー行を作成
      const headers = [
        '出荷予定日', '伝票番号', '商品コード', '商品名', '受注数', '引当数',
        '奥行き（cm）', '幅（cm）', '高さ（cm）', '発送方法コード', '発送方法',
        '重さ（g）', '受注状態区分', '送り先住所1'
      ];
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      
      // ヘッダー行のスタイル設定
      const headerRange = sheet.getRange(1, 1, 1, headers.length);
      headerRange.setFontWeight('bold');
      headerRange.setBackground('#f3f3f3');
      headerRange.setHorizontalAlignment('center');
    }
    
    // 既存データ行を完全に削除（ヘッダー行は残す）
    const lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      console.log(`既存データ行を削除: ${lastRow - 1}行`);
      sheet.deleteRows(2, lastRow - 1); // 2行目から最終行まで物理的に削除
    }


    // データを書き込み
    if (convertedData.length > 0) {
      console.log('データを書き込み中...');
      const range = sheet.getRange(2, 1, convertedData.length, convertedData[0].length);
      range.setValues(convertedData);
      console.log(`✓ 書き込み完了: ${convertedData.length}件`);
    } else {
      console.log('書き込むデータがありません');
    }
    
    console.log('');
    
    // ========================================
    // 5. 処理完了
    // ========================================
    const endTime = new Date();
    const elapsedTime = (endTime - startTime) / 1000; // 秒
    
    console.log('========================================');
    console.log('  処理完了');
    console.log('========================================');
    console.log(`取得期間: ${startDate} ～ ${endDate}`);
    console.log(`取得件数: ${apiData.length}件`);
    console.log(`処理時間: ${elapsedTime.toFixed(2)}秒`);
    console.log(`完了時刻: ${Utilities.formatDate(endTime, 'Asia/Tokyo', 'yyyy-MM-dd HH:mm:ss')}`);
    console.log('');
    
    // ========================================
    // 6. 実行ログを記録
    // ========================================
    logExecution(startDate, endDate, apiData.length, elapsedTime, 'success');
    
    return {
      success: true,
      startDate: startDate,
      endDate: endDate,
      recordCount: apiData.length,
      elapsedTime: elapsedTime
    };
    
  } catch (error) {
    console.error('');
    console.error('========================================');
    console.error('  エラーが発生しました');
    console.error('========================================');
    console.error('エラー内容:', error.message);
    console.error('エラー詳細:', error);
    console.error('');
    
    // エラーログを記録
    logExecution(startDate || 'N/A', endDate || 'N/A', 0, 0, 'error', error.message);
    
    throw error;
  }
}

/**
 * 実行ログを記録
 * 
 * @param {string} startDate - 開始日
 * @param {string} endDate - 終了日
 * @param {number} recordCount - 取得件数
 * @param {number} elapsedTime - 処理時間（秒）
 * @param {string} status - 実行ステータス（success/error）
 * @param {string} errorMessage - エラーメッセージ（エラー時のみ）
 */
function logExecution(startDate, endDate, recordCount, elapsedTime, status, errorMessage) {
  try {
    // スクリプトプロパティから設定を取得
    const properties = PropertiesService.getScriptProperties();
    const spreadsheetId = properties.getProperty('SPREADSHEET_ID');
    const logSheetName = properties.getProperty('LOG_SHEET_NAME') || '実行ログ';
    
    if (!spreadsheetId) {
      console.log('ログ記録: スプレッドシートIDが設定されていません');
      return;
    }
    
    // スプレッドシートとログシートを取得
    const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    let logSheet = spreadsheet.getSheetByName(logSheetName);
    
    // ログシートが存在しない場合は作成
    if (!logSheet) {
      logSheet = spreadsheet.insertSheet(logSheetName);
      
      // ヘッダー行を作成
      const headers = [
        '実行日時', '開始日', '終了日', '取得件数', 
        '処理時間（秒）', 'ステータス', 'エラー内容'
      ];
      logSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      
      // ヘッダー行のスタイル設定
      const headerRange = logSheet.getRange(1, 1, 1, headers.length);
      headerRange.setFontWeight('bold');
      headerRange.setBackground('#f3f3f3');
      headerRange.setHorizontalAlignment('center');
    }
    
    // ログを追加
    const now = new Date();
    const logRow = [
      Utilities.formatDate(now, 'Asia/Tokyo', 'yyyy-MM-dd HH:mm:ss'),
      startDate,
      endDate,
      recordCount,
      elapsedTime.toFixed(2),
      status === 'success' ? '成功' : 'エラー',
      errorMessage || ''
    ];
    
    logSheet.appendRow(logRow);
    
    console.log('実行ログを記録しました');
    
  } catch (error) {
    console.error('ログ記録エラー:', error.message);
    // ログ記録のエラーは処理を止めない
  }
}

/**
 * テスト実行: デフォルト（本日から3日分）
 */
function testUpdateShippingDataDefault() {
  console.log('=== デフォルト実行テスト ===');
  console.log('本日から3日分のデータを取得します');
  console.log('');
  
  updateShippingData();
}

/**
 * テスト実行: 期間指定
 */
function testUpdateShippingDataCustom() {
  console.log('=== 期間指定実行テスト ===');
  console.log('');
  
  updateShippingData('2025-10-03', '2025-10-05');
}