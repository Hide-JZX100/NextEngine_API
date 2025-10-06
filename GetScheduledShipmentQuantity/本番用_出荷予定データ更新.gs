/**
=============================================================================
本番用: ネクストエンジン出荷予定データ更新（エラーハンドリング強化版）
=============================================================================

* 【目的】
* ネクストエンジンから出荷予定データを取得し、スプレッドシートを更新する
* 
* 【機能】
* - 指定期間の出荷予定データを取得
* - スプレッドシートに一括書き込み（空白行問題解決済み）
* - 実行ログの記録
* - エラーハンドリング強化
* - API接続リトライ処理
* 
* 【使用方法】
* updateShippingData() - デフォルト（本日含む未来3日分）
* updateShippingData('2025-10-03', '2025-10-05') - 期間指定
* 
* 【トリガー設定】
* トリガー設定スクリプト.gs と連携
* スクリプトプロパティ: TRIGGER_FUNCTION_NAME = updateShippingData
* 
=============================================================================
*/

/**
 * 本番用: 出荷予定データを更新（エラーハンドリング強化版）
 * 
 * @param {string} startDate - 開始日（YYYY-MM-DD形式）省略時は本日
 * @param {string} endDate - 終了日（YYYY-MM-DD形式）省略時は本日+2日
 */
function updateShippingData(startDate, endDate) {
  const startTime = new Date();
  let currentStep = '初期化';
  
  try {
    // ========================================
    // 1. 初期設定
    // ========================================
    console.log('');
    console.log('========================================');
    console.log('  出荷予定データ更新開始');
    console.log('========================================');
    console.log('');
    
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
    // 2. データ取得（リトライ処理付き）
    // ========================================
    currentStep = 'データ取得';
    console.log('--- ステップ1: データ取得 ---');
    console.log('ネクストエンジンAPIからデータを取得中...');
    
    const apiData = fetchAllShippingDataWithRetry(startDate, endDate);
    
    if (!apiData || apiData.length === 0) {
      console.log('⚠ 取得データが0件でした');
      console.log('期間内に出荷予定データが存在しない可能性があります');
      console.log('');
      
      // データ0件でも正常終了として扱う
      logExecution(startDate, endDate, 0, 0, 'success', '取得データ0件');
      return {
        success: true,
        startDate: startDate,
        endDate: endDate,
        recordCount: 0,
        message: '取得データ0件'
      };
    }
    
    console.log(`✓ 取得完了: ${apiData.length}件`);
    console.log('');
    
    // ========================================
    // 3. データ変換
    // ========================================
    currentStep = 'データ変換';
    console.log('--- ステップ2: データ変換 ---');
    console.log('スプレッドシート形式に変換中...');
    
    const convertedData = convertAllDataToSheetRows(apiData);
    
    console.log(`✓ 変換完了: ${convertedData.length}件`);
    console.log('');
    
    // データ検証: 変換前後で件数が一致するか
    if (convertedData.length !== apiData.length) {
      throw new Error(`データ変換エラー: 変換前${apiData.length}件、変換後${convertedData.length}件で不一致`);
    }
    
    // ========================================
    // 4. スプレッドシート書き込み
    // ========================================
    currentStep = 'スプレッドシート書き込み';
    console.log('--- ステップ3: スプレッドシート書き込み ---');
    
    writeDataToSheet(convertedData);
    
    console.log('✓ 書き込み完了');
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
    // エラー処理
    const endTime = new Date();
    const elapsedTime = (endTime - startTime) / 1000;
    
    console.error('');
    console.error('========================================');
    console.error('  エラーが発生しました');
    console.error('========================================');
    console.error(`エラー発生ステップ: ${currentStep}`);
    console.error(`エラー内容: ${error.message}`);
    console.error(`エラー詳細:`, error);
    console.error(`発生時刻: ${Utilities.formatDate(endTime, 'Asia/Tokyo', 'yyyy-MM-dd HH:mm:ss')}`);
    console.error('');
    
    // エラーログを記録
    const errorMessage = `[${currentStep}] ${error.message}`;
    logExecution(startDate || 'N/A', endDate || 'N/A', 0, elapsedTime, 'error', errorMessage);
    
    // エラーを再スロー（トリガー実行時のエラー通知のため）
    throw new Error(`出荷予定データ更新エラー: ${errorMessage}`);
  }
}

/**
 * リトライ処理付きデータ取得
 * 
 * @param {string} startDate - 開始日
 * @param {string} endDate - 終了日
 * @param {number} maxRetries - 最大リトライ回数（デフォルト: 3）
 * @return {Array} 取得データ
 */
function fetchAllShippingDataWithRetry(startDate, endDate, maxRetries = 3) {
  let lastError = null;
  
  for (let attempt = 1; attempt <= maxRetries; attempt++) {
    try {
      console.log(`API接続試行: ${attempt}回目`);
      
      const data = fetchAllShippingData(startDate, endDate);
      
      // 成功したらデータを返す
      return data;
      
    } catch (error) {
      lastError = error;
      console.error(`API接続エラー（試行${attempt}回目）: ${error.message}`);
      
      // 最後の試行でなければ、少し待ってからリトライ
      if (attempt < maxRetries) {
        const waitSeconds = attempt * 2; // 2秒、4秒、6秒...と待機時間を増やす
        console.log(`${waitSeconds}秒後にリトライします...`);
        Utilities.sleep(waitSeconds * 1000);
      }
    }
  }
  
  // すべてのリトライが失敗した場合
  throw new Error(`API接続失敗（${maxRetries}回試行）: ${lastError.message}`);
}

/**
 * データ変換（エラーハンドリング付き）
 * 
 * @param {Array} apiData - APIから取得したデータ
 * @return {Array} 変換後のデータ（2次元配列）
 */
function convertAllDataToSheetRows(apiData) {
  const convertedData = [];
  
  for (let i = 0; i < apiData.length; i++) {
    try {
      const row = convertApiDataToSheetRow(apiData[i]);
      
      // データ検証: 14列あるか確認
      if (row.length !== 14) {
        throw new Error(`列数エラー: 期待値14列、実際${row.length}列`);
      }
      
      convertedData.push(row);
      
      // 進捗表示（50件ごと）
      if ((i + 1) % 50 === 0 || i === apiData.length - 1) {
        console.log(`  変換中: ${i + 1}/${apiData.length}件`);
      }
      
    } catch (error) {
      console.error(`データ変換エラー（${i + 1}件目）: ${error.message}`);
      console.error('エラーが発生したデータ:', JSON.stringify(apiData[i], null, 2));
      throw new Error(`データ変換エラー（${i + 1}件目）: ${error.message}`);
    }
  }
  
  return convertedData;
}

/**
 * スプレッドシート書き込み（改善版：空白行問題解決）
 * 
 * @param {Array} convertedData - 変換済みデータ（2次元配列）
 */
function writeDataToSheet(convertedData) {
  try {
    // スクリプトプロパティから設定を取得
    const properties = PropertiesService.getScriptProperties();
    const spreadsheetId = properties.getProperty('SPREADSHEET_ID');
    const sheetName = properties.getProperty('SHEET_NAME');
    
    if (!spreadsheetId || !sheetName) {
      throw new Error('スクリプトプロパティが設定されていません（SPREADSHEET_ID, SHEET_NAME）');
    }
    
    // スプレッドシートとシートを取得
    const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    let sheet = spreadsheet.getSheetByName(sheetName);
    
    // シートが存在しない場合は作成
    if (!sheet) {
      console.log(`シート「${sheetName}」が存在しないため、作成します`);
      sheet = createSheetWithHeaders(spreadsheet, sheetName);
    }
    
    // ヘッダー行の存在確認
    ensureHeaderExists(sheet);
    
    // 既存データ行を完全に削除（ヘッダー行は残す）
    // ★★★ 改善ポイント: 空白行が残らないように物理削除 ★★★
    const lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      console.log(`既存データ行を削除: ${lastRow - 1}行`);
      sheet.deleteRows(2, lastRow - 1); // 2行目から最終行まで物理的に削除
    }
    
    // データを書き込み
    if (convertedData.length > 0) {
      console.log(`データを書き込み中: ${convertedData.length}件`);
      const range = sheet.getRange(2, 1, convertedData.length, convertedData[0].length);
      range.setValues(convertedData);
    } else {
      console.log('書き込むデータがありません');
    }
    
  } catch (error) {
    throw new Error(`スプレッドシート書き込みエラー: ${error.message}`);
  }
}

/**
 * ヘッダー付きシートを作成
 * 
 * @param {Spreadsheet} spreadsheet - スプレッドシートオブジェクト
 * @param {string} sheetName - シート名
 * @return {Sheet} 作成したシート
 */
function createSheetWithHeaders(spreadsheet, sheetName) {
  const sheet = spreadsheet.insertSheet(sheetName);
  
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
  
  console.log('ヘッダー付きシートを作成しました');
  
  return sheet;
}

/**
 * ヘッダー行の存在を確認し、なければ作成
 * 
 * @param {Sheet} sheet - シートオブジェクト
 */
function ensureHeaderExists(sheet) {
  const headers = [
    '出荷予定日', '伝票番号', '商品コード', '商品名', '受注数', '引当数',
    '奥行き（cm）', '幅（cm）', '高さ（cm）', '発送方法コード', '発送方法',
    '重さ（g）', '受注状態区分', '送り先住所1'
  ];
  
  // 既存のヘッダー行を確認
  const existingHeaders = sheet.getRange(1, 1, 1, headers.length).getValues()[0];
  const hasHeaders = existingHeaders.some(cell => cell !== '');
  
  if (!hasHeaders) {
    console.log('ヘッダー行が存在しないため、作成します');
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    
    // ヘッダー行のスタイル設定
    const headerRange = sheet.getRange(1, 1, 1, headers.length);
    headerRange.setFontWeight('bold');
    headerRange.setBackground('#f3f3f3');
    headerRange.setHorizontalAlignment('center');
  }
}

/**
 * 実行ログを記録（エラーハンドリング強化版）
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
      typeof elapsedTime === 'number' ? elapsedTime.toFixed(2) : '0.00',
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

/**
 * テスト実行: エラーハンドリングのテスト
 * 存在しない日付を指定してエラー処理を確認
 */
function testErrorHandling() {
  console.log('=== エラーハンドリングテスト ===');
  console.log('存在しない未来の日付を指定してエラー処理を確認します');
  console.log('');
  
  try {
    // 遠い未来の日付（データが0件のはず）
    updateShippingData('2026-12-01', '2026-12-03');
  } catch (error) {
    console.log('エラーが正しくキャッチされました:', error.message);
  }
}