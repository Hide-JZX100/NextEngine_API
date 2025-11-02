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
主な処理関数
=============================================================================
updateShippingData(startDate, endDate)
💡 メイン処理を統括する関数
このスクリプトのメイン関数です。出荷予定データの更新処理全体を管理します。
機能:
初期設定: 処理開始時刻を記録し、startDateとendDateが未指定の場合は、本日を含む未来3日分の期間をデフォルトで設定します。
データ取得（リトライ付き）: 後述のfetchAllShippingDataWithRetry関数を呼び出し、指定期間の出荷予定データをネクストエンジンAPIから取得します。
データ変換: 後述のconvertAllDataToSheetRows関数を呼び出し、取得したAPIデータをスプレッドシートに書き込むための2次元配列形式に変換します。
スプレッドシート書き込み: 後述のwriteDataToSheet関数を呼び出し、変換後のデータをスプレッドシートに書き込みます。
処理完了・ログ記録: 処理時間、取得件数などを計算し、後述のlogExecution関数で実行ログを記録します。
エラーハンドリング: try...catchブロックで全体を囲み、処理中にエラーが発生した場合はそのステップと内容をコンソールに表示し、logExecutionでエラーログとして記録した後、エラーを再スローします（トリガー実行時の通知に役立ちます）。

fetchAllShippingDataWithRetry(startDate, endDate, maxRetries = 3)
💡 API接続エラーに対応するリトライ処理
ネクストエンジンAPIからデータを取得する処理（未定義のfetchAllShippingDataを呼び出し）にリトライ機能を追加します。
機能:
最大3回（maxRetries）までAPI接続を試行します。
接続が失敗した場合、試行回数に応じて待機時間（2秒、4秒、6秒...）を増やしてからリトライします（指数バックオフのような仕組み）。
規定回数リトライしても成功しない場合は、エラーをスローしてメイン処理にエラーを伝えます。

convertAllDataToSheetRows(apiData)
💡 一括書き込みのためのデータ形式変換
APIから取得したデータを、スプレッドシートに一括で書き込むための2次元配列（行と列の形式）に変換します。
機能:
取得したAPIデータ（オブジェクトの配列）を1件ずつ処理し、convertApiDataToSheetRow関数（未定義）を呼び出して1行分の配列に変換します。
変換後のデータが14列あるか検証し、列数に異常があればエラーをスローします。これにより、データ構造の変更などに早期に気づけます。
変換の進捗（50件ごと）をコンソールに表示します。
エラーハンドリング: 個別のデータ変換でエラーが発生した場合、そのデータの内容をコンソールに表示し、エラーをスローします。

writeDataToSheet(convertedData)
💡 空白行問題を解消したスプレッドシートへの書き込み
変換後のデータをスプレッドシートに書き込みます。特に、既存データを物理的に削除することで、GASでよくある空白行が残る問題（幽霊行）を解決しています。
機能:
スクリプトプロパティからスプレッドシートIDとシート名を取得します。
シートが存在しない場合は、後述のcreateSheetWithHeaders関数でヘッダー付きで作成します。
後述のensureHeaderExists関数でヘッダー行の存在を確認します。
2行目から最終行までの既存データ行を物理的にすべて削除します（これが「空白行問題解決済み」のポイント）。
新しいデータを2行目以降に一括で書き込みます。
=============================================================================
サポート・ユーティリティ関数
=============================================================================
createSheetWithHeaders(spreadsheet, sheetName)
💡 ヘッダー付きシート作成
指定されたスプレッドシート内に、定義された14列のヘッダー付きシートを新規作成します。ヘッダー行には太字や背景色などのスタイルも設定します。

ensureHeaderExists(sheet)
💡 ヘッダー行の存在確認と復元
シートの1行目にヘッダーが書き込まれているかを確認し、ヘッダーが存在しない場合は自動で作成します。

logExecution(startDate, endDate, recordCount, elapsedTime, status, errorMessage)
💡 実行履歴の記録
処理の実行結果（成功・失敗、処理時間、取得件数など）を専用の実行ログシートに追記します。
機能:
スクリプトプロパティからログシート名を取得し、シートがなければ作成します。
実行日時、期間、件数、時間、ステータス、エラー内容をログ行として記録します。
エラーログの記録失敗がメイン処理を妨げないよう、この関数自体にもtry...catchを設けています。

recordExecutionTimestamp()
💡 実行完了日時を指定されたシートに記録する関数
シート名はスクリプトプロパティに保存するので、任意のシート名を設定してください。
また、実行完了日時はそのシートのA1セルに記録するようにしていますので、
A1セルには他の情報を入力しないようにしてください。

=============================================================================
テスト関数
=============================================================================
testUpdateShippingDataDefault(): 引数なしでupdateShippingData()を実行し、デフォルトの期間（本日含む未来3日分）でテストします。

testUpdateShippingDataCustom(): 期間を指定してupdateShippingData('2025-10-03', '2025-10-05')を実行し、期間指定での動作をテストします。

testErrorHandling(): 遠い未来の日付を指定するなどして、データ取得が0件の場合や、エラー処理が正しく動作するかをテストします。

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
    // ★★★ 修正ポイント: トリガー実行時も正しく動作するように改善 ★★★
    if (!startDate || typeof startDate !== 'string') {
      const today = new Date();
      // タイムゾーンを考慮して日付文字列を生成
      const year = today.getFullYear();
      const month = String(today.getMonth() + 1).padStart(2, '0');
      const day = String(today.getDate()).padStart(2, '0');
      startDate = `${year}-${month}-${day}`;
    }
    
    if (!endDate || typeof endDate !== 'string') {
      const today = new Date();
      const threeDaysLater = new Date(today.getTime() + (2 * 24 * 60 * 60 * 1000)); // 本日+2日
      // タイムゾーンを考慮して日付文字列を生成
      const year = threeDaysLater.getFullYear();
      const month = String(threeDaysLater.getMonth() + 1).padStart(2, '0');
      const day = String(threeDaysLater.getDate()).padStart(2, '0');
      endDate = `${year}-${month}-${day}`;
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
      
      // データ0件でも正常終了として扱う（ログは記録しない）
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
    
    const writeStartTime = new Date(); // ←追加
    writeDataToSheet(convertedData);
    const writeEndTime = new Date(); // ←追加
    const writeElapsedTime = (writeEndTime - writeStartTime) / 1000; // ←追加
    
    console.log(`✓ 書き込み完了: ${writeElapsedTime.toFixed(2)}秒`); // ←修正
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
    
    // 実行完了日時を記録
    recordExecutionTimestamp();

    // ========================================
    // 6. 正常終了（ログは記録しない）
    // ========================================
    // エラー時のみログを記録するため、成功時のログ記録は行わない
    
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
    
    // エラーログを記録（エラー時のみ）
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
        const waitSeconds = Math.pow(2, attempt - 1); // 1秒、2秒、4秒...と待機時間を指数関数的に増やすように修正
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
 * スプレッドシート書き込み（改善版：空白行問題解決 + 固定行対応）
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
    
    // 既存データ行を削除または内容クリア
    // ★★★ 改善ポイント: 固定行対応 ★★★
    const deleteStartTime = new Date(); // ←追加
    const lastRow = sheet.getLastRow();
    console.log(`現在の行数: ${lastRow}行`); // ←追加

    if (lastRow > 1) {
      const rowsToDelete = lastRow - 1; // ヘッダーを除く行数
      console.log(`削除対象: ${rowsToDelete}行`); // ←追加
      console.log(`既存データ行を処理: ${rowsToDelete}行`);
      
      if (convertedData.length === 0) {
        // 新しいデータが0件の場合は、内容だけクリア（行は残す）
        console.log('データ0件のため、内容のみクリアします');
        sheet.getRange(2, 1, rowsToDelete, sheet.getLastColumn()).clearContent();
      } else if (rowsToDelete > 1) {
        // 新しいデータがある場合は、最後の1行を残して物理削除
        console.log(`既存データ行を削除: ${rowsToDelete - 1}行（1行は残します）`);
        sheet.deleteRows(2, rowsToDelete - 1); // 2行目から (最終行-1) まで削除
        
        // 残った1行（2行目）の内容をクリア
        sheet.getRange(2, 1, 1, sheet.getLastColumn()).clearContent();
      } else {
        // 既存データが1行のみの場合は、内容だけクリア
        console.log('既存データ1行のため、内容のみクリアします');
        sheet.getRange(2, 1, 1, sheet.getLastColumn()).clearContent();
      }
    }
    
    const deleteEndTime = new Date(); // ←追加
    const deleteElapsedTime = (deleteEndTime - deleteStartTime) / 1000; // ←追加
    console.log(`削除処理完了: ${deleteElapsedTime.toFixed(2)}秒`); // ←追加

    // データを書き込み
    const insertStartTime = new Date(); // ←追加
    if (convertedData.length > 0) {
      console.log(`データを書き込み中: ${convertedData.length}件`);
      const range = sheet.getRange(2, 1, convertedData.length, convertedData[0].length);
      range.setValues(convertedData);

      const insertEndTime = new Date(); // ←追加
      const insertElapsedTime = (insertEndTime - insertStartTime) / 1000; // ←追加
      console.log(`データ書き込み完了: ${insertElapsedTime.toFixed(2)}秒`); // ←追加

    } else {
      console.log('書き込むデータがありません（データ行は空になります）');
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

function recordExecutionTimestamp() {
  try {
    const properties = PropertiesService.getScriptProperties();
    const spreadsheetId = properties.getProperty('SPREADSHEET_ID');
    const sheetName = properties.getProperty('OPERATION_LOG_SHEET_NAME');

    if (!spreadsheetId || !sheetName) {
      throw new Error('スクリプトプロパティ SPREADSHEET_ID または OPERATION_LOG_SHEET_NAME が設定されていません。');
    }

    const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    const sheet = spreadsheet.getSheetByName(sheetName);

    if (!sheet) {
      console.error(`シート "${sheetName}" が見つかりません。日時の記録をスキップします。`);
      return;
    }
    
    sheet.getRange(1, 1).setValue(
      Utilities.formatDate(new Date(), 'JST', 'MM月dd日HH時mm分ss秒')
    );
    console.log(`実行日時をシート "${sheetName}" のA1セルに記録しました。`);

  } catch (error) {
    console.error('実行日時の記録中にエラーが発生しました:', error.message);
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