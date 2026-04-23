// ============================================================
// 06_ログと通知.gs
// ネクストエンジン Amazon受注データ取得 - ログ記録・エラー通知
// ============================================================

/**
 * ログ出力先シートタブ名を取得する
 * 
 * 【処理フロー】
 * 1. スクリプトプロパティ LOG_SHEET_NAME を取得
 * 2. 未設定の場合はデフォルト値 'LOG' を返す
 * 
 * @returns {string} ログシートのタブ名
 */
function getLogSheetName() {
  const props = PropertiesService.getScriptProperties();
  return props.getProperty('LOG_SHEET_NAME') || 'LOG';
}

/**
 * LOGシートが存在しない場合に新規作成し、ヘッダー行を書き込む
 * 
 * 【処理フロー】
 * 1. OUTPUT_SPREADSHEET_ID のスプレッドシートを開く
 * 2. getLogSheetName() でシート名を取得
 * 3. 同名のシートが既に存在する場合は何もせず終了
 * 4. 存在しない場合はシートを新規作成
 * 5. ヘッダー行を1行目に書き込み、スタイルを適用
 */
function initLogSheet() {
  const spreadsheetId = getOutputSpreadsheetId();
  if (!spreadsheetId) {
    console.warn('スプレッドシートIDが設定されていないため、LOGシートの初期化をスキップします。');
    return;
  }
  
  const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
  const sheetName = getLogSheetName();
  
  let sheet = spreadsheet.getSheetByName(sheetName);
  if (sheet) {
    // 既に存在する場合は何もしない（冪等性）
    return;
  }
  
  // シートの新規作成
  sheet = spreadsheet.insertSheet(sheetName);
  
  // ヘッダー行の定義と書き込み
  const headers = ['実行日時', '対象日付', '実行関数名', '取得件数', '実行結果', 'エラー内容'];
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setValues([headers]);
  
  // スタイルの適用（オレンジ背景・白文字・太字・中央揃え）
  headerRange.setBackground('#e67e22');
  headerRange.setFontColor('#ffffff');
  headerRange.setFontWeight('bold');
  headerRange.setHorizontalAlignment('center');
}

/**
 * LOGシートに実行結果の1行を追記する
 * 
 * 【処理フロー】
 * 1. OUTPUT_SPREADSHEET_ID のスプレッドシートを開く
 * 2. getLogSheetName() でログシートを取得
 * 3. LOGシートが存在しない場合は initLogSheet() を呼んで初期化する
 * 4. 現在日時を yyyy-MM-dd HH:mm:ss 形式で取得
 * 5. 引数の内容を1行の配列に組み立てる
 * 6. シートの最終行の次の行に appendRow() で追記する
 * ※ ログ記録の失敗で本処理が止まらないよう try-catch で囲む
 * 
 * @param {Object} params - ログパラメータ
 * @param {string} params.targetDate   - 取得対象日付（"YYYY-MM-DD"）
 * @param {string} params.funcName     - 実行関数名（"dailyRun" など）
 * @param {number} params.count        - 取得件数
 * @param {string} params.status       - 実行結果（"成功" / "0件" / "エラー"）
 * @param {string} [params.errorMsg]   - エラー内容（省略可。正常時は空文字）
 */
function writeLog(params) {
  try {
    const spreadsheetId = getOutputSpreadsheetId();
    if (!spreadsheetId) {
      console.warn('スプレッドシートIDが設定されていないため、ログ書き込みをスキップします。');
      return;
    }
    
    const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    const sheetName = getLogSheetName();
    let sheet = spreadsheet.getSheetByName(sheetName);
    
    if (!sheet) {
      initLogSheet();
      sheet = spreadsheet.getSheetByName(sheetName);
    }
    
    // 現在日時をフォーマット
    const now = new Date();
    const yyyy = now.getFullYear();
    const MM = ('0' + (now.getMonth() + 1)).slice(-2);
    const dd = ('0' + now.getDate()).slice(-2);
    const HH = ('0' + now.getHours()).slice(-2);
    const mm = ('0' + now.getMinutes()).slice(-2);
    const ss = ('0' + now.getSeconds()).slice(-2);
    const datetimeStr = yyyy + '-' + MM + '-' + dd + ' ' + HH + ':' + mm + ':' + ss;
    
    // 追記データの組み立て
    const rowData = [
      datetimeStr,
      params.targetDate || '',
      params.funcName || '',
      params.count !== undefined ? params.count : '',
      params.status || '',
      params.errorMsg || ''
    ];
    
    // シートの最終行へ追記
    sheet.appendRow(rowData);
    
  } catch (error) {
    // ログ記録の失敗はコンソールに出力するのみで、例外はスローしない
    console.error('ログの書き込みに失敗しました: ' + error.message);
  }
}

/**
 * 【Phase 5-1 テスト用】動作確認関数
 * LOGシートの初期化と、成功・0件・エラーのダミーログ書き込みを確認する
 */
function testPhase5_1() {
  console.log('=== Phase 5-1 テスト開始 ===');
  
  try {
    console.log('1. LOGシートの初期化を実行します');
    initLogSheet();
    
    console.log('2. 成功パターンのログを書き込みます');
    writeLog({ targetDate: '2026-04-22', funcName: 'dailyRun', count: 42, status: '成功', errorMsg: '' });
    
    console.log('3. 0件パターンのログを書き込みます');
    writeLog({ targetDate: '2026-04-21', funcName: 'dailyRun', count: 0,  status: '0件', errorMsg: '' });
    
    console.log('4. エラーパターンのログを書き込みます');
    writeLog({ targetDate: '2026-04-20', funcName: 'dailyRun', count: 0,  status: 'エラー', errorMsg: 'APIリクエストエラー: {"result":"error","message":"token error"}' });
    
    console.log('=== テスト完了。スプレッドシートのLOGタブに3行のデータが追記されているか確認してください ===');
  } catch (error) {
    console.error('テスト中にエラーが発生しました: ' + error.message);
  }
}

/**
 * エラー発生時に通知先メールアドレスへエラー通知メールを送信する
 * 
 * 【処理フロー】
 * 1. スクリプトプロパティ NOTIFY_EMAIL から通知先メールアドレスを取得
 * 2. 未設定の場合はコンソールにエラーを出力してメール送信をスキップ
 * 3. 件名・本文を組み立てて MailApp.sendEmail() で送信
 * 
 * @param {Object} params - 通知パラメータ
 * @param {string} params.targetDate - 取得対象日付
 * @param {string} params.funcName   - エラーが発生した関数名
 * @param {string} params.errorMsg   - エラー内容
 */
function sendErrorNotification(params) {
  try {
    const props = PropertiesService.getScriptProperties();
    const email = props.getProperty('NOTIFY_EMAIL');
    
    if (!email) {
      console.warn('通知先メールアドレス(NOTIFY_EMAIL)が未設定のため、エラー通知メールをスキップします。');
      return;
    }
    
    const subject = '[エラー] NE_(Amazon)受注情報取得失敗';
    
    const now = new Date();
    const yyyy = now.getFullYear();
    const MM = ('0' + (now.getMonth() + 1)).slice(-2);
    const dd = ('0' + now.getDate()).slice(-2);
    const HH = ('0' + now.getHours()).slice(-2);
    const mm = ('0' + now.getMinutes()).slice(-2);
    const ss = ('0' + now.getSeconds()).slice(-2);
    const datetimeStr = yyyy + '-' + MM + '-' + dd + ' ' + HH + ':' + mm + ':' + ss;
    
    const body = `NE(Amazon)受注データの自動取得中にエラーが発生しました。

■ 対象日付: ${params.targetDate || '不明'}
■ 実行関数: ${params.funcName || '不明'}
■ 発生時刻: ${datetimeStr}
■ エラー内容:
${params.errorMsg || '不明なエラー'}

【対処方法】
1. GASエディタのログを確認してください
2. トークンエラーの場合は testGenerateAuthUrl() を実行して再認証してください
3. ネクストエンジン側の一時的なエラーの場合は manualRun('YYYY-MM-DD') で手動再実行してください

このメールは自動送信されています。`;

    MailApp.sendEmail({
      to: email,
      subject: subject,
      body: body
    });
    
  } catch (error) {
    // メール送信エラーで本処理を止めない
    console.error('エラー通知メールの送信に失敗しました: ' + error.message);
  }
}
