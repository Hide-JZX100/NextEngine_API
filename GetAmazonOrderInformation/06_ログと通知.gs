/**
 * @file 06_ログと通知.gs
 * @description ネクストエンジン 受注データ取得ログ記録・エラー通知モジュール
 *
 * 【概要】
 * 本プロジェクトの実行結果をスプレッドシートへ記録し、異常終了時には
 * 管理者へメールでアラートを通知する機能を管理します。
 * 
 * 【主な機能】
 * - writeLog: 実行結果（成功・件数・エラー内容等）をログシートに一行追記。
 * - initLogSheet: ログ専用シートの自動判定・初期作成（ヘッダー付与）。
 * - sendErrorNotification: 実行エラー時の詳細（日時・関数・内容・対処法）をメール送信。
 *
 * 【運用上の注意】
 * - ログシート名はスクリプトプロパティ「LOG_SHEET_NAME」でカスタマイズ可能です（デフォルト: 'LOG'）。
 * - エラー通知を有効にするには、スクリプトプロパティ「NOTIFY_EMAIL」に受信アドレスの設定が必要です。
 * - ログ記録自体の失敗によってメイン処理が停止しないよう、内部で例外を捕捉(try-catch)しています。
 *
 * 【依存関係】
 * - 01_設定ファイル.gs (getOutputSpreadsheetId)
 * - Google Apps Script (SpreadsheetApp, MailApp, PropertiesService)
 */

/**
 * ログ出力先のシート名を取得するヘルパー関数です。
 *
 * 【詳細仕様】
 * - スクリプトプロパティ `LOG_SHEET_NAME` を参照します。
 * - プロパティが未設定（null または空文字）の場合は、システム標準のデフォルト値 'LOG' を採用します。
 * - ログ記録および初期化処理の両方で使用され、一貫したシート名へのアクセスを保証します。
 *
 * @returns {string} ログシートのタブ名
 */
function getLogSheetName() {
  const props = PropertiesService.getScriptProperties();
  return props.getProperty('LOG_SHEET_NAME') || 'LOG';
}

/**
 * ログ記録用のシートが存在しない場合に、新規作成と初期設定（ヘッダー付与）を行います。
 *
 * 【詳細仕様】
 * 1. 存在確認: `getLogSheetName()` で取得した名前のシートが既に存在するか確認します。
 *    存在する場合は何もせず終了します（冪等性の確保）。
 * 2. シート作成: シートが存在しない場合、スプレッドシートの末尾に新しいシートを挿入します。
 * 3. ヘッダー生成: 1行目に「実行日時」「対象日付」「実行関数名」「取得件数」「実行結果」「エラー内容」を書き込みます。
 * 4. スタイル適用: 視認性を高めるため、ヘッダー行に以下の書式を設定します。
 *    - 背景色: オレンジ (#e67e22)
 *    - 文字色: 白 (#ffffff)
 *    - フォント: 太字
 *    - 配置: 中央揃え
 *
 * 【使用タイミング】
 * - `writeLog()` の実行時にシートが見つからない場合、自動的に呼び出されます。
 * - 初回セットアップ時に手動で実行することも可能です。
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
 * 実行結果やエラー情報をログシートの最終行に追記します。
 *
 * 【詳細仕様】
 * 1. シート自動生成: 対象のログシートが存在しない場合、`initLogSheet()` を内部で呼び出して自動的に作成します。
 * 2. タイムスタンプ: 実行時のサーバー時刻を `yyyy-MM-dd HH:mm:ss` 形式で生成し、記録のキーとします。
 * 3. データの安全性: 引数 `params` 内の各値が未定義であっても、空文字にフォールバックして書き込みエラーを防ぎます。
 * 4. 耐障害性: ログ記録自体が失敗（シートの権限不足など）しても、`try-catch` によりメインの受注取得処理が
 *    中断されないよう設計されています。エラーはコンソールログに出力されます。
 *
 * 【使用タイミング】
 * - `dailyRun` や `manualRun` の処理終了時、または例外発生時に、実行ステータスを永続化するために使用します。
 *
 * @param {Object} params - ログに記録するパラメータオブジェクト
 * @param {string} params.targetDate - 処理対象とした出荷確定日（"YYYY-MM-DD"）
 * @param {string} params.funcName - 実行されたメイン関数名
 * @param {number} params.count - APIから取得した受注データの総件数
 * @param {string} params.status - 実行ステータス（例: "成功", "0件", "エラー"）
 * @param {string} [params.errorMsg] - エラー発生時の詳細メッセージ（正常終了時は省略可）
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
 * 【開発・検証用】Phase 5-1：ログ記録機能の単体テスト
 *
 * ログシートの初期化から、正常系・異常系それぞれのログ書き込みが正しく動作するかを検証します。
 *
 * 【検証項目】
 * - ログシートが存在しない場合に自動生成されるか。
 * - ヘッダーに適切なスタイルが適用されているか。
 * - `appendRow` によってデータが重複せず正しく追記されるか。
 */
function testPhase5_1() {
  console.log('=== Phase 5-1 テスト開始 ===');
  
  try {
    console.log('1. LOGシートの初期化を実行します');
    initLogSheet();
    
    console.log('2. 0件パターンのログを書き込みます');
    writeLog({ targetDate: '2026-04-21', funcName: 'dailyRun', count: 0,  status: '0件', errorMsg: '' });
    
    console.log('3. エラーパターンのログを書き込みます');
    writeLog({ targetDate: '2026-04-20', funcName: 'dailyRun', count: 0,  status: 'エラー', errorMsg: 'APIリクエストエラー: {"result":"error","message":"token error"}' });
    
    console.log('=== テスト完了。スプレッドシートのLOGタブに2行のデータが追記されているか確認してください ===');
  } catch (error) {
    console.error('テスト中にエラーが発生しました: ' + error.message);
  }
}

/**
 * システム実行中に発生した致命的なエラーを、管理者にメールで通知します。
 *
 * 【詳細仕様】
 * 1. 通知先制御: スクリプトプロパティ `NOTIFY_EMAIL` に登録されたアドレス宛に送信します。
 *    アドレスが未設定の場合は、警告をコンソールに出力し送信をスキップします。
 * 2. 文面生成: エラーが発生した日付、関数、時刻、および具体的なエラー内容を含む詳細な本文を生成します。
 * 3. リカバリガイド: 本文には、管理者向けに考えられる原因（トークン失効、一時的エラー等）と
 *    それに対応するリカバリ手順（再認証、手動再実行の案内）を付加し、迅速な復旧を支援します。
 * 4. 実行分離: 通知処理自体の失敗が、呼び出し元の例外処理を妨げないよう内部でエラーをトラップしています。
 *
 * 【使用タイミング】
 * - `dailyRun` や `manualRun` 内の `catch` ブロックで、即時の人間による介入が必要な場合に呼び出されます。
 *
 * @param {Object} params - 通知内容を構成するオブジェクト
 * @param {string} params.targetDate - エラーが発生した処理対象日
 * @param {string} params.funcName - エラーが発生した関数の名称
 * @param {string} params.errorMsg - 発生した例外の詳細メッセージ
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
