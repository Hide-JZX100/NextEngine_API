/**
 * @fileoverview 設定管理モジュール
 * 
 * このスクリプトは、Google Apps Script の「スクリプトプロパティ」に保存された
 * 設定情報を安全に取得・管理するためのユーティリティ関数を提供します。
 * 
 * 管理する設定項目:
 * - SPREADSHEET_ID: データの読み書き対象となるスプレッドシートのID
 * - SHEET_NAME: データの読み書き対象となるシート名
 * - LOG_LEVEL: 実行時のログ出力レベル（1:詳細, 2:概要のみ, 3:出力なし）
 * - CLIENT_ID, CLIENT_SECRET, REDIRECT_URI: Next Engine API 連携用の認証情報
 * - ACCESS_TOKEN, REFRESH_TOKEN: APIアクセス用トークン（動的に更新されます）
 */

/**
 * スプレッドシート設定を取得
 * 
 * スクリプトプロパティからスプレッドシートIDとシート名を取得します。
 * 設定が欠落している場合は、処理を続行できないため例外をスローします。
 * 
 * @return {{spreadsheetId: string, sheetName: string}} スプレッドシート設定オブジェクト
 * @throws {Error} SPREADSHEET_ID または SHEET_NAME がスクリプトプロパティに設定されていない場合
 */
function getSpreadsheetConfig() {
  const props = PropertiesService.getScriptProperties();
  const spreadsheetId = props.getProperty('SPREADSHEET_ID');
  const sheetName = props.getProperty('SHEET_NAME');
  
  if (!spreadsheetId) {
    throw new Error('SPREADSHEET_IDが設定されていません。スクリプトプロパティを確認してください。');
  }
  
  if (!sheetName) {
    throw new Error('SHEET_NAMEが設定されていません。スクリプトプロパティを確認してください。');
  }
  
  return {
    spreadsheetId: spreadsheetId,
    sheetName: sheetName
  };
}

/**
 * ログレベルを取得
 * 
 * スクリプトの動作ログの出力密度を制御するためのレベルを取得します。
 * 1: 全件出力 (デバッグ用)
 * 2: 先頭3行のみ (デフォルト / 正常確認用)
 * 3: 出力なし (パフォーマンス優先)
 * 
 * @return {number} ログレベル (1|2|3)
 */
function getLogLevel() {
  const props = PropertiesService.getScriptProperties();
  const logLevel = props.getProperty('LOG_LEVEL');
  
  // 未設定の場合はデフォルト値2(先頭3行のみ)
  if (!logLevel) {
    console.log('LOG_LEVELが未設定です。デフォルト値2(先頭3行のみ)を使用します。');
    return 2;
  }
  
  const level = parseInt(logLevel);
  
  // 1-3の範囲チェック
  if (level < 1 || level > 3) {
    console.log('LOG_LEVELが不正な値です。デフォルト値2(先頭3行のみ)を使用します。');
    return 2;
  }
  
  return level;
}

/**
 * 現在の設定情報をログに出力（デバッグ用）
 * 
 * スクリプトプロパティに格納されている現在の設定値を一覧表示します。
 * 認証情報（Secret等）はセキュリティのため、値そのものではなく「設定済み」か否かのみを表示します。
 * 開発環境のセットアップ確認に使用してください。
 */
function showAllConfig() {
  console.log('=== 現在の設定 ===');
  
  try {
    const spreadsheetConfig = getSpreadsheetConfig();
    console.log('SPREADSHEET_ID:', spreadsheetConfig.spreadsheetId);
    console.log('SHEET_NAME:', spreadsheetConfig.sheetName);
  } catch (error) {
    console.error('スプレッドシート設定エラー:', error.message);
  }
  
  const logLevel = getLogLevel();
  const logLevelText = ['', '全件ログ出力', '先頭3行のみ', 'ログ出力なし'];
  console.log('LOG_LEVEL:', logLevel, '(' + logLevelText[logLevel] + ')');
  
  console.log('');
  console.log('=== 認証情報 ===');
  const props = PropertiesService.getScriptProperties();
  console.log('CLIENT_ID:', props.getProperty('CLIENT_ID') ? '設定済み' : '未設定');
  console.log('CLIENT_SECRET:', props.getProperty('CLIENT_SECRET') ? '設定済み' : '未設定');
  console.log('REDIRECT_URI:', props.getProperty('REDIRECT_URI') || '未設定');
  console.log('ACCESS_TOKEN:', props.getProperty('ACCESS_TOKEN') ? '設定済み' : '未設定');
  console.log('REFRESH_TOKEN:', props.getProperty('REFRESH_TOKEN') ? '設定済み' : '未設定');
}

/**
 * スプレッドシートプロパティの設定テスト
 * 実際にスプレッドシートとシートが存在するか確認
 * 
 * 設定されたIDとシート名を使って、実際にGoogle Sheets API経由でアクセスできるかを検証します。
 * 権限エラーやシート名の間違いを早期に発見するために使用します。
 * @throws {Error} スプレッドシートが開けない、またはシートが見つからない場合
 */
function testSpreadsheetConfig() {
  console.log('=== スプレッドシート設定テスト ===');
  
  try {
    const config = getSpreadsheetConfig();
    console.log('設定取得成功');
    console.log('SPREADSHEET_ID:', config.spreadsheetId);
    console.log('SHEET_NAME:', config.sheetName);
    
    // スプレッドシートを開く
    const spreadsheet = SpreadsheetApp.openById(config.spreadsheetId);
    console.log('✅ スプレッドシートを開けました');
    console.log('スプレッドシート名:', spreadsheet.getName());
    
    // シートを取得
    const sheet = spreadsheet.getSheetByName(config.sheetName);
    if (!sheet) {
      throw new Error(`シート "${config.sheetName}" が見つかりません。`);
    }
    
    console.log('✅ シートを取得できました');
    console.log('シート名:', sheet.getName());
    console.log('最終行:', sheet.getLastRow());
    console.log('最終列:', sheet.getLastColumn());
    
    console.log('');
    console.log('✅ スプレッドシート設定テスト成功!');
    
  } catch (error) {
    console.error('❌ スプレッドシート設定テストエラー:', error.message);
    console.error('');
    console.error('【確認事項】');
    console.error('1. SPREADSHEET_IDが正しいか');
    console.error('2. スプレッドシートへのアクセス権限があるか');
    console.error('3. SHEET_NAMEが正確に一致しているか(大文字小文字、スペースに注意)');
    throw error;
  }
}