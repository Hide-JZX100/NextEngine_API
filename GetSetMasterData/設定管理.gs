/**
 * =============================================================================
 * 設定管理
 * =============================================================================
 * スクリプトプロパティから各種設定値を取得する関数群
 * 
 * 【管理する設定項目】
 * - SPREADSHEET_ID: 出力先スプレッドシートID
 * - SHEET_NAME: 出力先シート名
 * - LOG_LEVEL: ログ出力レベル(1:全件, 2:先頭3行, 3:なし)
 * - CLIENT_ID, CLIENT_SECRET, REDIRECT_URI: 認証情報(認証ライブラリで使用)
 * =============================================================================
 */

/**
 * スプレッドシート設定を取得
 * 
 * @return {Object} スプレッドシート設定 {spreadsheetId, sheetName}
 * @throws {Error} 設定が未設定の場合
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
 * @return {number} ログレベル(1:全件, 2:先頭3行のみ, 3:なし)
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
 * 現在の設定を全て表示(デバッグ用)
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