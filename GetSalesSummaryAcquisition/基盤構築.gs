/**
 * =============================================================================
 * ネクストエンジン受注明細取得スクリプト - Phase 1: 基盤構築
 * =============================================================================
 * 
 * 【概要】
 * ネクストエンジンAPIから受注明細を取得し、Googleスプレッドシートに書き込むための
 * 基盤となる機能を提供します。
 * 
 * 【必要なスクリプトプロパティ】
 * - ACCESS_TOKEN: ネクストエンジンAPIアクセストークン
 * - REFRESH_TOKEN: ネクストエンジンAPIリフレッシュトークン
 * - TARGET_SPREADSHEET_ID: 書き込み先スプレッドシートID
 * - TARGET_SHEET_NAME: 書き込み先シート名
 * - SHOP_MASTER_SPREADSHEET_ID: 店舗マスタスプレッドシートID
 * - SHOP_MASTER_SHEET_NAME: 店舗マスタシート名
 * - LOG_LEVEL: ログ出力レベル(1:全件, 2:先頭3行, 3:なし) ※省略時は3
 * - RETRY_COUNT: リトライ回数 ※省略時は3
 * - NOTIFICATION_EMAIL: 通知先メールアドレス ※省略時は実行者のメール
 * - NOTIFY_ON_SUCCESS: 成功時も通知するか(true/false) ※省略時はfalse
 * 
 * 【Phase 1 実装内容】
 * 1. スクリプトプロパティ取得処理
 * 2. 日付計算ロジック(先週日曜～本日)
 * 3. ログ出力制御機能
 * 4. 定数定義
 * 
 * @version 1.0
 * @date 2025-11-24
 */

// =============================================================================
// 定数定義
// =============================================================================

/**
 * ネクストエンジンAPIのベースURL
 */
const NE_API_BASE_URL = 'https://api.next-engine.org';

/**
 * 受注明細検索APIエンドポイント
 */
const NE_RECEIVEORDER_ROW_SEARCH_ENDPOINT = '/api_v1_receiveorder_row/search';

/**
 * API呼び出し時の待機フラグ(過負荷時も可能な限り実行)
 */
const NE_WAIT_FLAG = '1';

/**
 * API 1回あたりの最大取得件数
 */
const NE_API_LIMIT = 1000;

/**
 * ログレベル定義
 */
const LOG_LEVEL = {
  ALL: 1,      // 全件出力
  SAMPLE: 2,   // 先頭3行のみ出力
  NONE: 3      // 出力なし
};

// =============================================================================
// スクリプトプロパティ取得
// =============================================================================

/**
 * スクリプトプロパティから設定値を取得
 * 
 * 必須プロパティが未設定の場合はエラーをスローします。
 * オプションプロパティは未設定の場合デフォルト値を返します。
 * 
 * @return {Object} 設定オブジェクト
 * @throws {Error} 必須プロパティが未設定の場合
 */
function getScriptConfig() {
  const props = PropertiesService.getScriptProperties();
  
  // 必須プロパティ
  const requiredProps = {
    accessToken: props.getProperty('ACCESS_TOKEN'),
    refreshToken: props.getProperty('REFRESH_TOKEN'),
    targetSpreadsheetId: props.getProperty('TARGET_SPREADSHEET_ID'),
    targetSheetName: props.getProperty('TARGET_SHEET_NAME'),
    shopMasterSpreadsheetId: props.getProperty('SHOP_MASTER_SPREADSHEET_ID'),
    shopMasterSheetName: props.getProperty('SHOP_MASTER_SHEET_NAME')
  };
  
  // 必須プロパティの検証
  const missingProps = [];
  for (const [key, value] of Object.entries(requiredProps)) {
    if (!value) {
      missingProps.push(key);
    }
  }
  
  if (missingProps.length > 0) {
    throw new Error(
      `必須スクリプトプロパティが未設定です: ${missingProps.join(', ')}\n` +
      `スクリプトプロパティを設定してください。`
    );
  }
  
  // オプションプロパティ(デフォルト値あり)
  const logLevel = props.getProperty('LOG_LEVEL') || '3';
  const retryCount = props.getProperty('RETRY_COUNT') || '3';
  const notificationEmail = props.getProperty('NOTIFICATION_EMAIL') || Session.getActiveUser().getEmail();
  const notifyOnSuccess = props.getProperty('NOTIFY_ON_SUCCESS') === 'true';
  
  return {
    ...requiredProps,
    logLevel: parseInt(logLevel, 10),
    retryCount: parseInt(retryCount, 10),
    notificationEmail,
    notifyOnSuccess
  };
}

/**
 * スクリプトプロパティ設定テスト
 * 
 * 設定されているプロパティを確認し、問題がないかチェックします。
 * Phase 1のテスト用関数です。
 */
function testScriptConfig() {
  console.log('=== スクリプトプロパティ設定テスト ===');
  
  try {
    const config = getScriptConfig();
    
    console.log('✅ 必須プロパティ: すべて設定済み');
    console.log('');
    console.log('【設定内容】');
    console.log('- ACCESS_TOKEN:', config.accessToken ? '設定済み(先頭10文字: ' + config.accessToken.substring(0, 10) + '...)' : '未設定');
    console.log('- REFRESH_TOKEN:', config.refreshToken ? '設定済み(先頭10文字: ' + config.refreshToken.substring(0, 10) + '...)' : '未設定');
    console.log('- TARGET_SPREADSHEET_ID:', config.targetSpreadsheetId);
    console.log('- TARGET_SHEET_NAME:', config.targetSheetName);
    console.log('- SHOP_MASTER_SPREADSHEET_ID:', config.shopMasterSpreadsheetId);
    console.log('- SHOP_MASTER_SHEET_NAME:', config.shopMasterSheetName);
    console.log('');
    console.log('【オプション設定】');
    console.log('- LOG_LEVEL:', config.logLevel, getLogLevelName(config.logLevel));
    console.log('- RETRY_COUNT:', config.retryCount);
    console.log('- NOTIFICATION_EMAIL:', config.notificationEmail);
    console.log('- NOTIFY_ON_SUCCESS:', config.notifyOnSuccess);
    console.log('');
    console.log('✅ 設定確認完了!');
    
    return config;
    
  } catch (error) {
    console.error('❌ 設定エラー:', error.message);
    throw error;
  }
}

/**
 * ログレベル名を取得
 * 
 * @param {number} level - ログレベル
 * @return {string} ログレベル名
 */
function getLogLevelName(level) {
  switch (level) {
    case LOG_LEVEL.ALL:
      return '(全件出力)';
    case LOG_LEVEL.SAMPLE:
      return '(先頭3行のみ)';
    case LOG_LEVEL.NONE:
      return '(出力なし)';
    default:
      return '(不明)';
  }
}

// =============================================================================
// 日付計算ロジック
// =============================================================================

/**
 * 7日前(1週間前)の日付を取得
 * 
 * 例: 今日が2025/11/24(月)の場合 → 2025/11/17(月) 00:00:00を返す
 * 例: 今日が2025/11/25(火)の場合 → 2025/11/18(火) 00:00:00を返す
 * 
 * @return {Date} 7日前 00:00:00
 */
function getSevenDaysAgo() {
  const today = new Date();
  const sevenDaysAgo = new Date(today);
  sevenDaysAgo.setDate(today.getDate() - 7);
  sevenDaysAgo.setHours(0, 0, 0, 0);
  
  return sevenDaysAgo;
}

/**
 * 本日の終了時刻を取得
 * 
 * @return {Date} 本日 23:59:59
 */
function getTodayEnd() {
  const today = new Date();
  today.setHours(23, 59, 59, 999);
  
  return today;
}

/**
 * 日付を YYYY-MM-DD HH:MM:SS 形式にフォーマット
 * 
 * ネクストエンジンAPIの日時パラメータ形式に合わせます。
 * 
 * @param {Date} date - フォーマットする日付
 * @return {string} フォーマット済み日付文字列
 */
function formatDateTimeForNE(date) {
  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, '0');
  const day = String(date.getDate()).padStart(2, '0');
  const hours = String(date.getHours()).padStart(2, '0');
  const minutes = String(date.getMinutes()).padStart(2, '0');
  const seconds = String(date.getSeconds()).padStart(2, '0');
  
  return `${year}-${month}-${day} ${hours}:${minutes}:${seconds}`;
}

/**
 * 検索期間の日付範囲を取得
 * 
 * 7日前(1週間前) 00:00:00 から 本日 23:59:59 までの範囲を返します。
 * 
 * @return {Object} {startDate: Date, endDate: Date, startDateStr: string, endDateStr: string}
 */
function getSearchDateRange() {
  const startDate = getSevenDaysAgo();
  const endDate = getTodayEnd();
  
  return {
    startDate,
    endDate,
    startDateStr: formatDateTimeForNE(startDate),
    endDateStr: formatDateTimeForNE(endDate)
  };
}

/**
 * 日付計算テスト
 * 
 * Phase 1のテスト用関数です。
 */
function testDateCalculation() {
  console.log('=== 日付計算テスト ===');
  
  const today = new Date();
  console.log('今日:', today.toLocaleString('ja-JP'));
  console.log('曜日:', ['日', '月', '火', '水', '木', '金', '土'][today.getDay()]);
  console.log('');
  
  const sevenDaysAgo = getSevenDaysAgo();
  console.log('7日前(1週間前):', sevenDaysAgo.toLocaleString('ja-JP'));
  console.log('');
  
  const todayEnd = getTodayEnd();
  console.log('本日終了時刻:', todayEnd.toLocaleString('ja-JP'));
  console.log('');
  
  const dateRange = getSearchDateRange();
  console.log('【検索期間】');
  console.log('開始:', dateRange.startDateStr);
  console.log('終了:', dateRange.endDateStr);
  console.log('');
  
  // 日数計算
  const diffTime = dateRange.endDate - dateRange.startDate;
  const diffDays = Math.ceil(diffTime / (1000 * 60 * 60 * 24));
  console.log('検索期間日数:', diffDays, '日間');
  console.log('');
  console.log('✅ 日付計算テスト完了!');
}

// =============================================================================
// ログ出力制御
// =============================================================================

/**
 * ログ出力関数
 * 
 * LOG_LEVELの設定に応じてログ出力を制御します。
 * 
 * @param {string} message - ログメッセージ
 * @param {number} level - 出力レベル(省略時は LOG_LEVEL.ALL)
 */
function logMessage(message, level = LOG_LEVEL.ALL) {
  const config = getScriptConfig();
  
  // 設定されたログレベル以下の場合のみ出力
  if (config.logLevel <= level) {
    console.log(message);
  }
}

/**
 * データログ出力関数
 * 
 * LOG_LEVELの設定に応じてデータの出力を制御します。
 * - LOG_LEVEL.ALL (1): 全件出力
 * - LOG_LEVEL.SAMPLE (2): 先頭3行のみ出力
 * - LOG_LEVEL.NONE (3): 出力なし
 * 
 * @param {Array} dataArray - 出力するデータ配列
 * @param {string} title - データのタイトル
 */
function logData(dataArray, title = 'データ') {
  const config = getScriptConfig();
  
  if (config.logLevel === LOG_LEVEL.NONE) {
    return; // 出力なし
  }
  
  console.log(`=== ${title} (全${dataArray.length}件) ===`);
  
  if (config.logLevel === LOG_LEVEL.ALL) {
    // 全件出力
    dataArray.forEach((item, index) => {
      console.log(`[${index + 1}]`, JSON.stringify(item, null, 2));
    });
  } else if (config.logLevel === LOG_LEVEL.SAMPLE) {
    // 先頭3行のみ出力
    const sampleData = dataArray.slice(0, 3);
    sampleData.forEach((item, index) => {
      console.log(`[${index + 1}]`, JSON.stringify(item, null, 2));
    });
    
    if (dataArray.length > 3) {
      console.log(`... 残り ${dataArray.length - 3} 件は省略`);
    }
  }
  
  console.log('');
}

/**
 * ログ出力テスト
 * 
 * Phase 1のテスト用関数です。
 */
function testLogOutput() {
  console.log('=== ログ出力テスト ===');
  
  const config = getScriptConfig();
  console.log('現在のログレベル:', config.logLevel, getLogLevelName(config.logLevel));
  console.log('');
  
  // テストデータ
  const testData = [
    { id: 1, name: 'テスト商品1', price: 1000 },
    { id: 2, name: 'テスト商品2', price: 2000 },
    { id: 3, name: 'テスト商品3', price: 3000 },
    { id: 4, name: 'テスト商品4', price: 4000 },
    { id: 5, name: 'テスト商品5', price: 5000 }
  ];
  
  // ログレベル別出力テスト
  console.log('【通常メッセージ】');
  logMessage('これは通常メッセージです', LOG_LEVEL.ALL);
  console.log('');
  
  console.log('【データログ】');
  logData(testData, 'テストデータ');
  
  console.log('✅ ログ出力テスト完了!');
  console.log('');
  console.log('【確認事項】');
  console.log('- LOG_LEVEL=1 の場合: 全5件が表示される');
  console.log('- LOG_LEVEL=2 の場合: 先頭3件のみ表示される');
  console.log('- LOG_LEVEL=3 の場合: データログは表示されない');
}

// =============================================================================
// Phase 1 統合テスト
// =============================================================================

/**
 * Phase 1 統合テスト
 * 
 * Phase 1で実装した全機能をテストします。
 */
function testPhase1() {
  console.log('╔════════════════════════════════════════════════════════════╗');
  console.log('║  Phase 1: 基盤構築 - 統合テスト                           ║');
  console.log('╚════════════════════════════════════════════════════════════╝');
  console.log('');
  
  try {
    // 1. スクリプトプロパティテスト
    console.log('【1】スクリプトプロパティ設定テスト');
    testScriptConfig();
    console.log('');
    
    // 2. 日付計算テスト
    console.log('【2】日付計算テスト');
    testDateCalculation();
    console.log('');
    
    // 3. ログ出力テスト
    console.log('【3】ログ出力テスト');
    testLogOutput();
    console.log('');
    
    console.log('╔════════════════════════════════════════════════════════════╗');
    console.log('║  ✅ Phase 1 統合テスト: すべて成功!                       ║');
    console.log('╚════════════════════════════════════════════════════════════╝');
    console.log('');
    console.log('【次のステップ】');
    console.log('Phase 2: 店舗マスタ連携の開発に進みます。');
    
  } catch (error) {
    console.error('');
    console.error('╔════════════════════════════════════════════════════════════╗');
    console.error('║  ❌ Phase 1 統合テスト: エラー発生                       ║');
    console.error('╚════════════════════════════════════════════════════════════╝');
    console.error('');
    console.error('エラー内容:', error.message);
    console.error('');
    console.error('【確認事項】');
    console.error('- スクリプトプロパティが正しく設定されているか');
    console.error('- 必須項目がすべて入力されているか');
    
    throw error;
  }
}