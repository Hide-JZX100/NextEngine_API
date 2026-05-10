/**
 * @file 01_基盤構築.gs
 * @description ネクストエンジン受注明細取得スクリプト - Phase 1: 基盤構築。
 * ネクストエンジンAPIから受注明細を取得し、Googleスプレッドシートに書き込むための基盤となる機能を提供します。
 * 
 * ### 必要なスクリプトプロパティ
 * - ACCESS_TOKEN, REFRESH_TOKEN: API認証用
 * - TARGET_SPREADSHEET_ID, TARGET_SHEET_NAME: 書き込み先SS
 * - SHOP_MASTER_SPREADSHEET_ID, SHOP_MASTER_SHEET_NAME: 店舗マスタSS
 * - LOG_LEVEL: ログ出力レベル(1:全件, 2:先頭3行, 3:なし)
 * 
 * ### ファイル構成と依存関係
 * 1. 01_基盤構築.gs (本ファイル: 共通基盤)
 * 2. 02_店舗マスタ連携.gs
 * 3. 03_ネクストエンジンAPI接続.gs (NEAuthライブラリ必須)
 * 4. 04_データ整形・書き込み.gs
 * 5. 05_メイン処理とエラーハンドリング.gs
 * 
 * ### 注意事項
 * - GAS実行時間制限（6分）に注意。
 * - タイムゾーンを Asia/Tokyo に設定してください。
 *
 * @version 1.0
 * @see testPhase1 - 初回セットアップ時の統合テスト用
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
 * @details 
 * アプリケーションの動作に必要な認証情報やスプレッドシートIDを、Google Apps Scriptの
 * スクリプトプロパティ(PropertiesService)から一括で読み込みます。
 * 
 * 必須項目（ACCESS_TOKEN等）が欠落している場合、実行を中断しエラーをスローすることで、
 * 設定漏れによる予期せぬ動作を未然に防ぎます。
 * 数値項目（LOG_LEVEL等）は、計算に使用できるよう適切な型変換（parseInt）を行ってから返却します。
 * 
 * @return {Object} 設定オブジェクト。以下のプロパティを含みます：
 *   - accessToken, refreshToken, targetSpreadsheetId, targetSheetName, 
 *     shopMasterSpreadsheetId, shopMasterSheetName, logLevel, retryCount, 
 *     notificationEmail, notifyOnSuccess
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
 * @details 
 * `getScriptConfig`を呼び出し、取得した設定値をコンソールにダンプします。
 * セキュリティの観点から、トークン類は先頭数文字のみを表示するように制御しています。
 * 新しい環境でのセットアップ時や、プロパティを変更した直後の疎通確認に使用します。
 * @return {Object} 取得に成功した設定オブジェクト
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
 * @details 
 * 数値で管理されているログレベル（1, 2, 3）を、人間に分かりやすい文字列に変換します。
 * 主にテスト関数内での表示用に使用されます。
 * @param {number} level - LOG_LEVEL 定数で定義された数値
 * @return {string} ログレベルの名称（例: "(全件出力)"）。該当がない場合は "(不明)"
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
 * @details 
 * 実行時点から起算して7日前の日付を生成し、時刻を 00:00:00.000 にリセットします。
 * ネクストエンジンAPIから「過去1週間分」のデータを取得する際の開始日として使用されます。
 * 
 * 例: 今日が 2025/11/24 の場合、2025/11/17 00:00:00 を返します。
 * 
 * @return {Date} 7日前の 00:00:00 に設定された Date オブジェクト
 * @note タイムゾーンはGASプロジェクトの設定（通常は Asia/Tokyo）に依存します。
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
 * @details 
 * 実行時点の当日の日付に対して、時刻を 23:59:59.999 に設定した Date オブジェクトを返します。
 * 検索期間の終点として使用されます。
 * 
 * @return {Date} 本日の 23:59:59 に設定された Date オブジェクト
 */
function getTodayEnd() {
  const today = new Date();
  today.setHours(23, 59, 59, 999);

  return today;
}

/**
 * 日付を YYYY-MM-DD HH:MM:SS 形式にフォーマット
 * 
 * @details 
 * ネクストエンジンAPI（receiveorder_row/search等）が要求する日時の文字列形式に変換します。
 * 月・日・時・分・秒は常に2桁（ゼロパディング）で出力されます。
 * 
 * @param {Date} date - フォーマットする日付
 * @return {string} `YYYY-MM-DD HH:mm:ss` 形式の文字列
 * @example
 * formatDateTimeForNE(new Date(2025, 0, 1, 9, 0, 0)) // "2025-01-01 09:00:00"
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
 * @details 
 * `getSevenDaysAgo` と `getTodayEnd` を組み合わせ、APIリクエストに必要な
 * 日付オブジェクトとフォーマット済み文字列をセットで生成します。
 * 
 * @return {Object} 期間情報オブジェクト：
 *   - startDate {Date}: 7日前 00:00
 *   - endDate {Date}: 本日 23:59
 *   - startDateStr {string}: API用開始日時文字列
 *   - endDateStr {string}: API用終了日時文字列
 * @see getSevenDaysAgo, getTodayEnd, formatDateTimeForNE
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
 * @details 
 * 生成された日付範囲が意図通り（7日間）であるか、フォーマットが正しいかを検証します。
 * コンソールに出力される開始/終了日時を手動で確認するために使用します。
 */
function testDateCalculation1() {
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
 * @details 
 * スクリプトプロパティの `LOG_LEVEL` 設定に基づき、メッセージを出力するか否かを判定します。
 * 設定値（1:ALL, 2:SAMPLE, 3:NONE）と比較し、引数の `level` が設定値以上の重要度
 * であれば `console.log` を実行します。
 *  
 * @param {string} message - ログメッセージ
 * @param {number} [level=LOG_LEVEL.ALL] - 出力レベル。省略時は常に全件出力対象
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
 * @details 
 * 配列形式のデータを、設定されたログレベルに応じて整形して出力します。
 * - ALL (1): 配列内のすべての要素を JSON.stringify 形式で出力します。
 * - SAMPLE (2): 最初の3要素のみ出力し、残りの件数を表示します。
 * - NONE (3): 何も出力しません。
 * 
 * 大量のAPIレスポンスをコンソールに表示すると、GASのログ制限に抵触したり
 * ブラウザが重くなるのを防ぐためのユーティリティです。
 * 
 * @param {Array} dataArray - 出力するデータ配列
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
 * @details 
 * ダミーデータを用いて、`logMessage` および `logData` が現在のログレベル設定に従って
 * 正しくフィルタリングされて表示されるかを確認します。
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
 * @details 
 * 基盤構築フェーズで実装した「設定取得」「日付計算」「ログ出力」の全機能を一括で実行します。
 * 各関数の戻り値に不整合がないか、例外が発生しないかを検証します。
 * 本番処理（Phase 5等）を実行する前に、このテストがパスすることを確認してください。
 * 
 * @throws {Error} いずれかのテストで問題が検出された場合
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
    testDateCalculation1();
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