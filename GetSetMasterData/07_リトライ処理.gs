/**
 * =============================================================================
 * リトライ処理
 * =============================================================================
 * API呼び出しなどの処理が失敗した場合に自動的に再試行する機能
 * 
 * 【主な機能】
 * - 指数バックオフによる再試行
 * - リトライ回数の制御
 * - エラー種別による再試行判定
 * 
 * 【設定方法】
 * スクリプトプロパティに以下を追加(オプション):
 * - MAX_RETRY_COUNT: 最大リトライ回数(デフォルト: 3)
 * - RETRY_WAIT_TIME: 初回リトライ待機時間(秒)(デフォルト: 2)
 * =============================================================================
 */

/**
 * リトライ設定を取得
 * 
 * @return {Object} リトライ設定 {maxRetryCount, retryWaitTime}
 */
function getRetryConfig() {
  const props = PropertiesService.getScriptProperties();
  
  const maxRetryCount = parseInt(props.getProperty('MAX_RETRY_COUNT')) || 3;
  const retryWaitTime = parseInt(props.getProperty('RETRY_WAIT_TIME')) || 2;
  
  return {
    maxRetryCount: maxRetryCount,
    retryWaitTime: retryWaitTime
  };
}

/**
 * エラーがリトライ可能かどうかを判定
 * 
 * @param {Error} error - エラーオブジェクト
 * @return {boolean} リトライ可能な場合true
 */
function isRetryableError(error) {
  const errorMessage = error.message.toLowerCase();
  
  // リトライ可能なエラーパターン
  const retryablePatterns = [
    'timeout',
    'timed out',
    'connection',
    'network',
    'temporarily',
    'try again',
    'service unavailable',
    'too many requests',
    '503',
    '504',
    '429'
  ];
  
  // リトライ不可のエラーパターン(認証エラーなど)
  const nonRetryablePatterns = [
    'unauthorized',
    'forbidden',
    '401',
    '403',
    'invalid token',
    'authentication failed',
    'アクセストークンが見つかりません'
  ];
  
  // リトライ不可パターンに一致する場合はfalse
  for (const pattern of nonRetryablePatterns) {
    if (errorMessage.includes(pattern)) {
      console.log('リトライ不可のエラー:', pattern);
      return false;
    }
  }
  
  // リトライ可能パターンに一致する場合はtrue
  for (const pattern of retryablePatterns) {
    if (errorMessage.includes(pattern)) {
      console.log('リトライ可能なエラー:', pattern);
      return true;
    }
  }
  
  // デフォルトはリトライ可能
  return true;
}

/**
 * 指数バックオフで待機
 * リトライ回数に応じて待機時間を増やす
 * 
 * @param {number} retryCount - 現在のリトライ回数
 * @param {number} baseWaitTime - 基本待機時間(秒)
 */
function exponentialBackoff(retryCount, baseWaitTime) {
  // 待機時間 = 基本待機時間 * 2^リトライ回数
  const waitTime = baseWaitTime * Math.pow(2, retryCount);
  
  console.log(`${waitTime}秒待機します...`);
  Utilities.sleep(waitTime * 1000);
}

/**
 * 関数をリトライ機能付きで実行
 * 
 * @param {Function} func - 実行する関数
 * @param {Array} args - 関数の引数
 * @param {string} functionName - 関数名(ログ用)
 * @return {any} 関数の実行結果
 * @throws {Error} 最大リトライ回数を超えた場合
 */
function executeWithRetry(func, args = [], functionName = '関数') {
  const config = getRetryConfig();
  let lastError = null;
  
  for (let retryCount = 0; retryCount <= config.maxRetryCount; retryCount++) {
    try {
      if (retryCount > 0) {
        console.log(`${functionName} リトライ ${retryCount}/${config.maxRetryCount}`);
      }
      
      // 関数を実行
      const result = func.apply(null, args);
      
      // 成功した場合
      if (retryCount > 0) {
        console.log(`✅ ${functionName} リトライ成功 (${retryCount}回目)`);
      }
      
      return result;
      
    } catch (error) {
      lastError = error;
      console.error(`❌ ${functionName} エラー (試行 ${retryCount + 1}/${config.maxRetryCount + 1}):`, error.message);
      
      // 最大リトライ回数に達した場合
      if (retryCount >= config.maxRetryCount) {
        console.error(`最大リトライ回数(${config.maxRetryCount}回)に達しました`);
        break;
      }
      
      // リトライ可能なエラーかチェック
      if (!isRetryableError(error)) {
        console.error('リトライ不可のエラーのため、処理を中断します');
        break;
      }
      
      // 指数バックオフで待機
      exponentialBackoff(retryCount, config.retryWaitTime);
    }
  }
  
  // すべてのリトライが失敗した場合
  throw lastError;
}

/**
 * セット商品マスタ取得(リトライ機能付き)
 * 
 * @param {number} offset - 取得開始位置
 * @param {number} limit - 取得件数
 * @return {Object} APIレスポンス
 */
function fetchSetGoodsMasterWithRetry(offset = 0, limit = 1000) {
  return executeWithRetry(
    fetchSetGoodsMaster,
    [offset, limit],
    'セット商品マスタ取得API'
  );
}

/**
 * リトライ処理のテスト
 * 意図的にエラーを発生させてリトライ動作を確認
 */
function testRetryMechanism() {
  console.log('=== リトライ処理テスト ===');
  console.log('');
  
  // テスト用: 2回失敗して3回目に成功する関数
  let attemptCount = 0;
  
  function testFunction() {
    attemptCount++;
    console.log(`testFunction 実行 (${attemptCount}回目)`);
    
    if (attemptCount < 3) {
      throw new Error('Test timeout error - リトライ可能なエラー');
    }
    
    return '成功!';
  }
  
  try {
    console.log('【テスト1】リトライ可能なエラー(2回失敗後、成功)');
    const result = executeWithRetry(testFunction, [], 'テスト関数');
    console.log('結果:', result);
    console.log('');
    
  } catch (error) {
    console.error('テスト1失敗:', error.message);
  }
  
  // リトライ不可のエラーテスト
  function nonRetryableFunction() {
    throw new Error('401 Unauthorized - リトライ不可のエラー');
  }
  
  try {
    console.log('【テスト2】リトライ不可のエラー(即座に中断)');
    executeWithRetry(nonRetryableFunction, [], 'テスト関数2');
    
  } catch (error) {
    console.log('期待通りリトライせず失敗:', error.message);
    console.log('');
  }
  
  console.log('✅ リトライ処理テスト完了');
  console.log('');
  console.log('【確認ポイント】');
  console.log('- テスト1: 2回失敗後、3回目で成功');
  console.log('- テスト2: リトライ不可エラーで即座に中断');
}

/**
 * 実際のAPI呼び出しでリトライをテスト
 */
function testApiRetry() {
  console.log('=== API呼び出しリトライテスト ===');
  console.log('');
  
  try {
    console.log('リトライ機能付きでAPI呼び出しを実行します');
    const response = fetchSetGoodsMasterWithRetry(0, 10);
    
    console.log('');
    console.log('✅ API呼び出し成功');
    console.log('取得件数:', response.count);
    
  } catch (error) {
    console.error('❌ API呼び出し失敗:', error.message);
  }
}

/**
 * リトライ設定の確認
 */
function checkRetrySettings() {
  console.log('=== リトライ設定確認 ===');
  
  const config = getRetryConfig();
  
  console.log('最大リトライ回数:', config.maxRetryCount, '回');
  console.log('初回リトライ待機時間:', config.retryWaitTime, '秒');
  console.log('');
  console.log('【待機時間の例】');
  console.log('1回目のリトライ:', config.retryWaitTime * 2, '秒待機');
  console.log('2回目のリトライ:', config.retryWaitTime * 4, '秒待機');
  console.log('3回目のリトライ:', config.retryWaitTime * 8, '秒待機');
  console.log('');
  console.log('【カスタマイズ方法】');
  console.log('スクリプトプロパティに以下を追加:');
  console.log('- MAX_RETRY_COUNT: 最大リトライ回数(推奨: 3)');
  console.log('- RETRY_WAIT_TIME: 初回リトライ待機時間(秒)(推奨: 2)');
}