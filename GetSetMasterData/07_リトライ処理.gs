/**
 * @fileoverview リトライ処理モジュール
 * 
 * API 呼び出しなどの不安定なネットワーク通信を伴う処理において、
 * 失敗時に自動的に再試行（リトライ）を行うためのユーティリティを提供します。
 * 
 * 主な機能:
 * - 指数バックオフ（Exponential Backoff）による待機時間の調整
 * - リトライ回数の制限とエラー種別に基づいた継続判定
 * - 高階関数を用いた既存処理へのリトライ機能の付与
 * 
 * 管理する設定項目（スクリプトプロパティ）:
 * - MAX_RETRY_COUNT: 最大リトライ回数（デフォルト: 3）
 * - RETRY_WAIT_TIME: 初回リトライまでの基本待機時間（秒）（デフォルト: 2）
 */

/**
 * リトライ設定を取得します。
 * 
 * スクリプトプロパティから設定値を読み取り、未設定の場合はデフォルト値を使用します。
 * 
 * @return {{maxRetryCount: number, retryWaitTime: number}} リトライ設定オブジェクト
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
 * 発生したエラーが再試行（リトライ）可能かどうかを判定します。
 * 
 * タイムアウトやサーバー一時不在などの一時的なエラーは true を返し、
 * 認証エラーや権限不足などの恒久的なエラーは false を返します。
 * 
 * @param {Error} error - 判定対象のエラーオブジェクト
 * @return {boolean} リトライ可能な場合は true、不可能な場合は false
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
 * 指数バックオフに基づき、現在のスレッドを一時停止（スリープ）させます。
 * 
 * 待機時間は `baseWaitTime * 2^retryCount` で計算されます。
 * これにより、短時間の連続リトライによるサーバー負荷を軽減します。
 * 
 * @param {number} retryCount - 現在のリトライ試行回数（0始まり）
 * @param {number} baseWaitTime - 基本となる待機時間（秒）
 */
function exponentialBackoff(retryCount, baseWaitTime) {
  // 待機時間 = 基本待機時間 * 2^リトライ回数
  const waitTime = baseWaitTime * Math.pow(2, retryCount);
  
  console.log(`${waitTime}秒待機します...`);
  Utilities.sleep(waitTime * 1000);
}

/**
 * 指定された関数をリトライ機能付きで実行します。
 * 
 * 関数が例外をスローした場合、`isRetryableError` で判定を行い、
 * リトライ可能な場合は指数バックオフを挟んで再実行します。
 * 
 * @param {Function} func - 実行対象の関数
 * @param {any[]} [args=[]] - 関数に渡す引数の配列
 * @param {string} [functionName='関数'] - ログ出力用の識別名
 * @return {any} 関数の実行結果
 * @throws {Error} 最大リトライ回数を超過した、またはリトライ不可のエラーが発生した場合
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
 * セット商品マスタ取得 API をリトライ機能付きで呼び出します。
 * 
 * `fetchSetGoodsMaster` をラップし、ネットワークエラー等の際、
 * 自動的に再試行を行います。
 * 
 * @param {number} [offset=0] - 取得開始位置
 * @param {number} [limit=1000] - 取得件数
 * @return {Object} API レスポンスオブジェクト
 * @throws {Error} リトライ上限に達しても成功しなかった場合
 */
function fetchSetGoodsMasterWithRetry(offset = 0, limit = 1000) {
  return executeWithRetry(
    fetchSetGoodsMaster,
    [offset, limit],
    'セット商品マスタ取得API'
  );
}

/**
 * リトライ機構の単体テストを実施します。
 * 
 * 意図的にエラーを発生させるスタブ関数を用い、
 * 期待通りに再試行が行われるか、およびリトライ不可設定で即時中断するかを検証します。
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
 * 実際の API 呼び出しを伴うリトライテストを実施します。
 * 
 * ネットワーク環境等に起因する一時的なエラーに対する挙動を確認するために使用します。
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
 * 現在のリトライ設定（回数、待機時間）をコンソールに出力します。
 * 
 * 設定値が正しく反映されているか、および待機時間の増加推移を確認するために使用します。
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