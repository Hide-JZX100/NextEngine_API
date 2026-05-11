/**
 * @fileoverview API呼び出しモジュール
 * 
 * ネクストエンジン API との通信を担当します。
 * セット商品マスタの取得、ページネーション処理、およびアクセストークンの自動更新ロジックを提供します。
 * 
 * 主な機能:
 * - セット商品マスタ API (search) の呼び出し
 * - 大量データのページネーション取得 (offset/limit 方式)
 * - レスポンスに含まれる新しいトークンの自動保存
 * - ネットワークエラーや API エラーのハンドリング
 */

/** @constant {string} ネクストエンジンAPIのベースURL */
const NE_API_URL = 'https://api.next-engine.org';

/**
 * セット商品マスタを取得（1ページ分）
 * 
 * 指定された条件で API を呼び出し、セット商品マスタのデータを取得します。
 * レスポンスに新しいアクセストークンが含まれている場合、自動的にスクリプトプロパティを更新します。
 * 
 * @param {number} [offset=0] - 取得開始位置 (0始まり)
 * @param {number} [limit=1000] - 1ページあたりの取得件数 (最大1000)
 * @return {Object} APIレスポンスオブジェクト ({result: string, count: string, data: Object[]})
 * @throws {Error} 認証情報不足、HTTPエラー、またはAPIがエラーを返した場合
 */
function fetchSetGoodsMaster(offset = 0, limit = 1000) {
  try {
    const props = PropertiesService.getScriptProperties();
    const accessToken = props.getProperty('ACCESS_TOKEN');
    const refreshToken = props.getProperty('REFRESH_TOKEN');
    
    if (!accessToken || !refreshToken) {
      throw new Error('アクセストークンが見つかりません。認証を実行してください。');
    }
    
    // APIエンドポイント
    const url = `${NE_API_URL}/api_v1_master_setgoods/search`;
    
    // 取得するフィールドを指定
    const fields = [
      'set_goods_id',
      'set_goods_name',
      'set_goods_selling_price',
      'set_goods_detail_goods_id',
      'set_goods_detail_quantity'
    ].join(',');
    
    // リクエストパラメータ
    const payload = {
      'access_token': accessToken,
      'refresh_token': refreshToken,
      'fields': fields,
      'offset': offset.toString(),
      'limit': limit.toString()
    };
    
    // HTTPリクエストオプション
    const options = {
      'method': 'POST',
      'headers': {
        'Content-Type': 'application/x-www-form-urlencoded'
      },
      'payload': Object.keys(payload).map(key => 
        encodeURIComponent(key) + '=' + encodeURIComponent(payload[key])
      ).join('&'),
      'muteHttpExceptions': true // エラーレスポンスも取得
    };
    
    console.log(`API呼び出し: offset=${offset}, limit=${limit}`);
    
    // API呼び出し
    const response = UrlFetchApp.fetch(url, options);
    const responseCode = response.getResponseCode();
    const responseText = response.getContentText();
    
    // レスポンスコードチェック
    if (responseCode !== 200) {
      throw new Error(`HTTPエラー: ${responseCode} - ${responseText}`);
    }
    
    // JSONパース
    const responseData = JSON.parse(responseText);
    
    // API結果チェック
    if (responseData.result !== 'success') {
      throw new Error(`API呼び出し失敗: ${JSON.stringify(responseData)}`);
    }
    
    // トークンが更新された場合は保存
    if (responseData.access_token && responseData.refresh_token) {
      props.setProperties({
        'ACCESS_TOKEN': responseData.access_token,
        'REFRESH_TOKEN': responseData.refresh_token,
        'TOKEN_UPDATED_AT': new Date().getTime().toString()
      });
      console.log('トークンを更新しました');
    }
    
    console.log(`取得成功: ${responseData.count}件`);
    
    return responseData;
    
  } catch (error) {
    console.error('API呼び出しエラー:', error.message);
    throw error;
  }
}

/**
 * セット商品マスタを全件取得（ページネーション対応 + リトライ機能）
 * 
 * データの最終ページに到達するまで、繰り返し API を呼び出して全レコードを結合します。
 * 処理完了後、実行統計（件数、時間、呼び出し回数）をサマリーとして出力します。
 * ※内部で `fetchSetGoodsMasterWithRetry` を呼び出します。
 * 
 * @return {Object[]} 取得した全レコードの配列
 * @throws {Error} API取得プロセスで回復不能なエラーが発生した場合
 */
function fetchAllSetGoodsMasterWithRetry() {
  console.log('=== セット商品マスタ 全件取得開始 (リトライ機能付き) ===');
  
  const startTime = new Date().getTime();
  const limit = 1000; // 1回のAPI呼び出しで取得する件数
  let offset = 0;
  let allData = [];
  let apiCallCount = 0;
  
  try {
    while (true) {
      apiCallCount++;

      // ページ番号を先に表示
      console.log('');
      console.log(`--- ページ ${apiCallCount} 取得開始 ---`);

      // API呼び出し(リトライ機能付き)
      const response = fetchSetGoodsMasterWithRetry(offset, limit);
      const currentCount = parseInt(response.count);
      
      // データを追加
      if (response.data && response.data.length > 0) {
        allData = allData.concat(response.data);
      }

      // 取得結果を表示
      console.log(`取得成功: ${currentCount}件 (累積: ${allData.length}件)`);

      // ページネーション情報をログ出力
      // 最終ページかどうかで総ページ数を判定
      // const totalPages = currentCount < limit ? apiCallCount : '?';
      // logPaginationInfo(apiCallCount, totalPages, currentCount, allData.length);
      
      // 取得件数がlimitより少ない場合は最終ページ
      if (currentCount < limit) {
        console.log('最終ページに到達しました');
        break;
      }
      
      // 次のページへ
      offset += limit;
      
      // API制限を考慮して少し待機(1秒)
      Utilities.sleep(1000);
    }
    
    const endTime = new Date().getTime();
    const duration = endTime - startTime;
    
    // サマリー出力
    const summary = {
      apiCallCount: apiCallCount,
      totalCount: allData.length,
      writeCount: allData.length,
      duration: duration
    };
    
    logProcessSummary(summary);
    
    console.log('');
    console.log('✅ 全件取得完了');
    
    return allData;
    
  } catch (error) {
    console.error('❌ 全件取得エラー:', error.message);
    
    // エラー時も取得済みデータを返す
    console.log('取得済みデータ:', allData.length, '件');
    
    throw error;
  }
}

/**
 * セット商品マスタを全件取得（ページネーション対応）
 * 
 * 繰り返し API を呼び出して全レコードを取得します。
 * 現在はリトライ機能を持つ `fetchAllSetGoodsMasterWithRetry()` の使用を推奨します。
 * 
 * @deprecated リトライ機能付きの fetchAllSetGoodsMasterWithRetry() を推奨
 * @return {Object[]} 取得した全レコードの配列
 * @throws {Error} API取得中にエラーが発生した場合
 */
function fetchAllSetGoodsMaster() {
  console.log('=== セット商品マスタ 全件取得開始 ===');
  
  const startTime = new Date().getTime();
  const limit = 1000; // 1回のAPI呼び出しで取得する件数
  let offset = 0;
  let allData = [];
  let apiCallCount = 0;
  
  try {
    while (true) {
      apiCallCount++;
      
      // API呼び出し
      const response = fetchSetGoodsMaster(offset, limit);
      const currentCount = parseInt(response.count);
      
      // データを追加
      if (response.data && response.data.length > 0) {
        allData = allData.concat(response.data);
      }
      
      // ページネーション情報をログ出力
      // 最終ページかどうかで総ページ数を判定
      const totalPages = currentCount < limit ? apiCallCount : '?';
      logPaginationInfo(apiCallCount, totalPages, currentCount, allData.length);
      
      // 取得件数がlimitより少ない場合は最終ページ
      if (currentCount < limit) {
        console.log('最終ページに到達しました');
        break;
      }
      
      // 次のページへ
      offset += limit;
      
      // API制限を考慮して少し待機(1秒)
      Utilities.sleep(1000);
    }
    
    const endTime = new Date().getTime();
    const duration = endTime - startTime;
    
    // サマリー出力
    const summary = {
      apiCallCount: apiCallCount,
      totalCount: allData.length,
      writeCount: allData.length,
      duration: duration
    };
    
    logProcessSummary(summary);
    
    console.log('');
    console.log('✅ 全件取得完了');
    
    return allData;
    
  } catch (error) {
    console.error('❌ 全件取得エラー:', error.message);
    
    // エラー時も取得済みデータを返す
    console.log('取得済みデータ:', allData.length, '件');
    
    throw error;
  }
}

/**
 * API呼び出しテスト（1ページ分）
 * 
 * 最初の1ページ（最大1000件）を取得し、ログ出力制御モジュールを使用して
 * コンソールに内容を表示します。
 * 認証が通っているか、およびデータの構造が正しいかを確認するために使用します。
 * @throws {Error} 取得に失敗した場合
 */
function testFetchSetGoodsMaster() {
  console.log('=== セット商品マスタ取得テスト(1ページ) ===');
  
  try {
    // 1ページ目を取得
    const response = fetchSetGoodsMaster(0, 1000);
    
    console.log('');
    console.log('=== APIレスポンス ===');
    console.log('result:', response.result);
    console.log('count:', response.count);
    console.log('data件数:', response.data ? response.data.length : 0);
    
    // ログレベルに応じてデータを出力
    const logLevel = getLogLevel();
    console.log('');
    logApiData(response.data, logLevel, 'セット商品マスタ(1ページ目)');
    
    console.log('');
    console.log('✅ API呼び出しテスト成功!');
    
    if (parseInt(response.count) === 1000) {
      console.log('');
      console.log('【注意】取得件数が1000件です。');
      console.log('データがさらに存在する可能性があります。');
      console.log('全件取得する場合は testFetchAllSetGoodsMaster() を実行してください。');
    }
    
  } catch (error) {
    console.error('❌ API呼び出しテスト失敗:', error.message);
    throw error;
  }
}

/**
 * API呼び出しテスト（全件取得）
 * 
 * 全ページ分（全件）のデータを取得し、総件数をログに出力します。
 * スプレッドシートへの書き込みは行わずに、APIの負荷や取得時間の目安を確認するために使用します。
 * @throws {Error} 全件取得プロセスでエラーが発生した場合
 */
function testFetchAllSetGoodsMaster() {
  console.log('=== セット商品マスタ全件取得テスト ===');
  
  try {
    // 全件取得
    const allData = fetchAllSetGoodsMaster();
    
    console.log('');
    console.log('=== 取得結果 ===');
    console.log('総件数:', allData.length);
    
    // ログレベルに応じてデータを出力
    const logLevel = getLogLevel();
    console.log('');
    logApiData(allData, logLevel, 'セット商品マスタ(全件)');
    
    console.log('');
    console.log('✅ 全件取得テスト成功!');
    console.log('');
    console.log('【次のステップ】');
    console.log('スプレッドシートへの書き込みを含む完全な処理を実行する場合は');
    console.log('メイン処理関数を実行してください。');
    
  } catch (error) {
    console.error('❌ 全件取得テスト失敗:', error.message);
    throw error;
  }
}