/**
 * =============================================================================
 * API呼び出し
 * =============================================================================
 * ネクストエンジンAPIを呼び出す基本関数群
 * 
 * 【主な機能】
 * - セット商品マスタAPIの呼び出し
 * - ページネーション対応(offset/limit方式)
 * - トークン自動更新
 * - エラーハンドリング
 * =============================================================================
 */

// ネクストエンジンAPIのベースURL
const NE_API_URL = 'https://api.next-engine.org';

/**
 * セット商品マスタを取得(1ページ分)
 * 
 * @param {number} offset - 取得開始位置(0始まり)
 * @param {number} limit - 取得件数(最大1000)
 * @return {Object} APIレスポンス {result, count, data, access_token, refresh_token}
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
    
    // リクエストパラメータ
    const payload = {
      'access_token': accessToken,
      'refresh_token': refreshToken,
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
 * セット商品マスタを全件取得(ページネーション対応)
 * 
 * @return {Array} 全データの配列
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
      const estimatedTotal = currentCount === limit ? '不明(継続中)' : allData.length;
      logPaginationInfo(apiCallCount, '?', currentCount, allData.length);
      
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
 * API呼び出しテスト(1ページのみ)
 * 最初の1000件(または全件)を取得してログ出力
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
 * API呼び出しテスト(全件取得)
 * 全データを取得してログ出力(実際の書き込みは行わない)
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