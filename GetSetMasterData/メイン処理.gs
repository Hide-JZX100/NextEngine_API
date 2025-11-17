/**
 * =============================================================================
 * メイン処理
 * =============================================================================
 * セット商品マスタをネクストエンジンAPIから取得し、
 * Googleスプレッドシートに書き込むメイン処理
 * 
 * 【処理フロー】
 * 1. 設定確認
 * 2. API認証確認
 * 3. セット商品マスタ全件取得(ページネーション対応)
 * 4. スプレッドシートに書き込み
 * 5. 処理結果サマリー表示
 * 
 * 【実行方法】
 * - 手動実行: GASエディタで updateSetGoodsMaster() を実行
 * - 定期実行: トリガーで毎日/毎週などのスケジュール設定
 * =============================================================================
 */

/**
 * セット商品マスタ更新処理(メイン関数)
 * 
 * この関数を実行すると、ネクストエンジンから最新のセット商品マスタを取得し、
 * スプレッドシートに書き込みます。
 * 
 * @return {Object} 処理結果 {success: boolean, message: string, summary: Object}
 */
function updateSetGoodsMaster() {
  console.log('');
  console.log('='.repeat(70));
  console.log('セット商品マスタ更新処理 開始');
  console.log('実行日時:', new Date().toLocaleString('ja-JP'));
  console.log('='.repeat(70));
  console.log('');
  
  const startTime = new Date().getTime();
  
  try {
    // Step 1: 設定確認
    console.log('【Step 1】設定確認');
    const config = getSpreadsheetConfig();
    const logLevel = getLogLevel();
    console.log('スプレッドシートID:', config.spreadsheetId);
    console.log('シート名:', config.sheetName);
    console.log('ログレベル:', logLevel);
    console.log('✅ 設定確認完了');
    console.log('');
    
    // Step 2: API認証確認
    console.log('【Step 2】API認証確認');
    const props = PropertiesService.getScriptProperties();
    const accessToken = props.getProperty('ACCESS_TOKEN');
    const refreshToken = props.getProperty('REFRESH_TOKEN');
    
    if (!accessToken || !refreshToken) {
      throw new Error('認証トークンが見つかりません。testGenerateAuthUrl()で認証を実行してください。');
    }
    console.log('✅ 認証トークン確認完了');
    console.log('');
    
    // Step 3: セット商品マスタ全件取得
    console.log('【Step 3】セット商品マスタ取得');
    const allData = fetchAllSetGoodsMaster();
    console.log('');
    
    // Step 4: スプレッドシートに書き込み
    console.log('【Step 4】スプレッドシートへの書き込み');
    const writeResult = writeDataToSheet(allData);
    
    if (!writeResult.success) {
      throw new Error('スプレッドシートへの書き込みに失敗しました: ' + writeResult.message);
    }
    console.log('✅ 書き込み完了:', writeResult.message);
    console.log('');
    
    // Step 5: 処理完了
    const endTime = new Date().getTime();
    const duration = endTime - startTime;
    
    console.log('='.repeat(70));
    console.log('✅ セット商品マスタ更新処理 完了');
    console.log('='.repeat(70));
    console.log('');
    console.log('【最終結果】');
    console.log('取得件数:', allData.length, '件');
    console.log('書き込み行数:', writeResult.rowCount, '行');
    console.log('総処理時間:', (duration / 1000).toFixed(2), '秒');
    console.log('完了日時:', new Date().toLocaleString('ja-JP'));
    console.log('');
    
    return {
      success: true,
      message: '処理が正常に完了しました',
      summary: {
        dataCount: allData.length,
        writeCount: writeResult.rowCount,
        duration: duration,
        completedAt: new Date().toLocaleString('ja-JP')
      }
    };
    
  } catch (error) {
    const endTime = new Date().getTime();
    const duration = endTime - startTime;
    
    console.error('');
    console.error('='.repeat(70));
    console.error('❌ セット商品マスタ更新処理 失敗');
    console.error('='.repeat(70));
    console.error('');
    console.error('【エラー内容】');
    console.error('エラーメッセージ:', error.message);
    console.error('エラー発生時刻:', new Date().toLocaleString('ja-JP'));
    console.error('処理時間:', (duration / 1000).toFixed(2), '秒');
    console.error('');
    console.error('【トラブルシューティング】');
    console.error('1. 認証トークンの確認: testApiConnection()');
    console.error('2. スプレッドシート設定の確認: testSpreadsheetConfig()');
    console.error('3. ネクストエンジンのAPI権限確認');
    console.error('');
    
    return {
      success: false,
      message: error.message,
      summary: {
        duration: duration,
        failedAt: new Date().toLocaleString('ja-JP')
      }
    };
  }
}

/**
 * 処理状況確認
 * 現在の設定と認証状態を一括確認
 */
function checkSystemStatus() {
  console.log('='.repeat(70));
  console.log('システム状態確認');
  console.log('='.repeat(70));
  console.log('');
  
  try {
    // 1. スクリプトプロパティ確認
    console.log('【1. スクリプトプロパティ】');
    showAllConfig();
    console.log('');
    
    // 2. スプレッドシート接続確認
    console.log('【2. スプレッドシート接続】');
    testSpreadsheetConfig();
    console.log('');
    
    // 3. API認証確認
    console.log('【3. API認証】');
    testApiConnection();
    console.log('');
    
    console.log('='.repeat(70));
    console.log('✅ システム状態確認完了 - すべて正常です');
    console.log('='.repeat(70));
    console.log('');
    console.log('【次のステップ】');
    console.log('メイン処理を実行できます: updateSetGoodsMaster()');
    
  } catch (error) {
    console.error('');
    console.error('='.repeat(70));
    console.error('❌ システム状態確認でエラーが発生しました');
    console.error('='.repeat(70));
    console.error('');
    console.error('エラー:', error.message);
    console.error('');
    console.error('上記のエラーを解決してから、メイン処理を実行してください。');
  }
}

/**
 * クイックテスト
 * 少量のデータで動作確認(開発・テスト用)
 */
function quickTest() {
  console.log('=== クイックテスト(1ページのみ) ===');
  console.log('');
  
  try {
    // 1ページ(最大1000件)のみ取得
    console.log('【Step 1】データ取得(1ページのみ)');
    const response = fetchSetGoodsMaster(0, 1000);
    console.log('取得件数:', response.data.length, '件');
    console.log('');
    
    // スプレッドシートに書き込み
    console.log('【Step 2】スプレッドシート書き込み');
    const writeResult = writeDataToSheet(response.data);
    console.log('書き込み結果:', writeResult.message);
    console.log('');
    
    // ログレベルに応じてデータ表示
    const logLevel = getLogLevel();
    logApiData(response.data, logLevel, 'クイックテストデータ');
    
    console.log('');
    console.log('✅ クイックテスト完了!');
    console.log('スプレッドシートを確認してください。');
    console.log('');
    console.log('【注意】');
    console.log('これは1ページ(最大1000件)のみのテストです。');
    console.log('全件処理を実行する場合は updateSetGoodsMaster() を実行してください。');
    
  } catch (error) {
    console.error('❌ クイックテスト失敗:', error.message);
    throw error;
  }
}