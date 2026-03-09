/**
 * =============================================================================
 * ログ出力制御
 * =============================================================================
 * ログレベルに応じてデータの出力を制御する関数群
 * 
 * 【ログレベル】
 * 1: 全件ログ出力 - 全てのデータをログに出力(デバッグ用)
 * 2: 先頭3行のみ - 先頭3件のデータのみ出力(通常運用)
 * 3: ログ出力なし - データは出力せず、件数のみ表示(本番運用)
 * 
 * 【使用例】
 * const logLevel = getLogLevel();
 * logApiData(apiData, logLevel);
 * =============================================================================
 */

/**
 * APIデータをログ出力
 * ログレベルに応じて出力内容を制御
 * 
 * @param {Array} data - 出力するデータ配列
 * @param {number} logLevel - ログレベル(1:全件, 2:先頭3行, 3:なし)
 * @param {string} dataLabel - データのラベル(例: "セット商品マスタ")
 */
function logApiData(data, logLevel, dataLabel = 'データ') {
  if (!data || data.length === 0) {
    console.log(`${dataLabel}: 0件`);
    return;
  }
  
  const totalCount = data.length;
  
  switch (logLevel) {
    case 1:
      // 全件ログ出力
      console.log(`=== ${dataLabel} (全${totalCount}件) ===`);
      for (let i = 0; i < data.length; i++) {
        console.log(`[${i + 1}/${totalCount}]`, JSON.stringify(data[i], null, 2));
      }
      break;
      
    case 2:
      // 先頭3行のみ
      console.log(`=== ${dataLabel} (全${totalCount}件 / 先頭3件のみ表示) ===`);
      const displayCount = Math.min(3, totalCount);
      for (let i = 0; i < displayCount; i++) {
        console.log(`[${i + 1}/${totalCount}]`, JSON.stringify(data[i], null, 2));
      }
      if (totalCount > 3) {
        console.log(`... 他 ${totalCount - 3}件`);
      }
      break;
      
    case 3:
      // ログ出力なし
      console.log(`${dataLabel}: ${totalCount}件 (ログ出力OFF)`);
      break;
      
    default:
      console.log(`${dataLabel}: ${totalCount}件 (ログレベル不正: ${logLevel})`);
  }
}

/**
 * 処理結果のサマリーをログ出力
 * 
 * @param {Object} summary - サマリー情報
 * @param {number} summary.totalCount - 取得件数
 * @param {number} summary.writeCount - 書き込み件数
 * @param {number} summary.duration - 処理時間(ミリ秒)
 * @param {number} summary.apiCallCount - API呼び出し回数
 */
function logProcessSummary(summary) {
  console.log('');
  console.log('=== 処理サマリー ===');
  console.log('API呼び出し回数:', summary.apiCallCount || 0, '回');
  console.log('取得件数:', summary.totalCount || 0, '件');
  console.log('書き込み件数:', summary.writeCount || 0, '行');
  console.log('処理時間:', (summary.duration / 1000).toFixed(2), '秒');
  
  if (summary.errors && summary.errors.length > 0) {
    console.log('');
    console.log('=== エラー ===');
    summary.errors.forEach((error, index) => {
      console.log(`[${index + 1}]`, error);
    });
  }
}

/**
 * ページネーション情報をログ出力
 * 
 * @param {number} currentPage - 現在のページ番号
 * @param {number} totalPages - 総ページ数
 * @param {number} currentCount - 現在ページの件数
 * @param {number} accumulatedCount - 累積件数
 */
function logPaginationInfo(currentPage, totalPages, currentCount, accumulatedCount) {
  console.log('');
  console.log(`--- ページ ${currentPage}/${totalPages} ---`);
  console.log(`現在ページ件数: ${currentCount}件`);
  console.log(`累積取得件数: ${accumulatedCount}件`);
}

/**
 * ログ出力テスト
 * 各ログレベルでの出力を確認
 */
function testLogOutput() {
  console.log('=== ログ出力テスト ===');
  
  // テストデータ
  const testData = [
    { set_goods_id: 'TEST001', set_goods_name: 'テスト商品1' },
    { set_goods_id: 'TEST002', set_goods_name: 'テスト商品2' },
    { set_goods_id: 'TEST003', set_goods_name: 'テスト商品3' },
    { set_goods_id: 'TEST004', set_goods_name: 'テスト商品4' },
    { set_goods_id: 'TEST005', set_goods_name: 'テスト商品5' }
  ];
  
  console.log('');
  console.log('【ログレベル1: 全件出力】');
  logApiData(testData, 1, 'テストデータ');
  
  console.log('');
  console.log('【ログレベル2: 先頭3行のみ】');
  logApiData(testData, 2, 'テストデータ');
  
  console.log('');
  console.log('【ログレベル3: ログ出力なし】');
  logApiData(testData, 3, 'テストデータ');
  
  console.log('');
  console.log('【処理サマリーテスト】');
  const testSummary = {
    apiCallCount: 5,
    totalCount: 5000,
    writeCount: 5000,
    duration: 18234
  };
  logProcessSummary(testSummary);
  
  console.log('');
  console.log('【ページネーション情報テスト】');
  logPaginationInfo(1, 5, 1000, 1000);
  logPaginationInfo(2, 5, 1000, 2000);
  
  console.log('');
  console.log('✅ ログ出力テスト完了!');
}

/**
 * 現在の設定でのログレベルを確認
 */
function checkCurrentLogLevel() {
  console.log('=== 現在のログレベル確認 ===');
  
  const logLevel = getLogLevel();
  const logLevelText = {
    1: '全件ログ出力 (デバッグモード)',
    2: '先頭3行のみ出力 (通常モード)',
    3: 'ログ出力なし (本番モード)'
  };
  
  console.log('LOG_LEVEL:', logLevel);
  console.log('説明:', logLevelText[logLevel] || '不正な値');
  
  console.log('');
  console.log('【変更方法】');
  console.log('スクリプトプロパティのLOG_LEVELを変更してください');
  console.log('1: デバッグ時に全データを確認したい場合');
  console.log('2: 通常運用(推奨) - 先頭データのみ確認');
  console.log('3: 本番運用 - 件数のみ表示で高速化');
}