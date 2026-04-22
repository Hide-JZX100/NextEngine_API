// ============================================================
// 05_自動実行と通知.gs
// ネクストエンジン Amazon受注データ取得 - メイン・自動実行処理
// ============================================================

/**
 * 実行日の「前日」を YYYY-MM-DD 形式の文字列で取得するヘルパー関数
 * 
 * @returns {string} 前日の日付文字列
 */
function getYesterdayString() {
  const date = new Date();
  date.setDate(date.getDate() - 1);
  const yyyy = date.getFullYear();
  const mm = ('0' + (date.getMonth() + 1)).slice(-2);
  const dd = ('0' + date.getDate()).slice(-2);
  return yyyy + '-' + mm + '-' + dd;
}

/**
 * 毎日トリガーから自動実行される本番用関数
 * 
 * 【処理フロー】
 * 1. 実行前日の日付を YYYY-MM-DD 形式で取得
 * 2. fetchOrdersByShipDate(昨日の日付) でデータ取得
 * 3. formatOrderData() で整形
 * 4. writeToSheet() でスプレッドシートへ書き込み
 * 5. 実行結果（取得件数・日付）をコンソールログに出力
 * 
 * 【使用タイミング】
 * - 毎日の時間主導型トリガー（推奨: 午前1時〜2時）
 */
function dailyRun() {
  console.log('=== dailyRun 開始 ===');
  try {
    const targetDate = getYesterdayString();
    console.log('対象日付（前日）: ' + targetDate);

    // 1. データ取得
    const rawOrders = fetchOrdersByShipDate(targetDate);
    console.log('データ取得件数: ' + rawOrders.length + ' 件');

    // 2. データ整形
    const formattedData = formatOrderData(rawOrders);

    // 3. スプレッドシートへ書き込み
    writeToSheet(formattedData);

    console.log('スプレッドシートへの書き込みが完了しました。');
  } catch (error) {
    console.error('dailyRun 実行中にエラーが発生しました: ' + error.message);
    // エラーハンドリング・通知機能（Phase 5）を実装した場合はここで通知を行う
    throw error;
  }
  console.log('=== dailyRun 完了 ===');
}

/**
 * 任意の出荷確定日を指定して手動実行する
 * 
 * 【処理フロー】
 * 1. 引数の日付文字列を取得（省略時は前日を対象とする）
 * 2. fetchOrdersByShipDate() でデータ取得
 * 3. formatOrderData() で整形
 * 4. writeToSheet() で書き込み
 * 
 * 【使用タイミング】
 * - 過去分のデータを再取得・再出力したい場合（手動実行）
 * 
 * @param {string} [dateStr] - 対象日付（形式: "YYYY-MM-DD"）。省略時は前日
 */
function manualRun(dateStr) {
  console.log('=== manualRun 開始 ===');
  try {
    // 引数がない場合は前日を取得
    const targetDate = dateStr ? dateStr : getYesterdayString();
    console.log('対象日付: ' + targetDate);

    const rawOrders = fetchOrdersByShipDate(targetDate);
    console.log('データ取得件数: ' + rawOrders.length + ' 件');

    const formattedData = formatOrderData(rawOrders);
    writeToSheet(formattedData);

    console.log('スプレッドシートへの書き込みが完了しました。');
  } catch (error) {
    console.error('manualRun 実行中にエラーが発生しました: ' + error.message);
  }
  console.log('=== manualRun 完了 ===');
}

/**
 * 【Phase 4 テスト用】動作確認関数
 * 開発環境（楽天市場店 店舗ID:2）で動作確認を行うテスト用関数
 * 
 * 【処理フロー】
 * 1. スクリプトプロパティ DEVELOPMENT_MODE を一時的に "true" に設定
 * 2. 任意の出荷確定日（TEST_DATE）でデータ取得
 * 3. 取得件数とデータ先頭1件の内容をコンソールログに出力（書き込みは行わない）
 * 4. DEVELOPMENT_MODE を元の状態に戻す
 * 
 * 【使用タイミング】
 * - 本番実行前の総合テストや、安全にデータ取得だけを確認したい場合
 */
function testRun() {
  console.log('=== testRun 開始 ===');

  const props = PropertiesService.getScriptProperties();
  const originalMode = props.getProperty('DEVELOPMENT_MODE');

  // ★ ここを開発環境（テストデータが存在する日）の出荷確定日に変更してテストしてください
  const TEST_DATE = '2026-03-19';

  try {
    // 開発モードをONにする
    props.setProperty('DEVELOPMENT_MODE', 'true');
    console.log('開発モードに設定しました。対象店舗ID: ' + getShopId());
    console.log('対象日付: ' + TEST_DATE);

    // データ取得のみ行い、スプレッドシートへの書き込みは行わない
    const rawOrders = fetchOrdersByShipDate(TEST_DATE);
    console.log('データ取得件数: ' + rawOrders.length + ' 件');

    if (rawOrders.length > 0) {
      console.log('先頭1件目のデータ:');
      console.log(JSON.stringify(rawOrders[0]));
    } else {
      console.log('指定した日付（' + TEST_DATE + '）に該当するデータはありませんでした。');
    }

  } catch (error) {
    console.error('testRun 実行中にエラーが発生しました: ' + error.message);
  } finally {
    // 元のモードに戻す
    if (originalMode !== null) {
      props.setProperty('DEVELOPMENT_MODE', originalMode);
    } else {
      props.deleteProperty('DEVELOPMENT_MODE');
    }
  }

  console.log('=== testRun 完了 ===');
}

function manualRuntest() {
  manualRun('2026-03-19');
}