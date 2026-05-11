/**
 * @file 05_自動実行と通知.gs
 * @description ネクストエンジン 受注データ取得メイン・自動実行制御モジュール
 *
 * 【概要】
 * 本プロジェクトのメインエントリポイントです。各モジュール（通信・処理・出力）を連携させ、
 * 日次の自動実行、手動での再実行、および実行結果のロギングとエラー通知を管理します。
 * 
 * 【主な機能】
 * - dailyRun: 毎日、実行前日分の出荷確定データを自動取得し、シートへ追記します。
 * - manualRun: 任意の日付を指定してデータを取得・出力します。
 * - 統合エラーハンドリング: 実行中のエラーをキャッチし、通知およびログ記録を行います。
 *
 * 【運用上の注意】
 * - dailyRun は、Google Apps Script の「時間主導型トリガー」により、
 *   毎日深夜（API負荷の低い時間帯）に実行されることを想定しています。
 * - 実行結果はログシートおよび管理者へのメール通知（実装依存）にて確認可能です。
 *
 * 【依存関係】
 * - 01_設定ファイル.gs (設定値・プロパティ取得)
 * - 02_API通信.gs (受注データ取得)
 * - 03_データ処理.gs (データ整形)
 * - 04_スプレッドシート書き込み.gs (シート出力)
 * - ※ 外部関数: writeLog, sendErrorNotification (ログ出力・通知処理)
 *
 * 【認証仕様】
 * - 実行前に「00_NE_認証ライブラリ」にて、有効なトークンが保存されている必要があります。
 */

/**
 * 実行日の「前日」を YYYY-MM-DD 形式の文字列で取得するヘルパー関数です。
 *
 * 【詳細仕様】
 * - ネクストエンジンAPIの検索条件（`receive_order_send_date-eq`）に適合する日付形式を生成します。
 * - 深夜の自動実行時に「前日分」の確定データを特定するために使用されます。
 * - 月・日が1桁の場合は 0 埋め（パディング）を行い、常に10文字の固定長（YYYY-MM-DD）を返します。
 *
 * @returns {string} 前日の日付文字列（例: "2024-05-10"）
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
 * 毎日1回、自動実行トリガーから呼び出される本番用のメイン実行関数です。
 *
 * 【処理フロー】
 * 1. 日付特定: `getYesterdayString()` を呼び出し、処理対象日を前日に設定します。
 * 2. データ取得: `fetchOrdersByShipDate()` を実行し、APIから出荷確定データを全件取得します。
 * 3. 0件制御: 取得結果が0件の場合は、ログを記録して後続の処理（整形・書込）をスキップします。
 * 4. データ整形: `formatOrderData()` を用い、APIレスポンスをシート用の2次元配列に変換します。
 * 5. データ出力: `writeToSheet()` を実行し、スプレッドシートの最終行にデータを追記します。
 * 6. エラーハンドリング: 通信エラーや認証エラーが発生した場合、管理者への通知とログシートへの記録を行い、例外を再スローします。
 *
 * 【運用上の注意】
 * - Google Apps Script のトリガー設定にて「時間主導型」の「毎日」で設定してください。
 * - 推奨実行時刻は、日次処理が完了している深夜1時〜4時頃です。
 */
function dailyRun() {
  console.log('=== dailyRun 開始 ===');
  const targetDate = getYesterdayString();
  try {
    console.log('対象日付（前日）: ' + targetDate);

    // 1. データ取得
    const rawOrders = fetchOrdersByShipDate(targetDate);
    console.log('データ取得件数: ' + rawOrders.length + ' 件');

    if (rawOrders.length === 0) {
      writeLog({ targetDate: targetDate, funcName: 'dailyRun', count: 0, status: '0件' });
      console.log('取得データが0件のため終了します。');
      return;
    }

    // 2. データ整形
    const formattedData = formatOrderData(rawOrders);

    // 3. スプレッドシートへ書き込み
    writeToSheet(formattedData);

    console.log('スプレッドシートへの書き込みが完了しました。');
  } catch (error) {
    console.error('dailyRun 実行中にエラーが発生しました: ' + error.message);
    writeLog({ targetDate: targetDate, funcName: 'dailyRun', count: 0, status: 'エラー', errorMsg: error.message });
    sendErrorNotification({ targetDate: targetDate, funcName: 'dailyRun', errorMsg: error.message });
    throw error;
  } finally {
    console.log('=== dailyRun 完了 ===');
  }
}

/**
 * 任意の出荷確定日を指定してデータを取得・出力する手動実行用関数です。
 *
 * 【詳細仕様】
 * - 引数 `dateStr` が指定された場合はその日付を、省略された場合は `getYesterdayString()` の結果（前日）を対象にします。
 * - 内部ロジックは `dailyRun` と共通ですが、主に過去データのリカバリや欠損データの補填に使用することを目的としています。
 * - 既存のデータを上書きすることはありません。指定した日付のデータがシートの最後尾に追記されます。
 *
 * 【使用シーン】
 * - 運用開始前の過去ログ一括取得。
 * - トリガー実行時にエラーが発生し、特定の日のデータが欠損した場合の再実行。
 * - GASエディタから直接実行し、特定の日の動作を確認したい場合。
 * 
 * @param {string} [dateStr] - 対象日付（形式: "YYYY-MM-DD"）。省略時は前日
 */
function manualRun(dateStr) {
  console.log('=== manualRun 開始 ===');
  const targetDate = dateStr ? dateStr : getYesterdayString();
  try {
    console.log('対象日付: ' + targetDate);

    const rawOrders = fetchOrdersByShipDate(targetDate);
    console.log('データ取得件数: ' + rawOrders.length + ' 件');

    if (rawOrders.length === 0) {
      writeLog({ targetDate: targetDate, funcName: 'manualRun', count: 0, status: '0件' });
      console.log('取得データが0件のため終了します。');
      return;
    }

    const formattedData = formatOrderData(rawOrders);
    writeToSheet(formattedData);

    console.log('スプレッドシートへの書き込みが完了しました。');
  } catch (error) {
    console.error('manualRun 実行中にエラーが発生しました: ' + error.message);
    writeLog({ targetDate: targetDate, funcName: 'manualRun', count: 0, status: 'エラー', errorMsg: error.message });
    sendErrorNotification({ targetDate: targetDate, funcName: 'manualRun', errorMsg: error.message });
    throw error;
  } finally {
    console.log('=== manualRun 完了 ===');
  }
}

/**
 * 【開発・検証用】Phase 4：開発モードでのデータ取得テスト
 *
 * 開発用の店舗（楽天市場店等）を対象に、APIからのデータ取得のみを実行して内容を検証します。
 * スプレッドシートへの書き込みは行わないため、安全にAPIの通信状況やレスポンスの構造を確認できます。
 *
 * 【詳細仕様】
 * 1. `DEVELOPMENT_MODE` プロパティを一時的に `true` に書き換え、対象店舗を切り替えます。
 * 2. `fetchOrdersByShipDate()` を実行し、取得できた件数と先頭1件の生データをログ出力します。
 * 3. 処理完了後（またはエラー発生時）、プロパティを元の値に復元します。
 *
 * 【テスト時の注意点】
 * - 定数 `TEST_DATE` を、開発環境で確実にテストデータが存在する日付に書き換えて実行してください。
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

/**
 * `manualRun` を特定の日付で呼び出すためのショートカット関数です。
 * スクリプトエディタからワンクリックで特定の日の取得テストを行う際に利用します。
 */
function manualRuntest() {
  manualRun('2026-03-19');
}

/**
 * 【開発・検証用】Phase 5-3：全体統合動作テスト（開発モード）
 *
 * 開発モードを有効にした状態で `manualRun` を呼び出し、データの「取得」「整形」「書込」「ログ出力」「エラー通知」
 * の一連のパイプラインが正常に稼働するかを確認します。
 *
 * `testRun` との違いは、実際にスプレッドシートへの物理的な書き込みとログ記録までを行う点にあります。
 *
 * 【使用タイミング】
 * - リリース直前の最終チェック。
 * - ログシートや通知周りのロジックを変更した際のデバッグ。
 */
function testPhase5_3() {
  console.log('=== Phase 5-3 テスト開始 ===');
  const props = PropertiesService.getScriptProperties();
  const originalMode = props.getProperty('DEVELOPMENT_MODE');

  try {
    // 開発モードをONにする
    props.setProperty('DEVELOPMENT_MODE', 'true');
    console.log('開発モードに設定しました。');

    // 開発環境でテストデータが存在する日付
    const testDate = '2026-03-19';
    console.log('テスト対象日付: ' + testDate);

    // manualRunを実行して統合動作（取得・書込・ログ・エラー通知）を確認
    manualRun(testDate);

    console.log('=== テスト完了 ===');
    console.log('スプレッドシートの書き込み、および LOG シートへの記録を確認してください。');
  } catch (error) {
    console.error('テスト中にエラーが発生しました: ' + error.message);
  } finally {
    // 元のモードに戻す
    if (originalMode !== null) {
      props.setProperty('DEVELOPMENT_MODE', originalMode);
    } else {
      props.deleteProperty('DEVELOPMENT_MODE');
    }
  }
}