/**
 * 05_自動実行と通知.gs
 * Phase 4：自動実行と通知
 */

/**
 * 定期実行用メイン関数（トリガーから呼び出す）
 *
 * 【実行スケジュールと処理内容】
 * ┌──────────┬───────────────────────────────────────────────┐
 * │ 毎月 1日 │ 受注情報（前月16日〜末日）を取得して「追記」する │
 * │          │ キャンセル情報（前月1日〜末日）を取得           │
 * ├──────────┼───────────────────────────────────────────────┤
 * │ 毎月16日 │ 受注情報（当月1日〜15日）を取得して「上書き」する │
 * └──────────┴───────────────────────────────────────────────┘
 */
function scheduledRun() {
  const today = new Date();
  const date = today.getDate();

  // 動的トリガー経由でない場合（手動実行）はトリガー削除処理をスキップ
  const uid = PropertiesService.getScriptProperties()
    .getProperty('WARMUP_TRIGGER_UIDS');

  if (!uid) {
    console.log('手動実行のためトリガー削除処理をスキップします');
  } else {

    // warmUp が動的に作成したトリガーを削除（トリガーの蓄積を防ぐ）
    ScriptApp.getProjectTriggers().forEach(trigger => {
      if (
        trigger.getHandlerFunction() === 'scheduledRun' &&
        trigger.getEventType() === ScriptApp.EventType.CLOCK &&
        trigger.getTriggerSource() === ScriptApp.TriggerSource.CLOCK
      ) {
        // 月次固定トリガー（毎月1日・10日・20日）は削除しない
        // after() で作成された一回限りのトリガーのみ削除する
        const triggerUid = trigger.getUniqueId();

        const savedUids = (
          PropertiesService.getScriptProperties()
            .getProperty('WARMUP_TRIGGER_UIDS') || ''
        ).split(',');

        if (savedUids.includes(triggerUid)) {
          ScriptApp.deleteTrigger(trigger);
          console.log('動的トリガーを削除しました: ' + triggerUid);
        }
      }
    });
  }

  console.log(`定期実行開始 (実行日: ${date}日)`);

  // 対象日以外はスキップ（誤トリガー対策）
  if (date !== 1 && date !== 10 && date !== 20) {
    console.log('本日は定期実行対象日ではありません。処理をスキップします。');
    return;
  }

  // 受注情報は毎月1日・10日・20日すべてで実行
  try {
    updateOrders(null, null);
  } catch (e) {
    sendErrorNotification('updateOrders', e);
  }

  // キャンセル情報は1日のみ実行
  if (date === 1) {
    try {
      updateCancels(null, null);
    } catch (e) {
      sendErrorNotification('updateCancels', e);
    }
  }

  console.log('定期実行完了');
}

/**
 * 受注情報の手動実行用関数（任意の日付範囲で再取得したい場合に使用）
 * @param {string} startDateStr - 開始日文字列 'YYYY/MM/DD' 形式
 * @param {string} endDateStr - 終了日文字列 'YYYY/MM/DD' 形式
 * @param {boolean} append - 追記する場合は true
 *
 * 【使用例】
 * manualRun('2026/02/16', '2026/02/28', true) // 追記する
 */
function manualRun(startDateStr, endDateStr, append = false) {
  try {
    const start = new Date(startDateStr);
    const end = new Date(endDateStr);

    if (isNaN(start.getTime()) || isNaN(end.getTime())) {
      throw new Error('日付の形式が正しくありません。YYYY/MM/DD形式で指定してください。');
    }

    console.log(`手動実行(受注): ${startDateStr} - ${endDateStr} (追記=${append})`);
    updateOrders(start, end, append);
  } catch (e) {
    console.error('手動実行エラー:', e.message);
    throw e;
  }
}

/**
 * キャンセル情報の手動実行用関数（任意の受注日範囲で再取得したい場合に使用）
 * @param {string} startDateStr - 受注日の開始日文字列 'YYYY/MM/DD' 形式
 * @param {string} endDateStr - 受注日の終了日文字列 'YYYY/MM/DD' 形式
 *
 * 【使用例】
 * manualRunCancels('2026/02/01', '2026/02/28')
 */
function manualRunCancels(startDateStr, endDateStr) {
  try {
    const start = new Date(startDateStr);
    const end = new Date(endDateStr);

    if (isNaN(start.getTime()) || isNaN(end.getTime())) {
      throw new Error('日付の形式が正しくありません。YYYY/MM/DD形式で指定してください。');
    }

    console.log(`手動実行(キャンセル): ${startDateStr} - ${endDateStr}`);
    updateCancels(start, end);
  } catch (e) {
    console.error('手動実行エラー:', e.message);
    throw e;
  }
}

/**
 * エラー発生時にメール通知を送信する
 * @param {string} functionName - エラーが発生した関数名
 * @param {Error} error - エラーオブジェクト
 */
function sendErrorNotification(functionName, error) {
  const config = getConfig();
  let email = config.notificationEmail;
  if (!email) {
    email = Session.getActiveUser().getEmail();
  }

  const subject = '[エラー] 受注情報取得システム - ' + functionName;
  const body = `自動実行中にエラーが発生しました。

■ 発生日時
${new Date().toLocaleString('ja-JP')}

■ エラーが発生した処理
${functionName}

■ エラー内容
${error.message}
${error.stack ? error.stack : ''}

■ 対処方法
Google Apps Script エディタを開き、ログを確認してください。
必要に応じて、manualRun('YYYY/MM/DD', 'YYYY/MM/DD') または manualRunCancels() にて手動実行によるリカバリを行ってください。`;

  MailApp.sendEmail(email, subject, body);
  console.log(`エラーメールを送信しました: ${email}`);
}

/**
 * セットアップ状態を確認する診断関数
 * 初期設定後・トラブル発生時に実行して設定漏れを確認する
 */
function checkSetup() {
  console.log('=== セットアップ状態の確認 ===');
  const props = PropertiesService.getScriptProperties().getProperties();

  const checkProp = (key, name) => {
    if (props[key]) {
      console.log(`✅ ${name} (${key}): 設定済み`);
      return true;
    } else {
      console.error(`❌ ${name} (${key}): 未設定`);
      return false;
    }
  };

  checkProp('ACCESS_TOKEN', 'アクセストークン');
  checkProp('REFRESH_TOKEN', 'リフレッシュトークン');
  checkProp('TARGET_SPREADSHEET_ID', '書き込み先SS ID');
  checkProp('MASTER_SPREADSHEET_ID', '店舗マスタSS ID');
  checkProp('NOTIFICATION_EMAIL', '通知先メールアドレス');

  if (props['TARGET_SPREADSHEET_ID']) {
    try {
      const ss = SpreadsheetApp.openById(props['TARGET_SPREADSHEET_ID']);
      console.log(`✅ 書き込み先スプレッドシートへのアクセス: 成功 (${ss.getName()})`);
    } catch (e) {
      console.error(`❌ 書き込み先スプレッドシートへのアクセス: 失敗 (${e.message})`);
    }
  }

  if (props['MASTER_SPREADSHEET_ID']) {
    try {
      const ss = SpreadsheetApp.openById(props['MASTER_SPREADSHEET_ID']);
      console.log(`✅ 店舗マスタスプレッドシートへのアクセス: 成功 (${ss.getName()})`);
    } catch (e) {
      console.error(`❌ 店舗マスタスプレッドシートへのアクセス: 失敗 (${e.message})`);
    }
  }

  console.log('==============================');
}

/**
 * Phase 5 で指定されたスプレッドシートの確認項目用関数
 */
function checkSpreadsheet() {
  console.log('=== スプレッドシートの状態確認 ===');
  const config = getConfig();
  const ss = SpreadsheetApp.openById(config.targetSpreadsheetId);

  // 受注情報シート
  const orderSheet = ss.getSheetByName(config.sheetNameOrders);
  if (orderSheet) {
    const orderHeaders = orderSheet.getRange(1, 1, 1, orderSheet.getLastColumn()).getValues()[0];
    console.log(`✅ 受注情報シート: ヘッダー列数=${orderHeaders.length} (想定: ${ORDER_HEADERS.length}), データ行数=${orderSheet.getLastRow() - 1}`);
  } else {
    console.error(`❌ 受注情報シートが見つかりません: ${config.sheetNameOrders}`);
  }

  // キャンセル情報シート
  const cancelSheet = ss.getSheetByName(config.sheetNameCancel);
  if (cancelSheet) {
    const cancelHeaders = cancelSheet.getRange(1, 1, 1, cancelSheet.getLastColumn()).getValues()[0];
    console.log(`✅ キャンセル情報シート: ヘッダー列数=${cancelHeaders.length} (想定: ${CANCEL_HEADERS.length}), データ行数=${cancelSheet.getLastRow() - 1}`);
  } else {
    console.error(`❌ キャンセル情報シートが見つかりません: ${config.sheetNameCancel}`);
  }

  console.log('==============================');
}
