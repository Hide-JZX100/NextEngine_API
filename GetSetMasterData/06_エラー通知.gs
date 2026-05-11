/**
 * @fileoverview エラー通知モジュール
 * 
 * システムの実行中に発生したエラーや、処理の完了を管理者に通知する機能を提供します。
 * 
 * 主な機能:
 * - エラー内容の整形（発生日時、対処法、デバッグ情報等の付与）
 * - Google Mail サービスを使用したメール通知の送信
 * - スクリプトプロパティによる通知先および通知条件の管理
 * 
 * 管理する設定項目:
 * - NOTIFICATION_EMAIL: 通知先メールアドレス(未設定時は実行者のメール)
 * - SEND_SUCCESS_NOTIFICATION: 成功時にも通知を送る場合は 'true'
 */

/**
 * 通知先メールアドレスを取得します。
 * 
 * スクリプトプロパティ `NOTIFICATION_EMAIL` が設定されていない場合は、
 * スクリプトの実行ユーザーのアドレスをデフォルトとして返します。
 * 
 * @return {string} 通知対象のメールアドレス
 */
function getNotificationEmail() {
  const props = PropertiesService.getScriptProperties();
  const configuredEmail = props.getProperty('NOTIFICATION_EMAIL');
  
  // 設定されている場合はそれを使用、なければ実行者のメール
  if (configuredEmail) {
    return configuredEmail;
  }
  
  return Session.getActiveUser().getEmail();
}

/**
 * エラー通知メールを送信します。
 * 
 * 発生したエラーオブジェクトからメッセージを抽出し、スクリプトへのリンクや
 * 一般的なトラブルシューティング手順を本文に含めて送信します。
 * 
 * @param {Error} error - エラーオブジェクト
 * @param {string} [processName='セット商品マスタ更新処理'] - エラーが発生した処理の名称
 * @param {Object} [additionalInfo={}] - 本文に追加で記載するデバッグ情報（キー・値のペア）
 */
function sendErrorNotification(error, processName = 'セット商品マスタ更新処理', additionalInfo = {}) {
  try {
    const recipient = getNotificationEmail();
    const subject = `[エラー通知] ${processName}が失敗しました`;
    
    // メール本文の作成
    let body = `
${processName}でエラーが発生しました。

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
【エラー情報】
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

エラーメッセージ:
${error.message}

エラー種類:
${error.name || 'Error'}

発生日時:
${new Date().toLocaleString('ja-JP')}

スクリプトID:
${ScriptApp.getScriptId()}

実行者:
${Session.getActiveUser().getEmail()}
`;

    // 追加情報がある場合は追加
    if (Object.keys(additionalInfo).length > 0) {
      body += `
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
【処理情報】
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
`;
      
      for (const [key, value] of Object.entries(additionalInfo)) {
        body += `
${key}: ${value}`;
      }
    }

    // トラブルシューティング情報
    body += `

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
【対処方法】
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

1. GASエディタを開く
   https://script.google.com/home/projects/${ScriptApp.getScriptId()}/edit

2. システム状態を確認
   関数: checkSystemStatus() を実行

3. エラー内容に応じて対処:
   
   【認証エラーの場合】
   - testGenerateAuthUrl() で認証URLを生成
   - ブラウザで認証を実行
   
   【API権限エラーの場合】
   - ネクストエンジンの管理画面でAPI権限を確認
   
   【スプレッドシートエラーの場合】
   - SPREADSHEET_ID と SHEET_NAME を確認
   - スプレッドシートへのアクセス権限を確認
   
   【その他のエラー】
   - GASエディタの実行ログを確認
   - エラーメッセージで検索

4. 問題解決後、再実行
   関数: updateSetGoodsMaster() を実行

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

このメールは自動送信されています。
`;

    // メール送信
    MailApp.sendEmail(recipient, subject, body);
    
    console.log('エラー通知メールを送信しました:', recipient);
    
  } catch (mailError) {
    console.error('エラー通知メールの送信に失敗しました:', mailError.message);
    // メール送信失敗は致命的ではないので、エラーを投げない
  }
}

/**
 * 成功通知メールを送信します（オプション）。
 * 
 * 処理が正常に完了した際に、取得件数や処理時間を記載したメールを送信します。
 * スクリプトプロパティ `SEND_SUCCESS_NOTIFICATION` が 'true' の場合のみ実行されます。
 * 
 * @param {{completedAt?: string, dataCount: number, writeCount: number, duration: number}} summary - 処理結果のサマリー情報
 * @param {string} [processName='セット商品マスタ更新処理'] - 完了した処理の名称
 */
function sendSuccessNotification(summary, processName = 'セット商品マスタ更新処理') {
  try {
    const props = PropertiesService.getScriptProperties();
    const sendSuccessMail = props.getProperty('SEND_SUCCESS_NOTIFICATION');
    
    // 成功通知が無効の場合は送信しない
    if (sendSuccessMail !== 'true') {
      console.log('成功通知は無効化されています');
      return;
    }
    
    const recipient = getNotificationEmail();
    const subject = `[完了通知] ${processName}が正常に完了しました`;
    
    const body = `
${processName}が正常に完了しました。

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
【処理結果】
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

完了日時:
${summary.completedAt || new Date().toLocaleString('ja-JP')}

取得件数:
${summary.dataCount || 0} 件

書き込み行数:
${summary.writeCount || 0} 行

処理時間:
${summary.duration ? (summary.duration / 1000).toFixed(2) : '不明'} 秒

スプレッドシート:
${getSpreadsheetUrl()}

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

このメールは自動送信されています。
`;

    MailApp.sendEmail(recipient, subject, body);
    console.log('成功通知メールを送信しました:', recipient);
    
  } catch (mailError) {
    console.error('成功通知メールの送信に失敗しました:', mailError.message);
  }
}

/**
 * 操作対象スプレッドシートの閲覧用URLを生成します。
 * 
 * `SPREADSHEET_ID` 設定値を元に、管理者が直接データを確認できるURLを構築します。
 * 
 * @return {string} スプレッドシートのURL。取得できない場合はエラーメッセージ。
 */
function getSpreadsheetUrl() {
  try {
    const config = getSpreadsheetConfig();
    return `https://docs.google.com/spreadsheets/d/${config.spreadsheetId}/edit`;
  } catch (error) {
    return 'URL取得失敗';
  }
}

/**
 * エラー通知のテスト
 * 
 * 擬似的なエラーを生成し、実際にメールが正しくフォーマットされて届くかを検証します。
 */
function testErrorNotification() {
  console.log('=== エラー通知テスト ===');
  
  // テスト用のエラーを作成
  const testError = new Error('これはテスト用のエラーメッセージです');
  testError.name = 'TestError';
  
  // 追加情報
  const additionalInfo = {
    '取得件数': '1000件',
    'API呼び出し回数': '1回',
    '処理時間': '5.23秒'
  };
  
  console.log('通知先メール:', getNotificationEmail());
  console.log('エラー通知メールを送信します...');
  
  sendErrorNotification(testError, 'テスト処理', additionalInfo);
  
  console.log('');
  console.log('✅ エラー通知テスト完了');
  console.log('メールボックスを確認してください。');
}

/**
 * 成功通知のテスト
 * 
 * 擬似的な成功サマリーを生成し、成功通知の設定（有効化）と送信を検証します。
 */
function testSuccessNotification() {
  console.log('=== 成功通知テスト ===');
  
  // まず、成功通知を有効化
  const props = PropertiesService.getScriptProperties();
  props.setProperty('SEND_SUCCESS_NOTIFICATION', 'true');
  console.log('成功通知を一時的に有効化しました');
  
  // テスト用のサマリーを作成
  const testSummary = {
    dataCount: 5269,
    writeCount: 5269,
    duration: 17690,
    completedAt: new Date().toLocaleString('ja-JP')
  };
  
  console.log('通知先メール:', getNotificationEmail());
  console.log('成功通知メールを送信します...');
  
  sendSuccessNotification(testSummary, 'テスト処理');
  
  console.log('');
  console.log('✅ 成功通知テスト完了');
  console.log('メールボックスを確認してください。');
  console.log('');
  console.log('【注意】');
  console.log('成功通知は通常は無効化しておくことをおすすめします。');
  console.log('無効化する場合は、SEND_SUCCESS_NOTIFICATION を削除してください。');
}

/**
 * 通知設定の確認
 * 
 * 現在の通知先メールアドレスや、成功通知のオン・オフ設定をコンソールに表示します。
 * 設定不足による「メールが届かない」問題を解決するために使用します。
 */
function checkNotificationSettings() {
  console.log('=== 通知設定確認 ===');
  
  const props = PropertiesService.getScriptProperties();
  const notificationEmail = props.getProperty('NOTIFICATION_EMAIL');
  const sendSuccessMail = props.getProperty('SEND_SUCCESS_NOTIFICATION');
  
  console.log('通知先メール:', notificationEmail || `未設定(実行者: ${Session.getActiveUser().getEmail()})`);
  console.log('成功通知:', sendSuccessMail === 'true' ? '有効' : '無効(推奨)');
  
  console.log('');
  console.log('【設定方法】');
  console.log('特定のメールアドレスに通知したい場合:');
  console.log('スクリプトプロパティに NOTIFICATION_EMAIL を追加');
  console.log('');
  console.log('成功時にも通知したい場合:');
  console.log('スクリプトプロパティに SEND_SUCCESS_NOTIFICATION = true を追加');
  console.log('(推奨: エラー時のみ通知、成功時は通知なし)');
}