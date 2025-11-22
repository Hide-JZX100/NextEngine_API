/**
 * =============================================================================
 * エラー通知
 * =============================================================================
 * 処理エラー発生時にメール通知を送信する機能
 * 
 * 【主な機能】
 * - エラー内容の整形
 * - メール通知の送信
 * - 通知先の管理
 * 
 * 【設定方法】
 * スクリプトプロパティに以下を追加:
 * - NOTIFICATION_EMAIL: 通知先メールアドレス(未設定時は実行者のメール)
 * =============================================================================
 */

/**
 * 通知先メールアドレスを取得
 * 
 * @return {string} メールアドレス
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
 * エラー通知メールを送信
 * 
 * @param {Error} error - エラーオブジェクト
 * @param {string} processName - 処理名
 * @param {Object} additionalInfo - 追加情報(オプション)
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
 * 成功通知メールを送信(オプション)
 * 
 * @param {Object} summary - 処理サマリー
 * @param {string} processName - 処理名
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
 * スプレッドシートのURLを取得
 * 
 * @return {string} スプレッドシートURL
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