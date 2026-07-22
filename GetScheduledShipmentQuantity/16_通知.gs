/**
 * エラー発生時にスクリプト所有者へメール通知を送信する共通関数
 *
 * 【目的】
 * ネクストエンジンAPI連携処理（ウォームアップ・本番データ取得）で
 * 復旧困難なエラーが発生した際、即座に開発者へメールで知らせる。
 *
 * 【動作】
 * - 通知先は Session.getEffectiveUser().getEmail() で動的取得する
 * - severity に応じて件名の接頭辞を変更する ('WARNING' -> 【警告】, 'CRITICAL' -> 【緊急】)
 * - MailApp.sendEmail() を使用する
 * - メール送信自体が失敗した場合も、呼び出し元の処理を止めないよう
 *   try-catch で囲み、ログ出力のみに留める
 *
 * @param {string} severity - 'WARNING'（警告） または 'CRITICAL'（緊急）
 * @param {string} subject - 件名（重要度接頭辞を除いた本題部分）
 * @param {string} body - 本文（発生時刻・エラーメッセージ・関数名などを含める）
 * @returns {void}
 */
function notifyByEmail(severity, subject, body) {
  try {
    // 1. 重要度に応じた件名接頭辞の決定
    let prefix = '【通知】';
    if (severity === 'CRITICAL') {
      prefix = '【緊急】';
    } else if (severity === 'WARNING') {
      prefix = '【警告】';
    } else {
      console.warn(`[notifyByEmail] 不明な重要度(severity)が指定されました: ${severity}`);
    }

    const fullSubject = `${prefix} ${subject}`;

    // 2. スクリプト所有者のメールアドレスを動的に取得
    const recipient = Session.getEffectiveUser().getEmail();
    if (!recipient) {
      console.error('[notifyByEmail] 送信先メールアドレスを取得できませんでした (Session.getEffectiveUser().getEmail() が空)。送信をスキップします。');
      return;
    }

    // 3. メール送信
    MailApp.sendEmail(recipient, fullSubject, body);
    console.log(`[notifyByEmail] メールを送信しました。宛先: ${recipient}, 件名: ${fullSubject}`);

  } catch (error) {
    // 4. 送信失敗時も呼び出し元の処理を止めないよう例外はキャッチしてログ出力のみとする
    console.error(`[notifyByEmail] メール送信処理中にエラーが発生しました: ${error.message}`, error);
  }
}

/**
 * notifyByEmail() の動作確認用テスト関数
 *
 * 【検証方法】
 * GASエディタ上で本関数を選択して実行し、以下の2通のメールが届くことを目視確認する。
 * 1. 件名: 【警告】[テスト] ウォームアップ通知テスト
 * 2. 件名: 【緊急】[テスト] 本番データ更新エラー通知テスト
 */
function testNotifyByEmail() {
  console.log('--- notifyByEmail テスト開始 ---');

  const now = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd HH:mm:ss');

  // 1. WARNING テスト
  console.log('WARNINGメール送信テスト...');
  notifyByEmail(
    'WARNING',
    '[テスト] ウォームアップ通知テスト',
    `これはメール通知機能のテスト送信です（WARNING）。\n発生時刻: ${now}\n詳細: テストメッセージです。`
  );

  // 2. CRITICAL テスト
  console.log('CRITICALメール送信テスト...');
  notifyByEmail(
    'CRITICAL',
    '[テスト] 本番データ更新エラー通知テスト',
    `これはメール通知機能のテスト送信です（CRITICAL）。\n発生時刻: ${now}\n詳細: テストメッセージです。`
  );

  console.log('--- notifyByEmail テスト完了 ---');
}
