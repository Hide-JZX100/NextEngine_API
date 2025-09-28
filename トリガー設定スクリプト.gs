/**
 * =============================================================================
 * 【トリガー設定スクリプト】 スケジュールの管理（毎日実行版）
 * =============================================================================
 *
 * 【目的】
 * スクリプトプロパティで指定された関数を、
 * 指定した時間に毎日自動で実行するための時間ベースのトリガー（スケジュール）を
 * 安全に設定・管理します。
 * 
 * 【主な機能】
 * - 実行前に、既存の同名関数に紐づくトリガーのみを安全に削除します。
 * - executionTimes 配列に記述された時刻に合わせて、毎日実行する新しいトリガーを設定します。
 * 
 * 【スクリプトプロパティの設定方法と操作説明】
 * 1. GASエディタで「プロジェクトの設定」を開く（歯車のアイコン）
 * 2. 「スクリプトプロパティ」セクションまでスクロール
 * 3. 「スクリプトプロパティの追加」をクリックし、以下のキーと値を設定
 * キー                     | 値
   -------------------------|------------------------------------
   TRIGGER_FUNCTION_NAME    | 実行したい関数名
 * 4. 実行したい時刻を executionTimes 配列に [時, 分] の形式で追加・編集します。
 * 5. この setDailyTrigger() 関数を、**設定変更時または初回に限り**手動で一度実行してください。
 * 6. 設定後は指定時刻に毎日自動実行されます。
*/

// スクリプトプロパティから関数名を取得するためのキー
const PROPERTY_KEY = 'TRIGGER_FUNCTION_NAME';

/**
 * 【初回設定用】正確な時刻実行トリガーを設定
 * この関数を1回だけ手動実行してください
 */
function setupPreciseMultipleTriggers() {
  
  // スクリプトプロパティから実行したい関数名を取得
  const functionToTrigger = PropertiesService.getScriptProperties().getProperty(PROPERTY_KEY); 
  
  if (!functionToTrigger) {
    Logger.log(`エラー: スクリプトプロパティの'${PROPERTY_KEY}'が設定されていません。`);
    console.log(`スクリプトプロパティの設定方法:`);
    console.log(`1. プロジェクトの設定（歯車アイコン）を開く`);
    console.log(`2. 「スクリプトプロパティの追加」をクリック`);
    console.log(`3. キー: ${PROPERTY_KEY}, 値: 実行したい関数名（例：actualTask）を入力`);
    return;
  }

  // 既存のトリガーを削除
  deleteTriggersForFunction('preciseTimeChecker');

  // 実行したい時刻（[時, 分]）の配列を定義
  const executionTimes = [
    [9, 0],   // 09:00
    [13, 50], // 13:50
    [16, 0]   // 16:00
    // 例: [20, 30] と追加すると 20:30 にも実行されます。
  ];

  // 実行時刻情報をスクリプトプロパティに保存
  PropertiesService.getScriptProperties().setProperty('EXECUTION_TIMES', JSON.stringify(executionTimes));
  
  // 各時刻に対してトリガーを作成
  executionTimes.forEach(([hour, minute]) => {
    // その時間帯（hour時台）に実行されるトリガーを作成
    ScriptApp.newTrigger('preciseTimeChecker')
      .timeBased()
      .everyDays(1)
      .atHour(hour)
      .nearMinute(minute)  // ±15分の範囲で実行
      .create();
    
    Logger.log(`${String(hour).padStart(2, '0')}:${String(minute).padStart(2, '0')} 用のトリガーを作成しました。`);
  });

  Logger.log(`${executionTimes.length}件の正確な時刻実行トリガーを作成完了。`);
  Logger.log(`実行対象関数: ${functionToTrigger}`);
}

/**
 * 【時刻チェック関数】トリガーから自動実行される
 * この関数は編集しないでください
 */
function preciseTimeChecker() {
  const now = new Date();
  const currentHour = now.getHours();
  const currentMinute = now.getMinutes();
  
  // スクリプトプロパティから設定を取得
  const executionTimesJson = PropertiesService.getScriptProperties().getProperty('EXECUTION_TIMES');
  const functionToTrigger = PropertiesService.getScriptProperties().getProperty(PROPERTY_KEY);
  
  if (!executionTimesJson || !functionToTrigger) {
    Logger.log('エラー: 設定情報が見つかりません。');
    return;
  }
  
  const executionTimes = JSON.parse(executionTimesJson);
  
  // 現在時刻が指定時刻と一致するかチェック
  const isTargetTime = executionTimes.some(([targetHour, targetMinute]) => {
    return currentHour === targetHour && currentMinute === targetMinute;
  });
  
  if (isTargetTime) {
    const timeString = `${String(currentHour).padStart(2, '0')}:${String(currentMinute).padStart(2, '0')}`;
    Logger.log(`正確な時刻 ${timeString} に実行開始`);
    
    // 指定された関数を実行
    try {
      eval(functionToTrigger + '()');
      Logger.log(`${functionToTrigger}() の実行が完了しました。`);
    } catch (error) {
      Logger.log(`エラー: ${functionToTrigger}() の実行に失敗しました。${error.message}`);
    }
  } else {
    // 指定時刻でない場合は何もしない（ログも出力しない）
    console.log(`時刻チェック: ${String(currentHour).padStart(2, '0')}:${String(currentMinute).padStart(2, '0')} - 対象外`);
  }
}

/**
 * 【実際の処理関数】ここに毎日指定時刻に実行したい処理を記述
 * スクリプトプロパティのTRIGGER_FUNCTION_NAMEに「actualTask」を設定してください
 */
function actualTask() {
  const now = new Date();
  const timeString = `${String(now.getHours()).padStart(2, '0')}:${String(now.getMinutes()).padStart(2, '0')}`;
  
  Logger.log(`=== 定時処理実行 ${timeString} ===`);
  console.log('正確な時刻での処理実行: ' + now);
  
  // ここに実際に実行したい処理を記述してください
  // 例：スプレッドシート操作
  // const sheet = SpreadsheetApp.openById('your-sheet-id').getActiveSheet();
  // sheet.appendRow([now, '定時処理実行']);
  
  // 例：メール送信
  // MailApp.sendEmail('your-email@example.com', '定時実行完了', `${timeString}の処理が完了しました。`);
  
  // 例：外部API呼び出し
  // const response = UrlFetchApp.fetch('https://api.example.com/data');
  // Logger.log('API応答: ' + response.getContentText());
}

/**
 * 特定関数のトリガーを削除
 */
function deleteTriggersForFunction(functionName) {
  const triggers = ScriptApp.getProjectTriggers();
  let deletedCount = 0;
  
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === functionName) { 
      ScriptApp.deleteTrigger(trigger);
      deletedCount++;
    }
  });
  
  if (deletedCount > 0) {
    Logger.log(`関数'${functionName}'に紐づく既存のトリガーを ${deletedCount} 件削除しました。`);
  }
}

/**
 * 【デバッグ用】現在のトリガー設定を確認
 */
function showCurrentTriggers() {
  const triggers = ScriptApp.getProjectTriggers();
  
  if (triggers.length === 0) {
    Logger.log('現在設定されているトリガーはありません。');
    return;
  }
  
  Logger.log(`=== 現在のトリガー設定 (${triggers.length}件) ===`);
  
  triggers.forEach((trigger, index) => {
    const functionName = trigger.getHandlerFunction();
    Logger.log(`${index + 1}. 実行関数: ${functionName}`);
  });
  
  // 設定情報も表示
  const executionTimesJson = PropertiesService.getScriptProperties().getProperty('EXECUTION_TIMES');
  const targetFunction = PropertiesService.getScriptProperties().getProperty(PROPERTY_KEY);
  
  if (executionTimesJson && targetFunction) {
    const executionTimes = JSON.parse(executionTimesJson);
    Logger.log(`対象関数: ${targetFunction}`);
    Logger.log(`実行予定時刻: ${executionTimes.map(([h, m]) => `${String(h).padStart(2, '0')}:${String(m).padStart(2, '0')}`).join(', ')}`);
  }
}