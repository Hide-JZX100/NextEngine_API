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

function setDailyTrigger() {
  
  // スクリプトプロパティからスケジュールを設定したい関数名を取得
  const functionToTrigger = PropertiesService.getScriptProperties().getProperty(PROPERTY_KEY); 
  
  if (!functionToTrigger) {
    Logger.log(`エラー: スクリプトプロパティの'${PROPERTY_KEY}'が設定されていません。トリガー設定を中断します。`);
    console.log(`スクリプトプロパティの設定方法:`);
    console.log(`1. プロジェクトの設定（歯車アイコン）を開く`);
    console.log(`2. 「スクリプトプロパティの追加」をクリック`);
    console.log(`3. キー: ${PROPERTY_KEY}, 値: 実行したい関数名 を入力`);
    return;
  }

  // 1. 既存のトリガーをすべて削除し、重複を防ぎます。（対象関数に紐づくもののみ）
  deleteTriggersForFunction(functionToTrigger);

  // 2. 実行したい時刻（[時, 分]）の配列を定義します。
  // ここに新しい時刻を自由に追加できます。
  const executionTimes = [
    [9, 0],   // 09:00
    [13, 50], // 13:50
    [16, 0]   // 16:00
    // 例: [20, 30] と追加すると 20:30 にも実行されます。
  ];
  
  // 3. 配列をループして毎日実行のトリガーを作成します。
  executionTimes.forEach(([hour, minute]) => {
    
    // 毎日指定時刻に実行するトリガーを作成
    // 重要: .everyDays(1) で毎日実行、.at()ではなく.atHour()と.nearMinute()を使用
    ScriptApp.newTrigger(functionToTrigger)
      .timeBased()
      .everyDays(1)  // 毎日実行
      .atHour(hour)  // 指定した時間
      .nearMinute(minute)  // 指定した分（±15分の幅あり）
      .create();
    
    Logger.log(`${String(hour).padStart(2, '0')}:${String(minute).padStart(2, '0')} の毎日実行トリガーを作成しました。`);
  });

  Logger.log(`'${functionToTrigger}' の毎日実行トリガーを ${executionTimes.length} 件作成完了。`);
  Logger.log(`明日から毎日指定時刻に自動実行されます。`);
}

/**
 * 現在のスクリプトプロジェクトに設定されている、
 * 特定の関数に紐づく既存のトリガーのみを削除します。
 * @param {string} functionName 削除対象のトリガーが実行する関数名
 */
function deleteTriggersForFunction(functionName) {
  const triggers = ScriptApp.getProjectTriggers();
  let deletedCount = 0;
  
  triggers.forEach(trigger => {
    // ターゲット関数に一致する場合のみ削除
    if (trigger.getHandlerFunction() === functionName) { 
      ScriptApp.deleteTrigger(trigger);
      deletedCount++;
    }
  });
  
  if (deletedCount > 0) {
    Logger.log(`関数'${functionName}'に紐づく既存のトリガーを ${deletedCount} 件削除しました。`);
  } else {
    Logger.log(`関数'${functionName}'に紐づく既存のトリガーはありませんでした。`);
  }
}

/**
 * 【デバッグ用関数】現在設定されているトリガーの一覧を表示
 * トリガーの設定状況を確認したい場合に実行してください。
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
    const triggerSource = trigger.getTriggerSource();
    const eventType = trigger.getEventType();
    
    Logger.log(`${index + 1}. 関数: ${functionName}`);
    Logger.log(`   ソース: ${triggerSource}`);
    Logger.log(`   イベント: ${eventType}`);
    Logger.log('---');
  });
}