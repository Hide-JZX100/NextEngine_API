/**
 * =============================================================================
 * 【トリガー設定スクリプト】 スケジュールの管理
 * =============================================================================
 *
 * 【目的】
 * スクリプトプロパティで指定された関数を、
 * 指定した時間に毎日自動で実行するための時間ベースのトリガー（スケジュール）を
 * 安全に設定・管理します。
 * * 【主な機能】
 * - 実行前に、既存の同名関数に紐づくトリガーのみを安全に削除します。
 * - executionTimes 配列に記述された時刻に合わせて、新しいトリガーを設定します。
 * * 【スクリプトプロパティの設定方法と操作説明】
 * 1. GASエディタで「プロジェクトの設定」を開く（歯車のアイコン）
 * 2. 「スクリプトプロパティ」セクションまでスクロール
 * 3. 「スクリプトプロパティの追加」をクリックし、以下のキーと値を設定
 * キー                     | 値
   -------------------------|------------------------------------
   TRIGGER_FUNCTION_NAME    | 実行したい関数名
 * 4. 実行したい時刻を executionTimes 配列に [時, 分] の形式で追加・編集します。
 * 5. この setTrigger() 関数を、実行したい時間より前にトリガーで毎日実行させてください。
*/

// スクリプトプロパティから関数名を取得するためのキー
const PROPERTY_KEY = 'TRIGGER_FUNCTION_NAME';

function setTrigger() {
  
  // スクリプトプロパティからスケジュールを設定したい関数名を取得
  const functionToTrigger = PropertiesService.getScriptProperties().getProperty(PROPERTY_KEY); 
  
  if (!functionToTrigger) {
    Logger.log(`エラー: スクリプトプロパティの'${PROPERTY_KEY}'が設定されていません。トリガー設定を中断します。`);
    // スクリプトプロパティの設定方法をユーザーに促すメッセージは、必要に応じて追加してください
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
  
  // 3. 配列をループしてトリガーを作成します。
  executionTimes.forEach(([hour, minute]) => {
    
    // 実行時刻を設定するためのDateオブジェクトを作成
    const triggerTime = new Date();
    triggerTime.setHours(hour);
    triggerTime.setMinutes(minute);
    triggerTime.setSeconds(0);
    triggerTime.setMilliseconds(0);

    // 時間ベースのトリガーを作成し、指定時刻に実行するように設定
    ScriptApp.newTrigger(functionToTrigger)
      .timeBased()
      .at(triggerTime)
      .create();
  });

  Logger.log(`'${functionToTrigger}' の日次トリガーを ${executionTimes.length} 件作成しました。`);
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
  
  Logger.log(`関数'${functionName}'に紐づく既存のトリガーを ${deletedCount} 件削除しました。`);
}
