/**
 * =============================================================================
 * 【トリガー設定スクリプト 改良版】 当日/翌日選択可能なスケジュール管理
 * =============================================================================
 *
 * 【目的】
 * スクリプトプロパティで指定された関数を、
 * 指定した時間に自動で実行するための時間ベースのトリガー（スケジュール）を
 * 当日または翌日に設定できるように管理します。
 * 
 * 【主な機能】
 * - 実行前に、既存の同名関数に紐づくトリガーのみを安全に削除します。
 * - TRIGGER_MODE に応じて、当日または翌日にトリガーを設定します。
 *   - TODAY: 現在時刻より後の時刻を当日に設定
 *   - TOMORROW: すべての時刻を翌日に設定
 * 
 * 【スクリプトプロパティの設定方法】
 * 1. GASエディタで「プロジェクトの設定」を開く（歯車のアイコン）
 * 2. 「スクリプトプロパティ」セクションまでスクロール
 * 3. 「スクリプトプロパティの追加」をクリックし、以下のキーと値を設定
 * 
 * キー                     | 値
 * -------------------------|------------------------------------
 * TRIGGER_FUNCTION_NAME    | 実行したい関数名（例: updateStock）
 * TRIGGER_MODE             | TODAY または TOMORROW
 * 
 * 【TRIGGER_MODE の説明】
 * - TODAY: このスクリプト実行時刻より後の時刻のみ、当日に実行するトリガーを作成
 *   例）9:00に setTrigger() を実行した場合
 *       → 8:00はスキップ、10:00以降は当日に設定
 * 
 * - TOMORROW: すべての時刻を翌日に実行するトリガーとして作成
 *   例）23:00に setTrigger() を実行した場合
 *       → 8:00, 10:00... すべて翌日の時刻に設定
 * 
 * 【推奨運用】
 * - TODAY モード: 毎日早朝（例:0:30）に実行
 * - TOMORROW モード: 毎日夜（例:23:00）に実行
 */

// スクリプトプロパティのキー定義
const PROPERTY_KEY_FUNCTION = 'TRIGGER_FUNCTION_NAME';
const PROPERTY_KEY_MODE = 'TRIGGER_MODE';

function setTrigger() {
  
  // スクリプトプロパティから設定を取得
  const properties = PropertiesService.getScriptProperties();
  const functionToTrigger = properties.getProperty(PROPERTY_KEY_FUNCTION);
  const triggerMode = properties.getProperty(PROPERTY_KEY_MODE);
  
  // 必須プロパティのチェック
  if (!functionToTrigger) {
    Logger.log(`エラー: スクリプトプロパティ '${PROPERTY_KEY_FUNCTION}' が設定されていません。`);
    return;
  }
  
  if (!triggerMode) {
    Logger.log(`エラー: スクリプトプロパティ '${PROPERTY_KEY_MODE}' が設定されていません。`);
    Logger.log(`'TODAY' または 'TOMORROW' を設定してください。`);
    return;
  }
  
  // モードの検証
  if (triggerMode !== 'TODAY' && triggerMode !== 'TOMORROW') {
    Logger.log(`エラー: TRIGGER_MODE の値が不正です: '${triggerMode}'`);
    Logger.log(`'TODAY' または 'TOMORROW' を設定してください。`);
    return;
  }
  
  Logger.log(`=== トリガー設定開始 ===`);
  Logger.log(`実行関数: ${functionToTrigger}`);
  Logger.log(`実行モード: ${triggerMode}`);
  
  // 既存のトリガーを削除
  deleteTriggersForFunction(functionToTrigger);
  
  // 実行したい時刻（[時, 分]）の配列
  const executionTimes = [
    [8, 0],    // 8:00
    [10, 0],   // 10:00
    [12, 0],   // 12:00
    [14, 0],   // 14:00
    [17, 0],   // 17:00
    [21, 0],   // 21:00
    [23, 0]    // 23:00
  ];
  
  // 現在時刻を取得
  const now = new Date();
  let createdCount = 0;
  let skippedCount = 0;
  
  // 各時刻に対してトリガーを作成
  executionTimes.forEach(([hour, minute]) => {
    
    // トリガー実行時刻を設定
    const triggerTime = new Date();
    triggerTime.setHours(hour);
    triggerTime.setMinutes(minute);
    triggerTime.setSeconds(0);
    triggerTime.setMilliseconds(0);
    
    // モードに応じて日付を調整
    if (triggerMode === 'TOMORROW') {
      // 翌日に設定
      triggerTime.setDate(triggerTime.getDate() + 1);
    } else if (triggerMode === 'TODAY') {
      // 現在時刻より前の時刻はスキップ
      if (triggerTime <= now) {
        Logger.log(`  スキップ: ${hour}:${String(minute).padStart(2, '0')} (既に経過)`);
        skippedCount++;
        return;
      }
    }
    
    // トリガーを作成
    ScriptApp.newTrigger(functionToTrigger)
      .timeBased()
      .at(triggerTime)
      .create();
    
    const dateStr = `${triggerTime.getMonth() + 1}/${triggerTime.getDate()}`;
    const timeStr = `${hour}:${String(minute).padStart(2, '0')}`;
    Logger.log(`  作成: ${dateStr} ${timeStr}`);
    createdCount++;
  });
  
  Logger.log(`=== トリガー設定完了 ===`);
  Logger.log(`作成: ${createdCount} 件`);
  if (skippedCount > 0) {
    Logger.log(`スキップ: ${skippedCount} 件`);
  }
}

/**
 * 特定の関数に紐づく既存のトリガーをすべて削除します。
 * @param {string} functionName 削除対象のトリガーが実行する関数名
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
  
  Logger.log(`既存トリガー削除: ${deletedCount} 件`);
}