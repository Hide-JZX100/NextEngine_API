/**
 * @fileoverview トリガー管理モジュール
 * 
 * スクリプトの自動実行を制御するための「時間主導型（Time-driven）トリガー」の
 * 作成、一覧表示、削除などの管理機能を提供します。
 * 
 * 主な機能:
 * - メイン処理（データ更新）の定期実行予約
 * - トークン更新処理の定期実行予約（認証維持用）
 * - 既存トリガーの重複チェックと一覧表示
 * - 不要なトリガーの一括・個別削除
 * 
 * 推奨設定:
 * - メイン処理: 毎日午前3時実行(データ更新用)
 * - トークン更新: 毎日午前2時実行(認証維持用)
 */

/**
 * 毎日指定された時刻にメイン処理（updateSetGoodsMaster）を実行するトリガーを作成します。
 * 
 * 既に同じ関数のトリガーが存在する場合は、二重登録を防ぐため新規作成を行いません。
 * 
 * @param {number} [hour=3] - 実行時刻（0-23時）
 * @throws {Error} トリガー作成権限がない場合など
 */
function createDailyTrigger(hour = 3) {
  try {
    // 既存のトリガーを確認
    const existingTriggers = ScriptApp.getProjectTriggers();
    
    // 同じ関数のトリガーが既に存在するか確認
    for (const trigger of existingTriggers) {
      if (trigger.getHandlerFunction() === 'updateSetGoodsMaster') {
        console.log('既に定期実行トリガーが存在します');
        console.log('トリガーID:', trigger.getUniqueId());
        console.log('');
        console.log('新しいトリガーを作成する場合は、先に deleteAllTriggers() で既存トリガーを削除してください');
        return;
      }
    }
    
    // 新しいトリガーを作成
    ScriptApp.newTrigger('updateSetGoodsMaster')
      .timeBased()
      .atHour(hour)
      .everyDays(1)
      .create();
    
    console.log('✅ 定期実行トリガーを作成しました');
    console.log('関数: updateSetGoodsMaster');
    console.log('実行時刻: 毎日', hour, '時台');
    console.log('');
    console.log('【注意】');
    console.log('- 実行時刻は±15分の範囲で変動することがあります');
    console.log('- トリガーの確認: listAllTriggers()');
    console.log('- トリガーの削除: deleteAllTriggers()');
    
  } catch (error) {
    console.error('❌ トリガー作成エラー:', error.message);
    throw error;
  }
}

/**
 * 毎日指定された時刻にトークン更新処理（dailyTokenRefresh）を実行するトリガーを作成します。
 * 
 * ネクストエンジンのアクセストークン有効期限を維持するために重要です。
 * メイン処理の実行頻度に関わらず、このトリガーは毎日実行することを推奨します。
 * 
 * @param {number} [hour=2] - 実行時刻（0-23時）
 * @throws {Error} トリガー作成エラーが発生した場合
 */
function createTokenRefreshTrigger(hour = 2) {
  try {
    // 既存のトリガーを確認
    const existingTriggers = ScriptApp.getProjectTriggers();
    
    // 同じ関数のトリガーが既に存在するか確認
    for (const trigger of existingTriggers) {
      if (trigger.getHandlerFunction() === 'dailyTokenRefresh') {
        console.log('既にトークン更新トリガーが存在します');
        console.log('トリガーID:', trigger.getUniqueId());
        return;
      }
    }
    
    // 新しいトリガーを作成
    ScriptApp.newTrigger('dailyTokenRefresh')
      .timeBased()
      .atHour(hour)
      .everyDays(1)
      .create();
    
    console.log('✅ トークン更新トリガーを作成しました');
    console.log('関数: dailyTokenRefresh');
    console.log('実行時刻: 毎日', hour, '時台');
    console.log('');
    console.log('【重要】');
    console.log('このトリガーは認証トークンの有効期限を延長します。');
    console.log('メイン処理が週1回や月1回でも、このトリガーは毎日実行してください。');
    
  } catch (error) {
    console.error('❌ トリガー作成エラー:', error.message);
    throw error;
  }
}

/**
 * 毎週指定された曜日・時刻にメイン処理を実行するトリガーを作成します。
 * 
 * データの更新頻度が低い場合に適していますが、認証維持のための
 * `createTokenRefreshTrigger` は別途毎日実行するようにしてください。
 * 
 * @param {GoogleAppsScript.Script.WeekDay} [weekday=ScriptApp.WeekDay.MONDAY] - 実行する曜日
 * @param {number} [hour=3] - 実行時刻（0-23時）
 */
function createWeeklyTrigger(weekday = ScriptApp.WeekDay.MONDAY, hour = 3) {
  try {
    const weekdayNames = {
      [ScriptApp.WeekDay.MONDAY]: '月曜日',
      [ScriptApp.WeekDay.TUESDAY]: '火曜日',
      [ScriptApp.WeekDay.WEDNESDAY]: '水曜日',
      [ScriptApp.WeekDay.THURSDAY]: '木曜日',
      [ScriptApp.WeekDay.FRIDAY]: '金曜日',
      [ScriptApp.WeekDay.SATURDAY]: '土曜日',
      [ScriptApp.WeekDay.SUNDAY]: '日曜日'
    };
    
    // 既存のトリガーを確認
    const existingTriggers = ScriptApp.getProjectTriggers();
    
    for (const trigger of existingTriggers) {
      if (trigger.getHandlerFunction() === 'updateSetGoodsMaster') {
        console.log('既に定期実行トリガーが存在します');
        console.log('先に deleteAllTriggers() で削除してください');
        return;
      }
    }
    
    // 新しいトリガーを作成
    ScriptApp.newTrigger('updateSetGoodsMaster')
      .timeBased()
      .onWeekDay(weekday)
      .atHour(hour)
      .create();
    
    console.log('✅ 週次実行トリガーを作成しました');
    console.log('関数: updateSetGoodsMaster');
    console.log('実行日時: 毎週', weekdayNames[weekday], hour, '時台');
    console.log('');
    console.log('【注意】');
    console.log('週1回実行の場合でも、トークン更新トリガーは毎日実行してください');
    console.log('createTokenRefreshTrigger() で設定可能');
    
  } catch (error) {
    console.error('❌ トリガー作成エラー:', error.message);
    throw error;
  }
}

/**
 * 現在のプロジェクトに設定されている全てのトリガーをコンソールに表示します。
 * 
 * 各トリガーの ID、実行関数名、イベント種別などを確認できます。
 * 設定漏れや不要なトリガーの有無を確認するために使用します。
 */
function listAllTriggers() {
  console.log('=== プロジェクトのトリガー一覧 ===');
  console.log('');
  
  const triggers = ScriptApp.getProjectTriggers();
  
  if (triggers.length === 0) {
    console.log('トリガーは設定されていません');
    console.log('');
    console.log('【トリガー作成方法】');
    console.log('- 毎日実行: createDailyTrigger(3)  // 午前3時');
    console.log('- 毎週実行: createWeeklyTrigger(ScriptApp.WeekDay.MONDAY, 3)');
    console.log('- トークン更新: createTokenRefreshTrigger(2)  // 午前2時');
    return;
  }
  
  triggers.forEach((trigger, index) => {
    console.log(`【トリガー ${index + 1}】`);
    console.log('トリガーID:', trigger.getUniqueId());
    console.log('実行関数:', trigger.getHandlerFunction());
    console.log('トリガー種別:', trigger.getEventType());
    
    // 時間ベースのトリガーの場合、詳細情報を表示
    if (trigger.getEventType() === ScriptApp.EventType.CLOCK) {
      console.log('実行頻度:', getTriggerFrequency(trigger));
    }
    
    console.log('');
  });
  
  console.log('【トリガー削除】');
  console.log('全てのトリガーを削除: deleteAllTriggers()');
  console.log('特定のトリガーを削除: deleteTriggerById("トリガーID")');
}

/**
 * トリガーの実行頻度を人間が読みやすい形式の文字列で取得します（表示用ユーティリティ）。
 * 
 * @param {GoogleAppsScript.Script.Trigger} trigger - 頻度を調べたいトリガーオブジェクト
 * @return {string} 実行頻度の説明文字列
 * @private
 */
function getTriggerFrequency(trigger) {
  try {
    // 時間ベースのトリガーの場合
    const triggerSource = trigger.getTriggerSource();
    
    if (triggerSource === ScriptApp.TriggerSource.CLOCK) {
      // 詳細情報は取得できないため、シンプルな説明
      return '時間ベース(詳細はGUIで確認)';
    }
    
    return '不明';
  } catch (error) {
    return '取得できません';
  }
}

/**
 * プロジェクトに設定されている全てのトリガーを一括で削除します。
 * 
 * 開発環境の再構築時や、自動実行を完全に停止したい場合に使用します。
 * 削除後は復旧できないため、慎重に実行してください。
 */
function deleteAllTriggers() {
  console.log('=== 全トリガー削除 ===');
  
  const triggers = ScriptApp.getProjectTriggers();
  
  if (triggers.length === 0) {
    console.log('削除するトリガーがありません');
    return;
  }
  
  console.log('削除するトリガー数:', triggers.length);
  console.log('');
  
  triggers.forEach((trigger, index) => {
    console.log(`[${index + 1}] ${trigger.getHandlerFunction()} (ID: ${trigger.getUniqueId()})`);
    ScriptApp.deleteTrigger(trigger);
  });
  
  console.log('');
  console.log('✅ 全てのトリガーを削除しました');
}

/**
 * 指定された ID を持つトリガーを個別に削除します。
 * 
 * 特定のスケジュール設定のみを解除したい場合に使用します。
 * 
 * @param {string} triggerId - 削除対象のトリガー固有 ID
 */
function deleteTriggerById(triggerId) {
  const triggers = ScriptApp.getProjectTriggers();
  
  for (const trigger of triggers) {
    if (trigger.getUniqueId() === triggerId) {
      ScriptApp.deleteTrigger(trigger);
      console.log('✅ トリガーを削除しました');
      console.log('トリガーID:', triggerId);
      console.log('関数:', trigger.getHandlerFunction());
      return;
    }
  }
  
  console.log('❌ 指定されたIDのトリガーが見つかりません:', triggerId);
}

/**
 * 推奨されるトリガー設定（トークン更新 2時 / メイン処理 3時）を一括で作成します。
 * 
 * 環境構築後のセットアップを簡略化するために使用します。
 * @throws {Error} トリガー作成に失敗した場合
 */
function setupRecommendedTriggers() {
  console.log('=== 推奨トリガー設定 ===');
  console.log('');
  console.log('以下のトリガーを作成します:');
  console.log('1. セット商品マスタ更新: 毎日午前3時');
  console.log('2. トークン更新: 毎日午前2時');
  console.log('');
  
  try {
    // トークン更新トリガー(午前2時)
    createTokenRefreshTrigger(2);
    console.log('');
    
    // メイン処理トリガー(午前3時)
    createDailyTrigger(3);
    console.log('');
    
    console.log('='.repeat(70));
    console.log('✅ 推奨トリガー設定完了');
    console.log('='.repeat(70));
    console.log('');
    console.log('【確認】');
    console.log('トリガー一覧: listAllTriggers()');
    console.log('');
    console.log('【運用開始】');
    console.log('これで自動実行の準備が完了しました!');
    console.log('- 毎日午前2時: トークン更新');
    console.log('- 毎日午前3時: セット商品マスタ更新');
    
  } catch (error) {
    console.error('❌ トリガー設定エラー:', error.message);
    throw error;
  }
}