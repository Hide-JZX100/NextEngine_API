// ============================================
// Step 1.1: ログ機能（共通部分のみ）
// このファイルを「Step1_1_ログ機能.gs」として保存してください
// ============================================

// ============================================
// ログレベル管理
// ============================================

const LOG_LEVEL = {
  DEBUG: 3,
  INFO: 2,
  ERROR: 1,
  NONE: 0
};

/**
 * 現在のログレベルを取得
 * スクリプトプロパティから取得、未設定の場合はINFO
 */
function getCurrentLogLevel() {
  const scriptProperties = PropertiesService.getScriptProperties();
  const level = scriptProperties.getProperty('LOG_LEVEL');
  
  if (level === null || level === undefined) {
    return LOG_LEVEL.INFO; // デフォルトはINFO
  }
  
  return parseInt(level);
}

/**
 * ログレベルを設定
 * @param {number} level - ログレベル（LOG_LEVEL.DEBUG, LOG_LEVEL.INFO等）
 */
function setLogLevel(level) {
  const scriptProperties = PropertiesService.getScriptProperties();
  scriptProperties.setProperty('LOG_LEVEL', level.toString());
  Logger.log(`ログレベルを ${level} に設定しました`);
}

/**
 * 共通ログ関数
 * @param {string} level - ログレベル（'DEBUG', 'INFO', 'ERROR'）
 * @param {string} message - ログメッセージ
 */
function log(level, message) {
  const levels = {
    'DEBUG': LOG_LEVEL.DEBUG,
    'INFO': LOG_LEVEL.INFO,
    'ERROR': LOG_LEVEL.ERROR
  };
  
  const currentLevel = getCurrentLogLevel();
  
  if (levels[level] <= currentLevel) {
    const timestamp = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd HH:mm:ss');
    Logger.log(`[${timestamp}] [${level}] ${message}`);
  }
}

// ============================================
// 処理時間計測用クラス
// ============================================

class Timer {
  constructor(processName) {
    this.processName = processName;
    this.startTime = new Date();
    log('INFO', `=== ${processName} 開始 ===`);
  }
  
  end(additionalInfo = '') {
    const endTime = new Date();
    const elapsedTime = ((endTime - this.startTime) / 1000).toFixed(2);
    log('INFO', `=== ${this.processName} 完了 ===`);
    log('INFO', `処理時間: ${elapsedTime}秒 ${additionalInfo}`);
    return parseFloat(elapsedTime);
  }
  
  getElapsedTime() {
    const currentTime = new Date();
    return ((currentTime - this.startTime) / 1000).toFixed(2);
  }
}

// ============================================
// サンプルログ出力用関数
// ============================================

/**
 * 配列データのサンプルログ出力
 * @param {Array} dataArray - データ配列
 * @param {number} sampleSize - サンプル数（デフォルト3）
 * @param {string} label - ラベル
 */
function logSampleData(dataArray, sampleSize = 3, label = 'データ') {
  const totalCount = dataArray.length;
  
  if (totalCount === 0) {
    log('INFO', `${label}: 0件`);
    return;
  }
  
  log('DEBUG', `--- ${label} サンプル（全${totalCount}件中） ---`);
  
  // 最初のN件
  for (let i = 0; i < Math.min(sampleSize, totalCount); i++) {
    log('DEBUG', `[${i + 1}] ${JSON.stringify(dataArray[i])}`);
  }
  
  // 中間は省略メッセージ
  if (totalCount > sampleSize * 2) {
    log('DEBUG', `... (中間 ${totalCount - sampleSize * 2} 件省略) ...`);
  }
  
  // 最後のN件（データが十分にある場合のみ）
  if (totalCount > sampleSize) {
    for (let i = Math.max(sampleSize, totalCount - sampleSize); i < totalCount; i++) {
      log('DEBUG', `[${i + 1}] ${JSON.stringify(dataArray[i])}`);
    }
  }
}

// ============================================
// ログレベル設定用の補助関数
// ============================================

/**
 * ログレベルをDEBUGに設定（開発・デバッグ時）
 */
function setLogLevelDebug() {
  setLogLevel(LOG_LEVEL.DEBUG);
  log('INFO', 'ログレベルをDEBUGに変更しました（詳細ログモード）');
}

/**
 * ログレベルをINFOに設定（通常運用時）
 */
function setLogLevelInfo() {
  setLogLevel(LOG_LEVEL.INFO);
  log('INFO', 'ログレベルをINFOに変更しました（通常モード）');
}

/**
 * ログレベルをERRORに設定（エラーのみ）
 */
function setLogLevelError() {
  setLogLevel(LOG_LEVEL.ERROR);
  log('INFO', 'ログレベルをERRORに変更しました（エラーのみモード）');
}

/**
 * 現在のログレベルを確認
 */
function checkCurrentLogLevel() {
  const level = getCurrentLogLevel();
  const levelNames = {
    [LOG_LEVEL.DEBUG]: 'DEBUG（詳細ログモード）',
    [LOG_LEVEL.INFO]: 'INFO（通常モード）',
    [LOG_LEVEL.ERROR]: 'ERROR（エラーのみモード）',
    [LOG_LEVEL.NONE]: 'NONE（ログなし）'
  };
  Logger.log(`現在のログレベル: ${levelNames[level] || level}`);
  return level;
}

// ============================================
// 初期設定用関数（初回のみ実行）
// ============================================

/**
 * Step 1.1の初期設定
 * この関数を1度だけ実行してください
 */
function initStep1_1() {
  Logger.log('========================================');
  Logger.log('Step 1.1 初期設定を開始します');
  Logger.log('========================================');
  
  // ログレベルをINFOに設定
  setLogLevelInfo();
  
  // 設定確認
  checkCurrentLogLevel();
  
  Logger.log('========================================');
  Logger.log('初期設定完了');
  Logger.log('========================================');
  Logger.log('');
  Logger.log('次のステップ:');
  Logger.log('1. 既存のreceiveStockInfo関数を改善版に置き換えてください');
  Logger.log('2. testStockInfoOnly() を実行してテストしてください');
}