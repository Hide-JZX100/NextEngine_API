/**
 * =============================================================================
 * ネクストエンジン受注明細取得スクリプト - Phase 2: 店舗マスタ連携
 * =============================================================================
 * 
 * 【概要】
 * 店舗マスタスプレッドシートから店舗情報を読み込み、
 * 店舗コードから店舗名を取得する機能を提供します。
 * 
 * 【Phase 2 実装内容】
 * 1. 店舗マスタスプレッドシートの読み込み
 * 2. 店舗コード→店舗名の変換マップ(Dictionary)作成
 * 3. 店舗名取得関数
 * 4. テスト関数
 * 
 * 【店舗マスタの想定フォーマット】
 * A列: 店舗ID(店舗コード)
 * B列: 店舗名
 * ※1行目はヘッダー行として扱います
 * 
 * @version 1.0
 * @date 2025-11-24
 */

// =============================================================================
// 店舗マスタ読み込み
// =============================================================================

/**
 * 店舗マスタスプレッドシートを開く
 * 
 * スクリプトプロパティから店舗マスタのスプレッドシートIDとシート名を取得し、
 * 対象のシートオブジェクトを返します。
 * 
 * @return {GoogleAppsScript.Spreadsheet.Sheet} 店舗マスタシート
 * @throws {Error} スプレッドシートまたはシートが見つからない場合
 */
function openShopMasterSheet() {
  const config = getScriptConfig();
  
  try {
    // スプレッドシートを開く
    const spreadsheet = SpreadsheetApp.openById(config.shopMasterSpreadsheetId);
    
    // シートを取得
    const sheet = spreadsheet.getSheetByName(config.shopMasterSheetName);
    
    if (!sheet) {
      throw new Error(
        `シート "${config.shopMasterSheetName}" が見つかりません。\n` +
        `スプレッドシートID: ${config.shopMasterSpreadsheetId}`
      );
    }
    
    logMessage(`店舗マスタシートを開きました: ${spreadsheet.getName()} / ${sheet.getName()}`);
    
    return sheet;
    
  } catch (error) {
    throw new Error(
      `店舗マスタスプレッドシートを開けませんでした: ${error.message}\n` +
      `SHOP_MASTER_SPREADSHEET_ID と SHOP_MASTER_SHEET_NAME を確認してください。`
    );
  }
}

/**
 * 店舗マスタデータを読み込む
 * 
 * 店舗マスタシートからデータを全件読み込み、配列で返します。
 * 1行目(ヘッダー)は除外されます。
 * 
 * @return {Array<Array>} 店舗マスタデータ(2次元配列)
 */
function loadShopMasterData() {
  const sheet = openShopMasterSheet();
  
  // データ範囲を取得(1行目のヘッダーを除く)
  const lastRow = sheet.getLastRow();
  const lastColumn = sheet.getLastColumn();
  
  if (lastRow <= 1) {
    logMessage('⚠️ 店舗マスタにデータが存在しません(ヘッダー行のみ)');
    return [];
  }
  
  // 2行目からデータを取得
  const dataRange = sheet.getRange(2, 1, lastRow - 1, lastColumn);
  const data = dataRange.getValues();
  
  logMessage(`店舗マスタデータを読み込みました: ${data.length}件`);
  
  return data;
}

// =============================================================================
// 店舗コード→店舗名 変換マップ
// =============================================================================

/**
 * 店舗マスタから店舗コード→店舗名の変換マップを作成
 * 
 * A列(店舗ID)をキー、B列(店舗名)を値とするMapオブジェクトを作成します。
 * 空白行やA列が空のデータはスキップされます。
 * 
 * @return {Map<string, string>} 店舗コード→店舗名のマップ
 */
function createShopNameMap() {
  const data = loadShopMasterData();
  const shopMap = new Map();
  
  let skippedCount = 0;
  
  data.forEach((row, index) => {
    const shopCode = String(row[0]).trim(); // A列: 店舗ID
    const shopName = String(row[1]).trim(); // B列: 店舗名
    
    // 店舗コードが空の場合はスキップ
    if (!shopCode) {
      skippedCount++;
      logMessage(`⚠️ 店舗コードが空です(${index + 2}行目): スキップ`, LOG_LEVEL.SAMPLE);
      return;
    }
    
    // マップに登録
    shopMap.set(shopCode, shopName);
  });
  
  logMessage(`店舗名マップを作成しました: ${shopMap.size}件`);
  
  if (skippedCount > 0) {
    logMessage(`⚠️ スキップした行: ${skippedCount}件`);
  }
  
  return shopMap;
}

/**
 * 店舗コードから店舗名を取得
 * 
 * 店舗マスタマップから店舗コードに対応する店舗名を取得します。
 * マップに存在しない場合は「不明な店舗(コード: X)」を返します。
 * 
 * @param {Map<string, string>} shopMap - 店舗名マップ
 * @param {string|number} shopCode - 店舗コード
 * @return {string} 店舗名
 */
function getShopName(shopMap, shopCode) {
  const code = String(shopCode).trim();
  
  if (shopMap.has(code)) {
    return shopMap.get(code);
  } else {
    // マップに存在しない場合
    logMessage(`⚠️ 店舗マスタに存在しない店舗コード: ${code}`, LOG_LEVEL.SAMPLE);
    return `不明な店舗(コード: ${code})`;
  }
}

// =============================================================================
// キャッシュ機能(オプション)
// =============================================================================

/**
 * 店舗マスタマップをキャッシュから取得、またはキャッシュに保存
 * 
 * 同一実行内で複数回店舗マスタを参照する場合、
 * 毎回スプレッドシートを読み込むのは非効率なため、
 * グローバル変数にキャッシュします。
 * 
 * @return {Map<string, string>} 店舗名マップ
 */
let SHOP_MAP_CACHE = null;

function getShopMapWithCache() {
  if (SHOP_MAP_CACHE === null) {
    logMessage('店舗マスタをスプレッドシートから読み込みます...');
    SHOP_MAP_CACHE = createShopNameMap();
  } else {
    logMessage('店舗マスタをキャッシュから取得しました');
  }
  
  return SHOP_MAP_CACHE;
}

/**
 * キャッシュをクリア
 * 
 * テスト時や店舗マスタが更新された場合に使用します。
 */
function clearShopMapCache() {
  SHOP_MAP_CACHE = null;
  logMessage('店舗マスタキャッシュをクリアしました');
}

// =============================================================================
// テスト関数
// =============================================================================

/**
 * 店舗マスタシート接続テスト
 * 
 * スプレッドシートとシートが正しく開けるかテストします。
 */
function testShopMasterConnection() {
  console.log('=== 店舗マスタシート接続テスト ===');
  
  try {
    const sheet = openShopMasterSheet();
    
    console.log('✅ シート接続成功!');
    console.log('');
    console.log('【シート情報】');
    console.log('- スプレッドシート名:', sheet.getParent().getName());
    console.log('- シート名:', sheet.getName());
    console.log('- 最終行:', sheet.getLastRow());
    console.log('- 最終列:', sheet.getLastColumn());
    console.log('');
    
    // ヘッダー行を表示
    const headerRange = sheet.getRange(1, 1, 1, sheet.getLastColumn());
    const headers = headerRange.getValues()[0];
    console.log('【ヘッダー行】');
    headers.forEach((header, index) => {
      console.log(`${String.fromCharCode(65 + index)}列: ${header}`);
    });
    console.log('');
    
    console.log('✅ 店舗マスタシート接続テスト完了!');
    
  } catch (error) {
    console.error('❌ 接続エラー:', error.message);
    throw error;
  }
}

/**
 * 店舗マスタデータ読み込みテスト
 * 
 * 店舗マスタデータを読み込み、内容を確認します。
 */
function testShopMasterDataLoad() {
  console.log('=== 店舗マスタデータ読み込みテスト ===');
  
  try {
    const data = loadShopMasterData();
    
    console.log(`✅ データ読み込み成功: ${data.length}件`);
    console.log('');
    
    if (data.length > 0) {
      console.log('【先頭5件のデータ】');
      const sampleData = data.slice(0, 5);
      
      sampleData.forEach((row, index) => {
        console.log(`[${index + 1}] 店舗コード: ${row[0]}, 店舗名: ${row[1]}`);
      });
      
      if (data.length > 5) {
        console.log(`... 残り ${data.length - 5} 件`);
      }
      console.log('');
    }
    
    console.log('✅ 店舗マスタデータ読み込みテスト完了!');
    
    return data;
    
  } catch (error) {
    console.error('❌ 読み込みエラー:', error.message);
    throw error;
  }
}

/**
 * 店舗名マップ作成テスト
 * 
 * 店舗コード→店舗名の変換マップを作成し、内容を確認します。
 */
function testShopNameMap() {
  console.log('=== 店舗名マップ作成テスト ===');
  
  try {
    const shopMap = createShopNameMap();
    
    console.log(`✅ マップ作成成功: ${shopMap.size}件`);
    console.log('');
    
    console.log('【マップ内容(先頭10件)】');
    let count = 0;
    for (const [code, name] of shopMap) {
      console.log(`店舗コード: ${code} → 店舗名: ${name}`);
      count++;
      if (count >= 10) break;
    }
    
    if (shopMap.size > 10) {
      console.log(`... 残り ${shopMap.size - 10} 件`);
    }
    console.log('');
    
    console.log('✅ 店舗名マップ作成テスト完了!');
    
    return shopMap;
    
  } catch (error) {
    console.error('❌ マップ作成エラー:', error.message);
    throw error;
  }
}

/**
 * 店舗名取得テスト
 * 
 * 実際のデータで店舗名取得機能をテストします。
 */
function testGetShopName() {
  console.log('=== 店舗名取得テスト ===');
  
  try {
    const shopMap = createShopNameMap();
    
    // テストケース1: 存在する店舗コード
    console.log('【テストケース1: 存在する店舗コード】');
    
    // マップから実際の店舗コードを取得
    const testCodes = Array.from(shopMap.keys()).slice(0, 3);
    
    testCodes.forEach(code => {
      const shopName = getShopName(shopMap, code);
      console.log(`店舗コード: ${code} → 店舗名: ${shopName}`);
    });
    console.log('');
    
    // テストケース2: 存在しない店舗コード
    console.log('【テストケース2: 存在しない店舗コード】');
    const invalidCode = '99999';
    const shopName = getShopName(shopMap, invalidCode);
    console.log(`店舗コード: ${invalidCode} → 店舗名: ${shopName}`);
    console.log('');
    
    console.log('✅ 店舗名取得テスト完了!');
    
  } catch (error) {
    console.error('❌ 取得テストエラー:', error.message);
    throw error;
  }
}

/**
 * キャッシュ機能テスト
 * 
 * キャッシュが正しく動作するかテストします。
 */
function testShopMapCache() {
  console.log('=== キャッシュ機能テスト ===');
  
  try {
    // キャッシュをクリア
    clearShopMapCache();
    console.log('1回目: キャッシュなし');
    
    const startTime1 = new Date();
    const map1 = getShopMapWithCache();
    const endTime1 = new Date();
    const time1 = endTime1 - startTime1;
    
    console.log(`処理時間: ${time1}ms, 件数: ${map1.size}`);
    console.log('');
    
    console.log('2回目: キャッシュあり');
    const startTime2 = new Date();
    const map2 = getShopMapWithCache();
    const endTime2 = new Date();
    const time2 = endTime2 - startTime2;
    
    console.log(`処理時間: ${time2}ms, 件数: ${map2.size}`);
    console.log('');
    
    console.log(`高速化: ${time1 - time2}ms短縮`);
    console.log('');
    
    console.log('✅ キャッシュ機能テスト完了!');
    
  } catch (error) {
    console.error('❌ キャッシュテストエラー:', error.message);
    throw error;
  }
}

// =============================================================================
// Phase 2 統合テスト
// =============================================================================

/**
 * Phase 2 統合テスト
 * 
 * Phase 2で実装した全機能をテストします。
 */
function testPhase2() {
  console.log('╔════════════════════════════════════════════════════════════╗');
  console.log('║  Phase 2: 店舗マスタ連携 - 統合テスト                    ║');
  console.log('╚════════════════════════════════════════════════════════════╝');
  console.log('');
  
  try {
    // 1. シート接続テスト
    console.log('【1】店舗マスタシート接続テスト');
    testShopMasterConnection();
    console.log('');
    
    // 2. データ読み込みテスト
    console.log('【2】店舗マスタデータ読み込みテスト');
    testShopMasterDataLoad();
    console.log('');
    
    // 3. マップ作成テスト
    console.log('【3】店舗名マップ作成テスト');
    testShopNameMap();
    console.log('');
    
    // 4. 店舗名取得テスト
    console.log('【4】店舗名取得テスト');
    testGetShopName();
    console.log('');
    
    // 5. キャッシュ機能テスト
    console.log('【5】キャッシュ機能テスト');
    testShopMapCache();
    console.log('');
    
    console.log('╔════════════════════════════════════════════════════════════╗');
    console.log('║  ✅ Phase 2 統合テスト: すべて成功!                       ║');
    console.log('╚════════════════════════════════════════════════════════════╝');
    console.log('');
    console.log('【次のステップ】');
    console.log('Phase 3: ネクストエンジンAPI接続の開発に進みます。');
    
  } catch (error) {
    console.error('');
    console.error('╔════════════════════════════════════════════════════════════╗');
    console.error('║  ❌ Phase 2 統合テスト: エラー発生                       ║');
    console.error('╚════════════════════════════════════════════════════════════╝');
    console.error('');
    console.error('エラー内容:', error.message);
    console.error('');
    console.error('【確認事項】');
    console.error('- SHOP_MASTER_SPREADSHEET_ID が正しいか');
    console.error('- SHOP_MASTER_SHEET_NAME が正しいか');
    console.error('- スプレッドシートへのアクセス権限があるか');
    console.error('- 店舗マスタシートにデータが存在するか');
    
    throw error;
  }
}