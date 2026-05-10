/**
 * @file 02_店舗マスタ連携.gs
 * @description ネクストエンジン受注明細取得スクリプト - Phase 2: 店舗マスタ連携。
 * 店舗マスタスプレッドシートから店舗情報を読み込み、店舗コードから店舗名を取得する機能を提供します。
 * 
 * ### 実装内容
 * 1. 店舗マスタスプレッドシートの読み込み
 * 2. 店舗コード→店舗名の変換マップ(Dictionary)作成
 * 3. 店舗名取得関数
 * 
 * ### 店舗マスタのフォーマット
 * - A列: 店舗ID(店舗コード)
 * - B列: 店舗名
 * - ※1行目はヘッダー行として扱います。
 * 
 * ### 依存関係 (01_基盤構築.gs)
 * - getScriptConfig() : 設定値の取得
 * - logMessage() : ログ出力制御
 * - LOG_LEVEL : ログレベル定数
 * 
 * ### 推奨テスト実行順序
 * 1. testShopMasterConnection() -> testShopMasterDataLoad() -> testShopNameMap()
 * 2. testGetShopName() -> testShopMapCache()
 * 3. testPhase2() (統合テスト)
 *
 * ### 注意事項
 * - キャッシュ (SHOP_MAP_CACHE) は同一実行内（1回のスクリプト起動）でのみ有効です。
 * - 実行アカウントには店舗マスタSSへの閲覧権限が必要です。
 * 
 * @version 1.0
 * @date 2025-11-24
 * @see testPhase2 - 統合テスト
 */

// =============================================================================
// 店舗マスタ読み込み
// =============================================================================

/**
 * 店舗マスタスプレッドシートを開く
 * 
 * @details
 * `getScriptConfig`経由で取得した `shopMasterSpreadsheetId` および `shopMasterSheetName` を使用して、
 * 指定されたスプレッドシートの特定のシートオブジェクトを取得します。
 * 
 * 外部スプレッドシート（店舗マスタ）との接続の起点となる関数です。
 * 権限不足やIDの間違いがある場合は、詳細なエラーメッセージとともに例外をスローします。
 * 
 * @return {GoogleAppsScript.Spreadsheet.Sheet} 対象の店舗マスタシートオブジェクト
 * @throws {Error} スプレッドシートIDが見つからない、またはシート名が存在しない場合
 * @see getScriptConfig
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
 * @details
 * `openShopMasterSheet` を呼び出してシートを取得し、全データ範囲を2次元配列として取得します。
 * 実データのみを抽出するため、1行目のヘッダー行は読み込み範囲から除外します。
 * データが1行（ヘッダーのみ）以下の場合、警告ログを出力し空の配列を返却します。
 * 
 * @return {Array<Array<string|number>>} 店舗マスタのデータ行（2次元配列）。ヘッダー行は含みません。
 * @see openShopMasterSheet
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
 * @details
 * `loadShopMasterData` で取得した2次元配列を走査し、JavaScriptの Map オブジェクトを生成します。
 * - キー: A列（店舗ID/店舗コード）を文字列化・トリムしたもの
 * - 値: B列（店舗名）を文字列化・トリムしたもの
 * 
 * 店舗コードが空の行は、データ不備としてスキップし、ログに警告を表示します。
 * このマップを使用することで、大量の受注データに対して高速に店舗名変換を行うことが可能になります。
 * 
 * @return {Map<string, string>} 店舗コードをキー、店舗名を値とする Map オブジェクト
 * @see loadShopMasterData
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
 * @details
 * 引数で渡された `shopMap`（変換テーブル）を用いて、特定の店舗コードに対応する店舗名を検索します。
 * 
 * マップに該当するコードが存在しない場合、後続の集計処理でエラーにならないよう、
 * 「不明な店舗(コード: X)」という代替文字列を返却し、ログに記録します。
 * 
 * @param {Map<string, string>} shopMap - `createShopNameMap` で作成された変換マップ
 * @param {string|number} shopCode - 店舗コード
 * @return {string} 取得した店舗名、または不明時の代替文字列
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
 * @details
 * GASの実行性能を最適化するためのユーティリティです。
 * 同一のスクリプト実行内で複数回店舗マスタへのアクセスが発生する場合、
 * 2回目以降はスプレッドシートの再読み込みを行わず、メモリ上の `SHOP_MAP_CACHE` を再利用します。
 * 
 * 1万件以上の受注データを処理する場合など、頻繁なマスタ参照が発生するシーンで効果的です。
 * 
 * @return {Map<string, string>} 店舗名マップ（キャッシュ済み、または新規作成）
 * @note グローバル変数 `SHOP_MAP_CACHE` を使用します。
 * @see createShopNameMap
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
 * 
 * @details
 * グローバル変数 `SHOP_MAP_CACHE` を null にリセットします。
 * これにより、次に `getShopMapWithCache` が呼ばれた際に必ず最新のスプレッドシートから再読み込みが行われます。
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
 * @details
 * スプレッドシートIDとシート名が正しいか、および実行ユーザーに適切なアクセス権があるかを検証します。
 * 成功時には、シートのメタ情報（名前、最終行数、最終列数）と、
 * A列・B列が意図した項目（店舗ID/店舗名など）であるかを確認するためのヘッダー情報を出力します。
 * 
 * @throws {Error} 接続に失敗した場合
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
 * @details
 * `loadShopMasterData` を実行し、戻り値が正しい2次元配列であるかを確認します。
 * 最初の5件のデータを抽出し、店舗コードと店舗名が正しく取得できているか（列のズレがないか）を
 * 目視確認するためにログ出力します。
 * 
 * @return {Array<Array>} 取得された店舗マスタデータ
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
 * @details
 * `createShopNameMap` を実行し、Map オブジェクトが正しく構築されるかを検証します。
 * キーの重複や空データのスキップ状況、およびMapの要素数を確認します。
 * サンプルとして先頭10件の対応関係を出力します。
 * 
 * @return {Map<string, string>} 作成された店舗名マップ
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
 * @details
 * `getShopName` 関数の挙動を2つのパターンで検証します。
 * 1. 正常系: 実際にマスタに存在するコードを渡し、正しい名前が返るか
 * 2. 異常系: 存在しない（または空の）コードを渡し、定義済みの「不明な店舗」文字列が返るか
 * 
 * このテストにより、マスタにない店舗コードが含まれる受注データのハンドリングを保証します。
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
 * @details
 * `getShopMapWithCache` を連続して呼び出し、2回目以降の実行時間が短縮されるかを確認します。
 * 1回目：スプレッドシートへのIOが発生
 * 2回目：メモリ(キャッシュ)からの取得
 * ミリ秒単位での処理時間を比較することで、最適化の効果を可視化します。
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
 * @details
 * 「店舗マスタ連携」フェーズで実装した全機能の結合テストを一括実行します。
 * シート接続からデータ読み込み、マスタマップの構築、キャッシュの動作までをシーケンシャルに検証します。
 * 
 * Phase 3以降のAPI連携処理で店舗名変換を行うために、このテストがパスしていることが必須要件となります。
 * 
 * @throws {Error} いずれかのサブテストで失敗した場合
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