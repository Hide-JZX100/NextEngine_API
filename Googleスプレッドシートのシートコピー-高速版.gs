/**
=============================================================================
商品マスタ ハイブリッド更新スクリプト
=============================================================================

【概要】
・importrange関数の動作不安定を解決するため、GASで直接データをコピー
・追加メイン、低頻度更新、削除なしのデータ特性に最適化
・従来の590秒実行時間を大幅短縮（平日：数秒、完全更新：約150秒）

【スクリプトプロパティの設定方法】
1. GASエディタで「プロジェクトの設定」を開く（歯車のアイコン）
2. 「スクリプトプロパティ」セクションまでスクロール
3. 「スクリプトプロパティの追加」をクリックし、以下のキーと値を設定

   キー                     | 値
   -------------------------|------------------------------------
   SOURCE_SPREADSHEET_ID    | ---実際の値---
   SOURCE_SHEET_NAME        | ---実際のシート名---
   DEST_SPREADSHEET_ID      | ---実際の値---
   DEST_SHEET_NAME          | ---実際のシート名---


【使用方法】
1. Master_HybridUpdate を毎日1回実行するトリガーを設定
   - 新規データのみ差分更新（高速）
   - 週1回（月曜日）自動で完全更新（更新データも反映）

2. データ更新があった場合は Master_ForceFullUpdate を手動実行
   - 即座に全データを完全更新

【実行パターン】
・平日：新規データがあれば差分追加（5-20秒）
・月曜：完全更新で既存データの更新も反映（約150秒/約13,000行 BD列）
・手動：更新発生時のみ完全更新を実行

【注意事項】
・スプレッドシートIDは実際の環境に合わせて変更してください
・初回実行時は完全更新が実行されます
・PropertiesServiceを使用して更新履歴を管理しています

予備関数
・Master_SuperFast：一番シンプルで高速
・Master_Optimized：大容量データ処理時の安全策として用意した関数

使い分けの想定
データ量___推奨関数___理由
~15,000行___Master_HybridUpdate___一括処理で高速
15,000~30,000行___Master_Optimized___バッチ処理で安全
30,000行~___Master_Optimized + 分割実行___確実な処理

【作成日】2025年9月19日
【最終更新】2025年9月20日
=============================================================================
*/

/**
 * スクリプトプロパティからスプレッドシートとシートオブジェクトを取得するヘルパー関数
 * @returns {Object} 必要なスプレッドシートとシートのオブジェクト
 */
function getSheets() {
  const scriptProperties = PropertiesService.getScriptProperties();

  // スクリプトプロパティからIDとシート名を取得
  const sourceSsId = scriptProperties.getProperty('SOURCE_SPREADSHEET_ID');
  const sourceSheetName = scriptProperties.getProperty('SOURCE_SHEET_NAME');
  const destSsId = scriptProperties.getProperty('DEST_SPREADSHEET_ID');
  const destSheetName = scriptProperties.getProperty('DEST_SHEET_NAME');

  // プロパティが設定されているか確認
  if (!sourceSsId || !sourceSheetName || !destSsId || !destSheetName) {
    const errorMsg = 'スクリプトプロパティが設定されていません。プロジェクトの設定で必要なプロパティを追加してください。';
    Logger.log(errorMsg);
    throw new Error(errorMsg);
  }

  // スプレッドシートとシートオブジェクトを開く
  const ss_copyFrom = SpreadsheetApp.openById(sourceSsId);
  const sheet_copyFrom = ss_copyFrom.getSheetByName(sourceSheetName);
  const ss_copyTo = SpreadsheetApp.openById(destSsId);
  const sheet_copyTo = ss_copyTo.getSheetByName(destSheetName);

  return {
    ss_copyFrom: ss_copyFrom,
    sheet_copyFrom: sheet_copyFrom,
    ss_copyTo: ss_copyTo,
    sheet_copyTo: sheet_copyTo,
    sourceSsId: sourceSsId,      // API用に追加
    destSsId: destSsId,          // API用に追加
    destSheetName: destSheetName // API用に追加
  };
}

// ハイブリッド更新版（追加メイン + 低頻度更新に最適）
function Master_HybridUpdate() {
  try {
    const startTime = new Date();
    const sheets = getSheets();
    const sheet_copyFrom = sheets.sheet_copyFrom;
    const sheet_copyTo = sheets.sheet_copyTo;

    const lastRow_From = sheet_copyFrom.getLastRow();
    const lastColumn_From = sheet_copyFrom.getLastColumn();
    const lastRow_To = sheet_copyTo.getLastRow();

    Logger.log(`コピー元: ${lastRow_From}行, コピー先: ${lastRow_To}行`);

    const scriptProperties = PropertiesService.getScriptProperties();
    const lastFullUpdate = scriptProperties.getProperty('LAST_FULL_UPDATE');
    const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');

    const dayOfWeek = new Date().getDay();
    const shouldFullUpdate = !lastFullUpdate || lastFullUpdate !== today && dayOfWeek === 1;

    if (shouldFullUpdate) {
      Logger.log('完全更新を実行します');

      // データ取得
      const getDataStart = new Date();

      const allData = sheet_copyFrom.getRange(1, 1, lastRow_From, lastColumn_From).getValues();

      const getDataEnd = new Date();
      Logger.log(`[データ取得] 完了: ${(getDataEnd - getDataStart) / 1000}秒`);

      // シートクリア
      const clearStart = new Date();

      sheet_copyTo.clear();

      const clearEnd = new Date();
      Logger.log(`[シートクリア] 完了: ${(clearEnd - clearStart) / 1000}秒`);

      // データ書き込み
      const setValuesStart = new Date();

      sheet_copyTo.getRange(1, 1, lastRow_From, lastColumn_From).setValues(allData);

      const setValuesEnd = new Date();
      Logger.log(`[データ書き込み] 完了: ${(setValuesEnd - setValuesStart) / 1000}秒`);

      if (lastColumn_From >= 9) {
        const formatStart = new Date();

        sheet_copyTo.getRange(1, 9, lastRow_From, Math.min(2, lastColumn_From - 8))
          .setNumberFormat("@");

        const formatEnd = new Date();
        Logger.log(`[書式設定] 完了: ${(formatEnd - formatStart) / 1000}秒`);
      }

      scriptProperties.setProperty('LAST_FULL_UPDATE', today);
      Logger.log('完全更新完了');

    } else if (lastRow_From > lastRow_To) {
      const newRows = lastRow_From - lastRow_To;
      Logger.log(`差分更新: ${newRows}行の新しいデータを追加`);

      // データ取得
      const getDataStart = new Date();

      const newData = sheet_copyFrom.getRange(lastRow_To + 1, 1, newRows, lastColumn_From).getValues();

      const getDataEnd = new Date();
      Logger.log(`[差分データ取得] 完了: ${(getDataEnd - getDataStart) / 1000}秒`);

      // データ書き込み
      const setValuesStart = new Date();

      sheet_copyTo.getRange(lastRow_To + 1, 1, newRows, lastColumn_From).setValues(newData);

      const setValuesEnd = new Date();
      Logger.log(`[差分データ書き込み] 完了: ${(setValuesEnd - setValuesStart) / 1000}秒`);

      if (lastColumn_From >= 9) {
        const formatStart = new Date();

        sheet_copyTo.getRange(lastRow_To + 1, 9, newRows, Math.min(2, lastColumn_From - 8))
          .setNumberFormat("@");

        const formatEnd = new Date();
        Logger.log(`[差分書式設定] 完了: ${(formatEnd - formatStart) / 1000}秒`);
      }

      Logger.log(`${newRows}行の新しいデータを追加しました`);

    } else {
      Logger.log('更新するデータがありません');
    }

    const endTime = new Date();
    const executionTime = (endTime - startTime) / 1000;
    Logger.log(`ハイブリッド更新完了: 実行時間 ${executionTime}秒`);

  } catch (error) {
    Logger.log(`エラー発生: ${error.toString()}`);
    throw error;
  }
}

/**
=============================================================================
商品マスタ ハイブリッド更新スクリプト (API高速版・計測ログ付)
=============================================================================

【概要】
・Google Sheets API (Advanced Services) を使用し、大規模データの更新を高速化。
・追加した「工程別ログ計測」を搭載し、APIの恩恵を数値で確認可能。
・月曜日の完全更新時はAPI、平日の差分更新は標準命令を使い分けるハイブリッド仕様。

【更新履歴】
2026/02/09：Sheets API v4 導入 ＋ 計測ログ実装
=============================================================================
*/

function Master_HybridUpdate_API() {
  try {
    const startTime = new Date();
    const info = getSheets();
    const sheet_from = info.sheet_copyFrom;
    const sheet_to = info.sheet_copyTo;

    const lastRow_From = sheet_from.getLastRow();
    const lastColumn_From = sheet_from.getLastColumn();
    const lastRow_To = sheet_to.getLastRow();

    Logger.log(`[開始] コピー元: ${lastRow_From}行 / コピー先: ${lastRow_To}行`);

    const scriptProperties = PropertiesService.getScriptProperties();
    const lastFullUpdate = scriptProperties.getProperty('LAST_FULL_UPDATE');
    const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
    const dayOfWeek = new Date().getDay();

    const shouldFullUpdate = !lastFullUpdate || (lastFullUpdate !== today && dayOfWeek === 1);

    if (shouldFullUpdate) {
      Logger.log('【完全更新】APIモードで実行します');

      // 1. データ取得（ここは標準命令でも十分高速です）
      const getDataStart = new Date();
      const allData = sheet_from.getRange(1, 1, lastRow_From, lastColumn_From).getValues();
      Logger.log(`[1.データ取得] 完了: ${(new Date() - getDataStart) / 1000}秒`);

      // 2. APIによるシートクリア（超高速）
      const clearStart = new Date();
      // A1形式の指定が必要なため補助関数を使用
      const rangeLetter = columnToLetter(lastColumn_From);
      const clearRange = `'${info.destSheetName}'!A1:${rangeLetter}${lastRow_To + 1000}`;

      Sheets.Spreadsheets.Values.clear({}, info.destSsId, clearRange);
      Logger.log(`[2.APIクリア] 完了: ${(new Date() - clearStart) / 1000}秒`);

      // 3. APIによるデータ書き込み（ここが最大の短縮ポイント）
      const setValuesStart = new Date();
      const valueRange = Sheets.newRowData();
      valueRange.values = allData;

      // USER_ENTERED を使うことで、書式設定の一部（日付や数値認識）をAPIに任せます
      Sheets.Spreadsheets.Values.update(valueRange, info.destSsId, `'${info.destSheetName}'!A1`, {
        valueInputOption: 'USER_ENTERED'
      });
      Logger.log(`[3.API書き込み] 完了: ${(new Date() - setValuesStart) / 1000}秒`);

      // 4. 書式設定（APIでも可能ですが、まずは従来の方式で時間を計測）
      if (lastColumn_From >= 9) {
        const formatStart = new Date();
        sheet_to.getRange(1, 9, lastRow_From, Math.min(2, lastColumn_From - 8))
          .setNumberFormat("@");
        Logger.log(`[4.書式設定] 完了: ${(new Date() - formatStart) / 1000}秒`);
      }

      scriptProperties.setProperty('LAST_FULL_UPDATE', today);
      Logger.log('完全更新完了');

    } else if (lastRow_From > lastRow_To) {
      // --- 差分更新セクション（ヒデノリさんの計測ログを継承） ---
      const newRows = lastRow_From - lastRow_To;
      Logger.log(`【差分更新】${newRows}行の追加`);

      const getDataStart = new Date();
      const newData = sheet_from.getRange(lastRow_To + 1, 1, newRows, lastColumn_From).getValues();
      Logger.log(`[差分取得] 完了: ${(new Date() - getDataStart) / 1000}秒`);

      const setValuesStart = new Date();
      sheet_to.getRange(lastRow_To + 1, 1, newRows, lastColumn_From).setValues(newData);
      Logger.log(`[差分書込] 完了: ${(new Date() - setValuesStart) / 1000}秒`);

      if (lastColumn_From >= 9) {
        const formatStart = new Date();
        sheet_to.getRange(lastRow_To + 1, 9, newRows, Math.min(2, lastColumn_From - 8))
          .setNumberFormat("@");
        Logger.log(`[差分書式] 完了: ${(new Date() - formatStart) / 1000}秒`);
      }
    } else {
      Logger.log('更新不要：データに変化はありません');
    }

    const executionTime = (new Date() - startTime) / 1000;
    Logger.log(`[全工程完了] 総合実行時間: ${executionTime}秒`);

  } catch (error) {
    Logger.log(`[エラー発生] ${error.toString()}`);
    throw error;
  }
}

/**
 * 補助関数：列番号をアルファベットに変換
 * (N88-BASICのサブルーチンのように呼び出して使います)
 */
function columnToLetter(column) {
  let temp, letter = '';
  while (column > 0) {
    temp = (column - 1) % 26;
    letter = String.fromCharCode(temp + 65) + letter;
    column = (column - temp - 1) / 26;
  }
  return letter;
}

// 手動で完全更新を実行する関数
function Master_ForceFullUpdate() {
  try {
    const startTime = new Date();
    const sheets = getSheets();
    const sheet_copyFrom = sheets.sheet_copyFrom;
    const sheet_copyTo = sheets.sheet_copyTo;

    const lastRow_From = sheet_copyFrom.getLastRow();
    const lastColumn_From = sheet_copyFrom.getLastColumn();

    Logger.log(`強制完全更新: ${lastRow_From}行 × ${lastColumn_From}列`);

    const allData = sheet_copyFrom.getRange(1, 1, lastRow_From, lastColumn_From).getValues();

    sheet_copyTo.clear();

    sheet_copyTo.getRange(1, 1, lastRow_From, lastColumn_From).setValues(allData);

    if (lastColumn_From >= 9) {
      sheet_copyTo.getRange(1, 9, lastRow_From, Math.min(2, lastColumn_From - 8))
        .setNumberFormat("@");
    }

    const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
    PropertiesService.getScriptProperties().setProperty('LAST_FULL_UPDATE', today);

    const endTime = new Date();
    const executionTime = (endTime - startTime) / 1000;
    Logger.log(`強制完全更新完了: 実行時間 ${executionTime}秒`);

  } catch (error) {
    Logger.log(`エラー発生: ${error.toString()}`);
    throw error;
  }
}

// バッチ処理版（元のコード）
function Master_Optimized() {
  try {
    const startTime = new Date();
    const sheets = getSheets();
    const sheet_copyFrom = sheets.sheet_copyFrom;
    const sheet_copyTo = sheets.sheet_copyTo;

    const lastRow_From = sheet_copyFrom.getLastRow();
    const lastColumn_From = sheet_copyFrom.getLastColumn();

    Logger.log(`コピー元データ範囲: ${lastRow_From}行 × ${lastColumn_From}列`);

    if (lastRow_From === 0 || lastColumn_From === 0) {
      Logger.log('コピー元にデータが存在しません');
      return;
    }

    sheet_copyTo.clear();

    const batchSize = 5000;

    for (let startRow = 1; startRow <= lastRow_From; startRow += batchSize) {
      const endRow = Math.min(startRow + batchSize - 1, lastRow_From);
      const numRows = endRow - startRow + 1;

      Logger.log(`処理中: ${startRow}行目から${endRow}行目まで（${numRows}行）`);

      const copyValue = sheet_copyFrom.getRange(startRow, 1, numRows, lastColumn_From).getValues();

      if (lastColumn_From >= 9) {
        sheet_copyTo.getRange(startRow, 9, numRows, Math.min(2, lastColumn_From - 8))
          .setNumberFormat("@");
      }

      sheet_copyTo.getRange(startRow, 1, numRows, lastColumn_From).setValues(copyValue);

      Utilities.sleep(100);
    }

    const endTime = new Date();
    const executionTime = (endTime - startTime) / 1000;
    Logger.log(`処理完了: 実行時間 ${executionTime}秒`);

  } catch (error) {
    Logger.log(`エラー発生: ${error.toString()}`);
    throw error;
  }
}

// より高速な代替案：値のみコピー（元のコード）
function Master_SuperFast() {
  try {
    const startTime = new Date();
    const sheets = getSheets();
    const sheet_copyFrom = sheets.sheet_copyFrom;
    const sheet_copyTo = sheets.sheet_copyTo;

    const lastRow_From = sheet_copyFrom.getLastRow();
    const lastColumn_From = sheet_copyFrom.getLastColumn();

    Logger.log(`データ範囲: ${lastRow_From}行 × ${lastColumn_From}列`);

    if (lastRow_From === 0 || lastColumn_From === 0) {
      Logger.log('データが存在しません');
      return;
    }

    const allData = sheet_copyFrom.getRange(1, 1, lastRow_From, lastColumn_From).getValues();

    sheet_copyTo.clear();

    sheet_copyTo.getRange(1, 1, lastRow_From, lastColumn_From).setValues(allData);

    if (lastColumn_From >= 9) {
      sheet_copyTo.getRange(1, 9, lastRow_From, Math.min(2, lastColumn_From - 8))
        .setNumberFormat("@");
    }

    const endTime = new Date();
    const executionTime = (endTime - startTime) / 1000;
    Logger.log(`Super Fast完了: 実行時間 ${executionTime}秒`);

  } catch (error) {
    Logger.log(`エラー発生: ${error.toString()}`);
    throw error;
  }
}