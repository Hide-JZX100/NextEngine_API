//日次更新///////////////////////////////////////////////////////////////////////////////// 
function dailyUpdate() {

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('order_summary_');
  
  var Lrow = sheet.getLastRow(); //最終行
  var row = 2;
  var col = 1;
  var rng = sheet.getRange(row, col, Lrow);
  var num = rng.setNumberFormat("#");

//売れ数集計START////////////////////////////////////////////////////////////////////////////

//JCOシートで日付判定
  //コピー元
  var ss_From = SpreadsheetApp.openById('1h0jeG2XOu8rRXlMcUYh2Gi67mx2HADMztpqusNLBc0E');
  var sheet_From = ss_From.getSheetByName('週間集計_JCO');

  //貼付先
  var ss_To = SpreadsheetApp.openById('1199zcuvnGL0BmaurTc1NNDLdEEgr5qzU146cjaaKyd0');
  var sheet_To = ss_To.getSheetByName('SKU別売れ数');

  //貼付先の最終行判定
  var last_row = sheet_To.getLastRow();
  Logger.log(last_row);

　for(var i = last_row; i >= 1; i--){
　　var rng = sheet_To.getRange(i, 1);
　　if(rng.getValue() != ''){
      var rowL = rng.getRowIndex();
      var a1 = rowL  + 1
      break;
　　}
　}
  

//日付判定
  var d1 = rng.getValue();
  var d2 = sheet_From.getRange('A2').getValue();
  if(d1 != d2){
    var d3 = d2 - d1;
    a1 = rowL + d3   //貼付開始行
  }else{
    a1 = rowL
  }

  var startR = 2;     //コピー開始行
  var startC = 1;     //コピー開始列
  var lastR = 7;      //コピー最終行
  Logger.log(a1);  
  
//↓↓順番にコピペ↓↓  
  //JCO
  var lastC = sheet_From.getLastColumn();   //コピー最終列
  var copyValue = sheet_From.getRange(startR, startC, lastR ,lastC).getValues();
  sheet_To.getRange(a1, startC, lastR, lastC).setValues(copyValue);    //該当日付行へ値コピー
  Logger.log('JCO end');
  
  //ADD20151016START/////////////////////////////////////////////////////////////////////////
  //価格帯別集計
  var sheet_From = ss_From.getSheetByName('価格帯別集計');
  var ss_To = SpreadsheetApp.openById('1fL4K9HZJ72Hpe4lNP47deGpyYJI1qyM1WPOYugjTf7I');
  var sheet_To = ss_To.getSheetByName('日別集計データ');  
  
  var lastC = sheet_From.getLastColumn();   //コピー最終列
  var copyValue = sheet_From.getRange(startR, startC, lastR ,lastC).getValues();
  sheet_To.getRange(a1, startC, lastR, lastC).setValues(copyValue);    //該当日付行へ値コピー
  Logger.log('価格帯別集計 end');

  //ZHI
  var sheet_From = ss_From.getSheetByName('週間集計_ZHI');
  var ss_To = SpreadsheetApp.openById('1ueClrIM0FBV7T4kuwHNkK8fknkIqbGEEcJUuSt-bJgw');
  var sheet_To = ss_To.getSheetByName('SKU別売れ数');  
  
  var lastC = sheet_From.getLastColumn();   //コピー最終列
  var copyValue = sheet_From.getRange(startR, startC, lastR ,lastC).getValues();
  sheet_To.getRange(a1, startC, lastR, lastC).setValues(copyValue);    //該当日付行へ値コピー
  Logger.log('ZHI end');

  //DHF
  sheet_From = ss_From.getSheetByName('週間集計_DHF');
  ss_To = SpreadsheetApp.openById('1_A2sQpaB5uFUUKkf-CB-Mxv7JUs5tWtigJqPhyeYKtI');
  sheet_To = ss_To.getSheetByName('SKU別売れ数');  
  
  lastC = sheet_From.getLastColumn();   //コピー最終列
  copyValue = sheet_From.getRange(startR, startC, lastR ,lastC).getValues();
  sheet_To.getRange(a1, startC, lastR, lastC).setValues(copyValue);    //該当日付行へ値コピー
  Logger.log('DHF end');

  //マットレスバンド
  sheet_From = ss_From.getSheetByName('週間集計_MB');
  ss_To = SpreadsheetApp.openById('1RFdJbE2y9jZUoWOXy2lJp0bmJ2PElcDRVxG_OIIpyEY');
  sheet_To = ss_To.getSheetByName('SKU別売れ数');  
  
  lastC = sheet_From.getLastColumn();   //コピー最終列
  copyValue = sheet_From.getRange(startR, startC, lastR ,lastC).getValues();
  sheet_To.getRange(a1, startC, lastR, lastC).setValues(copyValue);    //該当日付行へ値コピー
  Logger.log('MB end');

  //FSF
  sheet_From = ss_From.getSheetByName('週間集計_FSF');
  ss_To = SpreadsheetApp.openById('15Q6fWIwCbdRX1D9OA--2xeFqJi7MfpWru-_4ysL9kQQ');
  sheet_To = ss_To.getSheetByName('SKU別売れ数');  
  
  lastC = sheet_From.getLastColumn();   //コピー最終列
  copyValue = sheet_From.getRange(startR, startC, lastR ,lastC).getValues();
  sheet_To.getRange(a1, startC, lastR, lastC).setValues(copyValue);    //該当日付行へ値コピー
  Logger.log('FSF end');

  //JTM
  sheet_From = ss_From.getSheetByName('週間集計_JTM');
  ss_To = SpreadsheetApp.openById('1Zbomb_q79I50zD0OVAsL1CBPgfmbse8HZfkED6L2nJQ');
  sheet_To = ss_To.getSheetByName('SKU別売れ数');  
  
  lastC = sheet_From.getLastColumn();   //コピー最終列
  copyValue = sheet_From.getRange(startR, startC, lastR ,lastC).getValues();
  sheet_To.getRange(a1, startC, lastR, lastC).setValues(copyValue);    //該当日付行へ値コピー
  Logger.log('JTM end');

  //SLM
  sheet_From = ss_From.getSheetByName('週間集計_SLM');
  ss_To = SpreadsheetApp.openById('1xoS-OlVhDD5qqvlNvZMGUpF0fF4GgcB792BMJedCP6I');
  sheet_To = ss_To.getSheetByName('SKU別売れ数');  
  
  lastC = sheet_From.getLastColumn();   //コピー最終列
  copyValue = sheet_From.getRange(startR, startC, lastR ,lastC).getValues();
  sheet_To.getRange(a1, startC, lastR, lastC).setValues(copyValue);    //該当日付行へ値コピー
  Logger.log('SLM end');

  //CDF
  sheet_From = ss_From.getSheetByName('週間集計_CDF');
  ss_To = SpreadsheetApp.openById('1P2_E4sZosuN2B7rGbQWa81obl9MU1Q5ryJQC6HNaedw');
  sheet_To = ss_To.getSheetByName('SKU別売れ数');  
  
  lastC = sheet_From.getLastColumn();   //コピー最終列
  copyValue = sheet_From.getRange(startR, startC, lastR ,lastC).getValues();
  sheet_To.getRange(a1, startC, lastR, lastC).setValues(copyValue);    //該当日付行へ値コピー
  Logger.log('CDF end');

  //QAT
  sheet_From = ss_From.getSheetByName('週間集計_QAT');
  ss_To = SpreadsheetApp.openById('1EfEQD-tW2UO7exzbRDm0rSKfyDD_VEqJpCoCBZUyIXM');
  sheet_To = ss_To.getSheetByName('SKU別売れ数');  
  
  lastC = sheet_From.getLastColumn();   //コピー最終列
  copyValue = sheet_From.getRange(startR, startC, lastR ,lastC).getValues();
  sheet_To.getRange(a1, startC, lastR, lastC).setValues(copyValue);    //該当日付行へ値コピー
  Logger.log('QAT end');

  //ZNZ
  sheet_From = ss_From.getSheetByName('週間集計_ZNZ');
  ss_To = SpreadsheetApp.openById('1iISRrQ-_BEzYXxNA5HlxXhxreItOCpdgeF5JShRS3LA');
  sheet_To = ss_To.getSheetByName('SKU別売れ数');  
  
  lastC = sheet_From.getLastColumn();   //コピー最終列
  copyValue = sheet_From.getRange(startR, startC, lastR ,lastC).getValues();
  sheet_To.getRange(a1, startC, lastR, lastC).setValues(copyValue);    //該当日付行へ値コピー
  Logger.log('ZNZ end');

  //COT
  sheet_From = ss_From.getSheetByName('週間集計_COT');
  ss_To = SpreadsheetApp.openById('1gMwhCENoSUrhJjboGBp5I0Wj_O3AijjAqWxKLN_LVwg');
  sheet_To = ss_To.getSheetByName('SKU別売れ数');  
  
  lastC = sheet_From.getLastColumn();   //コピー最終列
  copyValue = sheet_From.getRange(startR, startC, lastR ,lastC).getValues();
  sheet_To.getRange(a1, startC, lastR, lastC).setValues(copyValue);    //該当日付行へ値コピー
  Logger.log('COT end');

  //Enjoy
  sheet_From = ss_From.getSheetByName('週間集計_Enjoy');
  ss_To = SpreadsheetApp.openById('1BpSNJbMGoVMSC8_ozKdjgW-3SOm578-LVxxzj3DBl2Q');
  sheet_To = ss_To.getSheetByName('SKU別売れ数');  
  
  lastC = sheet_From.getLastColumn();   //コピー最終列
  copyValue = sheet_From.getRange(startR, startC, lastR ,lastC).getValues();
  sheet_To.getRange(a1, startC, lastR, lastC).setValues(copyValue);    //該当日付行へ値コピー
  Logger.log('Enjoy end');

  //万鵬家具
  sheet_From = ss_From.getSheetByName('週間集計_万鵬家具');
  ss_To = SpreadsheetApp.openById('1LnaHXiGFfCn7U78hX1gxxKI-Lq_ZEL0kVYwQLIcFUgg');
  sheet_To = ss_To.getSheetByName('SKU別売れ数');  
  
  lastC = sheet_From.getLastColumn();   //コピー最終列
  copyValue = sheet_From.getRange(startR, startC, lastR ,lastC).getValues();
  sheet_To.getRange(a1, startC, lastR, lastC).setValues(copyValue);    //該当日付行へ値コピー
  Logger.log('万鵬 end');

  //WAM
  sheet_From = ss_From.getSheetByName('週間集計_WAM');
  ss_To = SpreadsheetApp.openById('1Alzq_TqvoeykLPScN0UgM-TlUkBVna9Q8OmQyTRrwhc');
  sheet_To = ss_To.getSheetByName('SKU別売れ数');  
  
  lastC = sheet_From.getLastColumn();   //コピー最終列
  copyValue = sheet_From.getRange(startR, startC, lastR ,lastC).getValues();
  sheet_To.getRange(a1, startC, lastR, lastC).setValues(copyValue);    //該当日付行へ値コピー
  Logger.log('WAM end');

  //DIA
  sheet_From = ss_From.getSheetByName('週間集計_DIA');
  ss_To = SpreadsheetApp.openById('1-Href6oqzu1VOontNtaIAJHJjtHGhrnWn1C_I-sR-3g');
  sheet_To = ss_To.getSheetByName('SKU別売れ数');  
  
  lastC = sheet_From.getLastColumn();   //コピー最終列
  copyValue = sheet_From.getRange(startR, startC, lastR ,lastC).getValues();
  sheet_To.getRange(a1, startC, lastR, lastC).setValues(copyValue);    //該当日付行へ値コピー
  Logger.log('DIA end');

  //MSI
  sheet_From = ss_From.getSheetByName('週間集計_MSI');
  ss_To = SpreadsheetApp.openById('1obPmlVq8WCV1fG_EJXSMPeDK6PwaYu9AhZ1MpCjBT58');
  sheet_To = ss_To.getSheetByName('SKU別売れ数');  
  
  lastC = sheet_From.getLastColumn();   //コピー最終列
  copyValue = sheet_From.getRange(startR, startC, lastR ,lastC).getValues();
  sheet_To.getRange(a1, startC, lastR, lastC).setValues(copyValue);    //該当日付行へ値コピー
  Logger.log('MSI end');

  //EON
  sheet_From = ss_From.getSheetByName('週間集計_EON');
  ss_To = SpreadsheetApp.openById('1c1OcweBAV-OCwMnBqrQszOXuwtLeJJ4c8lDy2W1Du8Q');
  sheet_To = ss_To.getSheetByName('SKU別売れ数');  
  
  lastC = sheet_From.getLastColumn();   //コピー最終列
  copyValue = sheet_From.getRange(startR, startC, lastR ,lastC).getValues();
  sheet_To.getRange(a1, startC, lastR, lastC).setValues(copyValue);    //該当日付行へ値コピー
  Logger.log('EON end');

  //STL
  sheet_From = ss_From.getSheetByName('週間集計_STL');
  ss_To = SpreadsheetApp.openById('1xM0zNiBs26tCs5_2DZP_rKJLIgIP8YgQeVoMZHikDE8');
  sheet_To = ss_To.getSheetByName('SKU別売れ数');  
  
  lastC = sheet_From.getLastColumn();   //コピー最終列
  copyValue = sheet_From.getRange(startR, startC, lastR ,lastC).getValues();
  sheet_To.getRange(a1, startC, lastR, lastC).setValues(copyValue);    //該当日付行へ値コピー
  Logger.log('STL end');

  //VIA
  sheet_From = ss_From.getSheetByName('週間集計_VIA');
  ss_To = SpreadsheetApp.openById('1sUGiTXymLCZoetpsr1_biBGlf5UuDNuMDXzaYyL_OJM');
  sheet_To = ss_To.getSheetByName('SKU別売れ数');  
  
  lastC = sheet_From.getLastColumn();   //コピー最終列
  copyValue = sheet_From.getRange(startR, startC, lastR ,lastC).getValues();
  sheet_To.getRange(a1, startC, lastR, lastC).setValues(copyValue);    //該当日付行へ値コピー
  Logger.log('VIA end');

  //CAN
  sheet_From = ss_From.getSheetByName('週間集計_CAN');
  ss_To = SpreadsheetApp.openById('1QBRvzwYa2ngbKvAnyF6jf6c__eAR94jHGaIJqYODw3M');
  sheet_To = ss_To.getSheetByName('SKU別売れ数');  
  
  lastC = sheet_From.getLastColumn();   //コピー最終列
  copyValue = sheet_From.getRange(startR, startC, lastR ,lastC).getValues();
  sheet_To.getRange(a1, startC, lastR, lastC).setValues(copyValue);    //該当日付行へ値コピー
  Logger.log('CAN end');

  //PIT
  sheet_From = ss_From.getSheetByName('週間集計_PIT');
  ss_To = SpreadsheetApp.openById('1y1xucmACOUaZpOaITphy4WQAza9Aw3FCjk4TkuQxX7M');
  sheet_To = ss_To.getSheetByName('SKU別売れ数');  
  
  lastC = sheet_From.getLastColumn();   //コピー最終列
  copyValue = sheet_From.getRange(startR, startC, lastR ,lastC).getValues();
  sheet_To.getRange(a1, startC, lastR, lastC).setValues(copyValue);    //該当日付行へ値コピー
  Logger.log('PIT end');

  //GFT
  sheet_From = ss_From.getSheetByName('週間集計_GFT');
  ss_To = SpreadsheetApp.openById('13N3JXvoLERoOTlUz_p7rus1gjt39-H9s_WnZ0qsOCn0');
  sheet_To = ss_To.getSheetByName('SKU別売れ数');  
  
  lastC = sheet_From.getLastColumn();   //コピー最終列
  copyValue = sheet_From.getRange(startR, startC, lastR ,lastC).getValues();
  sheet_To.getRange(a1, startC, lastR, lastC).setValues(copyValue);    //該当日付行へ値コピー
  Logger.log('GFT end');

  //ADD20151104START/////////////////////////////////////////////////////////////////////////
  //サービス分集計
  sheet_From = ss_From.getSheetByName('サービス分集計');
  ss_To = SpreadsheetApp.openById('1LqgzV9egbflf1fxgzhUMUG2zOHCg-Sry5prxNM4e16A');
  sheet_To = ss_To.getSheetByName('SKU別売れ数');  
  
  lastC = sheet_From.getLastColumn();   //コピー最終列
  copyValue = sheet_From.getRange(startR, startC, lastR ,lastC).getValues();
  sheet_To.getRange(a1, startC, lastR, lastC).setValues(copyValue);    //該当日付行へ値コピー
  Logger.log('サービス分集計 end');

  //フランスベッド
  sheet_From = ss_From.getSheetByName('週間集計_FB_MT');
  ss_To = SpreadsheetApp.openById('1Gz0PEHr9Luek_EfqMkjO_itrCOzKkqfsucRreI-i-CA');
  sheet_To = ss_To.getSheetByName('SKU別売れ数');  
  
  lastC = sheet_From.getLastColumn();   //コピー最終列
  copyValue = sheet_From.getRange(startR, startC, lastR ,lastC).getValues();
  sheet_To.getRange(a1, startC, lastR, lastC).setValues(copyValue);    //該当日付行へ値コピー
  Logger.log('フランスベッド end');

  //国内マットレス
  sheet_From = ss_From.getSheetByName('週間集計_国内マットレス');
  ss_To = SpreadsheetApp.openById('1fMrHuf_LQ4N70JLyMkoDylKc0vQ0Xmpi05pRXiRhbDo');
  sheet_To = ss_To.getSheetByName('SKU別売れ数');  
  
  lastC = sheet_From.getLastColumn();   //コピー最終列
  copyValue = sheet_From.getRange(startR, startC, lastR ,lastC).getValues();
  sheet_To.getRange(a1, startC, lastR, lastC).setValues(copyValue);    //該当日付行へ値コピー
  Logger.log('国内マットレス end');

  //国内寝装品
  sheet_From = ss_From.getSheetByName('週間集計_国内寝装品');
  ss_To = SpreadsheetApp.openById('1L4ZuNvCYFedvTqMxYzdhGh5W4nOHMGtwozs5q2LPAII');
  sheet_To = ss_To.getSheetByName('SKU別売れ数');  
  
  lastC = sheet_From.getLastColumn();   //コピー最終列
  copyValue = sheet_From.getRange(startR, startC, lastR ,lastC).getValues();
  sheet_To.getRange(a1, startC, lastR, lastC).setValues(copyValue);    //該当日付行へ値コピー
  Logger.log('国内寝装品 end');

  //熊井綿業
  sheet_From = ss_From.getSheetByName('週間集計_熊井綿業');
  ss_To = SpreadsheetApp.openById('1gAti4hx74-Ap9y4W1jfG9DQO1_QH5NlzLK8pQkwCg9c');
  sheet_To = ss_To.getSheetByName('SKU別売れ数');  
  
  lastC = sheet_From.getLastColumn();   //コピー最終列
  copyValue = sheet_From.getRange(startR, startC, lastR ,lastC).getValues();
  sheet_To.getRange(a1, startC, lastR, lastC).setValues(copyValue);    //該当日付行へ値コピー
  Logger.log('熊井綿業 end');

  //水美家具　2021/03/03　集計終了
  //sheet_From = ss_From.getSheetByName('週間集計_水美家具');
  //ss_To = SpreadsheetApp.openById('1gmKncD9TaF4GxK0PutviQN5Le0_TdOcV18ZWObTy3dA');
  //sheet_To = ss_To.getSheetByName('SKU別売れ数');  
  //
  //lastC = sheet_From.getLastColumn();   //コピー最終列
  //copyValue = sheet_From.getRange(startR, startC, lastR ,lastC).getValues();
  //sheet_To.getRange(a1, startC, lastR, lastC).setValues(copyValue);    //該当日付行へ値コピー

  //麦昆 20190313集計終了
  //sheet_From = ss_From.getSheetByName('週間集計_麦昆');
  //ss_To = SpreadsheetApp.openById('1wVo6vICu6mC5KWzg8QSnOpoacPFR0eKEEL5_XesBS0s');
  //sheet_To = ss_To.getSheetByName('SKU別売れ数');  
  //
  //lastC = sheet_From.getLastColumn();   //コピー最終列
  //copyValue = sheet_From.getRange(startR, startC, lastR ,lastC).getValues();
  //sheet_To.getRange(a1, startC, lastR, lastC).setValues(copyValue);    //該当日付行へ値コピー

  //Lis　2021/03/03　集計終了
  //sheet_From = ss_From.getSheetByName('週間集計_Lis');
  //ss_To = SpreadsheetApp.openById('1diqrL8SR6vWon2ZOAlWnkipW1hWoaRVOGgsNI8XcQvs');
  //sheet_To = ss_To.getSheetByName('SKU別売れ数');  
  //
  //lastC = sheet_From.getLastColumn();   //コピー最終列
  //copyValue = sheet_From.getRange(startR, startC, lastR ,lastC).getValues();
  //sheet_To.getRange(a1, startC, lastR, lastC).setValues(copyValue);    //該当日付行へ値コピー

  //大鵬家具　2021/03/03　集計終了
  //sheet_From = ss_From.getSheetByName('週間集計_大鵬家具');
  //ss_To = SpreadsheetApp.openById('1voxbyc0t_YNeR0Gw0ONRUGTNbXIlf0uk8hCldNsNnXQ');
  //sheet_To = ss_To.getSheetByName('SKU別売れ数');  
  //
  //lastC = sheet_From.getLastColumn();   //コピー最終列
  //copyValue = sheet_From.getRange(startR, startC, lastR ,lastC).getValues();
  //sheet_To.getRange(a1, startC, lastR, lastC).setValues(copyValue);    //該当日付行へ値コピー

  //JIN 2020/5/3-集計停止
  // sheet_From = ss_From.getSheetByName('週間集計_JIN');
  // ss_To = SpreadsheetApp.openById('1PfvhEyrRF4asXfefZnAj4lCUmMsMTwPY7V5GUVEhbx8');
  // sheet_To = ss_To.getSheetByName('SKU別売れ数');  
  // 
  // lastC = sheet_From.getLastColumn();   //コピー最終列
  // copyValue = sheet_From.getRange(startR, startC, lastR ,lastC).getValues();
  // sheet_To.getRange(a1, startC, lastR, lastC).setValues(copyValue);    //該当日付行へ値コピー

  //MTF
  //  sheet_From = ss_From.getSheetByName('週間集計_MTF');
  //  ss_To = SpreadsheetApp.openById('17DRulfss0flYFX_jrJxJajLmJFcGh41KkE7ldDHRZ40');
  //  sheet_To = ss_To.getSheetByName('SKU別売れ数');  
  //  
  //  lastC = sheet_From.getLastColumn();   //コピー最終列
  //  copyValue = sheet_From.getRange(startR, startC, lastR ,lastC).getValues();
  //  sheet_To.getRange(a1, startC, lastR, lastC).setValues(copyValue);    //該当日付行へ値コピー

  //音部・ロリアン寝装 2020/02/27-集計終了
  //  sheet_From = ss_From.getSheetByName('週間集計_音部・ロリアン寝装');
  //  ss_To = SpreadsheetApp.openById('1YJZ9y1eKLzw61Fah_9dYOrEPHjHMdv0q-AJqubHqOCc');
  //  sheet_To = ss_To.getSheetByName('SKU別売れ数');  
  //  
  //  lastC = sheet_From.getLastColumn();   //コピー最終列
  //  copyValue = sheet_From.getRange(startR, startC, lastR ,lastC).getValues();
  //  sheet_To.getRange(a1, startC, lastR, lastC).setValues(copyValue);    //該当日付行へ値コピー
  
  //フランスベッド　GS3点パック 2020/02/27-集計終了
  //  sheet_From = ss_From.getSheetByName('週間集計_FB_GS3');
  //  ss_To = SpreadsheetApp.openById('1ftyD8W0EPhe0vtDiP6vsMdBIDOx9MOVAr0XvtFz7QvM');
  //  sheet_To = ss_To.getSheetByName('SKU別売れ数');  
  //  
  //  lastC = sheet_From.getLastColumn();   //コピー最終列
  //  copyValue = sheet_From.getRange(startR, startC, lastR ,lastC).getValues();
  //  sheet_To.getRange(a1, startC, lastR, lastC).setValues(copyValue);    //該当日付行へ値コピー
  
  // 2020/5/3集計停止
  // sheet_From = ss_From.getSheetByName('週間集計_FB_その他Sup');
  // ss_To = SpreadsheetApp.openById('1QABOENmE-Znt6a00cUGPv_f-AdULd9hj4E_yVMceyRU');
  // sheet_To = ss_To.getSheetByName('SKU別売れ数');  
  // 
  // lastC = sheet_From.getLastColumn();   //コピー最終列
  // copyValue = sheet_From.getRange(startR, startC, lastR ,lastC).getValues();
  // sheet_To.getRange(a1, startC, lastR, lastC).setValues(copyValue);    //該当日付行へ値コピー
  
  //ドリームベッド 2020/02/27-集計終了
  //  sheet_From = ss_From.getSheetByName('週間集計_ドリームベッド');
  //  ss_To = SpreadsheetApp.openById('1tNvDxKCEx5dQBrLeEcVMcoNF7qfNsF1ingqDKXb5gxw');
  //  sheet_To = ss_To.getSheetByName('SKU別売れ数');  
  //
  //  lastC = sheet_From.getLastColumn();   //コピー最終列
  //  copyValue = sheet_From.getRange(startR, startC, lastR ,lastC).getValues();
  //  sheet_To.getRange(a1, startC, lastR, lastC).setValues(copyValue);    //該当日付行へ値コピー
 
  ////フランスベッド→アイテムシート分割
  ////  sheet_From = ss_From.getSheetByName('週間集計_フランスベッド');
  ////  ss_To = SpreadsheetApp.openById('1Im-XYnD5d7Y5lDqhu_Ex_ukjcncRoXEPNTj3MOJdQhQ');
  ////  sheet_To = ss_To.getSheetByName('SKU別売れ数');  
  ////  
  ////  var lastC = sheet_From.getLastColumn();   //コピー最終列
  ////  var copyValue = sheet_From.getRange(startR, startC, lastR ,lastC).getValues();
  ////  sheet_To.getRange(a1, startC, lastR, lastC).setValues(copyValue);    //該当日付行へ値コピー  
  
  //20170102_集計中止  
  //  sheet_From = ss_From.getSheetByName('週間集計_FB_FR他');
  //  ss_To = SpreadsheetApp.openById('1V5GKdopwruTwwL8wsXvyWPFFXv_OrZKkaONPan1ozz8');
  //  sheet_To = ss_To.getSheetByName('SKU別売れ数');  
  //  
  //  lastC = sheet_From.getLastColumn();   //コピー最終列
  //  copyValue = sheet_From.getRange(startR, startC, lastR ,lastC).getValues();
  //  sheet_To.getRange(a1, startC, lastR, lastC).setValues(copyValue);    //該当日付行へ値コピー  
  
  //20190312_集計中止 
  //  sheet_From = ss_From.getSheetByName('週間集計_FB_effe');
  //  ss_To = SpreadsheetApp.openById('1tG7sMc8n8ctEoQHVys8WqxgxU69CdSNw5yLV4LFc7aA');
  //  sheet_To = ss_To.getSheetByName('SKU別売れ数');  
  //  
  //  lastC = sheet_From.getLastColumn();   //コピー最終列
  //  copyValue = sheet_From.getRange(startR, startC, lastR ,lastC).getValues();
  //  sheet_To.getRange(a1, startC, lastR, lastC).setValues(copyValue);    //該当日付行へ値コピー  
  
  //  //国内その他  20190312_集計中止  
  //  sheet_From = ss_From.getSheetByName('週間集計_国内その他');
  //  ss_To = SpreadsheetApp.openById('1yl05SCUuu33lnnMWThmApkCuQk_S_cAFm9LSs9Ce28s');
  //  sheet_To = ss_To.getSheetByName('SKU別売れ数');  
  //  
  //  lastC = sheet_From.getLastColumn();   //コピー最終列
  //  copyValue = sheet_From.getRange(startR, startC, lastR ,lastC).getValues();
  //  sheet_To.getRange(a1, startC, lastR, lastC).setValues(copyValue);    //該当日付行へ値コピー
  
  
  //20200602メーカー別発注状況一覧にタイムスタンプ
  //20250526メーカー別発注状況一覧のアドレス変更
  var ss = SpreadsheetApp.openById('1X2ZHsY15-p4j0tsuqiAyJYcxh9SuYY05nLfZyD0weNk');
  var sheet = ss.getSheetByName('EON');
  sheet.getRange('A1').setValue(new Date());

  //20210304データインポート用にタイムスタンプ
  var ss = SpreadsheetApp.openById('1grFbz3UQLMu7v3gmNTZOzW08EYBqsM6iMFZ7fFF9HLs');
  var sheet = ss.getSheetByName('使い方');
  sheet.getRange('F1').setValue(new Date());


  //BtoB集計START//////////////////////////////////////////////////////////////////////////////
  //テスト用
  //  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ss = SpreadsheetApp.openById('1ZX3fVnIqupZws-2N599qLZ05auGHDOdqf04E1OGXkTU');

  //コピー元
  var sheet_From = ss.getSheetByName('work');
  //貼付先
  var sheet_To = ss.getSheetByName('BtoB');
  //データ件数
  var tmp = ss.getSheetByName('list');
  var cnt = tmp.getRange('A1').getValue() + 1;
  var rowNum = sheet_To.getLastRow();
  //該当データがなければ終了
  if ( cnt  <= 1 ){
    return;
  }else{
    //NE伝票番号と商品番号でマッチング
    var r1 = rowNum + 1;
    var wkList = sheet_From.getDataRange().getValues();
    var BtoBList = sheet_To.getDataRange().getValues();
    var flg = false;
    
    for(var i = 1; i < cnt; i++){
      flg = false
      var w_dNo = wkList[i][0];  //NE伝票番号
      var w_iNo = wkList[i][2];  //商品番号
      for(var j= 1; j < BtoBList.length; j++){
        var b_dNo = BtoBList[j][0];  //NE伝票番号
        var b_iNo = BtoBList[j][2];  //商品番号
        if(w_dNo == b_dNo && w_iNo == b_iNo){
          flg = true
          continue;
        }
      }
      //BtoBシートに存在しなければ書出し
      if(flg == false){
        for(var c = 0; c < wkList[i].length; c++) {
          var val = wkList[i][c];
          sheet_To.getRange(r1, c + 1).setValue(val);
        }
      r1 = r1 + 1;
      }
    }
  }

//BtoB集計END////////////////////////////////////////////////////////////////////////////////


//20170317ADD suzue/////////////////////////////////////////////////////////////////////////  
//店舗別売上集計START//////////////////////////////////////////////////////////////////////////
//
// //コピー元
//  var ss_From = SpreadsheetApp.openById('16NCMCm8uGLa2mn1uUIc0XAGr2Vg8YoqWraED-n-phz0');
//  var uriage_From = ss_From.getSheetByName('uriage_work');
//  var gedai_From = ss_From.getSheetByName('gedai_work');
//  //貼付先
//  var ss_To = SpreadsheetApp.openById('13JQ-tK2qTM7T5fKdEbf7Z-_qAgd3W2iQAmnLZmrid6Y');
//  var uriage_To = ss_To.getSheetByName('売上集計');
//  var gedai_To = ss_To.getSheetByName('下代集計');
//  
////貼付先の最終行判定
//　var last_row = uriage_To.getLastRow();
//
//　for(var i = last_row; i >= 1; i--){
//　　var rng = uriage_To.getRange(i, 1);
//　　if(rng.getValue() != ''){
//      var rowL = rng.getRowIndex();
//      var a1 = rowL  + 1
//      break;
//　　}
//　}
//
////日付判定
//  var d1 = rng.getValue();
//  var d2 = uriage_From.getRange('A3').getValue();
//  if(d1 != d2){
//    var d3 = d2 - d1;
//    a1 = rowL + d3   //貼付開始行
//  }else{
//    a1 = rowL
//  }
//
//  var startR = 3;     //コピー開始行
//  var startC = 1;     //コピー開始列
//  var lastR = 123;      //コピー最終行  
//
//  var lastC = uriage_From.getLastColumn();   //コピー最終列
//  var copyValue = uriage_From.getRange(startR, startC, lastR ,lastC).getValues();
//  uriage_To.getRange(a1, startC, lastR, lastC).setValues(copyValue);    //該当日付行へ値コピー 
//  
//  var copyValue = gedai_From.getRange(startR, startC, lastR ,lastC).getValues();
//  gedai_To.getRange(a1, startC, lastR, lastC).setValues(copyValue);    //該当日付行へ値コピー 
//店舗別売上集計END/////////////////////////////////////////////////////////////////////////


}
//日次更新END////////////////////////////////////////////////////////////////////////////////