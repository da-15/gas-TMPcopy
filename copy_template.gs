'use strict';
/* --------------------------------------------------
 * メイン関数
 */
function main() {
  const sh = SpreadsheetApp.getActiveSpreadsheet();
  const shData = sh.getSheetByName(CONF.SHEET_NAME.SETTING);
  const shTemp = sh.getSheetByName(CONF.SHEET_NAME.TEMPLATE);
  
  // タイトル名を配列で取得
  const titles = getActiveRowValues(shData, 
    CONF.SHEET_COL.SHEET_NAME, CONF.SHEET_ROW.START_DATA);
  
  // 印刷要不要を配列で取得
  const printed =  getActiveRowValues(shData, 
    CONF.SHEET_COL.PRINT, CONF.SHEET_ROW.START_DATA);
  
  // レンジ指定行を配列で取得
  const arrRanges = getActiveColValues(shData, 
    CONF.SHEET_ROW.TITLE, CONF.SHEET_COL.START_DATA);


  // シートのクリア
  clearSheets();
  
  for(let i = 0, j = 1; i < titles.length; i++){
    if(printed[i] !== CONF.STATIC_NAME.DONE){
      // テンプレートを複製
      const shCopy = shTemp.copyTo(sh);
      // シート名をつける
      shCopy.setName((j++) + '_' + titles[i]); 
      // コピーしたシートにデータを書き出す
      setDataOnTemplate(shData, arrRanges, (i + CONF.SHEET_ROW.START_DATA), shCopy);

      // 出力したら出力済みフラグを立てる「済」
      shData.getRange(CONF.SHEET_ROW.START_DATA + i, CONF.SHEET_COL.PRINT).setValue(CONF.STATIC_NAME.DONE);
    }
  }
  
  //Browser.msgBox(MSGs[0], Browser.Buttons.OK);
}

/* --------------------------------------------------
 * コピーしたテンプレートに必要データを埋める 
 * 　fromSheet ... コピー元のシート
 *   arrRanges ... コピー先のカラム指定配列
 * 　fromIndex ... コピーするデータ行
 * 　toSheet ...   コピー先のシート
 */
function setDataOnTemplate(fromSheet, arrRanges, fromIndex, toSheet){

  for(let i = 0; i < arrRanges.length; i++){
    try{
       // 指定されたカラムにデータを書き出す
       toSheet.getRange(arrRanges[i]).setValue(
         fromSheet.getRange(
           fromIndex, CONF.SHEET_COL.START_DATA + i).getValue());
    }catch(e){
       Logger.log(e);
    }
  }
}

/* --------------------------------------------------
 * 指定した列を配列に取込
 */
function getActiveRowValues(sheet, numCol, numStartRow){
  let arrResult = [];

  if(numStartRow > 0 && numCol > 0){
    //指定された列を配列にする  
    for(let i = numStartRow; i <= sheet.getLastRow(); i++){
      arrResult.push(sheet.getRange(i, numCol).getValue());
    }
  }
  
  return arrResult;  
}

/* --------------------------------------------------
 * 指定した行を配列に取込
 */
function getActiveColValues(sheet, numRow, numStartCol){
  let arrResult = [];

  if(numStartCol > 0 && numRow > 0){
    //指定された行を配列にする  
    for(let i = numStartCol; i <= sheet.getLastColumn(); i++){
      arrResult.push(sheet.getRange(numRow, i).getValue());
    }
  }
  
  return arrResult;  
}

/* --------------------------------------------------
 * シートのクリア
 */
function clearSheets(){
  const thisBook = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = thisBook.getSheets();
  
  for(let i = 0; i < sheets.length; i++){
    // 設定シートとテンプレートシート意外を削除
    if(sheets[i].getName() != CONF.SHEET_NAME.SETTING && 
        sheets[i].getName() != CONF.SHEET_NAME.TEMPLATE){
      Logger.log(sheets[i].getName());
      thisBook.deleteSheet(sheets[i]);
    }
  }
}