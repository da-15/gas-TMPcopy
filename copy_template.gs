var ROW_TITLE = 1;
var ROW_START_DATA = 2;
var COL_START_DATA = 3;
var COL_SHEET_NAME = 2;

var NAME_SHEET_DATA = 'データ設定';
var NAME_SHEET_TEMPLATE = 'TEMPLATE';

var MSGs = [
'処理が完了しました。',
'PDF出力が完了しました。',
'エラー： \\n',
'保存先のフォルダIDを指定してください。\\n\\n(例)\\nhttps://drive.google.com/drive/folders/xxx ← xxxの部分',
'エラー： \\nフォルダIDが指定されてません。',
''];

/* --------------------------------------------------
メニュ-の表示
*/
function onOpen(){
  //メニュー配列
  var myMenu=[
    {name: '帳票を作成', functionName: "main"},
    // {name: 'PDF出力', functionName: "savePDF"},
    {name: 'シートのクリア', functionName: "clearSheets"}
  ];
  //メニューを追加
  SpreadsheetApp.getActiveSpreadsheet().addMenu("マクロ実行",myMenu);

}

/* --------------------------------------------------
メイン関数
*/
function main() {
  var arrTitles = [];
  var arrRanges = [];
  var thisBook = SpreadsheetApp.getActiveSpreadsheet();
  var shData = thisBook.getSheetByName(NAME_SHEET_DATA);
  var shTemp = thisBook.getSheetByName(NAME_SHEET_TEMPLATE);
  var shCopy;
  
  // シートのクリア
  clearSheets();
  
  // タイトル名を配列で取得
  arrTitles = getActiveRowValues(shData, COL_SHEET_NAME, ROW_START_DATA);
  Logger.log(arrTitles.length);
  
 
  // レンジ指定行を配列で取得
  arrRanges = getActiveColValues(shData, ROW_TITLE, COL_START_DATA);
  Logger.log(arrRanges.length);
  
  
  
  for(var i = 0; i < arrTitles.length; i++){
    // シートのコピー
    shCopy = shTemp.copyTo(thisBook);
    shCopy.setName((i + 1) + '_' + arrTitles[i]); // シート名をつける
    // コピーしたシートにデータを書き出す
    setDataOnTemplate(shData, arrRanges, (i + ROW_START_DATA), shCopy);
  }
  
  Browser.msgBox(MSGs[0], Browser.Buttons.OK);

}


/* --------------------------------------------------
コピーしたテンプレートに必要データを埋める 
　fromSheet ... コピー元のシート
  arrRanges ... コピー先のカラム指定配列
　fromIndex ... コピーするデータ行
　toSheet ...   コピー先のシート
*/
function setDataOnTemplate(fromSheet, arrRanges, fromIndex, toSheet){

  for(var i = 0; i < arrRanges.length; i++){
    try{
       // 指定されたカラムにデータを書き出す
       toSheet.getRange(arrRanges[i]).setValue(fromSheet.getRange(fromIndex, COL_START_DATA + i).getValue());
    }catch(e){
       Logger.log(e);
    }
  }
  
 
}

/* --------------------------------------------------
指定した列を配列に取込
*/
function getActiveRowValues(sheet, numCol, numStartRow){
  var arrResult = [];

  if(numStartRow > 0 && numCol > 0){
    //指定された列を配列にする  
    for(var i = numStartRow; i <= sheet.getLastRow(); i++){
      arrResult.push(sheet.getRange(i, numCol).getValue());
    }
  }
  
  return arrResult;  
}

/* --------------------------------------------------
指定した行を配列に取込
*/
function getActiveColValues(sheet, numRow, numStartCol){
  var arrResult = [];

  if(numStartCol > 0 && numRow > 0){
    //指定された行を配列にする  
    for(var i = numStartCol; i <= sheet.getLastColumn(); i++){
      arrResult.push(sheet.getRange(numRow, i).getValue());
    }
  }
  
  return arrResult;  
}

/* --------------------------------------------------
シートのクリア
*/
function clearSheets(){
  var thisBook = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = thisBook.getSheets();
  
  for(var i = 0; i < sheets.length; i++){
    // 設定シートとテンプレートシート意外を削除
    if(sheets[i].getName() != NAME_SHEET_DATA && sheets[i].getName() != NAME_SHEET_TEMPLATE){
      Logger.log(sheets[i].getName());
      thisBook.deleteSheet(sheets[i]);
    }
  }

}