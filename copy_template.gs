var ROW_START_DATA = 2;
var COL_TITLE = 2;
var COL_NAME = 3;
var COL_TEL = 4;
var COL_CONTENTS = 5;

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
    {name: 'PDF出力', functionName: "savePDF"},
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
  var thisBook = SpreadsheetApp.getActiveSpreadsheet();
  var shData = thisBook.getSheetByName(NAME_SHEET_DATA);
  var shTemp = thisBook.getSheetByName(NAME_SHEET_TEMPLATE);
  var shCopy;
  
  // シートのクリア
  clearSheets();
  
  // タイトルリストを取得
  arrTitles = getActiveRowValues(shData, COL_TITLE, ROW_START_DATA);
  Logger.log(arrTitles.length);
  
  
  for(var i = 0; i < arrTitles.length; i++){
    // シートのコピー
    shCopy = shTemp.copyTo(thisBook);
    shCopy.setName((i + 1) + '_' + arrTitles[i]); // シート名を付与
    setDataOnTemplate(shData, (i + ROW_START_DATA), shCopy);
  }
  
  Browser.msgBox(MSGs[0], Browser.Buttons.OK);

}


/* --------------------------------------------------
コピーしたテンプレートに必要データを埋める 
*/
function setDataOnTemplate(fromSheet, fromIndex, toSheet){
  // 名前
  toSheet.getRange(4, 4).setValue(fromSheet.getRange(fromIndex, COL_NAME).getValue());
  
  // タイトル
  toSheet.getRange(4, 15).setValue(fromSheet.getRange(fromIndex, COL_TITLE).getValue());
  
  // TEL
  toSheet.getRange(6, 4).setValue(fromSheet.getRange(fromIndex, COL_TEL).getValue());
  
  // 申請内容
  toSheet.getRange(10, 2).setValue(fromSheet.getRange(fromIndex, COL_CONTENTS).getValue());
  
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