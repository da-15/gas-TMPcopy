// グローバル定数
const CONF = function (){
  return {
    SHEET_NAME: {
      SETTING: 'データ設定',
      TEMPLATE: 'TEMPLATE'
    },
    SHEET_ROW:{
      TITLE: 1,
      START_DATA: 2
    },
    SHEET_COL:{
      PRINT: 1,
      SHEET_NAME: 2,
      START_DATA: 3
    }
  };
}();

// メッセージ（日本語）
const MSGs = [
  '処理が完了しました。',
  'PDF出力が完了しました。',
  'エラー： \\n',
  '保存先のフォルダIDを指定してください。\\n\\n(例)\\nhttps://drive.google.com/drive/folders/xxx ← xxxの部分',
  'エラー： \\nフォルダIDが指定されてません。',
  ''];

/* --------------------------------------------------
 * メニュ-の表示
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
