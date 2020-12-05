function savePDF(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = SpreadsheetApp.getActiveSheet();
  var urlExportTo = Browser.inputBox(MSGs[3]);
  
  Logger.log(urlExportTo);
  
  if(urlExportTo == ''){
    Browser.msgBox(MSGs[4]);
  }
  else if(urlExportTo != 'cancel'){
    createPDF(urlExportTo, ss.getId(), sheet.getSheetId(), 'PDF_' + sheet.getName());
  }
}

/* 一括出力用（ただしGoogleのリソース制限にひっかかるのでWait処理をかける必要あり） 
function createPDFs(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss.getSheets();
  
  for(var i = 0; i < sheets.length; i++){
    // 設定シートとテンプレートシート意外のPDFを作成
    if(sheets[i].getName() != NAME_SHEET_DATA && sheets[i].getName() != NAME_SHEET_TEMPLATE){
      Logger.log(sheets[i].getName());
      createPDF(URL_EXPORT_PDF, ss.getId(), sheets[3].getSheetId(), 'PDF2_' + i);
      
      // Googleのリソース制限にひっかからないよう処理に時間を置く
      Utilities.sleep(3000);
    }
  }
}
*/


function createPDF(folderid, ssid, sheetid, filename){
  try{
    // PDFファイルの保存先となるフォルダをフォルダIDで指定
    var folder = DriveApp.getFolderById(folderid);
    
    // スプレッドシートをPDFにエクスポートするためのURL。このURLに色々なオプションを付けてPDFを作成
    var url = "https://docs.google.com/spreadsheets/d/SSID/export?".replace("SSID", ssid);
    
    // PDF作成のオプションを指定
    var opts = {
      exportFormat: "pdf",    // ファイル形式の指定 pdf / csv / xls / xlsx
      format:       "pdf",    // ファイル形式の指定 pdf / csv / xls / xlsx
      size:         "A4",     // 用紙サイズの指定 legal / letter / A4
      portrait:     "true",   // true → 縦向き、false → 横向き
      fitw:         "true",   // 幅を用紙に合わせるか
      sheetnames:   "false",  // シート名をPDF上部に表示するか
      printtitle:   "false",  // スプレッドシート名をPDF上部に表示するか
      pagenumbers:  "false",  // ページ番号の有無
      gridlines:    "false",  // グリッドラインの表示有無
      fzr:          "false",  // 固定行の表示有無
      gid:          sheetid   // シートIDを指定 sheetidは引数で取得
    };
    
    var url_ext = [];
    
    // 上記のoptsのオプション名と値を「=」で繋げて配列url_extに格納
    for( optName in opts ){
      url_ext.push( optName + "=" + opts[optName] );
    }
    
    // url_extの各要素を「&」で繋げる
    var options = url_ext.join("&");
    
    // API使用のためのOAuth認証
    var token = ScriptApp.getOAuthToken();

    // PDF作成
    var response = UrlFetchApp.fetch(url + options, {
                                     headers: {
                                     'Authorization': 'Bearer ' +  token
                                     }
                                     });
    
    // 
    var blob = response.getBlob().setName(filename + '.pdf');
    
    
    
    //　PDFを指定したフォルダに保存
    folder.createFile(blob);
    
    Browser.msgBox(MSGs[1], Browser.Buttons.OK);
    
  } catch(ex){
    Logger.log('ERR: ' + filename + ': ' + ex.message);
    Browser.msgBox(MSGs[2] + ex.message, Browser.Buttons.OK);
  }
}