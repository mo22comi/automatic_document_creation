function EditDocument(number, template_id){
  var copy_dest_folder_id = CopyDestFolderId();
  var file_id = CopyTemplateFile(number, template_id, copy_dest_folder_id);
  var doc = DocumentApp.openById(file_id);
  var body = doc.getBody();

  var targetList;
  var list = Datalist();
  Logger.log(list);
  
  // 選択された番号のデータを取得
  for(var i = 0; i < list.length; i++){
    if(list[i]["番号"] == number){
      targetList = list[i];
      break;
    }
  }

  // 挿入する情報
  var name = (targetList["会社名"] == "")? targetList["名前"] : targetList["会社名"] + "\n" + targetList["名前"];
  var price = targetList["金額"];
  var quantity = targetList["数量"];
  var total = targetList["合計"];
  
  // 置換してダウンロードリンク生成
  body.replaceText("{name}", name).replaceText("{number}", number).replaceText("{price}", price).replaceText("{quantity}", quantity).replaceText("{total}", total);
  var downloadLink = "https://docs.google.com/document/d/" + file_id + "/export?format=doc";
  
  return downloadLink;
}


function CopyTemplateFile(num, file_id, folder_id){
  var templateFile = DriveApp.getFileById(file_id);
  var outputFolder = DriveApp.getFolderById(folder_id);
  var outputFileName = num + "_" + templateFile.getName();
  
  // テンプレートファイルをコピー
  templateFile.makeCopy(outputFileName, outputFolder);
  
  // コピー先のファイル一覧取得、コピーできていたらファイルIDを返す
  var files = DriveApp.getFolderById(folder_id).getFiles();
  
  for(var i = 0; files.hasNext(); i++){
    var file = files.next();
    if(file.getName() == outputFileName){
      return file.getId();
    }
  }
  return "";
}


function Datalist(){
  var list = [];
  var sheet = SpreadsheetObj();
  
  for(var i = 0; i < sheet.values.length; i++){
    var row = sheet.values[i];
    row = RowToHash(row, sheet.keys);
    list.push(row);
  }
  return list;
}


function ShowDialog(){
  // dialog.html をもとにHTMLファイルを生成
  // evaluate() は dialog.html 内の GAS を実行するため（ <?= => の箇所）
  var html = HtmlService.createTemplateFromFile("dialog").evaluate();
  
  // 上記HTMLファイルをダイアログ出力
  SpreadsheetApp.getUi().showModalDialog(html, "ファイルダウンロード");
}


function SpreadsheetObj(){
  // 今開いているスプレッドシートを読み込み
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getActiveSheet();

  // ヘッダーとデータ部分を読み取り
  var keys = HeaderKeys(sheet);
  var data = sheet.getRange(3, 1, sheet.getLastRow()-2, sheet.getLastColumn());
  // 値と背景色を取得
  var values = data.getValues();

  // Dictionaryのキーとデータを返す
  var obj = new Object();
  obj.keys = keys;
  obj.values = values;
  return obj;
}


/**
* returns keys located at top of spreadsheet 
*
* @param {sheet} sh Sheet class
* @return {array} array of keys
*/
function HeaderKeys(sh) {
  return sh.getRange(2,1,1, sh.getLastColumn()).getValues()[0];
}

/**
* Convert a row to key-value hash according to keys input parameter
*
* @param {array} array
* @param {array} keys
* @return {array} key-value mapped
*/
function RowToHash(array, keys) {
  var hash = {};
  array.forEach(function(value, i) {
    hash[keys[i]] = value;
  })
  return hash;
}