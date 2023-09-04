//シート追加
function addSheet(today, spreadsheet) {
  let newSheet = spreadsheet.insertSheet(); //新しいシートを挿入
  newSheet.setName(Utilities.formatDate(today, "JST","yyyyMMdd")); //シートに名前を設定（今日の日付）
}

//平日か確認
function isWorkday (targetDate) {
  // targetDate の曜日を確認、週末は休む (false)
  let rest_or_work = ["REST","mon","tue","wed","thu","fri","REST"]; //日〜土
  if ( rest_or_work [targetDate.getDay ()] == "REST" ) {
    return false;
  }; 

  //祝日カレンダーを確認する
  let calJpHolidayUrl = "ja.japanese#holiday@group.v.calendar.google.com";
  let calJpHoliday = CalendarApp.getCalendarById (calJpHolidayUrl);
  if (calJpHoliday.getEventsForDay (targetDate).length != 0) { //その日に予定がなにか入っている = 祝祭日 = 営業日じゃない (false)
    return false;
  } ;

  //全て当てはまらなければ営業日 (True)
  return true;
}

//最後尾のシートへ移動
function activeLast(spreadsheet){
  var numSheets = spreadsheet.getNumSheets();
  var sheet = spreadsheet.getSheets()[numSheets-1];
  sheet.activate();
}

function main(){
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet(); //アクティブなスプレッドシート
  let today = new Date ();
  let url = "\"this is URL\"" //クエリを引っ張ってくるURL
  let query = "\"table\"" //クエリ

  //平日だけ実行
  if (isWorkday(today) == true) {

    //最後尾にシートを追加するために最後尾のシートをアクティブ化
    activeLast(spreadsheet);
    addSheet(today, spreadsheet);
    activeLast(spreadsheet);

    var sheet = SpreadsheetApp.getActiveSheet(); //アクティブなシートを取得
    let buffer = sheet.getRange("A200"); //日付によってセル内容が変わるのを防ぐための捨てセル
    
    //tableの取得
    buffer.setFormula("=IMPORTHTML(" + url + "," + query + "," + "2" + ")"); //関数を設定して演算
    buffer.getValue(); //演算結果の取り出し
    Utilities.sleep(10000); //10s間delay
    sheet.getRange(200, 1, 250, 30).copyTo(sheet.getRange(1, 1), {contentsOnly:true}) //表をコピペ
    buffer.clear() //演算で利用したしたセルを初期状態に戻す
  }
}