//robots.txtファイルを確認
function checkRobotsTxt(url) {
  try {
    //robots.txtのURLを構築
    var robotsTxtUrl = url + '/robots.txt';
    
    //robots.txtファイルを取得
    var response = UrlFetchApp.fetch(robotsTxtUrl);
    var content = response.getContentText();
    
    //TODO:User-agentが"*"（すべてのクローラー）または自分のクローラー名が許可されているかを確認
    if (content.indexOf('User-agent: *\nDisallow:') !== -1 || content.indexOf('User-agent: YourCrawlerName\nDisallow:') !== -1) {
      return true; // スクレイピングが許可されている
    } else {
      return false; // スクレイピングが許可されていない
    }
  } catch (error) {
    // エラーが発生した場合もスクレイピングを許可しない
    return false;
  }
}

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
  let robots_url = "\"https://nikkeiyosoku.com\"" //スクレイピングルール確認用URL
  let url = "\"https://nikkeiyosoku.com/futures_volume/\"" //クエリを引っ張ってくるURL
  let query = "\"table\"" //クエリ

  //平日だけ実行
  if (isWorkday(today) == true) {
      var isScrapingAllowed = checkRobotsTxt(robots_url);
      
      if (isScrapingAllowed == true) {
        //最後尾にシートを追加するために最後尾のシートをアクティブ化
        activeLast(spreadsheet);
        addSheet(today, spreadsheet);
        activeLast(spreadsheet);

        var sheet = SpreadsheetApp.getActiveSheet(); //アクティブなシートを取得
        let overseasBuffer = sheet.getRange("A200"); //日付によってセル内容が変わるのを防ぐための捨てセル
        let domesticBuffer = sheet.getRange("L200"); 
        
        //海外tableの取得
        overseasBuffer.setFormula("=IMPORTHTML(" + url + "," + query + "," + "1" + ")"); //関数を設定して演算
        overseasBuffer.getValue(); //演算結果の取り出し
        sheet.getRange(200, 1, 230, 10).copyTo(sheet.getRange(1, 1), {contentsOnly:true}) //表をコピペ
        overseasBuffer.clear() //演算で利用したしたセルを初期状態に戻す

        //国内tableの取得
        domesticBuffer.setFormula("=IMPORTHTML(" + url + "," + query + "," + "2" + ")"); //関数を設定して演算
        domesticBuffer.getValue(); //演算結果の取り出し
        Utilities.sleep(10000); //10s間delay
        sheet.getRange(200, 12, 230, 22).copyTo(sheet.getRange(1, 12), {contentsOnly:true}) //表をコピペ
        domesticBuffer.clear() //演算で利用したしたセルを初期状態に戻す
      }
  }
}