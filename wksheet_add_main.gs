//勤務状況記入シート作成スクリプト

//スプレッドシート読み込み
var spreadsheet = SpreadsheetApp.openById('XXXXXXXXXX');

//明日の日にち取得
var date = new Date();
date.setDate(date.getDate() + 1);
//yyyyMMdd型に変換
var now = Utilities.formatDate(date,"JST","yyyyMMdd");
var sheet_name;
var today = new Date();

//日本の祝日判定　Googleカレンダーに登録されている祝日を判定
function isHoliday(date){
  var calendars = CalendarApp.getCalendarsByName("日本の祝日");
  var holidays = calendars[0];
  var events = holidays.getEventsForDay(date);
 
  if(events.length > 0)
  {
    return true;
  }
  return false;
}
 
//Date.getDay()は曜日番号として日曜始まりで0~6の値を返す
function isWeekend(date)
{
  var dayNum = date.getDay();
　//土日の場合はtrueを返す　0：日曜　1：土曜
  if(dayNum == 0 || dayNum == 6)
  {
    return true;
  }
　//平日の場合
  return false;
}

//シート作成処理
function work_sheet_add(now) 
{
  var month = now.substr( 4, 2 );
  var day = now.substr( 6, 2 );
  //シート名作成
  sheet_name = now.substr( 4, 4 );
  
  //シート追加　雛形をコピーしてシート名変更
  var sheet_temp = spreadsheet.getSheetByName("雛形"); //tempシートコピー
  var sheet_copy = sheet_temp.copyTo(spreadsheet);　　 //コピーを使用して新規のシート作成
  try
  {
  　//作成したシート名を変更
    sheet_copy.setName(sheet_name);
    //シートの位置を左から２番目にする
    spreadsheet.setActiveSheet(sheet_copy);
    spreadsheet.moveActiveSheet(2);
    
    //コピーしたシートのA１に【○○/□□：勤務状況】の形式の文言を追記1
    var sheet = spreadsheet.getSheetByName(sheet_name);
    var day_name = '【' + month + '/' + day + '：勤務状況】';
    sheet.getRange('A1').setValue(day_name);
  }
  catch(e)
  {
    //シートが既に存在した場合シート名を変更して削除
    sheet_copy.setName('削除シート');
    spreadsheet.deleteSheet(sheet_copy);
  }
}

//メイン処理
function main()
{
  //明日が営業日であれば新しいシートを作成する。
  if(!(isWeekend(date) ||　isHoliday(date)))
  {
  　//シート作成処理実行
    work_sheet_add(now);
  }
  else
  {
    var k = 1;
    //明日が営業日でない場合、営業日になるまで+1
    for(let i = 1; k == 1; i++)
    {
      date.setDate(date.getDate() + 1);
      if(!(isWeekend(date) ||　isHoliday(date)))
      {
      　//シート追加処理を実行したらループ終了
        now = Utilities.formatDate(date,"JST","yyyyMMdd");
        work_sheet_add(now);
        k++;
      }
  　}
  }
}
