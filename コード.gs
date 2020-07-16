function myFunction() {
  // スプレッドシート取得
  const sheet = SpreadsheetApp.getActive().getSheetByName('シート1');
  
  // データ取得範囲指定
  const row = 4;
  const column = 1;  
  const LastRow = sheet.getDataRange().getLastRow();
  const LastColumn = sheet.getDataRange().getLastColumn();
  const numRows = LastRow - row + 1;
  const numColumns = LastColumn - column + 1;  
  
  // データ取得
  let data = sheet.getRange(row, column, numRows, numColumns).getValues();

  // メール内容の変数
  let recipient = "アドレス1,アドレス2";
  let title = "※自動送信テスト【エンゲージメントスコア面談】対象者のお知らせ";
  let body = "面談の対象者を自動送信します。担当者は日程調整をお願いいたします。\n\n"

  // 面談対象者を調べて本文内に追記
  let date = new Date();
  date.setMonth(date.getMonth() - 1);
  date.setDate(1);
  let targetMonth = Utilities.formatDate(date, 'Asia/Tokyo', 'yyyy/MM/dd');
  let targetCnt = 0;

  for (let i = 0; i < data.length; i++){
    if ( Utilities.formatDate(new Date(data[i][0]), 'Asia/Tokyo','yyyy/MM/dd') == targetMonth && data[i][3] < 60){
      body += data[i][2] + "," + data[i][3] + "\n"
      targetCnt ++;
    }
  }
   
  if ( targetCnt == 0 ) {
    body += "対象者無し"; 
  }
  
  body +=  "\nこのメールは以下のスプレッドシートをもとに自動送信をしています。"+
           "\nツール→スクリプトエディタでコードが確認・編集できます。"+
           "\nスプレッドシートのURL";
  
  // メールを送信する
  GmailApp.sendEmail(recipient, title, body);

}
