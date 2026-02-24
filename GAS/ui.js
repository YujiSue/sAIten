/**
 * Make menu
 */
function onOpen(e) {
  var ui = SpreadsheetApp.getUi();
  var menu = ui.createMenu('教員');
  menu.addItem('採点', 'scoring');
  menu.addItem('返却', 'sendMail');
  menu.addItem('採点結果通知メールの送信者認証', 'authSender');
  menu.addToUi();
}

