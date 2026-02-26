/**  */
function onOpen(e) {
    var ui = SpreadsheetApp.getUi();
    var menu = ui.createMenu('課題');
    menu.addItem('入力欄の拡張', 'expandField');
    menu.addItem('フォームと集計シートの生成', 'generateFiles');
    menu.addToUi();
}