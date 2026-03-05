
/*********************************************************/
/** Open event handler */
function onOpen() {
    const props = PropertiesService.getScriptProperties();
    /** Get menu labels */
    const info = getInfo();

    /* Make menus */
    var ui = SpreadsheetApp.getUi();
    var menu = ui.createMenu('sAIten');
    var submenu = ui.createMenu(info.pref_menu);
    if (!props.getProperty('project')) submenu.addItem(info.init_menu, 'initialize');
    if (!props.getProperty('deploy')) submenu.addItem('deploy', 'showDeploy');
    submenu.addItem(info.ai_auth_menu, 'setAPIKey');
    //if (!props.getProperty('mail')) submenu.addItem(info.mail_auth_menu, 'authMail');
    submenu.addItem(info.user_menu, 'showPrefPane');
    menu
    .addSubMenu(submenu)
    .addSeparator()
    .addItem(info.generate_menu, 'generateFiles')
    .addItem('integrate', 'integrateResults')
    .addSeparator()
    .addItem(info.ver_menu, 'showAbout')
    .addItem(info.update_menu, 'updateFile')
    .addItem(info.help_menu, 'openHelp')
    .addToUi();
}

/*********************************************************/
/** Show alert pane */
function showAlert(message) {
    const ui = SpreadsheetApp.getUi();
    ui.alert(message);
}
/** Show prompt pane */
function showPrompt(message) {
    const ui = SpreadsheetApp.getUi();
    const res = ui.prompt(message, ui.ButtonSet.OK_CANCEL);
    if (res.getSelectedButton() === ui.Button.OK) 
        return {'status': true, 'text': res.getResponseText().trim()};
    else return {'status': false};
}
function showModal() {
    
    HtmlService.createHtmlOutput()
    const html = HtmlService.createHtmlOutputFromFile('index')
      .setWidth(400)
      .setHeight(300);
    SpreadsheetApp.getUi().showModalDialog(html, '');
}
/** Show notification */
function showToast(message) {
    SpreadsheetApp.getActiveSpreadsheet().toast(message);
}
/** Show sidebar */
function showSidebar(html) {
    SpreadsheetApp.getUi().showSidebar(html);
}
/*
function showPrefPane() {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const loc = new Localizer();
    spreadsheet.setActiveSheet(spreadsheet.getSheetByName(loc.trans('pref')));
}
*/
/*********************************************************/
/** Show version and copyright of main sheet */
function showAbout() {
    const props = PropertiesService.getScriptProperties();
    if (!props.getProperty('project')) {
        const info = getInfo();
        showAlert(info.init_alert_msg);
        return;
    }
    showAlert(`
${props.getProperty('app')} (${props.getProperty('mode')})
v${props.getProperty('version')}

Copyright (c) ${props.getProperty('publish').substring(0,4)} ${props.getProperty('developer')}

${props.getProperty('project')}
`);
}

/** Show information of each summary sheet */
function showInfo() {

}

function showDeploy() {


  //  https://script.google.com/u/0/home/projects/

}

/*********************************************************/
/** Show link to help */
function openHelp() {
    const info = getInfo();  
    const html = `<div style="font-family: sans-serif;">
      <p><a href="${info.help_url}" target="_blank">こちら</a>から詳細を確認してください。</p>
      <button onclick="google.script.host.close()" style="margin-top:10px;">閉じる</button>
      </div>`;
    const output = HtmlService.createHtmlOutput(html).setWidth(300).setHeight(150);
    SpreadsheetApp.getUi().showModalDialog(output, "Help");
}
/*********************************************************/


