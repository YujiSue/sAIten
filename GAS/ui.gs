/*********************************************************/
/**                      Common                         **/
/*********************************************************/
///////////////////////////////////////////////////////////
/** 
 * Show alert pane 
 */
function showAlert(message) {
  const ui = SpreadsheetApp.getUi();
  ui.alert(message);
}
///////////////////////////////////////////////////////////
/** 
 * Show prompt pane 
 */
function showPrompt(message) {
  const ui = SpreadsheetApp.getUi();
  const res = ui.prompt(message, ui.ButtonSet.OK_CANCEL);
  if (res.getSelectedButton() === ui.Button.OK) 
    return res.getResponseText().trim();
  return null;
}
///////////////////////////////////////////////////////////
/** 
 * Show modal window
 */
function showModal(title, opts) {
  const def_opt = { 'message': '', 'file' : null, 'preset': null, 'width': 400, 'height': 300};
  const option = { ...def_opt, ...opts };
  let template = option.file ? 
    HtmlService.createTemplateFromFile(option.file):
    HtmlService.createHtmlOutput(option.message);
  if (option.preset) {
    for (key in option.preset) template[key] = option.preset[key];
    const html = template.evaluate()
      .setWidth(option.width)
      .setHeight(option.height);
    SpreadsheetApp.getUi().showModalDialog(html, title);
  }
  else {
    const html = template
      .setWidth(option.width)
      .setHeight(option.height);
    SpreadsheetApp.getUi().showModalDialog(html, title);
  }
}
///////////////////////////////////////////////////////////
/** 
 * Show notification
 */
function showToast(message) {
  SpreadsheetApp.getActiveSpreadsheet().toast(message);
}
///////////////////////////////////////////////////////////
/** 
 * Show sidebar 
 */
function showSidebar(opts={'content': '', 'file' : null, 'preset': {}, 'width': 400, 'height': 300}) {
    const template = opts.file ?
                HtmlService.createHtmlOutputFromFile(opts.file) :
                HtmlService.createHtmlOutput(opts.content);
    for (key in opts.preset) html[key] = opts.preset[key];
    const html = template.evaluate()
                          .setWidth(opts.width)
                          .setHeight(opts.height);
    SpreadsheetApp.getUi().showSidebar(html);
}
///////////////////////////////////////////////////////////
/*********************************************************/

/*********************************************************/
/** 
 * Open event handler
 */
function onOpen() {
    /* Get sheet type */
    const sheet = props.getProperty('sheet');
    
    /* Make menus */
    var ui = SpreadsheetApp.getUi();
    var menu = ui.createMenu('sAIten');
    var submenu = ui.createMenu(trans('menu.setting'));
    //
    if (sheet === 'manage') {
        submenu.addItem(trans('menu.api_reg'), 'setAPIKey');
        /* Set unique menu items */
        submenu.addItem(trans('menu.pref'), 'showPreference');
        menu.addSubMenu(submenu)
            .addSeparator()
            .addItem(trans('menu.manage'), 'generateMain')
            .addItem(trans('menu.share'), 'shareMain')
            .addSeparator();
    }
    else if (sheet === 'main') {
        /* Setting */
        if (!props.getProperty('AUTHENTICATE')) submenu.addItem(trans('menu.auth'), 'setAuth');
        if (!props.getProperty('API_KEY')) submenu.addItem(trans('menu.api_reg'), 'setAPIKey');
        submenu.addItem(trans('menu.pref'), 'showPreference');
        /* Set unique menu items */
        menu.addSubMenu(submenu)
            .addSeparator()
            .addItem(trans('menu.generate'), 'generateFiles')
            .addSeparator();
    }
    else if (sheet === 'summary') {
        /* Set unique menu items */
        menu.addItem(trans('menu.return'), 'runSendMail')
          .addSeparator();
    }
    //
    menu.addItem(trans('menu.about'), 'showAbout')
        //.addItem(trans('menu.help'), 'openHelp')
        .addToUi();
}

/*********************************************************/
/** Show version and copyright of main sheet */
function showAbout() {
  //showModal()
    const props = PropertiesService.getScriptProperties();
    if (!props.getProperty('project')) initialize();
    showAlert(`
${props.getProperty('app')}
${trans(`file.${props.getProperty('sheet')}`)} v${props.getProperty('version')} [${props.getProperty('lang')}]

Copyright (c) ${dateStr(new Date(props.getProperty('publish'))).substring(0,4)} ${props.getProperty('developer')}

Project ID: ${props.getProperty('project')}
`);
}
/*********************************************************/
/** Show link to help */
function openHelp() {
  showModal(trans('menu.help'), {
    'message':`<html>${trans('msg.to_help')}<br/><a href="${props.getProperty('help')}" target="_blank">Help page</a></html>`,
    'width':360,
    'height': 200
  });
}
/*********************************************************/


