/*********************************************************/
/**                      Common                         **/
/*********************************************************/
///////////////////////////////////////////////////////////
/** Initialization */
function initialize() {
  try {
    /* Check props */
    if (props.getProperty('project')) return true;
    /* Load info, pref, and localize data */
    loadSetting();
    /* Set this Project ID */
    props.setProperty('app', 'sAIten');
    props.setProperty('project', ScriptApp.getScriptId());
    return true;
  } catch(e) {
    showAlert('[Error] Failed to complete initialization.');
    console.error(e.toString());
    return false;
  }
}
///////////////////////////////////////////////////////////
/** 
 * Check author's info 
 */
function checkAuthor() {
  let prefs = getPrefs();
  /* Set author's info. */
  if (prefs._default_.author_id.value === '') prefs._default_.author_id.value = '0000';
  if (prefs._default_.author_name.value === '') {
    const res = showPrompt(trans('msg.set_author_name'));
    if (res) prefs._default_.author_name.value = res;
  }
  if (prefs._default_.author_email.value === '') 
    prefs._default_.author_email.value = Session.getActiveUser().getEmail(); 
  props.setProperty('pref', JSON.stringify(prefs));
}
///////////////////////////////////////////////////////////
/** Authentication */
function setAuth() {
  const sheet = props.getProperty('sheet');
  if (sheet === 'main') {
    try {
      const pref = getPref();
      expandQField(SpreadsheetApp.getActiveSpreadsheet().getSheetByName(trans('tab.question')), pref.max_question.value);

      /* Drive & Form auth */
      const cur = getCurrentDir();
      console.log(cur.getName());
      const form = FormApp.create('init');
      const re_form = FormApp.openById(form.getId());
      console.log(re_form.getId());
      const file = DriveApp.getFileById(form.getId());
      file.setTrashed(true);

      checkAuthor();
      if (!props.getProperty('API_KEY')) setAPIKey();

      /* Fecth auth */
      UrlFetchApp.fetch('https://httpbin.org/get');
    }
    catch(e) {
      console.error(e.toString());
    }
  }
  props.setProperty('AUTHENTICATE', true);
}
///////////////////////////////////////////////////////////
/*********************************************************/
