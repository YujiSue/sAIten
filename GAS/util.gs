/*********************************************************/
/**                      Common                         **/
/*********************************************************/
///////////////////////////////////////////////////////////
/** Const values */
const keyCol = 0;
const valCol = 1;

const headerRow = 0;
const subHeaderRow = 1;

const mStartRow = 1;
const mIdCol = 0;
const mGroupCol = 1;
const mNameCol = 2;
const mMailCol = 3;
const mSheetCol = 4;
const mStateCol = 5;

const qStartRow = 2;
const qIdCol = 0;
const qTitleCol = 1;
const qNumCol = 2;
const qStartCol = 3;
const qqTextCol = 0;
const qqCriteriaCol = 1;
const qqImageCol = 2;
const qqScoreCol = 3;
const qqRequiredCol = 4;
const qColSize = 5;

const uStartRow = 2;
const uIdCol = 0;
const uTitleCol = 1;
const uFormCol = 2;
const uSumCol = 3;
const uURLCol = 4;
const uQRCol = 5;
const uStateCol = 6;

const rStartRow = 2;
const rDateCol = 0;
const rIdCol = 1;
const rStartCol = 2;
const rrAnsCol = 0;
const rrSCoreCol = 1;
const rrNoteCol = 2;
const rColSize = 3;

const STATUS_OK = 0;
const STATUS_NG = 1;
const STATUS_ERROR = 2;

const LIMIT_TIME = 355000;
///////////////////////////////////////////////////////////
const pGroups = ['managing', 'general', 'receiver', 'ai', 'grading', 'sending'];

///////////////////////////////////////////////////////////
const OpenErrorMsg =  '[Error] This file is broken. Please copy or install again.';

///////////////////////////////////////////////////////////
const props = PropertiesService.getScriptProperties();
const localizer = initialize() ? getLocal() : {};
const requiredRule = SpreadsheetApp.newDataValidation()
  .requireValueInList([trans('required'), trans('not_required')], true)
  .setAllowInvalid(false)
  .build();
const checkRule = SpreadsheetApp.newDataValidation()
  .requireCheckbox()
  .build();
const stateRule = SpreadsheetApp.newDataValidation()
  .requireValueInList([trans('pending'), trans('ready'), trans('done')], true)
  .setAllowInvalid(false)
  .build();
const openRule = SpreadsheetApp.newDataValidation()
  .requireValueInList([trans('opened'), trans('closed')], true)
  .setAllowInvalid(false)
  .build();
/*********************************************************/
///////////////////////////////////////////////////////////
function checkTrigger(func) {
    const triggers = ScriptApp.getProjectTriggers();
    for (const trigger of triggers) {
        if (trigger.getHandlerFunction() === func) return true;
    }
    return false;
}
///////////////////////////////////////////////////////////
function removeTrigger(func) {
  const triggers = ScriptApp.getProjectTriggers();
  for (const trigger of triggers) { 
    if (trigger.getHandlerFunction() === func) {
      ScriptApp.deleteTrigger(trigger); 
      return;
    }
  }
}
///////////////////////////////////////////////////////////
function dateStr(date) {
  return `${date.getFullYear()}/${date.getMonth()+1}/${date.getDate()} ${date.getHours()}:${date.getMinutes()}`;
}
/** 
 * Date check 
 */
function isInDate(start_date, end_date) {
  const startTime = new Date(start_date);
  const endTime = new Date(end_date);
  const now = new Date();
  if (now >= startTime && now <= endTime) return true;
  else return false;
}
///////////////////////////////////////////////////////////
/** 
 * Translate/Localize 
 */
function trans(wrd) { return (localizer && wrd in localizer) ? localizer[wrd] : wrd; } 
///////////////////////////////////////////////////////////
/** 
 * Parse version info 
 */
function parseVersionInfo(ver) { return ver.split('.'); }
///////////////////////////////////////////////////////////
/** Check email address */
function isValidEmail(email) {
    const pattern = /^[a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+$/;
    return pattern.test(email);
}
///////////////////////////////////////////////////////////

/*********************************************************/
/** Update */
function updateFile() {
    const props = PropertiesService.getScriptProperties();
    /* Check initialized */
    if(!props.getProperty('project')) initialize();
    /* Get current version */
    const current = props.getProperty('version');
    /* Localize */
    const loc = getLocal();
    /* Check update availability */
    const ref = new Reference();
    if (0 < comapreVersion(current, ref.version)) {
        showToast(trans('msg.update_start'));
        /**
         * 
         * 
         */
        showToast(trans('msg.update_complete'));
    }
    else showAlert(trans('msg.not_need_update'));
    return;
}
