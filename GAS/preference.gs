/*********************************************************/
/**                      Common                         **/
/*********************************************************/
/** 
 * Show preference pane 
 */
function showPreference() {
  /* Init. */
  showModal(trans('menu.pref'), { 
    'file' : "GAS/pref", 
    'width' : 500, 
    'height' : 600, 
    'preset': { 'sheets': getAssignments() }
  });
}
/** 
 * Get preference data for GUI setting 
 */
function getPrefData() {
  const sheet = props.getProperty('sheet');
  return {
    'tabs': (sheet === 'main' ? pGroups.slice(1) : pGroups),
    'prefs': getPrefs(), 
    'loc': localizer
  };
}
/** 
 * 
 */
function getAssignments() {
    let list = [{'value': '_default_', 'label': trans('pref.default')}];
    const qSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(trans('tab.question'));
    if (qSheet && qStartRow < qSheet.getLastRow()) {
        const values = qSheet.getRange(qStartRow + 1, qIdCol + 1, qSheet.getLastRow() - qStartRow, qTitleCol + 1).getValues();
        for (const record of values) {
            if (record[qIdCol] === '') continue;
            list.push({'value': record[qIdCol], 'label': record[qTitleCol]});
        }
    }
    return list;
}
/** 
 * 
 */
function updatePreference(prefs) {
  props.setProperty('pref', JSON.stringify(prefs));
  setSchedule();
}

/*********************************************************/
