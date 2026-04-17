/*********************************************************/
/**                      Common                         **/
/*********************************************************/
///////////////////////////////////////////////////////////
function getPrefs() {
    return JSON.parse(PropertiesService.getScriptProperties().getProperty('pref'));
}
///////////////////////////////////////////////////////////
function getPref(qid='_default_') {
    const prefs = getPrefs();
    if (qid in prefs) return prefs[qid];
    else return prefs['_default_'];
}
///////////////////////////////////////////////////////////
function getLocal() {
    return JSON.parse(PropertiesService.getScriptProperties().getProperty('localize'));
}
///////////////////////////////////////////////////////////
function readDict(sheet, order=false) {
  let data = {};
  if (order) data['_key_'] = [];
  const values = sheet.getDataRange().getValues();
  for (const row of values) {
    if (row[keyCol] === '') continue;
    if (order) data['_key_'].push(row[keyCol]);
    data[row[keyCol]] = row[valCol];
  }
  return data;
}
///////////////////////////////////////////////////////////
function readProps(sheet) {
  let data = {};
  let category = '';
  const values = sheet.getDataRange().getValues();
  for (const [idx, row] of values.entries()) {
    if (row[keyCol] === '') continue;
    if (row[keyCol][0] === '#') {
      category = row[keyCol].slice(1);
      data[category] = {
        type: 'group', content: [] 
      };
    }
    else {
      const cell = sheet.getRange(idx + 1, valCol + 1);
      const ctype = cellType(cell);
      const val = ctype === 'date' ? dateStr(row[valCol]) : row[valCol];
      data[row[keyCol]] = {
        row: idx, type: ctype, value: val
      };
      if (ctype === 'select') {
        data[row[keyCol]]['options'] = cell.getDataValidation().getCriteriaValues()[0];
      }
      if (category != '') data[category].content.push(row[keyCol]);
    }
  }
  return data;
}
///////////////////////////////////////////////////////////
function loadSetting(sheetId=null) {
    /* Property */
    const props = PropertiesService.getScriptProperties();
    /* Open spreadsheet */
    const spreadsheet = sheetId ? 
                        SpreadsheetApp.openById(sheetId) : 
                        SpreadsheetApp.getActiveSpreadsheet();
    /* Set info */
    const iSheet = spreadsheet.getSheetByName('_info_');
    if (iSheet) {
        const info = readDict(iSheet, order=true);
        for(const key in info) {
            if (key === '_key_') continue;
            props.setProperty(key, info[key]);
        }
        props.setProperty('info', JSON.stringify(info['_key_']));
    }
    /* Set pref */
    const pSheet = spreadsheet.getSheetByName('_pref_');
    if (pSheet) {
        const pref = readProps(pSheet);
        for (const key in pref) {
            if ('value' in pref[key] && pref[key].value[0] === '$') pref[key].value = trans(localize, pref[key].value.slice(1));
        }
        props.setProperty('pref', JSON.stringify({'_default_': pref}));
    }
    /* Set localize */    
    const lSheet = spreadsheet.getSheetByName('_localize_');
    if (lSheet) {
        const localize = readDict(lSheet);
        props.setProperty('localize', JSON.stringify(localize));
    }
    //ss.deleteSheet(ss.getSheetByName('_info_'));
    //ss.deleteSheet(ss.getSheetByName('_pref_'));
    //ss.deleteSheet(ss.getSheetByName('_localize_'));
}
///////////////////////////////////////////////////////////
/*********************************************************/
