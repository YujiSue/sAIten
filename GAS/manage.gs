/*********************************************************/
///////////////////////////////////////////////////////////
/** 
 * Create main sheet 
 */
function generateMain(){
    /* Check API Key */
    if (!props.getProperty('API_KEY')) setAPIKey();
    
    /* Get default pref and localizer */
    let pref = getPref();
    
    /* Get template */
    const main = props.getProperty('main');

    /* Current dir */
    const cur = getCurrentDir();

    /* Get author(s) information */
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();    
    const sheet = spreadsheet.getSheetByName(trans('tab.manage'));
    const values = sheet.getDataRange().getValues().slice(mStartRow);
    for ([idx, row] of values.entries()) {
        /* Check record */
        if (row[mIdCol] === '' || row[mSheetCol] != '') continue;

        /* Set output dir */
        const dir = cur.getFoldersByName(row[mGroupCol]).hasNext()?
                      cur.getFoldersByName(row[mGroupCol]).next() : 
                      cur.createFolder(row[mGroupCol]);
        
        /* Set file name */
        const fname = `${trans('file.main')}${row[mGroupCol]===''?'':`_${row[mGroupCol]}`}_${row[mIdCol]}`;

        /* Copy from template */
        let copied = DriveApp.getFileById(main).makeCopy(fname, dir);
        sheet.getRange(idx + mStartRow + 1, mSheetCol + 1).setValue(copied.getId());
        const ss = SpreadsheetApp.openById(copied.getId());

        /* Add info. */
        let iSheet = ss.getSheetByName('_info_');
        iSheet.appendRow(['admin', Session.getActiveUser().getEmail()]);
        if (pref.common_key.value && props.getProperty('API_KEY')) 
            iSheet.appendRow(['API_KEY', props.getProperty('API_KEY')]);

        /* Overwrite pref. */
        pref.author_id.value = row[mIdCol];
        pref.author_name.value = row[mNameCol];
        pref.author_email.value = row[mMailCol];
        props.setProperty('pref', JSON.stringify({'_default_': pref}));
        updatePrefSheet(copied.getId(), pGroups.slice(1));
        
        /* Expand pane */
        expandQField(ss.getSheetByName(trans('tab.question')), pref.max_question.value);

        /* Share */
        if (pref.sharing.value) {
            dir.addEditor(row[mMailCol]);
            copied.addEditor(row[mMailCol]);
            sheet.getRange(idx + mStartRow + 1, mStateCol + 1).setValue(trans('done'));
        }
        else {
            sheet.getRange(idx + mStartRow + 1, mStateCol + 1).setValue(trans('pending'));
        }
    }
}
///////////////////////////////////////////////////////////
/** 
 * Add authors to share each main sheet 
 */
function shareMain() {
    /* Get sheet */
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();    
    const sheet = spreadsheet.getSheetByName(trans('tab.manage'));
    /* Get values */
    const values = sheet.getDataRange().getValues().slice(mStartRow);
    for ([idx, row] of values.entries()) {
        /* Check record */
        if (row[mIdCol] === '' || 
            row[mSheetCol] === '' || 
            row[mStateCol] != trans('pending')) continue;
        const file = DriveApp.getFileById(row[mSheetCol]);
        /* Add editor */
        const dirs = file.getParents();
        if (dirs.hasNext()) dirs.next().addEditor(row[mMailCol]);
        file.addEditor(row[mMailCol]);
        sheet.getRange(idx + mStartRow + 1, mStateCol + 1).setValue(trans('done'));
    }
}
///////////////////////////////////////////////////////////
/*********************************************************/
