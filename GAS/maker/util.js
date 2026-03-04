/*********************************************************/
// Const values

const keyCol = 0;
const valCol = 2;

const qHeaderRow = 0;
const qSubheaderRow = 1;
const qStartRow = 2;
const qIDCol = 0;
const qTitleCol = 1;
const qNumCol = 2;
const qStartCol = 3;
const qqTextCol = 0;
const qqCriteriaCol = 1;
const qqImageCol = 2;
const qqRequiredCol = 3;
const qColSize = 4;
// Default locale
const defaultLocale = {
    "gset": { "en": "General Settings", "ja": "基本設定" },
    "preference": { "en": "User Preference", "ja": "ユーザ設定" }
}


/*********************************************************/
/** Get localed word/sentence */
function localed(dict, key, lang) {
    if (key in dict) {
        if (lang in dict[key]) return dict[key][lang];
        else return dict[key].en;
    }
    else return key;
}
/*********************************************************/
/** Get property */
function getProperty(sheet) {
    /** Read properties */
    let prop = {};
    const values = sheet.getDataRange().getValues();
    for (row of gValues) {
        if (!row[keyCol] || row[keyCol] === '') continue;
        prop[row[keyCol]] = row[valCol];
    }
    return prop;
}
/*********************************************************/
/** Load general settings and user preferences */
function loadSettings(spreadsheets) {
    let properties = {
        lang: '',
        locale: {}
    };
    // Default language
    properties.lang = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetLocale().substring(0, 2);
    // Load general settings
    const gSheet = spreadsheets.getSheetByName(localed(defaultLocale, 'gset', properties.lang));
    if (!gSheet) return properties;
    Object.assign(properties, getProperty(gSheet));
    // Load locale table

    // Load user preference
    const pSheet = spreadsheets.getSheetByName(localed(defaultLocale, 'preference', properties.lang));
    if (!pSheet) return properties;
    Object.assign(properties, getProperty(pSheet));
    //
    return properties;
}


const locale = {
    require: {"en": "Required", "ja": "必須"},
    not_require: {"en": "Not Required", "ja": "任意"},
};


/** Add Question Column(s) */
function addQestionColumns(sheet, current, num, props) {
    const qRequiredRule = SpreadsheetApp.newDataValidation()
                                    .requireValueInList([
                                        localed(props.locale, 'required', props.lang),
                                        localed(props.locale, 'not_required', props.lang)
                                    ], true)
                                    .setAllowInvalid(false)
                                    .build();
    for (let i = 0; i < num; i++) {
        const col = qStartCol + (current + i) * qColSize;
        sheet.getRange(qHeaderRow + 1, col + 1, 1, qColSize).merge();
        sheet.getRange(qHeaderRow + 1, col + 1).setValue(`Q.${current + num + 1}`);
        sheet.getRange(qSubheaderRow + 1, col + qqRequiredCol + 1).setDataValidation(qRequiredRule);
    }
}
/** Expand input fields */
function addQestionRow(sheet, props) {
    // Add new row
    sheet.appendRow([""]);
    const current = int((sheet.getLastColumn() - qStartCol) / qColSize);
    const qTextRule = SpreadsheetApp.newDataValidation()
        .requireTextIsNotEmpty()
        .setHelpText("ここに問題文を入力してください")
        .build();
    const qRequiredRule = SpreadsheetApp.newDataValidation()
        .requireValueInList([
            localed(props.locale, 'required', props.lang),
            localed(props.locale, 'not_required', props.lang)
        ], true)
        .setAllowInvalid(false)
        .build();
    
    const row = sheet.getLastRow();
    for (let i = 0; i < current; i++) {
        const col = qStartCol + i * qColSize;
        sheet.getRange(row, col + qqTextCol + 1).setDataValidation(qTextRule);
        sheet.getRange(row, col + qqRequiredCol + 1).setDataValidation(qRequiredRule);
    }
    // Set background color
    //sheet.getRange(sheet.getLastRow() + 1, 1, 1, sheet.getLastColumn()).setBackground("");    
}
/** Expand input fields */
function expandField(sheet) {
    const current = int((sheet.getLastColumn() - qStartCol) / qColSize);
    const values = sheet.getDataRange().getValues();
    for ([r, row] of values.entries()) {
        if (row[qIDCol] != '') continue;
        if (current < row[qNumCol]) addQestionColumns(sheet, current, row[qNumCol] - current);
    }
}
