///////////////////////////////////////////////////////////

const pKeyCol = 1;
const pValCol = 3;


const aEvalRow = 0;
const aGeneralCol = 0;
const aEvalCol = 1;


const STATUS_OK = 0;
const STATUS_NG = 1;
const STATUS_ERROR = 2;


///////////////////////////////////////////////////////////
/** 
 * API key setter 
 */
function setApiKey() {
    /** Get property */
    const prop = getProperty(spreadsheet.getSheetByName('property'));
    /** Display prompt to enter user's API key */
    const ui = SpreadsheetApp.getUi();
    const res = ui.prompt(locale.msg.enter_ai_api_key[prop.lang]);
    if (res.getSelectedButton() === ui.Button.OK) {
        /** Set key */
        PropertiesService.getScriptProperties().setProperty("API_KEY", res.getResponseText().trim());
    }
}

/**
 * Enclose message by bracket
 */
function enclose(msg, bracket="()") {
    return `${bracket.substring(0, bracket.length/2)}${msg}${bracket.slice(-bracket.length/2)}`;
}
/**
 * Get property
 */
function getProperty(sheet) {
    /** Read properties */
    let prop = {};
    for (let r = 0; r < sheet.getLastRow(); r++) {
        const row = sheet.getRange(r + 1, 1, 1, sheet.getLastColumn()).getValues()[0];
        if (!row[pKeyCol] || row[pKeyCol] === '') continue;
        prop[row[pKeyCol]] = row[pValCol];
    }
    return prop;
}



/**
 * Mail address domain check
 */
function isAllowedDomain(email, domain) {
    const parts = email.trim().toLowerCase().split("@");
    if (parts.length !== 2) return false;
    return parts[1] === domain;
}