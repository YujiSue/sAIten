/*********************************************************/
// Const values

const keyCol = 0;
const valCol = 2;

const qHeaderRow = 0;
const qSubheaderRow = 1;
const qStartRow = 2;
const qIdCol = 0;
const qTitleCol = 1;
const qNumCol = 2;
const qStartCol = 3;
const qqTextCol = 0;
const qqCriteriaCol = 1;
const qqImageCol = 2;
const qqRequiredCol = 3;
const qColSize = 4;

const uIdCol = 0;
const uTitleCol = 1;
const uFormCol = 2;
const uSumCol = 3;
const uURLCol = 4;

const sHeaderRow = 0;
const sStartCol = 1;


const STATUS_OK = 0;
const STATUS_NG = 1;
const STATUS_ERROR = 2;

/*********************************************************/
/** Get cell type */
function cellType(range) {
    /* Check formula */
    const formula = range.getFormula();
    if (formula !== "") return "func";
    /* Check validated */
    const rule = range.getDataValidation();
    if (rule) {
        const ctype = rule.getCriteriaType();
        const criteria = SpreadsheetApp.DataValidationCriteria;
        switch (ctype) {
            case criteria.CHECKBOX: return "checkbox";
            case criteria.VALUE_IN_LIST:
            case criteria.VALUE_IN_RANGE: return "select";
            default: 
        }
    }
    /* Check plain cell */
    const value = range.getValue();
    if (value === '') return "empty";
    else if (value instanceof Date) return "date";
    else { 
      const type = (typeof value);
      if (type === "boolean") return "bool";
      else if (type === "number") return "num";
      else if (type === "string") return"str";
      else return "any";
    }
}

/** Get property */
function getProperty(sheet) {
    /** Read properties */
    let prop = {};
    const values = sheet.getDataRange().getValues();
    for (row of values) {
        if (!row[keyCol] || row[keyCol] === '') continue;
        prop[row[keyCol]] = row[valCol];
    }
    return prop;
}
/** Get general info. */
function getInfo() { 
    const iSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('info');
    if (iSheet) return getProperty(iSheet);
    return {};  
}
/** Get preference */
function getPreference() {
    const loc = new Localizer();
    const pSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(loc.trans('pref'));
    if (pSheet) return getProperty(pSheet);
    return {}; 
}

/*********************************************************/
/** Notification email */
function sendNotification(data) {
    try {
        /* Send notification */
        GmailApp.sendEmail(
            data.to,
            data.subject, 
            `${data.header?data.header:''}\n\n${data.body?data.body:''}\n\n${data.footer?data.footer:''}`);
    } catch (e) {
        console.error(e.toString());
    }
}
/*********************************************************/
/** Check version */
function parseVersionInfo(ver) { return ver.split('.'); }
function compareVersion(v1, v2) {
    const v1_ = parseVersionInfo(v1);
    const v2_ = parseVersionInfo(v2);
    for (let v = 0; v < 3; v++) {
        if (v1_[v] < v2_[v]) return 1;
        else if (v2_[v] < v1_[v]) return -1;
    }
    return 0;
}

/*********************************************************/
/** Check email address */
function isValidEmail(email) {
    const pattern = /^[a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+$/;
    return pattern.test(email);
}
/*********************************************************/
/** Check the sheet is deployed as Web App */
function checkDeploy(props) {
    if (props.getProperty('WEB_APP')) return true;
    const webapp = ScriptApp.getService().getUrl();
    if (!webapp || webapp === '') return false;
    else {
        props.setProperty('WEB_APP', webapp);
        return true;
    }
        //showAlert('Deploy this notebook as webapp');
        //    
}
/*********************************************************/
