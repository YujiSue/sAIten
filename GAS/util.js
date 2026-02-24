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
 * Enclose message by bracket
 */
function enclose(msg, bracket="()") {
    return `${bracket.substring(0, bracket.length/2)}${msg}${bracket.slice(-bracket.length/2)}`;
}
/**
 * Get property
 */
function getProperty(propSheet) {
    /** Read properties */
    let prop = {};
    for (let r = 0; r < propSheet.getLastRow(); r++) {
        const row = propSheet.getRange(r + 1, 1, 1, propSheet.getLastColumn()).getValues()[0];
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