/*********************************************************/
/**                      Common                         **/
/*********************************************************/
///////////////////////////////////////////////////////////
/** Latest file IDs information */
class Reference {
    constructor(sheetId) {
        this.files; 
        this.version;
        this.load(sheetId);
    } 
    load(sheetId) {
        const ss = SpreadsheetApp.openById(sheetId);
        const latest = ss.getSheetByName('latest');
        this.files = {};
        const values = latest.getDataRange().getValues();
        for (const row of values) {
            this.version = row[0];
            if (!(row[1] in this.files)) this.files[row[1]] = {};
            if (!(row[2] in this.files[row[1]])) this.files[row[1]][row[2]] = {};
            this.files[row[1]][row[2]][row[3]] = row[4];
        }
    }
    getFile(sheet, mode, lang) {
        return this.files[sheet][mode][lang];
    }
}
///////////////////////////////////////////////////////////
/** Current directory */
function getCurrentDir() {
    const self = DriveApp.getFileById(SpreadsheetApp.getActiveSpreadsheet().getId());
    const parents = self.getParents();
    if (parents.hasNext()) return parents.next();
    else return DriveApp.getRootFolder();
}
///////////////////////////////////////////////////////////
/** Sub directory */
function searchDir(parent, name) {
    const children = parent.getFoldersByName(name);
    if (children.hasNext()) return children.next();
    return null;
}
///////////////////////////////////////////////////////////

function updateForm(formId, idx, text) {
    /* Open form */
    const form = FormApp.openById(formId);
    const items = form.getItems();
    /* Edit question text */
    items[idx].setTitle(text);
}
/*********************************************************/
function shareWONotify(fId, mail) {
  const resource = {
    'role': 'writer',
    'type': 'user',
    'emailAddress': mail
  };
  const options = {
    'sendNotificationEmails': false,
    'supportsAllDrives': true
  };
  try {
    Drive.Permissions.create(resource, fId, options);
  } catch (e) {
    console.error(e.toString());
  }
}
/*********************************************************/