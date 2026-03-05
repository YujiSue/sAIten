/*********************************************************/
/**                      Common                         **/
/*********************************************************/

///////////////////////////////////////////////////////////
/** Preference */ 
class Preference {
    constructor(sheet) {
        this.sheet = sheet;
        this.structure = [];
        this.preference = {};
        this.loadPreference();
    }
    loadPreference() {
        const values = this.sheet.getDataRange().getValues();
        let category = '';
        for (const [idx, row] of values.entries()) {
            if (row[keyCol] === '') continue;
            if (row[keyCol][0] === '#') {
                category = row[keyCol].slice(1);
                this.structure.push({title:category, keys: []});
            }
            else {
                this.preference[row[keyCol]] = {
                    'row': idx,
                    'category': category,
                    'value': row[valCol],
                    'klabel': row[keyCol+1],
                    'vlabel': row[valCol+1]
                };
                this.structure[-1].keys.push(row[keyCol]);
            }
        }
    }
    hasKey(key) { return key in this.preference; }
    getType(key) {}
    getChoice(key) {}
    getValue(key) {
        if (key in this.preference) return this.preference[key].value;
        return null;
    }
    setValue(key, val) {
        if (key in this.preference) {
            this.preference[key].value = val;
            this.sheet.getRange(this.preference[key].row + 1, valCol + 2).setValue(val);
        }
        else showAlert(`${key} is not defined.`);
    }
    writeOut(sheet) {
        sheet.setName('pref');
        let r = 0;
        for (const cat of this.structure) {
            
        }
    }
}
///////////////////////////////////////////////////////////
/*********************************************************/