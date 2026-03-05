/*********************************************************/
/**                      Common                         **/
/*********************************************************/

///////////////////////////////////////////////////////////
/** Localize */ 
class Localizer {
    constructor() {
        /* init */
        this.lang;
        this.reference;
        this.dict;

        /* Get properties */
        const props = PropertiesService.getScriptProperties();
        
        /* Set lang */
        this.lang = props.getProperty('lang');
        if(!this.lang) this.setLang(SpreadsheetApp.getActiveSheet().getSpreadsheetLocale().substring(0, 2));
        
        /* Set sheet URL */
        this.reference = props.getProperty('ldict');
        if (!this.reference) {
            const info = getInfo();
            this.setDict(info.ldict);
        }
        else {
            /* Translation dict */
            const data = props.getProperty('localize');
            if (data) this.dict = JSON.parse(data);
            else this.readDict();
        }
    }
    setLang(l) {
        this.lang = l;
        props.setProperty('lang', this.lang);
    }
    setDict(ref) {
        this.reference = ref;
        props.setProperty('ldict', this.reference);
        this.readDict();
    }
    readDict() {
        const sheet = SpreadsheetApp.openById(this.reference).getSheets()[0];
        const values = sheet.getDataRange().getValues();
        for (const row of values) {
            if (row[0] != '') this.dict[row[0]] = row[1];
        }
    }
    trans(wrd) {
        if (wrd in this.dict) return this.dict[wrd];
        else return wrd;
    }
}
