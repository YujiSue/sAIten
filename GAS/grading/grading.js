/**
 * Grading
 */
/** Call AI directly */
function grading() {
    /** Get current sheet */
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const crtSheet = spreadsheet.getSheetByName('criteria');
    const scoreSheet = spreadsheet.getSheetByName('score');
    const formSheet = spreadsheet.getSheetByName('form');
    /** Get property */
    const props = getProperty(spreadsheet.getSheetByName('property'));
    /** Containers */
    let trial = {};
    let criteria = [];
    let summary = {
        "rows": [],
        "answers": [],
        "results": []
    };
    /** Init. */
    const crtValues = crtSheet.getRange(1, 1, 1, crtSheet.getLastColumn()).getValues()[0];
    for (let q = 0; q < props.qnum; q++) {
        summary.answers.push([]);
        summary.results.push([]);
        criteria.push(`${crtValues[0]}\n${crtValues[q + 1]}`);
    }
    /**  */    
    try {
        for (let r = 1; r < formSheet.getLastRow(); r++) {
            const row = formSheet.getRange(r + 1, 1, formSheet.getLastColumn()).getValues()[0];
            /**  */ 
            if (!(row[fMailCol] in trial)) trial[row[fMailCol]] = 1;
            else trial[row[fMailCol]]++;
            /** Check status */
            if (row[fStatusCol] === locale.pending[props.lang]) {
                formSheet.getRange(r + 1, fStatusCol + 1).setValue(locale.scoring[props.lang]);
                // Check address
                if (props.filtering && !isAllowedDomain(to, props.allowed)) {
                    formSheet.getRange(r + 1, fStatusCol + 1).setValue(locale.done[props.lang]);
                    scoreSheet.getRange(r + 1, sQuizCol + 1).setValue(JSON.stringify({"score":"failed","note":props.msg.mail_filter_error[props.lang]}));
                    scoreSheet.getRange(r + 1, sStatusCol + 1).setValue(locale.pending[props.lang]);
                    continue;
                }
                // Check trial count
                if (props.maxtrial < summary.trial[row[fMailCol]]) {
                    formSheet.getRange(r + 1, fStatusCol + 1).setValue(locale.done[props.lang]);
                    scoreSheet.getRange(r + 1, sQuizCol + 1).setValue(JSON.stringify({"score":"failed","note":props.msg.over_trial_error[props.lang]}));
                    scoreSheet.getRange(r + 1, sStatusCol + 1).setValue(locale.pending[props.lang]);
                    continue;
                }
                //
                summary.rows.push(r);
                for (let q = 0; q < props.qnum; q++) {
                    if (props.maxword < row[fQuizCol + q].length) 
                        summary.answers[q].push(row[fQuizCol + q].substring(0, props.maxword));
                    else summary.answers[q].push(row[fQuizCol + q]);
                }                
            }
            //
            if (summary.rows.length == props.maxcount) break;
        }
        /** Check records */
        if (summary.rows.length == 0) return;
        /**  */
        for (let q = 0; q < qn; q++) {
            /**  */
            const prompt = `${prompts[q]}\n${locale.stdans[props.lang]}\n${JSON.stringify(summary.answers[q])}`;
            const res = askGemini(prompt);
            /** Check response */
            if (res.status == STATUS_ERROR) {
                notifyError(props, res.error);
                throw new Error(res.error);
            }
            /** Format AI response */
            let txt = res.text;
            if (0 <= txt.indexOf('```json')) 
                txt = txt.slice(txt.indexOf('```json'));
            txt = txt.replace("```json", "");
            if (0 <= txt.indexOf("```"))
                txt = txt.slice(0, res.indexOf("```"));
            summary.results[q] = JSON.parse(txt);
        }
        /**  */
        for (let r = 0; r < summary.rows.length; r++) {
            // Copy mail address
            scoreSheet.getRange(summary.rows[r] + 1, sMailCol + 1).setValue(formSheet.getRange(summary.rows[r] + 1, fMailCol + 1).getValue());
            // Set score and comment
            for (let q = 0; q < qn; q++) {
                scoreSheet.getRange(summary.rows[r] + 1, sScoreCol + q + 1).setValue(JSON.stringify(summary.results[q][r]));
            }
            // Update status
            formSheet.getRange(summary.rows[r] + 1, fStatusCol + 1).setValue(locale.done[props.lang]);
            scoreSheet.getRange(summary.rows[r] + 1, sStatusCol + 1).setValue(locale.pending[props.lang]);
        }
    } catch (e) {
        console.error(e.toString());
        /** Reset status */
        for(row of summary.rows) {
            formSheet.getRange(row + 1, fStatusCol + 1).setValue(locale.pending[props.lang]);
        }
    }
}
/** Call the grading app managed by admin. */
function callGrading() {
    
    
    // Summerize





}


