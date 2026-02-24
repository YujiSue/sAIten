/**
 * Return grading results: Scores and Feedback
 */
function sendMail() {
    try {
        /** Get current sheet */
        const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
        /** Get property */
        const props = getProperty(spreadsheet.getSheetByName('property'));
        /** Mail information */
        // Title
        const subject = `${props.title} ${locale.mail_sub[props.lang]}`;
        // Forwarding
        const options = {};
        if (props.forwarding && props.fwdto != '') options['cc'] = props.fwdto;
        /** Read score sheet */
        const scoreSheet = spreadsheet.getSheetByName('score');
        // Read records
        for(let r = 1; r < scoreSheet.getLastRow(); r++) {
            const row = scoreSheet.getRange(r + 1, 1, 1, scoreSheet.getLastColumn()).getValues()[0];
            // Check state
            if (row[sStateCol] != locale.pending[props.lang]) continue;
            // Get address
            const to = row[sMailCol];
            // Check address
            if (props.filtering && !isAllowedDomain(to, props.allowed)) continue;
            /** Make mail body */
            let body = "";
            body += "============================================================\n\n";
            for (let q = 0; q < props.qnum; q++) {
                const evaluated = JSON.parse(row[sQuizCol + q]);
                body += `${enclose(`Q.${q + 1}`, locale.bracket_rect[props.lang])}\n${locale['score:'][props.lang]}${evaluated['score']}\n${evaluated['note']}\n\n`;
            }
            body += "============================================================\n\n";
            /** Send mail */
            GmailApp.sendEmail(to, subject, body, options);
            scoreSheet.getRange(r + 1, stateCol).setValue(locale.done[props.lang]);
        }         
    } catch (e) {
        console.error('[Error] ' + e.toString());
    }
}
/**
 * Error Notification
 */
function notifyError(props, msg) {
    try {
        /** Send error notification */
        GmailApp.sendEmail(
            props.responsive,
            `${props.title} ${locale.err_notify_sub[props.lang]}`, 
            `${locale.err_msg[props.lang]}\n${msg}`);
    } catch (e) {
        console.error('[Error] ' + e.toString());
    }
}