/**
 * Make a message for a result
 */
function resultMessage(evaluated, props) {
    return `${localed(props.locale, 'rect_bracket', props.lang).replace('*', `Q.${q+1}`)}
${localed(props.locale, 'score', props.lang)}${localed(props.locale, 'colon', props.lang)}${evaluated.score}
${evaluated.note}

`;
}
/**
 * Send grading results: Scores and Feedback
 */
function notifyResult() {
    try {
        /** Get required sheets */
        const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
        const sSheet = spreadsheet.getSheetByName(localed(props.locale, 'ssheet', props.lang));
        if (!sSheet) console.error('The summary sheet could not be found.');

        /** Get settings */
        const props = loadSettings(spreadsheet);
        
        /** Mail setting */
        // Subject
        const subject = props.mail_sub
                            .replace('$title$', props.title)
                            .replace('$author', props.author)
                            .replace('$date', props.author);
        
        // Forwarding
        const options = {};
        if (props.forwarding && props.fwd_to != '') options.bcc = props.fwd_to;
        
        // Read records
        const sValues = sSheet.getRange(sStartRow + 1, 1, sSheet.getLastRow() - sStartRow, sSheet.getLastColumn()).getValues();
        for(row of sValues) {
            // Check state
            if (row[sStateCol] != localed(props.locale, 'pending', props.lang)) continue;
            // Get address
            const to = row[sMailCol];
            // Make mail body
            let body = props.mail_header;
            body += "\n\n============================================================\n\n";
            for (let q = 0; q < props.qnum; q++) {
                if (row[sQuizCol + q] === '') continue;
                body += resultMessage(JSON.parse(row[sQuizCol + q]), props);
            }
            body += "============================================================\n\n";
            body += props.mail_footer;
            // Send mail
            GmailApp.sendEmail(to, subject, body, options);
            scoreSheet.getRange(r + 1, sStateCol).setValue(localed( props.locale,'done', props.lang));
        }         
    } catch (e) {
        console.error('[Error] ' + e.toString());
    }
}
/**
 * Error Notification
 */
function notifyError(props, prefix, error, body) {
    try {
        /** Send error notification */
        GmailApp.sendEmail(
            props.contact,
            `[Error notification] ${localed.err_notify_sub[lang]}`, 
            `${prefix} ${error}\n${body}`);
    } catch (e) {
        console.error(e.toString());
    }
}

