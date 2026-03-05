


/*********************************************************/
/** Call onEdit */
function onEdit(e) {
    const loc = new Localizer();
    const range = e.range;
    const sheet = range.getSheet();
    
    /* When the max number of questions is changed */
    if (sheet.getName() === loc.trans('pref') && range.getColumn() == (valCol + 1)) {        
        const key = sheet.getRange(range.getRow(), keyCol + 1).getValue();
        if (key === 'max_question') {
            const num = range.getValue();
            expandField(num);
        }
    }
    /* When the number of questions is defined... */
    else if (sheet.getName() === loc.trans('qsheet') && range.getColumn() == (qNumCol + 1)) {
        //
        const qRequiredRule = SpreadsheetApp.newDataValidation()
              .requireValueInList([loc.trans('required'), loc.trans('not_required')], true)
              .setAllowInvalid(false)
              .build();
        // 
        const qn = range.getValue();
        const pSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(loc.trans('pref'));
        const pref = new Preference(pSheet);
        if (pref.getValue('max_question') < qn) {
            showAlert(loc.trans('over_qnum_msg').replace('%qnum%', String(qn)));
            range.setValue(pref.getValue('max_question'));
        }
        /* Active cells */
        const arange = sheet.getRange(range.getRow(), qStartCol + 1, 1, qn * qColSize);
        arange.setBackground(null);
        for (let q = 0; q < qn; q++) {
            let rrange = sheet.getRange(range.getRow(), qStartCol + q * qColSize + qqRequiredCol + 1);
            rrange.setDataValidation(qRequiredRule)
            rrange.setValue(loc.trans('required'));
        }
        /* Non-active cells */
        const start = qStartCol + qn * qColSize;
        const irange = sheet.getRange(range.getRow(), start + 1, 1, sheet.getLastColumn() - start + 1);
        irange.setBackground('gray');
    }
    return;
}
/*********************************************************/
/** Call doPost */
function findAssignment(aid, loc) {
    const qSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(loc.trans('qsheet'));
    const ids = qSheet.getRange(qStartRow+1, qIDCol+1, qSheet.getLastRow()-qStartRow, 1).getValues()[0];
    const r = ids.indexOf(aid);
    if (-1 < r) return qSheet.getRange(r+1, 1, 1, qSheet.getLastColumn()).getValues()[0];
    else return [];
}

function grading(data, loc) {
    /* Make prompts */
    const base_prompt = getPreference()['main_prompt'];
    
    const qinfo = findAssignment(data.assignment);
    if (qinfo.length == 0) return {"status": STATUS_ERROR, 'message': loc.trans('unknown_task_err') };
    
    let prompts = [];
    for (let q = 0; q < qinfo[qNumCol]; q++) {
        prompts.push({
            'criteria': qinfo[qStartCol + q * qColSize + qqCriteriaCol],
            'answers': []
        });        
    }
    
    for (const [idx,row] of data.answers.entries()) {
        for (let q = 0; q < qinfo[qNumCol]; q++) {
            prompts[idx].answers[q].push(row[aStartCol + q]);
        }
    }

    const res = askGemini(`${base_prompt}\n${JSON.stringify(prompts)}`);

    if (res.status == STATUS_OK) {
        /** Format AI response */
        let txt = res.text;
        if (0 <= txt.indexOf('```json')) 
            txt = txt.slice(txt.indexOf('```json'));
        txt = txt.replace("```json", "");
        if (0 <= txt.indexOf("```"))
            txt = txt.slice(0, res.indexOf("```"));

        data.scores = JSON.parse(txt);
        
        for (let [idx,row] of data.scores.entries()) {
            row.unshift(data.answers[idx][fMailCol]);
        }

    }
    return res;
}

/**  */
function sendScore(data, loc) {
    data.mail = [];
    /* Proc. */
    for(const row of data.scores) {
        // Get address
        const to = row[sMailCol];
        // Make mail body
        let body = data.header;
        body += "\n\n============================================================\n\n";
        for (let q = 0; q < data.qnum; q++) {
            if (row[sQuizCol + 2 * q - 1] === '') continue;
            body += `${loc.trans('bracket_rect').replace('*', `Q.${q+1}`)}\n`;
            body += `${loc.trans('score')} : ${row[sQuizCol + 2 * q - 1]}\n`;
            body += `${row[sQuizCol + 2 * q]}\n\n`;
        }
        body += "============================================================\n\n";
        body += data.footer;
        // Send mail
        try {
            GmailApp.sendEmail(to, data.subject, body, data.mailopts);
            data.mail.push(true);
        }
        catch(e) {
            console.error(e.toString());
            data.mail.push(false);
        }
    }
}

function doPost(e) {
    try {
        var response = {};
        var data = JSON.parse(e.postData.contents);
        const pref = getPreference();
        const loc = new Localizer();

        /* Grading */
        if (data.request === 'grading') {
            response = grading(data, loc);
            if (response.status == STATUS_OK) {
                response.grade = data.scores;
                /* Send scores */
                if (pref.trigger_return === 'grading') {
                    sendScore(data, loc);
                    response.mail = data.mail;
                }
            }
        }
        /* Send scores */
        else if (data.request === 'mail') {
            sendScore(data, loc);
            response.status = STATUS_OK;
            response.mail = data.mail;
        }
        /* Undefined */
        else {
            response.status = STATUS_NG;
            response.message = loc.trans('invalid_request_err');
        }
    } catch (error) {
        response.status = STATUS_ERROR;
        response.message = error.toString();
    } finally {
        return ContentService.createTextOutput(JSON.stringify(response))
                              .setMimeType(ContentService.MimeType.JSON);
    }
}

