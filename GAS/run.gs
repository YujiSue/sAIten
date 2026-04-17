/*********************************************************/
/**                      Schedule                       **/
/*********************************************************/
///////////////////////////////////////////////////////////
/**
 * Set/Reset schedules
 */
function setSchedule() {
  /* Schedule trigger can be controled only from "main" sheet. */
  if (props.getProperty('sheet') != 'main') return;
  
  /* Assignment Index */
  if (!props.getProperty('index')) return; // Ignrore if no assignment 
  const index = JSON.parse(props.getProperty('index'));
  
  /* URL sheet */
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const uSheet = spreadsheet.getSheetByName(trans('tab.url'));
  
  /* Pref */
  const prefs = getPrefs();
  /* Current Datetime */
  const now = new Date();
  /* Container */
  let schedules = [];
  
  /* Set schedules */
  for (const qId in index) {
    /* Get unique pref. */
    if (!(qId in prefs)) continue;
    const pref = prefs[qId];
    /* Open form */
    const form = FormApp.openById(index[qId].form);
    /* IF limited */
    if (pref.limit_period.value) {
      /* get start / end */
      const start = new Date(pref.start_date.value);
      const end = new Date(pref.end_date.value);
      /* IF pre-opened */
      if (now <= start) {
        form.setPublished(false);
        form.setAcceptingResponses(false);
        uSheet.getRange(index[qId].row, uStateCol+1).setValue(trans('closed'));
        schedules.push({
          'qid': qId,
          'date':pref.start_date.value, 
          'form': index[qId].form, 
          'do' : 'open'
        });
      }
      /* IF pre-closed */
      if (now <= end) {
        /* IF opened */
        if (start <= now) {
          form.setPublished(true);
          form.setAcceptingResponses(true);
          uSheet.getRange(index[qId].row, uStateCol+1).setValue(trans('opened'));
        }
        schedules.push({
          'qid': qId,
          'date':pref.end_date.value, 
          'form': index[qId].form, 
          'do' : 'close'
        });
      }
      /* IF closed */
      if (end < now) {
        form.setPublished(false);
        form.setAcceptingResponses(false);
        uSheet.getRange(index[qId].row, uStateCol+1).setValue(trans('closed'));
      }
    }
    /* NOT limited */
    else {
      /* Publish form */
      form.setPublished(true);
      form.setAcceptingResponses(true);
      /* State change */
      uSheet.getRange(index[qId].row, uStateCol+1).setValue(trans('opened'));
    }
  }
  /* Trigger & Props */
  if (0 < schedules.length) {    
    schedules.sort((e1, e2) => {
      return new Date(e1.date) < new Date(e2.date);
    });
    /* Lock */
    const lock = LockService.getScriptLock();
    lock.waitLock(prefs['_default_'].sync_delay.value);
    /* Set trigger */
    if (checkTrigger('runFormOpen')) removeTrigger('runFormOpen');
    ScriptApp.newTrigger('runFormOpen')
      .timeBased()
      .at(new Date(schedules[0].date))
      .create();
    /* Update schedules */
    props.setProperty('schedule', JSON.stringify(schedules));
    lock.releaseLock();
  }
}
///////////////////////////////////////////////////////////
/**
 * Form IO runner
 */
function runFormOpen() {
  /* URL sheet */
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const uSheet = spreadsheet.getSheetByName(trans('tab.url'));
  /* Assignment Index */
  const index = JSON.parse(props.getProperty('index'));
  /* Pref */
  const pref = getPref();
  /* Lock */
  const lock = LockService.getScriptLock();
  lock.waitLock(pref.sync_delay.value);
  /* Get schedules */
  const schedule = props.getProperty('schedule');
  let schedules = schedule ? JSON.parse(schedule) : [];
  /* Get current datetime */
  const now = new Date();  
  /* Check */
  while(0 < schedules.length) {
    /* Next plan */
    const next = schedules[0];
    /* Check */
    if (new Date(next.date) <= now) {
      const state = next.do === 'open';
      /* Update form */
      const form = FormApp.openById(next.form);
      form.setPublished(state);
      form.setAcceptingResponses(state);
      /* Update sheet */
      if (next.qid && next.qid in index) {
        uSheet.getRange(index[next.qid].row, uStateCol+1).setValue(trans(state?'opened':'closed'));
      }
      /* Update schedule list */
      schedules = schedules.slice(1);
    }
    else break;
  }
  /* Update props and trigger */
  removeTrigger('runFormOpen');
  if (0 < schedules.length) {
    props.setProperty('schedule', JSON.stringify(schedules));
    ScriptApp.newTrigger('runFormOpen')
      .timeBased()
      .at(new Date(schedules[0].date))
      .create();
  }
  else props.deleteProperty('schedule');
  lock.releaseLock();
}
///////////////////////////////////////////////////////////
/**
 * Form direct IO runner
 */
function runDirectFormOpen() {
  const pref = getPref();
  /* Lock */
  const lock = LockService.getScriptLock();
  lock.waitLock(pref.sync_delay.value);
  /* Get form list to change IO state */
  let list;
  const fio = props.getProperty('formio'); 
  if (fio) list = JSON.parse(fio);
  else list = [];
  /* Clear props */
  props.deleteProperty('formio');
  lock.releaseLock();
  
  /* Check */
  if (list.length == 0) return;
  /* Assignment Index */
  const index = JSON.parse(props.getProperty('index'));

  /* Change states */
  for (const elem of list) {
    /* Open form */
    const form = FormApp.openById(index[elem.qid].form);
    /* Update state */
    form.setPublished(elem.state);
    form.setAcceptingResponses(elem.state);
  }
}
/*********************************************************/

/*********************************************************/
/**                       Grading                       **/
/*********************************************************/
///////////////////////////////////////////////////////////
function checkGrading() {
  let checked = [];
  try{
    const index = JSON.parse(props.getProperty('index'));
    for (const qId in index) {
      const pref = getPref(qId);
      let count = {};
      let list = [];

      const ss = SpreadsheetApp.openById(qId);
      const sheet = ss.getSheetByName(trans('tab.response'));
      const rSheet = ss.getSheetByName(trans('tab.result'));
      const values = sheet.getDataRange().getValues();
      const num = sheet.getLastColumn() - rStartCol - 1;
      const stateCol = sheet.getLastColumn() - 1;
    
      for(const [idx, row] of values.entries()) {
        /* Count */
        if (row[rIdCol] in count) count[row[rIdCol]] += 1;
        else count[row[rIdCol]] = 1;
        /* Check */
        if (row[stateCol] != '') continue;
        sheet.getRange(idx + 1, sheet.getLastColumn()).setValue('-');
        rSheet.getRange(idx + 2, 1, 1, 2).setValues([[
          row[rDateCol], row[rIdCol]
          ]]);
        /* Check user */
        if (pref.filtering.value === 'address_filter' &&
          !checkAddress(row[rIdCol], pref.allow_domain.value)) {
          rSheet.getRange(idx + 2, rStartCol + rrSCoreCol + 1, 1, 2).setValues([[
            trans('error'), trans('msg.not_allowed_user')
            ]]);
          sheet.getRange(idx + 1, sheet.getLastColumn()).setValue(trans('done'));
          rSheet.getRange(idx + 2, rSheet.getLastColumn()).setValue(trans('done'));
          continue;
        }
        /* Check count */
        if (pref.limit_trial.value &&
          pref.max_trial.value < count[row[rIdCol]]) {
          rSheet.getRange(idx + 2, rStartCol + rrSCoreCol + 1, 1, 2).setValues([[
            trans('error'), trans('msg.over_trial')
            ]]);
          sendNotification({
            'to': row[rIdCol],
            'subject': pref.mail_sub.value.replace('%TITLE%', index[qId].title),
            'body': trans('msg.over_trial')
          });
          sheet.getRange(idx + 1, sheet.getLastColumn()).setValue(trans('done'));
          rSheet.getRange(idx + 2, rSheet.getLastColumn()).setValue(trans('done'));
          continue;
        }
        /* Check text */
        for (let q = 0; q < num; q++) {
          /* Get response text */
          let txt = row[rStartCol + q] === '' ? '' : String(row[rStartCol + q]);
          /* Check response */
          if (checkResSize(txt, pref)) {
            if (pref.max_words.value < txt.split(' ').length) 
              txt = txt.split(' ').slice(0, pref.max_words.value).join(' ');
            if (pref.max_chars.value < txt.length) 
              txt = txt.substring(0, pref.max_chars.value);
          }
          rSheet.getRange(idx + 2, rStartCol + q * rColSize + 1).setValue(txt);
        }
        list.push(idx + 1);
        sheet.getRange(idx + 1, sheet.getLastColumn()).setValue(trans('pending'));
        if (list.length == pref.max_count.value) break;
      }
      if (0 < list.length) {
        checked.push({'id': qId, 'list': list});
      }
    }
  } catch(e) {
    console.error(e.toString());

  }
  finally {
    return checked;
  }
}
///////////////////////////////////////////////////////////
function runGrading() {
  try { 
    let waiting = props.getProperty('waitg') ? JSON.parse(props.getProperty('waitg')): null;
    if (!waiting || waiting.length == 0) waiting = checkGrading();
    if (waiting.length == 0) return;
    const index = JSON.parse(props.getProperty('index'));
    const que = waiting[0];
    const qId = que.id;
    const pref = getPref(qId);
    /* Lock */
    const lock = LockService.getScriptLock();
    lock.waitLock(pref.sync_delay.value);
    waiting = waiting.slice(1);
    if (waiting.length == 0) props.deleteProperty('waitg');
    else props.setProperty('waitg', JSON.stringify(waiting));
    lock.releaseLock();
    /*  */
    const spreadsheet = SpreadsheetApp.openById(qId);
    const sheet = spreadsheet.getSheetByName(trans('tab.response'));
    const rSheet = spreadsheet.getSheetByName(trans('tab.result'));
    const num = sheet.getLastColumn() - rStartCol - 1;

      let questions = [];
      let responses = [];

      const header = sheet.getRange(1, rStartCol + 1, 1, num).getValues()[0];
      const footer = sheet.getRange(sheet.getLastRow(), rStartCol + 1, 1, num).getValues()[0];
      
      for(let q = 0; q < num; q++) {
        questions.push({
          'text': header[q],
          'point': rSheet.getRange(headerRow + 1, rStartCol + q * rColSize + rrSCoreCol + 1).getValue(),
        });
        responses.push({
          'question': header[q],
          'criteria': footer[q],
          'answers': []
        });
      }
      const rows = que.list;
      for (const row of rows) {
        sheet.getRange(row, sheet.getLastColumn()).setValue(trans('running'));
        for(let q = 0; q < num; q++) {
          responses[q].answers.push(String(rSheet.getRange(row+1, rStartCol + q * rColSize + 1).getValue()));
        }
      }
      /*  */
      const prompt = `${pref['base_prompt_header'].value}\n${pref['base_prompt'].value.replace('%NUM%', num)}\n${JSON.stringify(responses)}`;

      /* Call API */
      const result = askGemini(prompt, getEndpoint(pref), getSafety(pref));

      if (result.status == STATUS_OK) {
        /* Parse AI's answer */
        const answers = parseAIAnswer(result.text);

        /* Check AI answers */
        if (!checkAnser(answers, responses)) {
          /* Set flag */
          for (const row of rows) {
            sheet.getRange(row, sheet.getLastColumn()).setValue('');
          }
          return;
        }
        /* Arrange data */
        for (const [idx,row] of rows.entries()) {
          for (let q = 0; q < num; q++) {
            rSheet.getRange(row + 1, rStartCol + q * rColSize + rrSCoreCol + 1, 1, 2)
              .setValues([[answers[idx][q].score, answers[idx][q].note]]);
          }
          sheet.getRange(row, sheet.getLastColumn()).setValue(trans('done'));
          /* Return result */
          if (pref.trigger_return.value === 'auto') {
            rSheet.getRange(row + 1, rSheet.getLastColumn()).setValue(trans('ready'));
            /* Send mail */
            const result2 = sendScore({
              'id': qId,
              'title': index[qId].title,
              'author': {
                'id':pref.author_id.value, 
                'name':pref.author_name.value, 
                'email': pref.author_email.value
              },
              'questions': questions,
              'result': rSheet.getRange(row + 1, 1, 1, rSheet.getLastColumn()).getValues()[0]
            });
            /* IF successed */
            if (result2.status == STATUS_OK) 
              rSheet.getRange(row + 1, rSheet.getLastColumn()).setValue(trans('done'));
          }
          else {
            rSheet.getRange(row + 1, rSheet.getLastColumn()).setValue(trans('pending'));
            rSheet.getRange(row + 1, rSheet.getLastColumn()).setDataValidation(stateRule);
          }
        }
      }
      else {
        console.error(res.message);
        /* Set flag */
        for (const row of rows) {
          sheet.getRange(row, sheet.getLastColumn()).setValue('');
        }
        /* Notification */
        if (props.getProperty('admin')) {
          sendNotification({
            'to': props.getProperty('admin'),
            'subject': '[sAIten] Error notification',
            'body': `Error @ 'Grading'¥n${res.message}`
          })
        }
      }
    
  } catch(e) {
    console.error(e.toString());
  }
}
///////////////////////////////////////////////////////////
/*********************************************************/

/*********************************************************/
/**                        Mail                         **/
/*********************************************************/
///////////////////////////////////////////////////////////
/** 
 * Send notification email
 * Args: {to, subject, body, option}
 */
function sendNotification(data) {
    try {
        /* Send notification */
        GmailApp.sendEmail(
            data.to,
            data.subject, 
            `${data.header?data.header:''}\n\n${data.body?data.body:''}\n\n${data.footer?data.footer:''}`,
            data.opt);
    } catch (e) {
        console.error(e.toString());
    }
}
///////////////////////////////////////////////////////////
/*********************************************************/
///////////////////////////////////////////////////////////
/**
 * Send mail from summary sheet
 */
function runSendMail() {
  try {
    sendMail(SpreadsheetApp.getActiveSpreadsheet());
  }
  catch(e) {
    console.error(e.toString());
  }
}
///////////////////////////////////////////////////////////
/**
 * Send mail from main sheet
 */
function runSendMails() {
  try{
    if (!props.getProperty('index')) return;
    const index = JSON.parse(props.getProperty('index'));
    for (const key in index) {
      const pref = getPref(key);
      if (pref.trigger_return.value != 'auto') continue;
      const spreadsheet = SpreadsheetApp.openById(key);
      sendMail(spreadsheet);
    }
  }
  catch(e) {
    console.error(e.toString());
  }
}
///////////////////////////////////////////////////////////
function sendMail(ss) {
  try {
    const sheet = ss.getSheetByName(trans('tab.result'));
    if (sheet.getLastRow() == rStartRow) return;
    const num = (sheet.getLastColumn() - rStartCol - 1) / rColSize;

    /* Get header info */
    let questions = [];
    const header = sheet.getRange(headerRow + 1, 1, 1, sheet.getLastColumn()).getValues()[0];    
    for (let q = 0; q < num; q++) {
      questions.push({
        'text': header[rStartCol + q * rColSize],
        'point': header[rStartCol + q * rColSize + rrSCoreCol]
      });
    }
    /* Get responsed data */
    const stateCol = sheet.getLastColumn() - 1;
    const values = sheet.getRange(rStartRow + 1, 1, sheet.getLastRow() - rStartRow, sheet.getLastColumn()).getValues();
    for (const [idx,row] of values.entries()) {
      /* Check status */
      if (row[stateCol] === trans('ready')) {
        let data;
        if (props.getProperty('sheet') === 'main') {
          const index = JSON.parse(props.getProperty('index'));
          const pref = getPref();
          data = {
            'id': ss.getId(),
            'title': index[ss.getId()].title,
            'author': {
              "id":pref.author_id.value,
              "name":pref.author_name.value,
              "email":pref.author_email.value
            },
            'questions': questions,
            'result': row
          };
        }
        else if (props.getProperty('sheet') === 'summary') {
          data = {
            'id': props.getProperty('qid'),
            'title': props.getProperty('title'),
            'author': JSON.parse(props.getProperty('author')),
            'questions': questions,
            'result': row
          };
        }
        else continue;
        const res = sendScore(data);
        /* Successed */
        if (res.status == STATUS_OK) 
          sheet.getRange(rStartRow + idx + 1, sheet.getLastColumn()).setValue(trans('done'));
      }
    }
  }
  catch(e) {
    console.error(e.toString());
  }
}
///////////////////////////////////////////////////////////
/**
 * Send score mail
 */
function sendScore(data) {
  let response = {};
  try {
    const pref = getPref(data.id);
    const to = data.result[rIdCol];
    const subj = pref['mail_sub'].value.replace('%TITLE%', data.title);
    const opt = pref.forwarding.value ? {'cc':pref.fwd_to.value} : {};
    let body = `${pref['mail_header'].value}\n\n`;
    body += "============================================================\n\n";
    for (const [q, question] of data.questions.entries()) {
      body += `${question.text}\n`;
      body += `${trans('<>').replace('*', trans('answer'))}\n`;
      body += `${data.result[rStartCol + q * rColSize]}\n`;

      body += `${trans('<>').replace('*', trans('score'))}${trans(':')}`;
      if (pref.hide_score.value) body += `${trans('msg.response_hidden')}\n`;
      else body += `${data.result[rStartCol + q * rColSize + rrSCoreCol]} / ${question.point}\n`;

      body += `${trans('<>').replace('*', trans('note'))}\n`;
      if (pref.hide_note.value) body += `${trans('msg.response_hidden')}\n\n`;
      else body += `${data.result[rStartCol + q * rColSize + rrNoteCol]}\n\n`;
    }
    body += "============================================================\n\n";
    /* Send mail */
    GmailApp.sendEmail(to, subj, body, opt);
    //
    response.status = STATUS_OK;
  }
  catch(e) {
    response.status = STATUS_ERROR;
    response.message = e.toString();
  }
  finally {
    return response;
  }
}
///////////////////////////////////////////////////////////
/*********************************************************/
///////////////////////////////////////////////////////////
/**
 * Address filter
 */
function checkAddress(mail, domain) { return mail.endsWith(`@${domain}`); }
///////////////////////////////////////////////////////////
/*********************************************************/
///////////////////////////////////////////////////////////
function checkResSize(res, pref) {
  return pref.max_words.value < res.split(' ').length || 
        pref.max_chars.value < res.length;
}
///////////////////////////////////////////////////////////
function parseAIAnswer(text) {
  let str = text;
  if (str.includes('```json')) {
    str = str.slice(str.indexOf('```json') + "```json".length);
  }
  if (str.includes("```")) {
    str = str.slice(0, str.indexOf("```"))
  }
  try{
    const answers = JSON.parse(str);
    return answers;
  }
  catch(e) { return null; }
}
///////////////////////////////////////////////////////////
function checkAnser(answers, responses) {
  if (!answers ||
      answers.length != responses[0].answers.length ||
      answers[0].length != responses.length) return false;
  return true;
}
/*********************************************************/

