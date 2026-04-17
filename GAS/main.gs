/*********************************************************/
///////////////////////////////////////////////////////////
/** 
 * Generate form and summary sheet 
 */
function generateFiles() {
  try {
    if (!props.getProperty('AUTHENTICATE')) setAuth();
    /* General data */
    let prefs = getPrefs();
    const pref = prefs['_default_'];
      
  /* Set output directory */
  const curdir = getCurrentDir();
    
  /* Get required sheets */
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const qSheet = spreadsheet.getSheetByName(trans('tab.question'));
  const uSheet = spreadsheet.getSheetByName(trans('tab.url'));

  /* Proc. pending records */
  const qValues = qSheet.getDataRange().getValues().slice(qStartRow, qStartRow + pref.max_task.value);
  for (const [r,row] of qValues.entries()) {
    /* Check record */
    if (row[qIdCol] != '' || 
        row[qNumCol] === '' || 
        row[qTitleCol] === '') continue;
        
    /* Form and summary */
    const form = makeForm(row, pref);
    const summary = copySummary(props.getProperty('summary'), row, pref, spreadsheet.getId());

    /* Link between form and sheet */
    linkToSheet(summary, row, form, summary.getId());
          
    /* Move to subdir */
    const dir = curdir.createFolder(row[qTitleCol]);
    DriveApp.getFileById(form.getId()).moveTo(dir);
    DriveApp.getFileById(summary.getId()).moveTo(dir);
        
    /* Sharing */
    if (Session.getActiveUser().getEmail() != pref.author_email.value) {
      DriveApp.getFileById(form.getId()).addEditor(pref.author_email.value);
      DriveApp.getFileById(summary.getId()).addEditor(pref.author_email.value);
    }
    
    /* Update pref */
    prefs[summary.getId()] = JSON.parse(JSON.stringify(pref));
    props.setProperty('pref', JSON.stringify(prefs));
    
    /* Publish form */
    let opened = true;
    if (pref.limit_period.value) {
      //setSchedule();
      if (!isInDate(pref.start_date.value, pref.end_date.value)) opened = false;
    }
    form.setPublished(opened);
    form.setAcceptingResponses(opened);
    setSchedule();

    /* Make index */
    let index = props.getProperty('index') ? JSON.parse(props.getProperty('index')) : {};
    index[summary.getId()] = {
      'row': r + qStartRow + 1,
      'title': row[qTitleCol],
      'form': form.getId(),
      'state' : opened ? 'open': 'close'
    };
    props.setProperty('index', JSON.stringify(index));

    /* Update sheets */
    qSheet.getRange(r + qStartRow + 1, qIdCol + 1).setValue(summary.getId());
    uSheet.getRange(r + uStartRow + 1, 1, 1, uSheet.getLastColumn())
      .setValues([[summary.getId(), row[qTitleCol], form.getEditUrl(), summary.getUrl(), 
            form.getPublishedUrl(),
            `https://api.qrserver.com/v1/create-qr-code/?data=${form.getPublishedUrl()}&size=300x300`, 
            trans((opened ?'opened':'closed'))]]);
  } 
  } catch(e) {
    console.error(e.toString());
  }   
}
///////////////////////////////////////////////////////////
/** Make a form */
function makeForm(values, pref) {
    /* Create a new form */
    const fname = `${values[qTitleCol]} ${trans('file.form')}`;
    const form = FormApp.create(fname);
    if (pref.filtering.value === 'org_only') {
        form.setDescription(trans('msg.form_desc_org_only'));
        form.setCollectEmail(true);
        form.setRequireLogin(true);
        form.setEmailCollectionType(FormApp.EmailCollectionType.VERIFIED);
    }
    else {
        form.setDescription(trans('msg.form_desc'));
        form.setCollectEmail(true);
        form.setEmailCollectionType(FormApp.EmailCollectionType.RESPONDER_INPUT);
    }
    form.setTitle(values[qTitleCol]);
    /* Set questions */
    const num = values[qNumCol];
    for (let q = 0; q < num; q++) {
        const col = qStartCol + q * qColSize;
        if (values[col + qqImageCol] != '') {
            const imgFile = DriveApp.getFileById(values[col + qqImageCol]);
            if(imgFile) {
                const blob = imgFile.getBlob();
                response.form.addImageItem()
                        .setImage(blob)
                        .setAlignment(FormApp.Alignment.CENTER);
            }
        }
        /* Add question */
        form.addParagraphTextItem()
                .setTitle(`[Q.${q+1}] ${values[col + qqTextCol]}`)
                .setRequired(values[col + qqRequiredCol] === trans('required'));
    }
    return form;
}
///////////////////////////////////////////////////////////
/** Make a summary */
function copySummary(templateId, values, pref, srcid) {
    const summaryTemplate = DriveApp.getFileById(templateId);
    const fname = `${values[qTitleCol]} ${trans('file.summary')}`;
    const copied = summaryTemplate.makeCopy(fname);

    /* Add info. */
    addToInfoSheet(copied.getId(), [
      ['qid', copied.getId()],
      ['title', values[qTitleCol]],
      ['author', JSON.stringify({ 'id': pref.author_id.value, 'name': pref.author_name.value, 'email': pref.author_email.value })],
      ['admin', (props.getProperty('admin') ? props.getProperty('admin') : Session.getActiveUser().getEmail())],
      ['source', srcid]
    ]);

    /* Update pref */
    updatePrefSheet(copied.getId(), [pGroups[pGroups.length-1]]);

    const newSheet = SpreadsheetApp.openById(copied.getId());
    const rSheet = newSheet.getSheetByName(trans('tab.result'));

    /* Expand */
    const cur = (rSheet.getLastColumn() - rStartCol - 1) / rColSize;
    const num = values[qNumCol];
    if (cur < num) rSheet.insertColumnsBefore(rSheet.getLastColumn(), rColSize * (num - cur));    
    /* Resize columns and set labels */
    for (let r = 0; r < num; r++) {
        const col = rStartCol + rColSize * r;
        // 
        rSheet.getRange(headerRow + 1, col + 1).setValue(`[Q.${r + 1}] ${values[qStartCol + qColSize * r]}`);
        rSheet.getRange(headerRow + 1, col + rrSCoreCol + 1).setValue(values[qStartCol + qColSize * r + qqScoreCol]);
        //
        for (let rr = 0; rr < rColSize; rr++) {
            rSheet.getRange(subHeaderRow + 1, col + 1, 1, rColSize).setValues(
                rSheet.getRange(subHeaderRow + 1, rStartCol + 1, 1, rColSize).getValues()
            );
            rSheet.setColumnWidth(col + rr + 1, rSheet.getColumnWidth(rStartCol + rr + 1));
        }
    }
    return newSheet;
}
///////////////////////////////////////////////////////////
/** Link form to sheet */
function linkToSheet(ss, values, form, name) {
  const num = (values.length - qStartCol) / qColSize;
  /* Link */
  form.setDestination(FormApp.DestinationType.SPREADSHEET, ss.getId());
  SpreadsheetApp.flush();
  /* Rename sheet */
  const sheets = ss.getSheets();
  for (let sheet of sheets) {
    if (sheet.getName().startsWith(trans('tab.response'))) {
      //sheet.hideSheet();
      sheet.insertColumnAfter(sheet.getLastColumn());
      sheet.getRange(1, sheet.getLastColumn()).setValue(trans('status'));
      sheet.appendRow(['-', '-']);
      for (let q = 0; q < num; q++) {
        sheet.getRange(rStartRow, rStartCol + q + 1).setValue(values[qStartCol + q * qColSize + qqCriteriaCol]);
      }
      sheet.getRange(rStartRow, sheet.getLastColumn()).setValue('-');
      sheet.hideRows(rStartRow);
      sheet.setName(trans('tab.response'));
      break;
    }
  }
  /* Set trigger */
  if (!checkTrigger('runGrading')) {
    ScriptApp.newTrigger('runGrading')
      .timeBased()
      .everyMinutes(1)
      .create();
  }
  if (!checkTrigger('runSendMails')) {
    ScriptApp.newTrigger('runSendMails')
      .timeBased()
      .everyMinutes(1)
      .create();
  }
}
///////////////////////////////////////////////////////////
/*********************************************************/