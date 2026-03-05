/*********************************************************/
/** Add Question Column(s) */
function addQestionColumns(sheet, current, num) {
    //
    if (sheet.getLastColumn() < qStartCol + (current + num) * qColSize) {
        sheet.insertColumnsAfter(sheet.getLastColumn() + 1, qStartCol + (current + num) * qColSize - sheet.getLastColumn());
    }
    //
    for (let i = 0; i < num; i++) {
        const col = qStartCol + (current + i) * qColSize;
        // Add header 
        sheet.getRange(qHeaderRow + 1, col + 1, 1, qColSize).merge();
        sheet.getRange(qHeaderRow + 1, col + 1).setValue(`Q.${current + i + 1}`);
        // Add subheader
        sheet.getRange(qSubheaderRow + 1, col + qqTextCol + 1).setValue(sheet.getRange(qSubheaderRow + 1, qStartCol + qqTextCol + 1).getValue());
        sheet.setColumnWidth(col + qqTextCol + 1, sheet.getColumnWidth(qStartCol + qqTextCol + 1));
        //
        sheet.getRange(qSubheaderRow + 1, col + qqCriteriaCol + 1).setValue(sheet.getRange(qSubheaderRow + 1, qStartCol + qqCriteriaCol + 1).getValue());
        sheet.setColumnWidth(col + qqCriteriaCol + 1, sheet.getColumnWidth(qStartCol + qqCriteriaCol + 1));
        //
        sheet.getRange(qSubheaderRow + 1, col + qqImageCol + 1).setValue(sheet.getRange(qSubheaderRow + 1, qStartCol + qqImageCol + 1).getValue());
        sheet.setColumnWidth(col + qqImageCol + 1, sheet.getColumnWidth(qStartCol + qqImageCol + 1));
        //
        sheet.getRange(qSubheaderRow + 1, col + qqRequiredCol + 1).setValue(sheet.getRange(qSubheaderRow + 1, qStartCol + qqRequiredCol + 1).getValue());
        sheet.setColumnWidth(col + qqRequiredCol + 1, sheet.getColumnWidth(qStartCol + qqRequiredCol + 1));
    }
}
/** Expand input fields */
function expandField(num) {
    /** */
    const loc = new Localizer();
    /** */
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const qsheet = spreadsheet.getSheetByName(loc.trans('qsheet'));
    if (!qsheet) qsheet = makeQSheet(props);
    /**  */
    const current = (qsheet.getLastColumn() - qStartCol) / qColSize;
    if (current < num) {
        addQestionColumns(qsheet, current, num - current);
    }
}
/** Copy sheet from original version */
function copySheet(localized_name) {
    const props = PropertiesService.getScriptProperties();
    const origin = SpreadsheetApp.openById(props.origin);
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    spreadsheet.insertSheet(localized_name, {template:origin.getSheetByName(localized_name)});
    return spreadsheet.getSheetByName(localized_name);
}



/*********************************************************/
/** Initialization */
function initialize() {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

    /* Localize */
    const loc = new Localizer();
    
    let pSheet = spreadsheet.getSheetByName('pref');
    if (pSheet) {
        pSheet.setName(loc.trans('pref'));
        pSheet.showSheet();
    }
    else pSheet = spreadsheet.getSheetByName(loc.trans('pref'));
    
    /* Preference */
    const pref = new Preference(pSheet);

    /* Expand field for input questions and criteria */
    if (!spreadsheet.getSheetByName(loc.trans('qsheet'))) makeQSheet();
    if (!spreadsheet.getSheetByName(loc.trans('usheet'))) makeUSheet();
    expandField(pref.getValue('max_question'));

    /* Set author's name */
    if (pref.getValue('author') === '') {
        const author = showPrompt(loc.trans('in_name_msg'));
        if (author.status) pref.setValue('author', author.text);
    }

    /* Set author's email */
    if (pref.getValue('contact') === '')
        pref.setValue('contact', Session.getActiveUser().getEmail());

    /* Set default setting */
    const info = getInfo();
    var props = PropertiesService.getScriptProperties();
    //
    props.setProperty('project', ScriptApp.getScriptId());
    props.setProperty('app', 'sAIten');
    props.setProperty('developer', info.developer);
    props.setProperty('publish', info.publish);
    props.setProperty('mode', info.mode);
    props.setProperty('version', info.version);
    props.setProperty('license', info.license);
    props.setProperty('origin', info.origin);

    /* Notification and authenticate GMail */
    GmailApp.sendEmail(Session.getActiveUser().getEmail(), `[sAIten] Notification for email service`, `The authentication has been completed successfully.`);
    props.setProperty('AUTHENTICATE', true);
}
/*********************************************************/
/** Generate form and summary sheet */
function generateFiles() {
    /*  */
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const loc = new Localizer();  
  
    /* Check deploy */
    if (!checkDeploy()) {
        showAlert('not_deploy_msg');
        return;
    }

    /* Set output directory */
    let dir;
    
    const self = DriveApp.getFileById(spreadsheet.getId());
    const parents = self.getParents();
    if (parents.hasNext()) dir = parents.next();
    else dir = DriveApp.getRootFolder();

    
    /* */
    const qname = loc.trans('qsheet');
        const qsheet = spreadsheet.getSheetByName(qname);
    if (!qsheet) qsheet = copySheet(qname);
    
    
    
    const uname = loc.trans('usheet');
    const usheet = spreadsheet.getSheetByName(uname);
    if (!usheet) usheet = copySheet(uname);
    
    /**  */
    const qValues = qsheet.getRange(qStartRow + 1, 1, qsheet.getLastRow() - qStartRow, qsheet.getLastColumn()).getValues();
    for (const [r,row] of qValues.entries()) {
        /* Check pending record */
        if (row[qIDCol] != '') continue;

        /*  */
        



        

        if (props.storage === 'current') info.outdir = dir;
        else info.outdir = dir.createFolder();
        
    
        
    }



    /* */
    const stateColIdx = 1;
    const outDirIdx = 2;
  const authorColIdx = 3;
  const contactColIdx = 4;
  const titleColIdx = 5;
  const countColIdx = 6;
  const qstart = 7;

  
  /* Proc. */
  for(let r = 2; r < qsheet.getMaxRows(); r++) {
    try {
      /* Check */
      if (qsheet.getRange(r + 1, stateColIdx).getValue() === '済') continue;
    
      /* Make files */
      const qn = qsheet.getRange(r + 1, countColIdx).getValue();
      if (!qn) break; 
      let questions = [];
      const qvalues = qsheet.getRange(r + 1, qstart, 1, 2 * qn).getValues();
      for (let q = 0; q < qn; q++) {
        questions.push({
          'text': qvalues[0][2*q],
          'eval': qvalues[0][2*q + 1]
        });
      }
      //
      const info = {
        'author': qsheet.getRange(r + 1, authorColIdx).getValue(),
        'contact': qsheet.getRange(r + 1, contactColIdx).getValue(),
        'title': qsheet.getRange(r + 1, titleColIdx).getValue(),
        'count': qn,
        'question': questions
      };
      const copied = copyTemplate(info, qsheet.getRange(r + 1, outDirIdx).getValue());

      /* Update URLs */
      usheet.getRange(r + 1, uFIDColIdx).setValue(copied['form']);
      usheet.getRange(r + 1, uSIDColIdx).setValue(copied['sheet']);
      usheet.getRange(r + 1, uFormURLIdx).setValue(copied['answer']);
      usheet.getRange(r + 1, uSharedSheetURLIdx).setValue(copied['shared']['sheet']);

      /* Update status */
      qsheet.getRange(r + 1, stateColIdx).setValue('済');
    }
    catch(e) {
      Logger.log(e.stack);
    }
  }
}
//
function copyTemplate(info, props) {
    let copied = {};
    const title = info.title;

    /* Create a form */
    const fres = createForm(info, props.about.ftemplate);
    if (!fres.status) throw new Error(res.error);

    /* Create a summary */
    const sres = createSummary(info, fres.form, props);
    if (!sres.status) throw new Error(res.error);

    /* Share */
  //const editors = [info['contact']];
  //newForm.addEditors(editors);
  //newSheet.addEditors(editors);

    /* Return */
    copied.form = fres.formId;
    copied.url = fres.url;
    copied.shared.form = fres.share;
    copied.summary = sres.sheetId;
    copied.shared.summary = sres.share;
    return copied;



    

  /* Form */
  const formTemplateID = '1Tnex5MrvcjIQQealyzz4FruiXzp3xQBorxN-vgOdkhM';
  const fromTemplate = DriveApp.getFileById(formTemplateID);
  const newForm = fromTemplate.makeCopy(title, outdir);
  const form = FormApp.openById(newForm.getId());
  form.setTitle(title);
  
  //form.setAcceptingResponses(true);
  makeQuestions(info, form);
  
  /* Sheet */
  const sheetTemplateID = '1C94IpGmvy2LLJMHqTMrwrFB6pm4ZvwjM-M2kU51_PZo';
  const sheetTemplate = DriveApp.getFileById(sheetTemplateID);
  const newSheet = sheetTemplate.makeCopy(title, outdir);
  form.setDestination(FormApp.DestinationType.SPREADSHEET, newSheet.getId());
  const ss = SpreadsheetApp.openById(newSheet.getId());
  prepareAnsSheet(info, ss);

  /* Share */
  //const editors = [info['contact']];
  //newForm.addEditors(editors);
  //newSheet.addEditors(editors);

    /* Return */
    copied.form = fres.formId;
    copied.url = fres.url;
    copied.shared.form = fres.share;
    copied.summary = sres.sheetId;
    copied.shared.summary = sres.share;
    return copied;



  copied['form'] = newForm.getId();
  copied['answer'] = form.getPublishedUrl();
  copied['sheet'] = newSheet.getId();
  copied['shared'] = {
    'form': newForm.getUrl(),
    'sheet': newSheet.getUrl()
  };
}
//
/*********************************************************/

function copyTemplate(fid, name, dir) {
    const template = DriveApp.getFileById(fid);
    return template.makeCopy(name, dir);
}

/** Create a form */
function createForm(info, props) {
    /* Result container */
    let response = {};
    //
    try {
        /* Copy from template */
        response.file = copyTemplate(props.getProperty('ftemplate'), info.title, info.outdir);
        
        /* Edit the copied form */
        response.form = FormApp.openById(response.file.getId());
        response.form.setTitle(info.title);
        for ([q, quest] of info.questions.entries()) {
            /* Add question */
            response.form.addParagraphTextItem()
                .setTitle(`[Q.${q+1}] ${quest.text}`)
                .setRequired(quest.required);

            /* Add image */
            if (quest.image != '') {
                const blob = DriveApp.getFileById(quest.image).getBlob();
                const img = response.form.addImageItem();
                img.setImage(blob)
            }
        }
        response.status = STATUS_OK;

        
        /* Store object to link with a sheet later */
        //res.url = getPublishedUrl();
        //res.form = form;
        /* Share */
        //response.form.setAcceptingResponses(true);
        //newForm.addEditors([props.contact]);
    } catch(e) {
        response.status = STATUS_ERROR;
        response.message = e.toString();
    } finally {
        /* Return */
        return response;
    }
}
/*********************************************************/
/** Create a summary sheet */
function createSummary(info, form, props) {
    /* Result container */
    let response = {};
    //
    
    try {
        /* Copy from template */
        response.file = copyTemplate(props.stemplate, info.title, info.outdir);

        const sheetTemplate = DriveApp.getFileById(props.about.stemplate);
        const newSheet = sheetTemplate.makeCopy(info.title, info.outdir);
        res.sheetId = newSheet.getId();
        res.share = newSheet.getUrl();

        /* Link to the form */
        form.setDestination(FormApp.DestinationType.SPREADSHEET, res.sheetId);
        const fprefix = localized(props.localize, 'fresponse');

        /* Edit the copied sheet */
        const spreadsheet = SpreadsheetApp.openById(res.sheetId);
        const sheets = spreadsheet.getSheets();
        for (sheet of sheets) {
            if (sheet.getName().startsWith(fprefix)) {
                sheet.setName(fprefix);
                sheet.insertColumnAfter(sheet.getLastColumn());
                sheet.getRange(1, sheet.getLastColumn()).setValue(localized(props.localize, 'status'));
            }
            else if (sheet.getName() === localized(props.localize, 'result')) {
                sheet.getRange(sHeaderRow + 1,  sStartCol + info.count * 2 + 1).setValue(localized(props.localize, 'status'));
            }
            else if (sheet.getName() === localized(props.localize, 'criteria')) {
                for ([q, quest] of info.questions.entries()) {
                    sheet.getRange(sHeaderRow + 1, sStartCol + q + 1).setValue(quest.criteria);
                }
            }
            else if (sheet.getName() === localized(props.localize, 'pref')) {
                



            }
        }
        /* Share */
        newSheet.addEditors([props.contact]);
        
        form.setAcceptingResponses(true);
        
    } catch(e) {
        res.status = false;
        res.error = e.toString();
    }
    return res;
}
/*********************************************************/

/** Update */
function updateFile() {
    const info = getInfo();
    const current = PropertiesService.getScriptProperties().getProperty('version');
    if(!current) {
        showAlert(info.init_alert_msg);
        return;
    }
    const loc = new Localizer();
    /** Get original(latest) sheet */
    const origin = SpreadsheetApp.openById(info.original);
    const latest = getProperty(origin.getSheetByName('info'))['version'];
    if (0 < comapreVersion(current, latest)) {
        showToast(loc.trans('update_start_msg'));
        


        showToast(loc.trans('update_complete_msg'));
    }
    else showAlert(loc.trans('not_need_update_msg'));
    return;
}
