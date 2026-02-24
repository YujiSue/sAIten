//
function manageTask() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    /** */
  const qsheet = ss.getSheetByName('問題設定用');
  const usheet = ss.getSheetByName('生成済URL');

  /* */
  const stateColIdx = 1;
  const outDirIdx = 2;
  const authorColIdx = 3;
  const contactColIdx = 4;
  const titleColIdx = 5;
  const countColIdx = 6;
  const qstart = 7;

  const uFIDColIdx = 4;
  const uSIDColIdx = 5;
  const uFormURLIdx = 6;
  const uSharedSheetURLIdx = 7;
  
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
function copyTemplate(info, dirID) {
  let copied = {};
  const outdir = DriveApp.getFolderById(dirID);
  const title = `講義小テスト ${info['title']}`;
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
  copied['form'] = newForm.getId();
  copied['answer'] = form.getPublishedUrl();
  copied['sheet'] = newSheet.getId();
  copied['shared'] = {
    'form': newForm.getUrl(),
    'sheet': newSheet.getUrl()
  };
  return copied;
}
//
function makeQuestions(info, form) {
  for (let q = 0; q < info['count']; q++) {
    form.addParagraphTextItem()
      .setTitle(`[Q.${q+1}] ${info['question'][q]['text']}`)
      .setRequired(true);
  }
}
//
function prepareAnsSheet(info, ss) {
  const names = ss.getSheets().map(s => s.getName());
  for (name of names) {
    const sheet = ss.getSheetByName(name);
    /* */
    if (name.startsWith('フォームの回答')) {
      sheet.setName('フォームの回答');
      sheet.insertColumnAfter(sheet.getLastColumn());
      sheet.getRange(1, sheet.getLastColumn()).setValue('状態');
    }
    /* */
    else if (name === 'score') {
      sheet.getRange(1, info['count'] + 2).setValue('状態');
    }
    /* */
    else if (name === 'answer') {
      for (let q = 0; q < info['count']; q++) {
        sheet.getRange(1, q + 2).setValue(info['question'][q]['eval']);
      }
    }
    /* */
    else if (name === 'info') {
      sheet.getRange(1, 2).setValue(info['title']);
      sheet.getRange(2, 2).setValue(info['contact']);
      sheet.getRange(3, 2).setValue(info['count']);
    }
  }
  return true;
}
