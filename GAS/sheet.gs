/*********************************************************/
///////////////////////////////////////////////////////////
/** 
 * Get cell type 
 */
function cellType(range) {
    /* Check formula */
    const formula = range.getFormula();
    if (formula !== "") return "func";
    /* Check validated */
    const rule = range.getDataValidation();
    if (rule) {
        const ctype = rule.getCriteriaType();
        const criteria = SpreadsheetApp.DataValidationCriteria;
        switch (ctype) {
            case criteria.CHECKBOX: return "checkbox";
            case criteria.VALUE_IN_LIST: return "select";
            default: 
        }
    }
    /* Check plain cell */
    const value = range.getValue();
    if (value === '') return "empty";
    else if (value instanceof Date) return "date";
    else { 
        const type = (typeof value);
        if (type === "boolean") return "bool";
        else if (type === "number") return "num";
        else if (type === "string") return"str";
        else return "any";
    }
}
///////////////////////////////////////////////////////////

///////////////////////////////////////////////////////////
/** 
 * Expand input fields
 */
function expandQField(sheet, num) {
    /* Get current qnum */
    const cur = (sheet.getLastColumn() - qStartCol) / qColSize;
    /* Add columns */
    if (cur < num) {
        sheet.insertColumnsAfter(sheet.getLastColumn(), (num - cur) * qColSize);
        /* Adjust */
        for (let q = cur; q < num; q++) {
            const col = qStartCol + q * qColSize;
            /* Set header */
            sheet.getRange(headerRow + 1, col + 1).setValue(`Q.${q+1}`);
            for (let qq = 0; qq < qColSize; qq++) {
                /* Set subheader */
                sheet.getRange(subHeaderRow+1, col + 1, 1, qColSize).setValues(
                  sheet.getRange(subHeaderRow+1, qStartCol + 1, 1, qColSize).getValues()
                );
                /* Set width */
                sheet.setColumnWidth(col + qq + 1, 
                  sheet.getColumnWidth(qStartCol + qq + 1)
                );
            }
        }
    }
}
///////////////////////////////////////////////////////////
function adjustQField(sheet, num, row) {
    const qRequiredRule = SpreadsheetApp.newDataValidation()
              .requireValueInList([trans('required'), trans('not_required')], true)
              .setAllowInvalid(false)
              .build();
    /* Active cells */
    const range = sheet.getRange(row, qStartCol + 1, 1, num * qColSize);
    range.setBackground(null);
    for (let q = 0; q < num; q++) {
        let cells = sheet.getRange(row, qStartCol + q * qColSize + qqRequiredCol + 1);
        cells.setDataValidation(qRequiredRule)
        cells.setValue(trans('required'));
    }
    /* Non-active cells */
    const col = qStartCol + num * qColSize;
    const cells = sheet.getRange(row, col + 1, 1, sheet.getLastColumn() - col + 1);
    cells.setBackground('gray');
}

///////////////////////////////////////////////////////////
/** Expand response fields */
function expandRField(sheet, num) {
    /* Get current qnum */
    const cur = (sheet.getLastColumn() - qStartCol) / qColSize;
    /* Add columns */
    if (cur < num) {
        sheet.insertColumnsAfter(sheet.getLastColumn(), (num - cur) * qColSize);
        /* Adjust */
        for (let q = cur; q < num; q++) {
            const col = qStartCol + q * qColSize;
            /* Set header */
            sheet.getRange(headerRow + 1, col + 1).setValue(`Q.${q+1}`);
            for (let qq = 0; qq < qColSize; qq++) {
                /* Set subheader */
                sheet.getRange(subHeaderRow+1, col + 1, 1, qColSize).setValues(
                  sheet.getRange(subHeaderRow+1, qStartCol + 1, 1, qColSize).getValues()
                );
                /* Set width */
                sheet.setColumnWidth(col + qq + 1, 
                  sheet.getColumnWidth(qStartCol + qq + 1)
                );
            }
        }
    }
}
///////////////////////////////////////////////////////////
function addToInfoSheet(sheetId, data) {
  const spreadsheet = SpreadsheetApp.openById(sheetId);
  const sheet = spreadsheet.getSheetByName('_info_') ? 
                spreadsheet.getSheetByName('_info_') : 
                spreadsheet.insertSheet('_info_', 0);
  if (0 < data.length) 
    sheet.getRange(sheet.getLastRow() + 1, 1, data.length, data[0].length).setValues(data);
}

///////////////////////////////////////////////////////////
function updatePrefSheet(sheetId, items) {
  const spreadsheet = SpreadsheetApp.openById(sheetId);
  const sheet = spreadsheet.getSheetByName('_pref_') ? 
                spreadsheet.getSheetByName('_pref_') : 
                spreadsheet.insertSheet('_pref_', 1);
  sheet.clear();
  sheet.deleteColumn(valCol+1);
  const pref = getPref();
  for (const item of items) {
    if (pref[item].type === 'group') {
      sheet.appendRow([`#${item}`, '']);
      for (const key of pref[item].content) {
        const elem = pref[key];
        if (elem.type === 'num') sheet.appendRow([key, Number(elem.value)]);
        else if (elem.type === 'date') sheet.appendRow([key, new Date(elem.value)]);
        else if (elem.type === 'checkbox') {
          sheet.appendRow([key, elem.value]);
          sheet.getRange(sheet.getLastRow(), valCol + 1).setDataValidation(checkRule);
        }
        else if (elem.type === 'select') {
          sheet.appendRow([key, elem.value]);
          const rule = SpreadsheetApp.newDataValidation()
                        .requireValueInList(elem.options, true)
                        .setAllowInvalid(false)
                        .build();
          sheet.getRange(sheet.getLastRow(), valCol + 1).setDataValidation(rule);
        }
        else sheet.appendRow([key, elem.value]);
      }
    }
    else sheet.appendRow([item, pref[item].value]);
  }
}

///////////////////////////////////////////////////////////
/** Call onEdit */
function onEdit(e) {
  /* Get cell */
  const range = e.range;
  const sheet = range.getSheet();
  const pref = getPref();
      
  if (sheet.getName() === trans('tab.question')) {
    if (qStartRow + pref.max_task.value < range.getRow()) {
      showAlert(trans('msg.over_task'));
      return;
    }
    else if (qStartRow < range.getRow()) {
      const col = range.getColumn();
      /* Check the number of questions */
      if(col == (qNumCol + 1)) {
        /*  */
        if (sheet.getRange(range.getRow(), qIdCol).getValue() != '') {
          showAlert(trans('msg.not_qnum_change'));
          
          
        }
        /*  */
        if (!/^[0-9]+$/.test(range.getValue())) {
          const n = range.getValue().replace(/[０-９]/g, function(s) {
            return String.fromCharCode(s.charCodeAt(0) - 0xFEE0);
          });
          if (/^[0-9]+$/.test(n)) range.setValue(Number(n));
          else {
            showAlert(trans('msg.not_num_form'));
            range.setValue(1);
          }
        }
        /*  */
        const num = range.getValue();
        if (pref.max_question.value < range.getValue()) {
          showAlert(trans('msg.over_qnum').replace('%NUM%', `${range.getValue()}`));
          range.setValue(pref.max_question.value);
        }
        /*  */
        /* Adjust field */
        adjustQField(sheet, num, range.getRow());
      }
    }
  }
}
/*********************************************************/

/*********************************************************/
