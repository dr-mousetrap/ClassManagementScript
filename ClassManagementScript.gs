function onEdit(e) {
  const sheet = e.source.getSheetByName('Sheet2');
  const editedCell = e.range;
  const row = editedCell.getRow();
  const col = editedCell.getColumn();
  const value = editedCell.getValue();

  if (!value || !sheet) return;

  const dropdownOptions = ['Not Attempted', 'Ungraded', 'Completed', 'Failed'];

  if (value.toLowerCase() === 'z') {
    editedCell.setValue('');
    generateDropdown(sheet, row, col, 1);
    changeColours();
    return;
  }

  const bgColor = editedCell.getBackground();
  if (bgColor.toLowerCase() === '#9900ff') { // purple
    editedCell.clearDataValidations();
    editedCell.setValue('');
    editedCell.setBackground('#ffffff');
    return;
  }

  if (col === 3 && row >= 3) {
    const numClasses = parseInt(value);
    const startColumn = 5; // Column E
    const maxColumns = 26; // Up to column Z

    const rangeToClear = sheet.getRange(row, startColumn, 1, maxColumns - startColumn + 1);
    rangeToClear.clearDataValidations();
    rangeToClear.setBackground(null);

    if (!isNaN(numClasses) && numClasses > 0) {
      for (let i = 0; i < numClasses; i++) {
        generateDropdown(sheet, row, startColumn + i, 1);
      }
      changeColours();
    }
    return;
  }

  if (value.toLowerCase() === 'x') {
    const colLetter = String.fromCharCode(64 + col);
    const startRow = 3;
    let dropdownCount = 0;
    let firstRow = null;

    for (let r = row - 1; r >= startRow; r--) {
      const cell = sheet.getRange(r, col);
      const rule = cell.getDataValidation();

      if (
        rule &&
        rule.getCriteriaType() === SpreadsheetApp.DataValidationCriteria.VALUE_IN_LIST &&
        rule.getCriteriaValues()[0].join() === dropdownOptions.join()
      ) {
        dropdownCount++;
        firstRow = r;
      } else {
        break;
      }
    }

    if (dropdownCount > 0) {
      const lastRow = row - 1;
      const rangeFormula = `${colLetter}${firstRow}:${colLetter}${lastRow}`;
      const formula = `=IF(COUNTIF(${rangeFormula}, "Completed") = ${dropdownCount}, "Completed", IF(COUNTIF(${rangeFormula}, "Completed") >= 1, "Ungraded", "Not Attempted"))`;
      editedCell.setFormula(formula);
    } else {
      editedCell.setValue('Not Attempted');
    }
    changeColours();
    return;
  }

  if (/^C\d+$/i.test(value)) {
    const dropdownCount = parseInt(value.substring(1), 10);
    const colLetter = String.fromCharCode(64 + col);

    const assignmentName = "Assignment Name";
    const firstRowCell = sheet.getRange(row, col); // Same column as C#
    firstRowCell.setValue(assignmentName); // Set as plain text

    generateDropdown(sheet, row + 1, col, dropdownCount);

    const xRow = row + dropdownCount + 1;
    const rangeFormula = `${colLetter}${row + 1}:${colLetter}${xRow - 1}`;
    const formula = `=IF(COUNTIF(${rangeFormula}, "Completed") = ${dropdownCount}, "Completed", IF(COUNTIF(${rangeFormula}, "Completed") >= 1, "Ungraded", "Not Attempted"))`;
    const formulaCell = sheet.getRange(xRow, col);
    formulaCell.setFormula(formula);
    return;
  }

  if (/^V\d+$/i.test(value)) {
    const clearCount = parseInt(value.substring(1), 10);
    const startRow = row + 1;

    for (let i = 0; i < clearCount; i++) {
      const cell = sheet.getRange(startRow + i, col);
      cell.clearContent();
      cell.clearDataValidations();
      cell.setBackground(null);
    }

    editedCell.setValue('');
    return;
  }

  changeColours();
}

function generateDropdown(sheet, startRow, col, count) {
  const dropdownOptions = ['Not Attempted', 'Ungraded', 'Completed', 'Failed'];
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(dropdownOptions)
    .setAllowInvalid(false)
    .build();

  for (let i = 0; i < count; i++) {
    const cell = sheet.getRange(startRow + i, col);
    cell.setDataValidation(rule);
    if (cell.getValue() === '') cell.setValue('Not Attempted');
  }
}

function changeColours() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Sheet2');
  const startRow = 3;
  const startCol = 5;
  const endCol = 26;
  const lastRow = sheet.getLastRow();

  if (lastRow < startRow) return;

  const dataRange = sheet.getRange(startRow, startCol, lastRow - startRow + 1, endCol - startCol + 1);
  const values = dataRange.getValues();
  let backgrounds = [];

  for (let r = 0; r < values.length; r++) {
    let rowColors = [];
    for (let c = 0; c < values[r].length; c++) {
      const val = values[r][c];
      let color = '';
      switch (val) {
        case 'Not Attempted': color = '#FF7992'; break;
        case 'Ungraded': color = '#8C8E8D'; break;
        case 'Completed': color = '#85fd47'; break;
        case 'Failed': color = '#FF0000'; break;
        default: color = null;
      }
      rowColors.push(color);
    }
    backgrounds.push(rowColors);
  }

  dataRange.setBackgrounds(backgrounds);
}
