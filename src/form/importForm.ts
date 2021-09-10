import {formToSheet} from './formToSheet';

export function importForm(): void {
  const sheet = SpreadsheetApp.getActiveSheet();

  const idRows = sheet
    .getRange(1, 1, sheet.getLastRow(), 2)
    .getValues()
    .filter(function (row) {
      return row[0] === "id" && row[1] !== "";
    });
  if (idRows.length === 0) {
    throw "`forms id` row is not defined.";
  }
  const id = idRows[0][1];
  const form = FormApp.openById(id);
  formToSheet(form, sheet);
}

