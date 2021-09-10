type Form = GoogleAppsScript.Forms.Form;
type Sheet = GoogleAppsScript.Spreadsheet.Sheet;

import {getCreatedAtUpdatedAtValues} from './driveFileUtil';
import formToJson from './formToJson';
import jsonToSheet from './jsonToSheet';

export function formToSheet(
  form: Form,
  sheet: Sheet | null
): void {

  const {createdAt, updatedAt} = getCreatedAtUpdatedAtValues(form.getId());
  const json = formToJson(form, createdAt, updatedAt);

  if (!sheet) {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const title = "forms:" + json.metadata.title;
    sheet = spreadsheet.getSheetByName(title);
    if (!sheet) {
      sheet = spreadsheet.insertSheet();
      sheet.setName(title);
    } else {
      sheet.clear();
    }
    sheet.setTabColor("purple");
    spreadsheet.setActiveSheet(sheet);
  }

  jsonToSheet(json, sheet);
}
