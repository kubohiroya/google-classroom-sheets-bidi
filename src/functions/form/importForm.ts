import formToJson from "./formToJson";
import jsonToSheet from "./jsonToSheet";
import { messages } from "./messages";
type Form = GoogleAppsScript.Forms.Form;
type Sheet = GoogleAppsScript.Spreadsheet.Sheet;
import { getCreatedAtUpdatedAtValues } from "../driveFileUtil";
import { startFormPicker } from "../execute";

const message = messages[Session.getActiveUserLocale()]["ui"];

export function importFormWithPicker(): void {
  startFormPicker();
}

export function importFormWithDialog(): void {
  const inputBoxTitle = message["import form"];

  function importFormDialog(): Form {
    const input = Browser.inputBox(
      inputBoxTitle,
      "input source form URL",
      Browser.Buttons.OK_CANCEL
    );
    if (input === "cancel") {
      throw message["form import canceled"];
    }
    let form = null;
    if (input.endsWith("/edit")) {
      form = FormApp.openByUrl(input);
    }
    if (!form) {
      Browser.msgBox(message["invalid form URL"] + ": " + input);
      return importFormDialog();
    }
    return form;
  }

  const form = importFormDialog();
  const { createdAt, updatedAt } = getCreatedAtUpdatedAtValues(form.getId());
  formToSheet(form, createdAt, updatedAt, null);
}

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
  const { createdAt, updatedAt } = getCreatedAtUpdatedAtValues(id);
  formToSheet(form, createdAt, updatedAt, sheet);
}

export function formToSheet(
  form: Form,
  createdAt: number,
  updatedAt: number,
  sheet: Sheet | null
): void {
  try {
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
  } catch (exception) {
    Logger.log(exception);
    if (exception.stack) {
      Logger.log(exception.stack);
    }
    Browser.msgBox(
      message["form import failed."] +
        "\\n" +
        JSON.stringify(exception, null, " ")
    );
  }
}
