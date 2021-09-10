import sheetToJson from "./sheetToJson";
import { jsonToForm } from "./jsonToForm";
import { getCreatedAtUpdatedAtValues } from "./driveFileUtil";

export function exportForm(getFormTitle: ()=>string): string {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const { createdAt, updatedAt } = getCreatedAtUpdatedAtValues(ss.getId());
  const sheet = ss.getActiveSheet();

  const json = sheetToJson(sheet, createdAt, updatedAt);

  const title = json.metadata.title && json.metadata.title !== "" ? json.metadata.title : getFormTitle();

  const form = FormApp.create(title);
  jsonToForm(json, form);

  const file = DriveApp.getFileById(form.getId());
  form.setTitle(title);
  file.setName(title);

  return form.shortenFormUrl(form.getPublishedUrl())
}

export function updateForm(): string {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const { createdAt, updatedAt } = getCreatedAtUpdatedAtValues(ss.getId());
  const sheet = ss.getActiveSheet();
  const json = sheetToJson(sheet, createdAt, updatedAt);

  const form = FormApp.openById(json.metadata.id);
  for (let index = form.getItems().length - 1; index >= 0; index--) {
    form.deleteItem(index);
  }
  jsonToForm(json, form);

  const file = DriveApp.getFileById(form.getId());
  file.setName(form.getTitle());

  return form.shortenFormUrl(form.getPublishedUrl())
}
