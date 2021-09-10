import Form = GoogleAppsScript.Forms.Form;
import {
  exportForm,
  importForm,
  updateForm,
  formToSheet
} from './form';

import {messages} from './messages';
import { startFormPicker } from "./execute";
import {openUrl} from './openUrl';

const userLocale = Session.getActiveUserLocale();
const uiMessages = messages[userLocale]["ui"];

export function importFormWithPicker(): void {
  startFormPicker();
}

export function importFormWithDialog(): void {
  const inputBoxTitle = uiMessages["import form"];

  function importFormWithDialog(): Form {
    const input = Browser.inputBox(
      inputBoxTitle,
      "input source form URL",
      Browser.Buttons.OK_CANCEL
    );
    if (input === "cancel") {
      throw uiMessages["form import canceled"];
    }
    let form = null;
    if (input.endsWith("/edit")) {
      form = FormApp.openByUrl(input);
    }
    if (!form) {
      Browser.msgBox(uiMessages["invalid form URL"] + ": " + input);
      return importFormWithDialog();
    }
    return form;
  }

  try {
    const form = importFormWithDialog();
    formToSheet(form, null);
  } catch (exception) {
    Logger.log(exception);
    if (exception.stack) {
      Logger.log(exception.stack);
    }
    Browser.msgBox(
      uiMessages["form import failed."] +
      "\\n" +
      JSON.stringify(exception, null, " ")
    );
  }
}

global._importForm = importForm;
global._importFormWithDialog = importFormWithDialog;
global._importFormWithPicker = importFormWithPicker;
global._updateForm = updateForm;

global._exportFormWithDialog = function(){

  function getFormTitle(): string {
    const inputBoxTitle = uiMessages["update form"];
    const input = Browser.inputBox(
      inputBoxTitle,
      uiMessages["form title"],
      Browser.Buttons.OK_CANCEL
    );
    if (input === "cancel") {
      throw uiMessages["form update canceled"];
    }
    if (input === "") {
      return getFormTitle();
    }
    return input;
  }

  try {
    const url = exportForm(getFormTitle);
    openUrl("Please wait to open the forms page...", url);
    Browser.msgBox(
      uiMessages["form update succeed."] +
      "\\n" +
      "URL: \\n" + url
    );
  } catch (exception) {
    Logger.log(exception);
    if (exception.stack) {
      Logger.log(exception.stack);
    }
    Browser.msgBox(
      uiMessages["form update failed."] +
      "\\n" +
      JSON.stringify(exception, null, " ")
    );
  }
}

global._previewForm = function(): void {
  const title = "フォームをプレビュー...";
  const html = HtmlService.createTemplateFromFile("form");
  html.gid = SpreadsheetApp.getActiveSheet().getSheetId();
  SpreadsheetApp.getUi().showModelessDialog(
    html.evaluate().setWidth(1055).setHeight(654).setTitle(title),
    title
  );
}

export const importFormSubMenu = function(ui) {
  return ui.createMenu("C. フォームからの抽出")
    .addItem(
      "1. フォーム定義（forms:フォーム名）を抽出....",
      "_importFormWithPicker"
    )
    .addItem(
      "2. フォーム定義（forms:フォーム名）の抽出内容を更新",
      "_importForm"
    );
}

export const exportFormSubMenu = function(ui) {
  return ui.createMenu("D. フォームの作成/更新")
    .addItem(
      "1. フォーム定義（forms:フォーム名）をもとに新規作成",
      "_exportFormWithDialog"
    )
    .addItem("2. フォーム定義（forms:フォーム名）をもとに更新", "_updateForm")
    .addSeparator()
    .addItem(
      "3. フォーム定義（forms:フォーム名）をもとにプレビュー",
      "_previewForm"
    );
}

