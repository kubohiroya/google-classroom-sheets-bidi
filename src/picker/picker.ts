import {execute} from '../execute';

export function showPicker(
  title: string,
  mimeTypes: string[],
  command: string[],
  cursor: number
): void {
  const documentProps = PropertiesService.getDocumentProperties();
  let apiKey = documentProps.getProperty("PICKER_API_KEY");
  while(! apiKey){
    apiKey = Browser.inputBox("PICKER_API_KEY");
    documentProps.setProperty("PICKER_API_KEY", apiKey);
  }

  const picker = HtmlService.createTemplateFromFile("picker");
  picker.apiKey = apiKey;
  picker.mimeType = mimeTypes.join(",");
  picker.command = command.join(",");
  picker.cursor = cursor.toString();
  SpreadsheetApp.getUi().showModalDialog(
    picker.evaluate().setTitle(title).setWidth(1055).setHeight(654),
    title
  );
}

//Access Tokenを取得する
export function getOAuthToken(): string {
  DriveApp.getRootFolder();
  return ScriptApp.getOAuthToken();
}

//エラーメッセージの表示
export function showAlert(msg: string): void {
  const ui = SpreadsheetApp.getUi();
  ui.alert(msg);
}

export function pickerHandler(ev: {
  command: string;
  cursor: string;
  url: string;
}): void {
  Logger.log(JSON.stringify(ev, null, " "));
  const command = ev.command.split(",");
  const cursor = parseInt(ev.cursor);
  const url = ev.url;

  if (cursor < command.length - 1) {
    execute(command, cursor + 1, { url });
  }

}
