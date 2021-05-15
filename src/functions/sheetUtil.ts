type Spreadsheet = GoogleAppsScript.Spreadsheet.Spreadsheet;
type Sheet = GoogleAppsScript.Spreadsheet.Sheet;

export const createOrSelectSheetBySheetName = (
  name: string,
  tabColor: string
): Sheet => {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = spreadsheet.getSheetByName(name);
  if (!sheet) {
    sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet(name);
  }
  SpreadsheetApp.getActiveSpreadsheet().setActiveSheet(sheet);
  sheet.setTabColor(tabColor);
  return sheet;
};

export const getSchema = (sheetName: string): string => {
  return sheetName.split(":")[0];
};

export const openSheetByUrl = (url: string): Sheet | null => {
  const ss = SpreadsheetApp.openByUrl(url);
  const gid = parseInt(url.split(/#gid=/)[1]);
  return ss.getSheets().find((sheet) => sheet.getSheetId() === gid) || null;
};

export const isSpreadsheetUrl = (url: string): boolean => {
  return url.indexOf("/spreadsheets/d/") > 0 && url.indexOf("#gid=") > 0;
};

export const createSpreadsheetUrl = (ss: Spreadsheet, sheet: Sheet): string => {
  return `https://docs.google.com/spreadsheets/d/${ss.getId()}/edit#gid=${sheet.getSheetId()}`;
};
