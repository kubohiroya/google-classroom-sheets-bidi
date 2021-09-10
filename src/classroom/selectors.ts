import { getSchema } from "../form/sheetUtil";
import RangeList = GoogleAppsScript.Spreadsheet.RangeList;
type Sheet = GoogleAppsScript.Spreadsheet.Sheet;
type Range = GoogleAppsScript.Spreadsheet.Range;

type CourseData = {
  courseId: string;
  courseName: string;
};
type CourseWorkData = {
  courseWorkId: string;
  courseWorkTitle: string;
} & CourseData;

export class NoSelectedCourseError extends Error{}
export class NoSelectedCourseWorkError extends Error{}
export class InvalidSheetSchemaError extends Error{}
export class NoDefinedCourseSheetError extends Error{}

export function getSelectedCourse(): CourseData {

  function getSelectedCourseOnCoursesSheet (sheet: Sheet) {
    const activeRange = sheet.getActiveRange();
    const values =
      getSchema(sheet.getName()) === "courses" && activeRange != null
        ? activeRange.getValues()
        : sheet.getRange(2, 1, 1, 2).getValues();
    if (values.length != 1 && values[0].length < 2) {
      throw new NoSelectedCourseError();
    }
    const courseId = values[0][0] as string;
    const courseName = values[0][1] as string;

    if (
      courseId === null ||
      courseName === null ||
      courseId === "" ||
      courseName === ""
    ) {
      throw new InvalidSheetSchemaError();
    }
    return { courseId, courseName };
  }

  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const activeSheet = spreadsheet.getActiveSheet();
  const sheet = spreadsheet.getSheetByName("courses");

  if (!sheet) {
    throw new NoDefinedCourseSheetError();
  }

  switch (getSchema(activeSheet.getName())) {
    case "courses":
    case "teachers":
    case "students":
    case "courseworks":
      return getSelectedCourseOnCoursesSheet(activeSheet);
    default:
      throw new NoSelectedCourseError();
  }
}

export function getSelectedCourseWorksList(): Array<CourseWorkData> {

  function getSelectedCourseWorkList(values: Array<Array<string>>) {
    return Array.from(new Map<string, CourseWorkData>(values.map((row: string[])=> [row[2], {
      courseId: row[0],
      courseName: row[1],
      courseWorkId: row[2],
      courseWorkTitle: row[3],
    }])).values());
  }

  function getSelectedCourseWorkListOfSelectedRows (activeRangeList: RangeList): Array<CourseWorkData> {
    const values = activeRangeList.getRanges().map(activeRange=>activeRange.getValues()).flat();
    return getSelectedCourseWorkList(values);
  }

  function getSelectedCourseWorkListOfSelectedSheet (sheet: Sheet): Array<CourseWorkData> {
    const values = sheet.getRange(2, 1, sheet.getMaxRows() - 1, 4).getValues();
    return getSelectedCourseWorkList(values);
  }

  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const activeSheet = spreadsheet.getActiveSheet();
  const activeRangeList = activeSheet.getActiveRangeList();
  switch (getSchema(activeSheet.getName())) {
    case "courseworks":
      if (!activeRangeList || activeRangeList.getRanges().length == 0) {
        throw new NoSelectedCourseWorkError();

      }
      return getSelectedCourseWorkListOfSelectedRows(activeRangeList);
    case "submissions":
      return getSelectedCourseWorkListOfSelectedSheet(activeSheet);
    case "courses":
    case "teachers":
    case "students":
    default:
      throw new NoSelectedCourseWorkError()
  }
}
