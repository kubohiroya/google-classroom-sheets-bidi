import { getSchema } from "../sheetUtil";
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

export function getSelectedCourse(): CourseData {
  function getSelectedCourseOnCoursesSheet (sheet: Sheet) {
    const activeRange = sheet.getActiveRange();
    const values =
      getSchema(sheet.getName()) === "courses" && activeRange != null
        ? activeRange.getValues()
        : sheet.getRange(2, 1, 1, 2).getValues();
    if (values.length != 1 && values[0].length < 2) {
      throw new Error(
        "エラー：「courses」シートで、対象コースの行を、いずれか1行だけ選択状態にしてから、再実行してください。"
      );
    }
    const courseId = values[0][0] as string;
    const courseName = values[0][1] as string;

    if (courseId && courseName) {
      return { courseId, courseName };
    } else {
      throw new Error(
        "エラー：「courses」シートで、対象コースの行を、いずれか1行だけ選択状態にしてから、再実行してください。"
      );
    }
  }

  function getSelectedCourseOnSheet(
    sheet: Sheet,
    schema: string,
    caption: string
  ) {
    const values = sheet.getRange(2, 1, 1, 2).getValues();
    Logger.log("getSelectedCourseOnSheet:" + JSON.stringify(values));
    const courseId = values[0][0] as string;
    const courseName = values[0][1] as string;
    if (
      courseId === null ||
      courseName === null ||
      courseId === "" ||
      courseName === ""
    ) {
      throw new Error(
        `エラー：「${schema}」シート内容が不正です。「${caption}」を実行してから、こちらを再実行してください。`
      );
    }
    return { courseId, courseName };
  }

  function getSelectedCourseOnTeachersSheet(sheet: Sheet) {
    return getSelectedCourseOnSheet(
      sheet,
      "teachers",
      "2.教師一覧(シート名：teachers:コース名)を抽出"
    );
  }

  function getSelectedCourseOnStudentsSheet (sheet: Sheet) {
    return getSelectedCourseOnSheet(
      sheet,
      "students",
      "3.生徒一覧(シート名：students:コース名)を抽出"
    );
  }

  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const activeSheet = spreadsheet.getActiveSheet();
  const sheet = spreadsheet.getSheetByName("courses");

  switch (getSchema(activeSheet.getName())) {
    case "courses":
      return getSelectedCourseOnCoursesSheet(activeSheet);
    case "teachers":
      return getSelectedCourseOnTeachersSheet(activeSheet);
    case "students":
      return getSelectedCourseOnStudentsSheet(activeSheet);
    case "courseworks":
      Logger.log("courseworks");
      return getSelectedCourseOnCoursesSheet(activeSheet);
    default:
      if (!sheet) {
        throw new Error(
          "エラー：「courses」シートがありません。「1.コース一覧(シート名：courses)を抽出」を実行してから、こちらを再実行してください。"
        );
      } else {
        throw new Error(
          "* エラー：「courses」シートで、対象コースの行を、いずれか1行だけ選択状態にしてから、再実行してください。"
        );
      }
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
        throw new Error(
          "エラー：選択中のシート「課題一覧(courseworks:コース名)」において、課題の行を選択状態にしてから、再実行してください。"
        );
      }
      return getSelectedCourseWorkListOfSelectedRows(activeRangeList);
    case "submissions":
      return getSelectedCourseWorkListOfSelectedSheet(activeSheet);
    case "courses":
    case "teachers":
    case "students":
    default:
      throw new Error(
        "エラー：選択中のシート「課題一覧(courseworks:コース名)」において、課題の行を、いずれか1行だけ選択状態にしてから、再実行してください。"
      );
  }
}
