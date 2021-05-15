import { getSchema } from "../sheetUtil";
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

export const getSelectedCourse = (): CourseData => {
  const getSelectedCourseOnCoursesSheet = (sheet: Sheet) => {
    const activeRange = sheet.getActiveRange();
    const values =
      sheet.getName() === "courses" && activeRange != null
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
  };

  const getSelectedCourseOnSheet = (
    sheet: Sheet,
    schema: string,
    caption: string
  ) => {
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
  };

  const getSelectedCourseOnTeachersSheet = (sheet: Sheet) => {
    return getSelectedCourseOnSheet(
      sheet,
      "teachers",
      "2.教師一覧(シート名：teachers:コース名)を抽出"
    );
  };

  const getSelectedCourseOnStudentsSheet = (sheet: Sheet) => {
    return getSelectedCourseOnSheet(
      sheet,
      "students",
      "3.生徒一覧(シート名：students:コース名)を抽出"
    );
  };

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
};

export const getSelectedCourseWorks = (): CourseWorkData => {
  const getSelectedCourseWorkOfSelectedRow = (activeRange: Range) => {
    const values = activeRange.getValues();
    const courseId = values[0][0] as string;
    const courseName = values[0][1] as string;
    const courseWorkId = values[0][2] as string;
    const courseWorkTitle = values[0][3] as string;
    return { courseId, courseName, courseWorkId, courseWorkTitle };
  };

  const getSelectedCourseWorkOfSelectedSheet = (sheet: Sheet) => {
    const values = sheet.getRange(2, 1, 1, 4).getValues();
    const courseId = values[0][0] as string;
    const courseName = values[0][1] as string;
    const courseWorkId = values[0][2] as string;
    const courseWorkTitle = values[0][3] as string;
    return { courseId, courseName, courseWorkId, courseWorkTitle };
  };

  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const activeSheet = spreadsheet.getActiveSheet();
  const activeRange = activeSheet.getActiveRange();
  switch (getSchema(activeSheet.getName())) {
    case "courseworks":
      if (!activeRange || activeRange.getRowIndex() == 0) {
        throw new Error(
          "エラー：選択中のシート「課題一覧(courseworks:コース名)」において、課題の行を、いずれか1行だけ選択状態にしてから、再実行してください。"
        );
      }
      return getSelectedCourseWorkOfSelectedRow(activeRange);
    case "submissions":
      return getSelectedCourseWorkOfSelectedSheet(activeSheet);
    case "courses":
    case "teachers":
    case "students":
    default:
      throw new Error(
        "エラー：選択中のシート「課題一覧(courseworks:コース名)」において、課題の行を、いずれか1行だけ選択状態にしてから、再実行してください。"
      );
  }
};
