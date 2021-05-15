export function onOpen(): void {
  const ui = SpreadsheetApp.getUi();

  const importClassroomSubMenu = ui
    .createMenu("A. Classroomからの抽出")
    .addItem(
      "1. コース一覧（courses:教員名）を抽出/抽出内容を更新",
      "importCourses"
    )
    .addSeparator()
    .addItem(
      "2. 生徒一覧（students:コース名）を抽出/抽出内容を更新",
      "importCourseStudents"
    )
    .addItem(
      "3. 教員一覧（teachers:コース名）を抽出/抽出内容を更新",
      "importCourseTeachers"
    )
    .addItem(
      "4. 課題一覧（courseworks:コース名）を抽出/抽出内容を更新",
      "importCourseWorks"
    )
    .addSeparator()
    .addItem(
      "5. 提出物一覧（submissions:コース名 課題名）を抽出/抽出内容を更新",
      "importStudentSubmissions"
    );

  const exportClassroomSubMenu = ui
    .createMenu("B. Classroomの作成/更新")
    .addItem("1. コース一覧（courses:教員名）からの作成/更新", "updateCourses")
    .addSeparator()
    .addItem(
      "2. 生徒一覧（students:コース名）からの追加/削除",
      "updateCourseStudents"
    )
    .addItem(
      "3. 教員一覧（teachers:コース名）からの追加/削除",
      "updateCourseTeachers"
    )
    .addItem(
      "4. 課題一覧（courseworks:コース名）からの追加/削除/更新",
      "updateCourseWorks"
    );

  const importFormSubMenu = ui
    .createMenu("C. フォームからの抽出")
    .addItem(
      "1. フォーム定義（forms:フォーム名）を抽出....",
      "importFormWithPicker"
    )
    .addItem(
      "2. フォーム定義（forms:フォーム名）の抽出内容を更新",
      "importForm"
    );

  const exportFormSubMenu = ui
    .createMenu("D. フォームの作成/更新")
    .addItem(
      "1. フォーム定義（forms:フォーム名）をもとに新規作成",
      "exportFormWithDialog"
    )
    .addItem("2. フォーム定義（forms:フォーム名）をもとに更新", "updateForm")
    .addSeparator()
    .addItem(
      "3. フォーム定義（forms:フォーム名）をもとにプレビュー",
      "previewForm"
    );


  ui.createMenu("相互評価")
    .addSubMenu(importClassroomSubMenu)
    .addSubMenu(exportClassroomSubMenu)
    .addSeparator()
    .addSubMenu(importFormSubMenu)
    .addSubMenu(exportFormSubMenu)
    .addToUi();
}


