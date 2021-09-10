import {
  importCourses,
  importCourseTeachers,
  importCourseStudents,
  importCourseWorks,
  importStudentSubmissions,
  updateCourses,
  updateCourseStudents,
  updateCourseTeachers, updateCourseWorks
} from './classroom';

import {
  InvalidSheetSchemaError,
  NoDefinedCourseSheetError,
  NoSelectedCourseError,
  NoSelectedCourseWorkError
} from './classroom/selectors';

function handleErrors(err){
  if(err instanceof NoDefinedCourseSheetError) {
    Browser.msgBox("エラー：「courses」シートがありません。「1.コース一覧(シート名：courses)を抽出」を実行してから、こちらを再実行してください。");
  }else if(err instanceof InvalidSheetSchemaError){
    Browser.msgBox("エラー：シート内容が不正です。");
  }else if(err instanceof NoSelectedCourseError) {
    Browser.msgBox("エラー：「courses」シートで、対象コースの行を、いずれか1行だけ選択状態にしてから、再実行してください。");
  }else if(err instanceof NoSelectedCourseWorkError){
    Browser.msgBox("エラー：「課題一覧(courseworks:コース名)」において、課題の行を、いずれか1行だけ選択状態にしてから、再実行してください。");
  }else{
    Browser.msgBox(err);
  }
}

global._importCourses = function (){
  importCourses([
    "courseId",
    "コース名",
    "セクション",
    "説明",
    "部屋",
    "オーナー教員名",
    "オーナー教員Email",
    "作成日",
    "更新日",
    "クラスコード",
    "状態",
    "代替リンク",
    "教師グループEmail",
    "コースグループEmail",
    "教師フォルダId",
    "教師フォルダ代替リンク",
    "保護者機能有効化",
    "カレンダーID",
    "生徒数",
  ]);
}
global._importCourseTeachers = function(){
  try {
    importCourseTeachers([
      "courseId",
      "コース名",
      "メールアドレス",
      "氏名",
      "写真URL",
    ]);
  } catch (err) {
    handleErrors(err);
  }
}
global._importCourseStudents = function(){
  try {
    importCourseStudents([
      "courseId",
      "コース名",
      "メールアドレス",
      "氏名",
      "写真URL",
    ]);
  } catch (err) {
    handleErrors(err);
  }
}

global._importCourseWorks = function(){
  try {
    importCourseWorks([
      "courseId",
      "コース名",
      "courseWorkId",
      "課題タイトル",
      "説明",
      "状態",
      "代替リンク",
      "作成日",
      "更新日",
      "期限日",
      "期限時刻",
      "スケジュールされた時刻",
      "配点",
      "種類",
      "開発者モード",
      "割当",
      "個別割当対象者リスト",
      "作成者ID",
      "トピックID",
      "提出物フォルダ",
      "選択肢",
    ]);
  } catch (err) {
    handleErrors(err);
  }
};

global._importStudentSubmissions = function () {
  try {
    importStudentSubmissions([
      "courseId",
      "コース名",
      "courseWorkId",
      "課題タイトル",
      "氏名",
      "メールアドレス",
      "写真URL",
      "状態",
      "作成日",
      "更新日",
    ]);
  } catch (err) {
    handleErrors(err);
  }
}

global._updateCourses = updateCourses;
global._updateCourseTeachers = updateCourseTeachers;
global._updateCourseStudents = updateCourseStudents;
global._updateCourseWorks = updateCourseWorks;

export const importClassroomSubMenu = function(ui) {
  return ui.createMenu("A. Classroomからの抽出")
    .addItem(
      "1. コース一覧（courses:教員名）を抽出/抽出内容を更新",
      "_importCourses"
    )
    .addSeparator()
    .addItem(
      "2. 生徒一覧（students:コース名）を抽出/抽出内容を更新",
      "_importCourseStudents"
    )
    .addItem(
      "3. 教員一覧（teachers:コース名）を抽出/抽出内容を更新",
      "_importCourseTeachers"
    )
    .addItem(
      "4. 課題一覧（courseworks:コース名）を抽出/抽出内容を更新",
      "_importCourseWorks"
    )
    .addSeparator()
    .addItem(
      "5. 提出物一覧（submissions:コース名 課題名）を抽出/抽出内容を更新",
      "_importStudentSubmissions"
    );
}
export const exportClassroomSubMenu = function(ui) {
  return ui.createMenu("B. Classroomの作成/更新")
    .addItem("1. コース一覧（courses:教員名）からの作成/更新",
      "_updateCourses")
    .addSeparator()
    .addItem(
      "2. 生徒一覧（students:コース名）からの追加/削除",
      "_updateCourseStudents"
    )
    .addItem(
      "3. 教員一覧（teachers:コース名）からの追加/削除",
      "_updateCourseTeachers"
    )
    .addItem(
      "4. 課題一覧（courseworks:コース名）からの追加/削除/更新",
      "_updateCourseWorks"
    );
}
