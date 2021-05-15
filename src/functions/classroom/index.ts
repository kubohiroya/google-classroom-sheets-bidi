import {
  importCourses,
  importCourseStudents,
  importCourseTeachers,
  importCourseWorks,
  importStudentSubmissions,
} from "./importClassroom";
import {
  updateCourses,
  updateCourseStudents,
  updateCourseTeachers,
  updateCourseWorks,
} from "./updateClassroom";

global.importCourses = importCourses;
global.importCourseTeachers = importCourseTeachers;
global.importCourseStudents = importCourseStudents;
global.importCourseWorks = importCourseWorks;
global.importStudentSubmissions = importStudentSubmissions;

global.updateCourses = updateCourses;
global.updateCourseTeachers = updateCourseTeachers;
global.updateCourseStudents = updateCourseStudents;
global.updateCourseWorks = updateCourseWorks;
