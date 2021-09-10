import { getSelectedCourse, getSelectedCourseWorksList } from "./selectors";
import { createOrSelectSheetBySheetName } from "../form/sheetUtil";
type ListStudentsResponse = GoogleAppsScript.Classroom.Schema.ListStudentsResponse;
type Student = GoogleAppsScript.Classroom.Schema.Student;
type Teacher = GoogleAppsScript.Classroom.Schema.Teacher;
type UserProfile = GoogleAppsScript.Classroom.Schema.UserProfile;
type ListTeachersResponse = GoogleAppsScript.Classroom.Schema.ListTeachersResponse;
type ListCourseWorkResponse = GoogleAppsScript.Classroom.Schema.ListCourseWorkResponse;
type ListStudentSubmissionsResponse = GoogleAppsScript.Classroom.Schema.ListStudentSubmissionsResponse;
type StudentSubmission = GoogleAppsScript.Classroom.Schema.StudentSubmission;
type Sheet = GoogleAppsScript.Spreadsheet.Sheet;
type Course = GoogleAppsScript.Classroom.Schema.Course;
type CourseWork = GoogleAppsScript.Classroom.Schema.CourseWork;
import { formatDateTime } from "./dateTimeUtils";

const teacherProfiles = new Map<string, UserProfile>();
const studentProfiles = new Map<string, UserProfile>();

function listCourses(teacherId: string): Course[] {
  if (!Classroom.Courses) {
    throw new Error("Classroom.Courses is null");
  }
  let nextPageToken: string | undefined = "";
  const courses: Course[] = [];
  do {
    const optionalArgs = {
      pageSize: 100,
      teacherId: teacherId,
      courseStates: "ACTIVE",
      pageToken: "",
    };
    if (nextPageToken) {
      optionalArgs.pageToken = nextPageToken;
    }
    const coursesResponse = Classroom.Courses.list(optionalArgs);
    nextPageToken = coursesResponse.nextPageToken;
    if (coursesResponse.courses) {
      coursesResponse.courses.forEach(function (course) {
        courses.push(course);
        // course.id && CourseMap.set(course.id, course);
      });
    }
  } while (nextPageToken != undefined);
  return courses;
}

export function listStudentProfiles(courseId: string): Student[] {
  if (!Classroom.Courses || !Classroom.Courses.Students) {
    throw new Error("Classroom.Courses is null");
  }

  let students: Student[] = [];
  let nextPageToken = null;
  do {
    const optionalArgs = {
      pageSize: 64,
      pageToken: "",
    };
    if (nextPageToken) {
      optionalArgs.pageToken = nextPageToken;
    }
    const listStudentsResponse: ListStudentsResponse = Classroom.Courses.Students.list(
      courseId,
      optionalArgs
    );
    nextPageToken = listStudentsResponse.nextPageToken;
    if (listStudentsResponse.students) {
      students = students.concat(listStudentsResponse.students);
    }
  } while (nextPageToken != undefined);
  return students;
}

export function getTeacherProfile(
  courseId: string,
  teacherId: string
): UserProfile {
  if (!Classroom.Courses || !Classroom.Courses.Teachers) {
    throw new Error("Classroom.Courses is null");
  }
  const profile = teacherProfiles.get(teacherId);
  if (!profile) {
    const teacher = Classroom.Courses.Teachers.get(courseId, teacherId);
    if (teacher.profile) {
      teacherProfiles.set(teacherId, teacher.profile);
      return teacher.profile;
    } else {
      return {
        name: { fullName: "(" + teacherId + ")" },
        emailAddress: teacherId,
      };
    }
  } else {
    return profile;
  }
}

export function listTeacherProfiles(courseId: string): Teacher[] {
  if (!Classroom.Courses || !Classroom.Courses.Teachers) {
    throw new Error("Classroom.Courses is null");
  }
  let teachers: Teacher[] = [];
  let nextPageToken = null;
  do {
    const optionalArgs = {
      pageSize: 64,
      pageToken: "",
    };
    if (nextPageToken) {
      optionalArgs.pageToken = nextPageToken;
    }
    const teacherList: ListTeachersResponse = Classroom.Courses.Teachers.list(
      courseId,
      optionalArgs
    );
    nextPageToken = teacherList.nextPageToken;
    if (teacherList.teachers) {
      teachers = teachers.concat(teacherList.teachers);
    }
  } while (nextPageToken != undefined);
  return teachers;
}

export function getStudentProfile(
  courseId: string,
  studentId: string
): UserProfile {
  if (!Classroom.Courses || !Classroom.Courses.Students) {
    throw new Error("Classroom.Courses is null");
  }
  const profile = studentProfiles.get(studentId);
  if (!profile) {
    const student = Classroom.Courses.Students.get(courseId, studentId);
    if (student.profile) {
      studentProfiles.set(studentId, student.profile);
      return student.profile;
    } else {
      return {
        name: { fullName: "(" + studentId + ")" },
        emailAddress: studentId,
      };
    }
  } else {
    return profile;
  }
}

function getNumStudents(courseId: string): number {
  if (!Classroom.Courses || !Classroom.Courses.Students) {
    throw new Error("Classroom.Courses is null");
  }
  const students = Classroom.Courses.Students.list(courseId);
  if (students.students) {
    return students.students.length;
  } else {
    return 0;
  }
}

function listCourseWorks(courseId: string): CourseWork[] {
  if (!Classroom.Courses || !Classroom.Courses.CourseWork) {
    throw new Error("Classroom.Courses is null");
  }
  let nextPageToken: string | undefined = undefined;
  const courseWorks: CourseWork[] = [];
  do {
    const optionalArgs = {
      pageSize: 32,
      pageToken: "",
    };
    if (nextPageToken) {
      optionalArgs.pageToken = nextPageToken;
    }
    const courseWorksResponse: ListCourseWorkResponse = Classroom.Courses.CourseWork.list(
      courseId,
      optionalArgs
    );
    nextPageToken = courseWorksResponse.nextPageToken;
    if (!courseWorksResponse.courseWork) {
      throw new Error("NoCourseWorks");
    }
    courseWorksResponse.courseWork.forEach(function (courseWork) {
      if (courseWork && courseWork.id && courseWork.courseId) {
        // CourseMap.set(courseWork.courseId, { name: courseName });
        courseWorks.push(courseWork);
        // CourseWorkMap.set(courseWork.id, courseWork);
      }
    });
  } while (nextPageToken != undefined);
  return courseWorks;
}

export function listStudentSubmissions(
  courseId: string,
  courseName: string,
  courseWorkId: string
): StudentSubmission[] {
  if (
    !Classroom.Courses ||
    !Classroom.Courses.CourseWork ||
    !Classroom.Courses.CourseWork.StudentSubmissions
  ) {
    throw new Error("Classroom.Courses is null");
  }
  let nextPageToken: string | undefined = "";
  const submissions: StudentSubmission[] = [];

  do {
    const optionalArgs = {
      pageSize: 64,
      pageToken: "",
    };
    if (nextPageToken) {
      optionalArgs.pageToken = nextPageToken;
    }
    const studentSubmissions: ListStudentSubmissionsResponse = Classroom.Courses.CourseWork.StudentSubmissions.list(
      courseId,
      courseWorkId,
      optionalArgs
    );
    nextPageToken = studentSubmissions.nextPageToken;
    if (studentSubmissions.studentSubmissions) {
      studentSubmissions.studentSubmissions.forEach(function (submission) {
        submissions.push(submission);
      });
    }
  } while (nextPageToken != undefined);

  return submissions;
}

const updateCourseMemberSheet = (
  sheet: Sheet,
  courseId: string,
  courseName: string,
  data: Teacher[] | Student[],
  headers: string[]
): void => {
  sheet.clear();
  sheet.appendRow(headers);
  sheet.setFrozenRows(1);
  data.forEach(function (item) {
    if (item.profile && item.profile.name && item.profile.photoUrl) {
      sheet.appendRow([
        courseId,
        courseName,
        item.profile.emailAddress,
        item.profile.name.fullName,
        regulatePhotoUrl(item.profile.photoUrl),
      ]);
    }
  });
};

export function regulatePhotoUrl(photoUrl: string): string {
  return (photoUrl.startsWith("//") ? "https:" : "") + photoUrl;
}

function updateCoursesSheet(sheet: Sheet, data: Course[], headers: string[]): void {
  sheet.clear();
  sheet.appendRow(headers);
  sheet.setFrozenRows(1);
  data.forEach(function (course: Course) {
    if (course.id && course.ownerId) {
      const user = getTeacherProfile(course.id, course.ownerId);
      const row = [
        course.id,
        course.name,
        course.section || "",
        course.descriptionHeading || "",
        course.room || "",
        user.name?.fullName || user.emailAddress,
        user.emailAddress,
        formatDateTime(course.creationTime),
        formatDateTime(course.updateTime),
        course.enrollmentCode,
        course.courseState,
        course.alternateLink,
        course.teacherGroupEmail,
        course.courseGroupEmail,
        course.teacherFolder?.id,
        course.teacherFolder?.alternateLink,
        course.guardiansEnabled,
        course.calendarId,
        getNumStudents(course.id),
      ];
      sheet.appendRow(row);
    }
  });
}

function updateCourseWorksSheet(
  sheet: Sheet,
  courseName: string,
  data: CourseWork[],
  headers: string[]
): void {
  const courseWorkToRow = function (courseWork: CourseWork) {
    if (!courseWork.courseId) {
      throw new Error("CourseWork.courseId is not defined");
    }
    return [
      courseWork.courseId,
      courseName,
      courseWork.id,
      courseWork.title,
      courseWork.description,
      courseWork.state,
      courseWork.alternateLink,
      formatDateTime(courseWork.creationTime),
      formatDateTime(courseWork.updateTime),
      courseWork.dueDate, // TODO human readable format
      courseWork.dueTime, // TODO human readable format
      courseWork.scheduledTime,
      courseWork.maxPoints,
      courseWork.workType,
      courseWork.associatedWithDeveloper,
      courseWork.assigneeMode,
      courseWork.individualStudentsOptions?.studentIds?.join("\t"),
      courseWork.creatorUserId,
      courseWork.topicId, // TODO human readable topic name
      courseWork.assignment?.studentWorkFolder?.alternateLink,
      courseWork.multipleChoiceQuestion?.choices?.join("\t"),
    ];
  };

  sheet.clear();
  sheet.appendRow(headers);
  sheet.setFrozenRows(1);
  data.forEach(function (courseWork) {
    if (courseWork.courseId) {
      const row = courseWorkToRow(courseWork);
      sheet.appendRow(row);
    }
  });
}

function updateStudentSubmissionsSheet(
  sheet: Sheet,
  courseName: string,
  courseWorkTitle: string,
  data: StudentSubmission[],
  initialize: boolean,
  headers: string[]
): void {
  const submissionToRow = function (studentSubmission: StudentSubmission) {
    if (
      !studentSubmission.userId ||
      !studentSubmission.courseId ||
      !studentSubmission.courseWorkId
    ) {
      throw new Error(
        "studentSubmission is invalid:" + JSON.stringify(studentSubmission)
      );
    }
    const courseId: string = studentSubmission.courseId;

    const courseWorkId: string = studentSubmission.courseWorkId;

    const userProfile = getStudentProfile(courseId, studentSubmission.userId);

    const row = [
      courseId,
      courseName,
      courseWorkId,
      courseWorkTitle,
      userProfile.name?.fullName || userProfile.emailAddress,
      userProfile.emailAddress,
      regulatePhotoUrl(userProfile.photoUrl || ""),
      studentSubmission.state,
      formatDateTime(studentSubmission.creationTime),
      formatDateTime(studentSubmission.updateTime),
    ];

    if (
      studentSubmission.shortAnswerSubmission &&
      studentSubmission.shortAnswerSubmission.answer
    ) {
      row.push(studentSubmission.shortAnswerSubmission.answer);
    }

    if (
      studentSubmission.assignmentSubmission &&
      studentSubmission.assignmentSubmission.attachments
    ) {
      studentSubmission.assignmentSubmission.attachments.forEach(function (
        attachment
      ) {
        if (attachment.youTubeVideo) {
          const youTubeVideo = attachment.youTubeVideo;
          row.push(youTubeVideo.title);
          row.push(youTubeVideo.alternateLink);
        }
        if (attachment.driveFile) {
          row.push(attachment.driveFile.title);
          row.push(attachment.driveFile.alternateLink);
          row.push(attachment.driveFile.thumbnailUrl);
        }
        if (attachment.form) {
          row.push(attachment.form.title);
          row.push(attachment.form.responseUrl);
        }
      });
    }
    return row;
  };

  if (initialize) {
    sheet.clear();
    sheet.appendRow(headers);
    sheet.setFrozenRows(1);
  }
  data.forEach(function (submission) {
    sheet.appendRow(submissionToRow(submission));
  });
}

export function importCourses(headers: string[]): void {
  const user = Session.getActiveUser();
  const sheet = createOrSelectSheetBySheetName(
    "courses:" + user.getEmail(),
    "black"
  );
  updateCoursesSheet(sheet, listCourses(user.getEmail()), headers);
}

export function importCourseTeachers(headers: string[]): void {
  const { courseId, courseName } = getSelectedCourse();
  const teacherSheet = createOrSelectSheetBySheetName(
    "teachers:" + courseName,
    "yellow"
  );
  const teachers = listTeacherProfiles(courseId);
  if (teachers.length === 0) {
    throw new Error("NoRegisteredTeacher: "+courseName)
  }
  updateCourseMemberSheet(teacherSheet, courseId, courseName, teachers, headers);
}

export function importCourseStudents(headers: string[]): void {
  const { courseId, courseName } = getSelectedCourse();
  const studentSheet = createOrSelectSheetBySheetName(
    "students:" + courseName,
    "yellow"
  );
  const students = listStudentProfiles(courseId);
  if (students.length === 0) {
    throw new Error("NoRegisteredStudent: "+courseName);
  }
  updateCourseMemberSheet(studentSheet, courseId, courseName, students, headers);
}

export function importCourseWorks(headers: string[]): void {
  const { courseId, courseName } = getSelectedCourse();
  const targetSheet = createOrSelectSheetBySheetName(
    "courseworks:" + courseName,
    "yellow"
  );
  const courseWorks = listCourseWorks(courseId);
  updateCourseWorksSheet(targetSheet, courseName, courseWorks, headers);
}

type SelectedCourseWorks = {
  courseId: string;
  courseName: string;
  courseWorkId: string;
  courseWorkTitle: string;
};
export function importStudentSubmissions(studentSubmissionsHeaders: string[]): void {
  const selectedCourseWorksList: Array<SelectedCourseWorks> = getSelectedCourseWorksList();

  selectedCourseWorksList.forEach((selectedCourseWorks, index: number)=>{
    const {courseId, courseName, courseWorkId, courseWorkTitle} = selectedCourseWorks;
    if(! courseId || courseId === ""){
      return;
    }
    const studentSubmissions = listStudentSubmissions(
      courseId,
      courseName,
      courseWorkId
    );

    const targetSheet= createOrSelectSheetBySheetName(
      (selectedCourseWorksList.length == 1)?
        "submissions:" + courseName + " " + courseWorkTitle
        : "submissions:" + courseName,
      "orange"
    );

    updateStudentSubmissionsSheet(
      targetSheet,
      courseName,
      courseWorkTitle,
      studentSubmissions,
      index === 0,
      studentSubmissionsHeaders
    );
  });
}
