/**
 * Class for Google Classroom operations
 * This class is shared between declarationOrganiser.js and attachmentMover.js
 */
class ClassroomManager {
  /**
   * Gets assignment ID from assignment title
   * @param {string} courseId - The course ID
   * @param {string} assignmentTitle - The title of the assignment
   * @return {string|null} The assignment ID or null if not found
   */
  static getAssignmentId(courseId, assignmentTitle) {
    const classroom = Classroom.Courses.CourseWork;
    const courses = classroom.list(courseId).courseWork;
    const assignment = courses.find(a => a.title === assignmentTitle);
    return assignment ? assignment.id : null;
  }
  
  /**
   * Gets student submissions for an assignment
   * @param {string} courseId - The course ID
   * @param {string} assignmentId - The assignment ID
   * @param {string} userId - The user ID
   * @return {Array} Array of submissions
   */
  static getStudentSubmissions(courseId, assignmentId, userId) {
    const submissionService = Classroom.Courses.CourseWork.StudentSubmissions;
    return submissionService.list(courseId, assignmentId, { userId: userId }).studentSubmissions;
  }

  /**
   * Retrieves a list of students and teachers from a Google Classroom course
   * @param {string} courseId - The ID of the Google Classroom course
   * @returns {Object[]} An array of objects containing student/teacher information
   */
  static getClassroomMembers(courseId) {
    const classroomService = Classroom.Courses.Students;
    const students = classroomService.list(courseId).students;
    const teachers = Classroom.Courses.Teachers.list(courseId).teachers;
  
    const members = [...students, ...teachers].map(member => ({
      name: member.profile.name.fullName,
      userId: member.userId
    }));
  
    return members;
  }
}