function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Folder Populator')
    .addItem("1. Get names and IDs", "runScript")
    .addItem("2. Copy marksheets and declarations", "populateFoldersWithTemplates")
    .addItem("3. Copy coursework submissions", "populateFolders")
    .addToUi();
}

/**
 * Prompts the user for the Google Classroom URL, retrieves the course ID
 * and the ID of the active sheet, then calls the populateSheetWithClassroomMembers
 * function with the retrieved data.
 */
function runScript() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt(
    'Enter Google Classroom URL',
    'Please enter the URL of the Google Classroom course.',
    ui.ButtonSet.OK_CANCEL
  );

  if (response.getSelectedButton() === ui.Button.OK) {
    const classroomUrl = response.getResponseText().trim();
    const classroom = Classroom.Courses;
    const courses = classroom.list().courses;
    const course = courses.find(c => c.alternateLink === classroomUrl);
    const courseId = course ? course.id : null;

    if (!courseId) {
      ui.alert('Invalid Google Classroom URL or access denied.');
      return;
    }

    const folderResponse = ui.prompt(
      'Enter Root Folder ID',
      'Please enter the ID of the root folder on Google Drive where student folders will be created.',
      ui.ButtonSet.OK_CANCEL
    );

    if (folderResponse.getSelectedButton() === ui.Button.OK) {
      const rootFolderId = folderResponse.getResponseText().trim();
      const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

      const studentSheet = spreadsheet.getSheetByName("Student Info") || spreadsheet.insertSheet("Student Info");
      const courseSheet = spreadsheet.getSheetByName("Course Info") || spreadsheet.insertSheet("Course Info");

      studentSheet.clear();
      courseSheet.clear();

      populateSheetWithClassroomMembers(courseId, studentSheet, courseSheet, rootFolderId);
    } else {
      ui.alert('Operation canceled.');
    }
  } else {
    ui.alert('Operation canceled.');
  }
}

/**
 * Retrieves a list of students and teachers from a Google Classroom course.
 *
 * @param {string} courseId - The ID of the Google Classroom course.
 * @returns {Object[]} An array of objects containing student/teacher information.
 */
function getClassroomMembers(courseId) {
  const classroomService = Classroom.Courses.Students;
  const students = classroomService.list(courseId).students;
  const teachers = Classroom.Courses.Teachers.list(courseId).teachers;

  const members = [...students, ...teachers].map(member => ({
    name: member.profile.name.fullName,
    userId: member.userId
  }));

  return members;
}

/**
 * Writes the member data to a Google Sheet.
 *
 * @param {Object[]} members - An array of objects containing member information.
 * @param {Object} sheet - The Google Sheet to write the data to.
 * @param {string} rootFolderId - The ID of the root folder on Google Drive where student folders will be created.
 */
function writeMembersToSheet(members, sheet, rootFolderId) {
  const headers = ["Name", "User ID", "Folder ID"];
  sheet.appendRow(headers);

  const rootFolder = DriveApp.getFolderById(rootFolderId);
  members.forEach(member => {
    const folder = rootFolder.createFolder(member.name);
    const folderId = folder.getId();
    const row = [member.name, member.userId, folderId];
    sheet.appendRow(row);
  });
}

function populateSheetWithClassroomMembers(courseId, studentSheet, courseSheet, rootFolderId) {
  courseSheet.appendRow(["Course ID:", courseId]);
  courseSheet.appendRow(["Template File IDs"]);

  const members = getClassroomMembers(courseId);
  writeMembersToSheet(members, studentSheet, rootFolderId);
}

/**
 * Prompts the user for a Google Classroom assignment title and a prepend string,
 * then calls the processFolderAttachments function with the retrieved data.
 */
function populateFolders() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt(
    'Enter Google Classroom Assignment Title',
    'Please enter the title of the Google Classroom assignment.',
    ui.ButtonSet.OK_CANCEL
  );

  if (response.getSelectedButton() === ui.Button.OK) {
    const assignmentTitle = response.getResponseText().trim();
    const prependResponse = ui.prompt(
      'Enter Prepend String',
      'Please enter the string to prepend to the file attachments.',
      ui.ButtonSet.OK_CANCEL
    );

    if (prependResponse.getSelectedButton() === ui.Button.OK) {
      const prependString = prependResponse.getResponseText().trim();
      const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Student Info");
      const data = sheet.getDataRange().getValues();
      processFolderAttachments(assignmentTitle, prependString, data);
    } else {
      ui.alert('Operation canceled.');
    }
  } else {
    ui.alert('Operation canceled.');
  }
}

/**
 * Processes the folder attachments for a Google Classroom assignment.
 *
 * For each submission, the function checks if any PDF files exist:
 *   - If PDF files are present, only these are copied to the student’s folder.
 *   - Otherwise, any Google Docs files are converted to PDF and then copied over.
 *
 * @param {string} assignmentTitle - The title of the Google Classroom assignment.
 * @param {string} prependString - The string to prepend to the file attachments.
 * @param {Array[]} data - The data from the spreadsheet, including folderIds.
 */
function processFolderAttachments(assignmentTitle, prependString, data) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Course Info");
  const courseId = sheet.getRange(1, 2).getValue(); // Get the courseId from the first row, second column
  const classroom = Classroom.Courses.CourseWork;
  const courses = classroom.list(courseId).courseWork;
  const assignment = courses.find(a => a.title === assignmentTitle);
  const assignmentId = assignment ? assignment.id : null;

  if (!assignmentId) {
    SpreadsheetApp.getUi().alert('Invalid Google Classroom Assignment title or access denied.');
    return;
  }
  const submissionService = Classroom.Courses.CourseWork.StudentSubmissions;

  data.forEach((row, index) => {
    if (index < 1) return; // Skip the header row

    const [name, userId, folderId] = row;

    if (!folderId) {
      console.log(`No folder ID found for user ${userId}`);
      return;
    }

    const submissions = submissionService.list(courseId, assignmentId, { userId: userId }).studentSubmissions;

    submissions.forEach(submission => {
      const attachments = submission.assignmentSubmission.attachments || [];
      let pdfFiles = [];
      let googleDocsFiles = [];

      attachments.forEach(attachment => {
        if (attachment.driveFile) {
          const file = DriveApp.getFileById(attachment.driveFile.id);
          if (isPdf(file)) {
            pdfFiles.push(file);
          } else if (isGoogleDoc(file)) {
            googleDocsFiles.push(file);
          }
        }
      });

      const folder = DriveApp.getFolderById(folderId);
      if (pdfFiles.length > 0) {
        pdfFiles.forEach(file => copyFile(file, folder, prependString, name));
      } else {
        googleDocsFiles.forEach(file => copyGoogleDocAsPdf(file, folder, prependString, name));
      }
    });
  });
}

/**
 * Populates student folders with template files listed in the "Course Info" sheet.
 * The function retrieves the template file IDs and copies each file into each student's folder.
 * The new file names are prepended with the student's initials (first initial of first name and first two initials of surname).
 */
function populateFoldersWithTemplates() {
  // Get the active spreadsheet and the relevant sheets
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const studentSheet = spreadsheet.getSheetByName("Student Info");
  const courseSheet = spreadsheet.getSheetByName("Course Info");

  // Check if the required sheets exist
  if (!studentSheet || !courseSheet) {
    SpreadsheetApp.getUi().alert('Student Info or Course Info sheet not found.');
    return;
  }

  // Get student data from the "Student Info" sheet
  const studentData = studentSheet.getDataRange().getValues();
  // Get template file IDs from the "Course Info" sheet, starting from the third row
  const templateFileIds = courseSheet.getRange(3, 1, courseSheet.getLastRow() - 2, 1)
    .getValues().flat().filter(id => id);

  // Check if there are any template file IDs
  if (templateFileIds.length === 0) {
    SpreadsheetApp.getUi().alert('No template file IDs found.');
    return;
  }

  // Iterate over each student, starting from the second row (excluding header)
  studentData.slice(1).forEach(([name, userId, folderId]) => {
    // Extract first initial of first name and first two initials of surname
    const [firstName, lastName] = name.split(' ');
    const initials = `${firstName.charAt(0)}${lastName.charAt(0)}${lastName.charAt(1)}`.toUpperCase();

    // Iterate over each template file ID
    templateFileIds.forEach(templateFileId => {
      try {
        // Get the file by its ID
        const driveFile = DriveApp.getFileById(templateFileId);
        // Get the student's folder by its ID
        const folder = DriveApp.getFolderById(folderId);
        // Create a new file name with the initials prepended
        const newFileName = `${initials}_${driveFile.getName()}`;
        // Make a copy of the file in the student's folder with the new name
        driveFile.makeCopy(newFileName, folder);
        // Log the success message to the console
        console.log(`Copied file ${driveFile.getName()} to ${folder.getName()} as ${newFileName}`);
      } catch (e) {
        // Log any errors encountered during the file copy process
        console.error(`Failed to copy file with ID ${templateFileId} to folder with ID ${folderId}: ${e.message}`);
      }
    });
  });
}

/* Helper Functions */

/**
 * Checks if the provided file is a PDF.
 *
 * @param {File} file - The Drive file to check.
 * @returns {boolean} True if the file's MIME type is PDF.
 */
function isPdf(file) {
  return file.getMimeType() === "application/pdf";
}

/**
 * Checks if the provided file is a Google Docs file.
 *
 * @param {File} file - The Drive file to check.
 * @returns {boolean} True if the file's MIME type indicates a Google Docs document.
 */
function isGoogleDoc(file) {
  return file.getMimeType() === "application/vnd.google-apps.document";
}

/**
 * Copies the given file to the specified folder with a new name.
 *
 * @param {File} file - The Drive file to copy.
 * @param {Folder} folder - The destination folder.
 * @param {string} prependString - The string to prepend to the file name.
 * @param {string} userName - The name of the user (for logging purposes).
 */
function copyFile(file, folder, prependString, userName) {
  const newFileName = `${prependString}_${file.getName()}`;
  file.makeCopy(newFileName, folder);
  console.log(`${userName}'s document has been moved.`);
}

/**
 * Converts a Google Docs file to PDF and copies it to the specified folder with a new name.
 *
 * @param {File} file - The Google Docs file to convert.
 * @param {Folder} folder - The destination folder.
 * @param {string} prependString - The string to prepend to the file name.
 * @param {string} userName - The name of the user (for logging purposes).
 */
function copyGoogleDocAsPdf(file, folder, prependString, userName) {
  const pdfBlob = file.getAs("application/pdf");
  // Append '.pdf' to the original name for clarity.
  const newFileName = `${prependString}_${file.getName()}.pdf`;
  folder.createFile(pdfBlob).setName(newFileName);
  console.log(`${userName}'s document (converted to PDF) has been moved.`);
}
