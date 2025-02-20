/**
 * Copies and renames a Drive file.
 *
 * @param {string} driveFileId - The ID of the Drive file.
 * @param {string} folderId - The ID of the folder where the file will be copied.
 * @param {string} newFileName - The name of the file
 */
function copyAndRenameFile(driveFileId, folderId, newFileName) {
  const driveFile = DriveApp.getFileById(driveFileId);
  const folder = DriveApp.getFolderById(folderId);
  driveFile.makeCopy(newFileName, folder);
}

/**
 * Retrieves the text from a Google Docs attachment and searches for a 4 character string and a 5 character string consisting only of numbers.
 *
 * @param {string} driveFileId - The ID of the Drive file.
 * @return {Object} An object containing the first 4 character string (CandidateNo) and the first 5 character string (CentreNo) consisting only of numbers found in the document, or null if not found.
 */
function getCandidateAndCentreNo(driveFileId) {
  const doc = DocumentApp.openById(driveFileId);
  const body = doc.getBody();
  const text = body.getText();
  const candidateNoMatch = text.match(/\b\d{4}\b/); // Regular expression to find a 4 digit number
  const centreNoMatch = text.match(/\b\d{5}\b/); // Regular expression to find a 5 digit number
  console.log("Candidate Number is: " + candidateNoMatch + "\n Centre Number is: " + centreNoMatch)
  return {
    CandidateNo: candidateNoMatch ? candidateNoMatch[0] : null,
    CentreNo: centreNoMatch ? centreNoMatch[0] : null
  };
}

/**
 * Converts a Google Docs file to a PDF.
 *
 * @param {string} driveFileId - The ID of the Drive file.
 * @return {File} The converted PDF file.
 */
function convertToPdf(driveFileId) {
  const driveFile = DriveApp.getFileById(driveFileId);
  const blob = driveFile.getAs('application/pdf');
  const pdfFile = DriveApp.createFile(blob);
  return pdfFile;
}

/**
 * Formats a student name.
 *
 * @param {string} name - The student name in the format of {firstname} {surname}.
 * @return {string} The formatted string.
 */
function formatStudentName(name) {
  const [firstName, surname] = name.split(' ');
  const formattedSurname = surname.slice(0, 2).charAt(0).toUpperCase() + surname.slice(1, 2).toLowerCase();
  const formattedFirstName = firstName.charAt(0).toUpperCase();
  return `${formattedSurname}_${formattedFirstName}`;
}


/**
 * Creates a file name.
 *
 * @param {string} CentreNo - The centre number.
 * @param {string} CandidateNo - The candidate number.
 * @param {string} formattedName - The formatted name.
 * @return {string} The formatted file name.
 */
function createFileName(CentreNo, CandidateNo, name) {
  return `${CentreNo}_${CandidateNo}_${formatStudentName(name)}`;
}




/**
 * Processes the folder attachments for a Google Classroom assignment.
 *
 * @param {string} assignmentTitle - The title of the Google Classroom assignment.
 * @param {Array[]} data - The data from the spreadsheet, including folderIds.
 */
function processFolderAttachmentsForDeclarationsOnly(assignmentTitle, data) {
  const courseId = data[0][0]; // Get the courseId from the first row
  const classroom = Classroom.Courses.CourseWork;
  const courses = classroom.list(courseId).courseWork; // Use the courseId to fetch assignments
  const assignment = courses.find(a => a.title === assignmentTitle);
  const assignmentId = assignment ? assignment.id : null;

  if (!assignmentId) {
    SpreadsheetApp.getUi().alert('Invalid Google Classroom Assignment title or access denied.');
    return;
  }
  const submissionService = Classroom.Courses.CourseWork.StudentSubmissions;

  data.forEach((row, index) => {
    if (index < 2) return; // Skip the header row

    const [name, userId, folderId] = row;

    if (!folderId) {
      console.log(`No folder ID found for user ${userId}`);
      return;
    }

    const submissions = submissionService.list(courseId, assignmentId, { userId: userId }).studentSubmissions;

    submissions.forEach(submission => {
      const attachments = submission.assignmentSubmission.attachments || [];
      attachments.forEach(attachment => {
        if (attachment.driveFile) {
          const driveFileId = attachment.driveFile.id;
          const { CandidateNo, CentreNo } = getCandidateAndCentreNo(driveFileId);
          if (CandidateNo && CentreNo) {
            console.log(`Candidate number ${CandidateNo} and Centre number ${CentreNo} found in document.`);
            const pdfFile = convertToPdf(driveFileId);
            const pdfFileId = pdfFile.getId();
            const newFileName = createFileName(CentreNo, CandidateNo, name);
            copyAndRenameFile(pdfFileId, folderId, newFileName);
            console.log(name +"'s document has been moved and converted to PDF.")
          }
        }
      });
    });
  }
)}
