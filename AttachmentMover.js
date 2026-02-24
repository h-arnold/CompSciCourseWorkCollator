
/**
 * Main application class that orchestrates all operations
 */
class FolderPopulator {
  constructor() {
    // Initialize managers as needed
  }
  
  /**
   * Populates sheets with classroom members
   * @param {string} courseId - The ID of the Google Classroom course
   * @param {Object} studentSheet - The student info sheet
   * @param {Object} courseSheet - The course info sheet
   * @param {string} rootFolderId - The ID of the root folder
   */
  populateSheetWithClassroomMembers(courseId, studentSheet, courseSheet, rootFolderId) {
    courseSheet.appendRow(["Course ID:", courseId]);
    courseSheet.appendRow(["Template File IDs"]);
  
    const members = ClassroomManager.getClassroomMembers(courseId);
    SpreadsheetManager.writeMembersToSheet(members, studentSheet, rootFolderId);
  }
  
  /**
   * Initializes the classroom and folder setup
   */
  initializeClassroomAndFolders() {
    const classroomUrlResponse = UIManager.promptUser(
      'Enter Google Classroom URL',
      'Please enter the URL of the Google Classroom course.'
    );
  
    if (classroomUrlResponse.getSelectedButton() === SpreadsheetApp.getUi().Button.OK) {
      const classroomUrl = classroomUrlResponse.getResponseText().trim();
      const classroom = Classroom.Courses;
      const courses = classroom.list().courses;
      const course = courses.find(c => c.alternateLink === classroomUrl);
      const courseId = course ? course.id : null;
  
      if (!courseId) {
        UIManager.showAlert('Invalid Google Classroom URL or access denied.');
        return;
      }
  
      const folderResponse = UIManager.promptUser(
        'Enter Root Folder ID',
        'Please enter the ID of the root folder on Google Drive where student folders will be created.'
      );
  
      if (folderResponse.getSelectedButton() === SpreadsheetApp.getUi().Button.OK) {
        const rootFolderId = folderResponse.getResponseText().trim();
        const { studentSheet, courseSheet } = SpreadsheetManager.getSpreadsheetSheets();
  
        studentSheet.clear();
        courseSheet.clear();
  
        this.populateSheetWithClassroomMembers(courseId, studentSheet, courseSheet, rootFolderId);
      } else {
        UIManager.showAlert('Operation canceled.');
      }
    } else {
      UIManager.showAlert('Operation canceled.');
    }
  }
  
  /**
   * Processes attachments from Google Classroom assignments
   */
  processAssignmentAttachments() {
    const assignmentTitleResponse = UIManager.promptUser(
      'Enter Google Classroom Assignment Title',
      'Please enter the title of the Google Classroom assignment.'
    );
  
    if (assignmentTitleResponse.getSelectedButton() === SpreadsheetApp.getUi().Button.OK) {
      const assignmentTitle = assignmentTitleResponse.getResponseText().trim();
      const prependResponse = UIManager.promptUser(
        'Enter Prepend String',
        'Please enter the string to prepend to the file attachments.'
      );
  
      if (prependResponse.getSelectedButton() === SpreadsheetApp.getUi().Button.OK) {
        const prependString = prependResponse.getResponseText().trim();
        const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Student Info");
        const data = sheet.getDataRange().getValues();
        this.processFolderAttachments(assignmentTitle, prependString, data);
      } else {
        UIManager.showAlert('Operation canceled.');
      }
    } else {
      UIManager.showAlert('Operation canceled.');
    }
  }
  
  /**
   * Processes folder attachments for a Google Classroom assignment
   * @param {string} assignmentTitle - The title of the Google Classroom assignment
   * @param {string} prependString - The string to prepend to file attachments
   * @param {Array[]} data - The data from the spreadsheet, including folderIds
   */
  processFolderAttachments(assignmentTitle, prependString, data) {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Course Info");
    const courseId = sheet.getRange(1, 2).getValue(); // Get the courseId from the first row, second column
    const assignmentId = ClassroomManager.getAssignmentId(courseId, assignmentTitle);
  
    if (!assignmentId) {
      UIManager.showAlert('Invalid Google Classroom Assignment title or access denied.');
      return;
    }
  
    data.forEach((row, index) => {
      if (index < 1) return; // Skip the header row
  
      const [name, userId, folderId] = row;
  
      if (!folderId) {
        console.log(`No folder ID found for user ${userId}`);
        return;
      }
  
      const submissions = ClassroomManager.getStudentSubmissions(courseId, assignmentId, userId);
  
      submissions.forEach(submission => {
        const attachments = submission.assignmentSubmission.attachments || [];
        let pdfFiles = [];
        let googleDocsFiles = [];
        let zipFiles = []; // Array to hold zip files
  
        attachments.forEach(attachment => {
          if (attachment.driveFile) {
            try {
              const file = DriveApp.getFileById(attachment.driveFile.id);
              if (DriveManager.isPdf(file)) {
                pdfFiles.push(file);
              } else if (DriveManager.isGoogleDoc(file)) {
                googleDocsFiles.push(file);
              } else if (DriveManager.isZip(file)) { // Check for zip files
                zipFiles.push(file);
              }
            } catch (e) {
              console.error(`Error accessing file ID ${attachment.driveFile.id}: ${e.message}`);
            }
          }
        });
  
        try {
          const folder = DriveApp.getFolderById(folderId);
  
          // Always copy zip files if they exist
          zipFiles.forEach(file => DriveManager.copyFile(file, folder, prependString, name));
  
          // Then handle PDFs or Google Docs
          if (pdfFiles.length > 0) {
            pdfFiles.forEach(file => DriveManager.copyFile(file, folder, prependString, name));
          } else if (googleDocsFiles.length > 0) { // Only process Google Docs if no PDFs were found
            googleDocsFiles.forEach(file => DriveManager.copyGoogleDocAsPdf(file, folder, prependString, name));
          }
        } catch (e) {
           console.error(`Error processing folder ID ${folderId} for user ${name}: ${e.message}`);
        }
      });
    });
  }

  /**
   * Runs the ICT Content Creation collation workflow.
   * Creates student folders based on Student Info sheet and collates assignment attachments.
   */
  runICTContentCreationCollation() {
    const classroomUrlResponse = UIManager.promptUser(
      'ICT: Enter Google Classroom URL',
      'Please enter the URL of the Google Classroom course.'
    );

    if (classroomUrlResponse.getSelectedButton() !== SpreadsheetApp.getUi().Button.OK) {
      UIManager.showAlert('Operation canceled.');
      return;
    }

    const classroomUrl = classroomUrlResponse.getResponseText().trim();
    const classroom = Classroom.Courses;
    const courses = classroom.list().courses || [];
    const course = courses.find(c => c.alternateLink === classroomUrl);
    const courseId = course ? course.id : null;

    if (!courseId) {
      UIManager.showAlert('Invalid Google Classroom URL or access denied.');
      return;
    }

    const rootFolderResponse = UIManager.promptUser(
      'ICT: Enter Root Folder ID',
      'Please enter the root folder ID where ICT student folders should be created.'
    );

    if (rootFolderResponse.getSelectedButton() !== SpreadsheetApp.getUi().Button.OK) {
      UIManager.showAlert('Operation canceled.');
      return;
    }

    const assignmentTitleResponse = UIManager.promptUser(
      'ICT: Enter Assignment Title',
      'Please enter the Google Classroom assignment title to collate attachments from.'
    );

    if (assignmentTitleResponse.getSelectedButton() !== SpreadsheetApp.getUi().Button.OK) {
      UIManager.showAlert('Operation canceled.');
      return;
    }

    const rootFolderId = rootFolderResponse.getResponseText().trim();
    const assignmentTitle = assignmentTitleResponse.getResponseText().trim();

    const { studentSheet, courseSheet } = SpreadsheetManager.getSpreadsheetSheets();
    const ictStudents = SpreadsheetManager.getICTStudentsFromSheet(studentSheet);

    if (ictStudents.length === 0) {
      UIManager.showAlert('No ICT students found in Student Info sheet.');
      return;
    }

    courseSheet.clear();
    courseSheet.appendRow(['Course ID:', courseId]);

    this.createICTStudentFolders(rootFolderId, studentSheet, ictStudents);
    const result = this.processICTAssignmentAttachments(courseId, assignmentTitle, ictStudents);

    UIManager.showAlert(
      `ICT collation complete.\nProcessed students: ${result.processedStudents}\nCopied files: ${result.copiedFiles}\nSkipped (no sheet match): ${result.unmatchedSubmissions}\nSkipped (no submissions): ${result.studentsWithoutSubmissions}`
    );
  }

  /**
   * Initializes Student Info sheet for ICT collation.
   */
  setupICTStudentInfoSheet() {
    const { studentSheet } = SpreadsheetManager.getSpreadsheetSheets();
    const ui = SpreadsheetApp.getUi();
    const response = ui.alert(
      'ICT Student Info setup',
      'Do you want to clear existing Student Info data before writing ICT headers?',
      ui.ButtonSet.YES_NO_CANCEL
    );

    if (response === ui.Button.CANCEL) {
      UIManager.showAlert('Operation canceled.');
      return;
    }

    SpreadsheetManager.initializeICTStudentInfoSheet(studentSheet, response === ui.Button.YES);
    UIManager.showAlert('Student Info sheet is ready with ICT headers: Student ID, Name, User ID, Folder ID.');
  }

  /**
   * Creates ICT student folders using the format: "studentId studentName".
   * @param {string} rootFolderId - Root Drive folder ID
   * @param {Object} studentSheet - Student Info sheet
   * @param {Object[]} students - ICT student records
   */
  createICTStudentFolders(rootFolderId, studentSheet, students) {
    const rootFolder = DriveApp.getFolderById(rootFolderId);

    students.forEach(student => {
      if (!student.studentId || !student.name) {
        console.warn(`Skipping row ${student.rowIndex} due to missing Student ID or Name.`);
        return;
      }

      const folderName = `${student.studentId} ${student.name}`;
      const folder = DriveManager.createFolder(rootFolder, folderName);
      student.folderId = folder.getId();
    });

    SpreadsheetManager.updateICTFolderIds(studentSheet, students);
  }

  /**
   * Processes ICT assignment attachments for sheet-listed students.
   * Copies all Drive attachments preserving names, converting Google Docs to PDF.
   * @param {string} courseId - Classroom course ID
   * @param {string} assignmentTitle - Classroom assignment title
   * @param {Object[]} students - ICT student records
   * @returns {Object} Summary stats
   */
  processICTAssignmentAttachments(courseId, assignmentTitle, students) {
    const assignmentId = ClassroomManager.getAssignmentId(courseId, assignmentTitle);

    if (!assignmentId) {
      throw new Error('Invalid Google Classroom assignment title or access denied.');
    }

    let processedStudents = 0;
    let copiedFiles = 0;
    let studentsWithoutSubmissions = 0;
    let unmatchedSubmissions = 0;

    students.forEach(student => {
      if (!student.userId || !student.folderId) {
        console.warn(`Skipping student ${student.name} due to missing User ID or Folder ID.`);
        unmatchedSubmissions++;
        return;
      }

      let submissions = [];
      try {
        submissions = ClassroomManager.getStudentSubmissions(courseId, assignmentId, student.userId) || [];
      } catch (error) {
        console.error(`Unable to fetch submissions for ${student.name}: ${error.message}`);
        return;
      }

      if (submissions.length === 0) {
        studentsWithoutSubmissions++;
        return;
      }

      const destinationFolder = DriveApp.getFolderById(student.folderId);
      processedStudents++;

      submissions.forEach(submission => {
        const attachments = submission.assignmentSubmission?.attachments || [];

        attachments.forEach(attachment => {
          if (!attachment.driveFile?.id) {
            return;
          }

          try {
            const file = DriveApp.getFileById(attachment.driveFile.id);

            if (DriveManager.isGoogleDoc(file)) {
              DriveManager.copyGoogleDocAsPdfPreserveName(file, destinationFolder);
            } else {
              DriveManager.copyFilePreserveName(file, destinationFolder);
            }

            copiedFiles++;
          } catch (error) {
            console.error(`Error copying attachment ${attachment.driveFile.id} for ${student.name}: ${error.message}`);
          }
        });
      });
    });

    return {
      processedStudents,
      copiedFiles,
      studentsWithoutSubmissions,
      unmatchedSubmissions
    };
  }
  
  /**
   * Populates student folders with template files
   */
  populateFoldersWithTemplates() {
    // Get the active spreadsheet and the relevant sheets
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const studentSheet = spreadsheet.getSheetByName("Student Info");
    const courseSheet = spreadsheet.getSheetByName("Course Info");
  
    // Check if the required sheets exist
    if (!studentSheet || !courseSheet) {
      UIManager.showAlert('Student Info or Course Info sheet not found.');
      return;
    }
  
    // Get student data from the "Student Info" sheet
    const studentData = studentSheet.getDataRange().getValues();
    // Get template file IDs from the "Course Info" sheet, starting from the third row
    const templateFileIds = courseSheet.getRange(3, 1, courseSheet.getLastRow() - 2, 1)
      .getValues().flat().filter(Boolean);
  
    // Check if there are any template file IDs
    if (templateFileIds.length === 0) {
      UIManager.showAlert('No template file IDs found.');
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
  
          // Check if the file should be copied (handles existing file prompt)
          if (DriveManager.checkAndHandleExistingFile(folder, newFileName, name)) {
            // Make a copy of the file in the student's folder with the new name
            driveFile.makeCopy(newFileName, folder);
            // Log the success message to the console
            console.log(`Copied file ${driveFile.getName()} to ${folder.getName()} as ${newFileName}`);
          }
        } catch (e) {
          // Log any errors encountered during the file copy process
          console.error(`Failed to copy file with ID ${templateFileId} to folder with ID ${folderId}: ${e.message}`);
        }
      });
    });
  }
}

// Global entry point functions that maintain backward compatibility
function onOpen() {
  UIManager.createMenu();
}

function runScript() {
  const folderPopulator = new FolderPopulator();
  folderPopulator.initializeClassroomAndFolders();
}

function populateFolders() {
  const folderPopulator = new FolderPopulator();
  folderPopulator.processAssignmentAttachments();
}

function populateFoldersWithTemplates() {
  const folderPopulator = new FolderPopulator();
  folderPopulator.populateFoldersWithTemplates();
}

function runICTContentCreationCollation() {
  try {
    const folderPopulator = new FolderPopulator();
    folderPopulator.runICTContentCreationCollation();
  } catch (e) {
    UIManager.showAlert(`ICT collation failed: ${e.message}`);
  }
}

function setupICTStudentInfoSheet() {
  try {
    const folderPopulator = new FolderPopulator();
    folderPopulator.setupICTStudentInfoSheet();
  } catch (e) {
    UIManager.showAlert(`ICT sheet setup failed: ${e.message}`);
  }
}

/**
 * Entry point function for merging PDFs for all students
 * Displays a prompt to allow selection of recursive folder scanning
 */
async function mergeAllStudentPDFs() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    'Merge PDFs for all students',
    'Do you want to include PDFs from subfolders in the search?',
    ui.ButtonSet.YES_NO_CANCEL);
  
  if (response == ui.Button.CANCEL) {
    ui.alert('Operation cancelled.');
    return;
  }
  
  const recursive = (response == ui.Button.YES);
  
  // Show a loading message
  ui.alert('Starting PDF merge process for all students. This may take some time.');
  
  try {
    // Run the merge operation
    const result = await PDFMerger.getInstance().mergePDFsForAllStudents(recursive);
    
    // Show results to the user
    if (result.success) {
      ui.alert('Success', `${result.message}\n\nProcessed ${result.studentResults.length} students.`, ui.ButtonSet.OK);
    } else {
      ui.alert('Error', result.message, ui.ButtonSet.OK);
    }
  } catch (e) {
    ui.alert('Error', `An error occurred: ${e.message}`, ui.ButtonSet.OK);
  }
}
