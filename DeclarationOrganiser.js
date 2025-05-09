/**
 * Main processor class that orchestrates the declaration processing workflow
 */
/**
 * Processes and manages declaration forms for students by merging, converting, and organizing files
 * from Google Classroom and Google Drive.
 * 
 * This class provides functionality to:
 * - Merge signed declarations from Google Classroom with unsigned declarations containing marks
 * - Include marking grids in the final documents
 * - Convert documents to PDF format
 * - Rename files according to WJEC naming convention
 * - Handle file organization and management in student folders
 * 
 * Depends on:
 * - TextProcessor (for text manipulation and filename creation)
 * - ClassroomManager (static methods for Google Classroom operations)
 * - DriveManager (static methods for Google Drive operations)
 */
class DeclarationProcessor {
  constructor() {
    this.textProcessor = new TextProcessor();
    this.gClassroomDeclarationFile = null;
    this.studentFolderDeclarationFile = null;
    this.studentMarkingGridFile = null;
    // ClassroomManager is a global class with static methods
    // DriveManager is now also being used with static methods
  }

  /**
   * Finds the first Google Doc attachment in a Google Classroom submission
   * @param {Object} submission - The Google Classroom submission object
   * @return {Object|null} The Google Drive File object if found, null otherwise
   */
  findFirstGoogleDocAttachment(submission) {
    const attachments = submission.assignmentSubmission.attachments || [];

    for (const attachment of attachments) {
      if (attachment.driveFile) {
        //Check that the attachment is a Google Drive File
        const file = DriveApp.getFileById(attachment.driveFile.id);

        // Check if the file is a Google Doc
        if (DriveManager.isGoogleDoc(file)) {
          return file;
        }
      }
    }

    return null;
  }

  /**
   * Creates the final declaration forms for each student in the spreadsheet by merging:
   *   - the signed declaration on Google Classroom (which doesn't have marks)
   *   - the unsigned declaration in the student folder (which has marks)
   *    - the marking grid
   * Then converts it into a final PDF and renames it according to the
   * WJEC required convention which is:
   * {centreNumber}_{candidateNumber}_{firstInitial}_{firstTwoInitialOfSurname}
   * 
   * @param {string} assignmentTitle - The title of the Google Classroom assignment
   * @param {Array[]} studentFolderData - The data from the spreadsheet, including folderIds
   * @param {string} [gClassroomDeclarationFileId] - Optional ID of the base declaration file to merge with
   */
  createFinalDeclarationForms(assignmentTitle, studentFolderData) {
    const courseId = studentFolderData[0][0]; // Get the courseId from the first row
    const assignmentId = ClassroomManager.getAssignmentId(courseId, assignmentTitle);

    if (!assignmentId) {
      SpreadsheetApp.getUi().alert('Invalid Google Classroom Assignment title or access denied.');
      return;
    }


    studentFolderData.forEach((studentDataRow, index) => {
      if (index < 3) return; // Skip the header rows

      const name = studentDataRow[0];
      const userId = studentDataRow[1];
      const folderId = studentDataRow[2];

      // Initialise a student sample prefixes array for creating the sample folders later.



      if (!folderId) {
        console.log(`No folder ID found for user ${userId}`);
        return;
      }

      const studentFolder = DriveApp.getFolderById(folderId);

      // Get the declaration and marking grid files
      this.studentFolderDeclarationFile = DriveManager.findFilesBySubstring(
        studentFolder,
        "Declaration",
        false,
        "application/vnd.google-apps.document",
        "suffix");

      this.studentMarkingGridFile = DriveManager.findFilesBySubstring(
        studentFolder,
        "Marking Grid",
        false,
        "application/vnd.google-apps.document",
        "suffix");

      const submissions = ClassroomManager.getStudentSubmissions(courseId, assignmentId, userId);

      submissions.forEach(submission => {
        // Find the first Google Doc attachment
        this.gClassroomDeclarationFile = this.findFirstGoogleDocAttachment(submission);

        if (!this.gClassroomDeclarationFile) {
          console.error(`No Google Doc attachment found for user ${name}`);
          throw new Error(`No Google Doc attachment found for user ${name}`);
        }

        console.log(`Found Google Doc attachment for user ${name}`);
        const prefixAndFilename = this.generateStudentSubmissionPrefixAndFilename(name)

        // Now that we have all the files we need, create the final PDF.
        this.createFinalDeclarationPDF(
          name,
          prefixAndFilename.fileName,
          studentFolder
        );

        studentDataRow.push(prefixAndFilename.studentSubmissionPrefix)



      });

    });
    return studentFolderData//TODO: Come up with a more elegant way to get the submission prefix than buried down here.


  }

  /**
   * Generates a filename for a declaration document according to the 
   * WJEC required convention which is:
   * {centreNumber}_{candidateNumber}_{firstInitial}_{firstTwoInitialOfSurname}
   * @param {string} name - The student name
   * @return {Object|null} Object containing fileName and studentSubmissionPrefix or null if required information not found
   */
  generateStudentSubmissionPrefixAndFilename(name) {
    const { CandidateNo, CentreNo } = this.textProcessor.getCandidateAndCentreNo(this.gClassroomDeclarationFile)

    if (CandidateNo && CentreNo) {
      console.log(`Candidate number ${CandidateNo} and Centre number ${CentreNo} found in document.`);
      const studentSubmissionPrefix = this.textProcessor.createStudentSubmissionPrefix(CentreNo, CandidateNo, name);
      const newFileName = this.textProcessor.createFileName(studentSubmissionPrefix);
      console.log(`Generated filename for ${name}: ${newFileName}`);
      return {
        fileName: newFileName,
        studentSubmissionPrefix: studentSubmissionPrefix
      };
    } else {
      const fileId = this.gClassroomDeclarationFile.getId();
      console.log(`Could not find Candidate/Centre number in document ID ${fileId} for ${name}. Cannot generate filename.`);
      return null;
    }
  }

  /**
   * Checks which declaration files are available and determines which files to use
   * for the merge process
   * @private
   * @return {Object} Object containing information about available files and merge requirements
   */
  _checkDeclarationFilesAvailability() {
    let hasGClassroomFile = this.gClassroomDeclarationFile;
    let hasStudentFolderFile = this.studentFolderDeclarationFile;

    // If we get an error when trying to get an Id for either file, then the file is missing.

    try {
      this.gClassroomDeclarationFile.getId();
    } catch (e) {
      hasClassroomFile = null
      console.error(`Unable to find declaration file on Google Classroom. Error message: ${e.message}`)
    }

    try {
      this.studentFolderDeclarationFile.getId();
    } catch (e) {
      hasStudentFolderFile = null
      console.error(`Unable to find declaration file on Google Classroom. Error message: ${e.message}`)
    }



    console.log(`Declaration files available - Google Classroom: ${hasGClassroomFile}, Student Folder: ${hasStudentFolderFile}`);

    let baseFile = null, sourceFile = null;
    let isMergeNeeded = false;

    if (hasGClassroomFile && hasStudentFolderFile) {
      // Both files exist, so we need to perform a merge
      baseFile = this.gClassroomDeclarationFile;
      sourceFile = this.studentFolderDeclarationFile;
      isMergeNeeded = true;
    } else if (hasGClassroomFile) {
      // Only Google Classroom file exists
      console.log(`Only Google Classroom declaration file exists (ID: ${this.gClassroomDeclarationFile.getId()}). Using it directly.`);
      baseFile = this.gClassroomDeclarationFile;
    } else if (hasStudentFolderFile) {
      // Only student folder file exists
      console.log(`Only student folder declaration file exists (ID: ${this.studentFolderDeclarationFile.getId()}). Using it directly.`);
      baseFile = this.studentFolderDeclarationFile;
    }

    return {
      hasGClassroomFile,
      hasStudentFolderFile,
      baseFile,
      sourceFile,
      isMergeNeeded,
      canProceed: hasGClassroomFile || hasStudentFolderFile
    };
  }

  /**
   * Merges table data from a source document into a copy of a base document.
   * Specifically targets 'Title of Task:' and 'TOTAL' tables.
   * The copy of the base document is placed in the same folder as the source document.
   * If either file is missing, it will proceed with the available file.
   * @param {string} mergedFileName - The desired name for the newly created merged document.
   * @return {Object|null} The Google Drive File object of the newly created document, or null on failure.
   */
  mergeDeclarations(mergedFileName) {
    // Use the helper method to check file availability
    const fileAvailability = this._checkDeclarationFilesAvailability();

    // If neither file exists, return null as we can't proceed
    if (!fileAvailability.canProceed) {
      console.error("Both declaration files are missing. Cannot proceed with merge.");
      return null;
    }

    if (fileAvailability.isMergeNeeded) {
      console.log(`Starting merge process: Google Classroom Doc ID: ${this.gClassroomDeclarationFile.getId()}, Student Folder Doc ID: ${this.studentFolderDeclarationFile.getId()}, New Filename: ${mergedFileName}`);
    }

    // Determine the destination folder
    let destinationFolderId = null;
    try {
      const fileForParent = fileAvailability.hasStudentFolderFile ? this.studentFolderDeclarationFile : this.gClassroomDeclarationFile;
      const parents = fileForParent.getParents();
      if (parents.hasNext()) {
        destinationFolderId = parents.next().getId();
        console.log(`Target destination folder ID: ${destinationFolderId}`);
      } else {
        console.warn(`Source document has no parent folder. Copy will be placed in root.`);
      }
    } catch (e) {
      console.error(`Error getting parent folder: ${e}. Copy will be placed in root.`);
    }

    // Create a copy of the base document in the destination folder
    const mergedDoc = DriveManager.copyDocument(fileAvailability.baseFile, mergedFileName, destinationFolderId);
    if (!mergedDoc) {
      console.error("Failed to create a copy of the base document. Aborting merge.");
      return null;
    }
    const mergedDocId = mergedDoc.getId();
    console.log(`Created copy of base document with ID: ${mergedDocId}`);

    // If no merge is needed, we're done
    if (!fileAvailability.isMergeNeeded) {
      console.log(`No merge needed. Successfully created document: ${mergedFileName} (ID: ${mergedDocId})`);
      return mergedDoc;
    }

    // If merge is needed, proceed with extracting and replacing table data
    const titleTableData = this.textProcessor.extractTableText(fileAvailability.sourceFile, "Title of Task:");
    const totalTableData = this.textProcessor.extractTableText(fileAvailability.sourceFile, "TOTAL");

    if (!titleTableData) {
      console.warn(`Could not extract 'Title of Task:' table data from source doc ${fileAvailability.sourceFile.getId()}.`);
    }
    if (!totalTableData) {
      console.warn(`Could not extract 'TOTAL' table data from source doc ${fileAvailability.sourceFile.getId()}.`);
    }

    // Replace data in the new (merged) document tables
    let success = true;
    if (titleTableData) {
      success = this.textProcessor.replaceTableText(mergedDoc, "Title of Task:", titleTableData) && success;
    } else {
      console.log("Skipping replacement for 'Title of Task:' table as no data was extracted.");
    }

    if (totalTableData) {
      success = this.textProcessor.replaceTableText(mergedDoc, "TOTAL", totalTableData) && success;
    } else {
      console.log("Skipping replacement for 'TOTAL' table as no data was extracted.");
    }

    if (success) {
      console.log(`Successfully merged tables into new document: ${mergedFileName} (ID: ${mergedDocId})`);
      return mergedDoc;
    } else {
      console.error(`Failed to merge one or more tables into document: ${mergedFileName} (ID: ${mergedDocId}). Check logs for details.`);
      // Optionally trash the merged document if the merge failed
      DriveApp.getFileById(mergedDocId).setTrashed(true);
      console.log(`Trashed partially merged document ${mergedDocId} due to errors.`);
      return null;
    }
  }

  /**
   * Merges the declarations and the marking grid. 
   * @param {string} name - The student name
   * @param {string} folder - The Google Drive Folder Object
   * @return {Object|null} Object containing information about processed files or null if processing failed
   */
  createFinalDeclarationPDF(name, mergedFileName, folder) {
    console.log(`Starting merge and process workflow for ${name}`);


    // 1. Merge the two documents
    const mergedDeclarationFile = this.mergeDeclarations(mergedFileName);

    if (!mergedDeclarationFile) {
      console.error(`Failed to merge declarations for ${name}. Skipping PDF conversion.`);
      return null;
    }

    // 2. Convert the merged document and marking grid to PDF
    let filesToMerge = [];
    const mergedDeclarationPdf = DriveManager.copyGoogleDocAsPdf(mergedDeclarationFile, folder);
    filesToMerge.push(mergedDeclarationPdf);

    const markingGridPdf = DriveManager.copyGoogleDocAsPdf(this.studentMarkingGridFile, folder);
    filesToMerge.push(markingGridPdf);

    // 4. Merge the PDFs
    const finalMergedPDF = PDFMerger.getInstance().mergePDFs(filesToMerge, mergedFileName, folder);

    return finalMergedPDF;
  }
}

/**
 * Entry point function for processing declarations only
 * Gets user input for the assignment title and processes declarations
 * for all students in the spreadsheet
 */
function processDeclarationsOnly() {
  const ui = SpreadsheetApp.getUi();
  const assignmentTitleResponse = UIManager.promptUser(
    'Enter Google Classroom Assignment Title',
    'Please enter the title of the Google Classroom assignment containing the declarations.'
  );

  if (assignmentTitleResponse.getSelectedButton() === ui.Button.OK) {
    const assignmentTitle = assignmentTitleResponse.getResponseText().trim();


    try {
      // Get the data from the Student Info sheet
      const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Student Info");
      if (!sheet) {
        UIManager.showAlert('Student Info sheet not found. Please run "Get names and IDs" first.');
        return;
      }

      const data = sheet.getDataRange().getValues();

      // Create a separate array with courseId at the beginning
      const courseInfoSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Course Info");
      if (!courseInfoSheet) {
        UIManager.showAlert('Course Info sheet not found. Please run "Get names and IDs" first.');
        return;
      }

      const courseId = courseInfoSheet.getRange(1, 2).getValue();
      const processData = [[courseId]].concat([[""]]).concat(data); // Add courseId and a blank row

      // Process the declarations
      const declarationProcessor = new DeclarationProcessor();
      declarationProcessor.createFinalDeclarationForms(assignmentTitle, processData);

      UIManager.showAlert('Processing declarations completed.');
    } catch (e) {
      UIManager.showAlert(`Error: ${e.message}`);
      console.error(e);
    }
  } else {
    UIManager.showAlert('Operation canceled.');
  }
}

