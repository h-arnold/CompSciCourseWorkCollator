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
   * Creates the final declaration forms for each student in the spreadsheet by merging:
   *   - the signed declaration on Google Classroom (which doesn't have marks)
   *   - the unsigned declaration in the student folder (which has marks)
   *    - the marking grid
   * Then converts it into a final PDF and renames it according to the
   * WJEC required convention which is:
   * {centreNumber}_{candidateNumber}_{firstInitial}_{firstTwoInitialOfSurname}
   * 
   * @param {string} assignmentTitle - The title of the Google Classroom assignment
   * @param {Array[]} data - The data from the spreadsheet, including folderIds
   * @param {string} [gClassroomDeclarationFileId] - Optional ID of the base declaration file to merge with
   */
  createFinalDeclarationForms(assignmentTitle, data) {
    const courseId = data[0][0]; // Get the courseId from the first row
    const assignmentId = ClassroomManager.getAssignmentId(courseId, assignmentTitle);
    
    if (!assignmentId) {
      SpreadsheetApp.getUi().alert('Invalid Google Classroom Assignment title or access denied.');
      return;
    }
    
    // Track processed files for potential merging if requested
    const processedFiles = [];
    
    data.forEach((row, index) => {
      if (index < 3) return; // Skip the header rows
      
      const name = row[0];
      const userId = row[1];
      const folderId = row[2];
      
      
      if (!folderId) {
        console.log(`No folder ID found for user ${userId}`);
        return;
      }

      // Get the declaration and marking grid files
      this.studentFolderDeclarationFile = DriveManager.findFilesBySubstring(
        folderId, 
        "Declaration", 
        false, 
        "application/vnd.google-apps.document",
        "suffix");
        
      this.studentMarkingGridFile = DriveManager.findFilesBySubstring(      
        folderId, 
        "Marking Grid", 
        false, 
        "application/vnd.google-apps.document",
        "suffix");

      const submissions = ClassroomManager.getStudentSubmissions(courseId, assignmentId, userId);
      
      submissions.forEach(submission => {
        const attachments = submission.assignmentSubmission.attachments || [];
        attachments.forEach(attachment => {
          if (attachment.driveFile) {

            //Check that the attachment is a Google Drive File
            const file = DriveApp.getFileById(attachment.driveFile.id)

            // Check if the file is a Google Doc
            if (DriveManager.isGoogleDoc(file)) {
              // This attachment is most likely the declaration unless the student has done something strage
              this.gClassroomDeclarationFile = file;

              // Now that we have all the files we need, create the final PDF.
              this.createFinalDeclarationPDF(
                name, 
                folderId
              );

              return; // No need to continue the loop as there should only be one declaration.
            }
          }
        });
      });
    });
    


  }
  
  /**
   * Generates a filename for a declaration document according to the 
   * WJEC required convention which is:
   * {centreNumber}_{candidateNumber}_{firstInitial}_{firstTwoInitialOfSurname}
   * @param {string|File} fileIdOrFile - The Drive file ID or File object to extract candidate/center numbers from
   * @param {string} name - The student name
   * @return {string|null} The generated filename or null if required information not found
   */
  generateDeclarationFileName(name) {
    const { CandidateNo, CentreNo } = this.textProcessor.getCandidateAndCentreNo(this.gClassroomDeclarationFile)
  
    if (CandidateNo && CentreNo) {
      console.log(`Candidate number ${CandidateNo} and Centre number ${CentreNo} found in document.`);
      const studentSubmissionPrefix = this.textProcessor.createStudentSubmissionPrefix(CentreNo, CandidateNo, name);
      const newFileName = this.textProcessor.createFileName(studentSubmissionPrefix);
      console.log(`Generated filename for ${name}: ${newFileName}`);
      return newFileName;
    } else {
      const fileId = typeof fileIdOrFile === 'string' ? fileIdOrFile : fileIdOrFile.getId();
      console.log(`Could not find Candidate/Centre number in document ID ${fileId} for ${name}. Cannot generate filename.`);
      return null;
    }
  }

  /**
   * Merges table data from a source document into a copy of a base document.
   * Specifically targets 'Title of Task:' and 'TOTAL' tables.
   * The copy of the base document is placed in the same folder as the source document.
   * @param {string} gClassroomDeclarationFileId - The ID of the Google Doc to copy and merge into.
   * @param {string} studentFolderDeclarationFileId - The ID of the Google Doc containing the table data to merge.
   * @param {string} mergedFileName - The desired name for the newly created merged document.
   * @return {string|null} The ID of the newly created merged document, or null on failure.
   */
  mergeDeclarations(mergedFileName) {
    console.log(`Starting merge process: Google Classroom Doc ID: ${this.gClassroomDeclarationFile.getId()}, Student Folder Doc ID: ${this.studentFolderDeclarationFile.getId()}, New Filename: ${mergedFileName}`);

    // Determine the destination folder (parent of the source document)
    let destinationFolderId = null;
    try {
        const parents = this.studentFolderDeclarationFile.getParents();
        if (parents.hasNext()) {
            destinationFolderId = parents.next().getId();
            console.log(`Target destination folder ID (from source doc parent): ${destinationFolderId}`);
        } else {
            console.warn(`Source document ${studentFolderDeclarationFileId} has no parent folder. Copy will be placed relative to base document or in root.`);
        }
    } catch (e) {
        console.error(`Error getting parent folder for source document ${studentFolderDeclarationFileId}: ${e}. Copy will be placed relative to base document or in root.`);
    }


    // 1. Create a copy of the base document in the source document's folder
    const mergedDoc = DriveManager.copyDocument(this.gClassroomDeclarationFile, mergedFileName, destinationFolderId);
    if (!mergedDoc) {
      console.error("Failed to create a copy of the base document. Aborting merge.");
      return null;
    }
    const mergedDocId = mergedDoc.getId();
    console.log(`Created copy of base document with ID: ${mergedDocId}`);

    // 2. Extract data from source document tables
    const titleTableData = this.textProcessor.extractTableText(this.studentFolderDeclarationFile, "Title of Task:");
    const totalTableData = this.textProcessor.extractTableText(this.studentFolderDeclarationFile, "TOTAL");

    if (!titleTableData) {
        console.warn(`Could not extract 'Title of Task:' table data from source doc ${studentFolderDeclarationFileId.getId()}.`);
    }
    if (!totalTableData) {
        console.warn(`Could not extract 'TOTAL' table data from source doc ${studentFolderDeclarationFileId.getId()}.`);
    }

    // 3. Replace data in the new (merged) document tables
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
      // Optional: Clean up studentFolderDeclarationFileId or gClassroomDeclarationFileId if needed (e.g., trash them)
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
   * @param {string} folderId - The folder ID
   * @return {Object|null} Object containing information about processed files or null if processing failed
   */
  createFinalDeclarationPDF(name) {
    console.log(`Starting merge and process workflow for ${name}`);
    
    // 1. Generate a name for the merged document
    const mergedFileName = this.generateDeclarationFileName(name);
    if (!mergedFileName) {
      console.error(`Failed to generate filename for ${name}. Skipping PDF conversion.`);
      return null;
    }
    
    // 2. Merge the two documents
    const mergedDeclarationFile = this.mergeDeclarations(mergedFileName);
  
    if (!mergedDeclarationFile) {
      console.error(`Failed to merge declarations for ${name}. Skipping PDF conversion.`);
      return null;
    }
  
    // 3. Convert the merged document and marking grid to PDF
    let filesToMerge = [];
    const mergedDeclarationPdf = DriveManager.convertToPdf(mergedDeclarationFile);
    filesToMerge.push(mergedDeclarationPdf);
    
    const markingGridPdf = DriveManager.convertToPdf(this.studentMarkingGridFile);
    filesToMerge.push(markingGridPdf);
  
    // 4. Merge the PDFs
    const finalMergedPDF = PDFMerger.mergePDFs(filesToMerge, markingGridPdf);
  
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

