/**
 * Main processor class that orchestrates the declaration processing workflow
 */
class DeclarationProcessor {
  constructor() {
    this.textProcessor = new TextProcessor();
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
      const studentFolderDeclarationFile = DriveManager.findFilesBySubstring(
        folderId, 
        "Declaration", 
        false, 
        "application/vnd.google-apps.document",
        "suffix");
        
      const studentMarkingGridFile = DriveManager.findFilesBySubstring(      
        folderId, 
        "Marking Grid", 
        false, 
        "application/vnd.google-apps.document",
        "suffix");

      const studentFolderDeclarationFileId = studentFolderDeclarationFile.id
      const studentMarkingGridFileId = studentFolderDeclarationFileId.id
      
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
              const gClassroomDeclarationFileId = attachment.driveFile.id


              // If base declaration provided, merge first then process as PDF
              this.mergeAndProcessDeclarations(
                gClassroomDeclarationFileId, 
                studentFolderDeclarationFileId,
                folderId, 
                name
              );
            } else {
              // No merging needed, just process the original attachment
              this.createFinalDeclarationPDF(attachment.driveFile.id, folderId, name);
              return; //No need to continue the loop as there should only be one declaration.
            }
          }
        });
      });
    });
    


  }
  
  /**
   * Creates A SINGLE final declaration PDF which merges:
   *   - The signed declaration form with no marks on Google Classroom
   *   - The unsigned declaration form in the student folder with marks.
   *   - The marking grid
   * Then converts it into a final PDF and renames it according to the 
   * WJEC required convension which is:
   * {centreNumber}_{candidateNumber}_{firstInitial}_{firstTwoInitialOfSurname}
   * @param {string} driveFileId - The Drive file ID
   * @param {string} folderId - The folder ID
   * @param {string} name - The student name
   * @return {Object|null} Object containing information about processed files or null if processing failed
   */
  createFinalDeclarationPDF(driveFileId, folderId, name) {
    const { CandidateNo, CentreNo } = this.textProcessor.getCandidateAndCentreNo(driveFileId);

    if (CandidateNo && CentreNo) {
      console.log(`Candidate number ${CandidateNo} and Centre number ${CentreNo} found in document.`);
      // Convert the original Google Doc to PDF first
      const pdfBlob = DriveApp.getFileById(driveFileId).getAs('application/pdf');
      const newFileName = this.textProcessor.createFileName(CentreNo, CandidateNo, name);
      const pdfFileName = `${newFileName}.pdf`; // Ensure PDF extension
      const folder = DriveApp.getFolderById(folderId);

      // Check if PDF file with same name already exists
      if (DriveManager.checkAndHandleExistingFile(folder, pdfFileName)) {
          const pdfFile = folder.createFile(pdfBlob).setName(pdfFileName);
          console.log(`${name}'s document (ID: ${driveFileId}) converted to PDF "${pdfFileName}" (ID: ${pdfFile.getId()}) in folder ${folderId}.`);
          return {
            pdfFileId: pdfFile.getId(),
            pdfFileName: pdfFileName,
            studentName: name,
            candidateNo: CandidateNo,
            centreNo: CentreNo
          };
      } else {
          console.log(`Skipped creating PDF "${pdfFileName}" for ${name} as user chose not to replace existing file.`);
          return null;
      }
    } else {
        console.log(`Could not find Candidate/Centre number in document ID ${driveFileId} for ${name}. Skipping PDF conversion.`);
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
  mergeDeclarations(gClassroomDeclarationFileId, studentFolderDeclarationFileId, mergedFileName) {
    console.log(`Starting merge process: Google Classroom Doc ID: ${gClassroomDeclarationFileId}, Student Folder Doc ID: ${studentFolderDeclarationFileId}, New Filename: ${mergedFileName}`);

    // Determine the destination folder (parent of the source document)
    let destinationFolderId = null;
    try {
        const studentFolderDeclarationFile = DriveApp.getFileById(studentFolderDeclarationFileId);
        const parents = studentFolderDeclarationFile.getParents();
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
    const mergedDoc = DriveManager.copyDocument(gClassroomDeclarationFileId, mergedFileName, destinationFolderId);
    if (!mergedDoc) {
      console.error("Failed to create a copy of the base document. Aborting merge.");
      return null;
    }
    const mergedDocId = mergedDoc.getId();
    console.log(`Created copy of base document with ID: ${mergedDocId}`);

    // 2. Extract data from source document tables
    const titleTableData = this.textProcessor.extractTableText(studentFolderDeclarationFileId, "Title of Task:");
    const totalTableData = this.textProcessor.extractTableText(studentFolderDeclarationFileId, "TOTAL");

    if (!titleTableData) {
        console.warn(`Could not extract 'Title of Task:' table data from source doc ${studentFolderDeclarationFileId}.`);
    }
    if (!totalTableData) {
        console.warn(`Could not extract 'TOTAL' table data from source doc ${studentFolderDeclarationFileId}.`);
    }

    // 3. Replace data in the new (merged) document tables
    let success = true;
    if (titleTableData) {
        success = this.textProcessor.replaceTableText(mergedDocId, "Title of Task:", titleTableData) && success;
    } else {
        console.log("Skipping replacement for 'Title of Task:' table as no data was extracted.");
    }

    if (totalTableData) {
        success = this.textProcessor.replaceTableText(mergedDocId, "TOTAL", totalTableData) && success;
    } else {
        console.log("Skipping replacement for 'TOTAL' table as no data was extracted.");
    }


    if (success) {
      console.log(`Successfully merged tables into new document: ${mergedFileName} (ID: ${mergedDocId})`);
      // Optional: Clean up studentFolderDeclarationFileId or gClassroomDeclarationFileId if needed (e.g., trash them)
      return mergedDocId;
    } else {
      console.error(`Failed to merge one or more tables into document: ${mergedFileName} (ID: ${mergedDocId}). Check logs for details.`);
      // Consider trashing the partially merged document to avoid confusion
      // DriveApp.getFileById(mergedDocId).setTrashed(true);
      // console.log(`Trashed partially merged document ${mergedDocId} due to errors.`);
      return null;
    }
  }

  /**
   * Merges two declaration documents and then processes the merged document as a PDF
   * @param {string} baseDeclarationFileId - The ID of the base declaration Google Doc
   * @param {string} studentDeclarationFileId - The ID of the student's declaration Google Doc
   * @param {string} folderId - The folder ID to save the resulting files
   * @param {string} name - The student name
   * @return {Object|null} Object containing information about processed files or null if processing failed
   */
  mergeAndProcessDeclarations(baseDeclarationFileId, studentDeclarationFileId, folderId, name) {
    console.log(`Starting merge and process workflow for ${name}`);
    
    // 1. Generate a name for the merged document
    const mergedFileName = `Merged_Declaration_${name}`;
    
    // 2. Merge the two documents
    const mergedDocId = this.mergeDeclarations(
      baseDeclarationFileId, 
      studentDeclarationFileId, 
      mergedFileName
    );
    
    if (!mergedDocId) {
      console.error(`Failed to merge declarations for ${name}. Skipping PDF conversion.`);
      return null;
    }
    
    // 3. Process the merged document as a PDF
    console.log(`Successfully merged, now processing as PDF: ${mergedFileName} (ID: ${mergedDocId})`);
    return this.createFinalDeclarationPDF(mergedDocId, folderId, name);
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

