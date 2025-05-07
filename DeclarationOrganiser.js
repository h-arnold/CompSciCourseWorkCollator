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
   * Process folder attachments for declarations
   * @param {string} assignmentTitle - The title of the Google Classroom assignment
   * @param {Array[]} data - The data from the spreadsheet, including folderIds
   */
  processFolderAttachmentsForDeclarationsOnly(assignmentTitle, data) {
    const courseId = data[0][0]; // Get the courseId from the first row
    const assignmentId = ClassroomManager.getAssignmentId(courseId, assignmentTitle);
    
    if (!assignmentId) {
      SpreadsheetApp.getUi().alert('Invalid Google Classroom Assignment title or access denied.');
      return;
    }
    
    data.forEach((row, index) => {
      if (index < 2) return; // Skip the header rows
      
      const [name, userId, folderId] = row;
      
      if (!folderId) {
        console.log(`No folder ID found for user ${userId}`);
        return;
      }
      
      const submissions = ClassroomManager.getStudentSubmissions(courseId, assignmentId, userId);
      
      submissions.forEach(submission => {
        const attachments = submission.assignmentSubmission.attachments || [];
        attachments.forEach(attachment => {
          if (attachment.driveFile) {
            this.processAttachment(attachment.driveFile.id, folderId, name);
          }
        });
      });
    });
  }
  
  /**
   * Process a single attachment
   * @param {string} driveFileId - The Drive file ID
   * @param {string} folderId - The folder ID
   * @param {string} name - The student name
   */
  processAttachment(driveFileId, folderId, name) {
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
      } else {
          console.log(`Skipped creating PDF "${pdfFileName}" for ${name} as user chose not to replace existing file.`);
      }
    } else {
        console.log(`Could not find Candidate/Centre number in document ID ${driveFileId} for ${name}. Skipping PDF conversion.`);
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
    console.log(`Starting merge process: Base Doc ID: ${gClassroomDeclarationFileId}, Source Doc ID: ${studentFolderDeclarationFileId}, New Filename: ${mergedFileName}`);

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
}

