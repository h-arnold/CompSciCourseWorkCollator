/**
 * Class responsible for Google Drive file operations
 */
class DriveFileManager {
  /**
   * Copies and renames a Drive file
   * @param {string} driveFileId - The ID of the Drive file
   * @param {string} folderId - The ID of the folder where the file will be copied
   * @param {string} newFileName - The name of the file
   */
  copyAndRenameFile(driveFileId, folderId, newFileName) {
    const driveFile = DriveApp.getFileById(driveFileId);
    const folder = DriveApp.getFolderById(folderId);
    driveFile.makeCopy(newFileName, folder);
  }
  
  /**
   * Converts a Google Docs file to a PDF
   * @param {string} driveFileId - The ID of the Drive file
   * @return {File} The converted PDF file
   */
  convertToPdf(driveFileId) {
    const driveFile = DriveApp.getFileById(driveFileId);
    const blob = driveFile.getAs('application/pdf');
    const pdfFile = DriveApp.createFile(blob);
    return pdfFile;
  }
}

/**
 * Class responsible for text extraction and formatting
 */
class TextProcessor {
  /**
   * Retrieves candidate and center numbers from a Google Doc
   * @param {string} driveFileId - The ID of the Drive file
   * @return {Object} An object containing candidate and center numbers
   */
  getCandidateAndCentreNo(driveFileId) {
    const doc = DocumentApp.openById(driveFileId);
    const body = doc.getBody();
    const text = body.getText();
    const candidateNoMatch = text.match(/\b\d{4}\b/); // Regular expression for 4 digits
    const centreNoMatch = text.match(/\b\d{5}\b/); // Regular expression for 5 digits
    
    console.log("Candidate Number is: " + candidateNoMatch + "\n Centre Number is: " + centreNoMatch);
    
    return {
      CandidateNo: candidateNoMatch ? candidateNoMatch[0] : null,
      CentreNo: centreNoMatch ? centreNoMatch[0] : null
    };
  }
  
  /**
   * Formats a student name
   * @param {string} name - The student name in {firstname} {surname} format
   * @return {string} The formatted string
   */
  formatStudentName(name) {
    const [firstName, surname] = name.split(' ');
    const formattedSurname = surname.slice(0, 2).charAt(0).toUpperCase() + surname.slice(1, 2).toLowerCase();
    const formattedFirstName = firstName.charAt(0).toUpperCase();
    return `${formattedSurname}_${formattedFirstName}`;
  }
  
  /**
   * Creates a standardized file name
   * @param {string} centreNo - The centre number
   * @param {string} candidateNo - The candidate number
   * @param {string} name - The student name
   * @return {string} The formatted file name
   */
  createFileName(centreNo, candidateNo, name) {
    return `${centreNo}_${candidateNo}_${this.formatStudentName(name)}`;
  }
}

/**
 * Main processor class that orchestrates the declaration processing workflow
 */
class DeclarationProcessor {
  constructor() {
    this.driveManager = new DriveFileManager();
    this.textProcessor = new TextProcessor();
    // ClassroomManager is now a global class with static methods
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
      const pdfFile = this.driveManager.convertToPdf(driveFileId);
      const pdfFileId = pdfFile.getId();
      const newFileName = this.textProcessor.createFileName(CentreNo, CandidateNo, name);
      this.driveManager.copyAndRenameFile(pdfFileId, folderId, newFileName);
      console.log(`${name}'s document has been moved and converted to PDF.`);
    }
  }
}

/**
 * Entry point function to maintain backward compatibility
 * @param {string} assignmentTitle - The title of the Google Classroom assignment
 * @param {Array[]} data - The data from the spreadsheet, including folderIds
 */
function processFolderAttachmentsForDeclarationsOnly(assignmentTitle, data) {
  const processor = new DeclarationProcessor();
  processor.processFolderAttachmentsForDeclarationsOnly(assignmentTitle, data);
}
