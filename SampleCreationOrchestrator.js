/** Sample Orchestrator
 * This is a global orchestrator for creating the final sample folder.
 * It handles the creation of the parent folder, the student folders (appropriately named)
 * and merging all the PDFs and Declarations.
 *
*/
class SampleCreationOrchestrator {
    constructor() {
        this.uiManager = new UIManager();
    }

    // Main orchestator method
    createSample() {
        // Get the initial parameters needed
        const sampleDestinationFolderId = this.uiManager.promptUser("Sample Destination Folder", 
            "Please enter the destination folder for the samples:");
        const declarationAssignmentName = this.uiManager.promptUser("Declaration Assignment Name",
            "Please enter the name of the declaration assignment:");

        // Create the final declaration sheets for each student.

        studentFolderData = this.processDeclarationSheets();
    }

    processDeclarationSheets() {
        const spreadsheets = SpreadsheetManager.getSpreadsheetSheets();
        const studentSheet = spreadsheets.studentSheet;
        const courseSheet = spreadsheets.courseSheet;
        
        const data = studentSheet.getDataRange().getValues();
              
        const courseId = courseSheet.getRange(1, 2).getValue();
        const processData = [[courseId]].concat([[""]]).concat(data); // Add courseId and a blank row
        
        // Create the final declaration forms
        const declarationProcessor = new DeclarationProcessor();

        // Return the updated student data array with the properly formatted folder names for each student.
        return declarationProcessor.createFinalDeclarationForms(assignmentTitle, processData);
    }


}