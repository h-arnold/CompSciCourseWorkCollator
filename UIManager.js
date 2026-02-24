/**
 * Class for user interface management
 */
class UIManager {
    /**
     * Creates and adds the menu to the UI
     */
    static createMenu() {
      const ui = SpreadsheetApp.getUi();
      const computerScienceMenu = ui.createMenu('Computer Science')
        .addItem("Setup class and student folders", "runScript")
        .addItem("Copy marksheets and declarations", "populateFoldersWithTemplates")
        .addItem("Copy coursework submissions", "populateFolders")
        .addItem("Process declarations only", "processDeclarationsOnly")
        .addItem("Merge PDFs for all students", "mergeAllStudentPDFs");

      const ictMenu = ui.createMenu('ICT Content Creation')
        .addItem("Setup Student Info headers", "setupICTStudentInfoSheet")
        .addItem("Fetch roster to Student Info", "fetchICTRosterToSheet")
        .addItem("Create folders from Student Info", "createICTFoldersFromSheet")
        .addItem("Collate assignment content", "runICTContentCreationCollation");

      ui.createMenu('Coursework Collator')
        .addSubMenu(computerScienceMenu)
        .addSubMenu(ictMenu)
        .addToUi();
    }
    
    /**
     * Shows a prompt dialog and returns user response
     * @param {string} title - The title of the prompt
     * @param {string} message - The message to display
     * @returns {Object} User's response with selected button and response text
     */
    static promptUser(title, message) {
      const ui = SpreadsheetApp.getUi();
      return ui.prompt(title, message, ui.ButtonSet.OK_CANCEL);
    }
    
    /**
     * Shows an alert dialog
     * @param {string} message - The message to display
     */
    static showAlert(message) {
      const ui = SpreadsheetApp.getUi();
      ui.alert(message);
    }
  }
