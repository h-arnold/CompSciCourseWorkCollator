/**
 * Class for user interface management
 */
class UIManager {
    /**
     * Creates and adds the menu to the UI
     */
    static createMenu() {
      const ui = SpreadsheetApp.getUi();
      ui.createMenu('Folder Populator')
        .addItem("1. Get names and IDs", "runScript")
        .addItem("2. Copy marksheets and declarations", "populateFoldersWithTemplates")
        .addItem("3. Copy coursework submissions", "populateFolders")
        .addItem("4. Process declarations only", "processDeclarationsOnly")
        .addItem("5. Merge PDFs for all students", "mergeAllStudentPDFs")
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
