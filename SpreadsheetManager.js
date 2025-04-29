/**
 * Class for spreadsheet operations
 */
class SpreadsheetManager {
  /**
   * Writes the member data to a Google Sheet
   * @param {Object[]} members - An array of objects containing member information
   * @param {Object} sheet - The Google Sheet to write the data to
   * @param {string} rootFolderId - The ID of the root folder on Google Drive where student folders will be created
   */
  static writeMembersToSheet(members, sheet, rootFolderId) {
    const headers = ["Name", "User ID", "Folder ID"];
    sheet.appendRow(headers);
  
    const rootFolder = DriveApp.getFolderById(rootFolderId);
    members.forEach(member => {
      const folder = DriveManager.createFolder(rootFolder, member.name);
      const folderId = folder.getId();
      const row = [member.name, member.userId, folderId];
      sheet.appendRow(row);
    });
  }
  
  /**
   * Gets the active spreadsheet and ensures required sheets exist
   * @returns {Object} Object containing spreadsheet sheets
   */
  static getSpreadsheetSheets() {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const studentSheet = spreadsheet.getSheetByName("Student Info") || spreadsheet.insertSheet("Student Info");
    const courseSheet = spreadsheet.getSheetByName("Course Info") || spreadsheet.insertSheet("Course Info");
    const prefixSheet = spreadsheet.getSheetByName("Prefixes") || spreadsheet.insertSheet("Prefixes");
    return { studentSheet, courseSheet };
  }
}