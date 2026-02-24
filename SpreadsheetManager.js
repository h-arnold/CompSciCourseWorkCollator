const ICT_STUDENT_HEADERS = ["Student ID", "Name", "User ID", "Folder ID"];

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
    return { studentSheet, courseSheet, prefixSheet };
  }

  /**
   * Reads ICT student rows from the provided sheet.
   * Expected headers (case-insensitive):
   * - Student ID (required)
   * - Name (required)
   * - User ID (required for Classroom cross-check)
   * - Folder ID (optional, will be written if missing)
   * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - The sheet containing ICT student data
   * @returns {Object[]} Normalized student records with row index metadata
   */
  static getICTStudentsFromSheet(sheet) {
    const values = sheet.getDataRange().getValues();
    if (!values || values.length < 2) {
      return [];
    }

    const normalizeHeader = header => String(header || '').trim().toLowerCase();
    const headers = values[0].map(normalizeHeader);

    const findHeaderIndex = aliases => headers.findIndex(header => aliases.includes(header));

    const studentIdIndex = findHeaderIndex(['student id', 'studentid', 'id']);
    const nameIndex = findHeaderIndex(['name', 'student name', 'full name']);
    const userIdIndex = findHeaderIndex(['user id', 'userid', 'classroom user id', 'classroom id']);
    let folderIdIndex = findHeaderIndex(['folder id', 'folderid']);

    if (studentIdIndex === -1 || nameIndex === -1 || userIdIndex === -1) {
      throw new Error('Student Info sheet must include headers: Student ID, Name, and User ID.');
    }

    if (folderIdIndex === -1) {
      folderIdIndex = headers.length;
      sheet.getRange(1, folderIdIndex + 1).setValue('Folder ID');
    }

    const students = [];
    values.slice(1).forEach((row, offset) => {
      const rowIndex = offset + 2;
      const studentId = String(row[studentIdIndex] || '').trim();
      const name = String(row[nameIndex] || '').trim();
      const userId = String(row[userIdIndex] || '').trim();
      const folderId = String(row[folderIdIndex] || '').trim();

      if (!studentId && !name && !userId) {
        return;
      }

      students.push({
        rowIndex,
        studentId,
        name,
        userId,
        folderId,
        folderIdColumn: folderIdIndex + 1
      });
    });

    return students;
  }

  /**
   * Writes folder IDs back to the Student Info sheet for ICT rows.
   * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - The Student Info sheet
   * @param {Object[]} students - Normalized student records
   */
  static updateICTFolderIds(sheet, students) {
    if (!students || students.length === 0) {
      return;
    }

    students.forEach(student => {
      if (student.folderId && student.rowIndex && student.folderIdColumn) {
        sheet.getRange(student.rowIndex, student.folderIdColumn).setValue(student.folderId);
      }
    });
  }

  /**
   * Initializes Student Info sheet with ICT headers.
   * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - The Student Info sheet
   * @param {boolean} clearExisting - Whether to clear the sheet first
   */
  static initializeICTStudentInfoSheet(sheet, clearExisting = false) {
    if (clearExisting) {
      sheet.clear();
    }

    sheet.getRange(1, 1, 1, this.ICT_STUDENT_HEADERS.length).setValues([this.ICT_STUDENT_HEADERS]);
  }
}