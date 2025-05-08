/**
 * Class responsible for text extraction and formatting
 */
class TextProcessor {
  /**
   * Retrieves candidate and center numbers from a Google Doc
   * @param {string|File} fileIdOrFile - The ID of the Drive file or File object
   * @return {Object} An object containing candidate and center numbers
   */
  getCandidateAndCentreNo(fileIdOrFile) {
    const fileId = typeof fileIdOrFile === 'string' ? fileIdOrFile : fileIdOrFile.getId();
    const doc = DocumentApp.openById(fileId);
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
    const formattedSurname = surname.slice(0, 2).toUpperCase();
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
    return `0. Frontsheet_${centreNo}_${candidateNo}_${this.formatStudentName(name)}`;
  }

  /**
   * Finds a table in a document body based on the text in its first cell.
   * @private
   * @param {Body} body - The document body element.
   * @param {string} headerText - The text to match in the first cell (case-insensitive, trimmed).
   * @return {Table|null} The found table element or null.
   */
  _findTableByHeaderText(body, headerText) {
    const tables = body.getTables();
    for (let i = 0; i < tables.length; i++) {
      const table = tables[i];
      if (table.getNumRows() > 0 && table.getRow(0).getNumCells() > 0) {
        const cellText = table.getCell(0, 0).getText().trim();
        // Use startsWith for flexibility, e.g., "Title of Task:"
        if (cellText.toLowerCase().startsWith(headerText.toLowerCase())) {
          return table;
        }
      }
    }
    console.log(`Table starting with "${headerText}" not found.`);
    return null;
  }

  /**
   * Extracts the text content and formatting attributes of a table identified by its header text.
   * @param {string|File} docIdOrFile - The ID of the Google Document or File object
   * @param {string} tableHeaderText - The text identifying the table (e.g., "Title of Task:").
   * @return {Array<Array<{text: string, attributes: object}>>|null} A 2D array of cell data objects, or null if the table isn't found.
   */
  extractTableText(docIdOrFile, tableHeaderText) {
    try {
      const docId = typeof docIdOrFile === 'string' ? docIdOrFile : docIdOrFile.getId();
      const doc = DocumentApp.openById(docId);
      const body = doc.getBody();
      const table = this._findTableByHeaderText(body, tableHeaderText);

      if (!table) {
        return null;
      }

      const data = [];
      for (let r = 0; r < table.getNumRows(); r++) {
        const row = table.getRow(r);
        const rowData = [];
        for (let c = 0; c < row.getNumCells(); c++) {
          const cell = row.getCell(c);
          // Store both text and attributes
          rowData.push({
            text: cell.getText(),
            attributes: cell.getAttributes() 
          });
        }
        data.push(rowData);
      }
      return data;
    } catch (e) {
      console.error(`Error extracting table "${tableHeaderText}" from doc ${typeof docIdOrFile === 'string' ? docIdOrFile : docIdOrFile.getName()}: ${e}`);
      return null;
    }
  }

  /**
   * Replaces the text content and formatting of a table identified by its header text.
   * Assumes the target table has at least the same dimensions as the source data.
   * @param {string|File} docIdOrFile - The ID of the Google Document to modify or File object
   * @param {string} tableHeaderText - The text identifying the table to replace content in.
   * @param {Array<Array<{text: string, attributes: object}>>} cellDataArray - The 2D array of cell data (text and attributes) to insert.
   * @return {boolean} True if successful, false otherwise.
   */
  replaceTableText(docIdOrFile, tableHeaderText, cellDataArray) {
    if (!cellDataArray) {
      console.log(`No cell data provided for table "${tableHeaderText}" in doc ${typeof docIdOrFile === 'string' ? docIdOrFile : docIdOrFile.getName()}. Skipping replacement.`);
      return false;
    }

    try {
      const docId = typeof docIdOrFile === 'string' ? docIdOrFile : docIdOrFile.getId();
      const doc = DocumentApp.openById(docId);
      const body = doc.getBody();
      const table = this._findTableByHeaderText(body, tableHeaderText);

      if (!table) {
        return false;
      }

      for (let r = 0; r < cellDataArray.length; r++) {
        const rowData = cellDataArray[r];
        if (r < table.getNumRows()) {
          const tableRow = table.getRow(r);
          for (let c = 0; c < rowData.length; c++) {
            if (c < tableRow.getNumCells()) {
              const targetCell = tableRow.getCell(c);
              const sourceCellData = rowData[c];

              // Ensure sourceCellData and attributes exist
              if (sourceCellData) {
                // Clear existing content and set new text
                targetCell.clear().setText(sourceCellData.text || ''); // Use empty string if text is null/undefined
                
                // Apply attributes if they exist
                if (sourceCellData.attributes) {
                  // Create a mutable copy of attributes before setting
                  const attributesCopy = {};
                  for (const key in sourceCellData.attributes) {
                    if (sourceCellData.attributes.hasOwnProperty(key)) {
                      attributesCopy[key] = sourceCellData.attributes[key];
                    }
                  }
                  // Remove attributes that might cause issues if null/undefined
                  // (e.g., LINK_URL being null can sometimes cause errors)
                  if (attributesCopy[DocumentApp.Attribute.LINK_URL] === null) {
                    delete attributesCopy[DocumentApp.Attribute.LINK_URL];
                  }
                  targetCell.setAttributes(attributesCopy);
                }
              } else {
                 targetCell.clear(); // Clear cell if no source data
              }
            } else {
              console.warn(`Source data has more columns (${rowData.length}) than target table row ${r} (${tableRow.getNumCells()}) for table "${tableHeaderText}". Data truncated.`);
            }
          }
        } else {
          console.warn(`Source data has more rows (${cellDataArray.length}) than target table (${table.getNumRows()}) for table "${tableHeaderText}". Data truncated.`);
          break; 
        }
      }
      console.log(`Successfully replaced text and formatting in table "${tableHeaderText}" in doc ${docId}.`);
      return true;
    } catch (e) {
      console.error(`Error replacing text/formatting in table "${tableHeaderText}" in doc ${typeof docIdOrFile === 'string' ? docIdOrFile : docIdOrFile.getName()}: ${e} \n Stack: ${e.stack}`);
      return false;
    }
  }
}