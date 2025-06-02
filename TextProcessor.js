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
  createFileName(StudentSubmissionPrefix) {
    return `0. Frontsheet_${StudentSubmissionPrefix}`;
  }

  /**
   * Creates a standardized student submission prefix without the frontsheet prefix
   * @param {string} centreNo - The centre number
   * @param {string} candidateNo - The candidate number
   * @param {string} name - The student name
   * @return {string} The formatted submission prefix
   */
  createStudentSubmissionPrefix(centreNo, candidateNo, name) {
    return `${centreNo}_${candidateNo}_${this.formatStudentName(name)}`;
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

      // Determine the maximum number of columns in the table
      let maxColumns = 0;
      for (let r = 0; r < table.getNumRows(); r++) {
        const row = table.getRow(r);
        maxColumns = Math.max(maxColumns, row.getNumCells());
      }
      
      // Special handling for table headers to detect column count
      const headerRow = table.getRow(0);
      // If header names indicate more columns than physically present, use that higher count
      // This helps with merged cells in the header
      if (headerRow.getNumCells() > 0 && headerRow.getCell(headerRow.getNumCells() - 1).getText().trim() === "CENTRE COMMENTS") {
        // We found the "CENTRE COMMENTS" cell, ensure we have at least 4 columns
        maxColumns = Math.max(maxColumns, 4);
      }

      const data = [];
      for (let r = 0; r < table.getNumRows(); r++) {
        const row = table.getRow(r);
        const rowData = [];
        
        // Process each cell in the row up to the maximum columns
        for (let c = 0; c < maxColumns; c++) {
          // Check if this column exists in this row
          if (c < row.getNumCells()) {
            const cell = row.getCell(c);
            rowData.push({
              text: cell.getText(),
              attributes: cell.getAttributes() 
            });
          } else {
            // If this column doesn't exist in this row (due to merging),
            // push an empty cell placeholder
            rowData.push({
              text: "",
              attributes: {}
            });
          }
        }
        data.push(rowData);
      }
      
      console.log(`Extracted table with ${data.length} rows and up to ${maxColumns} columns`);
      return data;
    } catch (e) {
      console.error(`Error extracting table "${tableHeaderText}" from doc ${typeof docIdOrFile === 'string' ? docIdOrFile : docIdOrFile.getName()}: ${e}`);
      return null;
    }
  }

  /**
   * Replaces the text content and formatting of a table identified by its header text.
   * Handles tables with merged cells by ensuring all data is properly applied.
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

      // Determine the column count from the source data
      const maxSourceColumns = cellDataArray.reduce((max, row) => 
        Math.max(max, row.length), 0);
      
      // Process each row in the source data
      for (let r = 0; r < cellDataArray.length; r++) {
        const rowData = cellDataArray[r];
        
        // Ensure target table has enough rows
        let tableRow;
        if (r < table.getNumRows()) {
          tableRow = table.getRow(r);
        } else {
          console.warn(`Target table has fewer rows (${table.getNumRows()}) than source data (${cellDataArray.length}). Skipping extra rows.`);
          break;
        }
        
        // Ensure target row has enough cells by appending if needed
        const currentCellCount = tableRow.getNumCells();
        if (currentCellCount < rowData.length) {
          console.log(`Row ${r} has fewer cells (${currentCellCount}) than source data (${rowData.length}). This might be due to merged cells.`);
        }
        
        // Process each cell in the row
        for (let c = 0; c < rowData.length; c++) {
          // Check if this column exists in this row
          let targetCell;
          if (c < tableRow.getNumCells()) {
            targetCell = tableRow.getCell(c);
            const sourceCellData = rowData[c];

            // Apply source data if it exists
            if (sourceCellData) {
              // Clear existing content and set new text
              targetCell.clear().setText(sourceCellData.text || '');
              
              // Apply attributes if they exist
              if (sourceCellData.attributes) {
                // Create a mutable copy of attributes
                const attributesCopy = {};
                for (const key in sourceCellData.attributes) {
                  if (sourceCellData.attributes.hasOwnProperty(key)) {
                    attributesCopy[key] = sourceCellData.attributes[key];
                  }
                }
                
                // Remove attributes that might cause issues
                if (attributesCopy[DocumentApp.Attribute.LINK_URL] === null) {
                  delete attributesCopy[DocumentApp.Attribute.LINK_URL];
                }
                
                targetCell.setAttributes(attributesCopy);
              }
            } else {
              targetCell.clear(); // Clear cell if no source data
            }
          } else {
            // We have data for a column that doesn't exist in the target row
            // This can happen with merged cells or if source has more columns
            console.warn(`Cannot apply data for column ${c} in row ${r} as it doesn't exist in target table. This might be due to merged cells.`);
          }
        }
      }
      
      console.log(`Successfully replaced text and formatting in table "${tableHeaderText}" in doc ${docId}. Applied data for ${cellDataArray.length} rows with up to ${maxSourceColumns} columns.`);
      return true;
    } catch (e) {
      console.error(`Error replacing text/formatting in table "${tableHeaderText}" in doc ${typeof docIdOrFile === 'string' ? docIdOrFile : docIdOrFile.getName()}: ${e} \n Stack: ${e.stack}`);
      return false;
    }
  }
}