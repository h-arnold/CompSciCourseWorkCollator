/**
 * Class for handling PDF merge operations
 */
class PDFMerger {
  /**
   * Load the PDF-lib library from CDN
   * Docs for the library can be found here:
   * https://pdf-lib.js.org/
   * @returns {Promise<Object>} Promise that resolves with the PDFLib object
   */
  static async loadPdfLib() {
    const cdnUrl = "https://cdn.jsdelivr.net/npm/pdf-lib/dist/pdf-lib.min.js";
    const response = UrlFetchApp.fetch(cdnUrl);
    const content = response.getContentText();
    
    // Create a new context to evaluate the library in
    const context = {};
    
    // Evaluate the library code in the context
    try {
      // Define setTimeout for compatibility with pdf-lib
      context.setTimeout = function(f, t) {
        Utilities.sleep(t);
        return f();
      };
      
      // Use Function constructor instead of eval for better scoping
      const setupLibrary = new Function('context', `
        with(context) {
          ${content}
          return { PDFLib: PDFLib };
        }
      `);
      
      return setupLibrary(context);
    } catch(e) {
      console.error('Error loading PDF-lib:', e);
      throw new Error(`Failed to load PDF-lib: ${e.message}`);
    }
  }
  
  /**
   * Validates if all provided file IDs represent PDF files
   * @param {string[]} fileIds - Array of Google Drive file IDs
   * @returns {Object} Object containing valid and invalid files
   */
  static validateFiles(fileIds) {
    const validFiles = [];
    const invalidFiles = [];
    
    fileIds.forEach(fileId => {
      try {
        const file = DriveApp.getFileById(fileId);
        if (DriveManager.isPdf(file)) {
          validFiles.push(file);
        } else {
          invalidFiles.push({
            id: fileId, 
            name: file.getName(), 
            type: file.getMimeType()
          });
        }
      } catch (e) {
        invalidFiles.push({
          id: fileId, 
          error: e.message
        });
      }
    });
    
    return { validFiles, invalidFiles };
  }
  
  /**
   * Handles the case when only one PDF file is found - copies it instead of merging
   * @param {File} file - The single PDF file to copy
   * @param {string} outputFileName - Name for the output PDF file
   * @param {string} [outputFolderId=null] - Optional folder ID to save the PDF (if null, saves to root)
   * @returns {Promise<Object>} Object with status and result information
   */
  static async copySinglePdfFile(file, outputFileName, outputFolderId = null) {
    console.log(`Only one valid PDF found, copying instead of merging: ${file.getName()}`);
    let newFile;
    
    if (outputFolderId) {
      // Use the DriveManager helper to copy and rename
      const fileId = file.getId();
      DriveManager.copyAndRenameFile(fileId, outputFolderId, outputFileName);
      // Get the newly created file
      const folder = DriveApp.getFolderById(outputFolderId);
      const newFiles = folder.getFilesByName(outputFileName);
      if (newFiles.hasNext()) {
        newFile = newFiles.next();
      }
    } else {
      // If no folder specified, copy to root
      newFile = file.makeCopy(outputFileName);
    }
    
    if (newFile) {
      return {
        success: true,
        message: `Successfully copied the PDF file (skipped merging as only one file was found)`,
        file: {
          id: newFile.getId(),
          name: newFile.getName(),
          url: newFile.getUrl()
        }
      };
    }
    
    return {
      success: false,
      message: `Failed to copy the PDF file: ${file.getName()}`
    };
  }
  
  /**
   * Merges multiple PDF files into a single PDF document
   * @param {File[]} files - Array of PDF files to merge
   * @returns {Promise<Uint8Array>} PDF document bytes
   */
  static async mergeMultiplePdfFiles(files) {
    // Load PDF-lib library
    const lib = await this.loadPdfLib();
    const { PDFLib } = lib;
    
    // Create a new PDF document
    const pdfDoc = await PDFLib.PDFDocument.create();
    
    // Add each valid PDF to the merged document
    for (const file of files) {
      console.log(`Processing ${file.getName()} (${file.getMimeType()})`);
      try {
        const pdfData = await PDFLib.PDFDocument.load(new Uint8Array(file.getBlob().getBytes()));
        const pageIndices = [...Array(pdfData.getPageCount())].map((_, i) => i);
        const pages = await pdfDoc.copyPages(pdfData, pageIndices);
        pages.forEach(page => pdfDoc.addPage(page));
      } catch (e) {
        console.error(`Error processing ${file.getName()}: ${e.message}`);
      }
    }
    
    // Save the document
    return await pdfDoc.save();
  }
  
  /**
   * Saves a PDF document to Google Drive
   * @param {Uint8Array} pdfBytes - The PDF document as bytes
   * @param {string} outputFileName - Name for the output PDF file
   * @param {string} [outputFolderId=null] - Optional folder ID to save the PDF (if null, saves to root)
   * @returns {Object} Object with file information
   */
  static saveResultingPdf(pdfBytes, outputFileName, outputFolderId = null) {
    // Create the PDF file blob
    const blob = Utilities.newBlob([...new Int8Array(pdfBytes)], MimeType.PDF, outputFileName);
    let newFile;
    
    if (outputFolderId) {
      // Save to specified folder
      const folder = DriveApp.getFolderById(outputFolderId);
      newFile = folder.createFile(blob);
    } else {
      // Save to root
      newFile = DriveApp.createFile(blob);
    }
    
    return {
      id: newFile.getId(),
      name: newFile.getName(),
      url: newFile.getUrl()
    };
  }
  
  /**
   * Merges multiple PDF files into a single PDF
   * @param {string[]} fileIds - Array of Google Drive file IDs of PDF files to merge
   * @param {string} outputFileName - Name for the merged PDF file (default: "Merged.pdf")
   * @param {string} [outputFolderId=null] - Optional folder ID to save the merged PDF (if null, saves to root)
   * @returns {Promise<Object>} Object with status and result information
   */
  static async mergePDFs(fileIds, outputFileName = "Merged.pdf", outputFolderId = null) {
    try {
      // Validate files first
      const { validFiles, invalidFiles } = this.validateFiles(fileIds);
      
      if (validFiles.length === 0) {
        return {
          success: false,
          message: "No valid PDF files found to merge",
          invalidFiles
        };
      }
      
      // If there's only one valid file, simply copy it with the new name
      if (validFiles.length === 1) {
        const result = await this.copySinglePdfFile(validFiles[0], outputFileName, outputFolderId);
        if (invalidFiles.length > 0) {
          result.invalidFiles = invalidFiles;
        }
        return result;
      }
      
      // Multiple files - merge them
      const pdfBytes = await this.mergeMultiplePdfFiles(validFiles);
      
      // Save the merged PDF to Drive
      const fileInfo = this.saveResultingPdf(pdfBytes, outputFileName, outputFolderId);
      
      return {
        success: true,
        message: `Successfully merged ${validFiles.length} PDFs`,
        file: fileInfo,
        invalidFiles: invalidFiles.length > 0 ? invalidFiles : null
      };
    } catch (e) {
      return {
        success: false,
        message: `Error merging PDFs: ${e.message}`
      };
    }
  }
  
  /**
   * Gets prefixes from the Google Sheet and merges PDFs based on column groups
   * @param {string} sourceFolderId - ID of the folder containing the PDFs to merge
   * @param {string} [outputFolderId=null] - Optional folder ID to save the merged PDFs (if null, saves to root)
   * @param {boolean} [recursive=false] - Whether to search in subfolders recursively
   * @returns {Promise<Object>} Object with results of the merge operations
   */
  static async mergePDFsFromPrefixSheet(sourceFolderId, outputFolderId = null, recursive = false) {
    try {
      // Get the Prefixes sheet
      const { prefixSheet } = SpreadsheetManager.getSpreadsheetSheets();
      if (!prefixSheet) {
        return {
          success: false,
          message: "Prefixes sheet not found."
        };
      }
      
      // Get all data from the sheet
      const data = prefixSheet.getDataRange().getValues();
      if (data.length < 2) {
        return {
          success: false,
          message: "Prefixes sheet is empty or has only headers."
        };
      }
      
      const headers = data[0];
      const results = [];
      
      // Process each column (category)
      for (let col = 0; col < headers.length; col++) {
        if (!headers[col]) continue; // Skip columns with no header
        
        // Collect all non-empty prefixes for this column
        const prefixes = [];
        for (let row = 1; row < data.length; row++) {
          if (data[row][col] && data[row][col].toString().trim()) {
            prefixes.push(data[row][col].toString().trim());
          }
        }
        
        if (prefixes.length === 0) continue; // Skip if no prefixes for this category
        
        // Generate output filename from the header
        const outputFileName = `${headers[col]}.pdf`;
        console.log(`Processing category "${headers[col]}" with prefixes: ${prefixes.join(', ')}`);
        
        // Find all files matching these prefixes
        const fileIds = DriveManager.getFileIdsBySubstring(
          sourceFolderId, 
          prefixes,
          recursive, 
          ["application/pdf"],
          'prefix'  // We're still using prefix matching as before
        );
        
        if (fileIds.length === 0) {
          results.push({
            category: headers[col],
            success: false,
            message: "No matching PDF files found for this category."
          });
          continue;
        }
        
        // Merge PDFs for this category
        const mergeResult = await this.mergePDFs(fileIds, outputFileName, outputFolderId);
        
        // Store the result with category info
        results.push({
          category: headers[col],
          ...mergeResult
        });
      }
      
      return {
        success: true,
        message: `Processed ${results.length} categories from the prefixes sheet.`,
        results
      };
    } catch (e) {
      console.error(`Error merging PDFs from prefix sheet: ${e.message}`);
      return {
        success: false,
        message: `Error merging PDFs from prefix sheet: ${e.message}`
      };
    }
  }
  
  /**
   * Merges PDFs for each student folder based on the prefix sheet
   * @param {boolean} [recursive=false] - Whether to search in subfolders recursively
   * @returns {Promise<Object>} Object with results of the merge operations for each student
   */
  static async mergePDFsForAllStudents(recursive = false) {
    try {
      // Get the Student Info sheet
      const { studentSheet } = SpreadsheetManager.getSpreadsheetSheets();
      if (!studentSheet) {
        return {
          success: false,
          message: "Student Info sheet not found."
        };
      }
      
      // Get all data from the Student Info sheet
      const data = studentSheet.getDataRange().getValues();
      if (data.length < 2) {
        return {
          success: false,
          message: "Student Info sheet is empty or has only headers."
        };
      }
      
      const headers = data[0];
      const folderIdColumnIndex = headers.findIndex(header => header.toLowerCase().includes("folder")) || 2; // Default to 3rd column (index 2)
      const nameColumnIndex = 0; // Assume name is in the first column
      
      const studentResults = [];
      
      // Process each student row (skip header row)
      for (let row = 1; row < data.length; row++) {
        const studentName = data[row][nameColumnIndex];
        const sourceFolderId = data[row][folderIdColumnIndex];
        
        if (!sourceFolderId) {
          studentResults.push({
            student: studentName || `Row ${row + 1}`,
            success: false,
            message: "No folder ID found for this student."
          });
          continue;
        }
        
        try {
          // Get the source folder
          const sourceFolder = DriveApp.getFolderById(sourceFolderId);
          
          // Create a "MergedPDFs" subfolder
          const mergedPDFsFolder = DriveManager.createFolder(sourceFolder, "MergedPDFs");
          const outputFolderId = mergedPDFsFolder.getId();
          
          console.log(`Processing student: ${studentName}, creating merged PDFs in folder: ${outputFolderId}`);
          
          // Run the merge operation for this student's folder
          const mergeResult = await this.mergePDFsFromPrefixSheet(
            sourceFolderId,
            outputFolderId,
            recursive
          );
          
          // Store the result with student info
          studentResults.push({
            student: studentName,
            sourceFolderId,
            outputFolderId,
            ...mergeResult
          });
          
        } catch (e) {
          console.error(`Error processing student ${studentName}: ${e.message}`);
          studentResults.push({
            student: studentName,
            success: false,
            message: `Error: ${e.message}`
          });
        }
      }
      
      return {
        success: true,
        message: `Processed PDF merges for ${studentResults.length} students.`,
        studentResults
      };
    } catch (e) {
      console.error(`Error merging PDFs for students: ${e.message}`);
      return {
        success: false,
        message: `Error merging PDFs for students: ${e.message}`
      };
    }
  }
}
