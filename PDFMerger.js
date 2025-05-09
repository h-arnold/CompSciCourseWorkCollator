/**
 * Class for handling PDF merge operations using singleton pattern
 */
class PDFMerger {
  /**
   * Get the singleton instance of PDFMerger
   * @returns {PDFMerger} The singleton instance
   */
  static getInstance() {
    if (!PDFMerger._instance) {
      PDFMerger._instance = new PDFMerger();
    }
    return PDFMerger._instance;
  }
  
  /**
   * Private constructor to prevent direct instantiation
   * Use PDFMerger.getInstance() instead
   */
  constructor() {
    // Prevent multiple instances when using new PDFMerger()
    if (PDFMerger._instance) {
      console.warn('PDFMerger is a singleton. Use PDFMerger.getInstance() instead of new PDFMerger()');
      return PDFMerger._instance;
    }
    
    /**
     * Default destination folder for merged PDFs
     * @type {Folder|null}
     */
    this.destinationFolder = null;
    this.pdfLib = this.loadPdfLib();
    
    // Set this as the singleton instance
    PDFMerger._instance = this;
  }

  /**
   * Sets the destination folder for merged PDFs
   * @param {Folder} folder - The folder to set as destination
   */
  setDestinationFolder(folder) {
    this.destinationFolder = folder;
  }

  /**
   * Load the PDF-lib library from CDN
   * Docs for the library can be found here:
   * https://pdf-lib.js.org/
   * @returns {Promise<Object>} Promise that resolves with the PDFLib object
   */
  async loadPdfLib() {
    const cdnUrl = "https://cdn.jsdelivr.net/npm/pdf-lib/dist/pdf-lib.min.js";
    const response = UrlFetchApp.fetch(cdnUrl);
    const content = response.getContentText();

    // Create a new context to evaluate the library in
    const context = {};

    // Evaluate the library code in the context
    try {
      // Define setTimeout for compatibility with pdf-lib
      context.setTimeout = function (f, t) {
        Utilities.sleep(t);
        return f();
      };

      // Use Function constructor instead of eval for better scoping
      const setupLibrary = new Function(
        "context",
        `
        with(context) {
          ${content}
          return { PDFLib: PDFLib };
        }
      `
      );

      return setupLibrary(context);
    } catch (e) {
      console.error("Error loading PDF-lib:", e);
      throw new Error(`Failed to load PDF-lib: ${e.message}`);
    }
  }

  /**
   * Validates if all provided items represent PDF files
   * @param {(string|File)[]} items - Array of Google Drive file IDs or File objects
   * @returns {Object} Object containing valid and invalid files
   */
  validateFiles(items) {
    const validFiles = [];
    const invalidFiles = [];

    items.forEach((item) => {
      try {
        let file;
        // Check if the item is a string (fileId) or a File object
        if (typeof item === "string") {
          file = DriveApp.getFileById(item);
        } else if (item.getMimeType) {
          // Assume it's a File object if it has getMimeType method
          file = item;
        } else {
          throw new Error(
            "Invalid item type: must be a file ID string or File object"
          );
        }

        if (DriveManager.isPdf(file)) {
          validFiles.push(file);
        } else {
          // Log the file it could not process for debugging purposes.
          let idValue;
          if (typeof item === "string") {
            idValue = item;
          } else {
            idValue = file.getId();
          }

          invalidFiles.push({
            id: idValue,
            name: file.getName(),
            type: file.getMimeType(),
          });
        }
      } catch (e) {
        let idValue;
        if (typeof item === "string") {
          idValue = item;
        } else {
          idValue = "unknown";
        }

        invalidFiles.push({
          id: idValue,
          error: e.message,
        });
      }
    });

    return { validFiles, invalidFiles };
  }

  /**
   * Handles the case when only one PDF file is found - copies it instead of merging
   * @param {File} file - The single PDF file to copy
   * @param {string} outputFileName - Name for the output PDF file
   * @param {Folder} [outputFolder=null] - Optional folder to save the PDF (if null, saves to root)
   * @returns {Promise<Object>} Object with status and result information
   */
  async copySinglePdfFile(file, outputFileName, outputFolder = null) {
    console.log(
      `Only one valid PDF found, copying instead of merging: ${file.getName()}`
    );
    let newFile;
    const effectiveOutputFolder = outputFolder || this.destinationFolder;

    if (effectiveOutputFolder) {
      // Use the DriveManager helper to copy and rename
      const fileId = file.getId();
      DriveManager.copyAndRenameFile(fileId, effectiveOutputFolder, outputFileName);
      // Get the newly created file
      const newFiles = effectiveOutputFolder.getFilesByName(outputFileName);
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
          url: newFile.getUrl(),
        },
      };
    }

    return {
      success: false,
      message: `Failed to copy the PDF file: ${file.getName()}`,
    };
  }

  /**
   * Merges multiple PDF files into a single PDF document
   * @param {File[]} files - Array of PDF files to merge
   * @returns {Promise<Object>} Object containing PDF bytes
   */
  async mergeMultiplePdfFiles(files) {
    // Load PDF-lib library
    const lib = await this.loadPdfLib();
    const { PDFLib } = lib;

    // Create a new PDF document
    const pdfDoc = await PDFLib.PDFDocument.create();
    
    let totalPages = 0;

    // Add each valid PDF to the merged document
    for (const file of files) {
      console.log(`Processing ${file.getName()} (${file.getMimeType()})`);
      try {
        // Get the PDF data as bytes
        const fileBlob = file.getBlob();
        const base64Content = Utilities.base64Encode(fileBlob.getBytes());
        
        // Load the PDF document using base64 string
        const pdfStudentData = await PDFLib.PDFDocument.load(
          PDFLib.base64.decode(base64Content)
        );
        
        // Get page indices and copy them
        const pageCount = pdfStudentData.getPageCount();
        const pageIndices = [...Array(pageCount)].map((_, i) => i);
        totalPages += pageCount;
        
        console.log(`Copying ${pageCount} pages from ${file.getName()}`);
        const pages = await pdfDoc.copyPages(pdfStudentData, pageIndices);
        
        // Add pages to the output document
        pages.forEach((page) => pdfDoc.addPage(page));
      } catch (e) {
        console.error(`Error processing ${file.getName()}: ${e.message}`);
        console.error(e.stack); // Log the full stack trace for debugging
      }
    }

    console.log(`Merged document has ${totalPages} pages. Saving...`);
    
    try {
      // Save the document as bytes
      const pdfBytes = await pdfDoc.save();
      console.log(`Successfully saved PDF, byte length: ${pdfBytes.length}`);
      
      // Return only the bytes
      return {
        bytes: pdfBytes
      };
    } catch (e) {
      console.error(`Error saving merged PDF: ${e.message}`);
      console.error(e.stack);
      throw e; // Re-throw to be caught by the calling function
    }
  }

  /**
   * Process PDFs in smaller batches to avoid memory issues
   * @param {File[]} files - Array of PDF files to merge
   * @param {number} batchSize - Size of each batch
   * @returns {Promise<Object>} Object containing merged PDF data (bytes)
   */
  async processPDFsInBatches(files, batchSize = 5) {
    console.log(`Processing ${files.length} files in batches of ${batchSize}`);
    
    // If there are fewer files than the batch size, just process them directly
    if (files.length <= batchSize) {
      return await this.mergeMultiplePdfFiles(files);
    }
    
    // Create temporary storage for output files
    const tempFiles = [];
    const parentFolderForTempFiles = this.destinationFolder || DriveApp.getRootFolder();
    
    try {
      // Process files in batches
      for (let i = 0; i < files.length; i += batchSize) {
        const end = Math.min(i + batchSize, files.length);
        const batch = files.slice(i, end);
        console.log(`Processing batch ${Math.floor(i/batchSize) + 1}: files ${i+1} to ${end}`);
        
        // Merge this batch into a temporary PDF
        const batchData = await this.mergeMultiplePdfFiles(batch);
        
        // Save it as a temporary file in Drive root
        const tempName = `temp_batch_${Math.floor(i/batchSize) + 1}_of_${Math.ceil(files.length/batchSize)}.pdf`;
        const tempBlob = Utilities.newBlob(
          batchData.bytes, // Use bytes directly
          MimeType.PDF,
          tempName
        );
        
        const tempFile = parentFolderForTempFiles.createFile(tempBlob);
        tempFiles.push(tempFile);
        console.log(`Created temporary file: ${tempFile.getName()} (${tempFile.getId()}) in folder ${parentFolderForTempFiles.getName()}`);
      }
      
      console.log(`Created ${tempFiles.length} temporary files, now merging them`);
      
      // Now merge all temporary files
      const finalData = await this.mergeMultiplePdfFiles(tempFiles);
      
      // Clean up temp files
      tempFiles.forEach(file => {
        try {
          file.setTrashed(true);
          console.log(`Deleted temporary file: ${file.getName()}`);
        } catch (e) {
          console.warn(`Failed to delete temporary file ${file.getName()}: ${e.message}`);
        }
      });
      
      return finalData;
    } catch (e) {
      console.error(`Error in batch processing: ${e.message}`);
      console.error(e.stack);
      
      // Try to clean up temp files even if there was an error
      tempFiles.forEach(file => {
        try {
          file.setTrashed(true);
        } catch (cleanupErr) {
          console.warn(`Failed to delete temporary file during cleanup: ${cleanupErr.message}`);
        }
      });
      
      throw e;
    }
  }

  /**
   * Saves a PDF document to Google Drive
   * @param {Object} pdfData - Object with PDF bytes (`pdfData.bytes`)
   * @param {string} outputFileName - Name for the output PDF file
   * @param {Folder} [outputFolder=null] - Optional folder to save the PDF (if null, saves to root)
   * @returns {Object} Object with file information
   */
  saveResultingPdf(pdfData, outputFileName, outputFolder = null) {
    console.log(`Saving PDF to Drive with filename: ${outputFileName}`);
    
    // Get the PDF file bytes directly
    const bytes = pdfData.bytes;
    console.log(`Using PDF data with ${bytes.length} bytes`);
    
    // Create Blob
    const blob = Utilities.newBlob(bytes, MimeType.PDF, outputFileName);
    console.log(`Created blob with size: ${blob.getBytes().length} bytes`);
    
    let newFile;
    const effectiveOutputFolder = outputFolder || this.destinationFolder;

    if (effectiveOutputFolder) {
      // Save to specified folder
      console.log(`Saving to folder: ${effectiveOutputFolder.getName()}`);
      newFile = effectiveOutputFolder.createFile(blob);
    } else {
      // Save to root
      console.log("Saving to root folder");
      newFile = DriveApp.createFile(blob);
    }
    
    console.log(`File created with id: ${newFile.getId()}`);
    return {
      id: newFile.getId(),
      name: newFile.getName(),
      url: newFile.getUrl(),
    };
  }

  /**
   * Merges multiple PDF files into a single PDF
   * @param {(string|File)[]} items - Array of Google Drive file IDs or File objects to merge
   * @param {string} outputFileName - Name for the merged PDF file (default: "Merged.pdf")
   * @param {Folder} [outputFolder=null] - Optional folder to save the merged PDF (if null, saves to root)
   * @returns {Promise<Object>} Object with status and result information
   */
  async mergePDFs(
    items,
    outputFileName = "Merged.pdf",
    outputFolder = null
  ) {
    try {
      // Validate files first
      const { validFiles, invalidFiles } = this.validateFiles(items);

      if (validFiles.length === 0) {
        return {
          success: false,
          message: "No valid PDF files found to merge",
          invalidFiles,
        };
      }

      // If there's only one valid file, simply copy it with the new name
      if (validFiles.length === 1) {
        const result = await this.copySinglePdfFile(
          validFiles[0],
          outputFileName,
          outputFolder
        );
        if (invalidFiles.length > 0) {
          result.invalidFiles = invalidFiles;
        }
        return result;
      }

     
      // Process in batches if we have a large number of files
      let pdfData;
      if (validFiles.length > 10) {
        console.log("Large number of PDFs detected, processing in smaller batches");
        pdfData = await this.processPDFsInBatches(validFiles, 5);
      } else {
        // For smaller file sets, just merge directly
        pdfData = await this.mergeMultiplePdfFiles(validFiles);
      }
      
      if (!pdfData || !pdfData.bytes || pdfData.bytes.length === 0) {
        console.error("Error: No PDF data returned from merge operation");
        return {
          success: false,
          message: "Failed to create merged PDF - no data returned",
          invalidFiles: invalidFiles.length > 0 ? invalidFiles : null,
        };
      }
      
      console.log(`Successfully merged PDFs with byte length: ${pdfData.bytes.length}, saving to Drive with name: ${outputFileName}`);

      try {
        // Save the merged PDF to Drive
        const fileInfo = this.saveResultingPdf(
          pdfData,
          outputFileName,
          outputFolder
        );

        return {
          success: true,
          message: `Successfully merged ${validFiles.length} PDFs`,
          file: fileInfo,
          invalidFiles: invalidFiles.length > 0 ? invalidFiles : null,
        };
      } catch (saveError) {
        console.error(`Error saving merged PDF: ${saveError.message}`);
        console.error(saveError.stack);
        
        // Try an alternative save approach for large files
        return {
          success: false,
          message: `Error saving merged PDF: ${saveError.message}`,
          invalidFiles: invalidFiles.length > 0 ? invalidFiles : null,
        };
      }
    } catch (e) {
      console.error(`Error merging PDFs: ${e.message}`);
      console.error(e.stack);
      return {
        success: false,
        message: `Error merging PDFs: ${e.message}`,
      };
    }
  }

  /**
   * Gets prefixes from the Google Sheet and merges PDFs based on column groups
   * @param {Folder} sourceFolder - Folder containing the PDFs to merge
   * @param {Folder} [outputFolder=null] - Optional folder to save the merged PDFs (if null, saves to root)
   * @param {boolean} [recursive=false] - Whether to search in subfolders recursively
   * @returns {Promise<Object>} Object with results of the merge operations
   */
  async mergePDFsFromPrefixSheet(
    sourceFolder,
    outputFolder = null,
    recursive = false
  ) {
    try {
      // Get the Prefixes sheet
      const { prefixSheet } = SpreadsheetManager.getSpreadsheetSheets();
      if (!prefixSheet) {
        return {
          success: false,
          message: "Prefixes sheet not found.",
        };
      }

      // Get all data from the sheet
      const data = prefixSheet.getDataRange().getValues();
      if (data.length < 2) {
        return {
          success: false,
          message: "Prefixes sheet is empty or has only headers.",
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
          if (
            data[row][col] &&
            data[row][col].toString().trim()
          ) {
            prefixes.push(data[row][col].toString().trim());
          }
        }

        if (prefixes.length === 0) continue; // Skip if no prefixes for this category

        // Generate output filename from the header
        const outputFileName = `${headers[col]}.pdf`;
        console.log(
          `Processing category "${headers[col]}" with prefixes: ${prefixes.join(
            ", "
          )}`
        );

        // Find all files matching these prefixes
        const fileIds = DriveManager.getFileIdsBySubstring(
          sourceFolder,
          prefixes,
          recursive,
          ["application/pdf"],
          "prefix" // We're still using prefix matching as before
        );

        if (fileIds.length === 0) {
          results.push({
            category: headers[col],
            success: false,
            message: "No matching PDF files found for this category.",
          });
          continue;
        }

        // Merge PDFs for this category
        const mergeResult = await this.mergePDFs(
          fileIds,
          outputFileName,
          outputFolder
        );

        // Store the result with category info
        results.push({
          category: headers[col],
          ...mergeResult,
        });
      }

      return {
        success: true,
        message: `Processed ${results.length} categories from the prefixes sheet.`,
        results,
      };
    } catch (e) {
      console.error(`Error merging PDFs from prefix sheet: ${e.message}`);
      return {
        success: false,
        message: `Error merging PDFs from prefix sheet: ${e.message}`,
      };
    }
  }

  /**
   * Merges PDFs for each student folder based on the prefix sheet
   * @param {Array} studentData - Student data from spreadsheet
   * @param {string} destinationFolderId - ID of the destination folder
   * @param {boolean} [recursive=false] - Whether to search in subfolders recursively
   * @returns {Promise<Object>} Object with results of the merge operations for each student
   */
  async mergePDFsForAllStudents(
    studentData,
    destinationFolderId,
    recursive = false
  ) {
    try {
      const headers = studentData[0];
      const folderIdColumnIndex =  2

      const nameColumnIndex = 0; // Assume name is in the first column

      const studentResults = [];

      // Process each student row (skip header row)
      for (let row = 3; row < studentData.length; row++) {
        const studentName = studentData[row][nameColumnIndex];
        const sourceFolderId = studentData[row][folderIdColumnIndex];
        const destinationFolderName = studentData[row][3]; // 3rd column from studentData should be the destination folder name.

        if (!sourceFolderId) {
          studentResults.push({
            student: studentName || `Row ${row + 1}`,
            success: false,
            message: "No folder ID found for this student.",
          });
          continue;
        }

        try {
          // Get the source folder
          const sourceFolder = DriveApp.getFolderById(sourceFolderId);

          // Create a folder in the parent folder
          const destFolder = DriveApp.getFolderById(destinationFolderId);
          const mergedPDFsFolder = DriveManager.createFolderInParentFolder(
            destFolder,
            destinationFolderName || "MergedPDFs"
          );

          console.log(
            `Processing student: ${studentName}, creating merged PDFs in folder: ${mergedPDFsFolder.getName()}`
          );

          // Run the merge operation for this student's folder
          const mergeResult = await this.mergePDFsFromPrefixSheet(
            sourceFolder,
            mergedPDFsFolder,
            recursive
          );

          // Store the result with student info
          studentResults.push({
            student: studentName,
            sourceFolderId,
            outputFolderId: mergedPDFsFolder.getId(),
            ...mergeResult,
          });
        } catch (e) {
          console.error(`Error processing student ${studentName}: ${e.message}`);
          studentResults.push({
            student: studentName,
            success: false,
            message: `Error: ${e.message}`,
          });
        }
      }

      return {
        success: true,
        message: `Processed PDF merges for ${studentResults.length} students.`,
        studentResults,
      };
    } catch (e) {
      console.error(`Error merging PDFs for students: ${e.message}`);
      return {
        success: false,
        message: `Error merging PDFs for students: ${e.message}`,
      };
    }
  }
}

/**
 * Static instance property for the PDFMerger singleton
 * @type {PDFMerger|null}
 * @private
 */
PDFMerger._instance = null;
