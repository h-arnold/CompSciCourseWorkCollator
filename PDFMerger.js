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
      
      // Load PDF-lib library
      const lib = await this.loadPdfLib();
      const { PDFLib } = lib;
      
      // Create a new PDF document
      const pdfDoc = await PDFLib.PDFDocument.create();
      
      // Add each valid PDF to the merged document
      for (const file of validFiles) {
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
      const bytes = await pdfDoc.save();
      
      // Create the merged PDF file
      const blob = Utilities.newBlob([...new Int8Array(bytes)], MimeType.PDF, outputFileName);
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
        success: true,
        message: `Successfully merged ${validFiles.length} PDFs`,
        file: {
          id: newFile.getId(),
          name: newFile.getName(),
          url: newFile.getUrl()
        },
        invalidFiles: invalidFiles.length > 0 ? invalidFiles : null
      };
    } catch (e) {
      return {
        success: false,
        message: `Error merging PDFs: ${e.message}`
      };
    }
  }
}
