/**
 * Class for Google Drive operations
 */
class DriveManager {
    /**
     * Creates a copy of a Google Document.
     * @param {string|File} sourceFileIdOrFile - The ID of the Google Document to copy or the File object.
     * @param {string} newName - The name for the new copied document.
     * @param {string|Folder} [destinationFolderIdOrFolder=null] - Optional ID of the folder or Folder object to place the copy in. Defaults to the source file's parent folder.
     * @returns {File|null} The newly created File object, or null on error, if source is not a Doc, or if user cancels overwrite.
     */
    static copyDocument(sourceFileIdOrFile, newName, destinationFolderIdOrFolder = null) {
      try {
        // Handle whether we received a string ID or a File object
        const sourceFile = typeof sourceFileIdOrFile === 'string' 
          ? DriveApp.getFileById(sourceFileIdOrFile) 
          : sourceFileIdOrFile;

        // Ensure it's a Google Doc
        if (sourceFile.getMimeType() !== MimeType.GOOGLE_DOCS) {
            console.error(`File ${sourceFile.getId()} (${sourceFile.getName()}) is not a Google Doc. MimeType: ${sourceFile.getMimeType()}`);
            return null;
        }

        let targetFolder;
        if (destinationFolderIdOrFolder) {
            try {
                // Handle whether we received a string ID or a Folder object
                targetFolder = typeof destinationFolderIdOrFolder === 'string'
                  ? DriveApp.getFolderById(destinationFolderIdOrFolder)
                  : destinationFolderIdOrFolder;
                  
                console.log(`Using specified destination folder: "${targetFolder.getName()}" (ID: ${targetFolder.getId()})`);
            } catch (e) {
                console.error(`Error accessing specified destination folder: ${e}. Falling back to source file's parent.`);
                targetFolder = null; // Reset targetFolder so fallback logic runs
            }
        }

        // If no valid destination folder specified, use the source file's parent
        if (!targetFolder) {
            const parents = sourceFile.getParents();
            if (parents.hasNext()) {
                targetFolder = parents.next(); // Use the first parent
                console.log(`Using source file's parent folder: "${targetFolder.getName()}" (ID: ${targetFolder.getId()})`);
            } else {
                console.warn(`Source file ${sourceFile.getId()} has no parent folder. Placing copy in root folder.`);
                targetFolder = DriveApp.getRootFolder(); // Default to root folder
            }
        }

        // Check for existing file with the new name in the target folder
        if (this.checkAndHandleExistingFile(targetFolder, newName)) {
            const copiedFile = sourceFile.makeCopy(newName, targetFolder);
            console.log(`Document ${sourceFile.getId()} copied to "${newName}" (ID: ${copiedFile.getId()}) in folder "${targetFolder.getName()}".`);
            return copiedFile;
        } else {
            console.log(`Skipped copying document to "${newName}" as user chose not to replace existing file.`);
            return null; // Operation skipped by user
        }

      } catch (e) {
        console.error(`Error copying document ${sourceFileIdOrFile} to "${newName}": ${e}`);
        SpreadsheetApp.getUi().alert(`Error copying document: ${e.message}`);
        return null;
      }
    }

    /**
     * Creates a folder in the specified parent folder
     * @param {string|Folder} parentFolderIdOrFolder - The parent folder ID or Folder object
     * @param {string} folderName - The name of the folder to create
     * @returns {Folder} The created folder
     */
    static createFolder(parentFolderIdOrFolder, folderName) {
      // Convert string ID to Folder object if needed
      const parentFolder = typeof parentFolderIdOrFolder === 'string'
        ? DriveApp.getFolderById(parentFolderIdOrFolder)
        : parentFolderIdOrFolder;
        
      // Check if the folder already exists and get user confirmation if needed
      if (this.checkAndHandleExistingFolder(parentFolder, folderName)) {
        return parentFolder.createFolder(folderName);
      }
      // If the folder exists and the user chose not to replace it, find and return the existing folder
      return parentFolder.getFoldersByName(folderName).next();
    }
    
    /**
     * Checks if a file with the given name exists in the folder
     * @param {Folder} folder - The folder to check within
     * @param {string} newFileName - The name of the file to check for
     * @returns {boolean} True if the operation should proceed (file doesn't exist or user chose YES), false otherwise
     */
    static checkAndHandleExistingFile(folder, newFileName) {
      const existingFiles = folder.getFilesByName(newFileName);
      if (existingFiles.hasNext()) {
        const ui = SpreadsheetApp.getUi();
        const response = ui.alert(
          `File "${newFileName}" already exists in the destination folder. Replace?`,
          ui.ButtonSet.YES_NO);
    
        if (response == ui.Button.YES) {
          // Remove existing file(s)
          while (existingFiles.hasNext()) {
            existingFiles.next().setTrashed(true); // Move to trash
          }
          console.log(`Existing file "${newFileName}" marked for replacement.`);
          return true; // Proceed with operation
        } else {
          console.log(`Skipping replacement for existing file "${newFileName}".`);
          return false; // Skip operation
        }
      }
      return true; // File doesn't exist, proceed with operation
    }
    
    /**
     * Checks if a folder with the given name exists in the parent folder
     * @param {Folder} parentFolder - The parent folder to check within
     * @param {string} newFolderName - The name of the folder to check for
     * @returns {boolean} True if the operation should proceed (folder doesn't exist or user chose YES), false otherwise
     */
    static checkAndHandleExistingFolder(parentFolder, newFolderName) {
      const existingFolders = parentFolder.getFoldersByName(newFolderName);
      if (existingFolders.hasNext()) {
        const ui = SpreadsheetApp.getUi();
        const response = ui.alert(
          `Folder "${newFolderName}" already exists in the destination location. Replace?`,
          ui.ButtonSet.YES_NO);
    
        if (response == ui.Button.YES) {
          // Remove existing folder(s)
          while (existingFolders.hasNext()) {
            existingFolders.next().setTrashed(true); // Move to trash
          }
          console.log(`Existing folder "${newFolderName}" marked for replacement.`);
          return true; // Proceed with operation
        } else {
          console.log(`Skipping replacement for existing folder "${newFolderName}".`);
          return false; // Skip operation
        }
      }
      return true; // Folder doesn't exist, proceed with operation
    }
    
    /**
     * Checks if the provided file is a Zip archive
     * @param {File} file - The Drive file to check
     * @returns {boolean} True if the file's MIME type is application/x-zip-compressed
     */
    static isZip(file) {
      return file.getMimeType() === "application/x-zip-compressed";
    }
    
    /**
     * Checks if the provided file is a PDF
     * @param {File} file - The Drive file to check
     * @returns {boolean} True if the file's MIME type is PDF
     */
    static isPdf(file) {
      return file.getMimeType() === "application/pdf";
    }
    
    /**
     * Checks if the provided file is a Google Docs file
     * @param {File} file - The Drive file to check
     * @returns {boolean} True if the file's MIME type indicates a Google Docs document
     */
    static isGoogleDoc(file) {
      return file.getMimeType() === "application/vnd.google-apps.document";
    }
    
    /**
     * Copies the given file to the specified folder with a new name
     * @param {File} file - The Drive file to copy
     * @param {Folder} folder - The destination folder
     * @param {string} prependString - The string to prepend to the file name
     */
    static copyFile(file, folder, prependString) {
      const newFileName = `${prependString}_${file.getName()}`;
      if (this.checkAndHandleExistingFile(folder, newFileName)) {
        file.makeCopy(newFileName, folder);
        console.log(`File "${file.getName()}" copied as "${newFileName}".`);
      }
    }
    
    /**
     * Converts a Google Docs file to PDF and copies it to the specified folder with a new name
     * @param {File} file - The Google Docs file to convert
     * @param {Folder} folder - The destination folder
     * @param {string} prependString - The string to prepend to the file name
     */
    static copyGoogleDocAsPdf(file, folder, prependString) {
      const pdfBlob = file.getAs("application/pdf");
      // Append '.pdf' to the original name for clarity.
      const newFileName = `${prependString}_${file.getName()}.pdf`;
      if (this.checkAndHandleExistingFile(folder, newFileName)) {
        folder.createFile(pdfBlob).setName(newFileName);
        console.log(`Document "${file.getName()}" (converted to PDF) copied as "${newFileName}".`);
      }
    }
    
    /**
     * Finds files in a folder with names matching specified substrings
     * @param {string|Folder} folderIdOrFolder - The folder ID or Folder object to search in
     * @param {string|string[]} substrings - A single substring or array of substrings to match in file names
     * @param {boolean} [recursive=false] - Whether to search in subfolders recursively
     * @param {string[]} [mimeTypes=null] - Optional array of MIME types to filter by (e.g., ["application/pdf"])
     * @param {string} [searchType='prefix'] - Type of search: 'prefix', 'suffix', or 'contains'
     * @returns {File|File[]} A single File object if only one match, otherwise an array of Google Drive File objects
     */
    static findFilesBySubstring(folderIdOrFolder, substrings, recursive = false, mimeTypes = null, searchType = 'prefix') {
      try {
        const folder = (typeof folderIdOrFolder === 'string') 
          ? DriveApp.getFolderById(folderIdOrFolder) 
          : folderIdOrFolder;
            
        // Convert single string to array for consistent handling
        const substringArray = Array.isArray(substrings) ? substrings : [substrings];
        
        const matchingFiles = [];
        
        // Function to collect all files from a folder and its subfolders if recursive
        const getAllFiles = (currentFolder) => {
          const allFiles = [];
          
          // Get all files in the current folder
          const files = currentFolder.getFiles();
          while (files.hasNext()) {
            allFiles.push(files.next());
          }
          
          // If recursive search is enabled, process subfolders
          if (recursive) {
            const subFolders = currentFolder.getFolders();
            while (subFolders.hasNext()) {
              allFiles.push(...getAllFiles(subFolders.next()));
            }
          }
          
          return allFiles;
        };
        
        // Get all files from the folder (and subfolders if recursive)
        const allFiles = getAllFiles(folder);
        
        // Process each substring in the order they were provided
        for (const substring of substringArray) {
          // For each file, check if it matches according to the search type
          for (const file of allFiles) {
            const fileName = file.getName();
            let isMatch = false;
            
            // Determine if the file name matches based on the search type
            switch (searchType.toLowerCase()) {
              case 'prefix':
                isMatch = fileName.startsWith(substring);
                break;
              case 'suffix':
                isMatch = fileName.endsWith(substring);
                break;
              case 'contains':
              default:
                isMatch = fileName.includes(substring);
                break;
            }
            
            if (isMatch) {
              // If mimeTypes is provided, check if the file's mimeType matches any in the array
              if (!mimeTypes || mimeTypes.includes(file.getMimeType())) {
                // Check if this file has already been added to avoid duplicates
                const fileId = file.getId();
                if (!matchingFiles.some(existingFile => existingFile.getId() === fileId)) {
                  matchingFiles.push(file);
                }
              }
            }
          }
        }

        // Return matching file if the array only has one item
        if (matchingFiles.length === 1) {
          return matchingFiles[0];
        }
        
        return matchingFiles;
      } catch (e) {
        console.error(`Error finding files by substring: ${e.message}`);
        return [];
      }
    }
    
  
    /**
     * Copies and renames a Drive file
     * @param {string|File} driveFileIdOrFile - The ID of the Drive file or File object
     * @param {string|Folder} folderIdOrFolder - The ID of the folder or Folder object where the file will be copied
     * @param {string} newFileName - The name of the file
     * @returns {File|null} The copied file, or null if the operation was skipped
     */
    static copyAndRenameFile(driveFileIdOrFile, folderIdOrFolder, newFileName) {
      // Handle whether we received a string ID or a File object
      const driveFile = typeof driveFileIdOrFile === 'string' 
        ? DriveApp.getFileById(driveFileIdOrFile) 
        : driveFileIdOrFile;
      
      // Handle whether we received a string ID or a Folder object
      const folder = typeof folderIdOrFolder === 'string'
        ? DriveApp.getFolderById(folderIdOrFolder)
        : folderIdOrFolder;
      
      // Check if file with same name already exists
      if (this.checkAndHandleExistingFile(folder, newFileName)) {
        return driveFile.makeCopy(newFileName, folder);
      }
      return null;
    }
    
    /**
     * Converts a Google Docs file to a PDF
     * @param {string|File} fileIdOrFile - The ID of the Drive file or a File object
     * @param {string} [newFileName=null] - Optional new name for the PDF file (without extension)
     * @return {File} The converted PDF file
     */
    static convertToPdf(fileIdOrFile, newFileName = null) {
      // Handle whether we received a string ID or a File object
      const driveFile = typeof fileIdOrFile === 'string' 
        ? DriveApp.getFileById(fileIdOrFile) 
        : fileIdOrFile;
      
      const blob = driveFile.getAs('application/pdf');
      
      // Set the PDF filename, either using the provided name or the original filename
      let pdfFileName;
      if (newFileName) {
        pdfFileName = `${newFileName}.pdf`;
      } else {
        pdfFileName = `${driveFile.getName()}.pdf`;
      }
      
      const pdfFile = DriveApp.createFile(blob).setName(pdfFileName);
      return pdfFile;
    }
}