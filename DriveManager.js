/**
 * Class for Google Drive operations
 */
class DriveManager {
    /**
     * Creates a folder in the specified parent folder
     * @param {Folder} parentFolder - The parent folder
     * @param {string} folderName - The name of the folder to create
     * @returns {Folder} The created folder
     */
    static createFolder(parentFolder, folderName) {
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
     * @returns {Object[]} Array of objects with file information {id, name, mimeType, url}
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
                if (!matchingFiles.some(existingFile => existingFile.id === fileId)) {
                  matchingFiles.push({
                    id: fileId,
                    name: fileName,
                    mimeType: file.getMimeType(),
                    url: file.getUrl()
                  });
                }
              }
            }
          }
        }
        
        return matchingFiles;
      } catch (e) {
        console.error(`Error finding files by substring: ${e.message}`);
        return [];
      }
    }
    
    // For backward compatibility
    /**
     * @deprecated Use findFilesBySubstring with searchType='prefix' instead
     */
    static findFilesByPrefix(folderIdOrFolder, startsWithPrefixes, recursive = false, mimeTypes = null) {
      return this.findFilesBySubstring(folderIdOrFolder, startsWithPrefixes, recursive, mimeTypes, 'prefix');
    }
    
    /**
     * Copies and renames a Drive file
     * @param {string} driveFileId - The ID of the Drive file
     * @param {string} folderId - The ID of the folder where the file will be copied
     * @param {string} newFileName - The name of the file
     * @returns {File|null} The copied file, or null if the operation was skipped
     */
    static copyAndRenameFile(driveFileId, folderId, newFileName) {
      const driveFile = DriveApp.getFileById(driveFileId);
      const folder = DriveApp.getFolderById(folderId);
      
      // Check if file with same name already exists
      if (this.checkAndHandleExistingFile(folder, newFileName)) {
        return driveFile.makeCopy(newFileName, folder);
      }
      return null;
    }
    
    /**
     * Converts a Google Docs file to a PDF
     * @param {string} driveFileId - The ID of the Drive file
     * @return {File} The converted PDF file
     */
    static convertToPdf(driveFileId) {
      const driveFile = DriveApp.getFileById(driveFileId);
      const blob = driveFile.getAs('application/pdf');
      const pdfFile = DriveApp.createFile(blob);
      return pdfFile;
    }
    
    /**
     * Gets just the file IDs from a folder where filenames match specified substrings
     * @param {string|Folder} folderIdOrFolder - The folder ID or Folder object to search in
     * @param {string|string[]} substrings - A single substring or array of substrings to match in file names
     * @param {boolean} [recursive=false] - Whether to search in subfolders recursively
     * @param {string[]} [mimeTypes=null] - Optional array of MIME types to filter by
     * @param {string} [searchType='prefix'] - Type of search: 'prefix', 'suffix', or 'contains'
     * @returns {string[]} Array of file IDs
     */
    static getFileIdsBySubstring(folderIdOrFolder, substrings, recursive = false, mimeTypes = null, searchType = 'prefix') {
      const files = this.findFilesBySubstring(folderIdOrFolder, substrings, recursive, mimeTypes, searchType);
      return files.map(file => file.id);
    }
    
    // For backward compatibility
    /**
     * @deprecated Use getFileIdsBySubstring with searchType='prefix' instead
     */
    static getFileIdsByPrefix(folderIdOrFolder, startsWithPrefixes, recursive = false, mimeTypes = null) {
      return this.getFileIdsBySubstring(folderIdOrFolder, startsWithPrefixes, recursive, mimeTypes, 'prefix');
    }
  }