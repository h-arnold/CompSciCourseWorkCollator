# Folder Populator for A-Level Computer Science Coursework

This Google Apps Script automates administrative tasks for A-Level Computer Science coursework. It integrates with Google Classroom and Google Drive to create student folders, copy template files (e.g. marking grids, declaration forms), and retrieve coursework submissions from Google Classroom as PDFs.

## Features

- **Get Names and IDs:** Creates student folders and records details in the "Student Info" sheet.
- **Copy Marksheets and Declarations:** Copies template files into each folder, with filenames prepended by student initials.
- **Copy Coursework Submissions:** Copies PDF submissions or converts Google Docs to PDFs from a Google Classroom assignment to a folders for each student.

## Workflow

1. **Setup:**

   - Install the script in your Google Spreadsheet.
   - Enable Google Classroom and Google Drive APIs.
   - Authorise script access.

2. **Get Names and IDs:**

   - Select **Folder Populator > 1. Get names and IDs**.
   - Enter the Google Classroom course URL and root folder ID.
   - The script creates "Student Info" (student name, ID, folder ID) and "Course Info" (course ID, template file IDs) sheets.

3. **Copy Marksheets and Declarations:**

   - Enter template file IDs in "Course Info."
   - Select **Folder Populator > 2. Copy marksheets and declarations**.

4. **Copy Coursework Submissions:**

   - Select **Folder Populator > 3. Copy coursework submissions**.
   - Input the assignment title and prepend string.
   - The script copies PDFs or converts Google Docs to PDFs and places them in folders.

> **Note:** The script processes both students and teachers by default. Adjust `getClassroomMembers` to limit folders to students.

## Installation

1. Open or create a Google Spreadsheet.
2. Go to **Extensions > Apps Script**.
3. Paste the script into the editor.
4. Enable **Google Classroom API** and **Google Drive API** under **Services**.
5. Save the project and run any function to trigger authorisation.

## Additional Functionality

- Use `processFolderAttachmentsForDeclarationsOnly` to extract candidate and centre numbers from documents.

## Troubleshooting

- Ensure correct Google Classroom URL, root folder ID, and assignment title.
- Authorise script access to Google Classroom and Drive.
- Modify `getClassroomMembers` to exclude teachers if needed.

## Contributing

Submit issues or pull requests for improvements.

## License

MIT License.

