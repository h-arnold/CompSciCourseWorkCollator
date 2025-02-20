Below is an updated README with a new section near the top that explains a typical workflow for using the script.

---

# Folder Populator for A-Level Computer Science Coursework

This Google Apps Script automates tedious administrative tasks for organising A-Level Computer Science coursework. It interacts with Google Classroom and Google Drive to create individual student folders, copy essential template files (such as the marking grid and declaration form) into each folder, and retrieve coursework submissions—either as PDFs or as Google Docs converted to PDFs.

## Overview

The script provides three main functions, accessible via a custom menu in your Google Spreadsheet:

- **Get Names and IDs:** Retrieves members from a Google Classroom course, creates a folder for each (under a specified root folder), and records their details in a "Student Info" sheet.
- **Copy Marksheets and Declarations:** Copies template files (e.g. marking grid and declaration form) into each student’s folder, prepending filenames with the student’s initials.
- **Copy Coursework Submissions:** Pulls coursework attachments from a Google Classroom assignment. It copies PDFs directly or converts Google Docs to PDFs before placing them in the corresponding student folders.

## Typical Workflow

1. **Initial Setup:**  
   - Install the script in your Google Spreadsheet.
   - Enable the Google Classroom and Google Drive APIs.
   - Authorise the script to access your account.

2. **Get Names and IDs:**  
   - From the custom **Folder Populator** menu, select **1. Get names and IDs**.
   - Enter the Google Classroom course URL and the root folder ID where student folders should be created.
   - The script creates (or refreshes) two sheets:  
     - **Student Info:** Lists each member’s name, user ID, and created folder ID.
     - **Course Info:** Displays the course ID and (if needed) template file IDs.

3. **Copy Marksheets and Declarations:**  
   - Add the template file IDs (for the marking grid and declaration form) into the "Course Info" sheet.
   - From the menu, select **2. Copy marksheets and declarations**.
   - The script copies each template into every student’s folder, renaming the files by prepending the student’s initials.

4. **Copy Coursework Submissions:**  
   - From the menu, select **3. Copy coursework submissions**.
   - Enter the assignment title and a prepend string when prompted.
   - The script fetches student submissions from the specified assignment. It will copy any PDF attachments directly; if none are found, attached Google Docs are converted to PDFs before being copied into the student folders.

> **Note:** The current implementation retrieves both students and teachers from the Google Classroom course. To restrict folder creation to students only, adjust the `getClassroomMembers` function as required.

## Detailed Instructions

### Prerequisites

- A Google Workspace account with access to Google Classroom and Google Drive.
- Enabled Google Classroom API and Google Drive API in your Google Apps Script project.
- A Google Spreadsheet to install and run the script.

### Setup and Installation

1. **Open Your Google Spreadsheet:**  
   Go to [Google Sheets](https://sheets.google.com) and open or create the spreadsheet for this project.

2. **Access the Script Editor:**  
   Click on **Extensions > Apps Script**.

3. **Copy the Code:**  
   Paste the complete script into the editor.

4. **Enable APIs:**  
   Ensure that the **Google Classroom API** and **Google Drive API** are enabled under **Services** in the project menu.

5. **Save and Authorise:**  
   Save the project with a meaningful name (e.g. "Folder Populator for Coursework") and run any function (such as `onOpen`) to trigger the authorisation flow.

### Using the Script

- **Get Names and IDs:**  
  - Click **Folder Populator > 1. Get names and IDs**.
  - Provide the Google Classroom URL and the root folder ID when prompted.
  - Two sheets, "Student Info" and "Course Info", are generated with the relevant details.

- **Copy Marksheets and Declarations:**  
  - Enter the template file IDs (for the marking grid and declaration form) into the "Course Info" sheet.
  - Click **Folder Populator > 2. Copy marksheets and declarations**.
  - The script copies each template into every student’s folder, renaming them with the student’s initials.

- **Copy Coursework Submissions:**  
  - Click **Folder Populator > 3. Copy coursework submissions**.
  - Input the assignment title and a prepend string.
  - The script locates the assignment in Google Classroom, processes each student’s submission, and copies or converts files to PDFs, placing them in the corresponding student folders.

### Additional Functionality

There is also a variant function (`processFolderAttachmentsForDeclarationsOnly`) designed for processing declarations differently. This function extracts candidate and centre numbers from document text to customise file names. Use or modify this function as needed.

## Troubleshooting

- **Invalid Inputs:**  
  - If an incorrect Google Classroom URL or root folder ID is entered, the script will alert you and cancel the operation.
  - An invalid assignment title will also trigger an alert.

- **API Permissions:**  
  - Ensure that the script is authorised to access Google Classroom and Google Drive.

- **Teacher Folders:**  
  - The current implementation includes both students and teachers. Modify the `getClassroomMembers` function if you wish to create folders only for students.

## Contributing

Feel free to open issues or submit pull requests with suggestions or improvements. Contributions that refine the script or expand its functionality are most welcome.

## License

This project is licensed under the MIT License.

---

This README should provide comprehensive guidance on setting up, configuring, and using the script to streamline your coursework administration tasks. Enjoy your more efficient workflow!
