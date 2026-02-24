# Coursework Collator (Computer Science + ICT Content Creation)

This Google Apps Script helps collate coursework from Google Classroom into Google Drive folder structures.
It currently supports two workflows with separate UI submenus:

- **Computer Science**
- **ICT Content Creation**

## Menu Structure

From your spreadsheet, open:

- **Coursework Collator > Computer Science**
  - `Setup class and student folders`
  - `Copy marksheets and declarations`
  - `Copy coursework submissions`
  - `Process declarations only`
  - `Merge PDFs for all students`

- **Coursework Collator > ICT Content Creation**
  - `Setup Student Info headers`
  - `Fetch roster to Student Info`
  - `Create folders from Student Info`
  - `Collate assignment content`

## Prerequisites

- Google Workspace account with Classroom and Drive access.
- Apps Script project attached to a Google Sheet.
- Advanced Google services enabled in Apps Script:
  - Google Classroom API
  - Google Drive API

## Computer Science Workflow

1. Run **Setup class and student folders**.
   - Prompts for Classroom URL and root Drive folder ID.
   - Populates `Student Info` and `Course Info`.
   - Creates per-student folders (name format from Classroom full name).
2. Add template file IDs to `Course Info` (starting row 3, column A).
3. Run **Copy marksheets and declarations**.
   - Copies template files into each student folder.
   - Prepends filenames with generated initials.
4. Run **Copy coursework submissions**.
   - Prompts for assignment title and prepend string.
   - Copies zip files and PDFs.
   - If no PDFs are present for a submission, Google Docs are converted to PDF and copied.
5. Use **Process declarations only** and **Merge PDFs for all students** as needed for declaration/sample processing.

## ICT Content Creation Workflow

This workflow is intentionally staged so Student IDs can be added before folder creation.

1. (Optional) Run **Setup Student Info headers**.
   - Writes headers: `Student ID`, `Name`, `User ID`, `Folder ID`.
2. Run **Fetch roster to Student Info**.
   - Prompts for Classroom URL.
   - Pulls all students from Classroom into `Student Info` (`Name`, `User ID`).
   - Leaves `Student ID` and `Folder ID` empty for manual completion/creation.
3. Fill in `Student ID` values in `Student Info`.
4. Run **Create folders from Student Info**.
   - Prompts for root Drive folder ID.
   - Creates folders named `Student ID Name`.
   - Writes created folder IDs back to `Folder ID`.
   - Validates missing/duplicate Student IDs before creating folders.
5. Run **Collate assignment content**.
   - Prompts for Classroom URL and assignment title.
   - Uses existing `Folder ID` values to place files.
   - Copies all Drive attachments with original names.
   - Converts Google Docs attachments to PDF with the same base filename.

## Data Sheets Used

- `Student Info`
  - Computer Science flow: typically `Name`, `User ID`, `Folder ID`
  - ICT flow: `Student ID`, `Name`, `User ID`, `Folder ID`
- `Course Info`
  - Row 1 stores `Course ID`
  - Computer Science template file IDs are read from row 3 onward
- `Prefixes`
  - Used by declaration/PDF merge workflows

## Notes

- Classroom roster retrieval is paginated and fetches all students across pages.
- Current roster retrieval uses `Classroom.Courses.Students`, so it fetches students (not teachers).
- Many copy/folder operations prompt before replacing existing files/folders.

## Troubleshooting

- **Invalid Classroom URL / Assignment title**: verify teacher access and exact title match.
- **Missing ICT folder IDs during collation**: run `Create folders from Student Info` first.
- **Duplicate/missing Student IDs**: fix rows in `Student Info` before creating ICT folders.
- **Permissions errors**: re-authorize script and confirm enabled APIs.

## License

This project is licensed under the MIT License.
