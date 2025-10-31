# Lab Gradesheet Generator for CSE250, CSE251, and CSE350

This **Google Apps Script** automates the generation of lab gradesheets for each section of the courses **CSE250, CSE251, and CSE350**. It replicates template sheets, forms, and response sheets, links everything on the central gradesheet spreadsheet, and provides access to co-faculty members.

---

## Features

- Automatically generates **gradesheet spreadsheets** for each section.
- Creates and links **submission forms** for lab reports.
- Sets up **response sheets** for each form.
- Provides **co-faculty and faculty access** to the gradesheets.
- Ensures proper **permissions** are applied to all generated files and folders.
- Checks for previously generated sections to avoid duplication.

---

## User Sheet

The script is run using a **User Sheet**, where you select the course, enter the section number, and provide other necessary inputs.

- [Open User Sheet](https://docs.google.com/spreadsheets/d/1GgvR9vRk68b0s5l2F-pdIhcCs-KZjc3VpDM-xpdWcM0/edit?usp=sharing)  

> **Note:** The course must be selected in cell `B1` of the User Sheet, and the section number is entered when prompted by the script.

---

## Prerequisites

Before running the script:

1. **Add Drive API v2** from **Services** in Apps Script.  
   - Required because `Drive.Permissions.insert` will not work without it.
2. Ensure the faculty members have **edit access** to the gradesheet folders.
3. For debugging within Apps Script, bypass the `getUI()` function.

---

## Usage

1. Open the script in **Google Apps Script** attached to your master spreadsheet.
2. Open the **User Sheet** and select the desired course.  
3. Run the `main()` function.
4. Enter the **section number** when prompted.
5. The script will:
   - Create a folder for the section.
   - Copy the gradesheet template and forms.
   - Link the forms and response sheets to the gradesheet.
   - Apply permissions to the co-faculty and the generating user.

After completion, a toast message will confirm that the process is done, and emails will be sent automatically if notifications are enabled.

---

## Notes

- Always check if a section has already been generated. The script will warn you and ask for confirmation before overwriting.
- The script uses `Drive.Permissions.insert` to manage access, so the **Drive API v2** must be enabled.
- For debugging in Apps Script, bypass `getUI()` to avoid blocking prompts.

---

## File Structure

- `Gradesheet Templates` – Template spreadsheets for each course.
- `Forms` – Lab report submission forms.
- `Response Sheets` – Linked response spreadsheets for each form.
- `Script` – The Google Apps Script file that automates all of the above.
