# MCD Operations Reporting System

## Introduction
The **MCD Operations Reporting System** is an integrated suite of tools designed to monitor, report, and resolve issues related to **Basic Ambient Features & General Upkeep** across various property locations. 

The system digitizes the quality control workflow, moving away from manual tracking to a real-time environment where inspectors can flag deficiencies and agencies can report rectifications immediately via mobile devices.

### Core Components:
*   **Operations Landing Hub:** A central gateway for selecting specific checklists (Ambient Features, CR Fixtures, etc.) and navigating between Inspection and Rectification tasks.
*   **Inspection Portal:** A mobile-responsive digital form used by Housekeeping, Security, and Technical teams to check for deficiencies, take photos of issues, and log them into property-specific databases.
*   **Rectification Portal:** A dedicated interface for agencies to view pending deficiencies, upload "after" photos of fixed items, and submit completion reports.
*   **Automated Summary Engine:** A backend logic script that aggregates data from dozens of property tabs to generate daily and monthly performance metrics (Total Deficient vs. Total Rectified).

---

## Installation Instructions

### 1. Spreadsheet Setup
*   Create a Google Spreadsheet with tabs for each property (e.g., `EWCEastwoodCity`, `LCTLuckyChinatown`).
*   Ensure a `Summary` tab exists to host the aggregated reports.
*   Property tabs must follow a specific column structure (Reference numbers in Column K, Timestamps in Column B, etc.).

### 2. Google Drive Setup
You will need two dedicated folders in Google Drive:
*   **Inspection Photos Folder:** To host images of reported deficiencies.
*   **Rectification Photos Folder:** To host images of resolved issues.
*   Copy the **Folder IDs** from the browser URL to update the script constants.

### 3. Script Deployment
This system consists of three primary Apps Script deployments:
1.  **Landing Page:** Deploy `Landingpage.html` as a Web App.
2.  **Inspection Form:** Deploy `Inspection.html` with its corresponding `.gs` code.
3.  **Rectification Form:** Deploy `Rectificationpage.html` with its corresponding `.gs` code.

**Important Script Constants:**
Update the following IDs in your `.gs` files:
*   `targetSpreadsheetId`: The ID of your master database.
*   `folderId`: The ID of your image storage folder.

---

## Usage Examples

### Inspection: Reporting a Deficiency
When an inspector finds an issue (e.g., a busted light), they select the property and check the box in the grid. The frontend handles the file conversion to Base64 for the server:

```javascript
/**
 * Example: Handling Deficiency Submission
 * Converts image to base64 and maps it to a unique REF number.
 */
const handleFileChange = (event, key) => {
    const file = event.target.files[0];
    const reader = new FileReader();
    reader.onload = (e) => {
        uploadedFiles[key] = {
            name: file.name,
            data: e.target.result.split(",")[1], // Base64 Data
            mimeType: file.type
        };
    };
    reader.readAsDataURL(file);
};
```

### Backend: Monthly Summarization
The following script logic automatically populates the `Summary` tab with performance data across all months:

```javascript
/**
 * Example: Summarizing Counts
 * Loops through all property sheets and counts 
 * non-empty Deficiency and Rectification cells.
 */
function summarizeCounts() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();
  
  sheets.forEach(sheet => {
    const data = sheet.getDataRange().getValues();
    // Logic to count deficiencies based on Column K 
    // and Rectifications based on Column Q...
  });
  
  // Writes data starting at row 30 of the Summary tab.
}
```

---

## Tech Stack
*   **Frontend:** Vue.js 3 (Composition API), Bootstrap 5.3.
*   **Backend:** Google Apps Script (V8 Engine).
*   **Database:** Google Sheets.
*   **Storage:** Google Drive API (Base64 Image Reconstruction).

---

> **Warning**  
> **- For MCD Internal Use Only**  
> This application contains proprietary property management logic and internal mall data. Unauthorized distribution or reproduction of the source code or associated spreadsheets is strictly prohibited.
