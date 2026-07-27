/**
 * BEC Customer Feedback — Google Apps Script Web App
 *
 * Receives POSTed feedback from bec-feedback-form.html and appends it as a
 * row in the "BEC Customer Feedback" spreadsheet.
 *
 * Setup:
 * 1. Open the sheet: https://docs.google.com/spreadsheets/d/1LjV7nStTP-1iLswU-dCL4MprDxqP6HY0g_LNaOKa4hk/edit
 * 2. Extensions -> Apps Script, delete any starter code, paste this file.
 * 3. Deploy -> New deployment -> Web app:
 *      Execute as: Me
 *      Who has access: Anyone
 * 4. Copy the Web App URL (ends with /exec) and paste it into
 *    bec-feedback-form.html as SCRIPT_URL.
 */

const SHEET_ID = "1LjV7nStTP-1iLswU-dCL4MprDxqP6HY0g_LNaOKa4hk";
const SHEET_NAME = "Sheet1";

function doPost(e) {
  const data = JSON.parse(e.postData.contents);

  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_NAME);
  sheet.appendRow([
    new Date(),
    data.branch || "Unknown Branch",
    Number(data.neatness) || "",
    Number(data.service) || "",
    Number(data.staff) || "",
    data.remarks || ""
  ]);

  return ContentService
    .createTextOutput(JSON.stringify({ ok: true }))
    .setMimeType(ContentService.MimeType.JSON);
}
