/**
 * BEC Customer Feedback — Google Apps Script Web App
 *
 * doPost: receives feedback from bec-feedback-form.html, appends a row to Sheet1.
 * doGet:  ?action=branches returns the branch list from Sheet2 (col A, skipping
 *         the header row) as JSON, used by the form's branch dropdown.
 *
 * Sheet2 layout: A1 = "Branch" (header), A2..An = one branch name per row.
 *
 * After editing this code: Deploy -> Manage deployments -> pencil icon ->
 * Version: "New version" -> Deploy. The /exec URL stays the same.
 */

const SHEET_ID = "1LjV7nStTP-1iLswU-dCL4MprDxqP6HY0g_LNaOKa4hk";
const SHEET_NAME = "Sheet1";
const BRANCH_SHEET = "Sheet2";

function doGet(e) {
  if (e && e.parameter && e.parameter.action === "branches") {
    let branches = [];
    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(BRANCH_SHEET);
    if (sheet) {
      const last = sheet.getLastRow();
      if (last >= 2) {
        branches = sheet.getRange(2, 1, last - 1, 1).getValues()
          .map(function(row) { return String(row[0]).trim(); })
          .filter(function(name) { return name.length > 0; });
      }
    }
    return ContentService
      .createTextOutput(JSON.stringify({ branches: branches }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  return ContentService
    .createTextOutput("BEC Feedback endpoint is live ✓")
    .setMimeType(ContentService.MimeType.TEXT);
}

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
