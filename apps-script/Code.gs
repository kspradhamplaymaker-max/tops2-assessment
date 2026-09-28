/**
 * TOPS-2 assessment backend.
 *
 * Deployed as a Google Apps Script Web App, this receives one POST per
 * completed assessment from the page's sendToCoach() function, emails the
 * coach the respondent's PDF report (attached) with a short text summary,
 * and logs the submission to a Google Sheet.
 *
 * SETUP
 * 1. Go to script.google.com, create a new project, and paste this file's
 *    contents in as Code.gs (replace the default myFunction() stub).
 * 2. Run any function once from the editor (e.g. doPost stub call is not
 *    needed; just click Run on onOpenCheck below) so Google asks you to
 *    authorize Gmail and Sheets access for this script. Approve it.
 * 3. Deploy > New deployment > select type "Web app".
 *    - Execute as: Me
 *    - Who has access: Anyone
 *    Click Deploy, copy the Web App URL it gives you.
 * 4. Paste that URL into EMAIL_ENDPOINT near the top of index.html's
 *    inline script, replacing PASTE_YOUR_APPS_SCRIPT_WEB_APP_URL_HERE.
 * 5. Every submission also lands as a row in a "Submissions" sheet in
 *    whichever Google Sheet this script is bound to (or creates, the
 *    first time it runs, if you attach this script to a new Sheet
 *    instead of a standalone script project).
 */

var COACH_EMAIL = "kspradham.playmaker@gmail.com";
var SHEET_NAME = "Submissions";

function doPost(e) {
  var data = JSON.parse(e.postData.contents);
  try {
    logToSheet(data);
  } catch (err) {
    // Logging is a bonus, not a dependency, so a Sheet issue never
    // blocks the email from sending below.
  }
  sendEmail(data);
  return ContentService.createTextOutput(JSON.stringify({ ok: true }))
    .setMimeType(ContentService.MimeType.JSON);
}

function logToSheet(data) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) return; // standalone script with no bound Sheet; skip logging
  var sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(["Timestamp", "Name", "Answered", "Training (JSON)", "Competition (JSON)"]);
  }
  sheet.appendRow([
    new Date(),
    data.name || "(not given)",
    data.answeredCount || "",
    JSON.stringify(data.practice || {}),
    JSON.stringify(data.competition || {})
  ]);
}

function sendEmail(data) {
  var name = data.name || "(not given)";
  var subject = "TOPS-2 Mental Performance Profile (" + name + ")";
  var summary = data.reportText || "A new TOPS-2 assessment was completed, but no report text was attached.";
  var options = {};
  var intro;
  // The page sends the same PDF the respondent can download, base64-encoded.
  // If it's missing (the browser couldn't build it), the text summary still goes out.
  if (data.pdfBase64) {
    var filename = data.pdfFilename || "mental-performance-profile.pdf";
    options.attachments = [
      Utilities.newBlob(Utilities.base64Decode(data.pdfBase64), "application/pdf", filename)
    ];
    intro = name + " completed the TOPS-2 assessment. Their full report is attached as a PDF; a text summary follows.";
  } else {
    intro = name + " completed the TOPS-2 assessment. (The PDF couldn't be generated in their browser, so only the text summary is included.)";
  }
  GmailApp.sendEmail(COACH_EMAIL, subject, intro + "\n\n" + summary, options);
}

/**
 * Run this once from the Apps Script editor after pasting the code in.
 * It doesn't do anything on its own; running any function is just the
 * trigger Google uses to prompt you for the Gmail and Sheets permission
 * this script needs, before you deploy it.
 */
function onOpenCheck() {
  Logger.log("Ready. If this ran without an authorization prompt, run it once more after reloading the editor.");
}
