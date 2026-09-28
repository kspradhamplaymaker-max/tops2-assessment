/**
 * TOPS-2 assessment backend.
 *
 * Deployed as a Google Apps Script Web App, this receives one POST per
 * completed assessment from the page's sendToCoach() function, saves the
 * respondent's PDF report to a "TOPS-2 Submissions" folder in Drive, emails
 * the coach a link to it with a text summary, and logs the submission to a
 * Google Sheet.
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

// The PDF is saved to Drive and linked from the email rather than
// attached: Gmail rejected (5.7.1) the attachment version of this email.
function sendEmail(data) {
  var name = data.name || "(not given)";
  var subject = "TOPS-2 Mental Performance Profile (" + name + ")";
  var summary = data.reportText || "A new TOPS-2 assessment was completed, but no report text was attached.";
  var intro;
  var pdfUrl = null;
  // The page sends the same PDF the respondent can download, base64-encoded.
  // If it's missing or can't be saved, the text summary still goes out.
  if (data.pdfBase64) {
    try {
      pdfUrl = savePdfToDrive(data.pdfBase64, data.pdfFilename);
    } catch (err) {
      console.error("Saving PDF to Drive failed: " + err);
    }
  }
  if (pdfUrl) {
    intro = name + " completed the TOPS-2 assessment. Their full PDF report is saved in your Drive (" +
      SUBMISSIONS_FOLDER + "):\n" + pdfUrl + "\n\nA text summary follows.";
  } else if (data.pdfBase64) {
    intro = name + " completed the TOPS-2 assessment. (Their PDF couldn't be saved to Drive, so only the text summary is included.)";
  } else {
    intro = name + " completed the TOPS-2 assessment. (The PDF couldn't be generated in their browser, so only the text summary is included.)";
  }
  GmailApp.sendEmail(COACH_EMAIL, subject, intro + "\n\n" + summary);
}

var SUBMISSIONS_FOLDER = "TOPS-2 Submissions";

// Drive access goes through the Drive API (the "Drive" advanced service,
// enabled in appsscript.json), not DriveApp: DriveApp demands the full
// "all of your Drive files" scope even to create a folder, while the
// Drive API works with the narrow drive.file scope, which only covers
// files and folders this script created itself.

// Saves the PDF into the "TOPS-2 Submissions" Drive folder (created on
// first use) and returns the file's link. The file stays private to the
// script owner, who is also the coach receiving the email.
function savePdfToDrive(pdfBase64, pdfFilename) {
  var stamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HHmm");
  var filename = stamp + " " + (pdfFilename || "mental-performance-profile.pdf");
  var blob = Utilities.newBlob(Utilities.base64Decode(pdfBase64), "application/pdf", filename);
  var file = Drive.Files.create(
    { name: filename, mimeType: "application/pdf", parents: [getSubmissionsFolderId()] },
    blob,
    { fields: "id,webViewLink" }
  );
  return file.webViewLink;
}

// The folder is created once and then opened by its stored ID (Script
// Properties), never searched for by name, since drive.file can't see
// folders by name. It's only recreated if the stored ID is missing or the
// folder was trashed or deleted. A lock stops two simultaneous first
// submissions from each creating a folder.
function getSubmissionsFolderId() {
  var props = PropertiesService.getScriptProperties();
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var id = props.getProperty("SUBMISSIONS_FOLDER_ID");
    if (id) {
      try {
        var saved = Drive.Files.get(id, { fields: "id,trashed" });
        if (!saved.trashed) return saved.id;
      } catch (err) { /* deleted or no longer accessible; recreate below */ }
    }
    var folder = Drive.Files.create(
      { name: SUBMISSIONS_FOLDER, mimeType: "application/vnd.google-apps.folder" },
      null,
      { fields: "id" }
    );
    props.setProperty("SUBMISSIONS_FOLDER_ID", folder.id);
    return folder.id;
  } finally {
    lock.releaseLock();
  }
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
