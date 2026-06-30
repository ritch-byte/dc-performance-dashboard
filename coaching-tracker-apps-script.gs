/**
 * OA Coaching Log Tracker — Google Apps Script Web App
 * ----------------------------------------------------
 * Appends each coaching log sent from the DC dashboard to a "Logs" sheet.
 *
 * SETUP (one time):
 * 1. Create a Google Sheet named e.g. "OA Coaching Log Tracker".
 * 2. Extensions → Apps Script. Delete the sample code, paste THIS file. Save.
 * 3. Deploy → New deployment → gear ⚙ → "Web app".
 *      - Execute as:   Me
 *      - Who has access: Anyone
 *    Deploy → authorize → copy the "Web app" URL (ends with /exec)  ← COACH_LOG_POST_URL
 * 4. Back in the Sheet: File → Share → Publish to web →
 *      - Link tab → choose the "Logs" sheet → "Comma-separated values (.csv)"
 *      - Publish → copy that URL  ← COACH_LOG_CSV_URL
 * 5. Paste both URLs into the two constants at the top of the tracker block
 *    in index.html (COACH_LOG_POST_URL / COACH_LOG_CSV_URL), or send them over.
 *
 * Note: the published CSV can take ~1–5 min to reflect new rows (Google's
 * republish cadence). The dashboard shows new logs instantly in your own
 * session and pulls the authoritative sheet on Refresh / reload.
 */

const SHEET_NAME = 'Logs';
const HEADERS = ['timestamp', 'date', 'weekof', 'squad', 'rep', 'level', 'topic', 'coach', 'summary'];

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sh = ss.getSheetByName(SHEET_NAME);
    if (!sh) { sh = ss.insertSheet(SHEET_NAME); }
    if (sh.getLastRow() === 0) { sh.appendRow(HEADERS); }
    sh.appendRow(HEADERS.map(function (h) { return data[h] || ''; }));
    return ContentService
      .createTextOutput(JSON.stringify({ ok: true }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ ok: false, error: String(err) }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function doGet() {
  return ContentService.createTextOutput('OA Coaching Log Tracker endpoint is live.');
}
