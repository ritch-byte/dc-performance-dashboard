/**
 * OA Coaching Log Tracker + AI proxy — Google Apps Script Web App
 * --------------------------------------------------------------
 * ONE endpoint that does two jobs, so the dashboard needs no Netlify:
 *   1. Appends each coaching log to the "Logs" sheet  (payload has no `messages`)
 *   2. Relays coaching-AI calls to the Anthropic API  (payload has `messages`)
 *
 * SETUP (one time):
 * 1. Sheet → Extensions → Apps Script. Paste THIS whole file. Save.
 * 2. Add your Anthropic key so the AI relay can authenticate:
 *      Project Settings (⚙, left) → Script properties → Add script property
 *      Property: ANTHROPIC_API_KEY    Value: <your key from console.anthropic.com>
 * 3. Deploy → Manage deployments → ✏️ Edit → Version: New version → Deploy.
 *    (Execute as: Me · Who has access: Anyone.) The /exec URL stays the same.
 * 4. In the Sheet: File → Share → Publish to web → the "Logs" tab → CSV  (for reads).
 *
 * The dashboard posts to the SAME /exec URL for both logging and AI.
 */

const SHEET_NAME = 'Logs';
const HEADERS = ['timestamp', 'date', 'weekof', 'squad', 'rep', 'level', 'topic', 'coach', 'summary', 'type'];

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);

    // ── AI relay: any payload carrying `messages` is a Claude request ──
    if (data && data.messages) {
      const key = PropertiesService.getScriptProperties().getProperty('ANTHROPIC_API_KEY');
      if (!key) return json_({ error: 'ANTHROPIC_API_KEY script property is not set' });
      const resp = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', {
        method: 'post',
        contentType: 'application/json',
        headers: { 'x-api-key': key, 'anthropic-version': '2023-06-01' },
        payload: JSON.stringify({
          model: data.model || 'claude-haiku-4-5-20251001',
          max_tokens: data.max_tokens || 1024,
          messages: data.messages,
        }),
        muteHttpExceptions: true,
      });
      return ContentService.createTextOutput(resp.getContentText()).setMimeType(ContentService.MimeType.JSON);
    }

    // ── Otherwise: append a coaching-log row ──
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sh = ss.getSheetByName(SHEET_NAME);
    if (!sh) { sh = ss.insertSheet(SHEET_NAME); }
    if (sh.getLastRow() === 0) { sh.appendRow(HEADERS); }
    sh.appendRow(HEADERS.map(function (h) { return data[h] || ''; }));
    return json_({ ok: true });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  }
}

function doGet() {
  return ContentService.createTextOutput('OA Coaching Log Tracker + AI endpoint is live.');
}
