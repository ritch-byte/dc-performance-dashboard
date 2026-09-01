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
 * 5. Same again for the "Attendance" tab once a leader has tagged a day, and give that
 *    CSV link to the dashboard so attendance tags are shared rather than per browser.
 * 6. Same again for the "Schedule" tab once a schedule has been uploaded, so every leader
 *    sees the rostered shifts and off days rather than only the person who uploaded it.
 *
 * The dashboard posts to the SAME /exec URL for both logging and AI.
 */

const SHEET_NAME = 'Logs';
const HEADERS = ['timestamp', 'date', 'weekof', 'squad', 'rep', 'level', 'topic', 'coach', 'summary', 'type'];

// ── Spend guard on the AI relay ──────────────────────────────────────────
//
// This endpoint is deployed "Anyone" and its URL ships inside the Spiel Builder's
// JavaScript bundle on a public GitHub Pages site, so the URL is readable by anyone
// who views source, and the body arriving here is untrusted.
//
// Until now it forwarded data.model and data.max_tokens straight through. One request
// naming the most expensive model at 64k output costs about $1.60 against the org
// credits, and a loop empties the balance in minutes. These two constants bound what a
// single call can cost; they cannot stop someone calling it.
//
// The only real fix for a public static caller is to rotate the deployment URL, which
// kills whatever is using the current one. Deploy → New deployment (not Manage/Edit) to
// get a fresh /exec, then update AI_RELAY_URL in call-script-v2/src/lib/ai.ts.
const ALLOWED_MODELS = ['claude-haiku-4-5-20251001', 'claude-sonnet-4-6'];
const DEFAULT_MODEL = 'claude-haiku-4-5-20251001';
const MAX_TOKENS_CEILING = 4000;

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
      // Clamp before spending. An unrecognised model falls back to the cheap one rather
      // than erroring, so a legitimate caller never breaks; the ceiling is what stops one
      // request costing dollars. `tools` is dropped because nothing here uses server-side
      // search and it bills separately per request.
      var model = ALLOWED_MODELS.indexOf(data.model) !== -1 ? data.model : DEFAULT_MODEL;
      var maxTokens = Math.min(Number(data.max_tokens) || 1024, MAX_TOKENS_CEILING);
      if (data.model && model !== data.model) {
        console.warn('clamped model: ' + data.model);
      }
      if (Number(data.max_tokens) > MAX_TOKENS_CEILING) {
        console.warn('clamped max_tokens: ' + data.max_tokens);
      }

      const resp = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', {
        method: 'post',
        contentType: 'application/json',
        headers: { 'x-api-key': key, 'anthropic-version': '2023-06-01' },
        payload: JSON.stringify({
          model: model,
          max_tokens: maxTokens,
          messages: data.messages,
        }),
        muteHttpExceptions: true,
      });
      return ContentService.createTextOutput(resp.getContentText()).setMimeType(ContentService.MimeType.JSON);
    }

    // ── SDR Scorecard events: route to a separate "Scorecards" tab ──
    // Each generation posts {record:'scorecard', event:'sent', ...}; each
    // acknowledgement posts {record:'scorecard', event:'acknowledged', ...}.
    if (data && data.record === 'scorecard') {
      const SC_HEADERS = ['sdr_id', 'sdr_name', 'team', 'month', 'sent_day', 'sent_at', 'deadline_at', 'event', 'event_at'];
      const ssc = SpreadsheetApp.getActiveSpreadsheet();
      let sc = ssc.getSheetByName('Scorecards');
      if (!sc) { sc = ssc.insertSheet('Scorecards'); }
      if (sc.getLastRow() === 0) { sc.appendRow(SC_HEADERS); }
      sc.appendRow(SC_HEADERS.map(function (h) { return data[h] || ''; }));
      return json_({ ok: true });
    }

    // ── Attendance tags: route to a separate "Attendance" tab ──
    // A team leader tagging a day in the dashboard posts
    // {record:'attendance', date, sdr, sdr_id, team, status, taggedBy, at}.
    // Rows are appended rather than updated, so the tab is a log: re-tagging the same
    // person and day writes a second row, and the dashboard treats the newest as current.
    // That keeps who changed what, and when, instead of overwriting the earlier call.
    if (data && data.record === 'attendance') {
      const AT_HEADERS = ['at', 'date', 'sdr', 'sdr_id', 'team', 'status', 'taggedBy'];
      const ssa = SpreadsheetApp.getActiveSpreadsheet();
      let at = ssa.getSheetByName('Attendance');
      if (!at) { at = ssa.insertSheet('Attendance'); }
      if (at.getLastRow() === 0) { at.appendRow(AT_HEADERS); }
      at.appendRow(AT_HEADERS.map(function (h) { return data[h] || ''; }));
      return json_({ ok: true });
    }

    // ── Floor schedule: route to a separate "Schedule" tab ──
    // One upload posts every SDR row at once as {record:'schedule', period, uploadedAt,
    // uploadedBy, rows:[...]}. Rows are appended rather than replaced, so re-uploading a month
    // leaves the earlier version in place and the dashboard reads whichever upload is newest
    // for that period. That keeps a history of what the floor was told, which matters when an
    // appraisal is argued months later.
    if (data && data.record === 'schedule') {
      const SH_HEADERS = ['uploadedAt', 'uploadedBy', 'period', 'emp', 'name', 'batch', 'team', 'shift', 'off', 'rest', 'note'];
      const sss = SpreadsheetApp.getActiveSpreadsheet();
      let sd = sss.getSheetByName('Schedule');
      if (!sd) { sd = sss.insertSheet('Schedule'); }
      if (sd.getLastRow() === 0) { sd.appendRow(SH_HEADERS); }
      const rows = (data.rows || []).map(function (r) {
        return [data.uploadedAt || '', data.uploadedBy || '', data.period || '',
                r.emp || '', r.name || '', r.batch || '', r.team || '',
                r.shift || '', r.off || '', r.rest || '', r.note || ''];
      });
      if (rows.length) { sd.getRange(sd.getLastRow() + 1, 1, rows.length, SH_HEADERS.length).setValues(rows); }
      return json_({ ok: true, written: rows.length });
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

/**
 * AUTO-ACKNOWLEDGEMENT (optional)
 * ------------------------------------------------------------------
 * Scans Gmail for replies to scorecard emails and logs them as
 * "acknowledged" in the Scorecards tab, so the dashboard flips the
 * status automatically without anyone clicking "Mark ack".
 *
 * SETUP (one time):
 * 1. Paste this whole file. Save.
 * 2. Run `scanScorecardReplies` once from the editor and grant the
 *    Gmail permission it asks for.
 * 3. Triggers (clock icon, left) -> Add Trigger -> function
 *    scanScorecardReplies -> Time-driven -> Minutes timer -> Every 15
 *    minutes -> Save.
 *
 * How it matches: it reads the SDR name from the scorecard subject
 * (`[SDR SCORECARD] Month W## - Name - Band - from X`) and, when a
 * reply from anyone other than you appears in that thread, logs an
 * acknowledgement against that SDR's most recent pending scorecard.
 *
 * LIMITATION: it only sees replies that land in THIS account's Gmail,
 * so scorecards must be sent from (or Cc'd to, with reply-all) this
 * same account.
 */
function scanScorecardReplies() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sc = ss.getSheetByName('Scorecards');
  if (!sc || sc.getLastRow() < 2) return;

  var data = sc.getDataRange().getValues();
  var col = {};
  data[0].forEach(function (h, i) { col[h] = i; });

  var acked = {};        // "sdr_id|sent_day" already acknowledged
  var sentByName = {};   // lowercased sdr_name -> [{sdr_id, sent_day, sent_at}]
  for (var r = 1; r < data.length; r++) {
    var row = data[r];
    var ev = String(row[col.event] || '').toLowerCase();
    var key = row[col.sdr_id] + '|' + row[col.sent_day];
    if (ev === 'acknowledged') { acked[key] = true; continue; }
    var nm = String(row[col.sdr_name] || '').trim().toLowerCase();
    (sentByName[nm] = sentByName[nm] || []).push({ sdr_id: row[col.sdr_id], sent_day: row[col.sent_day], sent_at: row[col.sent_at] });
  }

  var me = Session.getActiveUser().getEmail();
  var label = GmailApp.getUserLabelByName('ScorecardAcked') || GmailApp.createLabel('ScorecardAcked');
  var threads = GmailApp.search('subject:"[SDR SCORECARD]" newer_than:21d -label:ScorecardAcked');
  var SC_HEADERS = ['sdr_id', 'sdr_name', 'team', 'month', 'sent_day', 'sent_at', 'deadline_at', 'event', 'event_at'];

  for (var t = 0; t < threads.length; t++) {
    var th = threads[t];
    var msgs = th.getMessages();
    var subj = msgs[0].getSubject() || '';
    if (subj.indexOf('[SDR SCORECARD]') < 0) continue;

    var replied = false;
    for (var m = 1; m < msgs.length; m++) {
      var from = msgs[m].getFrom() || '';
      if (me && from.indexOf(me) >= 0) continue;  // skip my own messages
      replied = true; break;
    }
    if (!replied) continue;

    var parts = subj.split('·');            // split on the middle dot
    var name = parts.length >= 3 ? String(parts[2]).trim() : '';
    if (!name) { th.addLabel(label); continue; }

    var list = (sentByName[name.toLowerCase()] || []).slice()
      .sort(function (a, b) { return String(b.sent_at).localeCompare(String(a.sent_at)); });
    var target = null;
    for (var i = 0; i < list.length; i++) {
      if (!acked[list[i].sdr_id + '|' + list[i].sent_day]) { target = list[i]; break; }
    }
    if (!target) { th.addLabel(label); continue; }

    var out = { sdr_id: target.sdr_id, sdr_name: name, sent_day: target.sent_day, event: 'acknowledged', event_at: new Date().toISOString() };
    sc.appendRow(SC_HEADERS.map(function (h) { return out[h] || ''; }));
    acked[target.sdr_id + '|' + target.sent_day] = true;
    th.addLabel(label);
  }
}
