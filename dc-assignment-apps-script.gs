/**
 * DC Assignment — Google Apps Script Web App
 * ------------------------------------------------------------------
 * Two jobs:
 *   1. syncCalendars()  — on a timer, reads upcoming events from the concierge calendars
 *                         into a "DC Assignments" sheet tab.
 *   2. doPost()         — records who a meeting is assigned to, passcode checked HERE
 *                         rather than only in the browser.
 *
 * This is deliberately a SEPARATE script from the coaching tracker. Its /exec URL ships inside
 * a page the whole floor can open, and the tracker's URL also relays to the Anthropic API, so
 * putting the two together would hand thirty people a way to spend the org's credits. Nothing
 * in this file can spend anything.
 *
 * SETUP (one time)
 * 1. Share all three calendars with the account that owns this script:
 *      Google Calendar → the calendar → Settings and sharing → Share with specific people
 *      → add your own address → "See all event details".
 *    Without this, getCalendarById returns null and the sync writes nothing.
 * 2. New Google Sheet → Extensions → Apps Script → paste this file → Save.
 * 3. Project Settings (gear) → Script properties → Add:
 *      DC_ASSIGN_PASS   = <the passcode leaders will type>
 *    Keep it out of this file. This file lives in a public GitHub repo.
 * 4. Run syncCalendars() once by hand and accept the permission prompts.
 * 5. Triggers (clock icon) → Add trigger → syncCalendars → Time-driven → every 30 minutes.
 * 6. Deploy → New deployment → Web app → Execute as: Me → Who has access: Anyone → Deploy.
 *    Send Claude the /exec URL.
 * 7. In the Sheet: File → Share → Publish to web → "DC Assignments" → CSV → Publish.
 *    Send Claude that link too. The dashboard reads that, and posts to /exec only to assign.
 */

const CALENDARS = [
  'concierge@outsourceaccelerator.com',
  'concierge2@outsourceaccelerator.com',
  'concierge-team@outsourceaccelerator.com'
];
const TAB          = 'DC Assignments';
const DAYS_AHEAD   = 21;
const DAYS_BEHIND  = 1;    // keep yesterday, so a meeting is still there to argue about
const HEADERS = ['eventId', 'calendar', 'title', 'start', 'end', 'durationMin', 'organizer',
                 'guests', 'assignedTo', 'assignedBy', 'assignedAt', 'syncedAt'];

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function sheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(TAB);
  if (!sh) { sh = ss.insertSheet(TAB); }
  if (sh.getLastRow() === 0) { sh.appendRow(HEADERS); }
  return sh;
}

/**
 * Pull the calendars into the tab.
 *
 * Assignments are held in the same rows as the events, so the sync has to put back what it
 * found: a rewrite that dropped the assignee columns would silently unassign the whole floor
 * every half hour. Existing assignments are read first, keyed by event id, and restored.
 */
function syncCalendars() {
  const sh = sheet_();
  const now = new Date();
  const from = new Date(now.getTime() - DAYS_BEHIND * 864e5);
  const to   = new Date(now.getTime() + DAYS_AHEAD  * 864e5);

  // what is already assigned, so the sync does not throw it away
  const kept = {};
  const last = sh.getLastRow();
  if (last > 1) {
    const rows = sh.getRange(2, 1, last - 1, HEADERS.length).getValues();
    rows.forEach(function (r) {
      const id = String(r[0] || '').trim();
      if (id && String(r[8] || '').trim()) {
        kept[id] = { to: r[8], by: r[9], at: r[10] };
      }
    });
  }

  const out = [];
  const missing = [];
  CALENDARS.forEach(function (id) {
    const cal = CalendarApp.getCalendarById(id);
    if (!cal) { missing.push(id); return; }
    cal.getEvents(from, to).forEach(function (ev) {
      const eid = ev.getId();
      const a = kept[eid] || { to: '', by: '', at: '' };
      let guests = '';
      try {
        guests = ev.getGuestList().map(function (g) { return g.getEmail(); }).join(' ');
      } catch (e) { guests = ''; }
      let organizer = '';
      try { organizer = ev.getCreators().join(' '); } catch (e) { organizer = ''; }
      out.push([
        eid, id, ev.getTitle(),
        Utilities.formatDate(ev.getStartTime(), 'Asia/Manila', "yyyy-MM-dd'T'HH:mm"),
        Utilities.formatDate(ev.getEndTime(),   'Asia/Manila', "yyyy-MM-dd'T'HH:mm"),
        Math.round((ev.getEndTime() - ev.getStartTime()) / 60000),
        organizer, guests,
        a.to, a.by, a.at,
        Utilities.formatDate(now, 'Asia/Manila', "yyyy-MM-dd'T'HH:mm")
      ]);
    });
  });

  out.sort(function (a, b) { return String(a[3]).localeCompare(String(b[3])); });

  if (sh.getLastRow() > 1) {
    sh.getRange(2, 1, sh.getLastRow() - 1, HEADERS.length).clearContent();
  }
  if (out.length) {
    sh.getRange(2, 1, out.length, HEADERS.length).setValues(out);
  }
  if (missing.length) {
    // Not an error worth throwing: two calendars syncing is better than none. It is logged
    // because a silently absent calendar looks exactly like a quiet week.
    console.warn('No access to: ' + missing.join(', ') + ' — share them with this account.');
  }
  return { written: out.length, missing: missing };
}

/**
 * Record an assignment.
 *
 * The passcode is checked here as well as in the page. The page can only ever be a deterrent,
 * since anyone can read its source; this is the check that actually holds.
 */
function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    if (!data || data.record !== 'dcassign') {
      return json_({ ok: false, error: 'unknown record type' });
    }
    const want = PropertiesService.getScriptProperties().getProperty('DC_ASSIGN_PASS');
    if (!want) return json_({ ok: false, error: 'DC_ASSIGN_PASS script property is not set' });
    if (String(data.pass || '') !== String(want)) {
      return json_({ ok: false, error: 'wrong passcode' });
    }

    const eid = String(data.eventId || '').trim();
    if (!eid) return json_({ ok: false, error: 'no eventId' });

    const sh = sheet_();
    const last = sh.getLastRow();
    if (last < 2) return json_({ ok: false, error: 'nothing synced yet' });

    const ids = sh.getRange(2, 1, last - 1, 1).getValues();
    for (let i = 0; i < ids.length; i++) {
      if (String(ids[i][0] || '').trim() !== eid) continue;
      const row = i + 2;
      const stamp = Utilities.formatDate(new Date(), 'Asia/Manila', "yyyy-MM-dd'T'HH:mm");
      // Clearing an assignment is a write like any other, so an empty name is allowed through
      // and blanks the row rather than being rejected as a mistake.
      sh.getRange(row, 9,  1, 3).setValues([[
        String(data.assignedTo || ''), String(data.assignedBy || ''),
        String(data.assignedTo || '') ? stamp : ''
      ]]);
      return json_({ ok: true, row: row });
    }
    return json_({ ok: false, error: 'that meeting is not in the sheet' });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  }
}

function doGet() {
  return ContentService.createTextOutput('DC Assignment endpoint is live.');
}
