/**
 * DC Assignment — Google Apps Script Web App
 * ------------------------------------------------------------------
 * Three jobs:
 *   1. syncCalendars()  — on a timer, reads upcoming events from the concierge calendars
 *                         into a "DC Assignments" sheet tab.
 *   2. syncAbsences()   — reads attendance notices sent to sd-attendance@ into an
 *                         "Absences" tab, so a meeting booked by someone who has called
 *                         in sick shows as needing cover. Status and date only, never
 *                         the reason: those mails carry medical and family detail and
 *                         the tab is published to a page the whole floor can open.
 *   3. doPost()         — records who a meeting is assigned to, passcode checked HERE
 *                         rather than only in the browser.
 *
 *   syncAll() runs both syncs, and is what the trigger should call.
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
 * 5. Triggers (clock icon) → Add trigger → syncAll → Time-driven → every 30 minutes.
 *    (If a syncCalendars trigger already exists, change it to syncAll or absences never sync.)
 * 6. Deploy → New deployment → Web app → Execute as: Me → Who has access: Anyone → Deploy.
 *    Send Claude the /exec URL.
 * 7. In the Sheet: File → Share → Publish to web → "DC Assignments" → CSV → Publish.
 *    Then do the same again for the "Absences" tab. Send Claude both links. The dashboard
 *    reads those, and posts to /exec only to assign.
 */

const CALENDARS = [
  'concierge@outsourceaccelerator.com',
  'concierge2@outsourceaccelerator.com',
  'concierge-team@outsourceaccelerator.com'
];
const TAB          = 'DC Assignments';
const DAYS_AHEAD   = 21;
const DAYS_BEHIND  = 1;    // keep yesterday, so a meeting is still there to argue about
const HEADERS = ['eventId', 'calendar', 'partner', 'sdrEmail', 'start', 'end', 'durationMin',
                 'assignedTo', 'assignedBy', 'assignedAt', 'syncedAt'];

/**
 * Who booked it, and for whom.
 *
 * A real event reads: organizer partner.six-eleven@, guests the prospect, the partner's own
 * people, concierge, bdr-team, and exactly one SDR. So the partner comes off the organizer
 * rather than out of the title, which is free text and will not always say "and Partner X",
 * and the SDR is the one internal address that is not a shared mailbox.
 *
 * The prospect's name and email are deliberately NOT written to the sheet. That tab gets
 * published to the web for a page the whole floor can open, and a client's name plus the
 * partner they are being sold to is not something to put on the open internet to save a
 * click. Partner and SDR are enough to say who owns the meeting and whether anyone is awake.
 */
function partnerFrom_(email) {
  const m = String(email || '').match(/^partner\.([^@]+)@/i);
  if (!m) return '';
  const words = m[1].replace(/[-_.]+/g, ' ').split(' ');
  return words.map(function (w) {
    if (!w) return w;
    // A short token with no vowel is an acronym, not a word: HGS OSS, not Hgsoss.
    if (w.length <= 4 && !/[aeiou]/i.test(w)) return w.toUpperCase();
    return w.charAt(0).toUpperCase() + w.slice(1);
  }).join(' ');
}
function sdrFrom_(emails) {
  for (let i = 0; i < emails.length; i++) {
    const e = String(emails[i] || '').toLowerCase();
    if (e.indexOf('@outsourceaccelerator.com') < 0) continue;
    if (/^partner\./.test(e)) continue;
    if (/^concierge/.test(e)) continue;
    if (/^bdr-team@/.test(e)) continue;
    return e;
  }
  return '';
}

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
      if (id && String(r[7] || '').trim()) {
        kept[id] = { to: r[7], by: r[8], at: r[9] };
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
      let guests = [];
      try { guests = ev.getGuestList().map(function (g) { return g.getEmail(); }); } catch (e) {}
      let creators = [];
      try { creators = ev.getCreators(); } catch (e) {}
      let partner = '';
      creators.concat(guests).forEach(function (e) { if (!partner) partner = partnerFrom_(e); });
      out.push([
        eid, id, partner, sdrFrom_(guests),
        Utilities.formatDate(ev.getStartTime(), 'Asia/Manila', "yyyy-MM-dd'T'HH:mm"),
        Utilities.formatDate(ev.getEndTime(),   'Asia/Manila', "yyyy-MM-dd'T'HH:mm"),
        Math.round((ev.getEndTime() - ev.getStartTime()) / 60000),
        a.to, a.by, a.at,
        Utilities.formatDate(now, 'Asia/Manila', "yyyy-MM-dd'T'HH:mm")
      ]);
    });
  });

  out.sort(function (a, b) { return String(a[4]).localeCompare(String(b[4])); });

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
      sh.getRange(row, 8,  1, 3).setValues([[
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

/* ════════════════════════════════════════════════════════════════════════════
 * ABSENCE SYNC
 *
 * SDRs mail sd-attendance@ when they will be absent, late or on leave. A meeting
 * booked by someone who has just called in sick needs covering, and nobody was
 * joining those two facts up.
 *
 * WHAT IS DELIBERATELY NOT STORED: the reason. Those mails carry medical and
 * family detail, someone recovering from illness, a parent's diagnosis. This tab
 * is published to the web for a page the whole floor can open. Status and date
 * answer "does this meeting need cover"; the reason answers nothing this page is
 * entitled to ask, and belongs between a rep and their leader.
 *
 * Matching is on the sender's address, not the name, because it is the same key
 * the calendar gives us and names in these mails are written six different ways.
 * ════════════════════════════════════════════════════════════════════════════ */

const ABS_TAB     = 'Absences';
const ABS_HEADERS = ['date', 'sdrEmail', 'name', 'status', 'notifiedAt', 'syncedAt'];
const ABS_QUERY   = '(to:sd-attendance@outsourceaccelerator.com OR cc:sd-attendance@outsourceaccelerator.com) newer_than:21d';
const ABS_MAX     = 250;

// Which kind of notice this is, from the subject. Anything that is not one of these
// is not an attendance notice at all: that mailbox also receives fulfilment
// summaries, coaching logs and rebooking requests.
function absStatus_(subject) {
  const s = String(subject || '').toLowerCase();
  if (/shift change|rebook|request for|fulfillment|coaching-log/.test(s)) return '';
  if (/absen/.test(s))                                       return 'Absent';
  if (/\bleave\b|\bvl\b|\bsl\b|vacation|sick leave/.test(s))  return 'On leave';
  if (/under\s*time/.test(s))                                 return 'Undertime';
  if (/\blate\b|tardy/.test(s))                               return 'Late';
  return '';
}

const ABS_MONTHS = ['january', 'february', 'march', 'april', 'may', 'june', 'july',
                    'august', 'september', 'october', 'november', 'december'];

// 09/08/2026, 9-8-2026, "September 8, 2026", "Sept 8 2026". Returns yyyy-MM-dd or ''.
function absDate_(raw, fallbackYear) {
  const s = String(raw || '').trim();
  let m = s.match(/(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2,4})/);
  if (m) {
    let y = Number(m[3]);
    if (y < 100) { y += 2000; }
    return Utilities.formatDate(new Date(y, Number(m[1]) - 1, Number(m[2])), 'Asia/Manila', 'yyyy-MM-dd');
  }
  m = s.match(/([A-Za-z]{3,9})\.?\s+(\d{1,2})(?:\s*,)?\s*(\d{4})?/);
  if (m) {
    const want = m[1].toLowerCase();
    let idx = -1;
    for (let i = 0; i < ABS_MONTHS.length; i++) {
      if (ABS_MONTHS[i].indexOf(want) === 0) { idx = i; break; }
    }
    if (idx >= 0) {
      const y = m[3] ? Number(m[3]) : fallbackYear;
      return Utilities.formatDate(new Date(y, idx, Number(m[2])), 'Asia/Manila', 'yyyy-MM-dd');
    }
  }
  return '';
}

// The date the notice is ABOUT, which is not the date it was sent: a leave request
// filed on Friday for Monday would otherwise mark the wrong day off.
function absFieldDate_(body, label, fallbackYear) {
  const re = new RegExp(label + '\\s*[:\\-]\\s*([^\\r\\n]{0,60})', 'i');
  const m = String(body || '').match(re);
  return m ? absDate_(m[1], fallbackYear) : '';
}

function absName_(body, email) {
  const m = String(body || '').match(/(?:full\s*name|name)\s*[:\-]\s*([^\r\n]{2,60})/i);
  if (m) { return m[1].trim().replace(/\s+/g, ' '); }
  return String(email || '').split('@')[0];
}

function syncAbsences() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(ABS_TAB);
  if (!sh) { sh = ss.insertSheet(ABS_TAB); }
  if (sh.getLastRow() === 0) { sh.appendRow(ABS_HEADERS); }

  const now = new Date();
  const stamp = Utilities.formatDate(now, 'Asia/Manila', "yyyy-MM-dd'T'HH:mm");
  const seen = {};
  let scanned = 0, notices = 0;

  GmailApp.search(ABS_QUERY, 0, ABS_MAX).forEach(function (thread) {
    thread.getMessages().forEach(function (msg) {
      scanned++;
      const status = absStatus_(msg.getSubject());
      if (!status) { return; }
      const hit = String(msg.getFrom()).match(/[\w.\-+]+@[\w.\-]+/);
      if (!hit) { return; }
      const email = hit[0].toLowerCase();
      // Only a person's own notice counts. A leader forwarding a batch on behalf of
      // several SDRs is about somebody else, and crediting it to the sender would mark
      // the wrong person off.
      if (email.indexOf('@outsourceaccelerator.com') < 0) { return; }

      let body = '';
      try { body = msg.getPlainBody(); } catch (e) { body = ''; }
      const sentYear = Number(Utilities.formatDate(msg.getDate(), 'Asia/Manila', 'yyyy'));

      const day1 = absFieldDate_(body, '(?:date \\(shift date\\)|shift\\s*date)', sentYear)
                || absFieldDate_(body, 'date', sentYear)
                || Utilities.formatDate(msg.getDate(), 'Asia/Manila', 'yyyy-MM-dd');
      const ret = absFieldDate_(body, 'date of return', sentYear);

      // A return date makes this a span. It is the day they are BACK, so the last day
      // off is the day before it, and a next-day return means a single day off.
      const days = [day1];
      if (ret && ret > day1) {
        const p = day1.split('-').map(Number);
        const d = new Date(p[0], p[1] - 1, p[2]);
        for (let i = 0; i < 30; i++) {
          d.setDate(d.getDate() + 1);
          const iso = Utilities.formatDate(d, 'Asia/Manila', 'yyyy-MM-dd');
          if (iso >= ret) { break; }
          days.push(iso);
        }
      }

      const name = absName_(body, email);
      notices++;
      days.forEach(function (day) {
        const key = email + '|' + day;
        // Newest notice for a person and day wins: a late that later becomes an absence
        // should read as an absence.
        if (seen[key] && seen[key].at >= msg.getDate()) { return; }
        seen[key] = { at: msg.getDate(),
                      row: [day, email, name, status,
                            Utilities.formatDate(msg.getDate(), 'Asia/Manila', "yyyy-MM-dd'T'HH:mm"),
                            stamp] };
      });
    });
  });

  const out = [];
  Object.keys(seen).forEach(function (k) { out.push(seen[k].row); });
  out.sort(function (a, b) { return String(a[0]).localeCompare(String(b[0])); });

  if (sh.getLastRow() > 1) {
    sh.getRange(2, 1, sh.getLastRow() - 1, ABS_HEADERS.length).clearContent();
  }
  if (out.length) {
    sh.getRange(2, 1, out.length, ABS_HEADERS.length).setValues(out);
  }
  console.log('scanned ' + scanned + ' messages, ' + notices + ' were attendance notices, '
    + out.length + ' person-days written');
  return { scanned: scanned, notices: notices, rows: out.length };
}

// One trigger for both, so there is one thing to schedule and one place to look when
// something has not updated.
function syncAll() {
  return { calendars: syncCalendars(), absences: syncAbsences() };
}
