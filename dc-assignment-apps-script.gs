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
// Copied on every assignment mail, so the leaders' mailbox carries a record of who was put on
// what without anyone having to remember to forward it. Set to '' to stop copying anyone.
const ASSIGN_CC    = 'sd-attendance@outsourceaccelerator.com';
const DAYS_AHEAD   = 21;
const DAYS_BEHIND  = 120;  // months of history, because the point is to look back
// How far back each run actually re-reads the calendar, which is a different question from how
// much history the tab keeps. Reading it is the expensive half: CalendarApp fetches the guest
// list and the creators of every event one round trip at a time, so the cost is the number of
// events in the window, not the number that changed. At 120 days behind that was roughly 2,800
// events and about 22,000 round trips, which ran for the full thirty minutes Apps Script allows
// and was killed twice in two days.
//
// A meeting three weeks past is settled: it will not be moved, renamed or cancelled again. So
// only the recent past is re-read, while DAYS_BEHIND above still governs what the tab retains.
// Nothing older is dropped or forgotten, it simply is not asked about again.
const REFRESH_BEHIND = 14;
const HEADERS = ['eventId', 'calendar', 'partner', 'lead', 'sdrEmail', 'start', 'end',
                 'durationMin', 'assignedTo', 'assignedBy', 'assignedAt', 'syncedAt',
                 'outcome', 'outcomeBy', 'outcomeAt'];

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

/**
 * A start or end read back out of the sheet, as the string the rest of this file expects.
 *
 * We write "2026-09-21T11:30" and Sheets recognises it as a datetime, so getValues hands back a
 * Date object rather than that text. String(thatDate) is "Mon Sep 21 2026 11:30:00 GMT+0800",
 * which sorts nowhere near an ISO string: every window comparison in the sync silently read a
 * September meeting as beyond the range, and the backfill built a key that could never match the
 * one from the mail. Normalised here so both sides speak the same language.
 */
function cellIso_(v) {
  if (v instanceof Date) { return Utilities.formatDate(v, 'Asia/Manila', "yyyy-MM-dd'T'HH:mm"); }
  return String(v || '').trim().replace(' ', 'T').slice(0, 16);
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
  // Rewritten every time rather than only when the tab is new: a sheet built before a column was
  // added would otherwise keep its old header and pair new values with the wrong names.
  sh.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
  return sh;
}

/**
 * Pull the calendars into the tab, keeping what is already there.
 *
 * This used to clear every row and rewrite from the calendar, which was fine while the tab only
 * ever looked forward. It cannot stay that way now the point is to look back: a meeting that has
 * happened tells you nothing if its row was deleted the next morning, and a cancelled meeting
 * simply stops coming back from the calendar, so a rewrite erases the very thing worth recording.
 *
 * So it merges. Rows already in the sheet are kept and updated; rows for events the calendar no
 * longer returns are kept too, and marked Cancelled when they sit inside the window that was
 * actually searched. A row missing because its date is outside that window is not cancelled, it
 * is merely out of sight, and is left alone.
 */
function syncCalendars() {
  const sh = sheet_();
  const now = new Date();
  const from = new Date(now.getTime() - REFRESH_BEHIND * 864e5);
  const to   = new Date(now.getTime() + DAYS_AHEAD  * 864e5);
  const iso  = function (d) { return Utilities.formatDate(d, 'Asia/Manila', "yyyy-MM-dd'T'HH:mm"); };
  const fromIso = iso(from), toIso = iso(to);
  // The retention edge, which is older than the refresh edge and is the only thing that decides
  // whether a row is dropped. Keeping these two apart is the whole point: a row can be kept
  // without being re-read.
  const keepFromIso = iso(new Date(now.getTime() - DAYS_BEHIND * 864e5));

  // everything already recorded, keyed by event
  const have = {};
  const order = [];
  const last = sh.getLastRow();
  if (last > 1) {
    sh.getRange(2, 1, last - 1, HEADERS.length).getValues().forEach(function (r) {
      const id = String(r[0] || '').trim();
      if (!id) { return; }
      have[id] = r.slice();
      order.push(id);
    });
  }

  const seen = {};
  const missing = [];
  CALENDARS.forEach(function (calId) {
    const cal = CalendarApp.getCalendarById(calId);
    if (!cal) { missing.push(calId); return; }
    cal.getEvents(from, to).forEach(function (ev) {
      const eid = ev.getId();
      seen[eid] = true;
      let guests = [];
      try { guests = ev.getGuestList().map(function (g) { return g.getEmail(); }); } catch (e) {}
      let creators = [];
      try { creators = ev.getCreators(); } catch (e) {}
      let partner = '';
      creators.concat(guests).forEach(function (e) { if (!partner) { partner = partnerFrom_(e); } });
      const rawTitle = String(ev.getTitle() || '').trim();
      const wasCancelled = CANCEL_RE.test(rawTitle);
      const title = rawTitle.replace(CANCEL_RE, '').trim();
      const lm = title.match(LEAD_RE);
      const prev = have[eid];
      // Never overwrites a person: a leader who recorded No-show on a meeting later renamed
      // keeps their answer, because they were there and the calendar was not.
      const hadOutcome = prev && String(prev[12] || '').trim();

      const row = [
        eid, calId, partner, (lm ? lm[1].trim() : title), sdrFrom_(guests),
        iso(ev.getStartTime()), iso(ev.getEndTime()),
        Math.round((ev.getEndTime() - ev.getStartTime()) / 60000),
        prev ? prev[8]  : '',      // assignedTo
        prev ? prev[9]  : '',      // assignedBy
        prev ? prev[10] : '',      // assignedAt
        iso(now),
        hadOutcome ? prev[12] : (wasCancelled ? 'Cancelled' : ''),
        hadOutcome ? prev[13] : (wasCancelled ? 'calendar' : ''),
        hadOutcome ? prev[14] : (wasCancelled ? iso(now) : '')
      ];
      if (!have[eid]) { order.push(eid); }
      have[eid] = row;
    });
  });

  // A row the calendar no longer returns, whose meeting sits inside the window we just searched,
  // is a meeting that was cancelled or deleted. Recorded rather than dropped, because "it was
  // cancelled" is an answer and a missing row is not. Never overwrites an outcome a person set.
  Object.keys(have).forEach(function (eid) {
    if (seen[eid]) { return; }
    const r = have[eid];
    const start = cellIso_(r[5]);
    if (!start || start < fromIso || start > toIso) { return; }
    if (!String(r[12] || '').trim()) {
      r[12] = 'Cancelled';
      r[13] = 'calendar';
      r[14] = iso(now);
    }
  });

  // Drop what is older than the retention window so the tab does not grow without limit. This
  // deliberately uses keepFromIso and not fromIso: measured against the refresh edge it would
  // delete every row older than a fortnight, which is most of the history the Previously
  // assigned tab exists to show.
  const out = [];
  order.forEach(function (eid) {
    const r = have[eid];
    if (!r) { return; }
    const start = cellIso_(r[5]);
    if (start && start < keepFromIso) { return; }
    out.push(r);
  });
  out.sort(function (a, b) { return String(a[5]).localeCompare(String(b[5])); });

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
  return { written: out.length, fromCalendar: Object.keys(seen).length, missing: missing };
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
    const kind = data && data.record;
    if (kind !== 'dcassign' && kind !== 'dcoutcome') {
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

      // How the meeting went. Recorded by a person, because it cannot be derived: the booking
      // sheet and these calendars share no key that holds, so guessing an outcome would be
      // inventing one. The calendar contributes the single fact it does know, that an event has
      // gone, and everything else is somebody saying what happened.
      if (kind === 'dcoutcome') {
        const val = String(data.outcome || '');
        sh.getRange(row, 13, 1, 3).setValues([[
          val, val ? String(data.outcomeBy || '') : '', val ? stamp : ''
        ]]);
        return json_({ ok: true, row: row, outcome: val || '(cleared)' });
      }
      // Clearing an assignment is a write like any other, so an empty name is allowed through
      // and blanks the row rather than being rejected as a mistake.
      sh.getRange(row, 9,  1, 3).setValues([[
        String(data.assignedTo || ''), String(data.assignedBy || ''),
        String(data.assignedTo || '') ? stamp : ''
      ]]);
      // The row is written before the mail goes out, and the mail cannot undo it. Somebody
      // being told twice is a nuisance; a call that nobody is recorded against is a problem.
      const calId = String(sh.getRange(row, 2).getValue() || '');
      const mailed = String(data.assignedTo || '') ? notifyAssignee_(data, calId, eid) : 'cleared, no mail';
      return json_({ ok: true, row: row, mailed: mailed });
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

// Titles read "Igor Matrosov and Partner Outposter", so everything before "and Partner" is
// the lead. A title written some other way keeps its whole text rather than being guessed at.
const LEAD_RE = /^(.+?)\s+and\s+Partner\b/i;
// A cancelled meeting is not deleted from these calendars, it is renamed: the title gains a
// "Canceled:" prefix and the event stays put. Thirteen per cent of the board carries one, and
// sometimes twice over when a meeting is cancelled, revived and cancelled again. So the prefix
// is both the outcome and a thing to strip, or the lead reads "Canceled: Ryan Blundell".
const CANCEL_RE = /^(\s*cancell?ed:\s*)+/i;

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
  const out = { calendars: syncCalendars(), absences: syncAbsences() };
  // The API sync only runs once a token exists; until then the pasted tab carries the outcomes,
  // and after that it stays harmless because HubSpot outranks it.
  const token = PropertiesService.getScriptProperties().getProperty('HUBSPOT_TOKEN');
  if (token) { out.outcomes = syncOutcomes(); }
  out.pasted = applyPastedOutcomes();
  // The confirmation emails, which are a record of booked DCs independent of the booking sheet.
  // Wrapped because a failure here must not take the calendar and absence syncs down with it:
  // those two are what the assignment board runs on, and this one is a second opinion.
  try { out.bookings = writeBookingsTab(); }
  catch (err) { out.bookings = { error: String(err) }; console.warn('booking email sync: ' + err); }
  return out;
}

/**
 * Tell the person they have been given a call.
 *
 * Sent from the account that runs this script, so it arrives from a real colleague rather than
 * a no-reply nobody reads. Only on an actual assignment: clearing one sends nothing, because a
 * mail saying a meeting is no longer yours is noise on a floor that already gets plenty.
 *
 * A failure here must not fail the assignment. The row is already written by the time this
 * runs, and refusing to record who is covering a call because a mailbox was full would be the
 * wrong way round. It is logged and swallowed.
 */
/**
 * The lead's name, read from the event at the moment the mail is sent.
 *
 * It is never written to the sheet. That tab is published as CSV, so anything in it is on the
 * open internet whatever the page chooses to draw; this is a private mail to one colleague, and
 * somebody being sent to cover a call needs to know who they are meeting. So it is fetched here
 * and thrown away, which gets the name to the one person entitled to it and nowhere else.
 */
function leadFrom_(calId, eventId) {
  try {
    const cal = CalendarApp.getCalendarById(calId);
    if (!cal) { return { title: '', lead: '' }; }
    const ev = cal.getEventById(eventId);
    if (!ev) { return { title: '', lead: '' }; }
    const title = String(ev.getTitle() || '').trim();
    // Titles read "Igor Matrosov and Partner Outposter". Everything before "and Partner" is the
    // lead; where a title is written some other way the whole title is used rather than a guess.
    const m = title.match(LEAD_RE);
    return { title: title, lead: m ? m[1].trim() : '' };
  } catch (err) {
    return { title: '', lead: '' };
  }
}

function notifyAssignee_(data, calId, eventId) {
  const to = String(data.assignedEmail || '').trim();
  if (!to) { return 'no address'; }
  if (!/^[^@\s]+@outsourceaccelerator\.com$/i.test(to)) { return 'refused: not an OA address'; }

  const when = absPrettyWhen_(String(data.startsAt || ''));
  const partner = String(data.partner || '').trim();
  const mins = Number(data.mins || 0);
  const booked = String(data.bookedBy || '').trim();
  const by = String(data.assignedBy || '').trim();

  const ev = leadFrom_(calId, eventId);
  const subject = 'You are covering a discovery call' + (when ? ' — ' + when : '')
    + (ev.lead ? ' with ' + ev.lead : '');
  const lines = [
    'Hi ' + String(data.assignedTo || '').split(' ')[0] + ',',
    '',
    'You have been assigned to cover a discovery call.',
    '',
    (ev.lead ? 'Lead:     ' + ev.lead : (ev.title ? 'Meeting:  ' + ev.title : '')),
    (when ? 'When:     ' + when + (mins ? ' (' + mins + ' min)' : '') : ''),
    (partner ? 'Partner:  ' + partner : ''),
    (booked ? 'Booked by: ' + booked : ''),
    (by ? 'Assigned by: ' + by : ''),
    '',
    'The meeting is on the concierge calendar. If you cannot take it, tell your team leader as',
    'soon as you can so it can go to somebody else.',
    '',
    'Sent automatically by the DC assignment board.'
  ].filter(function (l) { return l !== ''; });

  try {
    // Copied: the leaders' mailbox, and the rep who booked it. The booker is the one person
    // who knows what the call is about, and a handover where they are not told is how a lead
    // arrives at a meeting nobody has prepared for. Never copied to themselves.
    const cc = [];
    if (ASSIGN_CC) { cc.push(ASSIGN_CC); }
    const booker = String(data.bookedByEmail || '').trim().toLowerCase();
    if (booker && booker !== to.toLowerCase()
        && /^[^@\s]+@outsourceaccelerator\.com$/i.test(booker)) { cc.push(booker); }
    const mail = { to: to, subject: subject, body: lines.join(String.fromCharCode(10)) };
    if (cc.length) { mail.cc = cc.join(','); }
    MailApp.sendEmail(mail);
    return cc.length ? 'sent, cc ' + cc.join(' ') : 'sent';
  } catch (err) {
    // Quota exhausted, bad address, anything: say so in the log and let the assignment stand.
    console.warn('could not email ' + to + ': ' + err);
    return 'not sent: ' + err;
  }
}

// "Tuesday 9 Sep, 11:00 AM" from 2026-09-09T11:00. Written out rather than left as an ISO
// string, because a rep reading this on a phone should not have to decode a timestamp.
function absPrettyWhen_(iso) {
  const m = String(iso).match(/^(\d{4})-(\d{2})-(\d{2})[T ](\d{1,2}):(\d{2})/);
  if (!m) { return ''; }
  const d = new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]), Number(m[4]), Number(m[5]));
  return Utilities.formatDate(d, 'Asia/Manila', 'EEEE d MMM, h:mm a');
}

/* ════════════════════════════════════════════════════════════════════════════
 * BACKFILL — recover assignments the old sync destroyed
 *
 * Until the sync started merging, it cleared the tab and rewrote it from the calendar every half
 * hour, keeping one day behind. Every assignment older than yesterday went with it, and by the
 * time anyone looked the record showed six covers where there had been two hundred.
 *
 * The notification mail is the surviving copy. Each one names the assignee in its To line and
 * carries the lead, the time, the partner and who booked it in the body, which between them
 * identify the meeting precisely enough to put the assignment back on the right row.
 *
 * Safe to run more than once: it only ever fills a blank, so anything assigned since is left
 * exactly as it stands.
 * ════════════════════════════════════════════════════════════════════════════ */

const BACKFILL_QUERY = 'in:sent subject:"You are covering a discovery call"';
const BACKFILL_MAX   = 400;
const BF_MONTHS = ['jan','feb','mar','apr','may','jun','jul','aug','sep','oct','nov','dec'];

function bfField_(body, label) {
  const m = String(body || '').match(new RegExp('^' + label + ':\\s*(.+)$', 'im'));
  return m ? m[1].trim() : '';
}

// "Monday 21 Sep, 11:30 AM (30 min)" against the date the mail was sent, which supplies the year
// the line leaves out. A meeting that lands well before the mail belongs to the following year,
// which is what makes a December assignment for January work.
function bfWhen_(when, sentAt) {
  const m = String(when || '').match(/(\d{1,2})\s+([A-Za-z]{3,})[,\s]+(\d{1,2}):(\d{2})\s*(AM|PM)/i);
  if (!m) { return ''; }
  const mi = BF_MONTHS.indexOf(m[2].slice(0, 3).toLowerCase());
  if (mi < 0) { return ''; }
  let hh = Number(m[3]) % 12;
  if (/pm/i.test(m[5])) { hh += 12; }
  let y = Number(Utilities.formatDate(sentAt, 'Asia/Manila', 'yyyy'));
  let d = new Date(y, mi, Number(m[1]), hh, Number(m[4]));
  if (d.getTime() < sentAt.getTime() - 60 * 864e5) { d = new Date(y + 1, mi, Number(m[1]), hh, Number(m[4])); }
  return Utilities.formatDate(d, 'Asia/Manila', "yyyy-MM-dd'T'HH:mm");
}

function bfKey_(startIso, partner) {
  return String(startIso || '') + '|' + String(partner || '').toLowerCase().replace(/[^a-z0-9]/g, '');
}

function backfillAssignments() {
  const sh = sheet_();
  const last = sh.getLastRow();
  if (last < 2) { return { error: 'nothing synced yet, run syncAll first' }; }

  // What the mail says, newest last so a reassignment overwrites the earlier one.
  const found = {};
  let scanned = 0, parsed = 0;
  GmailApp.search(BACKFILL_QUERY, 0, BACKFILL_MAX).forEach(function (thread) {
    thread.getMessages().forEach(function (msg) {
      scanned++;
      let body = '';
      try { body = msg.getPlainBody(); } catch (e) { return; }
      if (body.indexOf('assigned to cover a discovery call') < 0) { return; }
      const to = String(msg.getTo() || '').match(/[\w.\-+]+@[\w.\-]+/);
      if (!to) { return; }
      const start = bfWhen_(bfField_(body, 'When'), msg.getDate());
      const partner = bfField_(body, 'Partner');
      if (!start || !partner) { return; }
      parsed++;
      const k = bfKey_(start, partner);
      const at = Utilities.formatDate(msg.getDate(), 'Asia/Manila', "yyyy-MM-dd'T'HH:mm");
      if (!found[k] || found[k].at <= at) {
        found[k] = { email: to[0].toLowerCase(), at: at, by: bfField_(body, 'Assigned by') };
      }
    });
  });

  const rows = sh.getRange(2, 1, last - 1, HEADERS.length).getValues();
  let filled = 0, already = 0, unmatched = 0;
  const seen = {};
  rows.forEach(function (r) {
    const k = bfKey_(cellIso_(r[5]), String(r[2] || ''));
    const hit = found[k];
    if (!hit) { return; }
    seen[k] = true;
    if (String(r[8] || '').trim()) { already++; return; }   // never overwrite a live assignment
    r[8] = hit.email;
    r[9] = hit.by || 'restored from sent mail';
    r[10] = hit.at;
    filled++;
  });
  Object.keys(found).forEach(function (k) { if (!seen[k]) { unmatched++; } });

  if (filled) { sh.getRange(2, 1, rows.length, HEADERS.length).setValues(rows); }
  console.log('scanned ' + scanned + ' sent mails, ' + parsed + ' were assignments, '
    + filled + ' rows filled, ' + already + ' already assigned, '
    + unmatched + ' mails matched no row in the sheet');
  return { scanned: scanned, parsed: parsed, filled: filled, already: already, unmatched: unmatched };
}

/* ════════════════════════════════════════════════════════════════════════════
 * OUTCOMES FROM HUBSPOT
 *
 * The outcome was never missing, it was in the CRM. Every discovery call is a HubSpot meeting
 * carrying hs_meeting_outcome, and the vocabulary is the floor's own: BOTH ATTENDED, BPO
 * ATTENDED, CANCELED BY LEAD, RESCHEDULED. The titles are identical to the calendar events
 * — "Nicole Hart and Partner Six Eleven" — so lead, partner and start time together identify a
 * meeting on both sides without anything having to be typed twice.
 *
 * This is why the booking sheet was the wrong place to look. It records what an SDR logged; this
 * records what happened, which is a different fact and the one being asked for.
 *
 * SETUP: HubSpot → Settings → Integrations → Private Apps → Create, with scope
 *        crm.objects.meetings.read. Copy the token into Script properties as HUBSPOT_TOKEN.
 *        Kept there rather than here, because this file is in a public repository.
 *
 * A leader's own answer still wins. Somebody who watched a call happen and recorded it keeps
 * their answer even if the CRM disagrees, because they were there.
 * ════════════════════════════════════════════════════════════════════════════ */

const HS_SEARCH_URL = 'https://api.hubapi.com/crm/v3/objects/meetings/search';
const HS_PAGE       = 100;
const HS_MAX_PAGES  = 40;

// HubSpot's vocabulary, reduced to the four things the board shows. Anything unrecognised is
// passed through as it stands rather than dropped, so a new outcome added in the CRM shows up
// here as itself instead of silently becoming nothing.
function hsOutcome_(v) {
  const s = String(v || '').trim().toUpperCase();
  if (!s || s === 'SCHEDULED') { return ''; }              // has not happened yet

  // Went ahead. Unqualified is kept apart: the call happened and the hour was spent, but it
  // should not count as a win, and averaging it with the rest would quietly say it did.
  if (s.indexOf('UNQUALIFIED') >= 0)    { return 'Showed, unqualified'; }
  if (s.indexOf('BOTH ATTENDED') === 0) { return 'Showed'; }
  if (s === 'COMPLETED' || s === 'ATTENDED') { return 'Showed'; }

  // Did not go ahead, and which side failed is the point. Three different conversations.
  if (s.indexOf('BPO ATTENDED') === 0)  { return 'Lead no-show'; }      // partner came, lead did not
  if (s.indexOf('LEAD ATTENDED') === 0) { return 'Partner no-show'; }   // lead came, partner did not
  if (s.indexOf('NO SHOW') >= 0 || s.indexOf('NO_SHOW') >= 0) { return 'No-show'; }

  // Called off, and by whom. A lead cancelling is a lead-quality signal; a partner cancelling
  // is ours to answer for, and one number covering both hides the difference.
  if (s.indexOf('CANCEL') >= 0 && s.indexOf('LEAD') >= 0)    { return 'Cancelled by lead'; }
  if (s.indexOf('CANCEL') >= 0 && s.indexOf('PARTNER') >= 0) { return 'Cancelled by partner'; }
  if (s.indexOf('CANCEL') >= 0)         { return 'Cancelled'; }

  if (s.indexOf('RESCHEDUL') >= 0)      { return 'Rescheduled'; }
  if (s === 'INVALID')                  { return 'Invalid'; }
  return String(v || '').trim();          // something new in the CRM, shown as itself
}

function hsKey_(lead, partner, startIso) {
  const n = function (x) { return String(x || '').toLowerCase().replace(/[^a-z0-9]/g, ''); };
  return n(lead) + '|' + n(partner) + '|' + String(startIso || '');
}

// "Nicole Hart and Partner Six Eleven" back into its two halves, the same split the sync makes
// when it reads the calendar.
function hsSplitTitle_(title) {
  const t = String(title || '').replace(CANCEL_RE, '').trim();
  const m = t.match(/^(.+?)\s+and\s+Partner\s+(.+)$/i);
  return m ? { lead: m[1].trim(), partner: m[2].trim() } : { lead: t, partner: '' };
}

function syncOutcomes() {
  const token = PropertiesService.getScriptProperties().getProperty('HUBSPOT_TOKEN');
  if (!token) { return { error: 'HUBSPOT_TOKEN script property is not set' }; }

  const sh = sheet_();
  const last = sh.getLastRow();
  if (last < 2) { return { error: 'nothing synced yet, run syncAll first' }; }

  const now = new Date();
  const from = new Date(now.getTime() - DAYS_BEHIND * 864e5);
  const to   = new Date(now.getTime() + DAYS_AHEAD  * 864e5);

  const found = {};
  let after = null, pages = 0, fetched = 0;
  do {
    const payload = {
      filterGroups: [{ filters: [{ propertyName: 'hs_meeting_start_time', operator: 'BETWEEN',
                                   value: String(from.getTime()), highValue: String(to.getTime()) }] }],
      properties: ['hs_meeting_title', 'hs_meeting_start_time', 'hs_meeting_outcome'],
      limit: HS_PAGE,
      sorts: [{ propertyName: 'hs_meeting_start_time', direction: 'ASCENDING' }]
    };
    if (after) { payload.after = after; }
    const res = UrlFetchApp.fetch(HS_SEARCH_URL, {
      method: 'post', contentType: 'application/json',
      headers: { Authorization: 'Bearer ' + token },
      payload: JSON.stringify(payload), muteHttpExceptions: true
    });
    if (res.getResponseCode() !== 200) {
      return { error: 'HubSpot ' + res.getResponseCode() + ': ' + res.getContentText().slice(0, 200) };
    }
    const body = JSON.parse(res.getContentText());
    (body.results || []).forEach(function (m) {
      const p = m.properties || {};
      if (!p.hs_meeting_start_time) { return; }
      const startIso = Utilities.formatDate(new Date(p.hs_meeting_start_time), 'Asia/Manila',
                                            "yyyy-MM-dd'T'HH:mm");
      const parts = hsSplitTitle_(p.hs_meeting_title);
      found[hsKey_(parts.lead, parts.partner, startIso)] = hsOutcome_(p.hs_meeting_outcome);
      fetched++;
    });
    after = body.paging && body.paging.next ? body.paging.next.after : null;
    pages++;
  } while (after && pages < HS_MAX_PAGES);

  const rows = sh.getRange(2, 1, last - 1, HEADERS.length).getValues();
  let set = 0, kept = 0, noMatch = 0;
  rows.forEach(function (r) {
    const k = hsKey_(r[3], r[2], cellIso_(r[5]));
    const hs = found[k];
    if (hs === undefined) { noMatch++; return; }
    const by = String(r[13] || '').trim().toLowerCase();
    // A person's answer stands. The calendar's does not: HubSpot is the better witness to a
    // cancellation, and a row marked Cancelled by the calendar can be corrected by it.
    if (String(r[12] || '').trim() && by !== '' && by !== 'calendar' && by !== 'hubspot') { kept++; return; }
    if (!hs) { return; }
    if (r[12] === hs && by === 'hubspot') { return; }
    r[12] = hs; r[13] = 'hubspot';
    r[14] = Utilities.formatDate(now, 'Asia/Manila', "yyyy-MM-dd'T'HH:mm");
    set++;
  });
  if (set) { sh.getRange(2, 1, rows.length, HEADERS.length).setValues(rows); }

  console.log('HubSpot meetings read ' + fetched + ' over ' + pages + ' page(s); '
    + set + ' outcomes written, ' + kept + ' left as a person recorded them, '
    + noMatch + ' sheet rows had no HubSpot meeting');
  return { fetched: fetched, set: set, kept: kept, noMatch: noMatch };
}

/* ════════════════════════════════════════════════════════════════════════════
 * PASTED OUTCOMES — the bridge until a HubSpot token exists
 *
 * Private apps are admin-only in this portal, so syncOutcomes has nothing to authenticate with
 * yet. Meanwhile the outcomes do exist in HubSpot and can be lifted out by hand. Paste them into
 * an "Outcomes" tab, four columns — lead, partner, start, outcome — and this matches them to the
 * board the same way the API sync would, on lead, partner and start time.
 *
 * It is a bridge and behaves like one. It never overrides a person, and never overrides HubSpot
 * once the token arrives, so leaving the tab in place after the real sync is running does no
 * harm: it simply stops being the freshest thing in the room.
 * ════════════════════════════════════════════════════════════════════════════ */

const PASTE_TAB = 'Outcomes';

function applyPastedOutcomes() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const src = ss.getSheetByName(PASTE_TAB);
  if (!src || src.getLastRow() < 2) { return { skipped: 'no Outcomes tab, or it is empty' }; }

  const vals = src.getRange(1, 1, src.getLastRow(), Math.max(4, src.getLastColumn())).getValues();
  const head = vals[0].map(function (h) { return String(h || '').trim().toLowerCase(); });
  const col = function (n) { return head.indexOf(n); };
  const iLead = col('lead'), iPartner = col('partner'), iStart = col('start'), iOut = col('outcome');
  if (iLead < 0 || iPartner < 0 || iStart < 0 || iOut < 0) {
    return { error: 'the Outcomes tab needs a header row: lead, partner, start, outcome' };
  }

  const want = {};
  let read = 0;
  vals.slice(1).forEach(function (r) {
    const lead = String(r[iLead] || '').trim();
    const start = cellIso_(r[iStart]);
    const outcome = String(r[iOut] || '').trim();
    if (!lead || !start || !outcome) { return; }
    want[hsKey_(lead, r[iPartner], start)] = outcome;
    read++;
  });

  const sh = sheet_();
  const last = sh.getLastRow();
  if (last < 2) { return { error: 'nothing synced yet, run syncAll first' }; }
  const rows = sh.getRange(2, 1, last - 1, HEADERS.length).getValues();

  let set = 0, kept = 0, noMatch = 0;
  const used = {};
  rows.forEach(function (r) {
    const k = hsKey_(r[3], r[2], cellIso_(r[5]));
    const v = want[k];
    if (v === undefined) { return; }
    used[k] = true;
    const by = String(r[13] || '').trim().toLowerCase();
    // A person's answer and HubSpot's both outrank a paste. The calendar's guess does not.
    if (String(r[12] || '').trim() && by !== '' && by !== 'calendar') { kept++; return; }
    if (r[12] === v && by === 'pasted') { return; }
    r[12] = v; r[13] = 'pasted';
    r[14] = Utilities.formatDate(new Date(), 'Asia/Manila', "yyyy-MM-dd'T'HH:mm");
    set++;
  });
  Object.keys(want).forEach(function (k) { if (!used[k]) { noMatch++; } });

  if (set) { sh.getRange(2, 1, rows.length, HEADERS.length).setValues(rows); }
  console.log('Outcomes tab: ' + read + ' rows read, ' + set + ' applied, '
    + kept + ' left as a person or HubSpot recorded them, ' + noMatch + ' matched no meeting');
  return { read: read, set: set, kept: kept, noMatch: noMatch };
}

/* ──────────────────────────────────────────────────────────────────────────
 * BOOKING CONFIRMATION EMAILS — READ-ONLY SCAN
 *
 * Run scanBookingEmails() and read the log. It writes nothing anywhere: the
 * point is to find out whether these emails can be counted as booked DCs
 * before any code depends on them being able to.
 *
 * Three things it is actually testing, each of which would quietly produce
 * wrong numbers if assumed:
 *
 *  1. Reminders. A confirmation thread carries a courtesy reminder with the
 *     SAME subject, sometimes days later. Counting messages would count the
 *     booking twice. This keys on partner + lead + meeting date and keeps the
 *     EARLIEST message, so a reminder collapses into the booking it reminds
 *     about whether it sits in the same thread or a new one.
 *
 *  2. Which date. The date in the subject is when the MEETING happens; the
 *     message date is when it was booked. The DC Dashboard counts by booking
 *     date, so that is what the per-day figure uses, and the scan prints both
 *     so the gap between them is visible rather than assumed away.
 *
 *  3. Coverage. Override leads get the confirmation only once the lead
 *     accepts, so some DCs have no email or a late one. The scan cannot see
 *     what was never sent, but comparing its daily counts against the booking
 *     sheet shows the size of the hole.
 * ────────────────────────────────────────────────────────────────────────── */

var SCAN_QUERY_SUBJECT = 'subject:"Meeting Confirmation:"';
var SCAN_DAYS          = 30;
var SCAN_MAX_THREADS   = 400;

// "Meeting Confirmation: VA Platinum <> Jamie Harawira - October 7, 2026"
// The separator is the literal "<>" the floor types between partner and lead.
var SCAN_SUBJ_RE = /^\s*(?:re:|fwd?:)*\s*Meeting\s+Confirmation:\s*(.+?)\s*<>\s*(.+?)\s*[-–]\s*(.+?)\s*$/i;

function scanParseSubject_(subject) {
  var m = String(subject || '').match(SCAN_SUBJ_RE);
  if (!m) { return null; }
  return { partner: m[1].trim(), lead: m[2].trim(), whenRaw: m[3].trim() };
}

function scanIsoDay_(d) {
  return Utilities.formatDate(d, 'Asia/Manila', 'yyyy-MM-dd');
}

// "October 7, 2026" -> 2026-10-07. Read off the string directly rather than
// through Date.parse: that returns an instant, and an instant formatted in
// another timezone slides to the day before. Here the day is just three
// numbers and never becomes a moment in time, so it cannot drift. It matters
// more than it looks: the meeting date is part of the dedupe key, so a
// one-day slide would stop a reminder matching its own confirmation and
// double-count the booking.
var SCAN_MONTHS = { jan:1, feb:2, mar:3, apr:4, may:5, jun:6, jul:7, aug:8, sep:9, oct:10, nov:11, dec:12 };
function scanMeetingDate_(raw) {
  var s = String(raw || '').replace(/\(.*?\)/g, ' ').replace(/(\d+)(st|nd|rd|th)/gi, '$1').trim();
  var mon = 0, day = 0, year = 0, m;

  // "October 7, 2026" / "Oct 7 2026"
  m = s.match(/([A-Za-z]{3,})\.?\s+(\d{1,2})\s*,?\s*(\d{4})/);
  if (m) { mon = SCAN_MONTHS[m[1].slice(0,3).toLowerCase()]; day = +m[2]; year = +m[3]; }

  // "7 October 2026"
  if (!mon) {
    m = s.match(/(\d{1,2})\s+([A-Za-z]{3,})\.?\s*,?\s*(\d{4})/);
    if (m) { day = +m[1]; mon = SCAN_MONTHS[m[2].slice(0,3).toLowerCase()]; year = +m[3]; }
  }

  // "10/7/2026", read month-first the way the booking sheet writes it
  if (!mon) {
    m = s.match(/(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})/);
    if (m) { mon = +m[1]; day = +m[2]; year = +m[3]; }
  }

  if (!mon || !day || !year || mon > 12 || day > 31) { return ''; }
  return year + '-' + (mon < 10 ? '0' : '') + mon + '-' + (day < 10 ? '0' : '') + day;
}

function scanKey_(p) {
  var n = function (x) { return String(x || '').toLowerCase().replace(/[^a-z0-9]/g, ''); };
  return n(p.partner) + '|' + n(p.lead) + '|' + (p.meetingDate || n(p.whenRaw));
}

function scanBookingEmails() {
  var query = SCAN_QUERY_SUBJECT + ' newer_than:' + SCAN_DAYS + 'd';
  var threads = GmailApp.search(query, 0, SCAN_MAX_THREADS);

  // Only the first message of each thread is read. That message IS the
  // confirmation; every courtesy reminder is a later one under the same
  // subject, so skipping them is both the cheap path and the correct one.
  // It also keeps the Gmail calls down to roughly one per thread instead of
  // one per message, which matters on a script that has hit the execution
  // ceiling before.
  var seen = {};            // dedupe key -> the earliest confirmation seen
  var msgTotal = 0, unparsed = [], parsedMsgs = 0;

  threads.forEach(function (th) {
    msgTotal++;
    var subject = th.getFirstMessageSubject();
    var p = scanParseSubject_(subject);
    if (!p) {
      if (unparsed.length < 10) { unparsed.push(subject); }
      return;
    }
    parsedMsgs++;
    p.meetingDate = scanMeetingDate_(p.whenRaw);

    var first = th.getMessages()[0];
    if (!first) { return; }
    var from = String(first.getFrom() || '');
    var hit  = from.match(/[\w.\-+]+@[\w.\-]+/);
    p.sender  = hit ? hit[0].toLowerCase() : from;
    p.sent    = first.getDate();
    p.sentDay = scanIsoDay_(p.sent);
    p.inThread = th.getMessageCount();

    // A reminder sent as its own thread rather than a reply still lands on the
    // same key, so the earliest one wins and the booking is counted once.
    var k = scanKey_(p);
    if (!seen[k] || p.sent < seen[k].sent) { seen[k] = p; }
  });

  var rows = Object.keys(seen).map(function (k) { return seen[k]; });

  // ── what it found ──────────────────────────────────────────────────────
  Logger.log('QUERY: ' + query);
  Logger.log('threads matched        : ' + threads.length
             + (threads.length >= SCAN_MAX_THREADS ? '  *** CAPPED, raise SCAN_MAX_THREADS ***' : ''));
  Logger.log('subjects that parsed   : ' + parsedMsgs);
  Logger.log('distinct bookings      : ' + rows.length
             + '   (' + (parsedMsgs - rows.length) + ' duplicate thread(s) collapsed)');
  var reminders = 0;
  rows.forEach(function (r) { reminders += Math.max(0, (r.inThread || 1) - 1); });
  Logger.log('reminders skipped      : ' + reminders
             + '   (later messages under the same subject, never counted)');
  Logger.log('');

  if (unparsed.length) {
    Logger.log('SUBJECTS THAT DID NOT PARSE (first ' + unparsed.length + '):');
    unparsed.forEach(function (s) { Logger.log('   ' + s); });
    Logger.log('');
  }

  var noDate = rows.filter(function (r) { return !r.meetingDate; });
  if (noDate.length) {
    Logger.log(noDate.length + ' booking(s) had a meeting date that could not be read, e.g. "'
               + noDate[0].whenRaw + '"');
    Logger.log('');
  }

  // ── who sent them ──────────────────────────────────────────────────────
  var bySender = {};
  rows.forEach(function (r) { bySender[r.sender] = (bySender[r.sender] || 0) + 1; });
  var senders = Object.keys(bySender).sort(function (a, b) { return bySender[b] - bySender[a]; });
  Logger.log('BOOKINGS PER SENDER (' + senders.length + ' senders)');
  senders.forEach(function (s) { Logger.log('   ' + pad_(s, 44) + bySender[s]); });
  Logger.log('');

  // ── per day, by the date it was SENT, which is the booking date ────────
  var byDay = {};
  rows.forEach(function (r) { byDay[r.sentDay] = (byDay[r.sentDay] || 0) + 1; });
  var days = Object.keys(byDay).sort();
  Logger.log('BOOKINGS PER DAY, by the day the confirmation was sent');
  days.forEach(function (d) { Logger.log('   ' + d + '   ' + byDay[d]); });
  Logger.log('');

  // ── how far ahead meetings are booked, which is why the two dates differ
  var lead = [], sameDay = 0;
  rows.forEach(function (r) {
    if (!r.meetingDate) { return; }
    var a = new Date(r.sentDay + 'T00:00:00Z').getTime();
    var b = new Date(r.meetingDate + 'T00:00:00Z').getTime();
    var days = Math.round((b - a) / 864e5);
    lead.push(days);
    if (days === 0) { sameDay++; }
  });
  if (lead.length) {
    lead.sort(function (x, y) { return x - y; });
    Logger.log('GAP BETWEEN BOOKING AND MEETING (days)');
    Logger.log('   soonest ' + lead[0] + ', median ' + lead[Math.floor(lead.length / 2)]
               + ', furthest ' + lead[lead.length - 1] + ', same-day ' + sameDay);
    Logger.log('   This is why "per day" has to say which date it means.');
    Logger.log('');
  }

  // ── a few whole rows, so the parse can be eyeballed ────────────────────
  Logger.log('FIRST 8 BOOKINGS AS PARSED');
  rows.slice(0, 8).forEach(function (r) {
    Logger.log('   sent ' + r.sentDay + ' | meet ' + (r.meetingDate || '??')
               + ' | ' + pad_(r.partner, 24) + ' | ' + pad_(r.lead, 22) + ' | ' + r.sender);
  });

  return { threads: threads.length, messages: msgTotal, bookings: rows.length, senders: senders.length };
}

function pad_(s, n) {
  s = String(s === null || s === undefined ? '' : s);
  while (s.length < n) { s += ' '; }
  return s.length > n ? s.slice(0, n) : s;
}

/* ──────────────────────────────────────────────────────────────────────────
 * BOOKING CONFIRMATION EMAILS — WRITE TO A TAB
 *
 * The scan proved the shape: one confirmation per partner per lead, sender is
 * the SDR, subject carries partner, lead and meeting date. So one email is one
 * booked DC, which is the same unit the booking sheet counts in its SP 1 and
 * SP 2 columns.
 *
 * This writes those bookings to a "Bookings" tab for the dashboard to read.
 * It does NOT replace the booking sheet. It sits beside it, because the two
 * will disagree and the disagreement is the useful part:
 *
 *   - An email with no sheet row is a DC somebody booked and never logged.
 *   - A sheet row with no email is an override lead whose confirmation has not
 *     gone out yet, or a booking that was logged but never confirmed.
 *
 * Neither is automatically the truth, so neither is silently preferred.
 *
 * Merges rather than rewrites, like the calendar sync: the Gmail query only
 * reaches back SCAN_DAYS, and a tab that forgot everything older every run
 * would be useless for exactly the ranges the dashboard is built to show.
 * ────────────────────────────────────────────────────────────────────────── */

var BOOK_TAB     = 'Bookings';
var BOOK_HEADERS = ['key', 'bookedOn', 'meetingDate', 'partner', 'lead', 'sdrEmail', 'syncedAt'];
var BOOK_KEEP_DAYS = 400;   // how much history the tab retains, not how far the query reads

function writeBookingsTab() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) { return { error: 'this script is not attached to a spreadsheet' }; }
  var sh = ss.getSheetByName(BOOK_TAB);
  if (!sh) { sh = ss.insertSheet(BOOK_TAB); }
  if (sh.getLastRow() === 0) { sh.appendRow(BOOK_HEADERS); }
  sh.getRange(1, 1, 1, BOOK_HEADERS.length).setValues([BOOK_HEADERS]);

  // What the tab already holds, so history older than the query window survives.
  var have = {}, order = [];
  var last = sh.getLastRow();
  if (last > 1) {
    sh.getRange(2, 1, last - 1, BOOK_HEADERS.length).getValues().forEach(function (r) {
      var k = String(r[0] || '').trim();
      if (!k) { return; }
      have[k] = r.slice();
      order.push(k);
    });
  }

  var now = new Date();
  var stamp = Utilities.formatDate(now, 'Asia/Manila', "yyyy-MM-dd'T'HH:mm");
  var query = SCAN_QUERY_SUBJECT + ' newer_than:' + SCAN_DAYS + 'd';
  var threads = GmailApp.search(query, 0, SCAN_MAX_THREADS);

  var fresh = {}, parsed = 0, skipped = 0;
  threads.forEach(function (th) {
    var p = scanParseSubject_(th.getFirstMessageSubject());
    if (!p) { skipped++; return; }
    p.meetingDate = scanMeetingDate_(p.whenRaw);

    var first = th.getMessages()[0];
    if (!first) { return; }
    var hit = String(first.getFrom() || '').match(/[\w.\-+]+@[\w.\-]+/);
    p.sender  = hit ? hit[0].toLowerCase() : '';
    p.sent    = first.getDate();
    p.sentDay = scanIsoDay_(p.sent);
    parsed++;

    var k = scanKey_(p);
    // The earliest confirmation is the booking; a later one under the same key
    // is a reminder or a resend and must not become a second DC.
    if (!fresh[k] || p.sent < fresh[k].sent) { fresh[k] = p; }
  });

  var added = 0, updated = 0;
  Object.keys(fresh).forEach(function (k) {
    var p = fresh[k];
    var row = [k, p.sentDay, p.meetingDate, p.partner, p.lead, p.sender, stamp];
    if (!have[k]) { order.push(k); added++; }
    else if (String(have[k][1]) !== p.sentDay || String(have[k][5]) !== p.sender) { updated++; }
    have[k] = row;
  });

  // Trim only by age, never by whether the query still returns it: a booking
  // from two months ago is absent from the search window and still true.
  var cutoff = Utilities.formatDate(new Date(now.getTime() - BOOK_KEEP_DAYS * 864e5),
                                    'Asia/Manila', 'yyyy-MM-dd');
  var out = [], dropped = 0;
  order.forEach(function (k) {
    var r = have[k];
    if (!r) { return; }
    var d = String(r[1] || '');
    if (d && d < cutoff) { dropped++; return; }
    out.push(r);
  });
  out.sort(function (a, b) { return String(a[1]).localeCompare(String(b[1])); });

  if (sh.getLastRow() > 1) {
    sh.getRange(2, 1, sh.getLastRow() - 1, BOOK_HEADERS.length).clearContent();
  }
  if (out.length) {
    sh.getRange(2, 1, out.length, BOOK_HEADERS.length).setValues(out);
  }

  console.log('Bookings: ' + threads.length + ' threads, ' + parsed + ' parsed, '
    + skipped + ' subjects ignored, ' + added + ' new, ' + updated + ' updated, '
    + out.length + ' rows on the tab' + (dropped ? ', ' + dropped + ' aged out' : ''));
  return { threads: threads.length, parsed: parsed, added: added,
           updated: updated, rows: out.length };
}

/* ──────────────────────────────────────────────────────────────────────────
 * WHAT THE MAILBOX KNOWS ABOUT WHETHER A MEETING HAPPENED — READ-ONLY SCAN
 *
 * The confirmation emails settled whether a DC was booked. The open question
 * is whether anything in the same mailbox settles whether it went ahead, which
 * would be worth a great deal: HubSpot carries an outcome on 3% of meetings,
 * so almost any email signal would be a better witness than the CRM.
 *
 * This guesses at nothing. It reads what is there and reports the vocabulary:
 *
 *   PART 1  every subject prefix in recent mail, counted. If the floor sends
 *           "Meeting Rescheduled:" or "No Show:" this is where it appears,
 *           named by the floor rather than by me.
 *
 *   PART 2  the messages that arrive AFTER a confirmation, on confirmation
 *           threads whose meeting date has already passed. That is where a
 *           "the lead did not join" would land, and it is the single most
 *           likely place for a real outcome signal to be hiding.
 *
 * Writes nothing. Run it, read the log, and then we know what can be built
 * instead of hoping.
 * ────────────────────────────────────────────────────────────────────────── */

var OSCAN_DAYS        = 45;
var OSCAN_MAX_THREADS = 150;
var OSCAN_SAMPLE      = 25;   // threads to read bodies from, kept small for runtime

// The text before the first colon, which is how this floor labels a mail type.
function oscanPrefix_(subject) {
  var s = String(subject || '').replace(/^\s*(?:re:|fwd?:)\s*/i, '').trim();
  var i = s.indexOf(':');
  if (i > 0 && i < 40) { return s.slice(0, i).trim(); }
  return '(no prefix) ' + s.slice(0, 34);
}

function scanOutcomeSignals() {
  Logger.log('=== PART 1: what kinds of mail are in here ===');
  var prefixes = {}, examples = {};
  var broad = GmailApp.search('newer_than:' + OSCAN_DAYS + 'd -category:promotions', 0, OSCAN_MAX_THREADS);
  broad.forEach(function (th) {
    var subj = th.getFirstMessageSubject();
    var p = oscanPrefix_(subj);
    prefixes[p] = (prefixes[p] || 0) + 1;
    if (!examples[p]) { examples[p] = subj; }
  });
  var keys = Object.keys(prefixes).sort(function (a, b) { return prefixes[b] - prefixes[a]; });
  Logger.log('sampled ' + broad.length + ' threads from the last ' + OSCAN_DAYS + ' days');
  keys.slice(0, 25).forEach(function (k) {
    Logger.log('   ' + pad_(String(prefixes[k]), 5) + pad_(k, 34) + ' e.g. ' + String(examples[k]).slice(0, 70));
  });
  Logger.log('');

  // Anything whose wording suggests a meeting did or did not happen.
  Logger.log('=== subject words that would signal an outcome ===');
  ['no show', 'no-show', 'did not', 'didn\'t', 'reschedul', 'cancel', 'recap',
   'follow up', 'follow-up', 'missed', 'summary'].forEach(function (w) {
    var n = GmailApp.search('subject:"' + w + '" newer_than:' + OSCAN_DAYS + 'd', 0, 50);
    if (n.length) {
      Logger.log('   ' + pad_('"' + w + '"', 16) + pad_(String(n.length) + (n.length >= 50 ? '+' : ''), 6)
                 + ' e.g. ' + String(n[0].getFirstMessageSubject()).slice(0, 68));
    }
  });
  Logger.log('');

  // ── PART 2 ───────────────────────────────────────────────────────────────
  Logger.log('=== PART 2: what follows a confirmation, once the meeting is past ===');
  var today = Utilities.formatDate(new Date(), 'Asia/Manila', 'yyyy-MM-dd');
  var threads = GmailApp.search(SCAN_QUERY_SUBJECT + ' newer_than:' + OSCAN_DAYS + 'd',
                                0, OSCAN_MAX_THREADS);

  var past = [], withExtra = 0, totalExtra = 0;
  threads.forEach(function (th) {
    var p = scanParseSubject_(th.getFirstMessageSubject());
    if (!p) { return; }
    p.meetingDate = scanMeetingDate_(p.whenRaw);
    if (!p.meetingDate || p.meetingDate >= today) { return; }   // not finished yet
    var n = th.getMessageCount();
    if (n > 1) { withExtra++; totalExtra += (n - 1); }
    past.push({ th: th, p: p, extra: n - 1 });
  });

  Logger.log('confirmation threads whose meeting is past : ' + past.length);
  Logger.log('   of those, ones with a later message      : ' + withExtra
             + '   (' + totalExtra + ' later messages in total)');
  Logger.log('   with nothing after the confirmation      : ' + (past.length - withExtra));
  Logger.log('');
  if (!withExtra) {
    Logger.log('   Nothing follows a confirmation once the meeting is done, so the mailbox');
    Logger.log('   cannot tell us whether it went ahead. Show-up has to come from elsewhere.');
    return { past: past.length, withExtra: 0 };
  }

  Logger.log('READING THE LATER MESSAGES ON UP TO ' + OSCAN_SAMPLE + ' OF THEM');
  Logger.log('');
  var shown = 0;
  past.forEach(function (row) {
    if (shown >= OSCAN_SAMPLE || !row.extra) { return; }
    shown++;
    var msgs = row.th.getMessages();
    Logger.log('[' + row.p.meetingDate + '] ' + row.p.lead + '  /  ' + row.p.partner
               + '   (' + row.extra + ' later message' + (row.extra === 1 ? '' : 's') + ')');
    msgs.slice(1).forEach(function (m) {
      var who = String(m.getFrom() || '').match(/[\w.\-+]+@[\w.\-]+/);
      var body = '';
      try { body = String(m.getPlainBody() || '').replace(/\s+/g, ' ').trim(); } catch (e) {}
      Logger.log('      ' + Utilities.formatDate(m.getDate(), 'Asia/Manila', 'MM-dd HH:mm')
                 + '  from ' + (who ? who[0] : '?'));
      Logger.log('      subj: ' + String(m.getSubject() || '').slice(0, 90));
      Logger.log('      body: ' + body.slice(0, 160));
    });
    Logger.log('');
  });

  return { past: past.length, withExtra: withExtra, laterMessages: totalExtra };
}

/**
 * Which spreadsheet is this script actually attached to, and what is on it?
 *
 * Run this first when a tab you expected is not there. A script created from a
 * new sheet is bound to THAT sheet, not to the one the dashboard reads, so the
 * code can run perfectly and write its tab somewhere nobody is looking. This
 * prints the URL to open and the tabs that exist, which settles it in one run.
 */
function whereAmI() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) {
    Logger.log('This script is NOT attached to any spreadsheet.');
    Logger.log('It is a standalone project, so syncCalendars, syncAbsences and');
    Logger.log('writeBookingsTab have nowhere to write. Paste this code into the');
    Logger.log('Apps Script project that belongs to the DC Assignments sheet instead:');
    Logger.log('open that sheet, then Extensions > Apps Script.');
    return { bound: false };
  }
  Logger.log('ATTACHED TO : ' + ss.getName());
  Logger.log('URL         : ' + ss.getUrl());
  Logger.log('');
  var names = ss.getSheets().map(function (s) { return s.getName(); });
  Logger.log('TABS ON IT  : ' + names.join(' | '));
  Logger.log('');
  var want = [TAB, ABS_TAB, BOOK_TAB];
  want.forEach(function (w) {
    var hit = names.indexOf(w) >= 0;
    var sh = hit ? ss.getSheetByName(w) : null;
    Logger.log('   ' + pad_(w, 18) + (hit ? 'yes, ' + Math.max(0, sh.getLastRow() - 1) + ' rows'
                                          : 'MISSING'));
  });
  Logger.log('');
  if (names.indexOf(TAB) < 0) {
    Logger.log('There is no "' + TAB + '" tab here, so this is NOT the sheet the dashboard');
    Logger.log('reads. Open the right spreadsheet, go to Extensions > Apps Script, and');
    Logger.log('paste the code there. Publishing a tab from this one would publish nothing.');
  }
  return { bound: true, name: ss.getName(), url: ss.getUrl(), tabs: names };
}
