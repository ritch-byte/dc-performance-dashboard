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
  const from = new Date(now.getTime() - DAYS_BEHIND * 864e5);
  const to   = new Date(now.getTime() + DAYS_AHEAD  * 864e5);
  const iso  = function (d) { return Utilities.formatDate(d, 'Asia/Manila', "yyyy-MM-dd'T'HH:mm"); };
  const fromIso = iso(from), toIso = iso(to);

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

  // Drop what is older than the window so the tab does not grow without limit.
  const out = [];
  order.forEach(function (eid) {
    const r = have[eid];
    if (!r) { return; }
    const start = cellIso_(r[5]);
    if (start && start < fromIso) { return; }
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
  if (!s || s === 'SCHEDULED') { return ''; }           // not yet happened
  if (s.indexOf('BOTH ATTENDED') === 0 || s === 'COMPLETED' || s === 'ATTENDED') { return 'Showed'; }
  if (s.indexOf('NO SHOW') >= 0 || s.indexOf('NO_SHOW') >= 0) { return 'Lead no-show'; }
  // The partner turned up and the lead did not. Its own outcome rather than folded into
  // no-show, because on this floor it is a distinct thing that happened and gets said out loud.
  if (s.indexOf('BPO ATTENDED') === 0) { return 'Lead no-show'; }
  if (s.indexOf('CANCEL') === 0 || s.indexOf('CANCEL') > 0) { return 'Cancelled'; }
  if (s.indexOf('RESCHEDUL') >= 0) { return 'Rescheduled'; }
  return String(v || '').trim();
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
