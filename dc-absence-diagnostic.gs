/**
 * ABSENCE SUBJECT DIAGNOSTIC — a SECOND file, not a replacement
 * ---------------------------------------------------------------------
 * Ctrl+A then paste over everything, including the myFunction stub Apps
 * Script puts in a new file. Then run scanAbsenceSubjects(). It writes
 * nothing.
 *
 * The first version printed every ignored subject, and the list was mostly
 * our own assignment mails: ASSIGN_CC copies "You are covering a discovery
 * call" to sd-attendance@, so that mailbox is full of them and the real
 * answer was buried.
 *
 * So this sorts the ignored subjects into three piles. Known noise is counted
 * and not listed. Anything carrying a word people actually use when they will
 * be off — can't make it, unable, emergency, hospital, typhoon, half day,
 * early out, WFH — is listed FIRST, because that is where a missed notice
 * will be. Everything else follows, in case the wording is something nobody
 * would have guessed.
 */

// Mail that is definitely not an attendance notice. Counted, never listed.
var ABSD_NOISE = [
  'you are covering a discovery call',
  'meeting confirmation:',
  'equipment inventory',
  'fulfillment', 'fulfilment',
  'coaching-log', 'coaching log',
  'shift change', 'rebook'
];

// Words somebody uses when they will not be at their desk. None of these are
// in the classifier today, which is the point: if a missed notice exists, it
// says one of these and absStatus_ never looked for it.
var ABSD_HINTS = [
  'cant make', "can't make", 'cannot make', 'unable to', 'will not be', "won't be",
  'not be able', 'off today', 'day off', 'offline', 'out of office', 'ooo',
  'emergency', 'hospital', 'clinic', 'doctor', 'medical', 'sick', 'ill ', 'fever',
  'typhoon', 'flood', 'storm', 'brownout', 'power outage', 'no internet', 'outage',
  'half day', 'halfday', 'early out', 'early-out', 'wfh', 'work from home',
  'bereavement', 'funeral', 'family', 'maternity', 'paternity',
  'notification', 'notice', 'advisory', 'excuse', 'permission'
];

function scanAbsenceSubjects() {
  var threads = GmailApp.search(ABS_QUERY, 0, ABS_MAX);
  var seen = {}, order = [], kept = 0, noise = 0, hinted = [], other = [];

  threads.forEach(function (th) {
    th.getMessages().forEach(function (msg) {
      var subj = String(msg.getSubject() || '').trim() || '(no subject)';
      if (seen[subj]) { seen[subj].n++; return; }
      var from = String(msg.getFrom() || '').match(/[\w.\-+]+@[\w.\-]+/);
      seen[subj] = { n: 1, st: absStatus_(subj),
                     from: from ? from[0].toLowerCase() : '?',
                     when: Utilities.formatDate(msg.getDate(), 'Asia/Manila', 'MM-dd') };
      order.push(subj);
    });
  });

  var p = function (s, n) {
    s = String(s === null || s === undefined ? '' : s);
    while (s.length < n) { s += ' '; }
    return s.length > n ? s.slice(0, n) : s;
  };
  var has = function (hay, list) {
    for (var i = 0; i < list.length; i++) { if (hay.indexOf(list[i]) >= 0) { return true; } }
    return false;
  };

  order.forEach(function (s) {
    var r = seen[s], low = s.toLowerCase();
    if (r.st) { kept++; return; }
    if (has(low, ABSD_NOISE)) { noise++; return; }
    if (has(low, ABSD_HINTS)) { hinted.push(s); } else { other.push(s); }
  });

  Logger.log('threads ' + threads.length + ', distinct subjects ' + order.length);
  Logger.log('recognised as attendance : ' + kept);
  Logger.log('known noise, not listed  : ' + noise);
  Logger.log('ignored but suspicious   : ' + hinted.length);
  Logger.log('ignored, everything else : ' + other.length);
  Logger.log('');

  Logger.log('=========== IGNORED BUT LOOKS LIKE SOMEBODY WILL BE OFF ===========');
  if (!hinted.length) { Logger.log('  (none)'); }
  hinted.forEach(function (s) {
    var r = seen[s];
    Logger.log('  ' + r.when + ' ' + p(r.n + 'x', 4) + p(r.from, 34) + s.slice(0, 68));
  });

  Logger.log('');
  Logger.log('=========== IGNORED, EVERYTHING ELSE ===========');
  if (!other.length) { Logger.log('  (none)'); }
  other.forEach(function (s) {
    var r = seen[s];
    Logger.log('  ' + r.when + ' ' + p(r.n + 'x', 4) + p(r.from, 34) + s.slice(0, 68));
  });

  return { subjects: order.length, recognised: kept, noise: noise,
           suspicious: hinted.length, other: other.length };
}
