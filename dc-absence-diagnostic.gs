/**
 * ABSENCE SUBJECT DIAGNOSTIC — add as a SECOND file, not a replacement
 * ---------------------------------------------------------------------
 * In the Apps Script editor click the + beside Files, choose Script, name it
 * Absences, and paste this in. Apps Script files share one scope, so this can
 * use ABS_QUERY, ABS_MAX and absStatus_ from Code.gs without anything being
 * moved or duplicated. Kept separate because the main file is now 1,400 lines
 * and pasting it whole has truncated three times.
 *
 * Then run scanAbsenceSubjects() and read the log. It writes nothing.
 *
 * Why this exists: a notice the classifier does not recognise is dropped in
 * silence, and a dropped notice looks exactly like a day nobody was off. The
 * last run scanned 303 messages and recognised 68, so if real absences are
 * being missed they are among the other 235, and the only way to know is to
 * read the subjects rather than guess at them.
 */

function scanAbsenceSubjects() {
  var threads = GmailApp.search(ABS_QUERY, 0, ABS_MAX);
  var seen = {}, order = [], kept = 0, dropped = 0;

  threads.forEach(function (th) {
    th.getMessages().forEach(function (msg) {
      var subj = String(msg.getSubject() || '').trim();
      if (!subj) { subj = '(no subject)'; }
      if (seen[subj]) { seen[subj].n++; return; }
      var from = String(msg.getFrom() || '').match(/[\w.\-+]+@[\w.\-]+/);
      var st = absStatus_(subj);
      seen[subj] = { n: 1, st: st,
                     from: from ? from[0].toLowerCase() : '?',
                     when: Utilities.formatDate(msg.getDate(), 'Asia/Manila', 'MM-dd') };
      order.push(subj);
      if (st) { kept++; } else { dropped++; }
    });
  });

  var p = function (s, n) {
    s = String(s === null || s === undefined ? '' : s);
    while (s.length < n) { s += ' '; }
    return s.length > n ? s.slice(0, n) : s;
  };

  Logger.log('QUERY : ' + ABS_QUERY);
  Logger.log('threads ' + threads.length + ', distinct subjects ' + order.length);
  Logger.log('recognised ' + kept + ', ignored ' + dropped);
  Logger.log('');

  Logger.log('================ RECOGNISED ================');
  order.forEach(function (s) {
    var r = seen[s];
    if (!r.st) { return; }
    Logger.log('  ' + r.when + '  ' + p(r.st, 11) + p(r.n + 'x', 5) + s.slice(0, 76));
  });

  Logger.log('');
  Logger.log('================ IGNORED — look here for real absences ================');
  Logger.log('Anything below that is actually someone saying they will be off is a');
  Logger.log('wording the classifier does not know. Send these to Claude.');
  Logger.log('');
  order.forEach(function (s) {
    var r = seen[s];
    if (r.st) { return; }
    Logger.log('  ' + r.when + '  ' + p(r.n + 'x', 5) + p(r.from, 34) + s.slice(0, 70));
  });

  return { subjects: order.length, recognised: kept, ignored: dropped };
}
