/**
 * Export starred Gmail messages to a Google Sheet (download as .xlsx via
 * File > Download > Microsoft Excel).
 *
 * Columns: Date | Time | Subject | Link | Reg No
 *
 * Setup: script.google.com > New project > paste this file > Run
 * `exportStarredMails` > approve access. The sheet URL is printed in the log.
 */

var STATE_CODES = ['AN', 'AP', 'AR', 'AS', 'BR', 'CG', 'CH', 'DD', 'DL', 'DN', 'GA', 'GJ',
  'HP', 'HR', 'JH', 'JK', 'KA', 'KL', 'LA', 'LD', 'MH', 'ML', 'MN', 'MP',
  'MZ', 'NL', 'OD', 'OR', 'PB', 'PY', 'RJ', 'SK', 'TG', 'TN', 'TR', 'TS',
  'UK', 'UA', 'UP', 'WB'];

var SEP = '[\\s\\-./]*';
var STANDARD_RE = new RegExp('(?<![A-Z0-9])([A-Z]{2})' + SEP + '(\\d{1,2})' + SEP +
  '([A-Z](?:' + SEP + '[A-Z]){0,2})' + SEP + '(\\d{1,4})(?![A-Z0-9])', 'g');
// Labelled numbers ("Regn. No. HR890648") may lack series letters.
var LABELLED_RE = new RegExp('(?:REGN?|REGISTRATION|VEH(?:ICLE)?)\\.?\\s*(?:NO|NUMBER)\\.?\\s*[:#\\-]?\\s*' +
  '([A-Z]{2})' + SEP + '(\\d{1,2})' + SEP + '()(\\d{1,4})(?![A-Z0-9])', 'g');
var BHARAT_RE = new RegExp('(?<![A-Z0-9])(\\d{2})' + SEP + '(BH)' + SEP + '(\\d{4})' + SEP +
  '([A-Z]{1,2})(?![A-Z0-9])', 'g');

function pad(s, n) {
  while (s.length < n) s = '0' + s;
  return s;
}

/** Returns unique reg numbers in `text`, normalised like MH12AB1234. */
function findRegNumbers(text) {
  if (!text) return [];
  var upper = text.toUpperCase();
  var found = [];
  var m;

  BHARAT_RE.lastIndex = 0;
  while ((m = BHARAT_RE.exec(upper)) !== null) {
    found.push({ pos: m.index, reg: m[1] + m[2] + m[3] + m[4] });
  }
  [STANDARD_RE, LABELLED_RE].forEach(function (re) {
    re.lastIndex = 0;
    while ((m = re.exec(upper)) !== null) {
      if (STATE_CODES.indexOf(m[1]) === -1) continue;
      var series = m[3].replace(/[^A-Z]/g, '');
      found.push({ pos: m.index, reg: m[1] + pad(m[2], 2) + series + pad(m[4], 4) });
    }
  });

  found.sort(function (a, b) { return a.pos - b.pos; });
  var result = [];
  found.forEach(function (f) {
    if (result.indexOf(f.reg) === -1) result.push(f.reg);
  });
  return result;
}

/** Same link format Gmail uses; authuser opens the right signed-in account. */
function mailLink(email, msgId, threadId) {
  return 'https://mail.google.com/mail/?authuser=' + email +
    '#all/thread-f:' + BigInt('0x' + threadId).toString() + '|msg-f:' + BigInt('0x' + msgId).toString();
}

function exportStarredMails() {
  var tz = Session.getScriptTimeZone();
  var email = Session.getEffectiveUser().getEmail();
  var rows = [];
  var start = 0;
  var batch;

  do {
    batch = GmailApp.search('is:starred', start, 100);
    batch.forEach(function (thread) {
      thread.getMessages().forEach(function (msg) {
        if (!msg.isStarred()) return;
        var date = msg.getDate();
        var subject = msg.getSubject() || '';
        rows.push([
          date,
          Utilities.formatDate(date, tz, 'dd-MM-yyyy'),
          Utilities.formatDate(date, tz, 'HH:mm:ss'),
          subject,
          mailLink(email, msg.getId(), thread.getId()),
          findRegNumbers(subject).join(', ')
        ]);
      });
    });
    start += batch.length;
  } while (batch.length === 100);

  rows.sort(function (a, b) { return b[0] - a[0]; }); // newest first
  rows = rows.map(function (r) { return r.slice(1); });

  var ss = SpreadsheetApp.create('Starred Mails ' + Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd HH:mm'));
  var sheet = ss.getActiveSheet();
  sheet.setName('Starred Mails');
  var headers = [['Date', 'Time', 'Subject', 'Link', 'Reg No']];
  sheet.getRange(1, 1, 1, 5).setValues(headers)
    .setFontWeight('bold').setFontColor('#ffffff').setBackground('#4472c4');
  if (rows.length) {
    // Date/Time as plain text so they stay in the dd-MM-yyyy / HH:mm:ss form.
    sheet.getRange(2, 1, rows.length, 2).setNumberFormat('@');
    sheet.getRange(2, 1, rows.length, 5).setValues(rows);
  }
  sheet.setFrozenRows(1);
  sheet.setColumnWidths(1, 2, 100);
  sheet.setColumnWidth(3, 450);
  sheet.setColumnWidth(4, 380);
  sheet.setColumnWidth(5, 150);

  Logger.log('Exported %s starred mails: %s', rows.length, ss.getUrl());
  Logger.log('Download as Excel: %s', 'https://docs.google.com/spreadsheets/d/' + ss.getId() + '/export?format=xlsx');
}
