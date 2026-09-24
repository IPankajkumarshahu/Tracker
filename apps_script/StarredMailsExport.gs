/**
 * Starred Mails Tracker: one-click export of starred Gmail messages to Excel.
 *
 * Columns: Date | Time | Subject | Link | Reg No
 * Reg No = vehicle registration number from the subject; if there is none,
 * the claim number is used instead (cell highlighted yellow).
 *
 * Setup (once): create a Google Sheet > Extensions > Apps Script > paste this
 * file > Save > reload the sheet. Then use the "Starred Mails" menu:
 *   Export to Excel     - refreshes the sheet and downloads it as .xlsx
 *   Refresh sheet only  - refreshes the sheet without downloading
 */

var STATE_CODES = ['AN', 'AP', 'AR', 'AS', 'BR', 'CG', 'CH', 'DD', 'DL', 'DN', 'GA', 'GJ',
  'HP', 'HR', 'JH', 'JK', 'KA', 'KL', 'LA', 'LD', 'MH', 'ML', 'MN', 'MP',
  'MZ', 'NL', 'OD', 'OR', 'PB', 'PY', 'RJ', 'SK', 'TG', 'TN', 'TR', 'TS',
  'UK', 'UA', 'UP', 'WB'];

var SEP = '[\\s\\-./]*';
// Written for both Apps Script runtimes (old Rhino has no lookbehind or BigInt),
// so a boundary is matched as (?:^|[^A-Z0-9]) instead of a lookbehind.
var START = '(?:^|[^A-Z0-9])';
var STANDARD_RE = new RegExp(START + '([A-Z]{2})' + SEP + '(\\d{1,2})' + SEP +
  '([A-Z](?:' + SEP + '[A-Z]){0,2})' + SEP + '(\\d{1,4})(?![A-Z0-9])', 'g');
// Labelled numbers ("Regn. No. HR890648") may lack series letters.
var LABELLED_RE = new RegExp('(?:REGN?|REGISTRATION|VEH(?:ICLE)?)\\.?\\s*(?:NO|NUMBER)\\.?\\s*[:#\\-]?\\s*' +
  '([A-Z]{2})' + SEP + '(\\d{1,2})' + SEP + '()(\\d{1,4})(?![A-Z0-9])', 'g');
var BHARAT_RE = new RegExp(START + '(\\d{2})' + SEP + '(BH)' + SEP + '(\\d{4})' + SEP +
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
    found.push({ pos: m.index + m[0].indexOf(m[1]), reg: m[1] + m[2] + m[3] + m[4] });
  }
  [STANDARD_RE, LABELLED_RE].forEach(function (re) {
    re.lastIndex = 0;
    while ((m = re.exec(upper)) !== null) {
      if (STATE_CODES.indexOf(m[1]) === -1) continue;
      var series = m[3].replace(/[^A-Z]/g, '');
      found.push({ pos: m.index + m[0].indexOf(m[1]), reg: m[1] + pad(m[2], 2) + series + pad(m[4], 4) });
    }
  });

  found.sort(function (a, b) { return a.pos - b.pos; });
  var result = [];
  found.forEach(function (f) {
    if (result.indexOf(f.reg) === -1) result.push(f.reg);
  });
  return result;
}

// Claim numbers, used when a subject has no registration number.
var CLAIM_LABELLED_RE = /CLAIM\s*(?:NO|NUMBER)?[\s.:#\-]*([A-Z]{0,4}\d[A-Z0-9]*(?:\/[A-Z0-9]+)*)(?![A-Z0-9])/g;
var CLAIM_BARE_RE = /(?:^|[^A-Z0-9])(CL\d{6,}|C\d{10,})(?![A-Z0-9])/g;

/** Returns unique claim numbers in `text`, in order of appearance. */
function findClaimNumbers(text) {
  if (!text) return [];
  var upper = text.toUpperCase();
  var found = [];
  var m;
  [CLAIM_LABELLED_RE, CLAIM_BARE_RE].forEach(function (re) {
    re.lastIndex = 0;
    while ((m = re.exec(upper)) !== null) {
      found.push({ pos: m.index + m[0].indexOf(m[1]), reg: m[1] });
    }
  });
  found.sort(function (a, b) { return a.pos - b.pos; });
  var result = [];
  found.forEach(function (f) {
    if (result.indexOf(f.reg) === -1) result.push(f.reg);
  });
  return result;
}

/** Reg number(s) in `text`, else claim number(s): {value, isClaim}. */
function regOrClaim(text) {
  var regs = findRegNumbers(text);
  if (regs.length) return { value: regs.join(', '), isClaim: false };
  var claims = findClaimNumbers(text);
  return { value: claims.join(', '), isClaim: claims.length > 0 };
}

/** Converts a hex id (e.g. 1a0d338b33fbc3f4) to its exact decimal string. */
function hexToDecimal(hex) {
  var digits = [0]; // base-10 digits, least significant first
  for (var i = 0; i < hex.length; i++) {
    var carry = parseInt(hex.charAt(i), 16);
    for (var j = 0; j < digits.length; j++) {
      var v = digits[j] * 16 + carry;
      digits[j] = v % 10;
      carry = Math.floor(v / 10);
    }
    while (carry) {
      digits.push(carry % 10);
      carry = Math.floor(carry / 10);
    }
  }
  return digits.reverse().join('');
}

/** Same link format Gmail uses; authuser opens the right signed-in account. */
function mailLink(email, msgId, threadId) {
  return 'https://mail.google.com/mail/?authuser=' + email +
    '#all/thread-f:' + hexToDecimal(threadId) + '|msg-f:' + hexToDecimal(msgId);
}

var SHEET_NAME = 'Starred Mails';
var CLAIM_COLOR = '#fff2cc';

function onOpen() {
  SpreadsheetApp.getUi().createMenu('Starred Mails')
    .addItem('Export to Excel', 'exportToExcel')
    .addItem('Refresh sheet only', 'refreshSheet')
    .addToUi();
}

/** Returns [{date, subject, link}] for every starred message, newest first. */
function fetchStarred() {
  var email = Session.getEffectiveUser().getEmail();
  var mails = [];
  var start = 0;
  var batch;
  do {
    batch = GmailApp.search('is:starred', start, 100);
    batch.forEach(function (thread) {
      thread.getMessages().forEach(function (msg) {
        if (!msg.isStarred()) return;
        mails.push({
          date: msg.getDate(),
          subject: msg.getSubject() || '',
          link: mailLink(email, msg.getId(), thread.getId())
        });
      });
    });
    start += batch.length;
  } while (batch.length === 100);
  mails.sort(function (a, b) { return b.date - a.date; });
  return mails;
}

/** Rewrites the "Starred Mails" sheet with the current starred mails. */
function refreshSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet() || SpreadsheetApp.create('Starred Mails Tracker');
  var tz = ss.getSpreadsheetTimeZone();
  var sheet = ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);
  var mails = fetchStarred();

  sheet.clear();
  sheet.getRange(1, 1, 1, 5).setValues([['Date', 'Time', 'Subject', 'Link', 'Reg No']])
    .setFontWeight('bold').setFontColor('#ffffff').setBackground('#4472c4');

  if (mails.length) {
    var values = [], links = [], backgrounds = [];
    mails.forEach(function (mail) {
      var reg = regOrClaim(mail.subject);
      values.push([
        Utilities.formatDate(mail.date, tz, 'dd-MM-yyyy'),
        Utilities.formatDate(mail.date, tz, 'HH:mm:ss'),
        mail.subject,
        reg.value
      ]);
      links.push([SpreadsheetApp.newRichTextValue().setText(mail.link).setLinkUrl(mail.link).build()]);
      backgrounds.push([reg.isClaim ? CLAIM_COLOR : null]);
    });
    var n = mails.length;
    // Date/Time as plain text so they stay in the dd-MM-yyyy / HH:mm:ss form.
    sheet.getRange(2, 1, n, 2).setNumberFormat('@');
    sheet.getRange(2, 1, n, 3).setValues(values.map(function (r) { return r.slice(0, 3); }));
    sheet.getRange(2, 4, n, 1).setRichTextValues(links);
    sheet.getRange(2, 5, n, 1).setValues(values.map(function (r) { return [r[3]]; }))
      .setBackgrounds(backgrounds);
  }
  sheet.setFrozenRows(1);
  sheet.setColumnWidths(1, 2, 100);
  sheet.setColumnWidth(3, 450);
  sheet.setColumnWidth(4, 380);
  sheet.setColumnWidth(5, 150);
  SpreadsheetApp.flush();

  Logger.log('%s starred mails written to %s', mails.length, ss.getUrl());
  return { ss: ss, sheet: sheet, count: mails.length };
}

/** Menu action: refresh the sheet, then download it as an .xlsx file. */
function exportToExcel() {
  var result = refreshSheet();
  var url = 'https://docs.google.com/spreadsheets/d/' + result.ss.getId() +
    '/export?format=xlsx&gid=' + result.sheet.getSheetId();
  var html = HtmlService.createHtmlOutput(
    '<p style="font-family:Arial">' + result.count + ' starred mails exported.</p>' +
    '<p><a id="dl" href="' + url + '" target="_blank" style="font-family:Arial;font-size:16px">' +
    'Download Excel file</a></p>' +
    '<script>window.open(document.getElementById("dl").href, "_blank");</script>'
  ).setWidth(320).setHeight(130);
  SpreadsheetApp.getUi().showModalDialog(html, 'Starred Mails');
}
