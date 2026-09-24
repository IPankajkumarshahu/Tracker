/**
 * Uploading Sheet Filler: fills the "my work" tab of the uploading sheet from
 * starred Gmail mails.
 *
 * For every starred mail (not yet processed) it:
 *  1. reads the mail (subject, sender, CC, body) and its attachments (RC copy,
 *     vehicle photos),
 *  2. asks Claude to read the RC / photos and return the vehicle details plus
 *     QC observations,
 *  3. fills the remaining columns from the lookup tabs of this spreadsheet
 *     (Seller id, Saller Master Sheet, MMV, State List, BD Team) and from the
 *     two helper tabs this script creates (CD Matrix, Seller Rules),
 *  4. appends the row to "my work", highlights cells that need a human check
 *     (with a note on the cell) and logs every mail in the "QC Report" tab.
 *
 * Set up: open the uploading sheet as a Google Sheet, Extensions > Apps Script,
 * paste this file, save, reload the sheet, then use the "Uploading Sheet" menu.
 * See README.md for the full steps.
 *
 * Written in plain ES5 so it runs on both Apps Script runtimes.
 */

var US_CONFIG = {
  MODEL: 'claude-opus-5',
  // low keeps each mail well inside the Apps Script time limits; raise to
  // 'medium' or 'high' if extraction quality needs it.
  EFFORT: 'low',
  USE_FALLBACKS: true,
  API_URL: 'https://api.anthropic.com/v1/messages',
  MAX_RC_FILES: 4,
  MAX_PHOTOS: 8,
  MIN_PHOTO_BYTES: 15000,        // smaller images are logos / signatures
  MAX_IMAGE_BYTES: 3700000,      // API limit is 5 MB after base64
  MAX_TOTAL_BYTES: 20000000,
  MAX_BODY_CHARS: 40000,
  TIME_BUDGET_MS: 270000,        // stop starting new mails after 4.5 minutes
  INTERNAL_DOMAINS: ['girnarsoft.com', 'girnarsoft.co.in', 'cardekho.com'],
  DEFAULT_END_TIME: '17:00',
  DEFAULT_TAT_DAYS: 3
};

var US_SHEETS = {
  WORK: 'my work',
  SELLER_ID: 'Seller id',
  MASTER: 'Saller Master Sheet',
  MMV: 'MMV',
  STATES: 'State List',
  BD: 'BD Team',
  MATRIX: 'CD Matrix',
  RULES: 'Seller Rules',
  QC: 'QC Report'
};

var US_DEFAULT_SUMMARY = 'Buyer Needs to Prepare Affidavit on Rs 200 Stamp Paper Along With DD of the approved amount. ' +
  'Buyer Himself Needs To Be Present At Spot While Vehicle Lifting; If due to  any reason he is not available ,   ' +
  'than the person authorized for Lifting needs to carry valid authorization letter issued by the winning buyer ,   ' +
  'along with winning buyer & and his own KYC copies. Please refer attached images for damage description ' +
  'Cardekho will not be responsible for any MISSING Stepney ,   Battery ,   Tool Kit And Extra Accessories fitted by insured  ' +
  'No amount will be hold for NOC ,   Buyer need to pay complete approved amount within 3 days of approval Via NEFT DD. ' +
  'Cardekho Insurance company will try to provide NOC within 45 working days from the date of lifting on best effort basis.';

var US_SUPERDARI_NOTE = 'Vehicle Released on Superdari //Taking Sale/Dispose off Permission or Cancellation of Superdari is totally Insured Responsibility.';
var US_THEFT_NOTE = 'Vehicle was stolen and Recovered &  currently under police custody . Vehicle has not been released from court and buyer will be responsible to release the vehicle from court and buyer will borne all the charges Please quote your offer on as it is where it is basis. Below photographs are for indicative purpose buyers are required to physically inspect the vehicle before giving their offer. No consideration will be given once vehicle is finalized. Buyers will be responsible to lift the wreck immediately within 3 days from the date of confirmation.';
var US_TATA_NCR_NOTE = 'Buyer needs to provide a Third-Party policy before or at the time of lifting the Vehicle in With RC cases.';

// Vehicle Type values written to the sheet, by the category Claude reports.
var US_VEHICLE_TYPES = {
  '2W': '2 Wheeler',
  '3W': '3 Wheeler',
  '4W': '4 Wheeler',
  '4W yellow plate': 'Passenger Carrying Vehicle',
  'E-Rickshaw': 'E-Rickshaw',
  'Commercial Vehicle': 'Commercial Vehicle',
  'Commercial Equipment': 'Commercial Equipment',
  'Farm Equipment': 'Farm Equipment'
};
// Vehicle types handled by the pan-India CV team in the CD Matrix.
var US_CV_TYPES = ['Commercial Vehicle', 'Commercial Equipment', 'Farm Equipment'];

var US_NCR_CITIES = ['delhi', 'newdelhi', 'gurgaon', 'gurugram', 'noida', 'greaternoida', 'ghaziabad', 'faridabad'];

// Seller Rules tab, seeded from the "Seller id" and "Seller T & C" tabs.
// Email domains are the insurers' usual mail domains: check and edit in the tab.
var US_RULES_HEADER = ['Seller ID', 'Seller Key', 'Email Domains', 'End Time (HH:MM)', 'Minutes Before Given Time',
  'ACR After End (min)', 'Days Before Asked Date', 'Default TAT Days', 'Notes'];
var US_RULES_SEED = [
  [109915, 'GODIGIT', 'godigit.com', '14:45', 0, 15, 0, 3, 'End 2:45 PM, ACR 3:00 PM. CV cases: close 1 day before the asked date.'],
  [2342, 'UNIVERSAL', 'universalsompo.com, universalsompo.co.in', '', 0, 0, 0, 3, 'As per mentioned date. UP/UK closing 6:30 PM. Mention highlighted note in remarks.'],
  [41, 'LIBERTY', 'libertyinsurance.in, libertygeneralinsurance.in', '17:15', 0, 15, 0, 3, 'Closing 5:15 PM, ACR 5:30 PM pan India.'],
  [45, 'TATA', 'tataaig.com', '', 30, 0, 0, 3, '30 min before given time; AP/TS close 5:45 PM.'],
  [43, 'IFFCO', 'iffcotokio.co.in', '18:30', 30, 0, 0, 3, '30 min before given time, else 6:30 PM; East by 5:45 PM; Salvage CHD 6:00 PM.'],
  [47, 'ROYAL', 'royalsundaram.in', '18:30', 0, 0, 1, 3, '1 day before the date in the link, at 6:30 PM.'],
  [81941, 'BAJAJ', 'bajajallianz.co.in, bajajgeneral.co.in, bajajgeneral.com', '15:30', 0, 0, 0, 3, '3:30 PM with RC, 3:25 PM without RC. Auto extend always No.'],
  [2350, 'SBI', 'sbigeneral.in', '17:15', 15, 0, 0, 3, 'As per mail, 5:15 PM; 15 min before a given time.'],
  [165906, 'HDFC', 'hdfcergo.com', '16:15', 0, 15, 0, 3, 'Closing 4:15 PM, ACR 4:30 PM. PCC mandatory.'],
  [15621, 'KOTAK', 'kotakgi.com, zurichkotak.com', '16:15', 15, 0, 0, 4, '15 min before a given time.'],
  [2346, 'FUTURE', 'futuregenerali.in, generalicentral.com', '16:00', 0, 0, 0, 3, 'Closing 4 PM / date as per mail.'],
  [106257, 'RAHEJA', 'rahejaqbe.com', '16:45', 0, 0, 0, 3, 'Closing 4:45 PM.'],
  [74719, 'ACKO', 'acko.com', '15:00', 0, 0, 0, 2, 'Live 2 days, cut off 3 PM; South 3 days.'],
  [35, 'CHOLA', 'cholams.murugappa.com, cholams.com', '', 0, 0, 0, 3, '3 days TAT from intimation.'],
  [39, 'MAGMA', 'magma-hdi.co.in, magmahdi.com', '', 0, 0, 0, 10, '10 calendar days; Gujarat 3 days.'],
  [2358, 'SHRIRAM', 'shriramgi.com', '', 0, 0, 0, 4, '4 days TAT or as per mail.'],
  [80154, 'DHFL', 'dhflinsurance.com, navi.com', '', 0, 0, 0, 4, '4 days TAT.'],
  [98306, 'ZUNO', 'hizuno.com, edelweissinsurance.com, zunogi.com', '', 0, 0, 0, 4, 'Edelweiss / Zuno. 4 days TAT.'],
  [37, 'BHARTI', 'bhartiaxa.com', '', 0, 0, 0, 3, ''],
  [4034, 'PSU', 'nic.co.in, nic.in', '15:30', 0, 0, 0, 3, 'National Insurance. Upload 3 days, ack 4 days.'],
  [4113, 'PSU', 'newindia.co.in', '15:30', 0, 0, 0, 3, 'New India Assurance.'],
  [4208, 'PSU', 'uiic.co.in', '15:30', 0, 0, 0, 3, 'United India Insurance.'],
  [14184, 'PSU', 'orientalinsurance.co.in', '15:30', 0, 0, 0, 3, 'Oriental Insurance.'],
  [2356, 'SURVEYOR', 'gmail.com, yahoo.com, yahoo.co.in, rediffmail.com, hotmail.com, outlook.com, ymail.com', '15:30', 0, 0, 0, 3, 'Surveyor seller (personal mail ids).'],
  [2360, 'INSTITUTIONAL', '', '', 0, 0, 0, 3, '']
];

// CD Matrix tab, transcribed from the "Metrix" tab. Rows are checked top to
// bottom; the first row whose Bucket, State and Seller match wins.
// States: state names, "Delhi NCR", "ZONE:<zone>" or ALL. Sellers: seller keys or ALL.
var US_MATRIX_HEADER = ['Bucket', 'States', 'Sellers', 'CD Person', 'Notes'];
var US_MATRIX_SEED = [
  ['LUXURY', 'Tamil Nadu, Puducherry, Pondicherry', 'SURVEYOR, PSU', 'Yanoke', 'TN surveyor/PSU and their luxury'],
  ['LUXURY', 'Tamil Nadu, Puducherry, Pondicherry', 'ALL', 'G Dinesh', 'TN private and their luxury'],
  ['LUXURY', 'Andhra Pradesh, Telangana', 'ALL', 'Kashish Arora', 'AP TS luxury bucket'],
  ['LUXURY', 'ALL', 'ALL', 'Ata Ullah Shaikh', 'Luxury, all sellers except TN, AP & TS'],
  ['CV', 'ALL', 'GODIGIT, BAJAJ, TATA, SHRIRAM, CHOLA', 'Sophia Johnson', ''],
  ['CV', 'ALL', 'UNIVERSAL, FUTURE, RAHEJA, SBI, ACKO', 'Ankit Kumar', ''],
  ['CV', 'ALL', 'IFFCO, ZUNO, KOTAK, MAGMA, LIBERTY, HDFC, ROYAL', 'Vishal Kumar Singh', ''],
  ['CV', 'ALL', 'SURVEYOR, PSU', 'Ravi Shankar Tripathi', 'Surveyor/PSU East & North and South & West'],
  ['NORMAL', 'Tamil Nadu, Puducherry, Pondicherry', 'SURVEYOR, PSU', 'Yanoke', ''],
  ['NORMAL', 'Tamil Nadu, Puducherry, Pondicherry', 'ALL', 'G Dinesh', ''],
  ['NORMAL', 'Kerala', 'SURVEYOR, PSU', 'Mahesh Babu', 'All PSU and surveyor'],
  ['NORMAL', 'Kerala', 'ALL', 'Bhuwan', 'All private sellers'],
  ['NORMAL', 'Andhra Pradesh, Telangana', 'SBI, GODIGIT, ZUNO, FUTURE, UNIVERSAL, KOTAK, TATA, ROYAL, MAGMA, ACKO', 'Shashank Sharma', 'AP TS region'],
  ['NORMAL', 'Andhra Pradesh, Telangana', 'RAHEJA, BAJAJ, CHOLA, SHRIRAM, HDFC, LIBERTY, IFFCO, SURVEYOR, PSU', 'Utsav Gupta', 'AP TS region'],
  ['NORMAL', 'Delhi NCR, Delhi, Uttar Pradesh, Gujarat, Uttarakhand', 'SURVEYOR, PSU', 'Manish Pandey', 'Surveyor and PSU except AP TS, TN, KL'],
  ['NORMAL', 'Karnataka, West Bengal, Odisha, Orissa, Assam, Arunachal Pradesh, Meghalaya, Manipur, Mizoram, Tripura, Nagaland, Sikkim', 'SURVEYOR, PSU', 'Shubham Arora', 'Surveyor and PSU except AP TS, TN, KL'],
  ['NORMAL', 'Chhattisgarh, Madhya Pradesh, Rajasthan', 'SURVEYOR, PSU', 'Pappu Yadav', 'Surveyor and PSU except AP TS, TN, KL'],
  ['NORMAL', 'Haryana, Himachal Pradesh, Jammu & Kashmir, Jammu, Ladakh, Bihar, Jharkhand', 'SURVEYOR, PSU', 'Anuradha Kumari', 'Surveyor and PSU except AP TS, TN, KL'],
  ['NORMAL', 'Maharashtra, Goa, Punjab, Chandigarh', 'SURVEYOR, PSU', 'Kamaldeep Kaur', 'Surveyor and PSU except AP TS, TN, KL'],
  ['NORMAL', 'ZONE:East, ZONE:North', 'SURVEYOR, PSU', 'Bhuwan', 'Surveyor/PSU East & North (fallback)'],
  ['NORMAL', 'ZONE:South, ZONE:West', 'SURVEYOR, PSU', 'Ravi Shankar Tripathi', 'Surveyor/PSU South & West (fallback)'],
  ['NORMAL', 'Delhi NCR', 'BAJAJ, CHOLA, FUTURE, KOTAK, SHRIRAM, RAHEJA, HDFC', 'Zakaullah Siddique', ''],
  ['NORMAL', 'Delhi NCR', 'ALL', 'Shashi Pandey', 'Rest all cases'],
  ['NORMAL', 'Karnataka', 'IFFCO, KOTAK, ROYAL, SBI, RAHEJA, ZUNO, BAJAJ, UNIVERSAL', 'Rahul rai', 'KA region'],
  ['NORMAL', 'Karnataka', 'TATA, MAGMA, FUTURE, CHOLA, ACKO, GODIGIT, HDFC, SHRIRAM, LIBERTY, DHFL', 'Ashutosh Kumar', 'KA region'],
  ['NORMAL', 'Gujarat', 'BAJAJ, SBI, HDFC, LIBERTY', 'Kshama Sharma', ''],
  ['NORMAL', 'Gujarat', 'ACKO, FUTURE, CHOLA, MAGMA, INSTITUTIONAL, KOTAK, SHRIRAM, IFFCO', 'Rinki Kanojia', ''],
  ['NORMAL', 'Gujarat', 'TATA, RAHEJA, ROYAL, GODIGIT, ZUNO, UNIVERSAL', 'Ranveer Kaur', ''],
  ['NORMAL', 'Haryana, Punjab, Himachal Pradesh, Jammu & Kashmir, Jammu, Ladakh, Chandigarh', 'KOTAK, BAJAJ, SBI, TATA, HDFC, LIBERTY, CHOLA, RAHEJA', 'Satish Srivastava', 'HR-PB-HP-JK-CHD'],
  ['NORMAL', 'Haryana, Punjab, Himachal Pradesh, Jammu & Kashmir, Jammu, Ladakh, Chandigarh', 'IFFCO, FUTURE, GODIGIT, ROYAL, SHRIRAM, ACKO, ZUNO, UNIVERSAL, MAGMA', 'Chetan Arora', 'HR-PB-HP-JK-CHD'],
  ['NORMAL', 'Uttar Pradesh, Uttarakhand', 'TATA, SHRIRAM, RAHEJA, UNIVERSAL, BAJAJ, IFFCO, ACKO, SBI', 'Vansh Malhotra', 'UP-UK'],
  ['NORMAL', 'Uttar Pradesh, Uttarakhand', 'KOTAK, LIBERTY, ROYAL, HDFC, MAGMA, DHFL, GODIGIT, FUTURE, CHOLA, ZUNO', 'Akash Srivastava', 'UP-UK'],
  ['NORMAL', 'Maharashtra, Goa', 'ACKO, SBI, FUTURE, ROYAL, RAHEJA, UNIVERSAL', 'Harinder Kumar Singh', ''],
  ['NORMAL', 'Maharashtra, Goa', 'ALL', 'Satnam Singh', 'Rest all'],
  ['NORMAL', 'West Bengal, Bihar, Jharkhand, Odisha, Orissa, Assam, Arunachal Pradesh, Meghalaya, Manipur, Mizoram, Tripura, Nagaland, Sikkim, Chhattisgarh', 'SBI, MAGMA, FUTURE, GODIGIT, ROYAL, ACKO, SHRIRAM, UNIVERSAL, IFFCO, HDFC', 'Pankaj Kumar', 'East region'],
  ['NORMAL', 'West Bengal, Bihar, Jharkhand, Odisha, Orissa, Assam, Arunachal Pradesh, Meghalaya, Manipur, Mizoram, Tripura, Nagaland, Sikkim, Chhattisgarh', 'ALL', 'Rachna', 'East region, rest all sellers'],
  ['NORMAL', 'Madhya Pradesh', 'ALL', 'Saurabh Pare', 'All private sellers'],
  ['NORMAL', 'Rajasthan', 'ALL', 'Shalini Kumari', 'RJ: one person for all sellers']
];

var US_QC_HEADER = ['Processed At', 'Status', 'Reg No', 'Claim No', 'Seller', 'Mail Date', 'Subject', 'Mail Link', 'Message ID', 'Checks'];

// ---------------------------------------------------------------------------
// Menu
// ---------------------------------------------------------------------------

function onOpen() {
  SpreadsheetApp.getUi().createMenu('Uploading Sheet')
    .addItem('Fill from starred mails', 'fillFromStarredMails')
    .addItem('Fill from one mail...', 'fillFromMailLink')
    .addSeparator()
    .addItem('Set Claude API key...', 'setClaudeApiKey')
    .addItem('Create helper tabs', 'createHelperTabs')
    .addToUi();
}

function setClaudeApiKey() {
  var ui = SpreadsheetApp.getUi();
  var res = ui.prompt('Claude API key', 'Paste your Anthropic API key (starts with sk-ant-). ' +
    'It is stored for your Google account only.', ui.ButtonSet.OK_CANCEL);
  if (res.getSelectedButton() !== ui.Button.OK) return;
  var key = res.getResponseText().trim();
  if (!key) return;
  PropertiesService.getUserProperties().setProperty('ANTHROPIC_API_KEY', key);
  ui.alert('API key saved.');
}

/** Creates the CD Matrix, Seller Rules and QC Report tabs if they are missing. */
function createHelperTabs() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ensureTab_(ss, US_SHEETS.MATRIX, US_MATRIX_HEADER, US_MATRIX_SEED);
  ensureTab_(ss, US_SHEETS.RULES, US_RULES_HEADER, US_RULES_SEED);
  ensureTab_(ss, US_SHEETS.QC, US_QC_HEADER, []);
}

function ensureTab_(ss, name, header, rows) {
  var sheet = findSheet_(ss, name);
  if (sheet) return sheet;
  sheet = ss.insertSheet(name);
  var data = [header].concat(rows);
  sheet.getRange(1, 1, data.length, header.length).setValues(data);
  sheet.getRange(1, 1, 1, header.length).setFontWeight('bold').setBackground('#d9e1f2');
  sheet.setFrozenRows(1);
  return sheet;
}

function findSheet_(ss, name) {
  var want = name.toLowerCase().trim();
  var sheets = ss.getSheets();
  for (var i = 0; i < sheets.length; i++) {
    if (sheets[i].getName().toLowerCase().trim() === want) return sheets[i];
  }
  return null;
}

// ---------------------------------------------------------------------------
// Entry points
// ---------------------------------------------------------------------------

/** Processes every starred mail that is not yet in the QC Report tab. */
function fillFromStarredMails() {
  var started = Date.now();
  var ctx = loadContext_();
  var done = processedMessageIds_(ctx.qcSheet);
  var messages = starredMessages_();
  var filled = 0, failed = 0, left = 0;
  for (var i = 0; i < messages.length; i++) {
    var msg = messages[i];
    if (done[msg.getId()]) continue;
    if (Date.now() - started > US_CONFIG.TIME_BUDGET_MS) {
      left++;
      continue;
    }
    if (processMessageSafely_(msg, ctx, null)) filled++; else failed++;
  }
  var text = filled + ' mail(s) added to "' + US_SHEETS.WORK + '".';
  if (failed) text += ' ' + failed + ' failed (see QC Report).';
  if (left) text += ' ' + left + ' still waiting: run "Fill from starred mails" again.';
  if (!filled && !failed && !left) text = 'No new starred mails.';
  notify_(text);
}

/** Processes one mail given its Gmail link, message id or a Gmail search, starred or not. */
function fillFromMailLink() {
  var ui = SpreadsheetApp.getUi();
  var res = ui.prompt('Fill from one mail', 'Paste the Gmail link of the mail, or type something to search ' +
    'for it (e.g. the reg number or claim number).', ui.ButtonSet.OK_CANCEL);
  if (res.getSelectedButton() !== ui.Button.OK) return;
  var input = res.getResponseText().trim();
  if (!input) return;
  var msg = null;
  var id = messageIdFromLink_(input);
  if (id) {
    try {
      msg = GmailApp.getMessageById(id);
    } catch (e) {
      msg = null;
    }
  }
  if (!msg && !/^https?:/i.test(input)) {
    var threads = GmailApp.search(input, 0, 1);
    if (threads.length) {
      var msgs = threads[0].getMessages();
      msg = msgs[msgs.length - 1];
      var ok = ui.alert('Use this mail?', msg.getSubject() + '\n' + msg.getFrom() + '\n' + msg.getDate(), ui.ButtonSet.YES_NO);
      if (ok !== ui.Button.YES) return;
    }
  }
  if (!msg) {
    ui.alert('Mail not found. Try searching by the reg number or claim number instead of the link.');
    return;
  }
  var sellerRes = ui.prompt('Seller ID (optional)', 'Leave blank to detect the seller from the mail, or type ' +
    'the Seller ID to use.', ui.ButtonSet.OK_CANCEL);
  if (sellerRes.getSelectedButton() !== ui.Button.OK) return;
  var sellerId = sellerRes.getResponseText().trim();
  var ctx = loadContext_();
  var done = processMessageSafely_(msg, ctx, sellerId || null);
  notify_(done ? 'Mail added to "' + US_SHEETS.WORK + '". Check the QC Report tab.' :
    'Failed: see the QC Report tab.');
}

function notify_(text) {
  Logger.log(text);
  try {
    SpreadsheetApp.getUi().alert(text);
  } catch (e) {
    // no UI (e.g. run from a trigger)
  }
}

/** Returns true when the row was written. Failures are logged in the QC Report. */
function processMessageSafely_(msg, ctx, sellerOverride) {
  try {
    processMessage_(msg, ctx, sellerOverride);
    return true;
  } catch (e) {
    ctx.qcSheet.appendRow([new Date(), 'ERROR', '', '', '', msg.getDate(), msg.getSubject(),
      mailLinkForMessage_(ctx.email, msg), msg.getId(), String(e && e.message || e)]);
    return false;
  }
}

// ---------------------------------------------------------------------------
// Per-mail pipeline
// ---------------------------------------------------------------------------

function processMessage_(msg, ctx, sellerOverride) {
  var mail = mailInfo_(msg, ctx.email);
  var files = collectAttachments_(msg);
  var ai = extractWithClaude_(ctx.apiKey, mail, files);
  var mmv = resolveMmv_(ai.make_rc, ai.model_rc, ctx.lookups.mmv, ctx.apiKey);
  var result = buildRow_({
    mail: mail,
    ai: ai,
    mmv: mmv,
    files: files,
    lookups: ctx.lookups,
    now: new Date(),
    tz: ctx.tz,
    sellerOverride: sellerOverride,
    existingRegNos: ctx.existingRegNos
  });
  writeRow_(ctx, result);
  ctx.existingRegNos[normKey_(result.values['Reg No'])] = true;
  ctx.qcSheet.appendRow([new Date(), result.flags.length ? 'CHECK' : 'OK', result.values['Reg No'],
    result.values['Claim No'], result.values['Customer Name'], mail.date, mail.subject, mail.link, mail.id,
    result.flags.join('\n')]);
}

function loadContext_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var apiKey = PropertiesService.getUserProperties().getProperty('ANTHROPIC_API_KEY');
  if (!apiKey) throw new Error('Set your Claude API key first: Uploading Sheet > Set Claude API key.');
  createHelperTabs();
  var work = findSheet_(ss, US_SHEETS.WORK);
  if (!work) throw new Error('Tab "' + US_SHEETS.WORK + '" not found.');
  var header = findHeaderRow_(work);
  var lookups = loadLookups_(ss);
  var existing = {};
  if (header.col['Reg No'] !== undefined && work.getLastRow() > header.row) {
    work.getRange(header.row + 1, header.col['Reg No'] + 1, work.getLastRow() - header.row, 1).getValues()
      .forEach(function (r) { if (r[0]) existing[normKey_(r[0])] = true; });
  }
  return {
    ss: ss,
    tz: ss.getSpreadsheetTimeZone(),
    apiKey: apiKey,
    email: Session.getEffectiveUser().getEmail(),
    work: work,
    header: header,
    qcSheet: findSheet_(ss, US_SHEETS.QC),
    lookups: lookups,
    existingRegNos: existing
  };
}

/** Finds the row holding the column names (the one with "Start Date" and "Reg No"). */
function findHeaderRow_(sheet) {
  var rows = sheet.getRange(1, 1, Math.min(5, sheet.getLastRow()), sheet.getLastColumn()).getValues();
  for (var r = 0; r < rows.length; r++) {
    var names = rows[r].map(function (v) { return String(v).trim(); });
    if (names.indexOf('Start Date') !== -1 && names.indexOf('Reg No') !== -1) {
      var col = {};
      names.forEach(function (n, i) { if (n && col[n] === undefined) col[n] = i; });
      return { row: r + 1, col: col, width: names.length };
    }
  }
  throw new Error('Could not find the header row (with "Start Date" and "Reg No") in "' + sheet.getName() + '".');
}

function processedMessageIds_(qcSheet) {
  var done = {};
  if (qcSheet.getLastRow() < 2) return done;
  qcSheet.getRange(2, 1, qcSheet.getLastRow() - 1, US_QC_HEADER.length).getValues().forEach(function (r) {
    if (r[1] !== 'ERROR' && r[8]) done[String(r[8])] = true;
  });
  return done;
}

/** Starred messages, oldest first so rows follow the order the mails came in. */
function starredMessages_() {
  var out = [];
  var start = 0, batch;
  do {
    batch = GmailApp.search('is:starred', start, 100);
    batch.forEach(function (thread) {
      thread.getMessages().forEach(function (m) { if (m.isStarred()) out.push(m); });
    });
    start += batch.length;
  } while (batch.length === 100);
  out.sort(function (a, b) { return a.getDate() - b.getDate(); });
  return out;
}

function mailInfo_(msg, email) {
  var body = msg.getPlainBody() || '';
  if (body.length > US_CONFIG.MAX_BODY_CHARS) body = body.substring(0, US_CONFIG.MAX_BODY_CHARS);
  return {
    id: msg.getId(),
    date: msg.getDate(),
    subject: msg.getSubject() || '',
    from: msg.getFrom() || '',
    to: msg.getTo() || '',
    cc: msg.getCc() || '',
    body: body,
    link: mailLinkForMessage_(email, msg)
  };
}

function mailLinkForMessage_(email, msg) {
  return 'https://mail.google.com/mail/?authuser=' + email + '#all/' + msg.getId();
}

/**
 * Splits attachments into RC files, vehicle photos and files the tool cannot
 * read. Falls back to the other mails of the thread when the starred mail has
 * no usable attachment (e.g. a reply on a forwarded chain).
 */
function collectAttachments_(msg) {
  // Inline images are included: surveyors often paste photos into the body.
  var result = classifyAttachments_(msg.getAttachments());
  if (!result.rc.length && !result.photos.length) {
    var others = msg.getThread().getMessages();
    for (var i = others.length - 1; i >= 0; i--) {
      if (others[i].getId() === msg.getId()) continue;
      var r = classifyAttachments_(others[i].getAttachments());
      if (r.rc.length || r.photos.length) {
        r.skipped = result.skipped.concat(r.skipped);
        return r;
      }
    }
  }
  return result;
}

var US_IMAGE_TYPES = ['image/jpeg', 'image/png', 'image/gif', 'image/webp'];

function classifyAttachments_(attachments) {
  var rc = [], photos = [], skipped = [];
  attachments.forEach(function (a) {
    var name = a.getName() || '';
    var type = (a.getContentType() || '').toLowerCase();
    var size = a.getSize();
    var file = { name: name, type: type, size: size, blob: a };
    if (type === 'application/pdf') {
      if (size > US_CONFIG.MAX_IMAGE_BYTES * 2) skipped.push(name + ' (PDF too large)');
      else rc.push(file);
    } else if (US_IMAGE_TYPES.indexOf(type) !== -1) {
      if (size > US_CONFIG.MAX_IMAGE_BYTES) skipped.push(name + ' (image too large)');
      else if (isRcFileName_(name)) rc.push(file);
      else if (size >= US_CONFIG.MIN_PHOTO_BYTES) photos.push(file);
    } else if (/sheet|excel|csv/.test(type) || /\.(xlsx?|csv)$/i.test(name)) {
      skipped.push(name + ' (Excel not read: check it manually)');
    } else if (/zip|rar/.test(type) || /\.(zip|rar)$/i.test(name)) {
      skipped.push(name + ' (zip not read: check it manually)');
    } else if (/heic|heif/.test(type) || /\.hei[cf]$/i.test(name)) {
      skipped.push(name + ' (HEIC image not read)');
    }
  });
  // PDFs named like RC first, then the rest.
  rc.sort(function (a, b) { return (isRcFileName_(b.name) ? 1 : 0) - (isRcFileName_(a.name) ? 1 : 0); });
  if (rc.length > US_CONFIG.MAX_RC_FILES) rc = rc.slice(0, US_CONFIG.MAX_RC_FILES);
  photos = pickEvenly_(photos, US_CONFIG.MAX_PHOTOS);
  var total = 0;
  var keep = function (f) {
    total += f.size;
    return total <= US_CONFIG.MAX_TOTAL_BYTES;
  };
  return { rc: rc.filter(keep), photos: photos.filter(keep), skipped: skipped };
}

function isRcFileName_(name) {
  return /(^|[^a-z])rc([^a-z]|$)|registration|reg[\s_-]?cert/i.test(name);
}

/** Picks up to n items spread across the list (so front, sides, rear and interior all get a chance). */
function pickEvenly_(items, n) {
  if (items.length <= n) return items;
  var out = [];
  for (var i = 0; i < n; i++) out.push(items[Math.floor(i * items.length / n)]);
  return out;
}

/** Message id (hex) from a Gmail link (#inbox/<hex>, msg-f:<decimal>) or a bare id. */
function messageIdFromLink_(text) {
  text = String(text || '').trim();
  var m = text.match(/msg-f:(\d+)/);
  if (m) return decimalToHex_(m[1]);
  m = text.match(/#[a-z]+\/([0-9a-f]{16})\b/i) || text.match(/^([0-9a-f]{16})$/i);
  if (m) return m[1].toLowerCase();
  m = text.match(/\b(\d{17,20})\b/);
  if (m) return decimalToHex_(m[1]);
  return '';
}

/** Exact decimal-string to hex conversion (message ids exceed double precision). */
function decimalToHex_(dec) {
  var digits = [0]; // base-16 digits, least significant first
  for (var i = 0; i < dec.length; i++) {
    var carry = parseInt(dec.charAt(i), 10);
    for (var j = 0; j < digits.length; j++) {
      var v = digits[j] * 10 + carry;
      digits[j] = v % 16;
      carry = Math.floor(v / 16);
    }
    while (carry) {
      digits.push(carry % 16);
      carry = Math.floor(carry / 16);
    }
  }
  return digits.reverse().map(function (d) { return d.toString(16); }).join('');
}

// ---------------------------------------------------------------------------
// Claude
// ---------------------------------------------------------------------------

var US_EXTRACT_SYSTEM = [
  'You fill the vehicle-auction uploading sheet of CarDekho\'s insurance salvage team. Each case is one',
  'email from an insurer or surveyor asking for salvage / wreck quotes for a vehicle, with the RC',
  '(registration certificate) and vehicle photos attached. Read the mail, the RC and the photos and return',
  'the fields of the JSON schema. Use "" for anything you cannot find; never guess numbers.',
  '',
  'Where each field comes from:',
  '- reg_no: registration number, upper case without spaces or dashes (MH43CB5179). Prefer the RC.',
  '- reg_no_seen_in_photos: the number plate as readable in the vehicle photos ("" if no plate is readable).',
  '- claim_no: claim number from the subject or body. Any reference that is not a phone number or the',
  '  registration number can be the claim number (e.g. CL26217676, 202601047651).',
  '- seller_contact_*: the insurer / surveyor person who sent the request: email address, the name in the',
  '  signature and the phone number below the signature. Mails are often forwarded by the internal team',
  '  (girnarsoft.com, girnarsoft.co.in, cardekho.com addresses): then use the original external sender from',
  '  the forwarded headers or signature, never an internal address.',
  '- seller_cc_emails: external email ids copied (CC) on the seller\'s request. No internal addresses.',
  '- quote_deadline_date / quote_deadline_time: the date (DD/MM/YYYY) and 24h time (HH:MM) by which the',
  '  seller wants the quotation, if the mail gives one.',
  '- yard_name: short name of where the vehicle is lying (garage / yard name and area).',
  '  yard_address: the full lying address. lying_city: the city or district of that address.',
  '  lying_state: the Indian state of that city.',
  '- vehicle_category: 2W, 3W, 4W, E-Rickshaw, Commercial Vehicle (goods carrier, pickup, truck, bus,',
  '  tipper, tanker), Commercial Equipment (JCB, crane, excavator), Farm Equipment (tractor, harvester).',
  '  A 4-wheeler car with a yellow plate is still 4W (plate_color tells it apart).',
  '- plate_color: colour of the number plate in the photos; if not visible, infer from the RC vehicle class',
  '  (Motor Cab / Maxi Cab / Goods carrier / Transport = yellow; Motor Car / LMV non-transport /',
  '  Motor Cycle non-transport = white; EV = green).',
  '- make_rc, model_rc, variant_rc: maker, model and variant as written on the RC (or mail body), e.g.',
  '  "MARUTI SUZUKI INDIA LTD", "SWIFT DZIRE", "VXI".',
  '- vehicle_in_photos: the make and model the photographed vehicle actually looks like.',
  '- photos_match_rc_model: "no" when the photographed vehicle is clearly a different model or body type',
  '  from the RC (RC says Swift, photos show an Alto), "unclear" when photos do not allow a judgement.',
  '- mfg_year, mfg_month (number 1-12), fuel_type, ownership_serial (owner serial number), reg_date',
  '  (DD/MM/YYYY), engine_no, chassis_no: from the RC, else the mail body.',
  '- tax_status: "OTT Paid" (one time tax), "Lifetime Tax Paid" (LTT / lifetime), "Tax valid upto date"',
  '  (RC shows a tax validity date) or "Not mentioned". tax_valid_upto, permit_valid_upto,',
  '  fitness_valid_upto as DD/MM/YYYY when shown.',
  '- transmission: "Automatic" or "Manual" only when a gear lever / selector is visible in the photos or the',
  '  model / variant name says so (AMT, AT, CVT, DCT = Automatic); otherwise "Not visible".',
  '- is_luxury: true if the vehicle is worth more than Rs 60 lakh new (4-wheelers) or Rs 5 lakh new',
  '  (2-wheelers).',
  '- flood / burnt / superdari (released on superdari or legal case) / theft_recovery / transit: true only',
  '  when the mail or photos say or clearly show so.',
  '- rc_attached: true if an RC copy is among the attachments.',
  '',
  'qc_flags: short, specific findings a QC checker must look at, for example:',
  '- number plate in photos differs from the RC / mail registration number;',
  '- photographed vehicle is a different model or body type than the RC;',
  '- photos seem to show more than one vehicle;',
  '- chassis or engine number differs between RC, mail and photos;',
  '- registration number, model or year differs between the subject, mail body and RC;',
  '- RC blurred, cut off, or only one side attached;',
  '- fitness, permit or tax expired (commercial vehicles);',
  '- hypothecation / bank lien shown on the RC;',
  '- flood or fire damage visible but not mentioned in the mail;',
  '- any other inconsistency that would change the quote.',
  'Return an empty list when everything matches.'
].join('\n');

var US_EXTRACT_SCHEMA = (function () {
  var str = { type: 'string' };
  var bool = { type: 'boolean' };
  var props = {
    reg_no: str,
    reg_no_seen_in_photos: str,
    claim_no: str,
    seller_contact_email: str,
    seller_contact_name: str,
    seller_contact_phone: str,
    seller_cc_emails: { type: 'array', items: str },
    quote_deadline_date: str,
    quote_deadline_time: str,
    yard_name: str,
    yard_address: str,
    lying_city: str,
    lying_state: str,
    vehicle_category: { type: 'string', enum: ['2W', '3W', '4W', 'E-Rickshaw', 'Commercial Vehicle', 'Commercial Equipment', 'Farm Equipment'] },
    plate_color: { type: 'string', enum: ['white', 'yellow', 'green', 'black', 'unknown'] },
    make_rc: str,
    model_rc: str,
    variant_rc: str,
    vehicle_in_photos: str,
    photos_match_rc_model: { type: 'string', enum: ['yes', 'no', 'unclear'] },
    mfg_year: str,
    mfg_month: str,
    fuel_type: { type: 'string', enum: ['Petrol', 'Diesel', 'CNG', 'LPG', 'Electric', 'Hybrid', ''] },
    ownership_serial: str,
    reg_date: str,
    engine_no: str,
    chassis_no: str,
    tax_status: { type: 'string', enum: ['OTT Paid', 'Lifetime Tax Paid', 'Tax valid upto date', 'Not mentioned'] },
    tax_valid_upto: str,
    permit_valid_upto: str,
    fitness_valid_upto: str,
    transmission: { type: 'string', enum: ['Manual', 'Automatic', 'Not visible'] },
    is_luxury: bool,
    flood: bool,
    burnt: bool,
    superdari: bool,
    theft_recovery: bool,
    transit: bool,
    rc_attached: bool,
    qc_flags: { type: 'array', items: str }
  };
  return { type: 'object', properties: props, required: Object.keys(props), additionalProperties: false };
})();

function extractWithClaude_(apiKey, mail, files) {
  var content = [{
    type: 'text',
    text: 'Subject: ' + mail.subject + '\nFrom: ' + mail.from + '\nTo: ' + mail.to + '\nCc: ' + mail.cc +
      '\nDate: ' + mail.date + '\n\n' + mail.body
  }];
  if (files.rc.length) {
    content.push({ type: 'text', text: 'Attachments that look like the RC / documents:' });
    files.rc.forEach(function (f) { content.push(fileBlock_(f)); });
  }
  if (files.photos.length) {
    content.push({ type: 'text', text: 'Vehicle photos (' + files.photos.length + '):' });
    files.photos.forEach(function (f) { content.push(fileBlock_(f)); });
  }
  if (!files.rc.length && !files.photos.length) {
    content.push({ type: 'text', text: 'No usable attachments: use the mail text only.' });
  }
  content.push({ type: 'text', text: 'Fill the fields for this case.' });
  return callClaude_(apiKey, US_EXTRACT_SYSTEM, content, US_EXTRACT_SCHEMA, 16000);
}

function fileBlock_(f) {
  var data = Utilities.base64Encode(f.blob.getBytes());
  if (f.type === 'application/pdf') {
    return { type: 'document', source: { type: 'base64', media_type: 'application/pdf', data: data } };
  }
  return { type: 'image', source: { type: 'base64', media_type: f.type, data: data } };
}

/** One Messages API call with a JSON-schema response. Retries once on 429 / 5xx. */
function callClaude_(apiKey, system, content, schema, maxTokens) {
  var body = {
    model: US_CONFIG.MODEL,
    max_tokens: maxTokens,
    system: system,
    messages: [{ role: 'user', content: content }],
    output_config: { effort: US_CONFIG.EFFORT, format: { type: 'json_schema', schema: schema } }
  };
  var headers = { 'x-api-key': apiKey, 'anthropic-version': '2023-06-01' };
  if (US_CONFIG.USE_FALLBACKS) {
    body.fallbacks = 'default';
    headers['anthropic-beta'] = 'server-side-fallback-2026-07-01';
  }
  var options = {
    method: 'post',
    contentType: 'application/json',
    headers: headers,
    payload: JSON.stringify(body),
    muteHttpExceptions: true
  };
  var resp, code;
  for (var attempt = 0; attempt < 2; attempt++) {
    resp = UrlFetchApp.fetch(US_CONFIG.API_URL, options);
    code = resp.getResponseCode();
    if (code !== 429 && code < 500) break;
    Utilities.sleep(attempt === 0 ? 15000 : 0);
  }
  var json;
  try {
    json = JSON.parse(resp.getContentText());
  } catch (e) {
    throw new Error('Claude API returned HTTP ' + code + ' with an unreadable body.');
  }
  if (code !== 200) {
    throw new Error('Claude API error ' + code + ': ' + (json.error && json.error.message || resp.getContentText()));
  }
  if (json.stop_reason === 'refusal') throw new Error('Claude declined this mail; fill it by hand.');
  if (json.stop_reason === 'max_tokens') throw new Error('Claude response was cut off; try again.');
  var text = '';
  (json.content || []).forEach(function (b) { if (b.type === 'text' && !text) text = b.text; });
  if (!text) throw new Error('Claude returned no answer.');
  return JSON.parse(text);
}

/** Maps the RC make/model to the MMV tab; asks Claude to choose when the names do not line up. */
function resolveMmv_(rcMake, rcModel, mmv, apiKey) {
  var hit = matchMmv_(rcMake, rcModel, mmv);
  if (hit.model) return hit;
  var makes = hit.candidateMakes.length ? hit.candidateMakes : distinctMakes_(mmv);
  if (!rcMake && !rcModel) return hit;
  try {
    var make = hit.candidateMakes.length === 1 ? hit.candidateMakes[0] :
      claudePick_(apiKey, 'Which make from the list is the vehicle maker "' + rcMake + '" (model "' + rcModel + '")?', makes);
    if (!make) return hit;
    var models = mmv.filter(function (r) { return r[0] === make; }).map(function (r) { return String(r[1]); });
    var model = claudePick_(apiKey, 'Vehicle maker "' + rcMake + '", model on RC "' + rcModel + '". Which model from ' +
      'the list is this vehicle? Pick the base model; variant details do not need to match.', models);
    if (!model) return { make: make, model: '', specification: '', candidateMakes: [make], via: 'claude' };
    return { make: make, model: model, specification: remainderAfter_(rcModel, model), candidateMakes: [make], via: 'claude' };
  } catch (e) {
    return hit;
  }
}

function claudePick_(apiKey, question, options) {
  var schema = {
    type: 'object',
    properties: { choice: { type: 'string' } },
    required: ['choice'],
    additionalProperties: false
  };
  var text = question + '\nAnswer with one entry copied exactly from the list, or "" if none fits.\n\nList:\n' +
    options.join('\n');
  var ans = callClaude_(apiKey, 'You map vehicle names to a fixed master list.', [{ type: 'text', text: text }], schema, 2000);
  return options.indexOf(ans.choice) !== -1 ? ans.choice : '';
}

// ---------------------------------------------------------------------------
// Lookups (read from the spreadsheet)
// ---------------------------------------------------------------------------

function loadLookups_(ss) {
  var values = function (name, minCols) {
    var sh = findSheet_(ss, name);
    if (!sh || sh.getLastRow() < 1) throw new Error('Tab "' + name + '" not found or empty.');
    return sh.getRange(1, 1, sh.getLastRow(), Math.max(minCols, sh.getLastColumn())).getValues();
  };
  var mmvSheet = findSheet_(ss, US_SHEETS.MMV);
  if (!mmvSheet) throw new Error('Tab "' + US_SHEETS.MMV + '" not found.');
  return parseLookups_({
    sellerId: values(US_SHEETS.SELLER_ID, 14),
    master: values(US_SHEETS.MASTER, 12),
    mmv: mmvSheet.getRange(1, 1, mmvSheet.getLastRow(), 2).getValues(),
    states: values(US_SHEETS.STATES, 3),
    bd: values(US_SHEETS.BD, 7),
    matrix: values(US_SHEETS.MATRIX, US_MATRIX_HEADER.length),
    rules: values(US_SHEETS.RULES, US_RULES_HEADER.length)
  });
}

/** Turns raw tab values (header row included) into lookup structures. Pure; unit tested. */
function parseLookups_(raw) {
  var cell = function (v) { return String(v === null || v === undefined ? '' : v).trim(); };
  var num = function (v) { return cell(v).replace(/\.0$/, ''); };

  var sellerNames = {}, contacts = [];
  raw.sellerId.slice(1).forEach(function (r) {
    if (cell(r[0]) && cell(r[1])) sellerNames[num(r[0])] = cell(r[1]);
    if (cell(r[11])) contacts.push({ name: cell(r[11]), phone: num(r[12]), email: cell(r[13]) });
  });

  var master = {}, blacklist = {};
  raw.master.slice(1).forEach(function (r) {
    var email = cell(r[0]).toLowerCase();
    if (email && !master[email]) master[email] = { name: cell(r[1]), phone: num(r[2]) };
    var bl = cell(r[11]).toLowerCase();
    if (bl) blacklist[bl] = true;
  });

  var mmv = [];
  raw.mmv.slice(1).forEach(function (r) {
    if (cell(r[0]) && cell(r[1])) mmv.push([cell(r[0]), cell(r[1])]);
  });

  var states = [];
  raw.states.slice(1).forEach(function (r) {
    if (cell(r[0])) states.push({ state: cell(r[0]), city: cell(r[1]), zone: cell(r[2]) });
  });

  var bd = {};
  raw.bd.slice(1).forEach(function (r) {
    if (cell(r[0])) bd[normName_(r[0])] = { zone: cell(r[1]), otherManager: num(r[5]), regionalHead: num(r[6]) };
  });

  var list = function (v) {
    return cell(v).split(',').map(function (s) { return s.trim(); }).filter(function (s) { return s; });
  };
  var matrix = [];
  raw.matrix.slice(1).forEach(function (r) {
    if (!cell(r[0])) return;
    matrix.push({ bucket: cell(r[0]).toUpperCase(), states: list(r[1]), sellers: list(r[2]).map(function (s) { return s.toUpperCase(); }), person: cell(r[3]) });
  });

  var rules = [];
  raw.rules.slice(1).forEach(function (r) {
    if (!cell(r[0])) return;
    rules.push({
      id: num(r[0]),
      key: cell(r[1]).toUpperCase(),
      domains: list(r[2]).map(function (d) { return d.toLowerCase().replace(/^@/, ''); }),
      endTime: cell(r[3]),
      minutesBefore: Number(r[4]) || 0,
      acrAfter: Number(r[5]) || 0,
      daysBefore: Number(r[6]) || 0,
      tatDays: Number(r[7]) || US_CONFIG.DEFAULT_TAT_DAYS
    });
  });

  return {
    sellerNames: sellerNames, contacts: contacts, master: master, blacklist: blacklist, mmv: mmv,
    states: states, bd: bd, matrix: matrix, rules: rules
  };
}

// ---------------------------------------------------------------------------
// Row building (pure; unit tested)
// ---------------------------------------------------------------------------

/**
 * Builds the "my work" row for one mail.
 * Returns {values: {column name: value}, formats: {column: number format},
 *          cellNotes: {column: note}, flags: [text]}.
 */
function buildRow_(c) {
  var ai = c.ai, mail = c.mail, L = c.lookups;
  var v = {}, notes = {}, flags = [];
  var flag = function (column, text) {
    flags.push((column ? column + ': ' : '') + text);
    if (column) notes[column] = notes[column] ? notes[column] + '\n' + text : text;
  };

  // Seller
  var candidates = [ai.seller_contact_email || ''].concat(emailsIn_(mail.from), emailsIn_(mail.to), emailsIn_(mail.cc));
  var seller = null, sellerEmail = '';
  if (c.sellerOverride) {
    seller = ruleById_(L.rules, c.sellerOverride) || { id: String(c.sellerOverride), key: '', endTime: '', minutesBefore: 0, acrAfter: 0, daysBefore: 0, tatDays: US_CONFIG.DEFAULT_TAT_DAYS };
    sellerEmail = firstExternal_(candidates);
  } else {
    var det = detectSeller_(candidates, L.rules);
    if (det) {
      seller = det.rule;
      sellerEmail = det.email;
    } else {
      sellerEmail = firstExternal_(candidates);
      flag('Seller ID', 'Seller not identified from ' + (sellerEmail || 'the mail') + '. Fill Seller ID and Customer Name.');
    }
  }

  // Contact person: Saller Master Sheet first, then the mail signature.
  var contactEmail = String(ai.seller_contact_email || '').trim();
  if (!contactEmail || isInternal_(contactEmail)) contactEmail = sellerEmail;
  var known = L.master[contactEmail.toLowerCase()];
  if (L.blacklist[contactEmail.toLowerCase()]) flag('Contact Person Email', 'BLACKLISTED seller contact: do not upload without approval.');
  if (!known && contactEmail) flag('Contact Person Name', 'Contact not in Saller Master Sheet; name/number taken from the mail signature.');

  // Location
  var loc = lookupLocation_(ai.lying_city, ai.lying_state, L.states);
  if (!loc.found) flag('City', 'City "' + (ai.lying_city || '?') + '" not found in State List; check State, City and Zone.');
  else if (loc.ambiguous) flag('State', 'City name exists in more than one state; check the state.');

  // Vehicle
  var category = ai.vehicle_category;
  var yellow = ai.plate_color === 'yellow';
  var vehicleType = US_VEHICLE_TYPES[category === '4W' && yellow ? '4W yellow plate' : category] || category || '';
  var isPrivate = (category === '2W' || category === '4W') && !yellow;
  var regType = isPrivate ? 'Private' : 'Commercial';
  var bucket = ai.is_luxury ? 'LUXURY' : (US_CV_TYPES.indexOf(vehicleType) !== -1 ? 'CV' : 'NORMAL');

  // CD contact person
  var sellerKey = seller ? seller.key : '';
  var regions = regionKeys_(loc);
  var person = seller ? pickCdPerson_(L.matrix, bucket, regions, sellerKey) : '';
  var cd = person ? findContact_(L.contacts, person) : null;
  if (!person) flag('CD Contact Person Name', 'No CD person in CD Matrix for ' + bucket + ' / ' + (loc.state || '?') + ' / ' + (sellerKey || 'unknown seller') + '.');
  else if (!cd) flag('CD Contact Person No', '"' + person + '" not found in the contact list of the Seller id tab.');

  // Dates
  var intimation = ymdOf_(mail.date, c.tz);
  var end = computeEndDate_(ai.quote_deadline_date, ai.quote_deadline_time, seller, intimation);
  if (end.assumed) flag('End Date', 'No quotation date in the mail: End Date set to ' + end.reason + '.');
  var acr = addMinutes_(end.parts, seller ? seller.acrAfter : 0);
  var today = ymdOf_(c.now, c.tz);

  // MMV
  var mmv = c.mmv || {};
  if (!mmv.model) flag('Model', 'RC model "' + [ai.make_rc, ai.model_rc].join(' ').trim() + '" not matched to the MMV tab; pick Make/Model by hand.');
  // The MMV tab has no variants: the variant and any other text of the RC
  // model go to the free-text Specification column.
  var spec = mmv.model ? mmv.specification : String(ai.model_rc || '').trim();
  if (ai.variant_rc && normKey_(spec).indexOf(normKey_(ai.variant_rc)) === -1) spec = (spec + ' ' + ai.variant_rc).trim();

  // Reg no / QC checks
  var regNo = normKey_(ai.reg_no);
  if (!regNo) flag('Reg No', 'Registration number not found.');
  var seen = normKey_(ai.reg_no_seen_in_photos);
  if (regNo && seen && seen !== regNo) flag('Reg No', 'Number plate in photos (' + seen + ') differs from RC/mail (' + regNo + ').');
  if (regNo && c.existingRegNos && c.existingRegNos[regNo]) flag('Reg No', 'Already in "' + US_SHEETS.WORK + '": possible duplicate.');
  if (ai.photos_match_rc_model === 'no') flag('Model', 'Photos show ' + (ai.vehicle_in_photos || 'a different vehicle') + ', RC says ' + [ai.make_rc, ai.model_rc].join(' ').trim() + '.');
  if (ai.transmission === 'Not visible') flag('Transmission', 'Transmission not visible in photos: check and fill.');
  if (!ai.rc_attached) flag('RC Available', 'No RC copy found in the mail: details come from the mail body; verify.');
  if (regType === 'Commercial' && !ai.fitness_valid_upto) flag('Fitness Expiry Date', 'Commercial vehicle: fitness date not found on RC.');
  if (ai.is_luxury) flag('Luxary Vehicle', 'Marked luxury: confirm.');
  if (sellerKey === 'SURVEYOR' || sellerKey === 'PSU') flag('', 'Surveyor/PSU seller: upload for 3 days, acknowledge for 4 days, closing 3:30 PM (Seller T & C).');
  if (ai.superdari) flag('Superdari Status', 'Superdari case: upload with and without RC both (Region Id 13).');
  if (ai.theft_recovery) flag('Theft Recovery', 'Theft case: upload with and without RC both (Region Id 10).');
  (c.files && c.files.skipped || []).forEach(function (s) { flag('', 'Attachment ' + s); });
  (ai.qc_flags || []).forEach(function (s) { flag('', s); });

  // BD Team: only for surveyor seller uploads.
  var bd = null;
  if (sellerKey === 'SURVEYOR') {
    bd = L.bd[normName_(bdStateName_(loc.state, regions))];
    if (!bd) flag('Other Manager', 'No BD Team row for state "' + loc.state + '".');
  }

  var master = known || {};
  var sellerId = seller ? seller.id : '';
  var contactName = master.name || ai.seller_contact_name || '';
  var contactNo = master.phone || digitsOnly_(ai.seller_contact_phone);
  var ccEmails = uniq_((ai.seller_cc_emails || []).concat(emailsIn_(mail.cc)).filter(function (e) {
    return !isInternal_(e) && e.toLowerCase() !== contactEmail.toLowerCase();
  }));

  v['Start Date'] = partsToDate_(nowParts_(c.now, c.tz), c.tz);
  v['End Date'] = partsToDate_(end.parts, c.tz);
  v['ACR date & time'] = partsToDate_(acr, c.tz);
  v['Intimation Date'] = partsToDate_([intimation[0], intimation[1], intimation[2], 0, 0], c.tz);
  v['Bid Sheet Send Date'] = v['ACR date & time'];
  v['Contact Person Email'] = contactEmail;
  v['Contact Person Name'] = contactName;
  v['Contact Person No'] = contactNo;
  v['ACR CC Email'] = ccEmails.join(', ');
  v['Reg No'] = regNo;
  v['ACR subject line'] = mail.subject;
  v['Engine No'] = ai.engine_no;
  v['Chassis No'] = ai.chassis_no;
  v['Claim No'] = ai.claim_no;
  v['Seller ID'] = sellerId ? Number(sellerId) : '';
  v['Enable Vahan'] = 'Yes';
  v['Auto Extend'] = 'No';
  v['Auto Extend Time Interval'] = '0';
  v['Auto Extend Time'] = '0';
  v['Auto Extent Time Counter'] = '0';
  v['Bid Limit'] = 20;
  v['Yard Name'] = ai.yard_name || ai.yard_address;
  v['State'] = loc.state;
  v['City'] = loc.city;
  v['Zone'] = loc.zone;
  v['CD Contact Person Name'] = cd ? cd.name : person;
  v['CD Contact Person No'] = cd ? cd.phone : '';
  v['CD Contact Person Email'] = cd ? cd.email : '';
  v['Other Manager'] = bd ? bd.otherManager : '';
  v['BD Regional Head'] = bd ? bd.regionalHead : '';
  v['Image Identifier'] = regNo;
  v['Yard Location'] = ai.yard_address || ai.yard_name;
  v['Vehicle Type'] = vehicleType;
  v['Make'] = mmv.make || '';
  v['Model'] = mmv.model || '';
  v['Variant'] = '';
  v['Specification'] = spec;
  v['Luxary Vehicle'] = ai.is_luxury ? 'Yes' : 'No';
  v['Make Year'] = ai.mfg_year ? Number(ai.mfg_year) || ai.mfg_year : '';
  v['Make Month'] = ai.mfg_month ? Number(ai.mfg_month) || ai.mfg_month : '';
  v['Fuel Type'] = ai.fuel_type;
  v['RC Available'] = 'Yes';
  v['Ownership'] = ai.ownership_serial ? Number(ai.ownership_serial) || ai.ownership_serial : '';
  v['Reg Date'] = dmyToDate_(ai.reg_date, c.tz);
  v['Reg Type'] = regType;
  v['Tax Validity Date'] = regType === 'Commercial' ? dmyToDate_(ai.tax_valid_upto, c.tz) : '';
  v['Transmission'] = ai.transmission === 'Not visible' ? '' : ai.transmission;
  v['Permit Expiry Date'] = regType === 'Commercial' ? dmyToDate_(ai.permit_valid_upto, c.tz) : '';
  v['Fitness Expiry Date'] = regType === 'Commercial' ? dmyToDate_(ai.fitness_valid_upto, c.tz) : '';
  v['Cardekho Region Id'] = ai.superdari ? 13 : (ai.theft_recovery ? 10 : '');
  v['Flood Infected'] = ai.flood ? 'Yes' : 'No';
  v['Superdari Status'] = ai.superdari ? 'Yes' : 'No';
  v['Theft Recovery'] = ai.theft_recovery ? 'Yes' : 'No';
  v['Transit Vehicle'] = ai.transit ? 'Yes' : 'No';
  v['Owner Type'] = 'Individual';
  v['Loan Paid Off'] = 'No';
  v['Vehicle Condition'] = 'Normal';
  v['Summary'] = buildSummary_(ai, sellerKey, regions);
  v['Customer Name'] = sellerId ? (L.sellerNames[String(sellerId)] || '') : '';
  v['Increment Type'] = 'Fixed';
  v['Increment Amount'] = category === '2W' ? 100 : 1000;
  v['Acknowledgement Date'] = partsToDate_([today[0], today[1], today[2], 0, 0], c.tz);

  var formats = {
    'Start Date': 'dd/mm/yyyy hh:mm:ss',
    'End Date': 'dd/mm/yyyy hh:mm:ss',
    'ACR date & time': 'dd/mm/yyyy hh:mm:ss',
    'Intimation Date': 'dd/mm/yyyy',
    'Bid Sheet Send Date': 'dd/mm/yyyy',
    'Reg Date': 'd-mmm-yy',
    'Tax Validity Date': 'd-mmm-yy',
    'Permit Expiry Date': 'd-mmm-yy',
    'Fitness Expiry Date': 'd-mmm-yy',
    'Acknowledgement Date': 'd-mmm-yy',
    'Claim No': '@',
    'Reg No': '@',
    'Image Identifier': '@'
  };
  return { values: v, formats: formats, cellNotes: notes, flags: flags };
}

/** "Flood vehicle // OTT Paid // <default summary>". */
function buildSummary_(ai, sellerKey, regions) {
  var parts = [];
  if (ai.flood) parts.push('Flood vehicle');
  if (ai.burnt) parts.push('Burnt Vehicle');
  if (ai.superdari) parts.push(US_SUPERDARI_NOTE.replace(/\s*\/\/\s*$/, ''));
  if (ai.theft_recovery) parts.push(US_THEFT_NOTE);
  if (ai.tax_status === 'OTT Paid') parts.push('OTT Paid');
  else if (ai.tax_status === 'Lifetime Tax Paid') parts.push('Lifetime Tax Paid');
  else if (ai.tax_status === 'Tax valid upto date' && ai.tax_valid_upto) parts.push('Tax valid upto ' + ai.tax_valid_upto);
  else parts.push('Tax details not mentioned please check on vahan');
  if (sellerKey === 'TATA' && regions.indexOf('delhincr') !== -1) parts.push(US_TATA_NCR_NOTE);
  parts.push(US_DEFAULT_SUMMARY);
  return parts.join(' // ');
}

function detectSeller_(emails, rules) {
  for (var i = 0; i < emails.length; i++) {
    var email = String(emails[i] || '').toLowerCase().trim();
    var at = email.lastIndexOf('@');
    if (at === -1 || isInternal_(email)) continue;
    var domain = email.substring(at + 1);
    for (var j = 0; j < rules.length; j++) {
      for (var k = 0; k < rules[j].domains.length; k++) {
        var d = rules[j].domains[k];
        if (domain === d || domain.slice(-(d.length + 1)) === '.' + d) return { rule: rules[j], email: email };
      }
    }
  }
  return null;
}

function ruleById_(rules, id) {
  for (var i = 0; i < rules.length; i++) if (rules[i].id === String(id).trim()) return rules[i];
  return null;
}

function isInternal_(email) {
  var domain = String(email).toLowerCase().split('@').pop().trim();
  return US_CONFIG.INTERNAL_DOMAINS.indexOf(domain) !== -1;
}

function firstExternal_(emails) {
  for (var i = 0; i < emails.length; i++) {
    if (emails[i] && String(emails[i]).indexOf('@') !== -1 && !isInternal_(emails[i])) return String(emails[i]).toLowerCase();
  }
  return '';
}

function emailsIn_(text) {
  return String(text || '').match(/[A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,}/g) || [];
}

function uniq_(list) {
  var seen = {}, out = [];
  list.forEach(function (x) {
    var k = String(x).toLowerCase();
    if (x && !seen[k]) { seen[k] = true; out.push(x); }
  });
  return out;
}

function digitsOnly_(s) {
  var d = String(s || '').replace(/\D/g, '');
  if (d.length === 12 && d.indexOf('91') === 0) d = d.substring(2);
  return d;
}

function normKey_(s) {
  return String(s || '').toUpperCase().replace(/[^A-Z0-9]/g, '');
}

function normName_(s) {
  return String(s || '').toLowerCase().replace(/\(.*?\)/g, '').replace(/[^a-z]/g, '');
}

var US_STATE_ALIASES = { orissa: 'odisha', pondicherry: 'puducherry', uttaranchal: 'uttarakhand', jammu: 'jammukashmir', delhincr: 'delhi' };

function stateKey_(s) {
  var k = normName_(s);
  return US_STATE_ALIASES[k] || k;
}

/** State / City / Zone from the State List for the lying city. */
function lookupLocation_(city, stateGuess, states) {
  var c = normName_(city), sg = stateKey_(stateGuess);
  var hits = c ? states.filter(function (r) { return normName_(r.city) === c; }) : [];
  if (hits.length) {
    var inState = hits.filter(function (r) { return stateKey_(r.state) === sg; });
    var pool = inState.length ? inState : hits;
    var exact = pool.filter(function (r) { return r.city.toLowerCase() === String(city).trim().toLowerCase(); });
    var hit = (exact.length ? exact : pool)[0];
    var distinct = uniq_(hits.map(function (r) { return r.state; }));
    return { state: hit.state, city: hit.city, zone: hit.zone, found: true, ambiguous: !inState.length && distinct.length > 1, cityKey: c };
  }
  var st = states.filter(function (r) { return stateKey_(r.state) === sg; });
  if (st.length) return { state: st[0].state, city: city || '', zone: st[0].zone, found: false, cityKey: c };
  return { state: stateGuess || '', city: city || '', zone: '', found: false, cityKey: c };
}

/** Keys the CD Matrix States column is matched against. */
function regionKeys_(loc) {
  var keys = [stateKey_(loc.state)];
  if (US_NCR_CITIES.indexOf(loc.cityKey) !== -1 || stateKey_(loc.state) === 'delhi') keys.push('delhincr');
  if (loc.zone) keys.push('zone:' + loc.zone.toLowerCase());
  return keys;
}

function pickCdPerson_(matrix, bucket, regions, sellerKey) {
  for (var i = 0; i < matrix.length; i++) {
    var row = matrix[i];
    if (row.bucket !== bucket) continue;
    var stateOk = row.states.some(function (s) {
      if (s.toUpperCase() === 'ALL') return true;
      if (/^zone:/i.test(s)) return regions.indexOf(s.toLowerCase().replace(/\s+/g, '')) !== -1;
      return regions.indexOf(normName_(s) === 'delhincr' ? 'delhincr' : stateKey_(s)) !== -1;
    });
    if (!stateOk) continue;
    if (row.sellers.indexOf('ALL') === -1 && row.sellers.indexOf(sellerKey) === -1) continue;
    return row.person;
  }
  return '';
}

/** Contact row for a CD person name: exact match first, then all name words contained. */
function findContact_(contacts, name) {
  var n = normName_(name);
  for (var i = 0; i < contacts.length; i++) if (normName_(contacts[i].name) === n) return contacts[i];
  var words = String(name).toLowerCase().split(/\s+/).filter(function (w) { return w; });
  var matches = contacts.filter(function (ct) {
    var cw = String(ct.name).toLowerCase().split(/[\s.]+/);
    return words.every(function (w) { return cw.indexOf(w) !== -1; });
  });
  return matches.length === 1 ? matches[0] : null;
}

/** State name as used in the BD Team tab. */
function bdStateName_(state, regions) {
  if (regions.indexOf('delhincr') !== -1) return 'Delhi/NCR';
  var k = stateKey_(state);
  var map = {
    odisha: 'Orissa', jammukashmir: 'Jammu', uttarakhand: 'Uttarakhand', tamilnadu: 'TN / PY', puducherry: 'TN / PY',
    arunachalpradesh: 'North East', meghalaya: 'North East', manipur: 'North East', mizoram: 'North East',
    tripura: 'North East', nagaland: 'North East', sikkim: 'North East'
  };
  return map[k] || state;
}

// --- MMV matching ---------------------------------------------------------

var US_MAKE_STOPWORDS = ['LTD', 'LIMITED', 'INDIA', 'PVT', 'PRIVATE', 'CO', 'COMPANY', 'THE', 'AND', 'MOTOR', 'MOTORS',
  'MOTORCYCLE', 'MOTORCYCLES', 'SCOOTER', 'SCOOTERS', 'CARS', 'CAR', 'AUTO', 'AUTOMOBILES', 'AUTOMOTIVE',
  'INTERNATIONAL', 'CORPORATION', 'CORP', 'GROUP', 'TWO', 'WHEELERS', 'WHEELER', 'VEHICLES', 'VEHICLE', 'SALES',
  'INDUSTRIES', 'MOTOCORP', 'OF'];

function makeTokens_(s) {
  return String(s || '').toUpperCase().replace(/&/g, ' ').split(/[^A-Z0-9]+/).filter(function (t) {
    return t && US_MAKE_STOPWORDS.indexOf(t) === -1;
  });
}

function distinctMakes_(mmv) {
  return uniq_(mmv.map(function (r) { return r[0]; }));
}

/**
 * Best MMV row for an RC make/model: the MMV model must be the RC model or
 * its beginning ("BOLERO PICK-UP FB PS 1.7 T XL" -> "Bolero Pick-Up"); the
 * longest such model wins. Returns {make, model, specification, candidateMakes}.
 */
function matchMmv_(rcMake, rcModel, mmv) {
  var rcTok = makeTokens_(rcMake);
  var byMake = {}, candidateMakes = [];
  mmv.forEach(function (r) {
    if (!byMake[r[0]]) {
      byMake[r[0]] = [];
      var mt = makeTokens_(r[0]);
      var contained = function (a, b) { return a.every(function (t) { return b.indexOf(t) !== -1; }); };
      if (rcTok.length && mt.length && (contained(mt, rcTok) || contained(rcTok, mt))) candidateMakes.push(r[0]);
    }
    byMake[r[0]].push(String(r[1]));
  });
  // Also try the model without a leading make word ("MARUTI SWIFT" -> "SWIFT").
  var texts = [String(rcModel || '').trim()];
  var firstWord = texts[0].split(/\s+/)[0];
  if (firstWord && rcTok.indexOf(normKey_(firstWord)) !== -1) texts.push(texts[0].substring(firstWord.length).trim());

  var best = null;
  candidateMakes.forEach(function (make) {
    byMake[make].forEach(function (m) {
      var k = normKey_(m);
      if (k.length < 2) return;
      texts.forEach(function (text) {
        // the MMV model must spell the start of the RC model and end on a word boundary
        var end = prefixEnd_(text, k);
        if (end < 0) return;
        var score = k.length + (end === text.length ? 1000 : 0);
        if (!best || score > best.score || (score === best.score && byMake[make].length > byMake[best.make].length)) {
          best = { make: make, model: m, score: score, rest: text.substring(end) };
        }
      });
    });
  });
  if (!best) return { make: candidateMakes.length === 1 ? candidateMakes[0] : '', model: '', specification: '', candidateMakes: candidateMakes };
  return { make: best.make, model: best.model, specification: best.rest.replace(/^[\s\-\/,.]+/, '').trim(), candidateMakes: candidateMakes, via: 'name' };
}

/**
 * Index in `text` just after the letters/digits spelling `key` from its start,
 * or -1 when `key` is not the start of `text` or ends inside a word.
 */
function prefixEnd_(text, key) {
  var seen = 0, i = 0;
  while (i < text.length && seen < key.length) {
    var ch = text.charAt(i).toUpperCase();
    if (/[A-Z0-9]/.test(ch)) {
      if (ch !== key.charAt(seen)) return -1;
      seen++;
    }
    i++;
  }
  if (seen < key.length) return -1;
  if (i < text.length && /[A-Za-z0-9]/.test(text.charAt(i))) return -1;
  return i;
}

/** Text of `full` left after the characters that spell `prefix` (letters/digits only). */
function remainderAfter_(full, prefix) {
  full = String(full || '');
  var want = normKey_(prefix);
  var pos = normKey_(full).indexOf(want);
  if (!want || pos < 0) return '';
  var seen = 0, i = 0;
  while (i < full.length && seen < pos + want.length) {
    if (/[A-Za-z0-9]/.test(full.charAt(i))) seen++;
    i++;
  }
  return full.substring(i).replace(/^[\s\-\/,.]+/, '').trim();
}

// --- Dates ------------------------------------------------------------------
// Dates are handled as [y, m, d, hh, mm] in the spreadsheet time zone and only
// turned into Date objects when written.

function ymdOf_(date, tz) {
  var s = Utilities.formatDate(date, tz, 'yyyy-MM-dd');
  return s.split('-').map(Number);
}

function nowParts_(date, tz) {
  return Utilities.formatDate(date, tz, 'yyyy-MM-dd-HH-mm').split('-').map(Number);
}

function partsToDate_(p, tz) {
  var pad = function (n) { return (n < 10 ? '0' : '') + n; };
  return Utilities.parseDate(p[0] + '-' + pad(p[1]) + '-' + pad(p[2]) + ' ' + pad(p[3] || 0) + ':' + pad(p[4] || 0),
    tz, 'yyyy-MM-dd HH:mm');
}

function addMinutes_(p, minutes) {
  var t = new Date(Date.UTC(p[0], p[1] - 1, p[2], p[3] || 0, p[4] || 0) + minutes * 60000);
  return [t.getUTCFullYear(), t.getUTCMonth() + 1, t.getUTCDate(), t.getUTCHours(), t.getUTCMinutes()];
}

function parseDmy_(s) {
  var m = String(s || '').match(/(\d{1,2})[\/\-.](\d{1,2})[\/\-.](\d{2,4})/);
  if (!m) return null;
  var y = Number(m[3]);
  if (y < 100) y += 2000;
  return [y, Number(m[2]), Number(m[1])];
}

function parseHm_(s) {
  var m = String(s || '').match(/(\d{1,2})[:.](\d{2})/);
  return m ? [Number(m[1]), Number(m[2])] : null;
}

function dmyToDate_(s, tz) {
  var d = parseDmy_(s);
  return d ? partsToDate_([d[0], d[1], d[2], 0, 0], tz) : '';
}

/**
 * End Date from the quotation date in the mail and the seller's closing rule.
 * Returns {parts: [y, m, d, hh, mm], assumed: bool, reason}.
 */
function computeEndDate_(deadlineDate, deadlineTime, rule, intimationYmd) {
  rule = rule || { endTime: '', minutesBefore: 0, daysBefore: 0, tatDays: US_CONFIG.DEFAULT_TAT_DAYS };
  var ruleTime = parseHm_(rule.endTime);
  var defaultTime = ruleTime || parseHm_(US_CONFIG.DEFAULT_END_TIME);
  var date = parseDmy_(deadlineDate);
  if (!date) {
    var p = addMinutes_([intimationYmd[0], intimationYmd[1], intimationYmd[2], defaultTime[0], defaultTime[1]], rule.tatDays * 1440);
    return { parts: p, assumed: true, reason: 'intimation + ' + rule.tatDays + ' days' };
  }
  var given = parseHm_(deadlineTime);
  var parts;
  if (given && rule.minutesBefore) {
    parts = addMinutes_([date[0], date[1], date[2], given[0], given[1]], -rule.minutesBefore);
  } else if (ruleTime) {
    parts = [date[0], date[1], date[2], ruleTime[0], ruleTime[1]];
  } else {
    var t = given || defaultTime;
    parts = [date[0], date[1], date[2], t[0], t[1]];
  }
  if (rule.daysBefore) parts = addMinutes_(parts, -rule.daysBefore * 1440);
  return { parts: parts, assumed: false, reason: '' };
}

// ---------------------------------------------------------------------------
// Writing
// ---------------------------------------------------------------------------

var US_CHECK_COLOR = '#fff2cc';

function writeRow_(ctx, result) {
  var sheet = ctx.work, header = ctx.header;
  var rowIndex = Math.max(sheet.getLastRow(), header.row) + 1;
  var row = [];
  for (var i = 0; i < header.width; i++) row.push('');
  Object.keys(result.values).forEach(function (name) {
    var col = header.col[name];
    if (col !== undefined) row[col] = result.values[name];
  });
  var range = sheet.getRange(rowIndex, 1, 1, header.width);
  Object.keys(result.formats).forEach(function (name) {
    var col = header.col[name];
    if (col !== undefined) sheet.getRange(rowIndex, col + 1).setNumberFormat(result.formats[name]);
  });
  range.setValues([row]);
  Object.keys(result.cellNotes).forEach(function (name) {
    var col = header.col[name];
    if (col === undefined) return;
    sheet.getRange(rowIndex, col + 1).setBackground(US_CHECK_COLOR).setNote(result.cellNotes[name]);
  });
}
