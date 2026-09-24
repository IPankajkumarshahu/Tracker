// Unit tests for the pure parts of apps_script/UploadingSheetFiller.gs.
// Run: node --test tests/
const test = require('node:test');
const assert = require('node:assert');
const fs = require('fs');
const path = require('path');
const vm = require('vm');

const IST_MS = 330 * 60000;
// Arrays built inside the vm context have a different prototype: compare as JSON.
const eq = (actual, expected) => assert.deepStrictEqual(JSON.parse(JSON.stringify(actual)), expected);

// Minimal Apps Script stand-ins; the tests use the Asia/Kolkata time zone only.
const Utilities = {
  formatDate(date, tz, fmt) {
    const d = new Date(date.getTime() + IST_MS);
    const pad = (n) => String(n).padStart(2, '0');
    return fmt
      .replace('yyyy', d.getUTCFullYear())
      .replace('MM', pad(d.getUTCMonth() + 1))
      .replace('dd', pad(d.getUTCDate()))
      .replace('HH', pad(d.getUTCHours()))
      .replace('mm', pad(d.getUTCMinutes()));
  },
  parseDate(s, tz, fmt) {
    const m = s.match(/^(\d{4})-(\d{2})-(\d{2}) (\d{2}):(\d{2})$/);
    return new Date(Date.UTC(+m[1], m[2] - 1, +m[3], +m[4], +m[5]) - IST_MS);
  },
};

function loadScript() {
  const ctx = vm.createContext({ Utilities });
  const src = fs.readFileSync(path.join(__dirname, '..', 'apps_script', 'UploadingSheetFiller.gs'), 'utf8');
  vm.runInContext(src, ctx);
  return ctx;
}

const S = loadScript();
const ist = (d) => Utilities.formatDate(d, 'Asia/Kolkata', 'yyyy-MM-dd HH:mm');

function lookups() {
  return S.parseLookups_({
    sellerId: [
      ['Seller ID', 'Insurance Company', '', '', '', '', '', '', '', '', '', 'CD name', 'CD contact', 'CD mail'],
      [2342, 'Universal Sompo General Insurance Company Limited', '', '', '', '', '', '', '', '', '', 'Pappu Yadav', 9982195739, 'pappu.yadav1@girnarsoft.co.in\t'],
      [109915, 'Go Digit', '', '', '', '', '', '', '', '', '', 'Shubham Arora', 7840801505, 'shubham.arora@girnarsoft.co.in'],
      [2356, 'Surveyor Seller', '', '', '', '', '', '', '', '', '', 'Satnam Singh', 8826865121, 'satnam.singh@girnarsoft.co.in'],
      ['', '', '', '', '', '', '', '', '', '', '', 'Ankit Kumar', 9155161636, 'ankit.kumar3@girnarsoft.com'],
      ['', '', '', '', '', '', '', '', '', '', '', 'Rahul rai ', 7827970281, 'rahul.rai@girnarsoft.com'],
    ],
    master: [
      ['Contact Person Email', 'Contact Person Name', 'Contct Person No', '', '', '', '', '', '', '', '', 'Blacklisted_Seller'],
      ['ravi.hs@godigit.com', 'Ravi HS', 9975846247, '', '', '', '', '', '', '', '', 'bad.surveyor@gmail.com'],
    ],
    mmv: [
      ['Make', 'Model'],
      ['Mahindra', 'Bolero'], ['Mahindra', 'Bolero Pick-Up'], ['Mahindra', 'Bolero Pickup'],
      ['MAHINDRA & MAHINDRA', 'Bolero'],
      ['Maruti', 'ALTO'], ['Maruti', 'ALTO K10'], ['Maruti', 'Swift'], ['Maruti', 'Swift Dzire'],
      ['Mahindra', 'SCORPIO S'], ['Mahindra', 'Scorpio'], ['Suzuki', 'Access 125'], ['Honda', 'SP 125'], ['Honda', 'City'],
    ],
    states: [
      ['State', 'City', 'Zone'],
      ['Chhattisgarh', 'Bilaspur', 'East'], ['Himachal Pradesh', 'Bilaspur(HP)', 'North'],
      ['Karnataka', 'Bellary', 'South'], ['Maharashtra', 'Mumbai', 'West'], ['Haryana', 'Gurgaon', 'North'],
      ['Tamil Nadu', 'Chennai', 'South'],
    ],
    bd: [
      ['State', 'Zone', 'CD Person', 'Manager (L1)', 'Zonal Manager (L2)', 'Other Manager', 'BD Regional Head'],
      ['Chhattisgarh', 'East', 'Pappu', 'Nilesh Santlani', 'Shaikh Aamir', 110763, 110746],
      ['Delhi/NCR', 'North', 'Manish', 'Gaurav Sharma', 'Swapnil Chalke', 123061, 158395],
    ],
    matrix: [S.US_MATRIX_HEADER].concat(S.US_MATRIX_SEED),
    rules: [S.US_RULES_HEADER].concat(S.US_RULES_SEED),
  });
}

function aiOutput(over) {
  return Object.assign({
    reg_no: 'CG09JG7842', reg_no_seen_in_photos: 'CG09JG7842', claim_no: 'CL26071371',
    seller_contact_email: 'info@universalsompo.co.in', seller_contact_name: 'Mukesh Pandey',
    seller_contact_phone: '7715042270', seller_cc_emails: ['Prasad.Kawle@universalsompo.com'],
    quote_deadline_date: '25/09/2026', quote_deadline_time: '19:00',
    yard_name: 'AUTOCENTRE, Bilaspur', yard_address: 'AUTOCENTRE, Bilaspur', lying_city: 'Bilaspur',
    lying_state: 'Chhattisgarh', vehicle_category: 'Commercial Vehicle', plate_color: 'yellow',
    make_rc: 'MAHINDRA & MAHINDRA', model_rc: 'BOLERO PICK-UP FB PS 1.7 T XL', variant_rc: '',
    vehicle_in_photos: 'Mahindra Bolero Pik-Up', photos_match_rc_model: 'yes', mfg_year: '2019',
    mfg_month: '2', fuel_type: 'Diesel', ownership_serial: '1', reg_date: '23/02/2019',
    engine_no: 'TBJ1M21790', chassis_no: 'MA1ZU2TBKJ1M99551', tax_status: 'Not mentioned',
    tax_valid_upto: '', permit_valid_upto: '', fitness_valid_upto: '22/02/2027', transmission: 'Manual',
    is_luxury: false, flood: false, burnt: false, superdari: false, theft_recovery: false, transit: false,
    rc_attached: true, qc_flags: [],
  }, over || {});
}

function build(aiOver, extra) {
  const L = lookups();
  const ai = aiOutput(aiOver);
  return S.buildRow_(Object.assign({
    mail: {
      id: 'x', date: new Date('2026-09-24T10:28:24Z'), subject: 'Request for Wreck offer<CL26071371>',
      from: 'qc_b2b@girnarsoft.com', to: 'FIU.Salvage@universalsompo.com', cc: '', body: '', link: 'l',
    },
    ai, mmv: S.matchMmv_(ai.make_rc, ai.model_rc, L.mmv), files: { skipped: [] }, lookups: L,
    now: new Date('2026-09-24T12:00:00Z'), tz: 'Asia/Kolkata', sellerOverride: null, existingRegNos: {},
  }, extra || {}));
}

test('detects the seller from the original sender, ignoring internal forwards', () => {
  const L = lookups();
  const det = S.detectSeller_(['qc_b2b@girnarsoft.com', 'info@universalsompo.co.in'], L.rules);
  assert.strictEqual(det.rule.id, '2342');
  assert.strictEqual(S.detectSeller_(['someone@hq.godigit.com'], L.rules).rule.key, 'GODIGIT');
  assert.strictEqual(S.detectSeller_(['surveyor.x@gmail.com'], L.rules).rule.key, 'SURVEYOR');
  assert.strictEqual(S.detectSeller_(['a@girnarsoft.com', 'b@unknown.in'], L.rules), null);
});

test('MMV: longest model at the start of the RC model wins, rest goes to specification', () => {
  const L = lookups();
  let m = S.matchMmv_('MAHINDRA & MAHINDRA', 'BOLERO PICK-UP FB PS 1.7 T XL', L.mmv);
  assert.strictEqual(m.make, 'Mahindra');
  assert.strictEqual(m.model, 'Bolero Pick-Up');
  assert.strictEqual(m.specification, 'FB PS 1.7 T XL');
  m = S.matchMmv_('MARUTI SUZUKI INDIA LTD', 'ALTO K10 VXI', L.mmv);
  eq([m.make, m.model, m.specification], ['Maruti', 'ALTO K10', 'VXI']);
  m = S.matchMmv_('MARUTI SUZUKI INDIA LTD', 'MARUTI SWIFT DZIRE VDI', L.mmv);
  eq([m.make, m.model, m.specification], ['Maruti', 'Swift Dzire', 'VDI']);
  m = S.matchMmv_('HONDA MOTORCYCLE & SCOOTER INDIA', 'SP 125 DISK', L.mmv);
  eq([m.make, m.model, m.specification], ['Honda', 'SP 125', 'DISK']);
  m = S.matchMmv_('MAHINDRA AND MAHINDRA LIMITED', 'SCORPIO S11', L.mmv);
  eq([m.make, m.model, m.specification], ['Mahindra', 'Scorpio', 'S11']);
  m = S.matchMmv_('TOYOTA KIRLOSKAR', 'INNOVA', L.mmv);
  assert.strictEqual(m.model, '');
});

test('location: city picked in the state Claude reported', () => {
  const L = lookups();
  let loc = S.lookupLocation_('Bilaspur', 'Chhattisgarh', L.states);
  eq([loc.state, loc.city, loc.zone, loc.found], ['Chhattisgarh', 'Bilaspur', 'East', true]);
  loc = S.lookupLocation_('Bilaspur', 'Himachal Pradesh', L.states);
  assert.strictEqual(loc.city, 'Bilaspur(HP)');
  loc = S.lookupLocation_('Ballari', 'Karnataka', L.states);
  eq([loc.state, loc.zone, loc.found], ['Karnataka', 'South', false]);
});

test('CD matrix: luxury, CV, surveyor and region rows', () => {
  const L = lookups();
  const regions = (city, state) => S.regionKeys_(S.lookupLocation_(city, state, L.states));
  assert.strictEqual(S.pickCdPerson_(L.matrix, 'NORMAL', regions('Bellary', 'Karnataka'), 'SBI'), 'Rahul rai');
  assert.strictEqual(S.pickCdPerson_(L.matrix, 'NORMAL', regions('Bellary', 'Karnataka'), 'GODIGIT'), 'Ashutosh Kumar');
  assert.strictEqual(S.pickCdPerson_(L.matrix, 'NORMAL', regions('Bellary', 'Karnataka'), 'SURVEYOR'), 'Shubham Arora');
  assert.strictEqual(S.pickCdPerson_(L.matrix, 'NORMAL', regions('Gurgaon', 'Haryana'), 'KOTAK'), 'Zakaullah Siddique');
  assert.strictEqual(S.pickCdPerson_(L.matrix, 'NORMAL', regions('Gurgaon', 'Haryana'), 'IFFCO'), 'Shashi Pandey');
  assert.strictEqual(S.pickCdPerson_(L.matrix, 'NORMAL', regions('Mumbai', 'Maharashtra'), 'GODIGIT'), 'Satnam Singh');
  assert.strictEqual(S.pickCdPerson_(L.matrix, 'CV', regions('Mumbai', 'Maharashtra'), 'UNIVERSAL'), 'Ankit Kumar');
  assert.strictEqual(S.pickCdPerson_(L.matrix, 'LUXURY', regions('Chennai', 'Tamil Nadu'), 'HDFC'), 'G Dinesh');
  assert.strictEqual(S.pickCdPerson_(L.matrix, 'LUXURY', regions('Mumbai', 'Maharashtra'), 'HDFC'), 'Ata Ullah Shaikh');
  assert.strictEqual(S.findContact_(L.contacts, 'Rahul rai').phone, '7827970281');
});

test('end date follows the seller closing rules', () => {
  const rule = (key) => lookups().rules.find((r) => r.key === key);
  const intim = [2026, 9, 24];
  eq(S.computeEndDate_('26/09/2026', '19:00', rule('GODIGIT'), intim).parts, [2026, 9, 26, 14, 45]);
  eq(S.computeEndDate_('26/09/2026', '18:00', rule('TATA'), intim).parts, [2026, 9, 26, 17, 30]);
  eq(S.computeEndDate_('26/09/2026', '', rule('ROYAL'), intim).parts, [2026, 9, 25, 18, 30]);
  eq(S.computeEndDate_('26/09/2026', '19:00', rule('UNIVERSAL'), intim).parts, [2026, 9, 26, 19, 0]);
  const assumed = S.computeEndDate_('', '', rule('SHRIRAM'), intim);
  eq([assumed.parts, assumed.assumed], [[2026, 9, 28, 17, 0], true]);
  eq(S.addMinutes_([2026, 9, 30, 23, 50], 15), [2026, 10, 1, 0, 5]);
});

test('summary prefixes', () => {
  const ai = aiOutput({ flood: true, tax_status: 'OTT Paid' });
  assert.strictEqual(S.buildSummary_(ai, 'UNIVERSAL', []), 'Flood vehicle // OTT Paid // ' + S.US_DEFAULT_SUMMARY);
  const burnt = aiOutput({ burnt: true, tax_status: 'Lifetime Tax Paid' });
  assert.ok(S.buildSummary_(burnt, 'GODIGIT', []).startsWith('Burnt Vehicle // Lifetime Tax Paid // Buyer Needs'));
  assert.ok(S.buildSummary_(aiOutput(), 'GODIGIT', []).startsWith('Tax details not mentioned please check on vahan // Buyer'));
});

test('full row for a forwarded Universal Sompo commercial vehicle mail', () => {
  const r = build();
  const v = r.values;
  assert.strictEqual(v['Seller ID'], 2342);
  assert.strictEqual(v['Customer Name'], 'Universal Sompo General Insurance Company Limited');
  assert.strictEqual(v['Contact Person Email'], 'info@universalsompo.co.in');
  assert.strictEqual(v['Contact Person No'], '7715042270');
  assert.strictEqual(v['ACR CC Email'], 'Prasad.Kawle@universalsompo.com');
  eq([v['State'], v['City'], v['Zone']], ['Chhattisgarh', 'Bilaspur', 'East']);
  assert.strictEqual(v['Vehicle Type'], 'Commercial Vehicle');
  assert.strictEqual(v['Reg Type'], 'Commercial');
  eq([v['Make'], v['Model'], v['Specification']], ['Mahindra', 'Bolero Pick-Up', 'FB PS 1.7 T XL']);
  assert.strictEqual(v['CD Contact Person Name'], 'Ankit Kumar');
  assert.strictEqual(v['CD Contact Person No'], '9155161636');
  assert.strictEqual(ist(v['End Date']), '2026-09-25 19:00');
  assert.strictEqual(ist(v['ACR date & time']), '2026-09-25 19:00');
  assert.strictEqual(ist(v['Intimation Date']), '2026-09-24 00:00');
  assert.strictEqual(ist(v['Fitness Expiry Date']), '2027-02-22 00:00');
  assert.strictEqual(v['Increment Amount'], 1000);
  assert.strictEqual(v['Other Manager'], '');
  eq(r.flags, ['Contact Person Name: Contact not in Saller Master Sheet; name/number taken from the mail signature.']);
});

test('private car, plate mismatch, model mismatch and missing transmission are flagged', () => {
  const r = build({
    vehicle_category: '4W', plate_color: 'white', reg_no: 'MH 12 AB 1234', reg_no_seen_in_photos: 'MH12AB1284',
    make_rc: 'MARUTI SUZUKI INDIA LTD', model_rc: 'SWIFT VXI', photos_match_rc_model: 'no',
    vehicle_in_photos: 'Maruti Alto', transmission: 'Not visible', lying_city: 'Mumbai', lying_state: 'Maharashtra',
    seller_contact_email: 'ravi.hs@godigit.com', seller_contact_name: 'Ravi', seller_contact_phone: '+91 99758 46247',
    tax_status: 'Lifetime Tax Paid', fitness_valid_upto: '',
  });
  const v = r.values;
  assert.strictEqual(v['Reg No'], 'MH12AB1234');
  assert.strictEqual(v['Reg Type'], 'Private');
  assert.strictEqual(v['Vehicle Type'], '4 Wheeler');
  assert.strictEqual(v['Seller ID'], 109915);
  assert.strictEqual(v['Contact Person Name'], 'Ravi HS');
  assert.strictEqual(v['CD Contact Person Name'], 'Satnam Singh');
  assert.strictEqual(v['Transmission'], '');
  assert.strictEqual(v['Fitness Expiry Date'], '');
  assert.ok(r.cellNotes['Reg No'].includes('MH12AB1284'));
  assert.ok(r.cellNotes['Model'].includes('Maruti Alto'));
  assert.ok(r.cellNotes['Transmission']);
  assert.ok(v['Summary'].startsWith('Lifetime Tax Paid // '));
});

test('yellow-plate car is passenger carrying and commercial; 2W gets 100 increment', () => {
  let v = build({ vehicle_category: '4W', plate_color: 'yellow' }).values;
  eq([v['Vehicle Type'], v['Reg Type']], ['Passenger Carrying Vehicle', 'Commercial']);
  v = build({ vehicle_category: '2W', plate_color: 'white' }).values;
  eq([v['Vehicle Type'], v['Reg Type'], v['Increment Amount']], ['2 Wheeler', 'Private', 100]);
  v = build({ vehicle_category: '2W', plate_color: 'yellow' }).values;
  assert.strictEqual(v['Reg Type'], 'Commercial');
});

test('surveyor seller fills BD Team columns and blacklist is flagged', () => {
  const r = build({ seller_contact_email: 'bad.surveyor@gmail.com', vehicle_category: '4W', plate_color: 'white' });
  assert.strictEqual(r.values['Seller ID'], 2356);
  assert.strictEqual(r.values['CD Contact Person Name'], 'Pappu Yadav');
  eq([r.values['Other Manager'], r.values['BD Regional Head']], ['110763', '110746']);
  assert.ok(r.cellNotes['Contact Person Email'].includes('BLACKLISTED'));
});

test('unknown seller, unknown city and duplicates are flagged', () => {
  const r = build({ seller_contact_email: 'x@unknown-insurer.in', lying_city: 'Atlantis', lying_state: '' }, {
    existingRegNos: { CG09JG7842: true },
    mail: { id: 'x', date: new Date('2026-09-24T10:28:24Z'), subject: 's', from: 'x@unknown-insurer.in', to: 'qc_b2b@girnarsoft.com', cc: '', body: '', link: 'l' },
  });
  assert.strictEqual(r.values['Seller ID'], '');
  assert.ok(r.cellNotes['Seller ID']);
  assert.ok(r.cellNotes['City']);
  assert.ok(r.cellNotes['Reg No'].includes('duplicate'));
});

test('Gmail link to message id', () => {
  assert.strictEqual(S.messageIdFromLink_('https://mail.google.com/mail/u/0/#inbox/1a0d2f66bc9439e6'), '1a0d2f66bc9439e6');
  assert.strictEqual(S.messageIdFromLink_('https://mail.google.com/mail/?authuser=a@b.com#all/thread-f:1877208737980365286|msg-f:1877208737980365286'), '1a0d2f66bc9439e6');
});

test('attachment picking spreads photos evenly', () => {
  eq(S.pickEvenly_([0, 1, 2, 3, 4, 5, 6, 7, 8, 9], 4), [0, 2, 5, 7]);
  assert.ok(S.isRcFileName_('RC_1-30-06-26-14-35-43.jpeg'));
  assert.ok(S.isRcFileName_('RC-30-06-26.jpeg'));
  assert.ok(!S.isRcFileName_('in_10-07-07-26-16-26-36.jpeg'));
  assert.ok(!S.isRcFileName_('source.jpeg'));
});

test('extraction schema is strict', () => {
  const schema = S.US_EXTRACT_SCHEMA;
  assert.strictEqual(schema.additionalProperties, false);
  eq(Object.keys(schema.properties).sort(), JSON.parse(JSON.stringify(schema.required)).sort());
  eq(Object.keys(aiOutput()).sort(), JSON.parse(JSON.stringify(schema.required)).sort());
});
