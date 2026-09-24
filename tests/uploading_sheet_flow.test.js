// End-to-end run of fillFromStarredMails() with mocked Gmail, Sheets and Claude.
// Run: node --test 'tests/*.test.js'
const test = require('node:test');
const assert = require('node:assert');
const { loadWithMocks, message, attachment, Utilities } = require('./apps_script_mocks');

const WORK_HEADER = ['Start Date', 'End Date', 'ACR date & time', 'Intimation Date', 'Bid Sheet Send Date',
  'Contact Person Email', 'Contact Person Name', 'Contact Person No', 'ACR CC Email', '', 'Reg No',
  'ACR subject line', 'Engine No', 'Chassis No', 'Claim No', 'Seller ID', 'Enable Vahan', 'Yard Name', 'State',
  'City', 'Zone', 'CD Contact Person Name', 'CD Contact Person No', 'CD Contact Person Email', 'Other Manager',
  'BD Regional Head', 'Vehicle Type', 'Make', 'Model', 'Variant', 'Specification', 'Reg Type', 'Transmission',
  'Fitness Expiry Date', 'Summary', 'Customer Name', 'Increment Amount', 'Reserve Price', 'Reserve Price'];

function tabs() {
  return {
    'my work': [
      WORK_HEADER.map((h, i) => (i === 1 ? 'eg. 30 10 2015 11:00:00' : '')),
      WORK_HEADER,
      WORK_HEADER.map((h) => (h === 'Reg No' ? 'MH43CB5179' : h ? 'example' : 'west 29')),
    ],
    'Seller id': [
      ['Seller ID', 'Insurance Company', '', '', '', '', '', '', '', '', '', 'CD name', 'CD contact', 'CD mail'],
      [2342, 'Universal Sompo General Insurance Company Limited', '', '', '', '', '', '', '', '', '', 'Ankit Kumar', 9155161636, 'ankit.kumar3@girnarsoft.com'],
      [2356, 'Surveyor Seller', '', '', '', '', '', '', '', '', '', 'Pappu Yadav', 9982195739, 'pappu.yadav1@girnarsoft.co.in'],
    ],
    'Saller Master Sheet': [
      ['Contact Person Email', 'Contact Person Name', 'Contct Person No', '', '', '', '', '', '', '', '', 'Blacklisted_Seller'],
      ['info@universalsompo.co.in', 'Mukesh Pandey', 7715042270, '', '', '', '', '', '', '', '', ''],
    ],
    MMV: [['Make', 'Model'], ['Mahindra', 'Bolero'], ['Mahindra', 'Bolero Pick-Up'], ['Maruti', 'Swift'], ['Maruti', 'Alto'], ['Suzuki', 'Access 125']],
    'State List': [['State', 'City', 'Zone'], ['Chhattisgarh', 'Bilaspur', 'East'], ['Rajasthan', 'Kota', 'North']],
    'BD Team': [
      ['State', 'Zone', 'CD Person', 'Manager (L1)', 'Zonal Manager (L2)', 'Other Manager', 'BD Regional Head'],
      ['Rajasthan', 'North', 'Pappu', 'Gaurav Sharma', 'Swapnil Chalke', 123061, 158395],
    ],
  };
}

function extraction(over) {
  return Object.assign({
    reg_no: 'CG09JG7842', reg_no_seen_in_photos: 'CG09JG7842', claim_no: 'CL26071371',
    seller_contact_email: 'info@universalsompo.co.in', seller_contact_name: 'Mukesh Pandey',
    seller_contact_phone: '7715042270', seller_cc_emails: [], quote_deadline_date: '25/09/2026',
    quote_deadline_time: '19:00', yard_name: 'AUTOCENTRE Bilaspur', yard_address: 'AUTOCENTRE, Bilaspur',
    lying_city: 'Bilaspur', lying_state: 'Chhattisgarh', vehicle_category: 'Commercial Vehicle',
    plate_color: 'yellow', make_rc: 'MAHINDRA & MAHINDRA', model_rc: 'BOLERO PICK-UP FB PS 1.7 T XL',
    variant_rc: '', vehicle_in_photos: 'Mahindra Bolero Pik-Up', photos_match_rc_model: 'yes', mfg_year: '2019',
    mfg_month: '2', fuel_type: 'Diesel', ownership_serial: '1', reg_date: '23/02/2019', engine_no: 'TBJ1M21790',
    chassis_no: 'MA1ZU2TBKJ1M99551', tax_status: 'Not mentioned', tax_valid_upto: '', permit_valid_upto: '',
    fitness_valid_upto: '22/02/2027', transmission: 'Manual', is_luxury: false, flood: false, burnt: false,
    superdari: false, theft_recovery: false, transit: false, rc_attached: true, qc_flags: [],
  }, over || {});
}

const ist = (d) => Utilities.formatDate(d, 'Asia/Kolkata', 'yyyy-MM-dd HH:mm');

function sompoMail() {
  return message({
    id: '1a0d2f66bc9439e6',
    date: new Date('2026-09-24T10:28:24Z'),
    subject: 'Request for Wreck offer<CL26071371>/<2374/78618707/00/001>/<YASHPAL KUMAR KAUSHAL>',
    from: 'QC B2B <qc_b2b@girnarsoft.com>',
    to: 'FIU.Salvage@universalsompo.com',
    body: 'Dear Team, Please refer to the attached RC copy ... Registration No CG-09-JG-7842',
    attachments: [
      attachment('image001.png', 'image/png', 3000),
      attachment('RC_1-30-06-26-14-35-43.jpeg', 'image/jpeg', 200000),
      attachment('RC_2-30-06-26-14-35-43.jpeg', 'image/jpeg', 180000),
      ...Array.from({ length: 20 }, (_, i) => attachment(`in_${i}-07-07-26.jpeg`, 'image/jpeg', 150000)),
      attachment('huge.jpeg', 'image/jpeg', 9000000),
      attachment('list.xlsx', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', 9000),
    ],
  });
}

test('starred mail is written to "my work" and logged in QC Report', () => {
  const { ctx, ss, requests } = loadWithMocks({
    tabs: tabs(),
    messages: [sompoMail()],
    claudeReplies: [() => extraction()],
  });
  ctx.fillFromStarredMails();

  // One Claude call with RC first, then an even spread of 8 photos; logos, huge files and Excel skipped.
  assert.strictEqual(requests.length, 1);
  const req = requests[0];
  assert.strictEqual(req.url, 'https://api.anthropic.com/v1/messages');
  assert.strictEqual(req.headers['x-api-key'], 'sk-ant-test');
  assert.strictEqual(req.headers['anthropic-version'], '2023-06-01');
  assert.strictEqual(req.body.model, 'claude-opus-5');
  assert.strictEqual(req.body.output_config.format.type, 'json_schema');
  const images = req.body.messages[0].content.filter((b) => b.type === 'image');
  assert.strictEqual(images.length, 10);
  assert.ok(images.every((b) => b.source.type === 'base64' && b.source.media_type === 'image/jpeg' && b.source.data));

  // Helper tabs created.
  for (const name of ['CD Matrix', 'Seller Rules', 'QC Report']) assert.ok(ss.sheet(name), name);

  // Row 4 of "my work" (after hint row, header row and the existing example row).
  const work = ss.sheet('my work');
  assert.strictEqual(work.getLastRow(), 4);
  const row = {};
  WORK_HEADER.forEach((h, i) => { if (h && !(h in row)) row[h] = work.get(4, i + 1); });
  assert.strictEqual(row['Reg No'], 'CG09JG7842');
  assert.strictEqual(row['Seller ID'], 2342);
  assert.strictEqual(row['Customer Name'], 'Universal Sompo General Insurance Company Limited');
  assert.strictEqual(row['Contact Person Name'], 'Mukesh Pandey');
  assert.strictEqual(row['Contact Person No'], '7715042270');
  assert.deepStrictEqual([row.State, row.City, row.Zone], ['Chhattisgarh', 'Bilaspur', 'East']);
  assert.deepStrictEqual([row.Make, row.Model, row.Specification], ['Mahindra', 'Bolero Pick-Up', 'FB PS 1.7 T XL']);
  assert.strictEqual(row['CD Contact Person Name'], 'Ankit Kumar');
  assert.strictEqual(row['CD Contact Person Email'], 'ankit.kumar3@girnarsoft.com');
  assert.strictEqual(ist(row['End Date']), '2026-09-25 19:00');
  assert.ok(row.Summary.startsWith('Tax details not mentioned please check on vahan // Buyer Needs'));
  assert.strictEqual(work.get(4, 10), '', 'unnamed column left empty');
  assert.strictEqual(work.meta['4,1'].format, 'dd/mm/yyyy hh:mm:ss');
  assert.strictEqual(work.meta['4,11'].format, '@');

  // QC Report: CHECK because two attachments could not be read.
  const qc = ss.sheet('QC Report');
  assert.strictEqual(qc.getLastRow(), 2);
  assert.strictEqual(qc.get(2, 2), 'CHECK');
  assert.strictEqual(qc.get(2, 9), '1a0d2f66bc9439e6');
  assert.match(qc.get(2, 10), /huge\.jpeg \(image too large\)/);
  assert.match(qc.get(2, 10), /list\.xlsx \(Excel not read/);
});

test('second run skips mails already in QC Report; errors are retried', () => {
  let calls = 0;
  const env = loadWithMocks({
    tabs: tabs(),
    messages: [sompoMail()],
    claudeReplies: [() => { calls++; throw new Error('boom'); }, () => { calls++; return extraction(); }],
  });
  env.ctx.fillFromStarredMails();
  const qc = env.ss.sheet('QC Report');
  assert.strictEqual(qc.get(2, 2), 'ERROR');
  assert.match(qc.get(2, 10), /boom/);
  assert.strictEqual(env.ss.sheet('my work').getLastRow(), 3, 'nothing written on error');

  env.ctx.fillFromStarredMails();
  assert.strictEqual(calls, 2);
  assert.strictEqual(env.ss.sheet('my work').getLastRow(), 4);

  env.ctx.fillFromStarredMails();
  assert.strictEqual(calls, 2, 'already processed mail is not sent again');
});

test('surveyor mail with unmatched model asks Claude to pick make and model', () => {
  const surveyor = message({
    id: '1a0d000000000001',
    date: new Date('2026-09-23T05:00:00Z'),
    subject: 'Salvage quote RJ20PB3198',
    from: 'Surveyor <ramesh.surveyor@gmail.com>',
    attachments: [attachment('photo.jpg', 'image/jpeg', 100000)],
  });
  const { ctx, ss, requests } = loadWithMocks({
    tabs: tabs(),
    messages: [surveyor],
    claudeReplies: [
      () => extraction({
        reg_no: 'RJ20PB3198', reg_no_seen_in_photos: '', seller_contact_email: 'ramesh.surveyor@gmail.com',
        lying_city: 'Kota', lying_state: 'Rajasthan', vehicle_category: '4W', plate_color: 'white',
        make_rc: 'MARUTI SUZUKI INDIA LTD', model_rc: 'NEW WAGON-R VXI', quote_deadline_date: '',
        rc_attached: false, tax_status: 'OTT Paid',
      }),
      (body) => {
        assert.match(body.messages[0].content[0].text, /List:\nMaruti\nSuzuki$/);
        return { choice: 'Maruti' };
      },
      (body) => {
        assert.match(body.messages[0].content[0].text, /\nAlto$/m);
        return { choice: 'Swift' };
      },
    ],
  });
  ctx.fillFromStarredMails();
  assert.strictEqual(requests.length, 3);
  const work = ss.sheet('my work');
  const col = (h) => WORK_HEADER.indexOf(h) + 1;
  assert.strictEqual(work.get(4, col('Seller ID')), 2356);
  assert.strictEqual(work.get(4, col('Make')), 'Maruti');
  assert.strictEqual(work.get(4, col('Model')), 'Swift');
  assert.strictEqual(work.get(4, col('Reg Type')), 'Private');
  assert.strictEqual(work.get(4, col('Other Manager')), '123061');
  assert.strictEqual(work.get(4, col('CD Contact Person Name')), 'Pappu Yadav');
  assert.ok(work.get(4, col('Summary')).startsWith('OTT Paid // '));
  assert.ok(work.meta[`4,${col('End Date')}`].note.includes('No quotation date'));
  assert.strictEqual(work.meta[`4,${col('End Date')}`].background, '#fff2cc');
});
