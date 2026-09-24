// In-memory stand-ins for the Apps Script services used by
// apps_script/UploadingSheetFiller.gs, so the whole flow can run under node.
const fs = require('fs');
const path = require('path');
const vm = require('vm');

const IST_MS = 330 * 60000;

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
  parseDate(s) {
    const m = s.match(/^(\d{4})-(\d{2})-(\d{2}) (\d{2}):(\d{2})$/);
    return new Date(Date.UTC(+m[1], m[2] - 1, +m[3], +m[4], +m[5]) - IST_MS);
  },
  base64Encode(bytes) {
    return Buffer.from(bytes).toString('base64');
  },
  sleep() {},
};

class Range {
  constructor(sheet, row, col, rows, cols) {
    Object.assign(this, { sheet, row, col, rows, cols });
  }
  cells(fn) {
    for (let r = 0; r < this.rows; r++) for (let c = 0; c < this.cols; c++) fn(this.row + r, this.col + c, r, c);
  }
  getValues() {
    const out = [];
    for (let r = 0; r < this.rows; r++) {
      const row = [];
      for (let c = 0; c < this.cols; c++) row.push(this.sheet.get(this.row + r, this.col + c));
      out.push(row);
    }
    return out;
  }
  setValues(values) {
    if (values.length !== this.rows || values.some((r) => r.length !== this.cols)) {
      throw new Error(`setValues: data ${values.length}x${values[0].length} does not match range ${this.rows}x${this.cols}`);
    }
    this.cells((r, c, i, j) => this.sheet.set(r, c, values[i][j]));
    return this;
  }
  meta(key, value) {
    this.cells((r, c) => { this.sheet.meta[`${r},${c}`] = Object.assign(this.sheet.meta[`${r},${c}`] || {}, { [key]: value }); });
    return this;
  }
  setNumberFormat(f) { return this.meta('format', f); }
  setBackground(b) { return this.meta('background', b); }
  setNote(n) { return this.meta('note', n); }
  setFontWeight() { return this; }
}

class Sheet {
  constructor(name, data) {
    this.name = name;
    this.data = (data || []).map((r) => r.slice());
    this.meta = {};
  }
  getName() { return this.name; }
  getLastRow() {
    for (let r = this.data.length; r > 0; r--) {
      if ((this.data[r - 1] || []).some((v) => v !== '' && v !== null && v !== undefined)) return r;
    }
    return 0;
  }
  getLastColumn() { return this.data.reduce((m, r) => Math.max(m, r.length), 0); }
  get(r, c) {
    const v = (this.data[r - 1] || [])[c - 1];
    return v === undefined ? '' : v;
  }
  set(r, c, v) {
    while (this.data.length < r) this.data.push([]);
    const row = this.data[r - 1];
    while (row.length < c) row.push('');
    row[c - 1] = v;
  }
  getRange(row, col, rows, cols) {
    if (row < 1 || col < 1) throw new Error(`getRange: bad start ${row},${col}`);
    if ((rows !== undefined && rows < 1) || (cols !== undefined && cols < 1)) throw new Error(`getRange: bad size ${rows}x${cols}`);
    return new Range(this, row, col, rows || 1, cols || 1);
  }
  appendRow(values) {
    const r = this.getLastRow() + 1;
    values.forEach((v, i) => this.set(r, i + 1, v));
  }
  setFrozenRows() {}
}

class Spreadsheet {
  constructor(tabs) {
    this.sheets = Object.keys(tabs).map((name) => new Sheet(name, tabs[name]));
  }
  getSheets() { return this.sheets; }
  insertSheet(name) {
    const s = new Sheet(name, []);
    this.sheets.push(s);
    return s;
  }
  getSpreadsheetTimeZone() { return 'Asia/Kolkata'; }
  sheet(name) { return this.sheets.find((s) => s.name === name); }
}

function attachment(name, type, size) {
  return {
    getName: () => name,
    getContentType: () => type,
    getSize: () => size,
    getBytes: () => Buffer.alloc(Math.min(size, 64), 7),
  };
}

function message(opts) {
  const msg = {
    getId: () => opts.id,
    getDate: () => opts.date,
    getSubject: () => opts.subject,
    getFrom: () => opts.from,
    getTo: () => opts.to || '',
    getCc: () => opts.cc || '',
    getPlainBody: () => opts.body || '',
    isStarred: () => opts.starred !== false,
    getAttachments: () => opts.attachments || [],
    getThread: () => ({ getMessages: () => [msg] }),
  };
  return msg;
}

/**
 * Loads the script with mocked services. `claudeReplies` is a list of
 * functions (requestBody) => JSON object returned as the model's text answer,
 * used in order. Every request is recorded in `requests`.
 */
function loadWithMocks({ tabs, messages, claudeReplies, apiKey = 'sk-ant-test' }) {
  const ss = new Spreadsheet(tabs);
  const requests = [];
  const replies = claudeReplies.slice();
  const logs = [];
  const ctx = vm.createContext({
    Utilities,
    Logger: { log: (...a) => logs.push(a.join(' ')) },
    SpreadsheetApp: {
      getActiveSpreadsheet: () => ss,
      getUi: () => { throw new Error('no UI in tests'); },
      flush() {},
    },
    PropertiesService: { getUserProperties: () => ({ getProperty: () => apiKey, setProperty() {} }) },
    Session: { getEffectiveUser: () => ({ getEmail: () => 'me@girnarsoft.com' }) },
    GmailApp: {
      search: (q, start) => (start === 0 ? messages.map((m) => ({ getMessages: () => [m] })) : []),
      getMessageById: (id) => messages.find((m) => m.getId() === id) || null,
    },
    UrlFetchApp: {
      fetch(url, options) {
        const body = JSON.parse(options.payload);
        requests.push({ url, headers: options.headers, body });
        const next = replies.shift();
        if (!next) throw new Error('unexpected Claude call');
        const answer = next(body);
        return {
          getResponseCode: () => 200,
          getContentText: () => JSON.stringify({
            stop_reason: 'end_turn',
            content: [{ type: 'thinking', thinking: '' }, { type: 'text', text: JSON.stringify(answer) }],
          }),
        };
      },
    },
  });
  const src = fs.readFileSync(path.join(__dirname, '..', 'apps_script', 'UploadingSheetFiller.gs'), 'utf8');
  vm.runInContext(src, ctx);
  return { ctx, ss, requests, logs };
}

module.exports = { loadWithMocks, message, attachment, Utilities };
