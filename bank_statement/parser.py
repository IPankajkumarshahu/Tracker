"""Read transactions out of a bank statement PDF.

Works on text-based PDFs (the ones downloaded from net banking), not scans.
No per-bank templates: the table header (Date / Narration / Debit / Credit /
Balance ...) is found by its wording, column edges come from the whitespace
between columns, and multi-line narrations are stitched back into one row.
Every row is then checked against the running balance.
"""

import datetime as dt
import re
import statistics
from dataclasses import dataclass, field

import pdfplumber


# ---------------------------------------------------------------- data model

@dataclass
class Transaction:
    date: dt.date
    description: str = ""
    value_date: dt.date | None = None
    ref: str = ""
    debit: float | None = None
    credit: float | None = None
    balance: float | None = None
    extra: dict = field(default_factory=dict)  # columns we have no fixed slot for
    page: int = 0
    check: str = ""  # "OK" / "Mismatch" / "" (no balance on this row)


@dataclass
class Statement:
    source: str
    bank: str = ""
    account_no: str = ""
    period: str = ""
    opening_balance: float | None = None
    closing_balance: float | None = None
    transactions: list = field(default_factory=list)
    extra_columns: list = field(default_factory=list)
    warnings: list = field(default_factory=list)

    @property
    def total_debit(self):
        return round(sum(t.debit or 0 for t in self.transactions), 2)

    @property
    def total_credit(self):
        return round(sum(t.credit or 0 for t in self.transactions), 2)

    @property
    def mismatches(self):
        return sum(t.check == "Mismatch" for t in self.transactions)


class StatementError(Exception):
    """The PDF could not be read as a bank statement (message is user-facing)."""


# ------------------------------------------------------------ header wording

NUMERIC_ROLES = {"debit", "credit", "amount", "balance"}


def _variants(bases, suffixes):
    return {b + s for b in bases for s in suffixes}


HEADER_KEYS = {
    "date": {"date", "txndate", "transactiondate", "trandate", "transdate", "trndate",
             "txndt", "trandt", "postdate", "postingdate", "postdt", "bookingdate",
             "entrydate", "dateoftransaction", "txnpostingdate"},
    "value_date": {"valuedate", "valuedt", "valdate", "valdt", "effectivedate"},
    "description": {"narration", "description", "particulars", "details", "remarks",
                    "transactiondetails", "transactionremarks", "transactionparticulars",
                    "transactiondescription", "narrative", "naration", "transactionnarration",
                    "accountdescription"},
    "ref": {"chqrefno", "chqno", "chequeno", "chequenumber", "refno", "refnumber",
            "referenceno", "referencenumber", "refnochequeno", "refnochqno", "chqrefnumber",
            "chequerefno", "chequereferenceno", "refchqno", "refchequeno", "instrumentno",
            "instrumentnumber", "instrno", "chq", "cheque", "utrno", "utrnumber",
            "chequedetails", "chqdetails", "chqnorefno", "chequenorefno", "instrumentid"},
    "debit": _variants(["withdrawal", "withdrawals", "withdrawl", "withdrawls", "debit",
                        "debits", "dr", "paidout", "withdrawn", "debitamount"],
                       ["", "amt", "amount", "dr", "amtdr", "amountdr"]),
    "credit": _variants(["deposit", "deposits", "credit", "credits", "cr", "paidin",
                         "creditamount", "lodgement", "lodgements"],
                        ["", "amt", "amount", "cr", "amtcr", "amountcr"]),
    "amount": {"amount", "amt", "transactionamount", "txnamount", "tranamount"},
    "drcr": {"drcr", "crdr", "type", "txntype", "debitcredit", "creditdebit", "drorcr"},
    "balance": {"balance", "bal", "closingbalance", "runningbalance", "availablebalance",
                "balanceamount", "availablebal", "closingbal", "ledgerbalance"},
}
KEY_TO_ROLE = {key: role for role, keys in HEADER_KEYS.items() for key in keys}
IGNORED_HEADER_TOKENS = {"inr", "rs", ""}
MAX_KEY_TOKENS = 4


def _norm(text):
    return re.sub(r"[^a-z]", "", text.lower())


# ------------------------------------------------------------- value parsing

MONTHS = {m: i for i, m in enumerate(
    ["jan", "feb", "mar", "apr", "may", "jun", "jul", "aug", "sep", "oct", "nov", "dec"], 1)}
DATE_RES = [
    ("dmy", re.compile(r"^(\d{1,2})[/\-.](\d{1,2})[/\-.](\d{4}|\d{2})(?!\d)")),
    ("ymd", re.compile(r"^(\d{4})[/\-.](\d{1,2})[/\-.](\d{1,2})(?!\d)")),
    ("dMy", re.compile(r"^(\d{1,2})[\s\-/.]*([A-Za-z]{3,9})[\s\-/.,']*(\d{4}|\d{2})(?!\d)")),
    ("Mdy", re.compile(r"^([A-Za-z]{3,9})[\s\-/.]*(\d{1,2}),?[\s\-/.,']*(\d{4})(?!\d)")),
]


def parse_date(text):
    """Return the date at the start of `text`, or None."""
    text = text.strip()
    for kind, rx in DATE_RES:
        m = rx.match(text)
        if not m:
            continue
        a, b, c = m.groups()
        try:
            if kind == "dmy":
                d, mo, y = int(a), int(b), int(c)
            elif kind == "ymd":
                y, mo, d = int(a), int(b), int(c)
            elif kind == "dMy":
                d, mo, y = int(a), MONTHS.get(b[:3].lower()), int(c)
            else:
                mo, d, y = MONTHS.get(a[:3].lower()), int(b), int(c)
            if mo is None:
                continue
            if y < 100:
                y += 2000
            return dt.date(y, mo, d)
        except ValueError:
            continue
    return None


NUMBER_RE = re.compile(r"\d[\d,]*(?:\.\d+)?|\.\d+")
AMOUNT_WORD_RE = re.compile(
    r"^[(\-+]?(?:₹|rs\.?|inr)?\s*\d{1,3}(?:,\d{2,3})*(?:\.\d{1,2})?\)?"
    r"(?:\s*\(?(?:cr|dr)\)?\.?)?$", re.I)
DRCR_WORD_RE = re.compile(r"^\(?(cr|dr)\)?\.?$", re.I)


def parse_amount(text):
    """Parse '1,23,456.70', '(500.00)', '-12.5', '9,000.00 Cr' -> (value, 'CR'/'DR'/None).

    Returns (None, None) for blanks and placeholders like '-'.
    """
    if not text:
        return None, None
    t = text.replace("₹", "").replace(" ", "")
    m = NUMBER_RE.search(t)
    if not m:
        return None, None
    value = float(m.group().replace(",", ""))
    before, after = t[:m.start()], t[m.end():]
    if "-" in before or ("(" in before and ")" in after and not re.search(r"cr|dr", after, re.I)):
        value = -value
    ind = re.search(r"(cr|dr)", after, re.I) or re.search(r"(cr|dr)", before, re.I)
    return value, (ind.group(1).upper() if ind else None)


# ---------------------------------------------------------- page geometry

@dataclass
class Word:
    text: str
    x0: float
    x1: float
    top: float
    bottom: float

    @property
    def mid(self):
        return (self.x0 + self.x1) / 2


@dataclass
class Line:
    words: list
    top: float
    bottom: float

    @property
    def text(self):
        return " ".join(w.text for w in self.words)


@dataclass
class Column:
    role: str
    label: str
    x0: float
    x1: float

    @property
    def mid(self):
        return (self.x0 + self.x1) / 2


def _merge_amount_suffix(words):
    """Glue '12,000.00' + 'Cr' (or '(Cr)') into one word so they land in the same column."""
    out = []
    for w in words:
        if (out and DRCR_WORD_RE.match(w.text) and AMOUNT_WORD_RE.match(out[-1].text)
                and w.x0 - out[-1].x1 < (w.bottom - w.top) * 1.2):
            prev = out[-1]
            out[-1] = Word(prev.text + " " + w.text, prev.x0, w.x1, prev.top, prev.bottom)
        else:
            out.append(w)
    return out


def _group_lines(raw_words):
    words = [Word(w["text"], w["x0"], w["x1"], w["top"], w["bottom"]) for w in raw_words]
    words.sort(key=lambda w: (w.top, w.x0))
    lines = []
    for w in words:
        mid = (w.top + w.bottom) / 2
        for line in reversed(lines[-3:]):
            h = line.bottom - line.top
            if abs(mid - (line.top + line.bottom) / 2) <= max(h, w.bottom - w.top) * 0.45:
                line.words.append(w)
                line.top, line.bottom = min(line.top, w.top), max(line.bottom, w.bottom)
                break
        else:
            lines.append(Line([w], w.top, w.bottom))
    for line in lines:
        line.words = _merge_amount_suffix(sorted(line.words, key=lambda w: w.x0))
    lines.sort(key=lambda l: l.top)
    return lines


# ------------------------------------------------------------ header finding

def _group(words):
    words = sorted(words, key=lambda w: (round(w.top), w.x0))
    return {
        "x0": min(w.x0 for w in words), "x1": max(w.x1 for w in words),
        "h": max(w.bottom - w.top for w in words),
        "text": " ".join(w.text for w in words),
        "key": "".join(_norm(w.text) for w in words if _norm(w.text) not in ("inr", "rs")),
    }


def _stack_groups(lines):
    """Header tokens, left to right. A label wrapped over several lines
    ("Withdrawal" above "Amount (INR)") becomes a single token."""
    phrases = []  # (line index, words closer together than half a letter height)
    for li, line in enumerate(lines):
        cur = None
        for w in line.words:
            if cur and w.x0 - cur[-1].x1 < (w.bottom - w.top) * 0.5:
                cur.append(w)
            else:
                cur = [w]
                phrases.append((li, cur))
    parent = list(range(len(phrases)))

    def root(i):
        while parent[i] != i:
            i = parent[i]
        return i

    for a in range(len(phrases)):
        for b in range(a + 1, len(phrases)):
            (la, wa), (lb, wb) = phrases[a], phrases[b]
            if la != lb and min(wa[-1].x1, wb[-1].x1) > max(wa[0].x0, wb[0].x0):
                parent[root(b)] = root(a)
    members = {}
    for i in range(len(phrases)):
        members.setdefault(root(i), []).append(phrases[i][1])
    groups = []
    for parts in members.values():
        if len(parts) == 1:  # not stacked: one token per word
            groups += [_group([w]) for w in parts[0]]
        else:
            groups.append(_group([w for p in parts for w in p]))
    groups.sort(key=lambda g: g["x0"])
    return groups


def _match_columns(groups):
    """Turn header word groups into columns by matching known header phrases."""
    cols, i = [], 0
    while i < len(groups):
        best = None
        for j in range(min(len(groups), i + MAX_KEY_TOKENS), i, -1):
            if any(groups[k + 1]["x0"] - groups[k]["x1"] > groups[k]["h"] * 1.2 for k in range(i, j - 1)):
                continue
            key = "".join(g["key"] for g in groups[i:j])
            if key in KEY_TO_ROLE:
                best = (j, KEY_TO_ROLE[key])
                break
        if best:
            j, role = best
        else:
            j, role = i + 1, "other"
            # keep unknown adjacent words together ("Sl No", "Init. Br")
            while (j < len(groups) and groups[j]["x0"] - groups[j - 1]["x1"] < groups[j]["h"] * 0.6
                   and not any("".join(g["key"] for g in groups[j:k]) in KEY_TO_ROLE
                               for k in range(j + 1, min(len(groups), j + MAX_KEY_TOKENS) + 1))):
                j += 1
        label = " ".join(g["text"] for g in groups[i:j])
        if role == "other" and _norm(label) in IGNORED_HEADER_TOKENS:
            if cols:  # stray "(INR)" under/after a label: widen the previous column
                cols[-1].x1 = max(cols[-1].x1, groups[j - 1]["x1"])
            i = j
            continue
        cols.append(Column(role, label, groups[i]["x0"], groups[j - 1]["x1"]))
        i = j
    # a second "date" column is the value/posting date; duplicates of other roles are extras
    seen = set()
    for c in cols:
        if c.role == "date" and "date" in seen and "value_date" not in seen:
            c.role = "value_date"
        elif c.role in seen and c.role != "other":
            c.role = "other"
        seen.add(c.role)
    return cols


def _header_score(cols):
    roles = {c.role for c in cols}
    if not roles & {"date", "value_date"}:
        return 0
    money = roles & {"debit", "credit", "amount", "balance"}
    if not money or not ({"debit", "credit"} <= roles or "amount" in roles or "balance" in roles):
        return 0
    known = len(roles - {"other"})
    if known < 3 or len([c for c in cols if c.role == "other"]) > known:
        return 0
    return known


def _find_header(lines):
    """Return (columns, index of first line after the header) or None."""
    for i, line in enumerate(lines[:80]):
        best = None
        for n in (1, 2, 3):
            chunk = lines[i:i + n]
            if len(chunk) < n:
                break
            h = max(w.bottom - w.top for w in chunk[0].words)
            if n > 1 and chunk[-1].top - chunk[-2].bottom > h * 1.2:
                break
            cols = _match_columns(_stack_groups(chunk))
            score = _header_score(cols)
            if score and (best is None or score > best[0]):
                best = (score, cols, i + n)
        if best:
            return best[1], best[2]
    return None


# ------------------------------------------------------------ column edges

def _naive_boundaries(cols):
    bounds = []
    for a, b in zip(cols, cols[1:]):
        if a.role in NUMERIC_ROLES or b.role in NUMERIC_ROLES:
            bounds.append((a.x1 + b.x0) / 2)
        else:
            bounds.append(b.x0 - 1)
    return bounds


def _river_boundaries(cols, lines):
    """Put each column edge in the widest vertical gap of whitespace between the two headers."""
    naive = _naive_boundaries(cols)
    if not lines:
        return naive
    spans = sorted((w.x0, w.x1) for l in lines for w in l.words)
    gaps, reach = [], spans[0][1]
    for x0, x1 in spans[1:]:
        if x0 > reach:
            gaps.append((reach, x0))
        reach = max(reach, x1)
    bounds = []
    for k, (a, b) in enumerate(zip(cols, cols[1:])):
        lo = a.x0 if a.role in NUMERIC_ROLES else a.mid
        hi = b.x1 if b.role not in NUMERIC_ROLES else b.mid
        lo = max(lo, a.x0 + 1)
        inside = [(max(g0, lo), min(g1, hi)) for g0, g1 in gaps if min(g1, hi) - max(g0, lo) >= 1.5]
        if inside:
            g0, g1 = max(inside, key=lambda g: g[1] - g[0])
            bounds.append((g0 + g1) / 2)
        else:
            bounds.append(naive[k])
    return bounds


def _cells(line, cols, bounds):
    cells = {}
    for w in line.words:
        k = sum(w.mid > b for b in bounds)
        cells.setdefault(k, []).append(w.text)
    return {k: " ".join(v) for k, v in cells.items()}


# --------------------------------------------------------------- page rows

OPENING_RE = re.compile(r"^(opening\s*balance|balance\s*(b/?f|brought\s*forward)|b/?f\b|brought\s*forward)", re.I)
CARRIED_RE = re.compile(r"^(balance\s*)?(c/?f|carried\s*forward)\b", re.I)
STOP_RE = re.compile(
    r"^[*\s]*(closing\s*balance|statement\s*summary|summary\b|total\b|grand\s*total|transaction\s*total"
    r"|end\s*of\s*(the\s*)?statement|\**\s*end\s*of|this\s*is\s*a\s*(computer|system)"
    r"|computer\s*generated|dr\s*count|cr\s*count|abbreviations|legends?\b|important\s*(note|information))",
    re.I)
SKIP_RE = re.compile(r"^(page\s*(no\.?\s*)?:?\s*\d+(\s*(of|/)\s*\d+)?|continued|contd)", re.I)


@dataclass
class Row:
    cells: dict  # role -> text (extras keyed "other:<label>")
    top: float
    bottom: float
    date: dt.date | None
    text: str


def _to_rows(lines, cols, bounds):
    rows = []
    for line in lines:
        by_col = _cells(line, cols, bounds)
        cells = {}
        for k, text in by_col.items():
            col = cols[k]
            key = col.role if col.role != "other" else "other:" + col.label
            cells[key] = (cells.get(key, "") + " " + text).strip()
        date = parse_date(cells.get("date", "")) if "date" in cells else None
        if date is None and not any(c.role == "date" for c in cols) and "value_date" in cells:
            date = parse_date(cells["value_date"])
        rows.append(Row(cells, line.top, line.bottom, date, line.text))
    return rows


def _table_lines(lines, h):
    """Cut the lines below the header where the transaction table ends."""
    kept, last_bottom = [], None
    for line in lines:
        text = line.text.strip()
        if SKIP_RE.match(text):
            continue
        first_is_date = parse_date(text) is not None
        if not first_is_date and STOP_RE.match(text):
            break
        if (not first_is_date and last_bottom is not None and kept
                and line.top - last_bottom > h * 4):
            break
        kept.append(line)
        last_bottom = line.bottom
    return kept


def _horizontal_rules(page, x0, x1):
    width = x1 - x0
    ys = []
    for e in page.edges:
        if e.get("orientation") == "h" and (e["x1"] - e["x0"]) >= 0.4 * width:
            ys.append(e["top"])
    ys.sort()
    out = []
    for y in ys:
        if not out or y - out[-1] > 1:
            out.append(y)
    return out


def _valid_blocks(blocks, strict=False):
    """At most one date per block. Undated blocks (other than a page's first, which
    continues the previous page) may not carry amounts; with `strict` they must be
    opening / carried-forward balance rows."""
    if not blocks:
        return False
    for i, b in enumerate(blocks):
        n = sum(r.date is not None for r in b)
        if n > 1:
            return False
        if n == 0 and i > 0:
            text = " ".join(r.cells.get("description", r.text) for r in b)
            if OPENING_RE.match(text) or CARRIED_RE.match(text):
                continue
            if strict or any(k in NUMERIC_ROLES for r in b for k in r.cells):
                return False
    return True


def _split(rows, cuts):
    blocks, cur, ci = [], [], 0
    for r in rows:
        mid = (r.top + r.bottom) / 2
        while ci < len(cuts) and cuts[ci] < mid:
            if cur:
                blocks.append(cur)
                cur = []
            ci += 1
        cur.append(r)
    if cur:
        blocks.append(cur)
    return blocks


def _blocks(rows, page, cols):
    """Group table lines into one block per transaction."""
    if not rows:
        return []
    # 1. ruled tables: one block between two horizontal rules
    rules = _horizontal_rules(page, cols[0].x0, cols[-1].x1)
    if len(rules) >= 3:
        blocks = _split(rows, rules)
        if _valid_blocks(blocks) and len(blocks) > 1:
            return blocks
    # 2. extra spacing between transactions (handles dates printed mid-cell)
    diffs = [b.top - a.top for a, b in zip(rows, rows[1:])]
    if len(diffs) >= 2:
        s = sorted(set(round(d, 1) for d in diffs))
        jumps = [(s[i + 1] / s[i], (s[i] + s[i + 1]) / 2) for i in range(len(s) - 1) if s[i] > 0]
        for ratio, thr in sorted(jumps, reverse=True):
            if ratio <= 1.3:
                break
            cuts = [(a.bottom + b.top) / 2 for a, b, d in zip(rows, rows[1:], diffs) if d > thr]
            blocks = _split(rows, cuts)
            if _valid_blocks(blocks, strict=True):
                return blocks
    # 3. default: every dated line starts a transaction, undated lines continue it
    blocks, cur = [], []
    for r in rows:
        if r.date is not None and cur:
            blocks.append(cur)
            cur = []
        cur.append(r)
    blocks.append(cur)
    return blocks


def _merge_block(block):
    cells = {}
    for r in block:
        for k, v in r.cells.items():
            cells[k] = (cells.get(k, "") + " " + v).strip()
    date = next((r.date for r in block if r.date), None)
    return date, cells


# ------------------------------------------------------------ metadata

IFSC_BANKS = {
    "HDFC": "HDFC Bank", "ICIC": "ICICI Bank", "SBIN": "State Bank of India", "UTIB": "Axis Bank",
    "KKBK": "Kotak Mahindra Bank", "YESB": "Yes Bank", "INDB": "IndusInd Bank",
    "PUNB": "Punjab National Bank", "BARB": "Bank of Baroda", "CNRB": "Canara Bank",
    "UBIN": "Union Bank of India", "IDFB": "IDFC FIRST Bank", "FDRL": "Federal Bank",
    "BKID": "Bank of India", "CBIN": "Central Bank of India", "IDIB": "Indian Bank",
    "IBKL": "IDBI Bank", "AUBL": "AU Small Finance Bank", "RATN": "RBL Bank",
    "IOBA": "Indian Overseas Bank", "UCBA": "UCO Bank", "MAHB": "Bank of Maharashtra",
    "PSIB": "Punjab & Sind Bank", "KARB": "Karnataka Bank", "KVBL": "Karur Vysya Bank",
    "SIBL": "South Indian Bank", "TMBL": "Tamilnad Mercantile Bank", "CSBK": "CSB Bank",
    "DBSS": "DBS Bank", "SCBL": "Standard Chartered", "HSBC": "HSBC", "CITI": "Citibank",
    "ESFB": "Equitas Small Finance Bank", "UJVN": "Ujjivan Small Finance Bank",
    "BDBL": "Bandhan Bank", "DLXB": "Dhanlaxmi Bank", "JAKA": "Jammu & Kashmir Bank",
}
BANK_NAMES = [(re.compile(p, re.I), n) for p, n in [
    (r"hdfc\s*bank", "HDFC Bank"), (r"icici\s*bank", "ICICI Bank"),
    (r"state\s*bank\s*of\s*india|\bsbi\b", "State Bank of India"), (r"axis\s*bank", "Axis Bank"),
    (r"kotak", "Kotak Mahindra Bank"), (r"yes\s*bank", "Yes Bank"), (r"indusind", "IndusInd Bank"),
    (r"punjab\s*national\s*bank|\bpnb\b", "Punjab National Bank"), (r"bank\s*of\s*baroda", "Bank of Baroda"),
    (r"canara\s*bank", "Canara Bank"), (r"union\s*bank", "Union Bank of India"),
    (r"idfc\s*first", "IDFC FIRST Bank"), (r"federal\s*bank", "Federal Bank"),
    (r"bank\s*of\s*india", "Bank of India"), (r"idbi", "IDBI Bank"), (r"rbl\s*bank", "RBL Bank"),
]]
ACCOUNT_RE = re.compile(
    r"(?:a/?c|account)\s*(?:no|number|num|#)?\.?\s*[:\-]?\s*([0-9Xx*]{6,20})\b", re.I)
PERIOD_RE = re.compile(
    r"(?:from|period)\s*(?:date)?\s*[:\-]?\s*(\d{1,2}[/\-.\s][\w]{2,9}[/\-.\s]\d{2,4})"
    r"\s*(?:to|till|-)\s*(?:date)?\s*[:\-]?\s*(\d{1,2}[/\-.\s][\w]{2,9}[/\-.\s]\d{2,4})", re.I)


def _metadata(stmt, text):
    m = re.search(r"\b([A-Z]{4})0[A-Z0-9]{6}\b", text)
    if m and m.group(1) in IFSC_BANKS:
        stmt.bank = IFSC_BANKS[m.group(1)]
    else:
        for rx, name in BANK_NAMES:
            if rx.search(text):
                stmt.bank = name
                break
    m = ACCOUNT_RE.search(text)
    if m:
        stmt.account_no = m.group(1)
    m = PERIOD_RE.search(text)
    if m:
        a, b = parse_date(m.group(1)), parse_date(m.group(2))
        if a and b:
            stmt.period = f"{a:%d-%m-%Y} to {b:%d-%m-%Y}"


# ------------------------------------------------------------ transactions

def _to_transaction(date, cells, page):
    t = Transaction(date=date, page=page)
    t.description = cells.get("description", "")
    t.ref = cells.get("ref", "")
    if "value_date" in cells:
        t.value_date = parse_date(cells["value_date"])
    t.extra = {k[6:]: v for k, v in cells.items() if k.startswith("other:")}

    dr, _ = parse_amount(cells.get("debit", ""))
    cr, _ = parse_amount(cells.get("credit", ""))
    t.debit = abs(dr) if dr else None
    t.credit = abs(cr) if cr else None
    if "amount" in cells:
        amt, ind = parse_amount(cells["amount"])
        flag = (cells.get("drcr", "") or ind or "").strip().upper()
        if amt:
            if flag.startswith("D") or (not flag and amt < 0):
                t.debit = abs(amt)
            elif flag.startswith("C") or (not flag and amt > 0 and "drcr" in cells):
                t.credit = abs(amt)
            else:
                t.extra["_amount"] = amt  # direction decided from the balance later
    bal, ind = parse_amount(cells.get("balance", ""))
    if bal is not None:
        t.balance = -abs(bal) if ind == "DR" else bal
    return t


def _count_mismatches(txns, opening):
    prev, bad = opening, 0
    for t in txns:
        if t.balance is None:
            continue
        if prev is not None and abs(prev - (t.debit or 0) + (t.credit or 0) - t.balance) > 0.005:
            bad += 1
        prev = t.balance
    return bad


def _finish(stmt):
    txns = stmt.transactions
    if not txns:
        return
    # Some banks list newest first: keep whichever order the running balance agrees with.
    if len(txns) > 1:
        rev = txns[::-1]
        fwd_bad = _count_mismatches(txns, None)
        rev_bad = _count_mismatches(rev, None)
        dates_desc = txns[0].date > txns[-1].date
        if rev_bad < fwd_bad or (rev_bad == fwd_bad and dates_desc):
            stmt.transactions = txns = rev
            stmt.warnings.append("Statement was listed newest-first; rows were put in date order.")

    # Amount column without Dr/Cr marker: direction from the balance movement.
    prev = stmt.opening_balance
    for t in txns:
        amt = t.extra.pop("_amount", None)
        if amt is not None:
            if prev is not None and t.balance is not None:
                if abs(prev + amt - t.balance) < 0.005:
                    t.credit = abs(amt)
                else:
                    t.debit = abs(amt)
            else:
                t.credit = amt if amt > 0 else None
                t.debit = -amt if amt < 0 else None
        if t.balance is not None:
            prev = t.balance

    if stmt.opening_balance is None:
        first = next((t for t in txns if t.balance is not None), None)
        if first is not None:
            idx = txns.index(first)
            stmt.opening_balance = round(
                first.balance - sum((t.credit or 0) - (t.debit or 0) for t in txns[:idx + 1]), 2)
    last = next((t for t in reversed(txns) if t.balance is not None), None)
    stmt.closing_balance = last.balance if last else None

    prev = stmt.opening_balance
    for t in txns:
        if t.balance is None:
            continue
        ok = prev is None or abs(prev - (t.debit or 0) + (t.credit or 0) - t.balance) <= 0.005
        t.check = "OK" if ok else "Mismatch"
        prev = t.balance
    if stmt.mismatches:
        stmt.warnings.append(
            f"{stmt.mismatches} row(s) do not agree with the running balance - check them against the PDF.")

    extras = []
    for t in txns:
        for k in t.extra:
            if k not in extras:
                extras.append(k)
    stmt.extra_columns = extras


def parse_statement(path, password=None, source=None):
    """Read a bank statement PDF and return a Statement."""
    stmt = Statement(source=source or str(path))
    try:
        pdf = pdfplumber.open(path, password=password or "")
    except Exception as e:  # pdfplumber wraps pdfminer's PDFPasswordIncorrect
        name = type(e).__name__
        causes = " ".join(type(x).__name__ for x in (e, *e.args, e.__cause__, e.__context__))
        if "Password" in causes or "Encryption" in causes:
            raise StatementError(
                "This PDF is password protected. Enter the password (often the customer ID, "
                "date of birth or the first letters of the name + DOB, as the bank's e-mail says).")
        raise StatementError(f"Could not open the PDF ({name}: {e}).")

    with pdf:
        cols = bounds = None
        any_text = False
        pending_open = True  # opening balance row allowed until the first transaction
        for pno, page in enumerate(pdf.pages, 1):
            words = page.extract_words(x_tolerance=1.5, y_tolerance=2, keep_blank_chars=False)
            if not words:
                continue
            any_text = True
            if pno == 1:
                _metadata(stmt, page.extract_text() or "")
            lines = _group_lines(words)
            hdr = _find_header(lines)
            heights = [w.bottom - w.top for l in lines for w in l.words]
            h = statistics.median(heights) if heights else 8
            if hdr:
                cols, start = hdr
                body = lines[start:]
            elif cols:
                # header not repeated on this page: start at the first dated line,
                # plus any narration lines just above it
                rows0 = _to_rows(lines, cols, bounds)
                first = next((i for i, r in enumerate(rows0) if r.date), None)
                if first is None:
                    continue
                start = first
                while start > 0 and lines[start].top - lines[start - 1].bottom < h * 1.2 \
                        and not rows0[start - 1].date:
                    start -= 1
                body = lines[start:]
            else:
                continue

            table = _table_lines(body, h)
            if hdr:
                bounds = _river_boundaries(cols, table)
            rows = _to_rows(table, cols, bounds)

            # re-cut the table with proper dates now that columns are known
            kept, last_bottom = [], None
            for r in rows:
                if r.date is None and last_bottom is not None and (
                        r.top - last_bottom > h * 4
                        # prose running across the date column: a footer, not a narration
                        or len(re.findall(r"[A-Za-z]{2,}", r.cells.get("date", ""))) >= 2):
                    break
                kept.append(r)
                last_bottom = r.bottom
            rows = kept

            for block in _blocks(rows, page, cols):
                date, cells = _merge_block(block)
                text = " ".join(r.text for r in block).strip()
                desc = cells.get("description", text)
                if OPENING_RE.match(desc) or OPENING_RE.match(text):
                    bal, ind = parse_amount(cells.get("balance", ""))
                    if bal is None:
                        nums = NUMBER_RE.findall(text)
                        bal = float(nums[-1].replace(",", "")) if nums else None
                        ind = "DR" if re.search(r"\bdr\b", text, re.I) else None
                    if bal is not None and pending_open and not stmt.transactions:
                        stmt.opening_balance = -abs(bal) if ind == "DR" else bal
                    continue
                if CARRIED_RE.match(desc) or CARRIED_RE.match(text):
                    continue
                if date is None:
                    # narration that ran over from the previous page / row
                    if stmt.transactions and not any(k in NUMERIC_ROLES for k in cells):
                        t = stmt.transactions[-1]
                        extra_desc = cells.get("description", "")
                        if extra_desc:
                            t.description = (t.description + " " + extra_desc).strip()
                        if cells.get("ref"):
                            t.ref = (t.ref + " " + cells["ref"]).strip()
                    continue
                t = _to_transaction(date, cells, pno)
                if t.debit is None and t.credit is None and "_amount" not in t.extra \
                        and t.balance is None:
                    stmt.warnings.append(f"Page {pno}: row dated {date:%d-%m-%Y} has no amount "
                                         f"(\"{t.description[:40]}\"); kept as is.")
                stmt.transactions.append(t)
                pending_open = False

        if not any_text:
            raise StatementError(
                "No text found in this PDF - it looks like a scanned image. Download the "
                "statement from net banking as a PDF (not a scan), or run it through OCR first.")
        if cols is None:
            raise StatementError(
                "Could not find the transaction table (a header row with Date / Narration / "
                "Debit / Credit / Balance). This layout is not supported yet.")

    _finish(stmt)
    return stmt
