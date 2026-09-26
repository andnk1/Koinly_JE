"""
koinly_engine.py
=================
Web-app engine for the Koinly -> QuickBooks JE tool. This is the same
ledger/JE math as the desktop koinly_to_je.py script (kept in sync by hand --
see that file's docstring for the full accounting rules), adapted to work on
uploaded file bytes instead of scanning a folder, and with no interactive
prompts / no sys.exit -- everything is returned as data so app.py can render
it or hand it to the xlsx/csv writers.

Nothing here writes to disk. All inputs come in as (filename, bytes) pairs
and all outputs are Python data structures (or, for CSV, plain strings).
"""
import csv
import io
from collections import defaultdict
from datetime import datetime, timedelta

DIGITAL_ASSETS = "Digital Asset"
CAPITAL_GAIN_LOSS = "Capital Gain/Loss"
CASH = "Cash"
ROUND = 2

# Accounts that are already real GL account names in their own right -- these
# are never looked up in the mapping file, just passed straight through into
# the GL Account column.
FIXED_ACCOUNTS = {DIGITAL_ASSETS, CAPITAL_GAIN_LOSS, CASH}

# Optional mapping-file rows that override the GL account used for the two
# built-in accounts. If the row is missing or its Account Name is blank, the
# built-in name ("Digital Asset" / "Capital Gain/Loss") is used as before.
DEFAULT_ACCOUNT_LABELS = {
    DIGITAL_ASSETS: "default digital asset account",
    CAPITAL_GAIN_LOSS: "default capital gain/loss account",
}

# Only these blank-Tag transaction types are listed in the "Blank-Tag
# transactions" review box, and only when their net value is > 0. Buy/Sell
# (same as exchanges), Fiat_Deposit/Fiat_Withdrawal and zero-value rows are
# left out of the review list (they are still posted in the JE as before).
BLANK_TAG_REVIEW_TYPES = {"crypto_deposit", "crypto_withdrawal"}

# Order in which category groups appear within a period; each group is then
# sorted largest-dollar-amount first.
TYPE_ORDER = [
    "crypto_deposit", "crypto_withdrawal", "exchange", "transfer",
    "buy", "sell", "fiat_deposit", "fiat_withdrawal",
]

JE_HEADER = ["JournalDate", "JournalNo", "AccountName", "GL Account",
             "Debit", "Credit", "Memo", "Name"]

QB_HEADER = ["*JournalNo", "*JournalDate", "*AccountName", "*Debits", "*Credits",
             "Description", "Name", "Currency", "Location", "Class"]


class EngineError(ValueError):
    """Raised for problems with the uploaded files themselves (bad columns,
    unreadable format, etc.) -- app.py catches this and shows it as a plain
    user-facing error rather than a stack trace."""
    pass


# ---------------------------------------------------------------------------
# number / CSV helpers
# ---------------------------------------------------------------------------

def clean_num(v):
    if v is None:
        return 0.0
    v = str(v).strip().replace(",", "").replace("$", "")
    if v == "":
        return 0.0
    try:
        return float(v)
    except ValueError:
        return 0.0


def fmt_amount(v):
    return f"{v:.2f}" if v else ""


def month_key(date_str):
    dt = datetime.strptime(date_str[:10], "%Y-%m-%d")
    return dt.strftime("%Y-%m")


def month_end(year, month):
    if month == 12:
        nxt = datetime(year + 1, 1, 1)
    else:
        nxt = datetime(year, month + 1, 1)
    return (nxt - timedelta(days=1)).strftime("%m/%d/%Y")


# ---------------------------------------------------------------------------
# Loading uploaded files
# ---------------------------------------------------------------------------

def load_koinly_csv_bytes(filename, content_bytes):
    """Koinly exports have a title row + blank row before the real header
    ('Date,...'). Returns a list of dict rows (DictReader-style)."""
    text = content_bytes.decode("utf-8-sig", errors="replace")
    lines = text.splitlines(keepends=True)
    header_idx = None
    for i, line in enumerate(lines):
        if line.startswith("Date,"):
            header_idx = i
            break
    if header_idx is None:
        raise EngineError(
            f"'{filename}' doesn't look like a Koinly transaction-history export "
            "(no 'Date,...' header row found). Make sure you uploaded the raw "
            "Koinly CSV, not a re-saved or edited copy."
        )
    reader = csv.DictReader(lines[header_idx:])
    rows = list(reader)
    required = {"Date", "Type"}
    if rows and not required.issubset(set(rows[0].keys())):
        missing = required - set(rows[0].keys())
        raise EngineError(
            f"'{filename}' is missing expected column(s): {', '.join(sorted(missing))}."
        )
    for r in rows:
        r["_source_file"] = filename
    return rows


def _find_columns(header_cells):
    """Tolerant of header variations like 'QB Account Name' -- looks for a
    cell that IS 'label' and a cell that CONTAINS 'account name'."""
    norm = [str(h).strip().lower() if h is not None else "" for h in header_cells]
    label_i = None
    acct_i = None
    for i, h in enumerate(norm):
        if h == "label" and label_i is None:
            label_i = i
        if "account name" in h and acct_i is None:
            acct_i = i
    if label_i is None or acct_i is None:
        raise EngineError(
            "Mapping file needs a 'Label' column and an 'Account Name' column "
            f"(found columns: {[str(h) for h in header_cells]})"
        )
    return label_i, acct_i


def load_mapping_bytes(filename, content_bytes):
    """Returns dict: normalized (stripped + lowercased) Label -> Account Name
    (stripped, original case), from an uploaded Label/Account Name mapping
    file (.csv, .xlsx or .xlsm)."""
    ext = filename.lower().rsplit(".", 1)[-1] if "." in filename else ""
    mapping = {}

    if ext == "csv":
        text = content_bytes.decode("utf-8-sig", errors="replace")
        reader = csv.reader(io.StringIO(text))
        header = next(reader, None)
        if not header:
            return mapping
        label_i, acct_i = _find_columns(header)
        for row in reader:
            if len(row) <= max(label_i, acct_i):
                continue
            label = (row[label_i] or "").strip()
            acct = (row[acct_i] or "").strip()
            if label:
                mapping[label.lower()] = acct

    elif ext in ("xlsx", "xlsm"):
        import openpyxl
        wb = openpyxl.load_workbook(io.BytesIO(content_bytes), data_only=True)
        ws = wb.worksheets[0]
        rows = list(ws.iter_rows(values_only=True))
        if not rows:
            return mapping
        label_i, acct_i = _find_columns(rows[0])
        for row in rows[1:]:
            if row is None or len(row) <= max(label_i, acct_i):
                continue
            label_v, acct_v = row[label_i], row[acct_i]
            label = str(label_v).strip() if label_v is not None else ""
            acct = str(acct_v).strip() if acct_v is not None else ""
            if label:
                mapping[label.lower()] = acct
    else:
        raise EngineError(f"Unsupported mapping file type: '{filename}' (use .csv or .xlsx)")
    return mapping


def gl_account_for(acct, mapping, unmapped_seen):
    if acct in DEFAULT_ACCOUNT_LABELS:
        override = mapping.get(DEFAULT_ACCOUNT_LABELS[acct], "")
        return override if override else acct
    if acct in FIXED_ACCOUNTS:
        return acct
    gl = mapping.get(acct.strip().lower())
    if gl:
        return gl
    unmapped_seen.add(acct)
    return ""


# ---------------------------------------------------------------------------
# Ledger: accumulate transactions into (month, type, category) groups
# ---------------------------------------------------------------------------

class Ledger:
    def __init__(self):
        self.groups = defaultdict(lambda: defaultdict(float))
        self.review = defaultdict(lambda: defaultdict(float))

    def add(self, month, typ, category, **amounts):
        g = self.groups[(month, typ, category)]
        for k, v in amounts.items():
            g[k] += v

    def flag(self, month, txn_type, tag_label, count=1, amount=0.0):
        key = (month, txn_type, tag_label)
        r = self.review[key]
        r["count"] += count
        r["amount"] += amount


def process_rows(rows, ledger, unmapped_types_seen):
    skipped_no_date = 0
    for r in rows:
        date = (r.get("Date") or "").strip()
        if not date:
            skipped_no_date += 1
            continue
        m = month_key(date)
        typ = (r.get("Type") or "").strip().lower()
        tag = (r.get("Tag") or "").strip()

        sent_cb = clean_num(r.get("Sent Cost Basis"))
        recv_cb = clean_num(r.get("Received Cost Basis"))
        sent_amt = clean_num(r.get("Sent Amount"))
        recv_amt = clean_num(r.get("Received Amount"))
        net_value = clean_num(r.get("Net Value (USD)"))

        if typ == "exchange":
            cat = "Exchange"
        elif typ == "transfer":
            cat = "Transfer"
        else:
            cat = tag if tag else (typ.title() if typ else "(blank type)")

        ledger.add(m, typ, cat, sent_cb=sent_cb, recv_cb=recv_cb,
                   net_value=net_value, sent_amt=sent_amt, recv_amt=recv_amt,
                   count=1)

        if not tag and typ in BLANK_TAG_REVIEW_TYPES and round(net_value, ROUND) > 0:
            ledger.flag(m, typ, cat, amount=net_value)

        if typ in ("buy", "sell", "fiat_deposit", "fiat_withdrawal"):
            unmapped_types_seen.add(typ)
        elif typ not in ("crypto_deposit", "crypto_withdrawal", "transfer", "exchange"):
            unmapped_types_seen.add(f"UNKNOWN TYPE: {typ}" if typ else "UNKNOWN TYPE: (blank)")
    return skipped_no_date


# ---------------------------------------------------------------------------
# Turning groups into balanced JE lines
# ---------------------------------------------------------------------------

def group_lines(typ, category, g, period_label):
    """Return the balanced debit/credit lines for one (period, type, category)
    group, all sharing one memo, in NATURAL order (matching the CLI script /
    QB_JE.csv exactly). Display-only debit-first reordering happens later,
    in reorder_debit_first() -- never here, so QB_JE.csv is unaffected."""
    memo = f"{category} - {period_label}"
    acct_for_category = category
    sent_cb, recv_cb, net_value = g["sent_cb"], g["recv_cb"], g["net_value"]
    sent_amt, recv_amt = g["sent_amt"], g["recv_amt"]
    lines = []

    def plug(diff):
        d = round(diff, ROUND)
        if d > 0:
            lines.append((CAPITAL_GAIN_LOSS, 0.0, d, memo))
        elif d < 0:
            lines.append((CAPITAL_GAIN_LOSS, -d, 0.0, memo))

    if typ == "crypto_deposit":
        lines.append((acct_for_category, 0.0, net_value, memo))
        lines.append((DIGITAL_ASSETS, net_value, 0.0, memo))

    elif typ == "crypto_withdrawal":
        lines.append((acct_for_category, net_value, 0.0, memo))
        lines.append((DIGITAL_ASSETS, 0.0, sent_cb, memo))
        plug(net_value - sent_cb)

    elif typ == "exchange":
        lines.append((DIGITAL_ASSETS, recv_cb, 0.0, memo))
        lines.append((DIGITAL_ASSETS, 0.0, sent_cb, memo))
        plug(recv_cb - sent_cb)

    elif typ == "transfer":
        lines.append((DIGITAL_ASSETS, recv_cb, 0.0, memo))
        lines.append((DIGITAL_ASSETS, 0.0, sent_cb, memo))
        plug(recv_cb - sent_cb)

    elif typ == "buy":
        basis = recv_cb if recv_cb else net_value
        lines.append((DIGITAL_ASSETS, basis, 0.0, memo))
        lines.append((CASH, 0.0, sent_amt, memo))

    elif typ == "sell":
        proceeds = recv_amt if recv_amt else net_value
        lines.append((CASH, proceeds, 0.0, memo))
        lines.append((DIGITAL_ASSETS, 0.0, sent_cb, memo))
        plug(proceeds - sent_cb)

    elif typ == "fiat_deposit":
        amt = recv_amt if recv_amt else net_value
        lines.append((CASH, amt, 0.0, memo))
        lines.append((acct_for_category, 0.0, amt, memo))

    elif typ == "fiat_withdrawal":
        amt = sent_amt if sent_amt else net_value
        lines.append((acct_for_category, amt, 0.0, memo))
        lines.append((CASH, 0.0, amt, memo))

    else:
        lines.append((acct_for_category, net_value, 0.0, memo))

    return lines


def sort_groups(items):
    def sort_key(item):
        typ, cat, g = item
        type_rank = TYPE_ORDER.index(typ) if typ in TYPE_ORDER else len(TYPE_ORDER)
        magnitude = max(abs(g["net_value"]), abs(g["sent_cb"]), abs(g["recv_cb"]))
        return (type_rank, -magnitude)
    return sorted(items, key=sort_key)


def emit_records(items, period_label, je_date, je_no, mapping, unmapped_accounts_seen):
    """Same emission as the CLI script -- natural (non-reordered) line order,
    used for JE_monthly/JE_total data and QB_JE.csv."""
    records = []
    for typ, cat, g in sort_groups(items):
        for acct, debit, credit, memo in group_lines(typ, cat, g, period_label):
            debit, credit = round(debit, ROUND), round(credit, ROUND)
            if not debit and not credit:
                continue
            gl = gl_account_for(acct, mapping, unmapped_accounts_seen)
            records.append({
                "je_date": je_date, "je_no": je_no,
                "acct": acct, "gl": gl,
                "debit": debit, "credit": credit,
                "memo": memo, "name": "",
            })
    return records


def combine_all_groups(ledger):
    combined = defaultdict(lambda: defaultdict(float))
    months = set()
    for (m, typ, cat), g in ledger.groups.items():
        months.add(m)
        dest = combined[(typ, cat)]
        for k, v in g.items():
            dest[k] += v
    return combined, sorted(months)


def records_to_je_rows(records):
    return [[r["je_date"], r["je_no"], r["acct"], r["gl"],
             fmt_amount(r["debit"]), fmt_amount(r["credit"]),
             r["memo"], r["name"]] for r in records]


def build_monthly_records(ledger, mapping, unmapped_accounts_seen):
    by_month = defaultdict(list)
    for (m, typ, cat), g in ledger.groups.items():
        by_month[m].append((typ, cat, g))

    all_records = []
    for m in sorted(by_month):
        year, mon = int(m[:4]), int(m[5:7])
        je_date = month_end(year, mon)
        je_no = f"JE-{m}"
        all_records.extend(
            emit_records(by_month[m], m, je_date, je_no, mapping, unmapped_accounts_seen)
        )
    return all_records


def build_total_records(ledger, mapping, unmapped_accounts_seen):
    combined, months = combine_all_groups(ledger)
    if not months:
        return []
    je_no = "JE-TOTAL"
    year, mon = int(months[-1][:4]), int(months[-1][5:7])
    je_date = month_end(year, mon)
    items = [(typ, cat, g) for (typ, cat), g in combined.items()]
    return emit_records(items, "Total", je_date, je_no, mapping, unmapped_accounts_seen)


def build_qb_je_rows(monthly_records):
    """QB_JE.csv rows -- unchanged mechanics/order from the CLI script. This
    is the file that gets imported into QuickBooks, so it always mirrors
    JE_monthly.csv's natural line order, never the display reordering."""
    rows_out = []
    seen_je_no = set()
    for r in monthly_records:
        first_line = r["je_no"] not in seen_je_no
        seen_je_no.add(r["je_no"])
        acct = r["gl"] if r["gl"] else r["acct"]
        rows_out.append([
            r["je_no"],
            r["je_date"] if first_line else "",
            acct,
            fmt_amount(r["debit"]),
            fmt_amount(r["credit"]),
            r["memo"],
            r["name"],
            "", "", "",
        ])
    return rows_out


def rows_to_csv_string(header, rows):
    buf = io.StringIO()
    w = csv.writer(buf)
    w.writerow(header)
    w.writerows(rows)
    return buf.getvalue()


# ---------------------------------------------------------------------------
# Display-only reordering: debit line(s) first within each JE group
# ---------------------------------------------------------------------------

def reorder_debit_first(records):
    """Groups consecutive records by (je_no, memo) -- each such block is one
    category's JE group -- and, within each block only, moves debit lines
    ahead of credit lines (stable sort, so relative order is otherwise
    unchanged). This never changes which records exist or their totals, only
    the on-screen row order in the 'JE by Month' tab."""
    out = []
    i = 0
    n = len(records)
    while i < n:
        key = (records[i]["je_no"], records[i]["memo"])
        j = i
        block = []
        while j < n and (records[j]["je_no"], records[j]["memo"]) == key:
            block.append(records[j])
            j += 1
        block.sort(key=lambda r: 0 if r["debit"] else 1)
        out.extend(block)
        i = j
    return out


# ---------------------------------------------------------------------------
# Reconciliation
# ---------------------------------------------------------------------------

def digital_asset_period_totals(ledger, mapping, unmapped_accounts_seen):
    combined, _months = combine_all_groups(ledger)
    debit = credit = 0.0
    for (typ, cat), g in combined.items():
        for acct, d, c, memo in group_lines(typ, cat, g, "Total"):
            if acct == DIGITAL_ASSETS:
                debit += d
                credit += c
    return round(debit, 2), round(credit, 2)


def build_reconciliation(ledger, mapping, unmapped_accounts_seen, beginning, ending):
    """beginning/ending may be None (left blank) -- treated as 0.0, never
    blocking the report."""
    debit, credit = digital_asset_period_totals(ledger, mapping, unmapped_accounts_seen)
    beginning = clean_num(beginning) if beginning not in (None, "") else 0.0
    ending = clean_num(ending) if ending not in (None, "") else 0.0
    calculated_ending = round(beginning + debit - credit, 2)
    difference = round(ending - calculated_ending, 2)
    return {
        "beginning": beginning,
        "ending": ending,
        "debit": debit,
        "credit": credit,
        "calculated_ending": calculated_ending,
        "difference": difference,
    }


# ---------------------------------------------------------------------------
# Top-level entry point used by app.py
# ---------------------------------------------------------------------------

def process(files, mapping_file, beginning, ending):
    """files: list of (filename, bytes) for the uploaded Koinly transaction
    history CSV(s). mapping_file: (filename, bytes) or None. beginning/ending:
    strings or None/'' from the form -- optional.

    Returns a dict with everything app.py / xlsx_report.py need. Raises
    EngineError for problems with the uploaded files themselves."""
    if not files:
        raise EngineError("Please upload at least one Koinly transaction history CSV.")

    mapping = {}
    if mapping_file is not None:
        mapping = load_mapping_bytes(mapping_file[0], mapping_file[1])

    ledger = Ledger()
    unmapped_types_seen = set()
    unmapped_accounts_seen = set()
    total_rows = 0
    skipped_no_date = 0
    for filename, content in files:
        rows = load_koinly_csv_bytes(filename, content)
        total_rows += len(rows)
        skipped_no_date += process_rows(rows, ledger, unmapped_types_seen)

    monthly_records = build_monthly_records(ledger, mapping, unmapped_accounts_seen)
    total_records = build_total_records(ledger, mapping, unmapped_accounts_seen)
    qb_rows = build_qb_je_rows(monthly_records)
    reconciliation = build_reconciliation(ledger, mapping, unmapped_accounts_seen, beginning, ending)

    return {
        "monthly_records": monthly_records,
        "monthly_records_display": reorder_debit_first(monthly_records),
        "total_records": total_records,
        "qb_rows": qb_rows,
        "reconciliation": reconciliation,
        "unmapped_types": sorted(unmapped_types_seen),
        "unmapped_accounts": sorted(unmapped_accounts_seen),
        "blank_tag_review": dict(ledger.review),
        "total_transactions": total_rows,
        "skipped_no_date": skipped_no_date,
        "n_files": len(files),
        "had_mapping": mapping_file is not None,
        "mapping_rows_loaded": len(mapping),
    }
