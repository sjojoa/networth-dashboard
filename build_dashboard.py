#!/usr/bin/env python3
"""Build a self-contained net-worth dashboard from an .xlsx tracker.

Reads the workbook, works out each account's owner, category and whether it is
still in use, and injects the result into template.html to produce
NetWorth_Dashboard.html (no network needed).

Nothing about the spreadsheet is hard-coded beyond its layout: owners come from
the legend in row 1, and a new account column is categorised automatically (see
category_for). Uses only the standard library, so it runs with any python3 —
including Apple's /usr/bin/python3, which the double-click launcher relies on.

Usage:
    python3 build_dashboard.py                 # always rebuild
    python3 build_dashboard.py --if-needed     # rebuild only if inputs changed
    python3 build_dashboard.py path/to/book.xlsx
"""

import datetime
import json
import re
import sys
import time
from pathlib import Path

import xlsx_reader

HERE = Path(__file__).resolve().parent
WORKBOOK_NAME = "NetWorthTracker.xlsx"
TEMPLATE = HERE / "template.html"
MODEL = HERE / "model.js"   # the same rules in JavaScript, so the page can redo them after an edit
OUTPUT = HERE / "NetWorth_Dashboard.html"

LEGEND_ROW = 1
CATEGORY_ROW = 3   # optional: type a category here to override the auto-guess
HEADER_ROW = 4
FIRST_DATA_ROW = 5

ZERO = 0.005       # dollars below this count as "no balance"
CLOSED_RUN = 3     # consecutive $0 snapshots at the end before an account is closed,
                   # or — never funded — since its first typed entry before it's hidden

# Explicit category overrides, for names no keyword would catch. Editing this is
# optional — it is the second of four resolution steps in category_for().
CATEGORY_OVERRIDES = {
    # "Grandma's gift fund": "Cash",
}

# Checked in order, so "Roth brokerage" lands in Retirement rather than Taxable.
CATEGORY_HINTS = [
    ("Retirement", ("401k", "401(k)", "403b", "403(b)", "457", " ira", "ira ", "roth",
                    "pension", "retirement", "tsp", "sep", "annuity")),
    ("Bonds", ("bond", "treasury", "t-bill", "tbill", "cd ladder", "gilt")),
    ("Taxable investing", ("brokerage", "taxable", "invest", "stocks", "etf", "crypto")),
    ("Cash", ("cash", "checking", "chequing", "savings", "money market", "hysa",
              "venmo", "paypal", "wallet")),
    ("Property", ("home", "house", "condo", "property", "real estate", "vehicle", "car")),
]

LIABILITY_HINTS = ("credit card", "card", "loan", "mortgage", "debt", "liability",
                   "heloc", "student", "owed", "payable")

# Categories render in this order; anything new is appended, liabilities last.
CANONICAL_ORDER = ["Cash", "Bonds", "Retirement", "Taxable investing", "Property", "Other assets"]
LIABILITY_CATEGORY = "Liabilities"


def find_workbook(args):
    """Explicit path wins; otherwise look beside this script, then one level up."""
    if args:
        return Path(args[0]).expanduser()
    for candidate in (HERE / WORKBOOK_NAME, HERE.parent / WORKBOOK_NAME):
        if candidate.exists():
            return candidate
    matches = sorted(HERE.glob("*.xlsx")) + sorted(HERE.parent.glob("*.xlsx"))
    return matches[0] if matches else HERE / WORKBOOK_NAME


def legend_owners(sheet):
    """Row 1 maps a fill colour to an owner name. Order here fixes owner colours,
    so an owner with no accounts yet still holds its slot."""
    names, by_fill = [], {}
    for cell in sheet.row(LEGEND_ROW):
        if isinstance(cell.value, str) and cell.value.strip() and cell.fill:
            name = cell.value.strip()
            if name.lower().rstrip(":") == "legend":
                continue
            if name not in names:
                names.append(name)
            by_fill[cell.fill] = name
    return names, by_fill


def normalize(text):
    """Lowercase, punctuation to spaces, padded — so hints can match whole words."""
    return " " + re.sub(r"[^a-z0-9]+", " ", text.lower()).strip() + " "


def mentions(name, hint):
    """Match a hint as a whole word (or its plural). Hints of 5+ characters also
    match as a substring, so 'invest' still catches 'Investments' — while short
    ones like 'car' and 'sep' can't fire on 'Carter' or 'September'."""
    padded, h = normalize(name), normalize(hint).strip()
    return (" " + h + " ") in padded or (" " + h + "s ") in padded or (len(h) >= 5 and h in padded)


def is_liability(name, series, explicit_category):
    if explicit_category:
        return explicit_category.strip().lower() in ("liability", "liabilities", "debt")
    if any(mentions(name, h) for h in LIABILITY_HINTS):
        return True
    # Fall back to the sign: a column that is never positive is money owed.
    values = [v for v in series if abs(v) > ZERO]
    return bool(values) and all(v < 0 for v in values)


def category_for(name, series, explicit_category):
    """1) the category row, 2) an override, 3) a keyword, 4) sign / unknown."""
    if explicit_category:
        text = explicit_category.strip()
        if text.lower() in ("liability", "liabilities", "debt"):
            return LIABILITY_CATEGORY
        return text
    if is_liability(name, series, None):
        return LIABILITY_CATEGORY
    if name in CATEGORY_OVERRIDES:
        return CATEGORY_OVERRIDES[name]
    for category, hints in CATEGORY_HINTS:
        if any(mentions(name, h) for h in hints):
            return category
    return "Other assets"


def is_blank(value):
    return value is None or (isinstance(value, str) and not value.strip())


def carry_forward(raw):
    """A blank cell means "not updated this time", not $0 — so people can update
    on different schedules (one weekly, one monthly) on the same sheet. Blanks
    take the account's last entered balance; blanks before its first entry are
    $0, since the account didn't exist yet. A typed 0 is still a real $0.
    Returns (series, carried) where carried lists the row indices filled in."""
    series, carried, last = [], [], None
    for i, v in enumerate(raw):
        if is_blank(v):
            if last is not None:
                carried.append(i)
            series.append(last if last is not None else 0.0)
        else:
            last = round(float(v), 2)
            series.append(last)
    return series, carried


def closed_state(series, first_entry):
    """An account is closed once it has been $0 for the last few snapshots after
    having held a balance — a rolled-over 401k, say. One that has never held a
    balance is hidden the same way, once CLOSED_RUN snapshots have passed since
    its first typed entry: the clock starts there rather than at row 1, so a
    column added to a sheet with years of history isn't hidden on day one, and
    an all-blank column (nothing typed yet) is left showing.
    Returns (hidden, last index with a balance, never funded)."""
    held = [i for i, v in enumerate(series) if abs(v) > ZERO]
    if not held:
        unfunded = first_entry is not None and len(series) - first_entry >= CLOSED_RUN
        return unfunded, None, unfunded
    tail = series[-CLOSED_RUN:] if len(series) >= CLOSED_RUN else series
    if all(abs(v) <= ZERO for v in tail):
        return True, held[-1], False
    return False, None, False


def order_categories(accounts):
    present = []
    for a in accounts:
        if a["category"] not in present:
            present.append(a["category"])
    liability_cats = [c for c in present if any(x["category"] == c and x["isLiability"] for x in accounts)]
    asset_cats = [c for c in present if c not in liability_cats]
    ordered = [c for c in CANONICAL_ORDER if c in asset_cats]
    ordered += [c for c in asset_cats if c not in ordered]
    ordered += [c for c in (LIABILITY_CATEGORY,) if c in liability_cats]
    ordered += [c for c in liability_cats if c not in ordered]
    return ordered


# Debt types file under Liabilities; asset types are categories themselves (as in model.js).
DEBT_TYPES = ("Credit card", "Loan", "Mortgage")
ACCOUNTS_SHEET = "Accounts"
ACCOUNT_HEADINGS = {
    "name": r"^(account|account name|name)$", "type": r"^(type|account type)$",
    "institution": r"^(institution|bank)$", "last4": r"^last ?4( digits)?$", "notes": r"^notes?$",
    "status": r"^status$", "creditLimit": r"^credit limit$", "interestRate": r"^(interest rate|apr)( ?\(%\))?$",
}
DETAIL_FIELDS = ("type", "institution", "last4", "notes", "status", "creditLimit", "interestRate")


def category_for_type(account_type):
    return LIABILITY_CATEGORY if account_type in DEBT_TYPES else account_type


def account_details(sheets):
    """The optional "Accounts" sheet: one row per account (matched by name) with
    its type, institution, last 4 digits, notes, Open/Closed status, credit limit
    and interest rate. Written by the dashboard's Save to workbook, or by hand."""
    sheet = next((sh for sh in sheets[1:] if sh.name.strip().lower() == ACCOUNTS_SHEET.lower()), None)
    if sheet is None:
        return {}
    head = [str(c.value).strip().lower() if c.value is not None else "" for c in sheet.row(1)]
    col = {f: next((i + 1 for i, h in enumerate(head) if re.match(rx, h)), None) for f, rx in ACCOUNT_HEADINGS.items()}
    if not col["name"]:
        return {}
    out = {}
    for r in range(2, sheet.max_row + 1):
        name = sheet.cell(r, col["name"]).value
        if not isinstance(name, str) or not name.strip():
            continue
        d = {}
        for f in DETAIL_FIELDS:
            v = sheet.cell(r, col[f]).value if col[f] else None
            if isinstance(v, str):
                v = v.strip() or None
            if f == "last4" and v is not None:
                v = str(int(v)).zfill(4) if isinstance(v, (int, float)) else str(v)
            elif f in ("creditLimit", "interestRate") and v is not None:
                try:
                    v = float(str(v).replace(",", "").replace("$", "").replace("%", ""))
                except ValueError:
                    v = None
            elif f == "status" and v is not None:
                v = str(v).lower() if str(v).lower() in ("open", "closed") else None
            elif v is not None and not isinstance(v, str):
                v = str(v)
            d[f] = v
        out[name.strip().lower()] = d
    return out


def sheet_table(sheets, name):
    """A sheet after the first, by name: [{heading: value}] for each row with something in it."""
    sheet = next((sh for sh in sheets[1:] if sh.name.strip().lower() == name.lower()), None)
    if sheet is None:
        return None
    head = [str(c.value).strip().lower() if c.value is not None else "" for c in sheet.row(1)]
    out = []
    for r in range(2, sheet.max_row + 1):
        row = {h: sheet.cell(r, k + 1).value for k, h in enumerate(head) if h}
        if any(v is not None and v != "" for v in row.values()):
            row["_row"] = r
            out.append(row)
    return out


def iso_date(v):
    if isinstance(v, (datetime.datetime, datetime.date)):
        return v.strftime("%Y-%m-%d")
    if isinstance(v, str) and re.match(r"^\d{4}-\d{2}-\d{2}$", v.strip()):
        return v.strip()
    return None


def statements_and_transactions(sheets, account_id):
    """The optional Statements and Transactions sheets, written by the Transactions
    tab's Save to workbook: uploaded statements, and one row per transaction
    (amounts in dollars here, kept as cents by the page)."""
    st_rows, tx_rows = sheet_table(sheets, "Statements"), sheet_table(sheets, "Transactions")
    if st_rows is None and tx_rows is None:
        return {"statements": [], "transactions": []}
    statements = []
    for r in st_rows or []:
        sid, acct = r.get("statement"), r.get("account")
        if not isinstance(sid, (int, float)) or not isinstance(acct, str) or acct.strip().lower() not in account_id:
            continue
        num = lambda v: float(v) if isinstance(v, (int, float)) else None
        statements.append({
            "id": int(sid), "account_id": account_id[acct.strip().lower()], "account_name": acct.strip(),
            "source": str(r.get("file") or "Statement"), "format": str(r.get("format") or "csv").lower(),
            "first_date": iso_date(r.get("first date")), "last_date": iso_date(r.get("last date")),
            "ending_balance": num(r.get("ending balance")), "ending_date": iso_date(r.get("ending date")),
            "uploaded_at": str(r.get("uploaded") or "") or iso_date(r.get("first date")) or "1970-01-01", "uploaded_by": None,
        })
    known = {st["id"] for st in statements}
    transactions = []
    for r in tx_rows or []:
        date, amount, sid = iso_date(r.get("date")), r.get("amount"), r.get("statement")
        if date is None or not isinstance(amount, (int, float)) or not isinstance(sid, (int, float)) or int(sid) not in known:
            continue
        kind = str(r.get("kind") or "").strip().lower()
        transactions.append({
            "id": r["_row"], "statement_id": int(sid), "date": date,
            "description": str(r.get("description") or "(no description)"), "amount": int(round(amount * 100)),
            "category": (str(r.get("category")).strip() or None) if r.get("category") is not None else None,
            "kind": kind if kind in ("transfer", "normal") else None,
        })
    return {"statements": statements, "transactions": transactions}


def col_letter(n):
    s = ""
    while n:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


def extract(xlsx_path):
    sheets = xlsx_reader.load(xlsx_path)
    sheet = sheets[0]
    owner_names, owner_by_fill = legend_owners(sheet)
    details = account_details(sheets)

    columns = []
    for col, cell in enumerate(sheet.row(HEADER_ROW), start=1):
        name = cell.value.strip() if isinstance(cell.value, str) else None
        if name and name.lower() == "total":
            break          # anything right of Total is notes, not an account
        if col == 1 or not name:
            continue
        raw_category = sheet.cell(CATEGORY_ROW, col).value
        d = details.get(name.lower(), {})
        typed = raw_category.strip() if isinstance(raw_category, str) and raw_category.strip() else None
        columns.append({
            "name": name,
            "col": col,
            "owner": owner_by_fill.get(cell.fill),
            # Row 3 wins; otherwise the type from the Accounts sheet picks the category.
            "rawCategory": typed or (category_for_type(d["type"]) if d.get("type") else None),
            "typed": typed,
            "status": d.get("status"),
            "details": {f: d.get(f) for f in DETAIL_FIELDS},
        })
    total_col = next((c for c, cell in enumerate(sheet.row(HEADER_ROW), start=1)
                      if isinstance(cell.value, str) and cell.value.strip().lower() == "total"), None)

    found = []
    for row in range(FIRST_DATA_ROW, sheet.max_row + 1):
        raw_date = sheet.cell(row, 1).value
        if not isinstance(raw_date, (datetime.datetime, datetime.date)):
            continue
        values = [sheet.cell(row, c["col"]).value for c in columns]
        if all(is_blank(v) or v == 0 for v in values):
            continue  # empty template row
        values = [None if is_blank(v) or not isinstance(v, (int, float)) else v for v in values]
        found.append((raw_date.strftime("%Y-%m-%d"), row, values))
    # Date order, as the hosted version keeps it, so a row typed out of order still lands in place.
    found.sort(key=lambda f: f[0])
    # Budgets, goals and statements find accounts by name, and edits find rows by date: twins would mix up.
    seen = {}
    for c in columns:
        if c["name"].lower() in seen:
            raise SystemExit(f'Two accounts are called "{c["name"]}" (columns {col_letter(seen[c["name"].lower()])} '
                             f'and {col_letter(c["col"])}). Give each a different name.')
        seen[c["name"].lower()] = c["col"]
    for (d1, r1, _), (d2, r2, _) in zip(found, found[1:]):
        if d1 == d2:
            raise SystemExit(f"Rows {r1} and {r2} both have the date {d1}. Merge them into one row.")
    dates = [f[0] for f in found]
    rows = [f[2] for f in found]
    sheet_rows = [f[1] for f in found]

    if not dates:
        raise SystemExit(f"No dated rows found in {xlsx_path}")

    accounts = []
    for i, c in enumerate(columns):
        raw = [r[i] for r in rows]
        series, carried = carry_forward(raw)
        first_entry = next((k for k, v in enumerate(raw) if not is_blank(v)), None)
        closed, last_held, unfunded = closed_state(series, first_entry)
        auto = ("unfunded" if unfunded else "closed") if closed else None
        # An Open/Closed status (the Accounts sheet) wins over the rules, as in model.js.
        if c["status"] == "open":
            closed, unfunded, last_held = False, False, None
        elif c["status"] == "closed":
            held = [k for k, v in enumerate(series) if abs(v) > ZERO]
            closed, unfunded, last_held = True, False, (held[-1] if held else None)
        accounts.append({
            "name": c["name"],
            "owner": c["owner"],
            "category": category_for(c["name"], series, c["rawCategory"]),
            "isLiability": is_liability(c["name"], series, c["rawCategory"]),
            "closed": closed,
            "unfunded": unfunded,
            "auto": auto,
            "lastHeld": dates[last_held] if last_held is not None else None,
            "series": series,
            "carried": carried,
        })

    owners = [{"name": n, "accounts": sum(1 for a in accounts if a["owner"] == n)} for n in owner_names]
    owner_id = {n: 1 + i for i, n in enumerate(owner_names)}

    return {
        "source": Path(xlsx_path).name,
        "generated": datetime.datetime.now().strftime("%Y-%m-%d %H:%M"),
        "dates": dates,
        "accounts": accounts,
        "owners": owners,
        "categoryOrder": order_categories(accounts),
        "closedRun": CLOSED_RUN,
        # Local editing: what the page needs to redo the rules after an edit (tracker),
        # ids for the editing UI (edit), and where each value sits for Save to workbook (sheet).
        "local": True,
        "builtAt": int(time.time() * 1000),
        "tracker": {
            "owners": owner_names,
            "accounts": [{"name": c["name"], "owner": c["owner"], "rawCategory": c["rawCategory"], "status": c["status"]} for c in columns],
            "dates": dates,
            "rows": rows,
        },
        "edit": {
            # Ids are where things sit in the sheet (row of a snapshot, column of an account),
            # so an edit kept in the browser still points at the same cell after a rebuild.
            "snapshotIds": sheet_rows,
            "accounts": [{"id": c["col"], "ownerId": owner_id.get(c["owner"]), "rawCategory": c["rawCategory"], **c["details"]}
                         for i, c in enumerate(columns)],
            "owners": [{"id": owner_id[n], "name": n} for n in owner_names],
        },
        "txStore": statements_and_transactions(sheets, {c["name"].lower(): c["col"] for c in columns}),
        "sheet": {
            "name": sheet.name,
            "headerRow": HEADER_ROW,
            "rows": sheet_rows,
            "cols": [c["col"] for c in columns],
            "categories": [c["typed"] for c in columns],
            "totalCol": total_col,
            "lastRow": sheet.max_row,
        },
    }


def is_stale(xlsx_path):
    """True if the dashboard is missing or older than the workbook/template."""
    if not OUTPUT.exists():
        return True
    built = OUTPUT.stat().st_mtime
    return any(p.exists() and p.stat().st_mtime > built for p in (xlsx_path, TEMPLATE, MODEL, Path(__file__)))


def build(xlsx_path):
    data = extract(xlsx_path)
    # model.js is an ES module; the page gets it as a plain script, inside a function so its names stay private.
    model = re.sub(r"^export ", "", MODEL.read_text(encoding="utf-8"), flags=re.M)
    html = TEMPLATE.read_text(encoding="utf-8").replace(
        "/*__DATA__*/null", json.dumps(data, separators=(",", ":"))
    ).replace("/*__MODEL__*/", model)
    OUTPUT.write_text(html, encoding="utf-8")
    return data


def main():
    argv = sys.argv[1:]
    if_needed = "--if-needed" in argv
    xlsx_path = find_workbook([a for a in argv if not a.startswith("--")])

    if not xlsx_path.exists():
        raise SystemExit(f"Workbook not found: {xlsx_path}")

    if if_needed and not is_stale(xlsx_path):
        print(f"Up to date: {OUTPUT.name} (workbook unchanged since last build)")
        return

    data = build(xlsx_path)
    latest = sum(a["series"][-1] for a in data["accounts"])
    closed = [a["name"] for a in data["accounts"] if a["closed"] and not a["unfunded"]]
    unfunded = [a["name"] for a in data["accounts"] if a["unfunded"]]
    print(f"Wrote {OUTPUT}")
    print(f"  {len(data['dates'])} snapshots, {len(data['accounts'])} accounts")
    print(f"  owners: " + ", ".join(f"{o['name']} ({o['accounts']})" for o in data["owners"]))
    print(f"  categories: " + ", ".join(data["categoryOrder"]))
    if closed:
        print(f"  closed: " + ", ".join(closed))
    if unfunded:
        print(f"  hidden, never funded: " + ", ".join(unfunded))
    stale = [a["name"] for a in data["accounts"] if a["carried"] and a["carried"][-1] == len(data["dates"]) - 1]
    if stale:
        print(f"  carried forward to the latest snapshot: " + ", ".join(stale))
    print(f"  latest ({data['dates'][-1]}): ${latest:,.2f}")


if __name__ == "__main__":
    main()
