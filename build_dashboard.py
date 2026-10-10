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
from pathlib import Path

import xlsx_reader

HERE = Path(__file__).resolve().parent
WORKBOOK_NAME = "NetWorthTracker.xlsx"
TEMPLATE = HERE / "template.html"
OUTPUT = HERE / "NetWorth_Dashboard.html"

LEGEND_ROW = 1
CATEGORY_ROW = 3   # optional: type a category here to override the auto-guess
HEADER_ROW = 4
FIRST_DATA_ROW = 5

ZERO = 0.005       # dollars below this count as "no balance"
CLOSED_RUN = 3     # consecutive $0 snapshots at the end before an account is closed

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


def closed_state(series):
    """An account is closed once it has been $0 for the last few snapshots after
    having held a balance — a rolled-over 401k, say. Returns (closed, last index
    with a balance)."""
    held = [i for i, v in enumerate(series) if abs(v) > ZERO]
    if not held:
        return False, None                    # never used; not "closed"
    tail = series[-CLOSED_RUN:] if len(series) >= CLOSED_RUN else series
    if all(abs(v) <= ZERO for v in tail):
        return True, held[-1]
    return False, None


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


def extract(xlsx_path):
    sheet = xlsx_reader.load(xlsx_path)[0]
    owner_names, owner_by_fill = legend_owners(sheet)

    columns = []
    for col, cell in enumerate(sheet.row(HEADER_ROW), start=1):
        name = cell.value.strip() if isinstance(cell.value, str) else None
        if name and name.lower() == "total":
            break          # anything right of Total is notes, not an account
        if col == 1 or not name:
            continue
        raw_category = sheet.cell(CATEGORY_ROW, col).value
        columns.append({
            "name": name,
            "col": col,
            "owner": owner_by_fill.get(cell.fill),
            "rawCategory": raw_category.strip() if isinstance(raw_category, str) else None,
        })

    dates, rows = [], []
    for row in range(FIRST_DATA_ROW, sheet.max_row + 1):
        raw_date = sheet.cell(row, 1).value
        if not isinstance(raw_date, (datetime.datetime, datetime.date)):
            continue
        values = [sheet.cell(row, c["col"]).value for c in columns]
        if all(is_blank(v) or v == 0 for v in values):
            continue  # empty template row
        dates.append(raw_date.strftime("%Y-%m-%d"))
        rows.append(values)

    if not dates:
        raise SystemExit(f"No dated rows found in {xlsx_path}")

    accounts = []
    for i, c in enumerate(columns):
        series, carried = carry_forward([r[i] for r in rows])
        closed, last_held = closed_state(series)
        accounts.append({
            "name": c["name"],
            "owner": c["owner"],
            "category": category_for(c["name"], series, c["rawCategory"]),
            "isLiability": is_liability(c["name"], series, c["rawCategory"]),
            "closed": closed,
            "lastHeld": dates[last_held] if last_held is not None else None,
            "series": series,
            "carried": carried,
        })

    owners = [{"name": n, "accounts": sum(1 for a in accounts if a["owner"] == n)} for n in owner_names]

    return {
        "source": Path(xlsx_path).name,
        "generated": datetime.datetime.now().strftime("%Y-%m-%d %H:%M"),
        "dates": dates,
        "accounts": accounts,
        "owners": owners,
        "categoryOrder": order_categories(accounts),
    }


def is_stale(xlsx_path):
    """True if the dashboard is missing or older than the workbook/template."""
    if not OUTPUT.exists():
        return True
    built = OUTPUT.stat().st_mtime
    return any(p.exists() and p.stat().st_mtime > built for p in (xlsx_path, TEMPLATE, Path(__file__)))


def build(xlsx_path):
    data = extract(xlsx_path)
    html = TEMPLATE.read_text(encoding="utf-8").replace(
        "/*__DATA__*/null", json.dumps(data, separators=(",", ":"))
    )
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
    closed = [a["name"] for a in data["accounts"] if a["closed"]]
    print(f"Wrote {OUTPUT}")
    print(f"  {len(data['dates'])} snapshots, {len(data['accounts'])} accounts")
    print(f"  owners: " + ", ".join(f"{o['name']} ({o['accounts']})" for o in data["owners"]))
    print(f"  categories: " + ", ".join(data["categoryOrder"]))
    if closed:
        print(f"  closed: " + ", ".join(closed))
    stale = [a["name"] for a in data["accounts"] if a["carried"] and a["carried"][-1] == len(data["dates"]) - 1]
    if stale:
        print(f"  carried forward to the latest snapshot: " + ", ".join(stale))
    print(f"  latest ({data['dates'][-1]}): ${latest:,.2f}")


if __name__ == "__main__":
    main()
