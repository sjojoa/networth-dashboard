#!/usr/bin/env python3
"""Generate a demo tracker workbook full of invented data.

Writes NetWorthTracker.xlsx in the same layout the dashboard expects:

    row 1   legend — one coloured cell per owner
    row 3   category per account (optional; the dashboard guesses without it)
    row 4   headers — Date, one column per account, then Total
    row 5+  one row per weekly snapshot

Everything here is fabricated by a seeded random walk, so it's reproducible and
resembles real accounts without being anyone's. Standard library only.

Usage:  python3 make_demo_workbook.py [out.xlsx] [options]

    --weeks N        how many weekly snapshots (default 105, about two years)
    --seed N         change the random draw (default 7)
    --scale N        multiply every balance, e.g. --scale 3 for a bigger portfolio
    --max-total N    fail if net worth would ever exceed N (off by default)

Nothing here limits what the dashboard can display — these only shape the
fabricated sample data.
"""

import datetime
import math
import random
import sys
import zipfile

EPOCH = datetime.date(1899, 12, 30)
OWNERS = ["User1", "User2", "Joint"]
# Light fills, one per owner: the dashboard reads ownership from these.
OWNER_FILL = ["FFC6E0B4", "FFF8CBAD", "FFBDD7EE"]


def serial(d):
    return (d - EPOCH).days


def col_letter(n):
    out = ""
    while n:
        n, r = divmod(n - 1, 26)
        out = chr(65 + r) + out
    return out


class Sim:
    """Weekly series generators. All amounts in dollars, rounded to the cent."""

    def __init__(self, weeks, rng):
        self.weeks = weeks
        self.rng = rng

    def market(self, start, annual_return=0.08, annual_vol=0.14, weekly_add=0.0):
        """Geometric random walk with a steady contribution — retirement/brokerage."""
        dt = 1.0 / 52.0
        mu, sig, v, out = annual_return, annual_vol, start, []
        for _ in range(self.weeks):
            shock = self.rng.gauss(0, 1)
            v *= math.exp((mu - 0.5 * sig * sig) * dt + sig * math.sqrt(dt) * shock)
            v += weekly_add
            out.append(round(max(v, 0), 2))
        return out

    def cash(self, start, weekly_pay, weekly_spend, pay_every=2, jitter=0.22):
        """Sawtooth: pay lands every other week, spending drains it in between."""
        v, out = start, []
        for w in range(self.weeks):
            if w % pay_every == 0:
                v += weekly_pay * pay_every * (1 + self.rng.gauss(0, 0.05))
            v -= weekly_spend * (1 + self.rng.gauss(0, jitter))
            out.append(round(max(v, 40), 2))
        return out

    def savings(self, start, apy=0.042, weekly_add=0.0, raid_chance=0.01):
        """Interest-bearing, occasionally raided for something big."""
        v, out = start, []
        for _ in range(self.weeks):
            v *= (1 + apy / 52.0)
            v += weekly_add
            if self.rng.random() < raid_chance:
                v -= v * self.rng.uniform(0.15, 0.4)
            out.append(round(max(v, 0), 2))
        return out

    def ladder(self, start, apy=0.045, top_up_every=13, top_up=500.0):
        """Bond/CD ladder: accrues, with a rung added now and then."""
        v, out = start, []
        for w in range(self.weeks):
            v *= (1 + apy / 52.0)
            if w and w % top_up_every == 0:
                v += top_up
            out.append(round(v, 2))
        return out

    def card(self, typical, spread=0.6, payoff_chance=0.12):
        """Revolving balance, held negative. Occasionally paid to zero."""
        out = []
        bal = typical
        for _ in range(self.weeks):
            bal = max(0.0, bal * (1 + self.rng.gauss(0, spread * 0.35))
                      + self.rng.gauss(typical * 0.25, typical * 0.4))
            if self.rng.random() < payoff_chance:
                bal = self.rng.uniform(0, typical * 0.1)
            out.append(-round(min(bal, typical * 4), 2))
        return out

    def loan(self, principal, weekly_payment, apr=0.062):
        """Amortising, held negative, and never past zero."""
        bal, out = principal, []
        for _ in range(self.weeks):
            bal = bal * (1 + apr / 52.0) - weekly_payment
            bal = max(bal, 0.0)
            out.append(-round(bal, 2))
        return out


def build_accounts(weeks, seed, scale=1.0):
    """Every generator here is linear in its dollar inputs, so scaling the
    finished series is the same as scaling each starting balance and
    contribution — and it keeps a given seed reproducible."""
    rng = random.Random(seed)
    s = Sim(weeks, rng)

    # name, owner, category, series
    acc = [
        ("Checking",         "User1", "Cash",              s.cash(3100, 1450, 1395)),
        ("Savings",          "User1", "Cash",              s.savings(6400, weekly_add=35)),
        ("HYSA",             "User2", "Cash",              s.savings(8900, apy=0.048, weekly_add=45)),
        ("Joint checking",   "Joint", "Cash",              s.cash(3800, 640, 625)),
        ("Joint HYSA",       "Joint", "Cash",              s.savings(10500, apy=0.047, weekly_add=60)),
        ("Bonds",            "User1", "Bonds",             s.ladder(4800, top_up=260)),
        ("Roth IRA",         "User1", "Retirement",        s.market(13800, weekly_add=48)),
        ("401(k)",           "User1", "Retirement",        s.market(19500, weekly_add=125)),
        ("Old 401(k)",       "User1", "Retirement",        s.market(11200, weekly_add=0)),
        ("Traditional IRA",  "User2", "Retirement",        s.market(8600, weekly_add=20)),
        ("401(k) 2",         "User2", "Retirement",        s.market(7900, weekly_add=90)),
        ("Brokerage",        "User1", "Taxable investing", s.market(6300, annual_vol=0.19, weekly_add=42)),
        ("Joint brokerage",  "Joint", "Taxable investing", s.market(5900, annual_vol=0.17, weekly_add=55)),
        ("Credit card",      "User1", "Liabilities",       s.card(1050)),
        ("Credit card 2",    "User2", "Liabilities",       s.card(560)),
        ("Car loan",         "Joint", "Liabilities",       s.loan(13800, 88)),
    ]

    # Roll the old 401(k) into the traditional IRA about two thirds through, so the
    # dashboard's "closed account" handling has something to show.
    by_name = {a[0]: a for a in acc}
    old, trad = by_name["Old 401(k)"], by_name["Traditional IRA"]
    at = int(weeks * 0.62)
    moved = old[3][at - 1]
    for i in range(at, weeks):
        old[3][i] = 0.0
        trad[3][i] = round(trad[3][i] + moved * (1 + 0.0012) ** (i - at), 2)

    if scale != 1.0:
        for a in acc:
            a[3][:] = [round(v * scale, 2) for v in a[3]]
    return acc


def write_xlsx(path, dates, accounts):
    strings, sidx = [], {}

    def sid(text):
        if text not in sidx:
            sidx[text] = len(strings)
            strings.append(text)
        return sidx[text]

    STYLE_DATE = 1
    STYLE_OWNER = {name: 2 + i for i, name in enumerate(OWNERS)}
    total_col = len(accounts) + 2

    def cell(ref, value=None, style=0, shared=False, formula=None):
        a = ' s="%d"' % style if style else ""
        if formula is not None:
            return '<c r="%s"%s><f>%s</f><v>%s</v></c>' % (ref, a, formula, value)
        if value is None:
            return '<c r="%s"%s/>' % (ref, a)
        if shared:
            return '<c r="%s"%s t="s"><v>%d</v></c>' % (ref, a, value)
        return '<c r="%s"%s><v>%s</v></c>' % (ref, a, value)

    rows = []

    # row 1 — legend
    cs = [cell("A1", sid("Legend:"), shared=True)]
    for i, name in enumerate(OWNERS):
        cs.append(cell("%s1" % col_letter(2 + i), sid(name), STYLE_OWNER[name], shared=True))
    rows.append("<row r=\"1\">%s</row>" % "".join(cs))

    # row 3 — categories
    cs = [cell("A3", sid("Category"), shared=True)]
    for j, (_, _, category, _) in enumerate(accounts):
        cs.append(cell("%s3" % col_letter(2 + j), sid(category), shared=True))
    rows.append("<row r=\"3\">%s</row>" % "".join(cs))

    # row 4 — headers, filled with the owner's colour
    cs = [cell("A4", sid("Date"), shared=True)]
    for j, (name, owner, _, _) in enumerate(accounts):
        cs.append(cell("%s4" % col_letter(2 + j), sid(name), STYLE_OWNER[owner], shared=True))
    cs.append(cell("%s4" % col_letter(total_col), sid("Total"), shared=True))
    rows.append("<row r=\"4\">%s</row>" % "".join(cs))

    # data rows
    for i, d in enumerate(dates):
        r = 5 + i
        cs = [cell("A%d" % r, serial(d), STYLE_DATE)]
        for j, (_, _, _, series) in enumerate(accounts):
            cs.append(cell("%s%d" % (col_letter(2 + j), r), "%.2f" % series[i]))
        total = sum(a[3][i] for a in accounts)
        cs.append(cell("%s%d" % (col_letter(total_col), r), "%.2f" % total,
                       formula="SUM(B%d:%s%d)" % (r, col_letter(total_col - 1), r)))
        rows.append('<row r="%d">%s</row>' % (r, "".join(cs)))

    fills = "".join('<fill><patternFill patternType="solid"><fgColor rgb="%s"/>'
                    '<bgColor indexed="64"/></patternFill></fill>' % c for c in OWNER_FILL)
    styles = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
        '<numFmts count="1"><numFmt numFmtId="164" formatCode="yyyy\\-mm\\-dd"/></numFmts>'
        '<fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>'
        '<fills count="%d"><fill><patternFill patternType="none"/></fill>'
        '<fill><patternFill patternType="gray125"/></fill>%s</fills>'
        '<borders count="1"><border/></borders>'
        '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
        '<cellXfs count="%d">'
        '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>'
        '<xf numFmtId="164" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>'
        '%s</cellXfs>'
        # Excel expects the built-in "Normal" style to be declared.
        '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>'
        '<dxfs count="0"/>'
        '</styleSheet>'
        % (2 + len(OWNER_FILL), fills, 2 + len(OWNER_FILL),
           "".join('<xf numFmtId="0" fontId="0" fillId="%d" borderId="0" xfId="0" applyFill="1"/>'
                   % (2 + i) for i in range(len(OWNER_FILL))))
    )

    esc = lambda t: t.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
    refs = sum(r.count('t="s"') for r in rows)
    sst = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
           '<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
           'count="%d" uniqueCount="%d">%s</sst>'
           % (refs, len(strings),
              "".join("<si><t>%s</t></si>" % esc(t) for t in strings)))

    widths = ('<cols><col min="1" max="1" width="11.5" customWidth="1"/>'
              '<col min="2" max="%d" width="15.5" customWidth="1"/></cols>' % total_col)
    sheet = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
             '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
             '<dimension ref="A1:%s%d"/><sheetViews><sheetView workbookViewId="0">'
             '<pane ySplit="4" topLeftCell="A5" activePane="bottomLeft" state="frozen"/>'
             '</sheetView></sheetViews>%s<sheetData>%s</sheetData></worksheet>'
             % (col_letter(total_col), 4 + len(dates), widths, "".join(rows)))

    parts = {
        "[Content_Types].xml":
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
            '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
            '<Default Extension="xml" ContentType="application/xml"/>'
            '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
            '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
            '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>'
            '<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>'
            '</Types>',
        "_rels/.rels":
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>'
            '</Relationships>',
        "xl/workbook.xml":
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
            'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
            '<sheets><sheet name="Net worth" sheetId="1" r:id="rId1"/></sheets></workbook>',
        "xl/_rels/workbook.xml.rels":
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>'
            '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
            '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>'
            '</Relationships>',
        "xl/worksheets/sheet1.xml": sheet,
        "xl/styles.xml": styles,
        "xl/sharedStrings.xml": sst,
    }
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        for name, body in parts.items():
            z.writestr(name, body)


def main():
    args = [a for a in sys.argv[1:] if not a.startswith("--")]
    out = args[0] if args else "NetWorthTracker.xlsx"
    opt = lambda name, cast, default: (
        cast(sys.argv[sys.argv.index(name) + 1]) if name in sys.argv else default)
    weeks = opt("--weeks", int, 105)
    seed = opt("--seed", int, 7)
    scale = opt("--scale", float, 1.0)
    max_total = opt("--max-total", float, None)   # no ceiling unless you ask

    end = datetime.date(2026, 8, 21)
    dates = [end - datetime.timedelta(days=7 * (weeks - 1 - i)) for i in range(weeks)]
    accounts = build_accounts(weeks, seed, scale)
    write_xlsx(out, dates, accounts)

    total = sum(a[3][-1] for a in accounts)
    peak = max(sum(a[3][i] for a in accounts) for i in range(weeks))
    if max_total is not None and peak > max_total:
        raise SystemExit("Demo net worth peaks at $%s, over the $%s you asked for — "
                         "try a smaller --scale."
                         % (format(peak, ",.0f"), format(max_total, ",.0f")))
    print("Wrote %s" % out)
    print("  %d weekly snapshots, %s to %s" % (weeks, dates[0], dates[-1]))
    print("  %d accounts, owners: %s" % (len(accounts), ", ".join(OWNERS)))
    print("  net worth: $%s final, $%s peak%s"
          % (format(total, ",.2f"), format(peak, ",.2f"),
             "" if scale == 1.0 else " (scale x%g)" % scale))


if __name__ == "__main__":
    main()
