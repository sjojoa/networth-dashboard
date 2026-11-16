#!/usr/bin/env python3
"""Checks for the local dashboard's editing, run in headless Chrome (sample workbook only).

1. Parity: the page's own DATA (model.js, redone in the browser) matches what
   build_dashboard.py wrote, field by field.
2. Edits: snapshots added, changed and deleted, an account's details and an
   Open/Closed status, through the same api() calls the hosted page makes;
   they're kept in this browser and show after a reload.
3. Save to workbook: the sample workbook patched with those edits is read back
   by build_dashboard.py, and must give the edited data (and keep its other cells).

    python3 test/local_check.py
"""
import base64, json, re, shutil, subprocess, sys, tempfile
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))
import build_dashboard  # noqa: E402

CHROME = "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome"
BOOK = ROOT / "NetWorthTracker.xlsx"

PROBE = r"""
<script>
window.__errors = [];
addEventListener("error", (e) => __errors.push(String(e.message)));
addEventListener("unhandledrejection", (e) => __errors.push(String(e.reason && e.reason.message || e.reason)));
</script>
"""

RUN = r"""
<script>
(async () => {
  const out = { errors: __errors };
  const tick = (ms) => new Promise((r) => setTimeout(r, ms));
  try {
    const step = localStorage.getItem("probe-step") || "1";
    if (step === "1") {
      localStorage.clear(); localStorage.setItem("probe-step", "1b");
      location.reload(); return;
    }
    if (step === "1b") {
      out.data = { source: DATA.source, generated: DATA.generated, dates: DATA.dates, accounts: DATA.accounts, owners: DATA.owners,
                   categoryOrder: DATA.categoryOrder, closedRun: DATA.closedRun };
      out.views = [];
      for (const v of ["dash", "accounts", "proj", "entries", "spending", "settings"]) { setView(v); await tick(30); out.views.push(state.view); }
      out.nav = [...document.querySelectorAll(".nav-item")].filter((n) => !n.hidden && !n.closest("[hidden]")).map((n) => n.textContent.trim());
      out.barHidden = $id("localBar").hidden;
      const ids = EDIT.accounts.map((a) => a.id), sid = EDIT.snapshotIds;
      const calls = [
        ["POST", "/api/snapshots", { date: "2026-08-28", balances: [[ids[0], 9000], [ids[1], null], [ids[2], 15000.5]] }],
        ["POST", "/api/snapshots", { date: "2024-08-16", balances: [[ids[0], 4000]] }],
        ["PUT", `/api/snapshots/${sid[sid.length - 1]}`, { date: "2026-08-21", balances: [[ids[0], 8300.25]] }],
        ["DELETE", `/api/snapshots/${sid[3]}`],
        ["PUT", `/api/accounts/${ids[14]}`, { name: "Visa card", type: "Credit card", ownerId: EDIT.accounts[14].ownerId, institution: "Demo Bank",
          last4: "0042", notes: "Paid in full", status: "open", creditLimit: 5000, interestRate: 21.9 }],
        ["PUT", "/api/accounts/status", { changes: [[ids[8], "open"]] }],
      ];
      out.results = [];
      for (const [m, p, b] of calls) { try { out.results.push(await api(m, p, b)); } catch (e) { out.results.push("ERR " + e.message); } }
      out.refused = [];
      for (const [m, p, b] of [["POST", "/api/snapshots", { date: "2026-08-28", balances: [[ids[0], 1]] }],
                               ["POST", "/api/snapshots", { date: "2026-02-30", balances: [[ids[0], 1]] }],
                               ["POST", "/api/snapshots", { date: "2026-09-04", balances: [[ids[0], 0]] }],
                               ["PUT", `/api/accounts/${ids[0]}`, { name: "Total", type: "Cash", ownerId: 1 }],
                               ["POST", "/api/owners", { name: "Zed" }]]) {
        try { await api(m, p, b); out.refused.push("accepted"); } catch (e) { out.refused.push(e.message); }
      }
      localStorage.setItem("probe-step", "2");
      localStorage.setItem("probe-out1", JSON.stringify(out));
      location.reload(); return;
    }
    if (step === "2") {
      const first = JSON.parse(localStorage.getItem("probe-out1"));
      out.first = first;
      out.after = { dates: DATA.dates.length, first: DATA.dates[0], last: DATA.dates[DATA.dates.length - 1],
                    card: DATA.accounts[14].name, cardDetails: EDIT.accounts[14], old401: [DATA.accounts[8].name, DATA.accounts[8].closed, DATA.accounts[8].auto],
                    checkingLast: DATA.accounts[0].series[DATA.accounts.length ? DATA.dates.length - 2 : 0] };
      out.bar = $id("localBar").hidden ? null : $id("localBarText").textContent;
      const bytes = Uint8Array.from(atob(window.__BOOK), (c) => c.charCodeAt(0));
      const patched = await patchWorkbook(bytes);
      let bin = ""; for (let i = 0; i < patched.length; i += 8192) bin += String.fromCharCode(...patched.subarray(i, i + 8192));
      out.patched = btoa(bin);
      // A workbook changed in Excel since the build (a renamed column) is refused.
      try { await patchWorkbook(Uint8Array.from(atob(window.__BOOK2), (c) => c.charCodeAt(0))); out.stale = "accepted"; } catch (e) { out.stale = e.message; }
      out.ops = LOCAL.store.ops;
      out.expect = { dates: DATA.dates, accounts: DATA.accounts.map((a) => ({ name: a.name, series: a.series, closed: a.closed })), edit: EDIT.accounts };
      out.errors = __errors;
      localStorage.clear();
      document.getElementById("probeOut").textContent = JSON.stringify(out);
    }
  } catch (e) { out.errors.push("probe: " + e.message + " " + e.stack); document.getElementById("probeOut").textContent = JSON.stringify(out); }
})();
</script>
<pre id="probeOut"></pre>
"""


def main():
    tmp = Path(tempfile.mkdtemp())
    for f in ("build_dashboard.py", "xlsx_reader.py", "template.html", "model.js", "NetWorthTracker.xlsx"):
        shutil.copy(ROOT / f, tmp / f)
    build = subprocess.run([sys.executable, str(tmp / "build_dashboard.py"), str(tmp / "NetWorthTracker.xlsx")], capture_output=True, text=True)
    if build.returncode:
        sys.exit(build.stderr)
    html = (tmp / "NetWorth_Dashboard.html").read_text(encoding="utf-8")
    expected = json.loads(re.search(r"withLocalEdits\((\{.*?\})\);\n", html, re.S).group(1))
    book64 = base64.b64encode(BOOK.read_bytes()).decode()
    # The same workbook with a column renamed in "Excel", for the stale check.
    import io, zipfile
    buf = io.BytesIO()
    with zipfile.ZipFile(BOOK) as zin, zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == "xl/sharedStrings.xml":
                data = data.replace(b"<t>Checking</t>", b"<t>Everyday checking</t>", 1)
            zout.writestr(item, data)
    book2 = base64.b64encode(buf.getvalue()).decode()
    html = html.replace("<head>", "<head>" + PROBE, 1).replace("</body>", f"<script>window.__BOOK = {json.dumps(book64)}; window.__BOOK2 = {json.dumps(book2)};</script>" + RUN + "</body>", 1)
    page = tmp / "probe.html"
    page.write_text(html, encoding="utf-8")
    profile = tmp / "profile"
    dom = subprocess.run([CHROME, "--headless=new", "--disable-gpu", f"--user-data-dir={profile}", "--allow-file-access-from-files",
                          "--window-size=1280,900", "--virtual-time-budget=60000", "--dump-dom", page.as_uri()],
                         capture_output=True, text=True, timeout=240).stdout
    m = re.search(r'<pre id="probeOut">(.*?)</pre>', dom, re.S)
    if not m or not m.group(1).strip():
        sys.exit("no result from the page")
    import html as H
    out = json.loads(H.unescape(m.group(1)))
    fails = []
    first = out["first"]
    if first["errors"] or out["errors"]:
        fails.append(("page errors", first["errors"] + out["errors"]))
    # The page adds cssVar and carriedSet to each account as it starts, and type (as the hosted server does).
    for a in first["data"]["accounts"]:
        a.pop("cssVar", None); a.pop("carriedSet", None); a.pop("type", None)
    for k in ("source", "generated", "dates", "accounts", "owners", "categoryOrder", "closedRun"):
        if first["data"][k] != expected[k]:
            detail = "differs"
            if k == "accounts":
                for a, b in zip(first["data"][k], expected[k]):
                    keys = [f for f in set(a) | set(b) if a.get(f) != b.get(f)]
                    if keys:
                        detail = (b["name"], {f: (a.get(f), b.get(f)) for f in keys}); break
            fails.append((f"parity: {k}", detail))
    print("views:", first["views"], "| nav:", first["nav"], "| bar hidden at start:", first["barHidden"])
    print("results:", first["results"])
    print("refused:", first["refused"])
    print("after reload:", {k: v for k, v in out["after"].items() if k != "cardDetails"})
    print("card details:", out["after"]["cardDetails"])
    print("bar:", out["bar"])
    print("stale workbook:", out["stale"])
    if "changed since this dashboard was built" not in out["stale"]:
        fails.append(("stale workbook accepted", out["stale"]))
    # Save to workbook, read back by build_dashboard.py
    saved = tmp / "saved.xlsx"
    saved.write_bytes(base64.b64decode(out["patched"]))
    got = build_dashboard.extract(saved)
    exp = out["expect"]
    if got["dates"] != exp["dates"]:
        fails.append(("saved: dates", [d for d in got["dates"] if d not in exp["dates"]] + ["/"] + [d for d in exp["dates"] if d not in got["dates"]]))
    for a, e in zip(got["accounts"], exp["accounts"]):
        if a["name"] != e["name"] or a["series"] != e["series"] or a["closed"] != e["closed"]:
            fails.append(("saved: account " + e["name"], {"name": a["name"], "closed": a["closed"]}))
    det = {k: got["edit"]["accounts"][14][k] for k in ("type", "institution", "last4", "notes", "status", "creditLimit", "interestRate")}
    print("saved card details:", det)
    if det != {k: exp["edit"][14][k] for k in det}:
        fails.append(("saved: details", det))
    import zipfile
    with zipfile.ZipFile(saved) as z:
        names = z.namelist()
        sheet = z.read("xl/worksheets/sheet1.xml").decode()
    tot = re.findall(r'<c r="R(\d+)"[^>]*><f>([^<]*)</f>', sheet)[-3:]
    print("saved parts:", names)
    print("last Total formulas:", tot)
    # Rebuild from the saved workbook: the kept edits replay as no change, so they're dropped.
    shutil.copy(saved, tmp / "NetWorthTracker.xlsx")
    subprocess.run([sys.executable, str(tmp / "build_dashboard.py"), str(tmp / "NetWorthTracker.xlsx")], check=True, capture_output=True)
    html2 = (tmp / "NetWorth_Dashboard.html").read_text(encoding="utf-8")
    ops = out["ops"]
    check = r"""<pre id="probeOut"></pre><script>
(() => {
  const o = document.getElementById("probeOut");
  if (!localStorage.getItem("ops-set")) {
    localStorage.clear(); localStorage.setItem("ops-set", "1");
    localStorage.setItem("nw-local-edits:" + DATA.source, JSON.stringify({ ops: __OPS, savedAt: 1, nextId: 1000002 }));
    location.reload(); return;
  }
  o.textContent = JSON.stringify({ kept: localStorage.getItem("nw-local-edits:" + DATA.source), bar: document.getElementById("localBar").hidden, n: DATA.dates.length });
})();
</script>""".replace("__OPS", json.dumps(ops))
    (tmp / "probe2.html").write_text(html2.replace("</body>", check + "</body>", 1), encoding="utf-8")
    dom2 = subprocess.run([CHROME, "--headless=new", "--disable-gpu", f"--user-data-dir={tmp / 'profile2'}", "--virtual-time-budget=30000",
                           "--dump-dom", (tmp / "probe2.html").as_uri()], capture_output=True, text=True, timeout=240).stdout
    m2 = re.search(r'<pre id="probeOut">(.*?)</pre>', dom2, re.S)
    after_rebuild = json.loads(H.unescape(m2.group(1))) if m2 and m2.group(1).strip() else None
    if after_rebuild is None:
        print("rebuild probe page:", (dom2 or "")[-600:])
    print("after rebuilding from the saved workbook:", after_rebuild)
    if not after_rebuild or after_rebuild["kept"] is not None or not after_rebuild["bar"]:
        fails.append(("edits not dropped after rebuilding from the saved workbook", after_rebuild))
    if fails:
        print("FAIL"); [print(" ", f) for f in fails]; sys.exit(1)
    print("local checks: OK")


if __name__ == "__main__":
    main()
