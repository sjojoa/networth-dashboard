/* Turns a stored tracker into the DATA object template.html renders.

   A port of build_dashboard.py's account logic (categories, carry-forward,
   closed and never-funded accounts), so the hosted dashboard shows exactly
   what the local build would for the same workbook. Change both together;
   test/parity.html checks they agree on the sample workbook.

   A tracker is { owners: [name], accounts: [{ name, owner, rawCategory }],
   dates: ["YYYY-MM-DD"], rows: [[number | null]] } — one row per snapshot,
   one value per account, null meaning "not updated this time". */

export const ZERO = 0.005;      // dollars below this count as "no balance"
export const CLOSED_RUN = 3;    // consecutive $0 snapshots at the end before an account is closed,
                                // or — never funded — since its first typed entry before it's hidden

// Explicit category overrides, for names no keyword would catch.
const CATEGORY_OVERRIDES = {
  // "Grandma's gift fund": "Cash",
};

// Checked in order, so "Roth brokerage" lands in Retirement rather than Taxable.
const CATEGORY_HINTS = [
  ["Retirement", ["401k", "401(k)", "403b", "403(b)", "457", " ira", "ira ", "roth",
                  "pension", "retirement", "tsp", "sep", "annuity"]],
  ["Bonds", ["bond", "treasury", "t-bill", "tbill", "cd ladder", "gilt"]],
  ["Taxable investing", ["brokerage", "taxable", "invest", "stocks", "etf", "crypto"]],
  ["Cash", ["cash", "checking", "chequing", "savings", "money market", "hysa",
            "venmo", "paypal", "wallet"]],
  ["Property", ["home", "house", "condo", "property", "real estate", "vehicle", "car"]],
];

const LIABILITY_HINTS = ["credit card", "card", "loan", "mortgage", "debt", "liability",
                         "heloc", "student", "owed", "payable"];

// Categories render in this order; anything new is appended, liabilities last.
const CANONICAL_ORDER = ["Cash", "Bonds", "Retirement", "Taxable investing", "Property", "Other assets"];
const LIABILITY_CATEGORY = "Liabilities";

/* Account types (#18), set in the app. A type decides the category: asset
   types are categories themselves, and debt types all file under Liabilities.
   Accounts without one (most imported ones) get a type inferred for display. */
export const ASSET_TYPES = ["Cash", "Bonds", "Retirement", "Taxable investing", "Property", "Other assets"];
export const DEBT_TYPES = ["Credit card", "Loan", "Mortgage"];
/** Categories a transaction can have (#20). A matched transfer is its own
    kind, not a category, so it's not in this list. */
export const TX_CATEGORIES = [
  "Income", "Groceries", "Dining & coffee", "Home & bills", "Transport", "Health & wellbeing", "Shopping & personal",
  "Subscriptions", "Travel", "Entertainment", "Kids & family", "Gifts & donations", "Fees & interest", "Other",
];
export const categoryForType = (type) => (DEBT_TYPES.includes(type) ? LIABILITY_CATEGORY : type);
/** The type to show for a DATA account: the one set in the app, or a best guess. */
export function typeOf(account, explicit) {
  if (explicit) return explicit;
  if (account.isLiability) return /card/i.test(account.name) ? "Credit card" : /mortgage|heloc/i.test(account.name) ? "Mortgage" : "Loan";
  return account.category;
}
const LIABILITY_WORDS = ["liability", "liabilities", "debt"];

/** Lowercase, punctuation to spaces, padded — so hints can match whole words. */
const normalize = (text) => " " + text.toLowerCase().replace(/[^a-z0-9]+/g, " ").trim() + " ";

/** Match a hint as a whole word (or its plural). Hints of 5+ characters also
    match as a substring, so 'invest' still catches 'Investments' — while short
    ones like 'car' and 'sep' can't fire on 'Carter' or 'September'. */
function mentions(name, hint) {
  const padded = normalize(name), h = normalize(hint).trim();
  return padded.includes(" " + h + " ") || padded.includes(" " + h + "s ") || (h.length >= 5 && padded.includes(h));
}

function isLiability(name, series, explicit) {
  if (explicit) return LIABILITY_WORDS.includes(explicit.trim().toLowerCase());
  if (LIABILITY_HINTS.some((h) => mentions(name, h))) return true;
  // Fall back to the sign: a column that is never positive is money owed.
  const values = series.filter((v) => Math.abs(v) > ZERO);
  return values.length > 0 && values.every((v) => v < 0);
}

/** 1) the category row, 2) an override, 3) a keyword, 4) sign / unknown. */
function categoryFor(name, series, explicit) {
  if (explicit) {
    const text = explicit.trim();
    return LIABILITY_WORDS.includes(text.toLowerCase()) ? LIABILITY_CATEGORY : text;
  }
  if (isLiability(name, series, null)) return LIABILITY_CATEGORY;
  if (Object.prototype.hasOwnProperty.call(CATEGORY_OVERRIDES, name)) return CATEGORY_OVERRIDES[name];
  for (const [category, hints] of CATEGORY_HINTS) {
    if (hints.some((h) => mentions(name, h))) return category;
  }
  return "Other assets";
}

/** Python's round(v, 2): nearest cent, exact ties to even. */
function roundCents(v) {
  const scaled = v * 100, floor = Math.floor(scaled);
  if (scaled - floor === 0.5) return (floor % 2 === 0 ? floor : floor + 1) / 100;
  return Number(v.toFixed(2));
}

/** A blank means "not updated this time", not $0 — so people can update on
    different schedules on the same sheet. Blanks take the account's last
    entered balance; blanks before its first entry are $0. A typed 0 is still
    a real $0. Returns [series, carried] where carried lists the rows filled in. */
function carryForward(raw) {
  const series = [], carried = [];
  let last = null;
  raw.forEach((v, i) => {
    if (v === null) {
      if (last !== null) carried.push(i);
      series.push(last !== null ? last : 0);
    } else {
      last = roundCents(v);
      series.push(last);
    }
  });
  return [series, carried];
}

/** Closed once $0 for the last CLOSED_RUN snapshots after having held a
    balance. Never held one: hidden once CLOSED_RUN snapshots have passed since
    its first typed entry; an all-blank column is left showing.
    Returns [hidden, last index with a balance, never funded]. */
function closedState(series, firstEntry) {
  const held = [];
  series.forEach((v, i) => { if (Math.abs(v) > ZERO) held.push(i); });
  if (!held.length) {
    const unfunded = firstEntry !== null && series.length - firstEntry >= CLOSED_RUN;
    return [unfunded, null, unfunded];
  }
  const tail = series.length >= CLOSED_RUN ? series.slice(-CLOSED_RUN) : series;
  if (tail.every((v) => Math.abs(v) <= ZERO)) return [true, held[held.length - 1], false];
  return [false, null, false];
}

function orderCategories(accounts) {
  const present = [];
  accounts.forEach((a) => { if (!present.includes(a.category)) present.push(a.category); });
  const liab = present.filter((c) => accounts.some((x) => x.category === c && x.isLiability));
  const assets = present.filter((c) => !liab.includes(c));
  const ordered = CANONICAL_ORDER.filter((c) => assets.includes(c));
  assets.forEach((c) => { if (!ordered.includes(c)) ordered.push(c); });
  if (liab.includes(LIABILITY_CATEGORY)) ordered.push(LIABILITY_CATEGORY);
  liab.forEach((c) => { if (!ordered.includes(c)) ordered.push(c); });
  return ordered;
}

/** The DATA object for template.html. meta: { source, generated }. */
export function buildData(tracker, meta) {
  const { dates, rows } = tracker;
  const accounts = tracker.accounts.map((c, i) => {
    const raw = rows.map((r) => r[i]);
    const [series, carried] = carryForward(raw);
    const first = raw.findIndex((v) => v !== null);
    let [closed, lastHeld, unfunded] = closedState(series, first < 0 ? null : first);
    // What the rules alone say, before any status set by hand: shown in Settings so a hidden account can be reopened.
    const auto = closed ? (unfunded ? "unfunded" : "closed") : null;
    // A status set by hand in the app wins over the rules.
    if (c.status === "open") { closed = false; unfunded = false; lastHeld = null; }
    if (c.status === "closed") {
      const held = series.reduce((k, v, i) => (Math.abs(v) > ZERO ? i : k), -1);
      closed = true; unfunded = false; lastHeld = held >= 0 ? held : null;
    }
    return {
      name: c.name,
      owner: c.owner,
      category: categoryFor(c.name, series, c.rawCategory),
      isLiability: isLiability(c.name, series, c.rawCategory),
      closed,
      unfunded,
      auto,
      lastHeld: lastHeld !== null ? dates[lastHeld] : null,
      series,
      carried,
    };
  });
  return {
    source: meta.source,
    generated: meta.generated,
    dates,
    accounts,
    owners: tracker.owners.map((n) => ({ name: n, accounts: accounts.filter((a) => a.owner === n).length })),
    categoryOrder: orderCategories(accounts),
    closedRun: CLOSED_RUN,
  };
}

/** Checks a tracker posted to the import endpoint. Returns an error message, or null. */
export function invalidTracker(t) {
  const isText = (s, max) => typeof s === "string" && s.trim().length > 0 && s.length <= max;
  if (!t || typeof t !== "object") return "Expected a tracker object.";
  const { owners, accounts, dates, rows } = t;
  if (!Array.isArray(owners) || owners.length > 20 || !owners.every((o) => isText(o, 100))) return "Owners must be up to 20 names.";
  if (new Set(owners).size !== owners.length) return "Owner names must be unique.";
  if (!Array.isArray(accounts) || !accounts.length || accounts.length > 200) return "Expected 1–200 accounts.";
  for (const a of accounts) {
    if (!a || !isText(a.name, 200)) return "Every account needs a name.";
    if (a.name.trim().toLowerCase() === "total") return `An account can't be called "Total".`;
    if (a.owner !== null && !owners.includes(a.owner)) return `Account "${a.name}" has an owner that isn't in the legend.`;
    if (a.rawCategory !== null && !isText(a.rawCategory, 100)) return `Account "${a.name}" has an unreadable category.`;
  }
  // Budgets, goals and statements find accounts by name, so two of the same name would mix up.
  const names = accounts.map((a) => a.name.trim().toLowerCase());
  const twin = accounts.find((a, i) => names.indexOf(names[i]) !== i);
  if (twin) return `Two accounts are called "${twin.name.trim()}". Give each a different name.`;
  if (!Array.isArray(dates) || !dates.length || dates.length > 10000) return "Expected 1–10,000 dated snapshots.";
  if (!dates.every((d) => typeof d === "string" && /^\d{4}-\d{2}-\d{2}$/.test(d))) return "Dates must be YYYY-MM-DD.";
  const twinDate = dates.find((d, i) => dates.indexOf(d) !== i);
  if (twinDate) return `Two rows have the date ${twinDate}. Merge them into one.`;
  if (!Array.isArray(rows) || rows.length !== dates.length) return "Every snapshot needs a row of balances.";
  for (const r of rows) {
    if (!Array.isArray(r) || r.length !== accounts.length) return "Every row needs one value per account.";
    if (!r.every((v) => v === null || (typeof v === "number" && Number.isFinite(v)))) return "Balances must be numbers or blank.";
  }
  return null;
}
