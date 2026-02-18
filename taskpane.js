/* ===== CSV header order for the export ===== */
const CSV_HEADERS = [
  "TRADE_SUBTYPE",
  "LEGAL_ENTITY_CODE",
  "INTERMEDIARY_BANK",
  "VALUE_DATE",
  "CLIENT_CODE",
  "CLIENT_MASTER_ACCOUNT_NAME",
  "CLIENT_SUB_ACCOUNT",
  "SIDE",
  "AMOUNT",
  "CURRENCY",
  "NOSTRO_BANK",
  "NOSTRO_CODE",
  "COMMENT",
  "FILE_TYPE",
  "COUNTERPARTY_BIC",
  "COUNTERPARTY_ACCOUNT_NUMBER",
  "CUSTODY"
];

/* ===== App state ===== */
const state = {
  parsed: null, 
  lookup: null,  
  email: null,  
  parsedList: [], 
  lookupList: []     

};

/* ===== Wire-up buttons ===== */
Office.onReady(() => {
  document.getElementById("readBodyBtn")?.addEventListener("click", extractFromEmail);
  document.getElementById("exportCsvBtn")?.addEventListener("click", exportCsv);
  write("Ready. Select the alert email, then click “Extract from Email”.");
});

function write(text) {
  const el = document.getElementById("output");
  if (el) el.textContent = text || "";
}

/* -----------------------------------------------------------
 * Step 1: Read email, parse fields, and perform lookup
 * -----------------------------------------------------------
 */
async function extractFromEmail() {
  const item = Office.context.mailbox?.item;
  if (!item || item.itemType !== Office.MailboxEnums.ItemType.Message) {
    write("Please select an email first.");
    return;
  }

  const createdByEmail  = Office.context.mailbox?.userProfile?.emailAddress || "";
  const receivedDate    = item.dateTimeCreated || item.dateTimeReceived || new Date();
  const receivedDateISO = toCorrectFormat(receivedDate); // keeps your current dd/mm/yyyy behaviour

  item.body.getAsync(Office.CoercionType.Html, async (result) => {
    if (result.status !== Office.AsyncResultStatus.Succeeded) {
      write("Error reading body: " + (result.error?.message || "Unknown"));
      return;
    }

    const bodyText = result.value || "";

    // Split into blocks and parse each independently
    const blocks = splitIntoTxnBlocks(bodyText);

    // DEBUG: preview found blocks (first line of each)
    //console.log("Split into blocks:", blocks.length);
    //blocks.forEach((b, i) => {
    //console.log(`Block ${i+1}:`, b);
    //});

    const parsedList = [];
    for (const block of blocks) {
    const p = parseEmail(block);
    if (p && p.ok) {
        p.valueDate = receivedDateISO;
        parsedList.push(p);
    }
    }

    if (parsedList.length === 0) {
      document.getElementById("exportCsvBtn").disabled = true;
      write("❌ No valid transaction blocks found. Ensure each has Amount, Currency, and Account code.");
      return;
    }

    // Batch lookup (tolerate failure and/or missing rows)
    const lookupList = await findLookupRowsFor(parsedList);

    // Update state (also keep single values for backward compatibility)
    state.parsedList = parsedList;
    state.lookupList = lookupList;
    state.parsed     = parsedList[0] || null;
    state.lookup     = lookupList[0] || null;
    state.email = {
      receivedDateISO,
      createdByEmail,
      from: item.from?.emailAddress || "",
      subject: item.subject || ""
    };

    // Allow export as long as parsing worked
    document.getElementById("exportCsvBtn").disabled = false;

    // Summarise in output
    const summary = parsedList.map((p, i) => ({
      i: i + 1,
      amount: p.amount,
      currency: p.currency,
      accountCode: p.accountCode,
      lookupFound: !!lookupList[i]
    }));
    write([
      `=== Parsed ${parsedList.length} transaction(s) ===`,
      JSON.stringify(summary, null, 2),
      (lookupList.some(l => !l)
        ? "\n⚠️ Some transactions missing lookup: those CSVs will have blank lookup fields."
        : "\n✅ All transactions have lookup data. Ready to export.")
    ].join("\n"));
  });
}

/* -----------------------------------------------------------
 * Step 2: Build the final CSV line and download
 * -----------------------------------------------------------
 */
async function exportCsv() {
  // Prefer multi-transaction path; fall back to single if needed
  const list = (state.parsedList && state.parsedList.length)
    ? state.parsedList
    : (state.parsed ? [state.parsed] : []);

  if (!list.length) {
    write("⚠️ Please run “Extract from Email” first.");
    return;
  }

  let exported = 0;
  for (let i = 0; i < list.length; i++) {
    const parsed = list[i];
    // Lookup row aligned by index if available; else fall back to single state.lookup; else null
    const lookup = (state.lookupList && state.lookupList[i]) || state.lookup || null;

    const row = buildCsvRow(parsed, lookup, state.email);
    const csvText = toCsv([CSV_HEADERS, row]);

    const code = parsed.accountCode || `UNKNOWN_${i + 1}`;
    const ccy  = parsed.currency || "CCY";
    // add suffix _01, _02, ... when multiple rows
    const suffix = list.length > 1 ? `_${String(i + 1).padStart(2, "0")}` : "";
    const filename = `cash_deposit_${code}_${ccy}_${state.email.receivedDateISO}${suffix}.csv`;

    downloadCsv(filename, csvText);
    exported += 1;

    // spacing to avoid multi-download throttling
    await delay(120);
  }

  write(`Exported ${exported} CSV file(s).${(state.lookupList || []).some(l => !l) ? " (Some with blank lookup fields)" : ""}`);
}

/* Map parsed + lookup to the required CSV columns */
function buildCsvRow(parsed, lookup, email) {
  const legalEntity = lookup?.LegalEntity || "";
  const nostroPrefix = (parsed.currency === "USD") ? "CS-SEG-BOANY-IFE11025-" : "CS-SEG-BOAN-IFE11025-";
  const nostroCode = `${nostroPrefix}${parsed.currency}`;
  const nostroBank = (parsed.currency === "USD") ? "BAML1" : "BAML";
  const clientCode  = lookup?.ClientCode || "";
  const clientMaster= lookup?.ClientMasterAccount || "";
  const clientSub   = lookup?.ClientSubAccount || "";

  return [
    /* TRADE_SUBTYPE */                 "Client Cash",
    /* LEGAL_ENTITY_CODE */             legalEntity,
    /* INTERMEDIARY_BANK */             "",
    /* VALUE_DATE */                    parsed.valueDate,
    /* CLIENT_CODE */                   clientCode,
    /* CLIENT_MASTER_ACCOUNT_NAME */    clientMaster,
    /* CLIENT_SUB_ACCOUNT */            clientSub,
    /* SIDE */                          "CREDIT",
    /* AMOUNT */                        parsed.amount?.toFixed(2) ?? "",
    /* CURRENCY */                      parsed.currency || "",
    /* NOSTRO_BANK */                   nostroBank,
    /* NOSTRO_CODE */                   nostroCode,
    /* COMMENT */                       "Cash Deposit",
    /* FILE_TYPE */                     "CASH",
    /* COUNTERPARTY_BIC */              "XXXXXXXXXXX",
    /* COUNTERPARTY_ACCOUNT_NUMBER */   "",
    /* CUSTODY */                       "TRUE"
  ];
}

/* ===== CSV generation helpers ===== */
function toCsv(rows) {
  const esc = (v) => {
    const s = String(v ?? "");
    return /[",\r\n]/.test(s) ? `"${s.replace(/"/g, '""')}"` : s;
  };
  return rows.map(r => r.map(esc).join(",")).join("\r\n") + "\r\n";
}

function downloadCsv(filename, text) {
  const blob = new Blob([text], { type: "text/csv;charset=utf-8;" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  a.style.display = "none";
  document.body.appendChild(a);
  a.click();
  setTimeout(() => {
    URL.revokeObjectURL(url);
    a.remove();
  }, 0);
}

/* -----------------------------------------------------------
 * Lookup CSV loader (served by your add-in at ./data/…)
 * -----------------------------------------------------------
 */
let __lookupRows = null; // cached rows

async function ensureLookupLoaded() {
  if (__lookupRows) return __lookupRows; // in-session cache
  const url = "./data/Cash_Deposit_Lookup.csv"; // relative to taskpane.html
  const res = await fetch(url, { cache: "no-cache" });
  if (!res.ok) throw new Error(`HTTP ${res.status} while fetching ${url}`);
  const text = await res.text();
  __lookupRows = parseCsvToObjects(text); // array of objects by header names
  return __lookupRows;
}

async function loadLookupByAccountCode(accountCode) {
  if (!accountCode) return null;
  const rows = await ensureLookupLoaded();
  // Strict match on the AccountCode column
  return rows.find(r => (r.AccountCode || "").trim() === accountCode) || null;
}

/* Tiny CSV parser that supports quoted fields with commas */
function parseCsvToObjects(csvText) {
  const lines = csvText.replace(/\r\n/g, "\n").split("\n").filter(line => line.length > 0);
  if (lines.length === 0) return [];
  const headers = splitCsvLine(lines[0]).map(h => h.trim());
  return lines.slice(1).map(line => {
    const cells = splitCsvLine(line);
    const obj = {};
    headers.forEach((h, i) => { obj[h] = (cells[i] ?? "").trim(); });
    return obj;
  });
}

function splitCsvLine(line) {
  const out = [];
  let cur = "", inQuotes = false;
  for (let i = 0; i < line.length; i++) {
    const c = line[i];
    if (inQuotes) {
      if (c === '"' && line[i + 1] === '"') { cur += '"'; i++; } // escaped quote
      else if (c === '"') { inQuotes = false; }
      else { cur += c; }
    } else {
      if (c === '"') inQuotes = true;
      else if (c === ',') { out.push(cur); cur = ""; }
      else { cur += c; }
    }
  }
  out.push(cur);
  return out;
}

/* -----------------------------------------------------------
 * Email parsing
 * -----------------------------------------------------------
 */
function parseEmail(bodyText) {
  const text = normalizeWhitespace(bodyText);

  // 1) Amount + Currency (prefer "Amount: 30,000.00 USD")
  const amountLineRegex = /([\d,]+(?:\.\d{1,4})?)\s*([A-Z]{3})\b/i;
  let amountRaw = null, amount = null, currency = null;

  let m = text.match(amountLineRegex);
  if (m) {
    amountRaw = m[1];
    currency = m[2].toUpperCase();
    amount = normalizeAmountNumber(amountRaw);
  }

  // 2) Account code (optional now): 6 digits + 1 uppercase letter
  const accountCodeRegex = /\b(\d{6}[A-Z])\b/g;
  let accountCode = null;
  const candidates = [];
  for (const match of text.matchAll(accountCodeRegex)) {
    const idx = match.index ?? 0;
    const windowStart = Math.max(0, idx - 40);
    const windowText = text.slice(windowStart, idx + match[0].length + 20);
    const bias = (/\bAC\b/i.test(windowText) || /\/BNF\b/i.test(windowText) || /\/FFC\b/i.test(windowText)) ? 1 : 0;
    candidates.push({ code: match[1], bias, index: idx });
  }
  if (candidates.length) {
    candidates.sort((a, b) => (b.bias - a.bias) || (a.index - b.index));
    accountCode = candidates[0].code;
  }

  return {
    amount,
    currency,
    amountRaw,
    accountCode,   
    ok: Boolean(amount) && Boolean(currency) 
  };
}


/* ===== Utilities ===== */
function cleanupName(s) {
  return String(s || "").replace(/\s+/g, " ").trim();
}

function normalizeWhitespace(s) {
  return (s || "")
    .replace(/\r\n/g, "\n")
    .replace(/&nbsp;/gi, " ")
    // remove thin/zero‑width/BOM characters that break regexes
    .replace(/[\u2000-\u200F\u202F\u2060\uFEFF]/g, "")
    .replace(/[ \t]+/g, " ")
    .trim();
}

function normalizeAmountNumber(amountStr) {
  if (!amountStr) return null;
  const cleaned = amountStr.replace(/,/g, "");
  const num = Number(cleaned);
  return Number.isFinite(num) ? num : null;
}

function toCorrectFormat(d) {
  const dt = new Date(d);
  const yyyy = dt.getUTCFullYear();
  const mm = String(dt.getUTCMonth() + 1).padStart(2, "0");
  const dd = String(dt.getUTCDate()).padStart(2, "0");
  return `${dd}/${mm}/${yyyy}`;
}

/* Find all candidate blocks that start at an "Amount:" line */
function splitIntoTxnBlocks(fullText) {
  // Normalize pitfalls that break regex scans
  const text = (fullText || "")
    .replace(/\r\n/g, "\n")
    .replace(/&nbsp;/gi, " ")
    // remove thin/zero‑width/BOM and similar Unicode spacing
    .replace(/[\u2000-\u200F\u2028\u2029\u202F\u2060\uFEFF]/g, "");

  // Find every "Amount:" (or "Amount -") occurrence

  const startIndex = text.indexOf("<table")
  const endIndex = text.lastIndexOf("</table")
  const searchText = text.slice(startIndex, endIndex)
  const starts = [];
  const amountMarker = /Amount: /gi;
  let m;
  while ((m = amountMarker.exec(searchText)) !== null) {
    starts.push(m.index);
  }

  if (!starts.length) return [fullText]; // fallback: treat entire body as single block

  // Validation regex: Amount line MUST contain "<number>[.<decimals>] <CUR>"
  // Allow optional whitespace between '.' and decimals (Outlook often inserts thin spaces)
  const amountCurrencyAfterMarker = /([\d,]+(?:\.\d{1,4})?)\s*([A-Z]{3})\b/i;

  const blocks = [];
  for (let i = 0; i < starts.length; i++) {
    const start = starts[i];
    const end = (i + 1 < starts.length) ? starts[i + 1] : searchText.length;
    const candidate = searchText.slice(start, end);

    if (amountCurrencyAfterMarker.test(candidate)) {
      blocks.push(candidate);
    } else {
      // Not a valid block; skip it
      // console.debug("Skipped non-block candidate (no amount+currency on Amount line):", firstLine);
    }
  }

  // If nothing validated, fall back to single block so user can still export
  return blocks.length ? blocks : [fullText];
}

/* Batch find lookups for an array of parsed items (uses your cached rows) */
async function findLookupRowsFor(parsedList) {
  let rows = null;
  try { rows = await ensureLookupLoaded(); } catch { /* keep null */ }
  return parsedList.map(p => {
    if (!rows) return null;
    const needle = String(p.accountCode || "").trim().toUpperCase();
    return rows.find(r => String(r.AccountCode || "").trim().toUpperCase() === needle) || null;
  });
}

/* Tiny delay to avoid some browsers throttling multiple downloads */
function delay(ms) { return new Promise(res => setTimeout(res, ms)); }