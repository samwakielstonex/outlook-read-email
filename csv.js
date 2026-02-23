import { state } from "./taskpane.js";

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

function write(text) {
  const el = document.getElementById("output");
  if (el) el.textContent = text || "";
}


/* -----------------------------------------------------------
 * Build the final CSV line and download
 * -----------------------------------------------------------
 */
export async function exportCsv() {
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

/* Tiny delay to avoid some browsers throttling multiple downloads */
function delay(ms) { return new Promise(res => setTimeout(res, ms)); }