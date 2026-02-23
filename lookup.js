/* -----------------------------------------------------------
 * Lookup CSV loader (served by your add-in at ./data/…)
 * -----------------------------------------------------------
 */

/* Batch find lookups for an array of parsed items (uses your cached rows) */
export async function findLookupRowsFor(parsedList) {
  let rows = null;
  try { rows = await ensureLookupLoaded(); } catch { /* keep null */ }
  return parsedList.map(p => {
    if (!rows) return null;
    const needle = String(p.accountCode || "").trim().toUpperCase();
    return rows.find(r => String(r.AccountCode || "").trim().toUpperCase() === needle) || null;
  });
}


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