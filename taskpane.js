// taskpane.js
Office.onReady(() => {
  const btn = document.getElementById("readBodyBtn");
  if (btn) btn.addEventListener("click", readBody);
});

function readBody() {
  const item = Office.context.mailbox?.item;
  if (!item || item.itemType !== Office.MailboxEnums.ItemType.Message) {
    write("Please select an email first.");
    return;
  }

  // Use CoercionType.Text: Outlook converts HTML into text reliably enough for regex parsing.
  item.body.getAsync(Office.CoercionType.Text, (result) => {
    if (result.status === Office.AsyncResultStatus.Succeeded) {
      const rawText = result.value || "";
      const parsed = parseEmail(rawText);

      // Present results in a clean way
      write(
        [
          "=== Raw Body (truncated to 3,000 chars) ===",
          rawText.slice(0, 3000),
          "",
          "=== Extracted Fields ===",
          JSON.stringify(parsed, null, 2),
        ].join("\n")
      );
    } else {
      write("Error: " + result.error.message);
    }
  });
}

/**
 * Parses the email text to extract:
 *  - amount (number)
 *  - currency (3-letter uppercase)
 *  - amountRaw (original match, e.g., "30,000.00")
 *  - accountCode (6 digits + 1 capital letter, e.g., 472852G)
 *
 * Strategy:
 *  - Prefer the "Amount:" line pattern like:   Amount:  30,000.00 USD
 *  - If not present, fallback to first occurrence of "<amount> <CURRENCY>"
 *  - Account code is searched across entire body by pattern: \b\d{6}[A-Z]\b
 */
function parseEmail(bodyText) {
  const text = normalizeWhitespace(bodyText);

  // 1) Extract Amount + Currency from an "Amount:" line, e.g.,
  //    "Amount: 30,000.00 USD"
  //    - allows thousands separators and decimals
  //    - currency: 3 uppercase letters
  // Matches groups:
  //   [1] amount number (with commas)
  //   [2] currency (3 letters)
  const amountLineRegex = /amount\s*[:\-]\s*([\d,]+(?:\.\d{1,4})?)\s*([A-Z]{3})\b/i;
  let amountRaw = null;
  let amount = null;
  let currency = null;

  let m = text.match(amountLineRegex);
  if (m) {
    amountRaw = m[1];
    currency = m[2].toUpperCase();
    amount = normalizeAmountNumber(amountRaw);
  } else {
    // Fallback: look for a "<number> <CUR>" pattern anywhere
    // This is less strict but can rescue cases where "Amount:" label is absent
    const anyAmountRegex = /\b([\d]{1,3}(?:,[\d]{3})*(?:\.\d{1,4})?)\s*([A-Z]{3})\b/g;
    let fallback = null;
    for (const match of text.matchAll(anyAmountRegex)) {
      // Heuristic: ignore likely "Account" numbers and prefer a place near "Amount" keyword
      // If text contains "Amount" nearby this match within 120 chars, prefer it.
      const idx = match.index ?? 0;
      const windowStart = Math.max(0, idx - 120);
      const windowText = text.slice(windowStart, idx + 60);
      if (/amount\b/i.test(windowText)) {
        fallback = match;
        break;
      }
      // Else keep the first reasonable candidate
      if (!fallback) fallback = match;
    }
    if (fallback) {
      amountRaw = fallback[1];
      currency = fallback[2].toUpperCase();
      amount = normalizeAmountNumber(amountRaw);
    }
  }

  // 2) Extract account code: exactly 6 digits followed by 1 uppercase letter.
  // Search full text because it often appears within a "Text:" or "/BNF/... AC 472852G" segment.
  const accountCodeRegex = /\b(\d{6}[A-Z])\b/g;
  let accountCode = null;
  // Collect all, but prefer ones preceded by "AC" or "/BNF" to reduce false positives
  const candidates = [];
  for (const match of text.matchAll(accountCodeRegex)) {
    const idx = match.index ?? 0;
    const windowStart = Math.max(0, idx - 40);
    const windowText = text.slice(windowStart, idx + match[0].length + 20);
    const bias =
      /\bAC\b/i.test(windowText) || /\/BNF\b/i.test(windowText) || /\/FFC\b/i.test(windowText)
        ? 1
        : 0;
    candidates.push({ code: match[1], bias, index: idx });
  }
  if (candidates.length) {
    // Sort by bias desc, then by earliest appearance
    candidates.sort((a, b) => (b.bias - a.bias) || (a.index - b.index));
    accountCode = candidates[0].code;
  }

  return {
    amount,        // number (e.g., 30000)
    currency,      // string (e.g., "USD")
    amountRaw,     // original amount text (e.g., "30,000.00")
    accountCode,   // string (e.g., "472852G")
    ok: Boolean(amount) && Boolean(currency) && Boolean(accountCode),
  };
}

/** Normalize all whitespace so regex is simpler and robust */
function normalizeWhitespace(s) {
  return (s || "")
    .replace(/\r\n/g, "\n")
    .replace(/[ \t]+/g, " ")
    .replace(/\u00A0/g, " ") // non-breaking spaces
    .trim();
}

/** Convert "30,000.00" -> 30000.00 (Number). Returns null if invalid. */
function normalizeAmountNumber(amountStr) {
  if (!amountStr) return null;
  const cleaned = amountStr.replace(/,/g, "");
  const num = Number(cleaned);
  return Number.isFinite(num) ? num : null;
}

function write(text) {
  document.getElementById("output").textContent = text || "";
}