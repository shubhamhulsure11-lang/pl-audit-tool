import * as XLSX from "xlsx";
import JSZip from "jszip";

// ─── Numeric & String Normalizers ─────────────────────────────────────────────
export const cleanNum = (v) => {
  if (typeof v === "number") return isNaN(v) ? 0 : v;
  if (!v) return 0;
  const s = String(v).replace(/[,₹\s'"]/g, "").trim();
  const num = parseFloat(s);
  return isNaN(num) ? 0 : num;
};

export const inrFormat = (v) =>
  new Intl.NumberFormat("en-IN", {
    style: "currency",
    currency: "INR",
    maximumFractionDigits: 0,
  }).format(v || 0);

export const ordinal = (n) => {
  const i = parseInt(n, 10);
  if (isNaN(i)) return String(n);
  const j = i % 10,
    k = i % 100;
  if (j === 1 && k !== 11) return `${i}st`;
  if (j === 2 && k !== 12) return `${i}nd`;
  if (j === 3 && k !== 13) return `${i}rd`;
  return `${i}th`;
};

export const formatWeekLabel = (start, end) => `${ordinal(start)} to ${ordinal(end)}`;

export const generateWeekRanges = (firstStart = 1, firstEnd = 7, lastStart = 29, lastEnd = 31) => {
  const fStart = parseInt(firstStart, 10) || 1;
  const fEnd = parseInt(firstEnd, 10) || 7;
  const lStart = parseInt(lastStart, 10) || 29;
  const lEnd = parseInt(lastEnd, 10) || 31;

  const ranges = [{ start: fStart, end: fEnd, label: formatWeekLabel(fStart, fEnd), weekNum: 1 }];
  let currentStart = fEnd + 1;
  let wNum = 2;

  while (currentStart + 6 < lStart) {
    const currentEnd = currentStart + 6;
    ranges.push({
      start: currentStart,
      end: currentEnd,
      label: formatWeekLabel(currentStart, currentEnd),
      weekNum: wNum++,
    });
    currentStart = currentEnd + 1;
  }

  ranges.push({
    start: lStart,
    end: lEnd,
    label: formatWeekLabel(lStart, lEnd),
    weekNum: wNum,
  });

  return ranges;
};

// ─── Multi-format Date Parser ──────────────────────────────────────────────────
export const parseDayFromDate = (dateVal) => {
  if (!dateVal) return null;
  if (typeof dateVal === "number") {
    try {
      const d = XLSX.SSF.parse_date_code(dateVal);
      if (d && d.d) return d.d;
    } catch (_) {}
  }
  if (dateVal instanceof Date && !isNaN(dateVal)) {
    return dateVal.getDate();
  }
  const s = String(dateVal).trim();

  // ISO format first: YYYY-MM-DD or YYYY-MM-DD HH:MM:SS (e.g. Zomato/Swiggy exports).
  // Must be checked before the generic regex below, otherwise the year gets
  // misread as the day (e.g. "2026-07-01" -> "26" from "...26-07-01").
  const iso = s.match(/^(\d{4})[\s/-](\d{1,2})[\s/-](\d{1,2})/);
  if (iso) {
    const d = parseInt(iso[3], 10);
    if (d >= 1 && d <= 31) return d;
  }

  // Regex matches DD-MM-YYYY, "01 Oct", "1st Oct"
  const m = s.match(/^(\d{1,2})(?:st|nd|rd|th)?[\s/-]+([A-Za-z]{3,}|\d{1,2})/i);
  if (m) {
    const d = parseInt(m[1], 10);
    if (d >= 1 && d <= 31) return d;
  }
  const parts = s.match(/\d+/g);
  if (parts && parts.length >= 3) {
    if (parts[0].length === 4) return parseInt(parts[2], 10); // YYYY-MM-DD
    return parseInt(parts[0], 10); // DD-MM-YYYY
  }
  return null;
};

// ─── File Readers & Sheet Extractors ──────────────────────────────────────────

// Some platform exports (seen on Zomato "Order Level" reports) ship with a
// stale <dimension> tag in the sheet XML — e.g. it claims the sheet starts at
// row 9 when real content (including the header row) starts at row 1. Excel
// and openpyxl silently recompute the true range from the actual cells, but
// SheetJS trusts the declared tag and drops everything above it. That makes
// every downstream header lookup fail silently and return 0. Recompute the
// real range from the cell keys actually present so nothing gets dropped.
function fixStaleSheetRange(sheet) {
  if (!sheet) return sheet;
  let minRow = Infinity, minCol = Infinity, maxRow = -Infinity, maxCol = -Infinity;
  for (const key of Object.keys(sheet)) {
    if (key[0] === "!") continue;
    const addr = XLSX.utils.decode_cell(key);
    if (addr.r < minRow) minRow = addr.r;
    if (addr.c < minCol) minCol = addr.c;
    if (addr.r > maxRow) maxRow = addr.r;
    if (addr.c > maxCol) maxCol = addr.c;
  }
  if (minRow === Infinity) return sheet; // empty sheet, nothing to fix

  const declared = sheet["!ref"] ? XLSX.utils.decode_range(sheet["!ref"]) : null;
  if (!declared || declared.s.r > minRow || declared.s.c > minCol) {
    sheet["!ref"] = XLSX.utils.encode_range({
      s: { r: minRow, c: minCol },
      e: { r: Math.max(maxRow, declared ? declared.e.r : maxRow), c: Math.max(maxCol, declared ? declared.e.c : maxCol) },
    });
  }
  return sheet;
}

export async function readWorkbookFromFile(file) {
  let buffer;
  if (file instanceof ArrayBuffer) {
    buffer = file;
  } else if (file.data instanceof ArrayBuffer) {
    buffer = file.data;
  } else if (typeof file.arrayBuffer === "function") {
    buffer = await file.arrayBuffer();
  } else {
    throw new Error("Invalid file object");
  }
  const wb = XLSX.read(new Uint8Array(buffer), { type: "array", cellDates: true, raw: false });
  (wb.SheetNames || []).forEach((name) => fixStaleSheetRange(wb.Sheets[name]));
  return wb;
}

export function findSheetByPattern(wb, patterns) {
  const names = wb.SheetNames || [];
  for (const pat of patterns) {
    const p = pat.toLowerCase();
    const found = names.find((n) => n.toLowerCase().includes(p));
    if (found) return wb.Sheets[found];
  }
  return wb.Sheets[names[0]];
}

export function sheetToRows(sheet, headerRow = 1) {
  if (!sheet) return [];
  const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "" });
  if (rows.length < headerRow) return [];
  
  const headers = (rows[headerRow - 1] || []).map((h) => String(h || "").trim());
  const data = [];

  for (let r = headerRow; r < rows.length; r++) {
    const row = rows[r];
    if (!row || row.every((c) => c === "" || c === null || c === undefined)) continue;
    const obj = {};
    headers.forEach((h, colIdx) => {
      if (h) obj[h] = row[colIdx] ?? "";
    });
    // also keep numeric indexed row
    obj._raw = row;
    data.push(obj);
  }
  return data;
}

// ─── Helper to find header key by fuzzy matching ──────────────────────────────
export function getRowVal(row, candidates) {
  if (!row) return 0;
  const keys = Object.keys(row);
  for (const c of candidates) {
    const cLower = c.toLowerCase();
    // Prefer an exact (trimmed, case-insensitive) header match first, then a
    // substring match that is NOT a rate/percentage column (e.g. "Base
    // service fee %" must not satisfy the search for "Base service fee" —
    // that's a rate, not the fee amount). Only headers that literally END in
    // "%" are treated as rate columns; a "%" appearing inside a formula note
    // like "Base service fee\n[(12)% * (B)]" doesn't count.
    const exactKey = keys.find((k) => k.toLowerCase().trim() === cLower);
    const substringKeyNoPct = keys.find(
      (k) => k.toLowerCase().includes(cLower) && !k.trim().endsWith("%")
    );
    const anySubstringKey = keys.find((k) => k.toLowerCase().includes(cLower));
    const matchedKey = exactKey || substringKeyNoPct || anySubstringKey;
    if (matchedKey && row[matchedKey] !== undefined && row[matchedKey] !== "") {
      return cleanNum(row[matchedKey]);
    }
  }
  return 0;
}

export function getRowString(row, candidates) {
  if (!row) return "";
  const keys = Object.keys(row);
  for (const c of candidates) {
    const cLower = c.toLowerCase();
    const matchedKey = keys.find((k) => k.toLowerCase().includes(cLower));
    if (matchedKey && row[matchedKey] !== undefined) {
      return String(row[matchedKey]).trim();
    }
  }
  return "";
}

// ─── Match Day to Week Range ──────────────────────────────────────────────────
export function findWeekForDay(day, weekRanges) {
  if (day === null || day === undefined) return weekRanges[0];
  for (const w of weekRanges) {
    if (w.start <= w.end) {
      if (day >= w.start && day <= w.end) return w;
    } else {
      // Month boundary spillover
      if (day >= w.start || day <= w.end) return w;
    }
  }
  if (day > weekRanges[weekRanges.length - 1].end) return weekRanges[weekRanges.length - 1];
  return weekRanges[0];
}

// ─── Bank Statement Parser ────────────────────────────────────────────────────
export async function parseBankTransactions(bankFile) {
  if (!bankFile) return [];
  try {
    const wb = await readWorkbookFromFile(bankFile);
    const sheet = wb.Sheets[wb.SheetNames[0]];
    const rawRows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "" });
    const transactions = [];

    for (const row of rawRows) {
      if (!row || row.length === 0) continue;
      const rowStr = row.join(" ").toLowerCase();
      // Look for credit amounts and descriptions
      const nums = row.map(cleanNum).filter((v) => v > 0);
      const desc = row.map((c) => String(c || "")).join(" ");
      if (nums.length > 0) {
        transactions.push({
          rawText: desc,
          amounts: nums,
          maxAmount: Math.max(...nums),
          date: row[0] || "",
        });
      }
    }
    return transactions;
  } catch (e) {
    console.warn("Could not parse bank file:", e);
    return [];
  }
}

export function matchBankPayout(expectedAmt, bankTransactions, tolerance = 25) {
  if (!expectedAmt || !bankTransactions || bankTransactions.length === 0) {
    return { matched: false, actual: 0, diff: -expectedAmt };
  }
  const match = bankTransactions.find((tx) =>
    tx.amounts.some((amt) => Math.abs(amt - expectedAmt) <= tolerance)
  );
  if (match) {
    const actual = match.amounts.find((amt) => Math.abs(amt - expectedAmt) <= tolerance) || expectedAmt;
    return { matched: true, actual, diff: actual - expectedAmt, note: match.rawText };
  }
  return { matched: false, actual: 0, diff: -expectedAmt };
}

// ═════════════════════════════════════════════════════════════════════════════
// 1. ZOMATO RECONCILIATION ENGINE
// ═════════════════════════════════════════════════════════════════════════════
export async function runZomatoRecon({
  files = [],
  bankFile = null,
  clientName = "Client",
  month = "Current",
  firstWeekStart = 1,
  firstWeekEnd = 7,
  lastWeekStart = 29,
  lastWeekEnd = 31,
  mode = "weekly", // "weekly" | "consolidated"
}) {
  const weekRanges = generateWeekRanges(firstWeekStart, firstWeekEnd, lastWeekStart, lastWeekEnd);
  const bankTxs = await parseBankTransactions(bankFile);

  const weekDataMap = {};
  weekRanges.forEach((w) => {
    weekDataMap[w.weekNum] = {
      week: w,
      orders: 0,
      itemTotal: 0,
      packagingCharges: 0,
      compensationCancelled: 0,
      discounts: 0,
      gstCollected: 0,
      commission: 0,
      paymentMechanismFee: 0,
      discountOnServiceFee: 0,
      longDistanceFee: 0,
      customerCompensation: 0,
      tcs: 0,
      tds: 0,
      gstDeduction: 0,
      marketingAds: 0,
      hyperpure: 0,
      expectedPayout: 0,
      bankActual: 0,
      utrList: [],
    };
  });

  for (const file of files) {
    try {
      const wb = await readWorkbookFromFile(file);
      const fileName = file.name || "Invoice.xlsx";

      // Detect start day from filename or sheet
      let fileDay = parseDayFromDate(fileName);

      // Sheet 1: Order Level / Order Details / Orders
      const orderSheet = findSheetByPattern(wb, [
        "order level",
        "order details",
        "orders",
        "addition deductions details",
        "summary",
      ]);

      // Scan header row
      let headerRowIdx = 1;
      const rawGrid = XLSX.utils.sheet_to_json(orderSheet, { header: 1, defval: "" });
      for (let i = 0; i < Math.min(15, rawGrid.length); i++) {
        const rowStr = (rawGrid[i] || []).join(" ").toLowerCase();
        if (
          rowStr.includes("subtotal") ||
          rowStr.includes("item total") ||
          rowStr.includes("order id") ||
          rowStr.includes("bill amount") ||
          rowStr.includes("service fee")
        ) {
          headerRowIdx = i + 1;
          break;
        }
      }

      const rows = sheetToRows(orderSheet, headerRowIdx);

      // Sheet 2: Addition Deductions (D2W) if present
      const d2Sheet = findSheetByPattern(wb, [
        "addition deductions",
        "deductions",
        "tax report",
        "other charges",
      ]);
      const d2Rows = d2Sheet && d2Sheet !== orderSheet ? sheetToRows(d2Sheet, 1) : [];

      for (const r of rows) {
        const orderDateStr = getRowString(r, [
          "order date",
          "date",
          "transaction date",
          "settlement date",
        ]);
        const orderDay = parseDayFromDate(orderDateStr) || fileDay || 1;
        const targetWeek = findWeekForDay(orderDay, weekRanges);
        const wData = weekDataMap[targetWeek.weekNum];

        wData.orders += 1;
        wData.itemTotal += getRowVal(r, ["subtotal", "items total", "item total", "bill amount", "order amount"]);
        wData.packagingCharges += getRowVal(r, ["packaging charge", "packing charge"]);
        wData.compensationCancelled += getRowVal(r, ["net additions", "cancellation refund", "compensation"]);
        wData.discounts += getRowVal(r, ["restaurant discount", "promo", "bogo", "freebies", "gold", "discount"]);
        wData.gstCollected += getRowVal(r, ["total gst collected", "gst collected", "gst"]);
        wData.commission += getRowVal(r, ["base service fee", "service fee", "platform fee", "commission"]);
        wData.paymentMechanismFee += getRowVal(r, ["payment mechanism fee", "convenience fee"]);
        wData.discountOnServiceFee += getRowVal(r, ["discount on service fee", "service fee capping"]);
        wData.longDistanceFee += getRowVal(r, ["long distance", "fulfilment fee"]);
        wData.customerCompensation += getRowVal(r, ["customer compensation", "recoupment"]);
        wData.tcs += getRowVal(r, ["tcs"]);
        wData.tds += getRowVal(r, ["tds"]);
        wData.gstDeduction += getRowVal(r, ["gst deduction", "gst liability"]);

        const utr = getRowString(r, ["bank utr", "utr", "ctr"]);
        if (utr && !wData.utrList.includes(utr)) {
          wData.utrList.push(utr);
        }
      }

      // Process D2 additions/deductions (Ads, Hyperpure, etc.)
      for (const d2r of d2Rows) {
        const d2Text = JSON.stringify(d2r).toLowerCase();
        const amt = getRowVal(d2r, ["amount", "total", "net amount", "debit"]);
        const targetWeek = weekRanges[0];
        if (d2Text.includes("ads") || d2Text.includes("marketing") || d2Text.includes("brandverse")) {
          weekDataMap[targetWeek.weekNum].marketingAds += amt;
        }
        if (d2Text.includes("hyperpure")) {
          weekDataMap[targetWeek.weekNum].hyperpure += amt;
        }
      }
    } catch (err) {
      console.warn("Error parsing Zomato file:", file.name, err);
    }
  }

  // Build final structured table rows
  const weeks = weekRanges.map((w) => {
    const d = weekDataMap[w.weekNum];
    const grossSales = d.itemTotal + d.packagingCharges + d.compensationCancelled;
    const netSalesExclGst = Math.max(0, grossSales - d.discounts);
    const netCommission = Math.max(0, d.commission + d.paymentMechanismFee - d.discountOnServiceFee + d.longDistanceFee);
    const totalDeductions =
      netCommission +
      d.customerCompensation +
      d.tcs +
      d.tds +
      d.gstDeduction +
      d.marketingAds +
      d.hyperpure;
    const netPayout = Math.max(0, netSalesExclGst + d.gstCollected - totalDeductions);

    // Bank match
    const bankCheck = matchBankPayout(netPayout, bankTxs);

    return {
      weekNum: w.weekNum,
      label: w.label,
      orders: d.orders,
      grossSales,
      packagingCharges: d.packagingCharges,
      discounts: d.discounts,
      gstCollected: d.gstCollected,
      netSalesExclGst,
      commission: netCommission,
      marketingAds: d.marketingAds,
      otherDeductions: d.customerCompensation + d.tcs + d.tds + d.gstDeduction + d.hyperpure,
      expectedPayout: netPayout,
      bankActual: bankCheck.matched ? bankCheck.actual : 0,
      bankDiff: bankCheck.matched ? bankCheck.diff : -netPayout,
      bankMatched: bankCheck.matched,
      utr: d.utrList.join(", ") || "—",
    };
  });

  // Calculate Totals
  const total = {
    label: "Total",
    orders: weeks.reduce((s, w) => s + w.orders, 0),
    grossSales: weeks.reduce((s, w) => s + w.grossSales, 0),
    packagingCharges: weeks.reduce((s, w) => s + w.packagingCharges, 0),
    discounts: weeks.reduce((s, w) => s + w.discounts, 0),
    gstCollected: weeks.reduce((s, w) => s + w.gstCollected, 0),
    netSalesExclGst: weeks.reduce((s, w) => s + w.netSalesExclGst, 0),
    commission: weeks.reduce((s, w) => s + w.commission, 0),
    marketingAds: weeks.reduce((s, w) => s + w.marketingAds, 0),
    otherDeductions: weeks.reduce((s, w) => s + w.otherDeductions, 0),
    expectedPayout: weeks.reduce((s, w) => s + w.expectedPayout, 0),
    bankActual: weeks.reduce((s, w) => s + w.bankActual, 0),
    bankDiff: weeks.reduce((s, w) => s + w.bankDiff, 0),
    bankMatched: weeks.every((w) => w.bankMatched),
    utr: "—",
  };

  return {
    platform: "Zomato",
    clientName,
    month,
    mode,
    weeks,
    total,
    filesCount: files.length,
    hasBank: bankFile != null,
  };
}

// ═════════════════════════════════════════════════════════════════════════════
// 2. SWIGGY RECONCILIATION ENGINE
// ═════════════════════════════════════════════════════════════════════════════
export async function runSwiggyRecon({
  files = [],
  bankFile = null,
  clientName = "Client",
  month = "Current",
  firstWeekStart = 1,
  firstWeekEnd = 7,
  lastWeekStart = 29,
  lastWeekEnd = 31,
}) {
  const weekRanges = generateWeekRanges(firstWeekStart, firstWeekEnd, lastWeekStart, lastWeekEnd);
  const bankTxs = await parseBankTransactions(bankFile);

  const weekDataMap = {};
  weekRanges.forEach((w) => {
    weekDataMap[w.weekNum] = {
      week: w,
      orders: 0,
      itemTotal: 0,
      packagingCharges: 0,
      compensationCancelled: 0,
      discounts: 0,
      gstCollected: 0,
      platformFee: 0,
      callCenterFee: 0,
      swiggyOneFee: 0,
      pocketHeroFee: 0,
      longDistanceFee: 0,
      cancellationCharges: 0,
      customerComplaints: 0,
      tcs: 0,
      tds: 0,
      gstDeduction: 0,
      collectionCharges: 0,
      marketingAds: 0,
      utrList: [],
    };
  });

  for (const file of files) {
    try {
      const wb = await readWorkbookFromFile(file);
      const fileName = file.name || "Swiggy_Invoice.xlsx";
      let fileDay = parseDayFromDate(fileName);

      const orderSheet = findSheetByPattern(wb, ["order level", "orders", "summary"]);

      let headerRowIdx = 1;
      const rawGrid = XLSX.utils.sheet_to_json(orderSheet, { header: 1, defval: "" });
      for (let i = 0; i < Math.min(15, rawGrid.length); i++) {
        const rowStr = (rawGrid[i] || []).join(" ").toLowerCase();
        if (
          rowStr.includes("item total") ||
          rowStr.includes("order status") ||
          rowStr.includes("bill amount") ||
          rowStr.includes("commission")
        ) {
          headerRowIdx = i + 1;
          break;
        }
      }

      const rows = sheetToRows(orderSheet, headerRowIdx);

      for (const r of rows) {
        const orderDateStr = getRowString(r, ["order date", "date", "order_date"]);
        const orderDay = parseDayFromDate(orderDateStr) || fileDay || 1;
        const targetWeek = findWeekForDay(orderDay, weekRanges);
        const wData = weekDataMap[targetWeek.weekNum];

        const status = getRowString(r, ["order status", "status"]).toLowerCase();
        if (status && !status.includes("delivered") && !status.includes("cancelled")) {
          // skip non-orders if applicable
        }

        wData.orders += 1;
        wData.itemTotal += getRowVal(r, ["item total", "order total", "bill amount"]);
        wData.packagingCharges += getRowVal(r, ["packaging charges", "packing charges"]);
        wData.compensationCancelled += getRowVal(r, [
          "total customer paid",
          "complaint & cancellation charges",
          "compensation",
        ]);
        wData.discounts += getRowVal(r, [
          "restaurant discounts",
          "restaurant discount",
          "swiggy one exclusive offer discount",
          "discount",
        ]);
        wData.gstCollected += getRowVal(r, ["gst collected", "gst"]);
        wData.platformFee += getRowVal(r, ["commission", "platform fee"]);
        wData.callCenterFee += getRowVal(r, ["call center charges", "call center fees"]);
        wData.swiggyOneFee += getRowVal(r, ["swiggy one fees"]);
        wData.pocketHeroFee += getRowVal(r, ["pocket hero fees"]);
        wData.longDistanceFee += getRowVal(r, ["long distance charges"]);
        wData.cancellationCharges += getRowVal(r, ["restaurant cancellation charges"]);
        wData.customerComplaints += getRowVal(r, ["customer complaints"]);
        wData.tds += getRowVal(r, ["tds"]);
        wData.tcs += getRowVal(r, ["tcs"]);
        wData.gstDeduction += getRowVal(r, ["gst deduction", "gst liability"]);
        wData.collectionCharges += getRowVal(r, ["payment collection charges", "collection charges"]);

        const utr = getRowString(r, ["bank utr", "utr", "ctr"]);
        if (utr && !wData.utrList.includes(utr)) {
          wData.utrList.push(utr);
        }
      }
    } catch (err) {
      console.warn("Error parsing Swiggy file:", file.name, err);
    }
  }

  const weeks = weekRanges.map((w) => {
    const d = weekDataMap[w.weekNum];
    const totalIncome = d.itemTotal + d.packagingCharges + d.compensationCancelled;
    const commission =
      d.platformFee +
      d.callCenterFee +
      d.swiggyOneFee +
      d.pocketHeroFee +
      d.longDistanceFee +
      d.collectionCharges;
    const taxesAndDeductions =
      d.discounts +
      d.cancellationCharges +
      d.customerComplaints +
      d.tds +
      d.tcs +
      d.gstDeduction;
    const moneyReceived = Math.max(0, totalIncome + d.gstCollected - commission - taxesAndDeductions);

    const bankCheck = matchBankPayout(moneyReceived, bankTxs);

    return {
      weekNum: w.weekNum,
      label: w.label,
      orders: d.orders,
      totalIncome,
      packagingCharges: d.packagingCharges,
      discounts: d.discounts,
      gstCollected: d.gstCollected,
      commission,
      otherCharges: d.callCenterFee + d.swiggyOneFee + d.longDistanceFee + d.collectionCharges,
      taxesAndDeductions,
      expectedPayout: moneyReceived,
      bankActual: bankCheck.matched ? bankCheck.actual : 0,
      bankDiff: bankCheck.matched ? bankCheck.diff : -moneyReceived,
      bankMatched: bankCheck.matched,
      utr: d.utrList.join(", ") || "—",
    };
  });

  const total = {
    label: "Total",
    orders: weeks.reduce((s, w) => s + w.orders, 0),
    totalIncome: weeks.reduce((s, w) => s + w.totalIncome, 0),
    packagingCharges: weeks.reduce((s, w) => s + w.packagingCharges, 0),
    discounts: weeks.reduce((s, w) => s + w.discounts, 0),
    gstCollected: weeks.reduce((s, w) => s + w.gstCollected, 0),
    commission: weeks.reduce((s, w) => s + w.commission, 0),
    otherCharges: weeks.reduce((s, w) => s + w.otherCharges, 0),
    taxesAndDeductions: weeks.reduce((s, w) => s + w.taxesAndDeductions, 0),
    expectedPayout: weeks.reduce((s, w) => s + w.expectedPayout, 0),
    bankActual: weeks.reduce((s, w) => s + w.bankActual, 0),
    bankDiff: weeks.reduce((s, w) => s + w.bankDiff, 0),
    bankMatched: weeks.every((w) => w.bankMatched),
    utr: "—",
  };

  return {
    platform: "Swiggy",
    clientName,
    month,
    weeks,
    total,
    filesCount: files.length,
    hasBank: bankFile != null,
  };
}

// ═════════════════════════════════════════════════════════════════════════════
// 3. SWIGGY DINEOUT RECONCILIATION ENGINE
// ═════════════════════════════════════════════════════════════════════════════
export async function runDineoutRecon({
  files = [],
  bankFile = null,
  clientName = "Client",
  month = "Current",
  firstWeekStart = 1,
  firstWeekEnd = 7,
  lastWeekStart = 29,
  lastWeekEnd = 31,
}) {
  const weekRanges = generateWeekRanges(firstWeekStart, firstWeekEnd, lastWeekStart, lastWeekEnd);
  const bankTxs = await parseBankTransactions(bankFile);

  const weekDataMap = {};
  weekRanges.forEach((w) => {
    weekDataMap[w.weekNum] = {
      week: w,
      salesExclGst: 0,
      discounts: 0,
      excessPaymentAdjustments: 0,
      packagingServiceCharges: 0,
      tips: 0,
      platformFee: 0,
      discountOnServiceFee: 0,
      accessCharges: 0,
      callCenterFee: 0,
      longDistanceFee: 0,
      bannerAdsFee: 0,
      deliveryServiceFee: 0,
      collectionCharges: 0,
      tds: 0,
      tcs: 0,
      netSettlement: 0,
    };
  });

  for (const file of files) {
    try {
      const wb = await readWorkbookFromFile(file);
      const sheet = findSheetByPattern(wb, ["dineout", "orders", "sheet1", "summary"]);
      const rows = sheetToRows(sheet, 1);

      for (const r of rows) {
        const dateStr = getRowString(r, ["date", "transaction date", "bill date"]);
        const day = parseDayFromDate(dateStr) || parseDayFromDate(file.name) || 1;
        const targetWeek = findWeekForDay(day, weekRanges);
        const wData = weekDataMap[targetWeek.weekNum];

        wData.salesExclGst += getRowVal(r, ["sales", "bill amount", "gross amount", "subtotal"]);
        wData.discounts += getRowVal(r, ["discount", "merchant discount", "dineout discount"]);
        wData.tips += getRowVal(r, ["tip", "tips"]);
        wData.platformFee += getRowVal(r, ["platform service fee", "commission", "service fee"]);
        wData.discountOnServiceFee += getRowVal(r, ["discount on swiggy service fee", "service fee discount"]);
        wData.collectionCharges += getRowVal(r, ["collection charges", "payment gateway fee"]);
        wData.tds += getRowVal(r, ["tds"]);
        wData.tcs += getRowVal(r, ["tcs"]);
        wData.netSettlement += getRowVal(r, ["net settlement", "settlement amount", "payout"]);
      }
    } catch (e) {
      console.warn("Error parsing Dineout file:", file.name, e);
    }
  }

  const weeks = weekRanges.map((w) => {
    const d = weekDataMap[w.weekNum];
    const salesAfterDiscounts = Math.max(0, d.salesExclGst - d.discounts - d.excessPaymentAdjustments);
    const gst5Pct = salesAfterDiscounts * 0.05;
    const salesInclGst = salesAfterDiscounts + gst5Pct + d.packagingServiceCharges + d.tips;
    const totalCommission = Math.max(
      0,
      d.platformFee -
        d.discountOnServiceFee +
        d.accessCharges +
        d.callCenterFee +
        d.longDistanceFee +
        d.bannerAdsFee +
        d.deliveryServiceFee +
        d.collectionCharges
    );
    const expectedPayout = Math.max(0, salesInclGst - totalCommission - d.tds - d.tcs);

    const bankCheck = matchBankPayout(expectedPayout, bankTxs);

    return {
      weekNum: w.weekNum,
      label: w.label,
      salesExclGst: d.salesExclGst,
      discounts: d.discounts,
      salesAfterDiscounts,
      gst5Pct,
      salesInclGst,
      commission: totalCommission,
      tdsTcs: d.tds + d.tcs,
      expectedPayout,
      bankActual: bankCheck.matched ? bankCheck.actual : 0,
      bankDiff: bankCheck.matched ? bankCheck.diff : -expectedPayout,
      bankMatched: bankCheck.matched,
    };
  });

  const total = {
    label: "Total",
    salesExclGst: weeks.reduce((s, w) => s + w.salesExclGst, 0),
    discounts: weeks.reduce((s, w) => s + w.discounts, 0),
    salesAfterDiscounts: weeks.reduce((s, w) => s + w.salesAfterDiscounts, 0),
    gst5Pct: weeks.reduce((s, w) => s + w.gst5Pct, 0),
    salesInclGst: weeks.reduce((s, w) => s + w.salesInclGst, 0),
    commission: weeks.reduce((s, w) => s + w.commission, 0),
    tdsTcs: weeks.reduce((s, w) => s + w.tdsTcs, 0),
    expectedPayout: weeks.reduce((s, w) => s + w.expectedPayout, 0),
    bankActual: weeks.reduce((s, w) => s + w.bankActual, 0),
    bankDiff: weeks.reduce((s, w) => s + w.bankDiff, 0),
    bankMatched: weeks.every((w) => w.bankMatched),
  };

  return {
    platform: "Swiggy Dineout",
    clientName,
    month,
    weeks,
    total,
    filesCount: files.length,
    hasBank: bankFile != null,
  };
}

// ═════════════════════════════════════════════════════════════════════════════
// 4. ZOMATO PAY RECONCILIATION ENGINE
// ═════════════════════════════════════════════════════════════════════════════
export async function runZomatoPayRecon({
  files = [],
  adsFiles = [],
  bankFile = null,
  clientName = "Client",
  month = "Current",
  firstWeekStart = 1,
  firstWeekEnd = 7,
  lastWeekStart = 29,
  lastWeekEnd = 31,
}) {
  const weekRanges = generateWeekRanges(firstWeekStart, firstWeekEnd, lastWeekStart, lastWeekEnd);
  const bankTxs = await parseBankTransactions(bankFile);

  const weekDataMap = {};
  weekRanges.forEach((w) => {
    weekDataMap[w.weekNum] = {
      week: w,
      billAmount: 0,
      discounts: 0,
      failedTransactions: 0,
      tips: 0,
      commission: 0,
      ads: 0,
    };
  });

  for (const file of files) {
    try {
      const wb = await readWorkbookFromFile(file);
      const sheet = findSheetByPattern(wb, ["zomato pay", "calculations", "orders", "sheet1"]);
      const rows = sheetToRows(sheet, 1);

      for (const r of rows) {
        const dateStr = getRowString(r, ["date", "transaction date", "payment date"]);
        const day = parseDayFromDate(dateStr) || parseDayFromDate(file.name) || 1;
        const targetWeek = findWeekForDay(day, weekRanges);
        const wData = weekDataMap[targetWeek.weekNum];

        wData.billAmount += getRowVal(r, ["bill amount", "order amount", "gross amount", "amount"]);
        wData.discounts += getRowVal(r, ["discount", "promo", "offers"]);
        wData.failedTransactions += getRowVal(r, ["failed", "reversed", "refund"]);
        wData.tips += getRowVal(r, ["tip", "tips"]);
        wData.commission += getRowVal(r, ["commission", "service fee", "platform fee"]);
      }
    } catch (e) {
      console.warn("Error parsing Zomato Pay file:", file.name, e);
    }
  }

  // Process Ads
  for (const file of adsFiles) {
    try {
      const wb = await readWorkbookFromFile(file);
      const sheet = wb.Sheets[wb.SheetNames[0]];
      const rows = sheetToRows(sheet, 1);
      for (const r of rows) {
        const dateStr = getRowString(r, ["date"]);
        const day = parseDayFromDate(dateStr) || 1;
        const targetWeek = findWeekForDay(day, weekRanges);
        weekDataMap[targetWeek.weekNum].ads += getRowVal(r, ["amount", "ad spend", "cost"]);
      }
    } catch (e) {
      console.warn("Error parsing Ads file:", file.name, e);
    }
  }

  const weeks = weekRanges.map((w) => {
    const d = weekDataMap[w.weekNum];
    // Formula from Python repo: (billAmount * 100/105) - (discounts * 100/105)
    const salesExclGstBefore = d.billAmount * (100.0 / 105.0);
    const discountsExclGst = d.discounts * (100.0 / 105.0);
    const salesExclGstAfter = Math.max(0, salesExclGstBefore - discountsExclGst - d.failedTransactions);
    const gst5Pct = salesExclGstAfter * 0.05;
    const salesInclGst = salesExclGstAfter + gst5Pct + d.tips;
    const commissionInclGst = d.commission * 1.18;
    const expectedPayout = Math.max(0, salesInclGst - commissionInclGst - d.ads);

    const bankCheck = matchBankPayout(expectedPayout, bankTxs);

    return {
      weekNum: w.weekNum,
      label: w.label,
      salesExclGstBefore,
      discounts: discountsExclGst,
      failedTransactions: d.failedTransactions,
      salesExclGstAfter,
      gst5Pct,
      tips: d.tips,
      salesInclGst,
      commissionInclGst,
      ads: d.ads,
      expectedPayout,
      bankActual: bankCheck.matched ? bankCheck.actual : 0,
      bankDiff: bankCheck.matched ? bankCheck.diff : -expectedPayout,
      bankMatched: bankCheck.matched,
    };
  });

  const total = {
    label: "Total",
    salesExclGstBefore: weeks.reduce((s, w) => s + w.salesExclGstBefore, 0),
    discounts: weeks.reduce((s, w) => s + w.discounts, 0),
    failedTransactions: weeks.reduce((s, w) => s + w.failedTransactions, 0),
    salesExclGstAfter: weeks.reduce((s, w) => s + w.salesExclGstAfter, 0),
    gst5Pct: weeks.reduce((s, w) => s + w.gst5Pct, 0),
    tips: weeks.reduce((s, w) => s + w.tips, 0),
    salesInclGst: weeks.reduce((s, w) => s + w.salesInclGst, 0),
    commissionInclGst: weeks.reduce((s, w) => s + w.commissionInclGst, 0),
    ads: weeks.reduce((s, w) => s + w.ads, 0),
    expectedPayout: weeks.reduce((s, w) => s + w.expectedPayout, 0),
    bankActual: weeks.reduce((s, w) => s + w.bankActual, 0),
    bankDiff: weeks.reduce((s, w) => s + w.bankDiff, 0),
    bankMatched: weeks.every((w) => w.bankMatched),
  };

  return {
    platform: "Zomato Pay",
    clientName,
    month,
    weeks,
    total,
    filesCount: files.length,
    hasBank: bankFile != null,
  };
}

// ═════════════════════════════════════════════════════════════════════════════
// 5. PAYTM RECONCILIATION ENGINE
// ═════════════════════════════════════════════════════════════════════════════
export async function runPaytmRecon({
  files = [],
  bankFile = null,
  clientName = "Client",
  month = "Current",
  firstWeekStart = 1,
  firstWeekEnd = 7,
  lastWeekStart = 29,
  lastWeekEnd = 31,
}) {
  const weekRanges = generateWeekRanges(firstWeekStart, firstWeekEnd, lastWeekStart, lastWeekEnd);
  const bankTxs = await parseBankTransactions(bankFile);

  const weekDataMap = {};
  weekRanges.forEach((w) => {
    weekDataMap[w.weekNum] = {
      week: w,
      salesExclGst: 0,
      failedTransactions: 0,
      commission: 0,
      netSettlement: 0,
    };
  });

  for (const file of files) {
    try {
      const wb = await readWorkbookFromFile(file);
      const sheet = findSheetByPattern(wb, ["paytm", "settlement", "sheet1"]);
      const rows = sheetToRows(sheet, 1);

      for (const r of rows) {
        const dateStr = getRowString(r, ["date", "transaction date", "settlement date"]);
        const day = parseDayFromDate(dateStr) || parseDayFromDate(file.name) || 1;
        const targetWeek = findWeekForDay(day, weekRanges);
        const wData = weekDataMap[targetWeek.weekNum];

        wData.salesExclGst += getRowVal(r, ["amount", "order amount", "gross amount", "total"]);
        wData.failedTransactions += getRowVal(r, ["failed", "refund", "chargeback"]);
        wData.commission += getRowVal(r, ["fee", "commission", "mdr", "charges"]);
        wData.netSettlement += getRowVal(r, ["settlement amount", "net payout", "net amount"]);
      }
    } catch (e) {
      console.warn("Error parsing Paytm file:", file.name, e);
    }
  }

  const weeks = weekRanges.map((w) => {
    const d = weekDataMap[w.weekNum];
    const salesAfterFailed = Math.max(0, d.salesExclGst - d.failedTransactions);
    const gst5Pct = salesAfterFailed * 0.05;
    const salesInclGst = salesAfterFailed + gst5Pct;
    const commissionInclGst = d.commission * 1.18;
    const expectedReceipt = Math.max(0, salesInclGst - commissionInclGst);

    const bankCheck = matchBankPayout(expectedReceipt, bankTxs);

    return {
      weekNum: w.weekNum,
      label: w.label,
      salesExclGst: d.salesExclGst,
      failedTransactions: d.failedTransactions,
      salesAfterFailed,
      gst5Pct,
      salesInclGst,
      commissionInclGst,
      expectedReceipt,
      bankActual: bankCheck.matched ? bankCheck.actual : 0,
      bankDiff: bankCheck.matched ? bankCheck.diff : -expectedReceipt,
      bankMatched: bankCheck.matched,
    };
  });

  const total = {
    label: "Total",
    salesExclGst: weeks.reduce((s, w) => s + w.salesExclGst, 0),
    failedTransactions: weeks.reduce((s, w) => s + w.failedTransactions, 0),
    salesAfterFailed: weeks.reduce((s, w) => s + w.salesAfterFailed, 0),
    gst5Pct: weeks.reduce((s, w) => s + w.gst5Pct, 0),
    salesInclGst: weeks.reduce((s, w) => s + w.salesInclGst, 0),
    commissionInclGst: weeks.reduce((s, w) => s + w.commissionInclGst, 0),
    expectedReceipt: weeks.reduce((s, w) => s + w.expectedReceipt, 0),
    bankActual: weeks.reduce((s, w) => s + w.bankActual, 0),
    bankDiff: weeks.reduce((s, w) => s + w.bankDiff, 0),
    bankMatched: weeks.every((w) => w.bankMatched),
  };

  return {
    platform: "Paytm",
    clientName,
    month,
    weeks,
    total,
    filesCount: files.length,
    hasBank: bankFile != null,
  };
}

// ═════════════════════════════════════════════════════════════════════════════
// 6. POS CLEANER & CHANNEL AGGREGATOR
// ═════════════════════════════════════════════════════════════════════════════
export async function runPosCleaner({
  file,
  firstWeekStart = 1,
  firstWeekEnd = 7,
  lastWeekStart = 29,
  lastWeekEnd = 31,
}) {
  if (!file) throw new Error("No POS file provided");
  const weekRanges = generateWeekRanges(firstWeekStart, firstWeekEnd, lastWeekStart, lastWeekEnd);
  const wb = await readWorkbookFromFile(file);
  const sheet = wb.Sheets[wb.SheetNames[0]];

  let headerRowIdx = 1;
  const rawGrid = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "" });
  for (let i = 0; i < Math.min(20, rawGrid.length); i++) {
    const rowStr = (rawGrid[i] || []).join(" ").toLowerCase();
    if (
      rowStr.includes("payment type") ||
      rowStr.includes("payment_type") ||
      rowStr.includes("order type") ||
      rowStr.includes("invoice") ||
      rowStr.includes("bill no")
    ) {
      headerRowIdx = i + 1;
      break;
    }
  }

  const rows = sheetToRows(sheet, headerRowIdx);
  const channels = ["Zomato", "Swiggy", "Dineout", "Paytm", "Cash", "Card", "UPI", "Other"];
  const channelData = {};

  channels.forEach((ch) => {
    channelData[ch] = {
      channel: ch,
      weeks: weekRanges.map((w) => ({ weekNum: w.weekNum, label: w.label, amount: 0, orders: 0 })),
      total: 0,
      totalOrders: 0,
    };
  });

  for (const r of rows) {
    const paymentType = getRowString(r, ["payment type", "payment_type", "order type", "source"]).toLowerCase();
    const dateStr = getRowString(r, ["date", "transaction date", "order date", "bill date"]);
    const day = parseDayFromDate(dateStr) || 1;
    const amount = getRowVal(r, ["amount", "total", "net amount", "grand total", "bill amount"]);

    let matchedChannel = "Other";
    if (paymentType.includes("zomato") && !paymentType.includes("pay")) matchedChannel = "Zomato";
    else if (paymentType.includes("zomato pay") || paymentType.includes("zpay")) matchedChannel = "Dineout";
    else if (paymentType.includes("swiggy") && !paymentType.includes("dineout")) matchedChannel = "Swiggy";
    else if (paymentType.includes("dineout")) matchedChannel = "Dineout";
    else if (paymentType.includes("paytm")) matchedChannel = "Paytm";
    else if (paymentType.includes("cash")) matchedChannel = "Cash";
    else if (paymentType.includes("card") || paymentType.includes("credit") || paymentType.includes("debit") || paymentType.includes("edc")) matchedChannel = "Card";
    else if (paymentType.includes("upi") || paymentType.includes("gpay") || paymentType.includes("phonepe") || paymentType.includes("qr")) matchedChannel = "UPI";

    const targetWeek = findWeekForDay(day, weekRanges);
    const chObj = channelData[matchedChannel];
    const wObj = chObj.weeks.find((w) => w.weekNum === targetWeek.weekNum);
    if (wObj) {
      wObj.amount += amount;
      wObj.orders += 1;
    }
    chObj.total += amount;
    chObj.totalOrders += 1;
  }

  return {
    fileName: file.name || "POS_Report.xlsx",
    channels: Object.values(channelData),
    grandTotal: Object.values(channelData).reduce((s, c) => s + c.total, 0),
    grandOrders: Object.values(channelData).reduce((s, c) => s + c.totalOrders, 0),
    weekRanges,
  };
}

// ═════════════════════════════════════════════════════════════════════════════
// 7. EXPORT RECONCILIATION WORKBOOK (.XLSX)
// ═════════════════════════════════════════════════════════════════════════════
export function exportReconWorkbook(report) {
  const wb = XLSX.utils.book_new();
  const title = `${report.platform} Reconciliation - ${report.clientName} (${report.month})`;

  // Summary Sheet Data
  const sheetData = [
    [title],
    ["Client Name:", report.clientName, "Month:", report.month, "Generated:", new Date().toLocaleString("en-IN")],
    [],
  ];

  if (report.platform === "Zomato" || report.platform === "Swiggy") {
    sheetData.push([
      "Week / Period",
      "No. of Orders",
      "Gross Sales (₹)",
      "Discounts (₹)",
      "GST Collected (₹)",
      "Commission (₹)",
      "Other Deductions (₹)",
      "Expected Payout (₹)",
      "Bank Actual (₹)",
      "Bank Difference (₹)",
      "Status",
    ]);

    report.weeks.forEach((w) => {
      sheetData.push([
        w.label,
        w.orders,
        w.grossSales || w.totalIncome,
        w.discounts,
        w.gstCollected,
        w.commission,
        w.otherDeductions || w.taxesAndDeductions,
        w.expectedPayout,
        w.bankActual,
        w.bankDiff,
        w.bankMatched ? "MATCH" : "DISCREPANCY",
      ]);
    });

    sheetData.push([
      report.total.label,
      report.total.orders,
      report.total.grossSales || report.total.totalIncome,
      report.total.discounts,
      report.total.gstCollected,
      report.total.commission,
      report.total.otherDeductions || report.total.taxesAndDeductions,
      report.total.expectedPayout,
      report.total.bankActual,
      report.total.bankDiff,
      report.total.bankMatched ? "ALL MATCHED" : "FLAGGED",
    ]);
  } else {
    sheetData.push([
      "Week / Period",
      "Sales Excl. GST (₹)",
      "Discounts (₹)",
      "GST 5% (₹)",
      "Sales Incl. GST (₹)",
      "Commission (₹)",
      "Expected Payout (₹)",
      "Bank Actual (₹)",
      "Difference (₹)",
      "Status",
    ]);

    report.weeks.forEach((w) => {
      sheetData.push([
        w.label,
        w.salesExclGst || w.salesExclGstBefore,
        w.discounts,
        w.gst5Pct,
        w.salesInclGst || w.salesAfterFailed,
        w.commission || w.commissionInclGst,
        w.expectedPayout || w.expectedReceipt,
        w.bankActual,
        w.bankDiff,
        w.bankMatched ? "MATCH" : "DISCREPANCY",
      ]);
    });

    sheetData.push([
      report.total.label,
      report.total.salesExclGst || report.total.salesExclGstBefore,
      report.total.discounts,
      report.total.gst5Pct,
      report.total.salesInclGst || report.total.salesAfterFailed,
      report.total.commission || report.total.commissionInclGst,
      report.total.expectedPayout || report.total.expectedReceipt,
      report.total.bankActual,
      report.total.bankDiff,
      report.total.bankMatched ? "ALL MATCHED" : "FLAGGED",
    ]);
  }

  const ws = XLSX.utils.aoa_to_sheet(sheetData);
  XLSX.utils.book_append_sheet(wb, ws, "Reconciliation Summary");

  const fileName = `${report.clientName}_${report.platform.replace(/\s+/g, "_")}_Recon_${report.month}.xlsx`;
  XLSX.writeFile(wb, fileName);
}
