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

// ─── Exact (whitespace-normalized) header matching for columns that must ─────
// never be fuzzy-matched (e.g. summing two distinctly-named discount columns
// without risking a double count, or picking the exact TCS column instead of
// a similarly-worded "Applicable amount for TCS" base column).
export function normalizeHeader(s) {
  return String(s || "").toLowerCase().replace(/\s+/g, " ").trim();
}

export function sumExactCols(row, candidateNames) {
  if (!row) return 0;
  const normCandidates = candidateNames.map(normalizeHeader);
  let total = 0;
  for (const k of Object.keys(row)) {
    if (normCandidates.includes(normalizeHeader(k))) {
      total += cleanNum(row[k]);
    }
  }
  return total;
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

  // Zomato's own settlement report labels several fee columns as their
  // pre-GST amount; Zomato charges 18% GST on its own service/convenience/
  // long-distance fees, so those columns must be grossed up by 1.18x to
  // match what the platform actually deducts. This constant mirrors that.
  const ZOMATO_FEE_GST_MULTIPLIER = 1.18;

  const DISCOUNT_EXACT_COLS = [
    "Restaurant discount (Promo)",
    "Restaurant discount [Promo]",
    "Restaurant discount (BOGO, Freebies, Gold, Brand pack & others)",
    "Restaurant Discount [BOGO, Freebies, Gold, Brand pack & others]",
    "Restaurant discount [Flat offs, Freebies, Gold, Brand pack & others]",
    "Delivery charge discount/ Relisting discount",
    "Delivery charge discount / Relisting discount",
  ];
  const TCS_EXACT_COLS = ["TCS IGST amount", "Tax collected at source"];
  const CANCELLED_STATUSES = ["CANCELLED", "TIMEDOUT", "TIMEOUT", "REJECTED"];

  const weekDataMap = {};
  weekRanges.forEach((w) => {
    weekDataMap[w.weekNum] = {
      week: w,
      orders: 0,
      // Delivered-orders-only figures (Zomato only counts a delivered
      // order's item value/fees toward payout; cancelled orders don't).
      itemTotalDelivered: 0,
      packagingDelivered: 0,
      discountDelivered: 0,
      gstDelivered: 0,
      baseServiceFeeDelivered: 0,
      convenienceFeeDelivered: 0,
      discountOnServiceFeeDelivered: 0,
      longDistanceFeeDelivered: 0,
      discountOnLongDistanceDelivered: 0,
      // Cancelled-orders-only figure.
      compensationCancelled: 0,
      // Applies across delivered + cancelled rows (matches Zomato's own
      // "total" aggregation for tax lines).
      tdsAll: 0,
      tcsAll: 0,
      gstPaidByZomatoAll: 0,
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

        const status = getRowString(r, ["order status"]).toUpperCase();
        const isDelivered = status === "DELIVERED";
        const isCancelledGroup = CANCELLED_STATUSES.includes(status);

        // Only count recognized (delivered/cancelled) rows — matches the
        // reference tool's order count, which excludes malformed/unrecognized
        // rows (e.g. a stray #REF! row) that aren't real orders.
        if (isDelivered || isCancelledGroup) {
          wData.orders += 1;
        }

        if (isDelivered) {
          wData.itemTotalDelivered += getRowVal(r, [
            "subtotal (items total)",
            "subtotal",
            "item total",
            "bill amount",
            "order amount",
          ]);
          wData.packagingDelivered += getRowVal(r, ["packaging charge", "packing charge"]);
          wData.discountDelivered += sumExactCols(r, DISCOUNT_EXACT_COLS);
          wData.gstDelivered += getRowVal(r, ["total gst collected from customers", "total gst collected", "gst collected"]);
          wData.baseServiceFeeDelivered += getRowVal(r, ["base service fee", "service fee", "platform fee"]);
          wData.convenienceFeeDelivered += getRowVal(r, ["payment mechanism fee", "convenience fee"]);
          wData.discountOnServiceFeeDelivered += getRowVal(r, ["discount on service fee", "service fee capping"]);
          wData.longDistanceFeeDelivered += getRowVal(r, ["long distance enablement fee", "long distance", "fulfilment fee"]);
          wData.discountOnLongDistanceDelivered += getRowVal(r, ["discount on long distance enablement fee"]);
        }
        if (isCancelledGroup) {
          wData.compensationCancelled += getRowVal(r, ["net additions", "cancellation refund", "compensation"]);
        }
        // Tax lines are Zomato's "total" (delivered + cancelled combined) aggregation.
        if (isDelivered || isCancelledGroup) {
          wData.tdsAll += getRowVal(r, ["tds 194o amount", "tds"]);
          wData.tcsAll += sumExactCols(r, TCS_EXACT_COLS);
          wData.gstPaidByZomatoAll += getRowVal(r, ["gst paid by zomato on behalf of restaurant", "gst paid by zomato"]);
        }

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
    const grossSales = d.itemTotalDelivered + d.packagingDelivered + d.compensationCancelled;
    const discounts = d.discountDelivered;
    const gstCollected = d.gstDelivered;
    const netSalesExclGst = Math.max(0, grossSales - discounts);
    const netSalesInclGst = netSalesExclGst + gstCollected;

    // Zomato's own service/convenience/long-distance fees carry 18% GST on
    // top of the raw column value in the settlement report.
    const commission = d.baseServiceFeeDelivered * ZOMATO_FEE_GST_MULTIPLIER;
    const convenienceFee = d.convenienceFeeDelivered * ZOMATO_FEE_GST_MULTIPLIER;
    const discountOnServiceFee = d.discountOnServiceFeeDelivered * ZOMATO_FEE_GST_MULTIPLIER;
    const longDistanceFee = (d.longDistanceFeeDelivered - d.discountOnLongDistanceDelivered) * ZOMATO_FEE_GST_MULTIPLIER;
    const otherChargesGrossed = convenienceFee + longDistanceFee - discountOnServiceFee;
    const taxAdjustments = d.tdsAll + d.tcsAll + d.gstPaidByZomatoAll; // used for the total deduction math (unchanged)
    const tdsAndTcsOnly = d.tdsAll + d.tcsAll; // matches the reference sheet's "Tax Adjustments" line, which excludes GST-paid-by-Zomato (that sits inside Profit instead)
    const otherDeductions = otherChargesGrossed + taxAdjustments + d.marketingAds + d.hyperpure;
    const profitFromZomato = Math.max(
      0,
      netSalesInclGst - commission - otherChargesGrossed - d.marketingAds - d.hyperpure - d.gstPaidByZomatoAll
    );

    const netPayout = Math.max(0, netSalesExclGst + gstCollected - commission - otherDeductions);

    // Summary/Profit-statement-sheet derived metrics (match the reference
    // tool's own definitions, verified against real settlement data):
    //  - Commissionable Amount = item sales minus discount (pre-packaging/GST)
    //  - "Total Income" (Profit statement) = item sales + packaging + GST
    //    collected − discount − GST paid by Zomato
    //  - "Net Sales (Numerize)" = that Total Income + cancellation compensation
    const commissionableAmount = d.itemTotalDelivered - discounts;
    const totalIncomeProfitStmt = d.itemTotalDelivered + d.packagingDelivered + gstCollected - discounts - d.gstPaidByZomatoAll;
    const netSalesNumerize = totalIncomeProfitStmt + d.compensationCancelled;
    const otherChargesAndTaxAdj = otherChargesGrossed + taxAdjustments; // Summary sheet's combined line (excludes marketing/hyperpure, which are separate rows)

    // Bank match
    const bankCheck = matchBankPayout(netPayout, bankTxs);

    return {
      weekNum: w.weekNum,
      label: w.label,
      orders: d.orders,
      // Cashflow-sheet detail (used by the full workbook export; also
      // available for the on-screen table).
      itemSales: d.itemTotalDelivered,
      packagingCharges: d.packagingDelivered,
      compensation: d.compensationCancelled,
      discounts,
      gstCollected,
      netSales: netSalesInclGst,
      grossSales,
      netSalesExclGst,
      commission,
      convenienceFee,
      discountOnServiceFee,
      longDistanceFee,
      otherChargesGrossed,
      profitFromZomato,
      tds: d.tdsAll,
      tcs: d.tcsAll,
      gstPaidByZomato: d.gstPaidByZomatoAll,
      taxAdjustments,
      tdsAndTcsOnly,
      marketingAds: d.marketingAds,
      hyperpure: d.hyperpure,
      otherDeductions,
      commissionableAmount,
      totalIncomeProfitStmt,
      netSalesNumerize,
      otherChargesAndTaxAdj,
      expectedPayout: netPayout,
      bankActual: bankCheck.matched ? bankCheck.actual : 0,
      bankDiff: bankCheck.matched ? bankCheck.diff : -netPayout,
      bankMatched: bankCheck.matched,
      utr: d.utrList.join(", ") || "—",
    };
  });

  // Calculate Totals
  const sum = (key) => weeks.reduce((s, w) => s + w[key], 0);
  const total = {
    label: "Total",
    orders: sum("orders"),
    itemSales: sum("itemSales"),
    packagingCharges: sum("packagingCharges"),
    compensation: sum("compensation"),
    discounts: sum("discounts"),
    gstCollected: sum("gstCollected"),
    netSales: sum("netSales"),
    grossSales: sum("grossSales"),
    netSalesExclGst: sum("netSalesExclGst"),
    commission: sum("commission"),
    convenienceFee: sum("convenienceFee"),
    discountOnServiceFee: sum("discountOnServiceFee"),
    longDistanceFee: sum("longDistanceFee"),
    otherChargesGrossed: sum("otherChargesGrossed"),
    profitFromZomato: sum("profitFromZomato"),
    tds: sum("tds"),
    tcs: sum("tcs"),
    gstPaidByZomato: sum("gstPaidByZomato"),
    taxAdjustments: sum("taxAdjustments"),
    tdsAndTcsOnly: sum("tdsAndTcsOnly"),
    marketingAds: sum("marketingAds"),
    hyperpure: sum("hyperpure"),
    otherDeductions: sum("otherDeductions"),
    commissionableAmount: sum("commissionableAmount"),
    totalIncomeProfitStmt: sum("totalIncomeProfitStmt"),
    netSalesNumerize: sum("netSalesNumerize"),
    otherChargesAndTaxAdj: sum("otherChargesAndTaxAdj"),
    expectedPayout: sum("expectedPayout"),
    bankActual: sum("bankActual"),
    bankDiff: sum("bankDiff"),
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

  // Swiggy fee columns in their settlement report are already GST-inclusive
  // (unlike Zomato where the raw column values are pre-GST and need × 1.18).

  const weekDataMap = {};
  weekRanges.forEach((w) => {
    weekDataMap[w.weekNum] = {
      week: w,
      orders: 0,
      itemTotal: 0,
      packagingCharges: 0,
      discounts: 0,
      gstCollected: 0,
      platformFee: 0,       // Commission — already incl. 18% GST in Swiggy report
      callCenterFee: 0,
      swiggyOneFee: 0,
      pocketHeroFee: 0,
      longDistanceFee: 0,
      collectionCharges: 0,
      // Net compensation: total customer paid − complaint & cancellation charges
      compensationCancelled: 0,
      cancellationCharges: 0,
      customerComplaints: 0,   // Paid by restaurant
      tds: 0,
      tcs: 0,
      gstDeduction: 0,         // GST collected & paid by Swiggy (≡ gstPaidByZomato)
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

        const status = getRowString(r, ["order status", "status"]).toUpperCase();
        const isDelivered = status === "DELIVERED" || status === "COMPLETED" || status === "";
        const isCancelled = status.includes("CANCEL");

        if (isDelivered || isCancelled) {
          wData.orders += 1;
        } else if (!status) {
          wData.orders += 1; // count rows with no status field
        }

        // All order rows contribute to totals (matches Swiggy's own aggregation)
        wData.itemTotal += getRowVal(r, ["item total", "order total", "bill amount"]);
        wData.packagingCharges += getRowVal(r, ["packaging charges", "packing charges", "packing charge"]);
        wData.discounts += getRowVal(r, [
          "restaurant discounts",
          "restaurant discount",
          "swiggy one exclusive offer discount",
          "discount",
        ]);
        wData.gstCollected += getRowVal(r, ["gst collected", "gst"]);

        // Fee columns — incl. GST in Swiggy format, no gross-up needed
        wData.platformFee += getRowVal(r, ["commission", "platform fee"]);
        wData.callCenterFee += getRowVal(r, ["call center charges", "call center fees"]);
        wData.swiggyOneFee += getRowVal(r, ["swiggy one fees"]);
        wData.pocketHeroFee += getRowVal(r, ["pocket hero fees"]);
        wData.longDistanceFee += getRowVal(r, ["long distance charges"]);
        wData.collectionCharges += getRowVal(r, ["payment collection charges", "collection charges"]);
        wData.cancellationCharges += getRowVal(r, ["restaurant cancellation charges"]);
        wData.customerComplaints += getRowVal(r, ["customer complaints"]);

        // Net compensation: Total Customer Paid − Complaint & Cancellation Charges
        const totalCustomerPaid = getRowVal(r, ["total customer paid"]);
        const complaintCancellation = getRowVal(r, ["complaint & cancellation charges", "complaint cancellation"]);
        wData.compensationCancelled += (totalCustomerPaid - complaintCancellation);

        // Tax lines
        wData.tds += getRowVal(r, ["tds"]);
        wData.tcs += getRowVal(r, ["tcs"]);
        wData.gstDeduction += getRowVal(r, ["gst deduction", "gst liability", "gst paid by swiggy"]);

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

    // Income
    const grossSales = d.itemTotal + d.packagingCharges;
    const discounts = d.discounts;
    const netSalesExclGst = Math.max(0, grossSales - discounts);
    const gstCollected = d.gstCollected;
    const netSales = netSalesExclGst + gstCollected; // = Total Income on Swiggy dashboard
    const compensation = d.compensationCancelled;

    // Commission — all platform fees (already GST inclusive in Swiggy report)
    const commission = d.platformFee + d.callCenterFee + d.swiggyOneFee + d.pocketHeroFee + d.longDistanceFee + d.collectionCharges;

    // Deductions
    const cancellationCharges = d.cancellationCharges;
    const customerComplaints = d.customerComplaints;
    const tds = d.tds;
    const tcs = d.tcs;
    const gstPaidBySwiggy = d.gstDeduction;
    const taxAdjustments = tds + tcs;
    const otherDeductions = cancellationCharges + customerComplaints + taxAdjustments + gstPaidBySwiggy + d.marketingAds;

    // Profit from Swiggy (mirrors Zomato's profitFromZomato logic)
    const profitFromSwiggy = Math.max(
      0,
      netSales + compensation - commission - customerComplaints - cancellationCharges - gstPaidBySwiggy - d.marketingAds
    );

    // Commissionable amount = item sales − discounts
    const commissionableAmount = d.itemTotal - discounts;

    // Total Income for profit statement (mirrors reference tool's definition)
    const totalIncomeProfitStmt = d.itemTotal + d.packagingCharges + gstCollected - discounts - gstPaidBySwiggy;
    const netSalesNumerize = totalIncomeProfitStmt + compensation;

    // Expected payout
    const expectedPayout = Math.max(0, netSales + compensation - commission - otherDeductions);

    const bankCheck = matchBankPayout(expectedPayout, bankTxs);

    return {
      weekNum: w.weekNum,
      label: w.label,
      orders: d.orders,
      // Detailed fields for the 4-sheet workbook export
      itemSales: d.itemTotal,
      packagingCharges: d.packagingCharges,
      compensation,
      discounts,
      gstCollected,
      grossSales,
      netSalesExclGst,
      netSales,
      commission,
      platformFee: d.platformFee,
      callCenterFee: d.callCenterFee,
      swiggyOneFee: d.swiggyOneFee,
      pocketHeroFee: d.pocketHeroFee,
      longDistanceFee: d.longDistanceFee,
      collectionCharges: d.collectionCharges,
      cancellationCharges,
      customerComplaints,
      tds,
      tcs,
      gstPaidBySwiggy,
      taxAdjustments,
      otherDeductions,
      profitFromSwiggy,
      commissionableAmount,
      totalIncomeProfitStmt,
      netSalesNumerize,
      marketingAds: d.marketingAds,
      expectedPayout,
      bankActual: bankCheck.matched ? bankCheck.actual : 0,
      bankDiff: bankCheck.matched ? bankCheck.diff : -expectedPayout,
      bankMatched: bankCheck.matched,
      utr: d.utrList.join(", ") || "—",
    };
  });

  const sum = (key) => weeks.reduce((s, w) => s + (w[key] || 0), 0);
  const total = {
    label: "Total",
    orders: sum("orders"),
    itemSales: sum("itemSales"),
    packagingCharges: sum("packagingCharges"),
    compensation: sum("compensation"),
    discounts: sum("discounts"),
    gstCollected: sum("gstCollected"),
    grossSales: sum("grossSales"),
    netSalesExclGst: sum("netSalesExclGst"),
    netSales: sum("netSales"),
    commission: sum("commission"),
    platformFee: sum("platformFee"),
    callCenterFee: sum("callCenterFee"),
    swiggyOneFee: sum("swiggyOneFee"),
    pocketHeroFee: sum("pocketHeroFee"),
    longDistanceFee: sum("longDistanceFee"),
    collectionCharges: sum("collectionCharges"),
    cancellationCharges: sum("cancellationCharges"),
    customerComplaints: sum("customerComplaints"),
    tds: sum("tds"),
    tcs: sum("tcs"),
    gstPaidBySwiggy: sum("gstPaidBySwiggy"),
    taxAdjustments: sum("taxAdjustments"),
    otherDeductions: sum("otherDeductions"),
    profitFromSwiggy: sum("profitFromSwiggy"),
    commissionableAmount: sum("commissionableAmount"),
    totalIncomeProfitStmt: sum("totalIncomeProfitStmt"),
    netSalesNumerize: sum("netSalesNumerize"),
    marketingAds: sum("marketingAds"),
    expectedPayout: sum("expectedPayout"),
    bankActual: sum("bankActual"),
    bankDiff: sum("bankDiff"),
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

  // Dineout settlement report: Order Total is incl. 5% GST.
  // Per Python reference: orderTotalExclGst = orderTotalInclGst × (100/105)
  const weekDataMap = {};
  weekRanges.forEach((w) => {
    weekDataMap[w.weekNum] = {
      week: w,
      orderTotalInclGst: 0,   // Raw Order Total from settlement (incl. 5% GST)
      discountInclGst: 0,     // Total merchant discount (incl. GST)
      serviceFee: 0,
      discountOnServiceFee: 0,
      collectionCharges: 0,
      ads: 0,
      tips: 0,
      tds: 0,
      tcs: 0,
    };
  });

  for (const file of files) {
    try {
      const wb = await readWorkbookFromFile(file);
      const sheet = findSheetByPattern(wb, ["dineout", "orders", "sheet1", "summary"]);

      let headerRowIdx = 1;
      const rawGrid = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "" });
      for (let i = 0; i < Math.min(15, rawGrid.length); i++) {
        const rowStr = (rawGrid[i] || []).join(" ").toLowerCase();
        if (rowStr.includes("order total") || rowStr.includes("bill amount") || rowStr.includes("date")) {
          headerRowIdx = i + 1;
          break;
        }
      }
      const rows = sheetToRows(sheet, headerRowIdx);

      for (const r of rows) {
        const dateStr = getRowString(r, ["date", "transaction date", "bill date", "order date"]);
        const day = parseDayFromDate(dateStr) || parseDayFromDate(file.name) || 1;
        const targetWeek = findWeekForDay(day, weekRanges);
        const wData = weekDataMap[targetWeek.weekNum];

        // Order Total in Dineout = incl. GST; we extract excl-GST via ÷ 1.05 below
        wData.orderTotalInclGst += getRowVal(r, ["order total", "bill amount", "gross amount", "subtotal", "sales"]);
        wData.discountInclGst += getRowVal(r, ["total merchant discount", "merchant discount", "discount", "dineout discount"]);
        wData.serviceFee += getRowVal(r, ["platform service fee", "commission", "service fee"]);
        wData.discountOnServiceFee += getRowVal(r, ["discount on swiggy service fee", "discount on service fee", "service fee discount"]);
        wData.collectionCharges += getRowVal(r, ["collection charges", "payment gateway fee", "payment collection charges"]);
        wData.tips += getRowVal(r, ["tip", "tips"]);
        wData.ads += getRowVal(r, ["ads", "advertising", "banner ads"]);
        wData.tds += getRowVal(r, ["tds"]);
        wData.tcs += getRowVal(r, ["tcs"]);
      }
    } catch (e) {
      console.warn("Error parsing Dineout file:", file.name, e);
    }
  }

  const weeks = weekRanges.map((w) => {
    const d = weekDataMap[w.weekNum];

    // Per Python reference: order_total * 100/105 = excl-GST amount
    const orderTotalExclGst = d.orderTotalInclGst * (100.0 / 105.0);
    const discountExclGst = d.discountInclGst * (100.0 / 105.0);
    const salesAfterDiscounts = Math.max(0, orderTotalExclGst - discountExclGst);
    const gst5Pct = salesAfterDiscounts * 0.05;
    const salesInclGst = salesAfterDiscounts + gst5Pct + d.tips;

    // Commission = service fee − discount on service fee + collection charges
    const commission = Math.max(0, d.serviceFee - d.discountOnServiceFee + d.collectionCharges);
    const tds = d.tds;
    const tcs = d.tcs;
    const tdsTcs = tds + tcs;

    // Profit and expected payout
    const profitFromDineout = Math.max(0, salesInclGst - commission - d.ads - tdsTcs);
    const expectedPayout = profitFromDineout;

    const bankCheck = matchBankPayout(expectedPayout, bankTxs);

    return {
      weekNum: w.weekNum,
      label: w.label,
      orderTotalInclGst: d.orderTotalInclGst,
      orderTotalExclGst,
      discountExclGst,
      salesAfterDiscounts,
      salesExclGst: salesAfterDiscounts,   // alias for generic table display
      gst5Pct,
      tips: d.tips,
      salesInclGst,
      serviceFee: d.serviceFee,
      discountOnServiceFee: d.discountOnServiceFee,
      collectionCharges: d.collectionCharges,
      commission,
      ads: d.ads,
      tds,
      tcs,
      tdsTcs,
      profitFromDineout,
      expectedPayout,
      bankActual: bankCheck.matched ? bankCheck.actual : 0,
      bankDiff: bankCheck.matched ? bankCheck.diff : -expectedPayout,
      bankMatched: bankCheck.matched,
    };
  });

  const sum = (key) => weeks.reduce((s, w) => s + (w[key] || 0), 0);
  const total = {
    label: "Total",
    orderTotalInclGst: sum("orderTotalInclGst"),
    orderTotalExclGst: sum("orderTotalExclGst"),
    discountExclGst: sum("discountExclGst"),
    salesAfterDiscounts: sum("salesAfterDiscounts"),
    salesExclGst: sum("salesAfterDiscounts"),
    gst5Pct: sum("gst5Pct"),
    tips: sum("tips"),
    salesInclGst: sum("salesInclGst"),
    serviceFee: sum("serviceFee"),
    discountOnServiceFee: sum("discountOnServiceFee"),
    collectionCharges: sum("collectionCharges"),
    commission: sum("commission"),
    ads: sum("ads"),
    tds: sum("tds"),
    tcs: sum("tcs"),
    tdsTcs: sum("tdsTcs"),
    profitFromDineout: sum("profitFromDineout"),
    expectedPayout: sum("expectedPayout"),
    bankActual: sum("bankActual"),
    bankDiff: sum("bankDiff"),
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

  // Zomato Pay: Bill Amount in settlement report is incl. 5% GST.
  // Per Python reference: salesExclGst = billAmount × (100/105)
  // Commission is pre-GST; gross-up × 1.18 before deducting.
  const weekDataMap = {};
  weekRanges.forEach((w) => {
    weekDataMap[w.weekNum] = {
      week: w,
      billAmountInclGst: 0,   // Raw Bill Amount from settlement (incl. 5% GST)
      discountsInclGst: 0,    // Discounts/promos (incl. GST)
      failedTransactions: 0,
      tips: 0,
      commission: 0,          // Pre-GST in Zomato Pay report; grossed up × 1.18
      ads: 0,
    };
  });

  for (const file of files) {
    try {
      const wb = await readWorkbookFromFile(file);
      // Zomato Pay uses a "Transactions summary" sheet for order data
      const sheet = findSheetByPattern(wb, ["transactions summary", "zomato pay", "calculations", "orders", "sheet1"]);

      let headerRowIdx = 1;
      const rawGrid = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "" });
      for (let i = 0; i < Math.min(10, rawGrid.length); i++) {
        const rowStr = (rawGrid[i] || []).join(" ").toLowerCase();
        if (rowStr.includes("date") && (rowStr.includes("bill") || rowStr.includes("amount"))) {
          headerRowIdx = i + 1;
          break;
        }
      }
      const rows = sheetToRows(sheet, headerRowIdx);

      for (const r of rows) {
        const dateStr = getRowString(r, ["date", "transaction date", "payment date", "order date"]);
        const day = parseDayFromDate(dateStr) || parseDayFromDate(file.name) || 1;
        const targetWeek = findWeekForDay(day, weekRanges);
        const wData = weekDataMap[targetWeek.weekNum];

        wData.billAmountInclGst += getRowVal(r, ["bill amount", "order amount", "gross amount", "amount", "total amount"]);
        wData.discountsInclGst += getRowVal(r, ["discount", "promo", "offers", "zomato discount"]);
        wData.failedTransactions += getRowVal(r, ["failed", "reversed", "refund", "failed amount"]);
        wData.tips += getRowVal(r, ["tip", "tips"]);
        wData.commission += getRowVal(r, ["commission", "service fee", "platform fee", "convenience fee"]);
      }
    } catch (e) {
      console.warn("Error parsing Zomato Pay file:", file.name, e);
    }
  }

  // Process Ads from "Additions & deductions" sheet (separate file or second sheet)
  for (const file of adsFiles) {
    try {
      const wb = await readWorkbookFromFile(file);
      const sheet = findSheetByPattern(wb, ["additions & deductions", "additions deductions", "ads", "deductions"]);
      const rows = sheetToRows(sheet, 1);
      for (const r of rows) {
        const dateStr = getRowString(r, ["date"]);
        const day = parseDayFromDate(dateStr) || 1;
        const targetWeek = findWeekForDay(day, weekRanges);
        weekDataMap[targetWeek.weekNum].ads += getRowVal(r, ["amount", "ad spend", "cost", "debit"]);
      }
    } catch (e) {
      console.warn("Error parsing Zomato Pay Ads file:", file.name, e);
    }
  }

  const weeks = weekRanges.map((w) => {
    const d = weekDataMap[w.weekNum];
    // Per Python reference: billAmount * 100/105 = excl-GST
    const salesExclGstBefore = d.billAmountInclGst * (100.0 / 105.0);
    const discountsExclGst = d.discountsInclGst * (100.0 / 105.0);
    const salesExclGstAfter = Math.max(0, salesExclGstBefore - discountsExclGst - d.failedTransactions);
    const gst5Pct = salesExclGstAfter * 0.05;
    const salesInclGst = salesExclGstAfter + gst5Pct + d.tips;
    // Commission is pre-GST in Zomato Pay settlement; gross up × 1.18
    const commissionInclGst = d.commission * 1.18;
    const profitFromZomatoPay = Math.max(0, salesInclGst - commissionInclGst - d.ads);
    const expectedPayout = profitFromZomatoPay;

    const bankCheck = matchBankPayout(expectedPayout, bankTxs);

    return {
      weekNum: w.weekNum,
      label: w.label,
      billAmountInclGst: d.billAmountInclGst,
      salesExclGstBefore,
      discounts: discountsExclGst,
      failedTransactions: d.failedTransactions,
      salesExclGstAfter,
      gst5Pct,
      tips: d.tips,
      salesInclGst,
      commission: d.commission,
      commissionInclGst,
      ads: d.ads,
      profitFromZomatoPay,
      expectedPayout,
      bankActual: bankCheck.matched ? bankCheck.actual : 0,
      bankDiff: bankCheck.matched ? bankCheck.diff : -expectedPayout,
      bankMatched: bankCheck.matched,
    };
  });

  const sum = (key) => weeks.reduce((s, w) => s + (w[key] || 0), 0);
  const total = {
    label: "Total",
    billAmountInclGst: sum("billAmountInclGst"),
    salesExclGstBefore: sum("salesExclGstBefore"),
    discounts: sum("discounts"),
    failedTransactions: sum("failedTransactions"),
    salesExclGstAfter: sum("salesExclGstAfter"),
    gst5Pct: sum("gst5Pct"),
    tips: sum("tips"),
    salesInclGst: sum("salesInclGst"),
    commission: sum("commission"),
    commissionInclGst: sum("commissionInclGst"),
    ads: sum("ads"),
    profitFromZomatoPay: sum("profitFromZomatoPay"),
    expectedPayout: sum("expectedPayout"),
    bankActual: sum("bankActual"),
    bankDiff: sum("bankDiff"),
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

  // Paytm: Sales are excl. GST in the settlement report.
  // MDR (commission) is pre-GST; gross-up × 1.18 before deducting.
  const weekDataMap = {};
  weekRanges.forEach((w) => {
    weekDataMap[w.weekNum] = {
      week: w,
      salesExclGst: 0,        // Gross sales (excl. 5% GST)
      failedTransactions: 0,  // Refunds / chargebacks / failed
      commission: 0,          // MDR (pre-GST); will be grossed up × 1.18
      netSettlement: 0,
    };
  });

  for (const file of files) {
    try {
      const wb = await readWorkbookFromFile(file);
      const sheet = findSheetByPattern(wb, ["paytm", "settlement", "transactions", "sheet1"]);

      let headerRowIdx = 1;
      const rawGrid = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "" });
      for (let i = 0; i < Math.min(10, rawGrid.length); i++) {
        const rowStr = (rawGrid[i] || []).join(" ").toLowerCase();
        if (rowStr.includes("amount") || rowStr.includes("date") || rowStr.includes("transaction")) {
          headerRowIdx = i + 1;
          break;
        }
      }
      const rows = sheetToRows(sheet, headerRowIdx);

      for (const r of rows) {
        const dateStr = getRowString(r, ["date", "transaction date", "settlement date", "txn date"]);
        const day = parseDayFromDate(dateStr) || parseDayFromDate(file.name) || 1;
        const targetWeek = findWeekForDay(day, weekRanges);
        const wData = weekDataMap[targetWeek.weekNum];

        wData.salesExclGst += getRowVal(r, ["amount", "order amount", "gross amount", "total", "transaction amount"]);
        wData.failedTransactions += getRowVal(r, ["failed", "refund", "chargeback", "reversal"]);
        wData.commission += getRowVal(r, ["fee", "commission", "mdr", "charges", "paytm charges"]);
        wData.netSettlement += getRowVal(r, ["settlement amount", "net payout", "net amount", "net settlement"]);
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
    // MDR is pre-GST in Paytm settlement report; gross up × 1.18
    const commissionInclGst = d.commission * 1.18;
    const profitFromPaytm = Math.max(0, salesInclGst - commissionInclGst);
    const expectedReceipt = profitFromPaytm;

    const bankCheck = matchBankPayout(expectedReceipt, bankTxs);

    return {
      weekNum: w.weekNum,
      label: w.label,
      salesExclGst: d.salesExclGst,
      failedTransactions: d.failedTransactions,
      salesAfterFailed,
      gst5Pct,
      salesInclGst,
      commission: d.commission,
      commissionInclGst,
      profitFromPaytm,
      expectedPayout: expectedReceipt,
      expectedReceipt,
      bankActual: bankCheck.matched ? bankCheck.actual : 0,
      bankDiff: bankCheck.matched ? bankCheck.diff : -expectedReceipt,
      bankMatched: bankCheck.matched,
    };
  });

  const sum = (key) => weeks.reduce((s, w) => s + (w[key] || 0), 0);
  const total = {
    label: "Total",
    salesExclGst: sum("salesExclGst"),
    failedTransactions: sum("failedTransactions"),
    salesAfterFailed: sum("salesAfterFailed"),
    gst5Pct: sum("gst5Pct"),
    salesInclGst: sum("salesInclGst"),
    commission: sum("commission"),
    commissionInclGst: sum("commissionInclGst"),
    profitFromPaytm: sum("profitFromPaytm"),
    expectedPayout: sum("expectedReceipt"),
    expectedReceipt: sum("expectedReceipt"),
    bankActual: sum("bankActual"),
    bankDiff: sum("bankDiff"),
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

// ═════════════════════════════════════════════════════════════════════════════
// SHARED LINE-ITEM ROW DEFINITIONS — used by both the on-screen breakdown
// tables (page.js) and the downloadable multi-sheet workbook below, so the
// two views can never drift out of sync.
// Row structure: [label, getter(weekOrTotal), kind]
// kind = "add" | "less" | "subtotal" | "total" | "variance" | "percent" | "plain" | "section"
// ═════════════════════════════════════════════════════════════════════════════
const r2 = (v) => Math.round((v || 0) * 100) / 100;

// ── Platform detection guards ─────────────────────────────────────────────────
export function isZomatoDetailed(report) {
  return report.platform === "Zomato" && report.weeks.length > 0 && "itemSales" in report.weeks[0];
}
export function isSwiggyDetailed(report) {
  return report.platform === "Swiggy" && report.weeks.length > 0 && "platformFee" in report.weeks[0];
}
export function isDineoutDetailed(report) {
  return report.platform === "Swiggy Dineout" && report.weeks.length > 0 && "orderTotalExclGst" in report.weeks[0];
}
export function isZomatoPayDetailed(report) {
  return report.platform === "Zomato Pay" && report.weeks.length > 0 && "salesExclGstBefore" in report.weeks[0];
}
export function isPaytmDetailed(report) {
  return report.platform === "Paytm" && report.weeks.length > 0 && "salesAfterFailed" in report.weeks[0];
}

// Rows that are always 0 (manual-entry placeholders kept for structural parity)
const zero = () => 0;

// Each row: [label, getter(weekOrTotal), kind]
export function getSummaryRowDefs(report) {
  if (isZomatoDetailed(report)) {
    return [
      ["No. of Orders", (x) => x.orders, "plain"],
      ["Total Income (Number you'd see on Zomato dashboard)", (x) => x.netSales, "plain"],
      ["Avg. Order Value", (x) => (x.orders ? x.netSales / x.orders : 0), "plain"],
      ["No. of Disputed Orders (manual entry)", zero, "plain"],
      ["Refund Amount for Disputed Orders", zero, "plain"],
      ["Discounts", (x) => x.discounts, "less"],
      ["Commissionable Amount", (x) => x.commissionableAmount, "plain"],
      ["Commission", (x) => x.commission, "less"],
      ["Marketing & Ads", (x) => x.marketingAds, "less"],
      ["Other Charges & Tax Adjustments", (x) => x.otherChargesAndTaxAdj, "less"],
      ["Zomato Hyperpure Adjustments", (x) => x.hyperpure, "less"],
      ["Money Received (Expected)", (x) => x.expectedPayout, "total"],
    ];
  }
  if (isSwiggyDetailed(report)) {
    return [
      ["No. of Orders", (x) => x.orders, "plain"],
      ["Total Income (Swiggy Dashboard)", (x) => x.netSales, "plain"],
      ["Avg. Order Value", (x) => (x.orders ? x.netSales / x.orders : 0), "plain"],
      ["Discounts", (x) => x.discounts, "less"],
      ["Commissionable Amount", (x) => x.commissionableAmount, "plain"],
      ["Commission (Platform + Other Fees, incl. GST)", (x) => x.commission, "less"],
      ["Customer Complaints & Cancellation Charges", (x) => (x.customerComplaints || 0) + (x.cancellationCharges || 0), "less"],
      ["Tax Adjustments (TDS + TCS)", (x) => x.taxAdjustments, "less"],
      ["GST collected & paid by Swiggy", (x) => x.gstPaidBySwiggy, "less"],
      ["Marketing & Ads", (x) => x.marketingAds, "less"],
      ["Money Received (Expected)", (x) => x.expectedPayout, "total"],
    ];
  }
  if (isDineoutDetailed(report)) {
    return [
      ["Order Total (Incl. 5% GST)", (x) => x.orderTotalInclGst, "plain"],
      ["Order Total (Excl. GST, ÷1.05)", (x) => x.orderTotalExclGst, "plain"],
      ["Less:- Discounts (Excl. GST)", (x) => x.discountExclGst, "less"],
      ["Sales After Discounts", (x) => x.salesAfterDiscounts, "subtotal"],
      ["Add:- GST 5%", (x) => x.gst5Pct, "add"],
      ["Add:- Tips", (x) => x.tips, "add"],
      ["Sales Incl. GST", (x) => x.salesInclGst, "subtotal"],
      ["Less:- Commission (Service Fee net)", (x) => x.commission, "less"],
      ["Less:- Ads", (x) => x.ads, "less"],
      ["Less:- TDS + TCS", (x) => x.tdsTcs, "less"],
      ["Expected Payout", (x) => x.expectedPayout, "total"],
    ];
  }
  if (isZomatoPayDetailed(report)) {
    return [
      ["Bill Amount (Incl. 5% GST)", (x) => x.billAmountInclGst, "plain"],
      ["Sales Excl. GST (Bill × 100/105)", (x) => x.salesExclGstBefore, "plain"],
      ["Less:- Discounts (Excl. GST)", (x) => x.discounts, "less"],
      ["Less:- Failed / Reversed Transactions", (x) => x.failedTransactions, "less"],
      ["Net Sales (Excl. GST)", (x) => x.salesExclGstAfter, "subtotal"],
      ["Add:- GST 5%", (x) => x.gst5Pct, "add"],
      ["Add:- Tips", (x) => x.tips, "add"],
      ["Sales Incl. GST", (x) => x.salesInclGst, "subtotal"],
      ["Less:- Commission (Incl. 18% GST)", (x) => x.commissionInclGst, "less"],
      ["Less:- Ads", (x) => x.ads, "less"],
      ["Expected Payout", (x) => x.expectedPayout, "total"],
    ];
  }
  if (isPaytmDetailed(report)) {
    return [
      ["Gross Sales (Excl. GST)", (x) => x.salesExclGst, "plain"],
      ["Less:- Failed / Refunded Transactions", (x) => x.failedTransactions, "less"],
      ["Net Sales (Excl. GST)", (x) => x.salesAfterFailed, "subtotal"],
      ["Add:- GST 5%", (x) => x.gst5Pct, "add"],
      ["Sales Incl. GST", (x) => x.salesInclGst, "subtotal"],
      ["Less:- MDR / Commission (Incl. 18% GST)", (x) => x.commissionInclGst, "less"],
      ["Expected Receipt", (x) => x.expectedReceipt, "total"],
    ];
  }
  return [
    ["No. of Orders", (x) => x.orders, "plain"],
    ["Gross Sales / Total Income", (x) => x.grossSales ?? x.totalIncome ?? x.salesInclGst ?? 0, "plain"],
    ["Discounts", (x) => x.discounts || 0, "less"],
    ["Commission", (x) => x.commission ?? x.commissionInclGst ?? 0, "less"],
    ["Other Deductions/Taxes", (x) => x.otherDeductions ?? x.taxesAndDeductions ?? 0, "less"],
    ["Money Received (Expected)", (x) => x.expectedPayout ?? x.expectedReceipt ?? 0, "total"],
  ];
}

export function getCashflowRowDefs(report) {
  if (isZomatoDetailed(report)) {
    return [
      ["Item sales (Delivered orders)", (x) => x.itemSales, "plain"],
      ["Add:- Packing charges", (x) => x.packagingCharges, "add"],
      ["Add:- Compensation paid for cancelled orders", (x) => x.compensation, "add"],
      ["Less:- Discount", (x) => -x.discounts, "less"],
      ["Add:- GST", (x) => x.gstCollected, "add"],
      ["Net Sales", (x) => x.netSales, "subtotal"],
      ["Less:- Commission (Platform Fee, incl. GST)", (x) => -x.commission, "less"],
      ["Less:- Other Charges", (x) => -(x.convenienceFee + x.longDistanceFee - x.discountOnServiceFee), "less"],
      ["   Convenience Fee", (x) => -x.convenienceFee, "less"],
      ["   Discount on Service Fee", (x) => x.discountOnServiceFee, "add"],
      ["   Long Distance Fee", (x) => -x.longDistanceFee, "less"],
      ["Merchant Cancellation Charges", zero, "plain"],
      ["Less:- Refunds", zero, "plain"],
      ["   Paid by Restaurant", zero, "plain"],
      ["Received by Restaurant", zero, "plain"],
      ["Less:- Marketing Charges", (x) => -x.marketingAds, "less"],
      ["   High Priority", zero, "plain"],
      ["   Call Center Service Fees", zero, "plain"],
      ["   PocketHero Fee", zero, "plain"],
      ["Profit from Zomato", (x) => x.profitFromZomato, "subtotal"],
      ["Less:- Tax Adjustments", (x) => -x.tdsAndTcsOnly, "less"],
      ["   TDS deduction for aggregators", (x) => -x.tds, "less"],
      ["   TCS", (x) => -x.tcs, "less"],
      ["Less:- GST collected and paid by Zomato", (x) => -x.gstPaidByZomato, "less"],
      ["Expected Receipts", (x) => x.expectedPayout, "total"],
      ["   Service Fee Reversal", zero, "plain"],
      ["   Interim payments", zero, "plain"],
      ["   Opening Week Adjustments", zero, "plain"],
      ["   Following Week Adjustments", zero, "plain"],
      ["   Closing Week Adjustments", zero, "plain"],
      ["   Zomato Hyperpure Payment Adjustments", (x) => -x.hyperpure, "less"],
      ["   EMI Loan deductions", zero, "plain"],
      ["Actual receipts", (x) => x.bankActual, "plain"],
      ["Difference (Expected − Actual)", (x) => x.expectedPayout - x.bankActual, "variance"],
    ];
  }
  if (isSwiggyDetailed(report)) {
    return [
      ["Item sales", (x) => x.itemSales, "plain"],
      ["Add:- Packing charges", (x) => x.packagingCharges, "add"],
      ["Add:- Compensation (net, cancelled orders)", (x) => x.compensation, "add"],
      ["Less:- Discount", (x) => -x.discounts, "less"],
      ["Add:- GST 5%", (x) => x.gstCollected, "add"],
      ["Net Sales", (x) => x.netSales, "subtotal"],
      ["Less:- Platform Fee (Commission, incl. GST)", (x) => -x.platformFee, "less"],
      ["   Swiggy One Fees", (x) => -(x.swiggyOneFee || 0), "less"],
      ["   Call Center Service Fees", (x) => -(x.callCenterFee || 0), "less"],
      ["   PocketHero Fee", (x) => -(x.pocketHeroFee || 0), "less"],
      ["   Long Distance Charges", (x) => -(x.longDistanceFee || 0), "less"],
      ["   Payment Collection Charges", (x) => -(x.collectionCharges || 0), "less"],
      ["Less:- Merchant Cancellation Charges", (x) => -(x.cancellationCharges || 0), "less"],
      ["Less:- Customer Complaints (Paid by restaurant)", (x) => -(x.customerComplaints || 0), "less"],
      ["Less:- Marketing & Ads", (x) => -(x.marketingAds || 0), "less"],
      ["Profit from Swiggy", (x) => x.profitFromSwiggy, "subtotal"],
      ["Less:- Tax Adjustments", (x) => -(x.taxAdjustments || 0), "less"],
      ["   TDS deduction for aggregators", (x) => -(x.tds || 0), "less"],
      ["   TCS", (x) => -(x.tcs || 0), "less"],
      ["Less:- GST collected and paid by Swiggy", (x) => -(x.gstPaidBySwiggy || 0), "less"],
      ["Expected Receipts", (x) => x.expectedPayout, "total"],
      ["Actual receipts", (x) => x.bankActual || 0, "plain"],
      ["Difference (Expected − Actual)", (x) => (x.expectedPayout || 0) - (x.bankActual || 0), "variance"],
    ];
  }
  if (isDineoutDetailed(report)) {
    return [
      ["Order Total (Incl. GST from Dineout report)", (x) => x.orderTotalInclGst, "plain"],
      ["Order Total (Excl. GST, ÷1.05)", (x) => x.orderTotalExclGst, "plain"],
      ["Less:- Total Merchant Discount (Excl. GST)", (x) => -(x.discountExclGst || 0), "less"],
      ["Net Sales (Excl. GST)", (x) => x.salesAfterDiscounts, "subtotal"],
      ["Add:- GST 5%", (x) => x.gst5Pct, "add"],
      ["Add:- Tips", (x) => x.tips || 0, "add"],
      ["Net Sales (Incl. GST)", (x) => x.salesInclGst, "subtotal"],
      ["Less:- Service Fee (Platform Commission)", (x) => -(x.serviceFee || 0), "less"],
      ["Add:- Discount on Service Fee", (x) => x.discountOnServiceFee || 0, "add"],
      ["Less:- Collection Charges", (x) => -(x.collectionCharges || 0), "less"],
      ["Less:- Ads", (x) => -(x.ads || 0), "less"],
      ["Profit from Dineout", (x) => x.profitFromDineout, "subtotal"],
      ["Less:- TDS + TCS", (x) => -(x.tdsTcs || 0), "less"],
      ["Expected Payout", (x) => x.expectedPayout, "total"],
      ["Actual receipts", (x) => x.bankActual || 0, "plain"],
      ["Difference (Expected − Actual)", (x) => (x.expectedPayout || 0) - (x.bankActual || 0), "variance"],
    ];
  }
  if (isZomatoPayDetailed(report)) {
    return [
      ["Bill Amount (Incl. 5% GST)", (x) => x.billAmountInclGst, "plain"],
      ["Sales Excl. GST (Bill × 100/105)", (x) => x.salesExclGstBefore, "plain"],
      ["Less:- Discounts (Excl. GST)", (x) => -(x.discounts || 0), "less"],
      ["Less:- Failed / Reversed Transactions", (x) => -(x.failedTransactions || 0), "less"],
      ["Net Sales (Excl. GST)", (x) => x.salesExclGstAfter, "subtotal"],
      ["Add:- GST 5%", (x) => x.gst5Pct, "add"],
      ["Add:- Tips", (x) => x.tips || 0, "add"],
      ["Net Sales (Incl. GST)", (x) => x.salesInclGst, "subtotal"],
      ["Less:- Commission (Incl. 18% GST)", (x) => -(x.commissionInclGst || 0), "less"],
      ["Less:- Ads", (x) => -(x.ads || 0), "less"],
      ["Profit / Expected Payout", (x) => x.expectedPayout, "total"],
      ["Actual receipts", (x) => x.bankActual || 0, "plain"],
      ["Difference (Expected − Actual)", (x) => (x.expectedPayout || 0) - (x.bankActual || 0), "variance"],
    ];
  }
  if (isPaytmDetailed(report)) {
    return [
      ["Gross Sales (Excl. GST)", (x) => x.salesExclGst, "plain"],
      ["Less:- Failed / Refunded Transactions", (x) => -(x.failedTransactions || 0), "less"],
      ["Net Sales (Excl. GST)", (x) => x.salesAfterFailed, "subtotal"],
      ["Add:- GST 5%", (x) => x.gst5Pct, "add"],
      ["Net Sales (Incl. GST)", (x) => x.salesInclGst, "subtotal"],
      ["Less:- MDR / Commission (Incl. 18% GST)", (x) => -(x.commissionInclGst || 0), "less"],
      ["Expected Receipt", (x) => x.expectedReceipt, "total"],
      ["Actual receipts", (x) => x.bankActual || 0, "plain"],
      ["Difference (Expected − Actual)", (x) => (x.expectedReceipt || 0) - (x.bankActual || 0), "variance"],
    ];
  }
  return [
    ["Gross Sales / Total Income", (x) => x.grossSales ?? x.totalIncome ?? x.salesInclGst ?? 0, "plain"],
    ["Less:- Discount", (x) => -(x.discounts || 0), "less"],
    ["Add:- GST", (x) => x.gstCollected ?? x.gst5Pct ?? 0, "add"],
    ["Less:- Commission", (x) => -(x.commission ?? x.commissionInclGst ?? 0), "less"],
    ["Less:- Other Deductions/Taxes", (x) => -(x.otherDeductions ?? x.taxesAndDeductions ?? 0), "less"],
    ["Expected Payout / Receipts", (x) => x.expectedPayout ?? x.expectedReceipt ?? 0, "total"],
    ["Bank Actual", (x) => x.bankActual || 0, "plain"],
    ["Variance", (x) => x.bankDiff || 0, "variance"],
  ];
}

export function getProfitRowDefs(report) {
  if (isZomatoDetailed(report)) {
    return [
      ["A. Net Sales", (x) => "", "section"],
      ["   Total Income", (x) => x.totalIncomeProfitStmt, "plain"],
      ["   Add:- Adjust: Cancelled orders (by Zomato)", (x) => x.compensation, "add"],
      ["   Net Sales (Number you'd see on Numerize app)", (x) => x.netSalesNumerize, "subtotal"],
      ["B. Less:- Commission", (x) => -x.commission, "less"],
      ["C. Less:- Marketing & Ads", (x) => -x.marketingAds, "less"],
      ["D. Less:- Refund on Disputed Orders (manual entry)", zero, "plain"],
      ["E. Other Charges", (x) => -x.otherChargesGrossed, "less"],
      ["F. Profit (Net Sales − B − C − D − E)", (x) => x.profitFromZomato, "total"],
      ["Percentage of Net Sales", (x) => (x.netSalesNumerize ? x.profitFromZomato / x.netSalesNumerize : 0), "percent"],
    ];
  }
  if (isSwiggyDetailed(report)) {
    return [
      ["A. Net Sales", (x) => "", "section"],
      ["   Total Income (Item + Packaging + GST − Discounts − GST paid by Swiggy)", (x) => x.totalIncomeProfitStmt, "plain"],
      ["   Add:- Compensation (cancelled orders, net)", (x) => x.compensation || 0, "add"],
      ["   Net Sales (Numerize)", (x) => x.netSalesNumerize, "subtotal"],
      ["B. Less:- Commission (all platform fees, incl. GST)", (x) => -(x.commission || 0), "less"],
      ["C. Less:- Marketing & Ads", (x) => -(x.marketingAds || 0), "less"],
      ["D. Less:- Customer Complaints & Cancellation Charges", (x) => -((x.customerComplaints || 0) + (x.cancellationCharges || 0)), "less"],
      ["E. Profit from Swiggy (A − B − C − D)", (x) => x.profitFromSwiggy, "total"],
      ["Percentage of Net Sales", (x) => (x.netSalesNumerize ? x.profitFromSwiggy / x.netSalesNumerize : 0), "percent"],
    ];
  }
  if (isDineoutDetailed(report)) {
    return [
      ["A. Sales (Excl. GST, after discounts)", (x) => x.salesAfterDiscounts, "plain"],
      ["B. Add:- GST 5%", (x) => x.gst5Pct, "add"],
      ["C. Add:- Tips", (x) => x.tips || 0, "add"],
      ["   Net Sales (Incl. GST)", (x) => x.salesInclGst, "subtotal"],
      ["D. Less:- Commission (Service Fee)", (x) => -(x.commission || 0), "less"],
      ["E. Less:- Ads", (x) => -(x.ads || 0), "less"],
      ["F. Less:- TDS + TCS", (x) => -(x.tdsTcs || 0), "less"],
      ["Profit from Dineout (A+B+C − D − E − F)", (x) => x.profitFromDineout, "total"],
      ["Percentage of Net Sales", (x) => (x.salesInclGst ? x.profitFromDineout / x.salesInclGst : 0), "percent"],
    ];
  }
  if (isZomatoPayDetailed(report)) {
    return [
      ["A. Net Sales Excl. GST (after discounts & failed)", (x) => x.salesExclGstAfter, "plain"],
      ["B. Add:- GST 5%", (x) => x.gst5Pct, "add"],
      ["C. Add:- Tips", (x) => x.tips || 0, "add"],
      ["   Net Sales (Incl. GST)", (x) => x.salesInclGst, "subtotal"],
      ["D. Less:- Commission (Incl. 18% GST)", (x) => -(x.commissionInclGst || 0), "less"],
      ["E. Less:- Ads", (x) => -(x.ads || 0), "less"],
      ["Profit from Zomato Pay", (x) => x.profitFromZomatoPay, "total"],
      ["Percentage of Net Sales", (x) => (x.salesInclGst ? x.profitFromZomatoPay / x.salesInclGst : 0), "percent"],
    ];
  }
  if (isPaytmDetailed(report)) {
    return [
      ["A. Gross Sales (Excl. GST)", (x) => x.salesExclGst, "plain"],
      ["B. Less:- Failed / Refunded Transactions", (x) => -(x.failedTransactions || 0), "less"],
      ["   Net Sales (Excl. GST)", (x) => x.salesAfterFailed, "subtotal"],
      ["C. Add:- GST 5%", (x) => x.gst5Pct, "add"],
      ["   Net Sales (Incl. GST)", (x) => x.salesInclGst, "subtotal"],
      ["D. Less:- MDR / Commission (Incl. 18% GST)", (x) => -(x.commissionInclGst || 0), "less"],
      ["Profit from Paytm", (x) => x.profitFromPaytm, "total"],
      ["Percentage of Net Sales", (x) => (x.salesInclGst ? x.profitFromPaytm / x.salesInclGst : 0), "percent"],
    ];
  }
  return [
    ["A. Net Sales (Gross − Discounts + GST)", (x) => (x.grossSales ?? x.totalIncome ?? x.salesInclGst ?? 0) - (x.discounts || 0) + (x.gstCollected ?? x.gst5Pct ?? 0), "plain"],
    ["B. Less:- Commission", (x) => -(x.commission ?? x.commissionInclGst ?? 0), "less"],
    ["C. Less:- Other Deductions/Taxes", (x) => -(x.otherDeductions ?? x.taxesAndDeductions ?? 0), "less"],
    ["Expected Payout (A − B − C)", (x) => x.expectedPayout ?? x.expectedReceipt ?? 0, "total"],
  ];
}

// Discrepancies: mirrors the original tool's A. POS / B. Order / C. Payment layout.
// Sign convention: Difference = Expected − Actual (positive = shortfall).
export function getDiscrepancyRowDefs(report) {
  if (isZomatoDetailed(report)) {
    return [
      ["A. POS Discrepancies", (x) => "", "section"],
      ["   Sales as per Zomato", (x) => x.netSales, "plain"],
      ["   Sale as per POS (not uploaded)", zero, "plain"],
      ["   Difference (POS − Zomato)", (x) => -x.netSales, "less"],
      ["B. Order Discrepancies", (x) => "", "section"],
      ["   No. of delivered orders unpaid (manual entry)", zero, "plain"],
      ["   Value of such orders (manual entry)", zero, "plain"],
      ["C. Payment Discrepancies", (x) => "", "section"],
      ["   Expected receipts in the bank account", (x) => x.expectedPayout, "plain"],
      ["   Investment in Hyperpure", (x) => -(x.hyperpure || 0), "less"],
      ["   Actual receipts in the bank", (x) => x.bankActual || 0, "plain"],
      ["   Difference (Expected − Actual)", (x) => (x.expectedPayout || 0) - (x.hyperpure || 0) - (x.bankActual || 0), "total"],
    ];
  }
  if (isSwiggyDetailed(report)) {
    return [
      ["A. POS Discrepancies", (x) => "", "section"],
      ["   Sales as per Swiggy", (x) => x.netSales, "plain"],
      ["   Sale as per POS (not uploaded)", zero, "plain"],
      ["   Difference (POS − Swiggy)", (x) => -(x.netSales || 0), "less"],
      ["B. Order Discrepancies", (x) => "", "section"],
      ["   No. of orders unpaid (manual entry)", zero, "plain"],
      ["   Value of such orders (manual entry)", zero, "plain"],
      ["C. Payment Discrepancies", (x) => "", "section"],
      ["   Expected receipts in the bank account", (x) => x.expectedPayout, "plain"],
      ["   Actual receipts in the bank", (x) => x.bankActual || 0, "plain"],
      ["   Difference (Expected − Actual)", (x) => (x.expectedPayout || 0) - (x.bankActual || 0), "total"],
    ];
  }
  if (isDineoutDetailed(report) || isZomatoPayDetailed(report) || isPaytmDetailed(report)) {
    const expGetter = (x) => x.expectedPayout ?? x.expectedReceipt ?? 0;
    return [
      ["A. POS Discrepancies", (x) => "", "section"],
      ["   Sales as per Platform", (x) => x.salesInclGst || 0, "plain"],
      ["   Sale as per POS (not uploaded)", zero, "plain"],
      ["   Difference (POS − Platform)", (x) => -(x.salesInclGst || 0), "less"],
      ["B. Payment Discrepancies", (x) => "", "section"],
      ["   Expected receipts in the bank account", expGetter, "plain"],
      ["   Actual receipts in the bank", (x) => x.bankActual || 0, "plain"],
      ["   Difference (Expected − Actual)", (x) => expGetter(x) - (x.bankActual || 0), "total"],
    ];
  }
  return [
    ["Expected Payout / Receipts", (x) => x.expectedPayout ?? x.expectedReceipt ?? 0, "plain"],
    ["Bank Actual", (x) => x.bankActual || 0, "plain"],
    ["Difference (Expected − Actual)", (x) => (x.expectedPayout ?? x.expectedReceipt ?? 0) - (x.bankActual || 0), "total"],
  ];
}

function buildLineRowsAoa(report, rowDefs) {
  const header = ["Details", ...report.weeks.map((w) => w.label), "Total"];
  const body = rowDefs.map(([label, getter]) => [
    label,
    ...report.weeks.map((w) => (typeof getter(w) === "number" ? r2(getter(w)) : getter(w))),
    typeof getter(report.total) === "number" ? r2(getter(report.total)) : getter(report.total),
  ]);
  return [header, ...body];
}

export function exportFullReconWorkbook(report) {
  const wb = XLSX.utils.book_new();
  const generated = new Date().toLocaleString("en-IN");

  // ── Summary sheet ──────────────────────────────────────────────────────
  const summaryRows = [
    [`${report.clientName} — ${report.platform} Summary (${report.month})`],
    ["Generated:", generated],
    [],
    ...buildLineRowsAoa(report, getSummaryRowDefs(report)),
    [],
    ["Bank Statement Attached?", report.hasBank ? "Yes" : "No"],
    ["Source Files Processed", report.filesCount],
  ];
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(summaryRows), "Summary");

  // ── Cashflow sheet ─────────────────────────────────────────────────────
  const cashflowRows = [
    [`${report.clientName} ${report.platform} Cash Flow — ${report.month}`],
    [],
    ...buildLineRowsAoa(report, getCashflowRowDefs(report)),
  ];
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(cashflowRows), "Cashflow");

  // ── Profit statement sheet ─────────────────────────────────────────────
  const profitRows = [
    [`${report.clientName} ${report.platform} Profit Statement — ${report.month}`],
    [],
    ...buildLineRowsAoa(report, getProfitRowDefs(report)),
  ];
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(profitRows), "Profit statement");

  // ── Discrepancies sheet ────────────────────────────────────────────────
  const discRows = [
    [`${report.clientName} ${report.platform} Discrepancies — ${report.month}`],
    [
      report.hasBank
        ? "Compares expected payout against the uploaded bank statement."
        : "No bank statement was uploaded — every week shows as unmatched (POS/bank rows default to 0). Upload one to check actual settlement vs expected payout.",
    ],
    [],
    ...buildLineRowsAoa(report, getDiscrepancyRowDefs(report)),
  ];
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(discRows), "Discrepancies");

  const fileName = `${report.clientName}_${report.platform.replace(/\s+/g, "_")}_Full_Report_${report.month}.xlsx`;
  XLSX.writeFile(wb, fileName);
}
