import * as XLSX from "xlsx";
import JSZip from "jszip";

// ─── numeric normaliser ───────────────────────────────────────────────────────
const n = (v) => {
  if (typeof v === "number") return v;
  const s = String(v ?? "").replace(/[,₹\s]/g, "").trim();
  return parseFloat(s) || 0;
};

// ─── fuzzy column finder ──────────────────────────────────────────────────────
// Given a row's keys and a list of keyword patterns, returns the first matching value
function findCol(keys, patterns) {
  const lower = keys.map((k) => ({ key: k, lc: k.toLowerCase() }));
  for (const pat of patterns) {
    const lp = pat.toLowerCase();
    const match = lower.find((k) => k.lc.includes(lp));
    if (match) return match.key;
  }
  return null;
}

function getVal(row, patterns) {
  const key = findCol(Object.keys(row), patterns);
  return key != null ? n(row[key]) : 0;
}

// ─── week label from day-of-month ────────────────────────────────────────────
function weekFromDay(day) {
  if (day <= 7) return "Week 1";
  if (day <= 14) return "Week 2";
  if (day <= 21) return "Week 3";
  if (day <= 28) return "Week 4";
  return "Week 5";
}

function weekLabelFromDate(dateStr) {
  try {
    // Handle Excel serial numbers
    if (typeof dateStr === "number") {
      const d = XLSX.SSF.parse_date_code(dateStr);
      return weekFromDay(d.d);
    }
    const d = new Date(dateStr);
    if (!isNaN(d)) return weekFromDay(d.getDate());
  } catch (_) {}
  return null;
}

// ─── Week label from filename ─────────────────────────────────────────────────
function weekLabelFromFilename(name) {
  const lc = name.toLowerCase();
  const m = lc.match(/week[-_\s]?(\d)/i) || lc.match(/\bw(\d)\b/i);
  if (m) return `Week ${m[1]}`;
  return null;
}

// ─────────────────────────────────────────────────────────────────────────────
// EXTRACTION
// ─────────────────────────────────────────────────────────────────────────────

/**
 * Recursively extract files from a zip (including nested zips).
 * Returns Array<{ name, path, data: ArrayBuffer }>
 */
export async function extractZipRecursively(file) {
  const results = [];

  async function processZip(source) {
    const zip = await JSZip.loadAsync(source);
    const entries = [];
    zip.forEach((path, entry) => {
      if (!entry.dir) entries.push({ path, entry });
    });

    await Promise.all(
      entries.map(async ({ path, entry }) => {
        const buf = await entry.async("arraybuffer");
        if (path.toLowerCase().endsWith(".zip")) {
          try {
            await processZip(buf);
          } catch (e) {
            console.warn("Nested zip failed:", path, e);
          }
        } else {
          results.push({ name: path.split("/").pop(), path, data: buf });
        }
      })
    );
  }

  await processZip(file);
  return results;
}

// ─────────────────────────────────────────────────────────────────────────────
// CLASSIFICATION
// ─────────────────────────────────────────────────────────────────────────────

/**
 * Classify files into pos / zomatoRaw / swiggyRaw / zomatoSummaries / swiggySummaries.
 * Strategy: filename keywords first, then sheet-name inspection, then first-row column sniff.
 */
export function classifyDeliverableFiles(files) {
  const out = {
    pos: [],
    zomatoRaw: [],
    swiggyRaw: [],
    swiggySummaries: [],
    zomatoSummaries: [],
    unknown: [],
  };

  for (const file of files) {
    const lc = file.name.toLowerCase();

    // ── Accountant summaries (filename contains "summary") ─────────────────
    if (lc.includes("swiggy") && lc.includes("summary")) {
      out.swiggySummaries.push(file); continue;
    }
    if (lc.includes("zomato") && lc.includes("summary")) {
      out.zomatoSummaries.push(file); continue;
    }

    // ── Non-spreadsheet files ──────────────────────────────────────────────
    const isExcel = lc.endsWith(".xlsx") || lc.endsWith(".xls");
    const isCsv   = lc.endsWith(".csv");
    if (!isExcel && !isCsv) { out.unknown.push(file); continue; }

    // ── Parse workbook ─────────────────────────────────────────────────────
    let wb;
    try {
      wb = XLSX.read(new Uint8Array(file.data), {
        type: "array",
        bookSheets: true,
        sheetRows: 5, // only read first 5 rows for classification speed
      });
    } catch (_) {
      out.unknown.push(file); continue;
    }

    const sheets = (wb.SheetNames || []).map((s) => s.toLowerCase());
    const hasSheet = (kw) => sheets.some((s) => s.includes(kw));

    // ── Sheet-name heuristics ──────────────────────────────────────────────
    if (hasSheet("order level") || (hasSheet("settlement") && lc.includes("swiggy"))) {
      out.swiggyRaw.push(file); continue;
    }
    if (hasSheet("addition deductions") || hasSheet("tax report") || (hasSheet("payout") && lc.includes("zomato"))) {
      out.zomatoRaw.push(file); continue;
    }
    if (hasSheet("payment wise") || hasSheet("petpooja") || hasSheet("posist") || lc.includes("petpooja")) {
      out.pos.push(file); continue;
    }

    // ── Filename keyword fallback ──────────────────────────────────────────
    if (lc.includes("zomato")) { out.zomatoRaw.push(file); continue; }
    if (lc.includes("swiggy")) { out.swiggyRaw.push(file); continue; }
    if (lc.includes("pos") || lc.includes("petpooja") || lc.includes("posist") || lc.includes("sale")) {
      out.pos.push(file); continue;
    }

    // ── Column-sniff the first sheet ──────────────────────────────────────
    try {
      const fullWb = XLSX.read(new Uint8Array(file.data), { type: "array", sheetRows: 10 });
      const sheet  = fullWb.Sheets[fullWb.SheetNames[0]];
      const rows   = XLSX.utils.sheet_to_json(sheet, { defval: "" });
      if (rows.length > 0) {
        const keys = Object.keys(rows[0]).map((k) => k.toLowerCase()).join(" ");
        if (keys.includes("zomato") || keys.includes("restaurant id") || keys.includes("tax report")) {
          out.zomatoRaw.push(file); continue;
        }
        if (keys.includes("swiggy") || keys.includes("order level") || keys.includes("settlement amount")) {
          out.swiggyRaw.push(file); continue;
        }
        if (keys.includes("order_type") || keys.includes("payment_type") || keys.includes("petpooja")) {
          out.pos.push(file); continue;
        }
      }
    } catch (_) {}

    out.unknown.push(file);
  }

  return out;
}

// ─────────────────────────────────────────────────────────────────────────────
// SWIGGY RAW PARSING
// ─────────────────────────────────────────────────────────────────────────────
// Swiggy weekly payout reports have a sheet called "Order Level" or similar.
// Columns (may vary slightly): Order_Date | Bill_Amount / Order_Amount | Commission | TDS | Net_Payout / Settlement_Amount

function parseSwiggyRaw(file) {
  const wb = XLSX.read(new Uint8Array(file.data), { type: "array" });
  const sheetName =
    wb.SheetNames.find((s) => /order/i.test(s) || /settlement/i.test(s) || /payout/i.test(s)) ||
    wb.SheetNames[0];
  const sheet = wb.Sheets[sheetName];
  const rows  = XLSX.utils.sheet_to_json(sheet, { defval: "" });

  const weekMap = new Map();

  for (const row of rows) {
    const keys = Object.keys(row);

    // ── Date ──────────────────────────────────────────────────────────────
    const dateCols = ["order_date", "date", "transaction_date", "order date", "settlement date"];
    const dateKey  = findCol(keys, dateCols);
    const dateRaw  = dateKey ? row[dateKey] : "";
    const week     = weekLabelFromDate(dateRaw) || weekLabelFromFilename(file.name) || "Week 1";

    // ── Sales (gross order value) ─────────────────────────────────────────
    const sales = getVal(row, [
      "order_total", "order total", "bill_amount", "bill amount",
      "gross_amount", "gross amount", "order_amount", "order amount",
      "total_amount", "total amount",
    ]);

    // ── Commission / deduction ────────────────────────────────────────────
    const commission = getVal(row, [
      "commission", "service_fee", "service fee", "platform_fee", "platform fee",
      "total_deduction", "total deduction", "deduction",
    ]);

    // ── Net payout ────────────────────────────────────────────────────────
    const payout = getVal(row, [
      "net_payout", "net payout", "settlement_amount", "settlement amount",
      "amount_paid", "amount paid", "payment_amount", "payment amount",
    ]);

    const prev = weekMap.get(week) || { sales: 0, commission: 0, netPayout: 0 };
    weekMap.set(week, {
      sales:      prev.sales      + sales,
      commission: prev.commission + commission,
      netPayout:  prev.netPayout  + payout,
    });
  }

  return weekMap;
}

// ─────────────────────────────────────────────────────────────────────────────
// ZOMATO RAW PARSING
// ─────────────────────────────────────────────────────────────────────────────
// Zomato settlement reports: sheet "Addition Deductions Details" or "Tax Report"
// Columns: Order_Date | Gross_Sales | Commission | Net_Settlement

function parseZomatoRaw(file) {
  const wb = XLSX.read(new Uint8Array(file.data), { type: "array" });
  const sheetName =
    wb.SheetNames.find((s) => /addition|deduction|order|settlement/i.test(s)) ||
    wb.SheetNames[0];
  const sheet = wb.Sheets[sheetName];
  const rows  = XLSX.utils.sheet_to_json(sheet, { defval: "" });

  const weekMap = new Map();

  for (const row of rows) {
    const keys = Object.keys(row);

    const dateKey = findCol(keys, ["order_date", "date", "transaction_date", "order date", "settlement_date"]);
    const dateRaw = dateKey ? row[dateKey] : "";
    const week    = weekLabelFromDate(dateRaw) || weekLabelFromFilename(file.name) || "Week 1";

    const sales = getVal(row, [
      "gross_sales", "gross sales", "order_amount", "order amount",
      "bill_amount", "bill amount", "total_amount", "total amount",
      "customer_total", "customer total",
    ]);

    const commission = getVal(row, [
      "commission", "platform_commission", "platform commission",
      "total_deductions", "total deductions", "deductions",
    ]);

    const payout = getVal(row, [
      "net_settlement", "net settlement", "settlement_amount", "settlement amount",
      "net_payout", "net payout", "amount_paid",
    ]);

    const prev = weekMap.get(week) || { sales: 0, commission: 0, netPayout: 0 };
    weekMap.set(week, {
      sales:      prev.sales      + sales,
      commission: prev.commission + commission,
      netPayout:  prev.netPayout  + payout,
    });
  }

  return weekMap;
}

// ─────────────────────────────────────────────────────────────────────────────
// POS PARSING
// ─────────────────────────────────────────────────────────────────────────────

function parsePOS(file, platform /* "zomato" | "swiggy" */) {
  const wb    = XLSX.read(new Uint8Array(file.data), { type: "array" });
  const sheet = wb.Sheets[wb.SheetNames[0]];
  const rows  = XLSX.utils.sheet_to_json(sheet, { defval: "" });
  const weekMap = new Map();

  for (const row of rows) {
    const keys = Object.keys(row);

    // Detect if this row belongs to our platform
    const typeCols = ["order_type", "payment_type", "order type", "payment type", "source", "platform"];
    const typeKey  = findCol(keys, typeCols);
    if (typeKey) {
      const typeVal = String(row[typeKey] || "").toLowerCase();
      if (!typeVal.includes(platform)) continue;
    }

    const dateKey = findCol(keys, ["transaction_date", "order_date", "date", "bill_date"]);
    const dateRaw = dateKey ? row[dateKey] : "";
    const week    = weekLabelFromDate(dateRaw) || "Week 1";

    const amt = getVal(row, ["amount", "total", "net_amount", "net amount", "grand_total", "grand total"]);

    weekMap.set(week, (weekMap.get(week) || 0) + amt);
  }

  return weekMap;
}

// ─────────────────────────────────────────────────────────────────────────────
// ACCOUNTANT SUMMARY PARSING
// ─────────────────────────────────────────────────────────────────────────────
// Swiggy/Zomato summary sheets created by accountants.
// We try two strategies:
// A) json rows with clear header columns (Week / Sales / Commission / Payout)
// B) raw row scanning for rows that mention "week" or have date ranges

function parseAccountantSummary(file) {
  const wb = XLSX.read(new Uint8Array(file.data), { type: "array" });
  const sheetName =
    wb.SheetNames.find((s) => /summary|recon/i.test(s)) || wb.SheetNames[0];
  const sheet = wb.Sheets[sheetName];

  // Strategy A: JSON with headers
  const jsonRows = XLSX.utils.sheet_to_json(sheet, { defval: "" });
  const weekResults = [];

  if (jsonRows.length > 0) {
    for (const row of jsonRows) {
      const keys = Object.keys(row);
      const weekKey = findCol(keys, ["week", "period", "dates", "date range"]);
      const weekRaw = weekKey ? String(row[weekKey]) : "";
      const week = extractWeekLabel(weekRaw);
      if (!week) continue;

      const sales      = getVal(row, ["sales", "gross_sales", "gross sales", "total sales", "revenue"]);
      const commission = getVal(row, ["commission", "platform_fee", "deductions", "charges"]);
      const payout     = getVal(row, ["net payout", "net_payout", "settlement", "amount received"]);

      weekResults.push({ label: week, sales, commission, payout });
    }
  }

  // Strategy B: raw 2D rows fallback
  if (weekResults.length === 0) {
    const rawRows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "" });
    for (const row of rawRows) {
      const rowStr = row.join(" ").toLowerCase();
      if (!rowStr.includes("week") && !rowStr.match(/\bw[1-5]\b/)) continue;

      const week = extractWeekLabel(row[0]) || extractWeekLabel(row[1]);
      if (!week) continue;

      // Try to find numeric columns, skipping the first (label) column
      const nums = row.slice(1).map(n).filter((v) => v > 0);
      if (nums.length >= 3) {
        weekResults.push({
          label: week,
          sales:      nums[0],
          commission: nums[1],
          payout:     nums[2],
        });
      } else if (nums.length === 2) {
        weekResults.push({
          label: week,
          sales:      nums[0],
          commission: 0,
          payout:     nums[1],
        });
      }
    }
  }

  return weekResults;
}

function extractWeekLabel(val) {
  const s = String(val || "").toLowerCase();
  if (s.includes("week 1") || /\bw1\b/.test(s)) return "Week 1";
  if (s.includes("week 2") || /\bw2\b/.test(s)) return "Week 2";
  if (s.includes("week 3") || /\bw3\b/.test(s)) return "Week 3";
  if (s.includes("week 4") || /\bw4\b/.test(s)) return "Week 4";
  if (s.includes("week 5") || /\bw5\b/.test(s)) return "Week 5";
  return null;
}

// ─────────────────────────────────────────────────────────────────────────────
// MERGE week maps from multiple files
// ─────────────────────────────────────────────────────────────────────────────
function mergeWeekMaps(maps) {
  const merged = new Map();
  for (const wm of maps) {
    wm.forEach((data, week) => {
      const prev = merged.get(week) || { sales: 0, commission: 0, netPayout: 0 };
      merged.set(week, {
        sales:      prev.sales      + (data.sales      || 0),
        commission: prev.commission + (data.commission || 0),
        netPayout:  prev.netPayout  + (data.netPayout  || 0),
      });
    });
  }
  return merged;
}

function mergePosWeekMaps(maps) {
  const merged = new Map();
  for (const wm of maps) {
    wm.forEach((amt, week) => {
      merged.set(week, (merged.get(week) || 0) + amt);
    });
  }
  return merged;
}

// ─────────────────────────────────────────────────────────────────────────────
// PUBLIC: SWIGGY AUDIT
// ─────────────────────────────────────────────────────────────────────────────

export function auditSwiggyReconciliation(swiggyRawFiles, posFiles, swiggySummaries) {
  const report = { platform: "Swiggy", weeks: [], diagnostics: [] };

  // Parse raw Swiggy files
  const rawMaps = swiggyRawFiles.map((f) => {
    try { return parseSwiggyRaw(f); }
    catch (e) {
      report.diagnostics.push(`⚠️ ${f.name}: ${e.message}`);
      return new Map();
    }
  });
  const rawWeeks = mergeWeekMaps(rawMaps);

  // Parse POS for Swiggy sales
  const posMaps = posFiles.map((f) => {
    try { return parsePOS(f, "swiggy"); }
    catch (e) { return new Map(); }
  });
  const posWeeks = mergePosWeekMaps(posMaps);

  // Parse accountant summaries
  const accWeeks = swiggySummaries.flatMap((f) => {
    try { return parseAccountantSummary(f); }
    catch (e) {
      report.diagnostics.push(`⚠️ Summary ${f.name}: ${e.message}`);
      return [];
    }
  });

  // ── Build result weeks ──────────────────────────────────────────────────
  if (accWeeks.length > 0) {
    for (const acc of accWeeks) {
      const raw = rawWeeks.get(acc.label) || { sales: 0, commission: 0, netPayout: 0 };
      const pos = posWeeks.get(acc.label) || 0;
      report.weeks.push(buildWeekRow(acc.label, acc, raw, pos));
    }
  } else if (rawWeeks.size > 0) {
    // No accountant summary — show raw data with zero baseline
    const allWeeks = ["Week 1", "Week 2", "Week 3", "Week 4", "Week 5"].filter(w => rawWeeks.has(w));
    for (const w of allWeeks) {
      const raw = rawWeeks.get(w);
      const pos = posWeeks.get(w) || 0;
      report.weeks.push(buildWeekRow(w, { sales: 0, commission: 0, payout: 0 }, raw, pos));
    }
  }

  return report;
}

// ─────────────────────────────────────────────────────────────────────────────
// PUBLIC: ZOMATO AUDIT
// ─────────────────────────────────────────────────────────────────────────────

export function auditZomatoReconciliation(zomatoRawFiles, posFiles, zomatoSummaries) {
  const report = { platform: "Zomato", weeks: [], diagnostics: [] };

  const rawMaps = zomatoRawFiles.map((f) => {
    try { return parseZomatoRaw(f); }
    catch (e) {
      report.diagnostics.push(`⚠️ ${f.name}: ${e.message}`);
      return new Map();
    }
  });
  const rawWeeks = mergeWeekMaps(rawMaps);

  const posMaps = posFiles.map((f) => {
    try { return parsePOS(f, "zomato"); }
    catch (e) { return new Map(); }
  });
  const posWeeks = mergePosWeekMaps(posMaps);

  const accWeeks = zomatoSummaries.flatMap((f) => {
    try { return parseAccountantSummary(f); }
    catch (e) {
      report.diagnostics.push(`⚠️ Summary ${f.name}: ${e.message}`);
      return [];
    }
  });

  if (accWeeks.length > 0) {
    for (const acc of accWeeks) {
      const raw = rawWeeks.get(acc.label) || { sales: 0, commission: 0, netPayout: 0 };
      const pos = posWeeks.get(acc.label) || 0;
      report.weeks.push(buildWeekRow(acc.label, acc, raw, pos));
    }
  } else if (rawWeeks.size > 0) {
    const allWeeks = ["Week 1", "Week 2", "Week 3", "Week 4", "Week 5"].filter(w => rawWeeks.has(w));
    for (const w of allWeeks) {
      const raw = rawWeeks.get(w);
      const pos = posWeeks.get(w) || 0;
      report.weeks.push(buildWeekRow(w, { sales: 0, commission: 0, payout: 0 }, raw, pos));
    }
  }

  return report;
}

// ─────────────────────────────────────────────────────────────────────────────
// Helper to build a week comparison row
// ─────────────────────────────────────────────────────────────────────────────
function buildWeekRow(label, acc, raw, posSales) {
  const accSales  = acc.sales      || 0;
  const accComm   = acc.commission || 0;
  const accPayout = acc.payout     || 0;

  const calcSales  = raw.sales      || 0;
  const calcComm   = raw.commission || 0;
  const calcPayout = raw.netPayout  || 0;

  return {
    label,
    accountant:  { sales: accSales,  commission: accComm,  payout: accPayout  },
    calculated:  { sales: calcSales, commission: calcComm, payout: calcPayout },
    pos:         { sales: posSales },
    discrepancy: {
      sales:      accSales  - calcSales,
      commission: accComm   - calcComm,
      payout:     accPayout - calcPayout,
    },
  };
}
