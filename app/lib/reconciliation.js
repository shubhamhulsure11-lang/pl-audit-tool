import * as XLSX from "xlsx";
import JSZip from "jszip";

// Helper to sanitize cell values
const n = (v) => typeof v === "number" ? v : Number(String(v ?? "").replace(/[,₹]/g, "")) || 0;

/**
 * Recursively unzips a zip file and extracts all contained files (including nested zips).
 * Returns an array of objects: { name: string, data: ArrayBuffer/Uint8Array }
 */
export async function extractZipRecursively(file) {
  const filesList = [];
  const zip = await JSZip.loadAsync(file);

  const processEntries = [];
  zip.forEach((relativePath, zipEntry) => {
    if (!zipEntry.dir) {
      processEntries.push((async () => {
        const buffer = await zipEntry.async("arraybuffer");
        const lowerName = relativePath.toLowerCase();
        
        if (lowerName.endsWith(".zip")) {
          // Nested zip extraction
          try {
            const nestedFiles = await extractZipRecursively(buffer);
            filesList.push(...nestedFiles);
          } catch (e) {
            console.error("Error processing nested zip:", relativePath, e);
          }
        } else {
          filesList.push({
            name: relativePath.split("/").pop(),
            path: relativePath,
            data: buffer
          });
        }
      })());
    }
  });

  await Promise.all(processEntries);
  return filesList;
}

/**
 * Classifies extracted files into:
 * - POS reports
 * - Zomato raw reports
 * - Swiggy raw reports
 * - Accountant Swiggy summaries
 * - Accountant Zomato summaries
 */
export function classifyDeliverableFiles(files) {
  const classified = {
    pos: [],
    zomatoRaw: [],
    swiggyRaw: [],
    swiggySummaries: [],
    zomatoSummaries: [],
    unknown: []
  };

  for (const file of files) {
    const name = file.name.toLowerCase();
    
    // Accountant summary checks
    if (name.includes("swiggy summary") || (name.includes("swiggy") && name.includes("reconciliation") && !name.includes("order"))) {
      classified.swiggySummaries.push(file);
      continue;
    }
    if (name.includes("zomato summary") || (name.includes("zomato") && name.includes("reconciliation") && !name.includes("order"))) {
      classified.zomatoSummaries.push(file);
      continue;
    }

    // Try parsing the sheet names if it is an excel file
    if (name.endsWith(".xlsx") || name.endsWith(".xls")) {
      try {
        const workbook = XLSX.read(new Uint8Array(file.data), { type: "array", bookSheets: true });
        const sheets = workbook.SheetNames;

        if (sheets.includes("Order Level") || (sheets.includes("Summary") && sheets.some(s => s.toLowerCase().includes("swiggy")))) {
          classified.swiggyRaw.push(file);
        } else if (sheets.includes("Addition Deductions Details") || sheets.includes("Tax Report") || sheets.some(s => s.toLowerCase().includes("deduction"))) {
          classified.zomatoRaw.push(file);
        } else if (sheets.some(s => s.toLowerCase().includes("payment wise") || s.toLowerCase().includes("petpooja"))) {
          classified.pos.push(file);
        } else {
          classified.unknown.push(file);
        }
      } catch (e) {
        classified.unknown.push(file);
      }
    } else if (name.endsWith(".csv")) {
      if (name.includes("zomato")) classified.zomatoRaw.push(file);
      else if (name.includes("swiggy")) classified.swiggyRaw.push(file);
      else if (name.includes("pos") || name.includes("petpooja") || name.includes("order")) classified.pos.push(file);
      else classified.unknown.push(file);
    } else {
      classified.unknown.push(file);
    }
  }

  return classified;
}

/**
 * Audits Zomato reconciliation.
 * Recalculates Zomato figures from raw data and POS, then compares them with the accountant's summary.
 */
export function auditZomatoReconciliation(zomatoRawFiles, posFiles, zomatoSummaries) {
  const show = (v) => new Intl.NumberFormat("en-IN", { style: "currency", currency: "INR", maximumFractionDigits: 0 }).format(v || 0);
  const auditReport = {
    platform: "Zomato",
    posChecked: posFiles.length > 0,
    summaryFound: zomatoSummaries.length > 0,
    rawCount: zomatoRawFiles.length,
    weeks: []
  };

  // 1. Process Raw Zomato Payout Files
  const rawWeeks = new Map();
  zomatoRawFiles.forEach(file => {
    try {
      const wb = XLSX.read(new Uint8Array(file.data), { type: "array" });
      
      // Look for Order Level / Details sheet
      const orderSheetName = wb.SheetNames.find(s => 
        s.includes("Addition Deductions") || s.includes("Order") || s.includes("Invoice") || s.includes("Details")
      ) || wb.SheetNames[0];

      const sheet = wb.Sheets[orderSheetName];
      const rows = XLSX.utils.sheet_to_json(sheet, { defval: "" });

      // Identify week from filename or data
      const weekLabel = getWeekLabelFromFilename(file.name);

      let sales = 0;
      let commission = 0;
      let netPayout = 0;

      rows.forEach(row => {
        // Zomato headers mapping
        const keys = Object.keys(row);
        const findVal = (choices) => {
          const key = keys.find(k => choices.some(c => k.toLowerCase().includes(c.toLowerCase())));
          return key ? n(row[key]) : 0;
        };

        sales += findVal(["sales", "order amount", "gross", "bill amount"]);
        commission += findVal(["commission", "charge", "fee"]);
        netPayout += findVal(["net payout", "settlement", "payout amount"]);
      });

      rawWeeks.set(weekLabel, { sales, commission, netPayout });
    } catch (e) {
      console.error("Error parsing Zomato raw file:", file.name, e);
    }
  });

  // 2. Process POS Files
  let posWeeklySales = new Map();
  posFiles.forEach(file => {
    try {
      const wb = XLSX.read(new Uint8Array(file.data), { type: "array" });
      const sheet = wb.Sheets[wb.SheetNames[0]];
      const rows = XLSX.utils.sheet_to_json(sheet, { defval: "" });

      rows.forEach(row => {
        const orderType = String(row["Order_Type"] || row["order_type"] || "").toLowerCase();
        const paymentType = String(row["Payment_Type"] || row["payment_type"] || "").toLowerCase();
        const isZomato = orderType.includes("zomato") || paymentType.includes("zomato");
        
        if (isZomato) {
          const dateStr = String(row["Transaction_Date"] || row["date"] || "");
          const weekLabel = getWeekLabelFromDate(dateStr);
          const totalAmt = n(row["Amount"] || row["Total"] || row["total"]);
          
          posWeeklySales.set(weekLabel, (posWeeklySales.get(weekLabel) || 0) + totalAmt);
        }
      });
    } catch (e) {
      console.error("Error parsing POS file:", file.name, e);
    }
  });

  // 3. Process Accountant Summaries & Compare
  zomatoSummaries.forEach(file => {
    try {
      const wb = XLSX.read(new Uint8Array(file.data), { type: "array" });
      const sheetName = wb.SheetNames.find(s => s.toLowerCase().includes("summary") || s.toLowerCase().includes("recon")) || wb.SheetNames[0];
      const sheet = wb.Sheets[sheetName];
      const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "" });

      // Scan rows to find week summary rows (contain week indicators)
      rows.forEach((row, idx) => {
        const rowStr = row.join(" ").toLowerCase();
        if (rowStr.includes("week") || rowStr.includes("w1") || rowStr.includes("w2")) {
          // Assume columns: [Week, Sales, Commission, Payout]
          const weekName = getWeekLabelFromText(row[0] || row[1]);
          if (!weekName) return;

          const accSales = n(row[2] || row[1]);
          const accComm = n(row[3] || row[2]);
          const accPayout = n(row[4] || row[3]);

          const rawData = rawWeeks.get(weekName) || { sales: 0, commission: 0, netPayout: 0 };
          const posSales = posWeeklySales.get(weekName) || 0;

          auditReport.weeks.push({
            label: weekName,
            accountant: { sales: accSales, commission: accComm, payout: accPayout },
            calculated: { sales: rawData.sales, commission: rawData.commission, payout: rawData.netPayout },
            pos: { sales: posSales },
            discrepancy: {
              sales: accSales - rawData.sales,
              commission: accComm - rawData.commission,
              payout: accPayout - rawData.netPayout
            }
          });
        }
      });
    } catch (e) {
      console.error("Error auditing accountant summary:", file.name, e);
    }
  });

  // Fallback: If no accountant summary matched but raw files exist, present raw data values directly
  if (auditReport.weeks.length === 0 && rawWeeks.size > 0) {
    rawWeeks.forEach((data, weekName) => {
      const posSales = posWeeklySales.get(weekName) || 0;
      auditReport.weeks.push({
        label: weekName,
        accountant: { sales: 0, commission: 0, payout: 0 },
        calculated: data,
        pos: { sales: posSales },
        discrepancy: { sales: -data.sales, commission: -data.commission, payout: -data.netPayout }
      });
    });
  }

  return auditReport;
}

/**
 * Audits Swiggy reconciliation.
 */
export function auditSwiggyReconciliation(swiggyRawFiles, posFiles, swiggySummaries) {
  const show = (v) => new Intl.NumberFormat("en-IN", { style: "currency", currency: "INR", maximumFractionDigits: 0 }).format(v || 0);
  const auditReport = {
    platform: "Swiggy",
    posChecked: posFiles.length > 0,
    summaryFound: swiggySummaries.length > 0,
    rawCount: swiggyRawFiles.length,
    weeks: []
  };

  const rawWeeks = new Map();
  swiggyRawFiles.forEach(file => {
    try {
      const wb = XLSX.read(new Uint8Array(file.data), { type: "array" });
      const sheetName = wb.SheetNames.find(s => s.toLowerCase().includes("order")) || wb.SheetNames[0];
      const sheet = wb.Sheets[sheetName];
      const rows = XLSX.utils.sheet_to_json(sheet, { defval: "" });

      const weekLabel = getWeekLabelFromFilename(file.name);

      let sales = 0;
      let commission = 0;
      let netPayout = 0;

      rows.forEach(row => {
        const keys = Object.keys(row);
        const findVal = (choices) => {
          const key = keys.find(k => choices.some(c => k.toLowerCase().includes(c.toLowerCase())));
          return key ? n(row[key]) : 0;
        };

        sales += findVal(["bill amount", "order amount", "gross", "sales"]);
        commission += findVal(["commission", "service fee", "charge"]);
        netPayout += findVal(["net payout", "settlement", "payout amount"]);
      });

      rawWeeks.set(weekLabel, { sales, commission, netPayout });
    } catch (e) {
      console.error("Error parsing Swiggy raw file:", file.name, e);
    }
  });

  // POS Swiggy sales match
  let posWeeklySales = new Map();
  posFiles.forEach(file => {
    try {
      const wb = XLSX.read(new Uint8Array(file.data), { type: "array" });
      const sheet = wb.Sheets[wb.SheetNames[0]];
      const rows = XLSX.utils.sheet_to_json(sheet, { defval: "" });

      rows.forEach(row => {
        const orderType = String(row["Order_Type"] || row["order_type"] || "").toLowerCase();
        const paymentType = String(row["Payment_Type"] || row["payment_type"] || "").toLowerCase();
        const isSwiggy = orderType.includes("swiggy") || paymentType.includes("swiggy");
        
        if (isSwiggy) {
          const dateStr = String(row["Transaction_Date"] || row["date"] || "");
          const weekLabel = getWeekLabelFromDate(dateStr);
          const totalAmt = n(row["Amount"] || row["Total"] || row["total"]);
          
          posWeeklySales.set(weekLabel, (posWeeklySales.get(weekLabel) || 0) + totalAmt);
        }
      });
    } catch (e) {
      console.error("Error parsing POS file:", file.name, e);
    }
  });

  swiggySummaries.forEach(file => {
    try {
      const wb = XLSX.read(new Uint8Array(file.data), { type: "array" });
      const sheetName = wb.SheetNames.find(s => s.toLowerCase().includes("summary") || s.toLowerCase().includes("recon")) || wb.SheetNames[0];
      const sheet = wb.Sheets[sheetName];
      const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "" });

      rows.forEach(row => {
        const rowStr = row.join(" ").toLowerCase();
        if (rowStr.includes("week") || rowStr.includes("w1") || rowStr.includes("w2")) {
          const weekName = getWeekLabelFromText(row[0] || row[1]);
          if (!weekName) return;

          const accSales = n(row[2] || row[1]);
          const accComm = n(row[3] || row[2]);
          const accPayout = n(row[4] || row[3]);

          const rawData = rawWeeks.get(weekName) || { sales: 0, commission: 0, netPayout: 0 };
          const posSales = posWeeklySales.get(weekName) || 0;

          auditReport.weeks.push({
            label: weekName,
            accountant: { sales: accSales, commission: accComm, payout: accPayout },
            calculated: { sales: rawData.sales, commission: rawData.commission, payout: rawData.netPayout },
            pos: { sales: posSales },
            discrepancy: {
              sales: accSales - rawData.sales,
              commission: accComm - rawData.commission,
              payout: accPayout - rawData.netPayout
            }
          });
        }
      });
    } catch (e) {
      console.error("Error auditing Swiggy accountant summary:", file.name, e);
    }
  });

  if (auditReport.weeks.length === 0 && rawWeeks.size > 0) {
    rawWeeks.forEach((data, weekName) => {
      const posSales = posWeeklySales.get(weekName) || 0;
      auditReport.weeks.push({
        label: weekName,
        accountant: { sales: 0, commission: 0, payout: 0 },
        calculated: data,
        pos: { sales: posSales },
        discrepancy: { sales: -data.sales, commission: -data.commission, payout: -data.netPayout }
      });
    });
  }

  return auditReport;
}

// Private helper to extract a week label from filename
function getWeekLabelFromFilename(filename) {
  const cleanName = filename.toLowerCase();
  const match = cleanName.match(/(week[-_ ]?\d|w\d)/);
  if (match) {
    const digit = match[0].match(/\d/)[0];
    return `Week ${digit}`;
  }
  // Guess based on dates
  if (cleanName.includes("01_") || cleanName.includes("1st")) return "Week 1";
  if (cleanName.includes("08_") || cleanName.includes("8th")) return "Week 2";
  if (cleanName.includes("15_") || cleanName.includes("15th")) return "Week 3";
  if (cleanName.includes("22_") || cleanName.includes("22nd")) return "Week 4";
  return "Week 1"; // Default
}

// Private helper to identify week index from transaction date string
function getWeekLabelFromDate(dateStr) {
  try {
    const dt = new Date(dateStr);
    const day = dt.getDate();
    if (isNaN(day)) return "Week 1";
    if (day <= 7) return "Week 1";
    if (day <= 14) return "Week 2";
    if (day <= 21) return "Week 3";
    return "Week 4";
  } catch (e) {
    return "Week 1";
  }
}

// Helper to sanitize week labels in accountant sheet
function getWeekLabelFromText(val) {
  const s = String(val || "").toLowerCase();
  if (s.includes("week 1") || s.includes("w1")) return "Week 1";
  if (s.includes("week 2") || s.includes("w2")) return "Week 2";
  if (s.includes("week 3") || s.includes("w3")) return "Week 3";
  if (s.includes("week 4") || s.includes("w4")) return "Week 4";
  if (s.includes("week 5") || s.includes("w5")) return "Week 5";
  return null;
}
