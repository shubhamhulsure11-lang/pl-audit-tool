/**
 * Automated Account Head Auditor Engine
 * 
 * Multi-layer accounting audit:
 * 1. Data extraction & normalization
 * 2. Deterministic Hard Accounting Rules (Overrides AI)
 *    - Rule 1: ICE / ICE CUBES / ICE PURCHASE / PREMIUM ICE CUBES / ICE PURCHASES -> Other Purchases
 *    - Rule 2: CHARCOAL / CHARCOALS / KOYLA -> Other Purchases (not Kitchen Gas)
 *    - Rule 3: STAFF RICE / RICE - STAFF MASOORI / RICE-STAFF MASOORI -> Staff Welfare
 * 3. Domain Heuristics & Direct Matching
 * 4. Dataset Consistency / Similar Item Check (Pass 2)
 * 5. Modular AI Classification (Groq) with Safe Batching
 * 6. Confidence & Non-Existent Account Guard (DO NOT MAKE UP ACCOUNT HEADS)
 * 7. 100% Row Reconciliation: Rows Reviewed = Definite + Needs Verification + Correct
 */

import * as XLSX from "xlsx";

// ─── 1. ITEM NORMALIZATION ───────────────────────────────────────────────────

/**
 * Normalizes an item description while preserving essential accounting keywords.
 * Preserves attributes like 'staff', 'rice', 'cube', 'cubes', 'charcoal', 'frozen',
 * 'fresh', 'chicken', 'mutton', 'veg', 'non veg', '1l', '500ml', etc.
 */
export function normalizeItem(str) {
  if (!str) return "";
  return String(str)
    .trim()
    .replace(/[\r\n\t]+/g, " ")
    .replace(/[,\/\\#\-–—_:]+/g, " ") // replace separators with single space
    .replace(/\s+/g, " ")            // collapse multi spaces
    .toLowerCase();
}

/**
 * Extracts key semantic tokens for similarity comparison.
 * Filters out generic packing sizes, noise words, but keeps product identity.
 */
export function extractCoreTokens(str) {
  const norm = normalizeItem(str);
  const words = norm.split(" ").filter(w => w.length > 1);
  const stopWords = new Set(["pkt", "pack", "packet", "box", "case", "bottle", "btl", "tin", "can", "nos", "no", "pcs", "pc", "brand", "grade", "imported", "local", "best", "special", "fresh", "frozen"]);
  return words.filter(w => !stopWords.has(w) && !/^\d+(?:kg|gm|g|ltr|lt|l|ml)?$/.test(w));
}

// ─── 2. ACCOUNT HEAD MATCHING & VALIDATION ───────────────────────────────────

/**
 * Finds a matching account head from available heads (case-insensitive).
 * NEVER invents a new account head.
 */
export function matchAvailableHead(targetHead, availableHeads) {
  if (!targetHead || !availableHeads || availableHeads.length === 0) return null;
  const t = targetHead.trim().toLowerCase();
  
  // Exact match (case-insensitive)
  const exact = availableHeads.find(h => h.trim().toLowerCase() === t);
  if (exact) return exact;

  // Plural/singular or substring match (e.g. "Other Purchase" vs "Other Purchases")
  const fuzzy = availableHeads.find(h => {
    const hl = h.trim().toLowerCase();
    return hl.replace(/s$/, "") === t.replace(/s$/, "") ||
           hl.includes(t) || t.includes(hl);
  });
  return fuzzy || null;
}

// ─── 3. DETERMINISTIC HARD ACCOUNTING RULES ──────────────────────────────────

export const HARD_RULES = [
  {
    id: "rule_ice",
    name: "Ice Purchases",
    targetHead: "Other Purchases",
    reason: "Ice is governed by the established Other Purchases rule (do not classify as Beverages).",
    test: (norm) => {
      // Must not match ice cream or iced tea
      if (/\b(ice\s*cream|icecream|tea|juice|spiced)\b/i.test(norm)) return false;
      return /\b(ice\s*cubes?|ice\s*purchases?|premium\s*ice(\s*cubes?)?|^ice$)\b/i.test(norm) ||
             norm === "ice" || norm === "ice cube" || norm === "ice cubes";
    }
  },
  {
    id: "rule_charcoal",
    name: "Charcoal / Fuel",
    targetHead: "Other Purchases",
    reason: "Charcoal/Koyla is governed by the established Other Purchases rule (do not classify as Kitchen Gas).",
    test: (norm) => {
      return /\b(charcoals?|koyla)\b/i.test(norm);
    }
  },
  {
    id: "rule_staff_rice",
    name: "Staff Rice",
    targetHead: "Staff Welfare",
    reason: "Staff food/rice is earmarked specifically for Staff Welfare.",
    test: (norm) => {
      // Matches: STAFF RICE, RICE - STAFF MASOORI, RICE-STAFF MASOORI, STAFF RICE MASOORI
      return /\b(staff\s*rice|rice\s*staff|staff\s*food|staff\s*meal)\b/i.test(norm) ||
             (/\brice\b/i.test(norm) && /\bstaff\b/i.test(norm));
    }
  }
];

/**
 * Evaluates an item against deterministic hard rules.
 * Overrides any AI or consistency classification.
 */
export function evaluateHardRules(itemDesc, currentHead, availableHeads) {
  const norm = normalizeItem(itemDesc);
  if (!norm) return null;

  for (const rule of HARD_RULES) {
    if (rule.test(norm)) {
      const matchedHead = matchAvailableHead(rule.targetHead, availableHeads);
      const isAlreadyCorrect = matchedHead && currentHead && currentHead.trim().toLowerCase() === matchedHead.trim().toLowerCase();

      if (isAlreadyCorrect) {
        return {
          classification: "correct",
          recommended_account_head: matchedHead,
          confidence: 1.0,
          reason: `Properly classified under '${matchedHead}'.`,
          rule_source: "hard_rule"
        };
      }

      if (matchedHead) {
        return {
          classification: "definite_correction",
          recommended_account_head: matchedHead,
          confidence: 1.0,
          reason: rule.reason,
          rule_source: "hard_rule"
        };
      } else {
        // Target head not present in available account heads!
        return {
          classification: "needs_verification",
          recommended_account_head: "Needs invoice/use verification",
          confidence: 0.85,
          reason: `${rule.name} normally belongs to '${rule.targetHead}', but this head is not available in the client's current Chart of Accounts.`,
          rule_source: "hard_rule"
        };
      }
    }
  }

  return null;
}

// ─── 4. DOMAIN PATTERN & OBVIOUS CLASSIFICATION LAYER ────────────────────────

const DOMAIN_CATEGORY_PATTERNS = [
  {
    headPattern: /dairy/i,
    itemPattern: /\b(milk|curd|paneer|butter|cheese|ghee|cream|yogurt|lassi|khoya|dahi)\b/i,
    headName: "Dairy products purchases"
  },
  {
    headPattern: /poultry|meat/i,
    itemPattern: /\b(chicken|mutton|lamb|goat|beef|pork|egg|eggs)\b/i,
    headName: "Poultry and meat purchases"
  },
  {
    headPattern: /sea\s*food|fish/i,
    itemPattern: /\b(fish|prawn|prawns|crab|pomfret|salmon|surmai|squid|rohu|katla)\b/i,
    headName: "Sea food purchases"
  },
  {
    headPattern: /clean|housekeep/i,
    itemPattern: /\b(colin|harpic|bleach|detergent|mop|wiper|broom|garbage\s*bag|phenyl|sanitizer|handwash|duster)\b/i,
    headName: "Cleaning and Housekeeping"
  },
  {
    headPattern: /print|station/i,
    itemPattern: /\b(bill\s*roll|thermal\s*roll|pen|paper|register|stapler|toner|cartridge)\b/i,
    headName: "Printing & Stationery"
  },
  {
    headPattern: /pack/i,
    itemPattern: /\b(container|foil|cling\s*film|box|bag|carry\s*bag|pouch|straw|meal\s*tray|aluminium\s*foil)\b/i,
    headName: "Packing material"
  },
  {
    headPattern: /repair|maint/i,
    itemPattern: /\b(plumbing|electrical|wire|switch|bulb|tube\s*light|ac\s*service|welding|hardware)\b/i,
    headName: "Repairs and Maintenance"
  }
];

export function checkDomainAlignment(itemDesc, currentHead, availableHeads) {
  const norm = normalizeItem(itemDesc);
  const cur = (currentHead || "").trim().toLowerCase();

  for (const cat of DOMAIN_CATEGORY_PATTERNS) {
    if (cat.itemPattern.test(norm)) {
      // If current head already matches this category, it is verified correct!
      if (cat.headPattern.test(cur)) {
        return {
          classification: "correct",
          recommended_account_head: currentHead,
          confidence: 0.95,
          reason: `Item description clearly aligns with '${currentHead}'.`,
          rule_source: "domain_rule"
        };
      }

      // Check if the appropriate head exists in availableHeads
      const suitableHead = availableHeads.find(h => cat.headPattern.test(h));
      if (suitableHead) {
        // High confidence reclassification
        return {
          classification: "definite_correction",
          recommended_account_head: suitableHead,
          confidence: 0.92,
          reason: `'${itemDesc}' belongs under '${suitableHead}' rather than '${currentHead}'.`,
          rule_source: "domain_rule"
        };
      }
    }
  }

  return null;
}

// ─── 5. SIMILAR ITEM / CONSISTENCY CHECK ENGINE (PASS 2) ─────────────────────

/**
 * Checks consistency across similar items in the dataset.
 * e.g., if 8 items like "Mojito Mint Syrup" are in "Beverages", but 1 is in "Groceries purchases",
 * the outlier is identified and corrected.
 */
export function performConsistencyCheck(uniqueProfiles, availableHeads) {
  const consistencyMap = new Map(); // tokenKey -> list of profiles
  
  uniqueProfiles.forEach(p => {
    const tokens = extractCoreTokens(p.item);
    if (tokens.length >= 2) {
      const key = tokens.slice(0, 2).sort().join("_");
      if (!consistencyMap.has(key)) consistencyMap.set(key, []);
      consistencyMap.get(key).push(p);
    }
  });

  const updates = new Map(); // profileKey -> update result

  consistencyMap.forEach((profiles, key) => {
    if (profiles.length < 2) return;

    // Count distribution of current account heads
    const headCounts = {};
    profiles.forEach(p => {
      const h = p.current_account_head;
      headCounts[h] = (headCounts[h] || 0) + 1;
    });

    const entries = Object.entries(headCounts).sort((a, b) => b[1] - a[1]);
    if (entries.length > 1) {
      const [dominantHead, dominantCount] = entries[0];
      const dominanceRatio = dominantCount / profiles.length;

      // If dominant head represents >= 75% of similar items
      if (dominanceRatio >= 0.75 && dominantCount >= 2) {
        const targetHead = matchAvailableHead(dominantHead, availableHeads);
        if (targetHead) {
          profiles.forEach(p => {
            if (p.current_account_head !== dominantHead) {
              const matchedProfile = p;
              const similarNames = profiles.filter(x => x.current_account_head === dominantHead).map(x => x.item).slice(0, 2);
              updates.set(p.profileKey, {
                classification: "definite_correction",
                recommended_account_head: targetHead,
                confidence: 0.90,
                reason: `Consistency check: ${dominantCount} similar items (${similarNames.map(s => `'${s}'`).join(", ")}) are categorized under '${targetHead}'.`,
                rule_source: "consistency_check"
              });
            }
          });
        }
      } else {
        // Conflicting classification without strong majority -> needs verification
        profiles.forEach(p => {
          updates.set(p.profileKey, {
            classification: "needs_verification",
            recommended_account_head: "Needs invoice/use verification",
            confidence: 0.70,
            reason: `Conflicting classifications found across similar items (${entries.map(e => `'${e[0]}' (${e[1]})`).join(" vs ")}). Requires invoice review.`,
            rule_source: "consistency_check"
          });
        });
      }
    }
  });

  return updates;
}

// ─── 6. AI CLASSIFICATION LAYER (BATCHED) ───────────────────────────────────

/**
 * Batched AI classification for ambiguous items.
 * Uses Groq via the server endpoint or graceful fallback.
 */
export async function classifyItemsWithAI(itemsToAudit, availableHeads, apiKey, model) {
  if (!itemsToAudit || itemsToAudit.length === 0) return {};
  if (!apiKey) {
    // Graceful fallback when no AI key is provided
    const fallbackResults = {};
    itemsToAudit.forEach(it => {
      fallbackResults[it.profileKey] = {
        classification: "needs_verification",
        recommended_account_head: "Needs invoice/use verification",
        confidence: 0.50,
        reason: "Requires manual invoice/use verification (AI integration not configured).",
        rule_source: "fallback"
      };
    });
    return fallbackResults;
  }

  const results = {};
  const BATCH_SIZE = 15;

  for (let i = 0; i < itemsToAudit.length; i += BATCH_SIZE) {
    const batch = itemsToAudit.slice(i, i + BATCH_SIZE);
    try {
      const res = await fetch("/api/ai-audit", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          batchAuditItems: batch.map(b => ({
            profileKey: b.profileKey,
            item: b.item,
            vendor: b.vendor,
            current_account_head: b.current_account_head,
            amount: b.amount
          })),
          availableAccounts: availableHeads,
          apiKey,
          model: model || "llama-3.1-8b-instant"
        })
      });

      if (!res.ok) {
        throw new Error(`AI API error ${res.status}`);
      }

      const data = await res.json();
      if (data.results && Array.isArray(data.results)) {
        data.results.forEach(r => {
          const validHead = matchAvailableHead(r.recommended_account_head, availableHeads);
          const isDefinite = r.classification === "definite_correction" && validHead && (r.confidence >= 0.85);

          if (isDefinite) {
            results[r.profileKey || r.item] = {
              classification: "definite_correction",
              recommended_account_head: validHead,
              confidence: r.confidence || 0.90,
              reason: r.reason || "AI classification determined stronger accounting fit.",
              rule_source: "ai"
            };
          } else if (r.classification === "correct" || r.current_verdict === "CORRECT") {
            results[r.profileKey || r.item] = {
              classification: "correct",
              recommended_account_head: r.current_account_head,
              confidence: r.confidence || 0.90,
              reason: r.reason || "Current classification verified.",
              rule_source: "ai"
            };
          } else {
            results[r.profileKey || r.item] = {
              classification: "needs_verification",
              recommended_account_head: "Needs invoice/use verification",
              confidence: r.confidence || 0.60,
              reason: r.reason || "Description requires invoice review.",
              rule_source: "ai"
            };
          }
        });
      } else {
        throw new Error("Invalid response format from AI");
      }
    } catch (err) {
      // Failure Handling: NEVER lose a row
      batch.forEach(b => {
        results[b.profileKey] = {
          classification: "needs_verification",
          recommended_account_head: "Needs invoice/use verification",
          confidence: 0.50,
          reason: `AI classification unavailable; requires manual verification (${err.message}).`,
          rule_source: "fallback"
        };
      });
    }
  }

  return results;
}

// ─── 7. MAIN AUDIT ENGINE ───────────────────────────────────────────────────

/**
 * Runs the complete automated Account Head Audit on 100% of the dataset rows.
 * Reconciles: totalReviewed === definiteCorrections + needsVerification + correct
 */
export async function runAccountHeadAudit({
  records,
  availableHeads,
  clientName,
  periodName,
  apiKey,
  model,
  onProgress
}) {
  if (!records || records.length === 0) {
    throw new Error("No records available to audit.");
  }
  if (!clientName || !periodName) {
    throw new Error("Both Client Name and Period must be specified for the audit context.");
  }

  const totalRows = records.length;
  if (onProgress) onProgress(0, totalRows, "Extracting and grouping dataset records...");

  // Unique list of available heads
  const validHeads = (availableHeads && availableHeads.length > 0)
    ? availableHeads
    : Array.from(new Set(records.map(r => r.account).filter(a => a && a !== "Unassigned Account")));

  // Step 1: Group identical (item, account, vendor) tuples to optimize processing
  const profileMap = new Map();
  records.forEach((r, idx) => {
    const norm = normalizeItem(r.item);
    const profileKey = `${norm}|||${(r.account || "").trim().toLowerCase()}|||${(r.vendor || "").trim().toLowerCase()}`;
    if (!profileMap.has(profileKey)) {
      profileMap.set(profileKey, {
        profileKey,
        item: r.item || "(No Item Name)",
        vendor: r.vendor || "(No Vendor)",
        current_account_head: r.account || "Unassigned Account",
        amount: r.total || 0,
        rowIndices: []
      });
    }
    const profile = profileMap.get(profileKey);
    profile.rowIndices.push(idx);
    profile.amount += (r.total || 0);
  });

  const uniqueProfiles = Array.from(profileMap.values());
  const profileDecisions = new Map(); // profileKey -> decision

  // Step 2: Deterministic Hard Rules
  if (onProgress) onProgress(Math.floor(totalRows * 0.2), totalRows, "Evaluating hard accounting rules...");
  const remainingForDomain = [];
  uniqueProfiles.forEach(p => {
    const hardVerdict = evaluateHardRules(p.item, p.current_account_head, validHeads);
    if (hardVerdict) {
      profileDecisions.set(p.profileKey, hardVerdict);
    } else {
      remainingForDomain.push(p);
    }
  });

  // Step 3: Domain Heuristic Alignment
  if (onProgress) onProgress(Math.floor(totalRows * 0.4), totalRows, "Checking category alignment...");
  const remainingForConsistency = [];
  remainingForDomain.forEach(p => {
    const domainVerdict = checkDomainAlignment(p.item, p.current_account_head, validHeads);
    if (domainVerdict) {
      profileDecisions.set(p.profileKey, domainVerdict);
    } else {
      remainingForConsistency.push(p);
    }
  });

  // Step 4: Consistency & Similar Item Check
  if (onProgress) onProgress(Math.floor(totalRows * 0.6), totalRows, "Running similar item consistency check...");
  const consistencyUpdates = performConsistencyCheck(remainingForConsistency, validHeads);
  const remainingForAI = [];
  remainingForConsistency.forEach(p => {
    if (consistencyUpdates.has(p.profileKey)) {
      profileDecisions.set(p.profileKey, consistencyUpdates.get(p.profileKey));
    } else {
      remainingForAI.push(p);
    }
  });

  // Step 5: AI Classification for Ambiguous Items
  if (remainingForAI.length > 0) {
    if (onProgress) onProgress(Math.floor(totalRows * 0.8), totalRows, `Analyzing ${remainingForAI.length} ambiguous item profiles...`);
    const aiResults = await classifyItemsWithAI(remainingForAI, validHeads, apiKey, model);
    remainingForAI.forEach(p => {
      const res = aiResults[p.profileKey];
      if (res) {
        profileDecisions.set(p.profileKey, res);
      } else {
        // Fallback default
        profileDecisions.set(p.profileKey, {
          classification: "needs_verification",
          recommended_account_head: "Needs invoice/use verification",
          confidence: 0.50,
          reason: "Item description requires manual verification.",
          rule_source: "fallback"
        });
      }
    });
  }

  // Step 6: Map decisions back to EVERY SINGLE original row (100% review)
  if (onProgress) onProgress(totalRows, totalRows, "Reconciling audit results across all rows...");

  const allRowAudits = records.map((r, idx) => {
    const norm = normalizeItem(r.item);
    const profileKey = `${norm}|||${(r.account || "").trim().toLowerCase()}|||${(r.vendor || "").trim().toLowerCase()}`;
    const dec = profileDecisions.get(profileKey) || {
      classification: "correct",
      recommended_account_head: r.account || "Unassigned Account",
      confidence: 0.90,
      reason: "No irregularities found.",
      rule_source: "default"
    };

    return {
      rowId: r.id ?? idx,
      date: r.date || "",
      bill: r.bill || "",
      branch: r.branch || "",
      item: r.item || "(No Item Name)",
      vendor: r.vendor || "(No Vendor)",
      amount: r.total || 0,
      current_account_head: r.account || "Unassigned Account",
      recommended_account_head: dec.recommended_account_head,
      classification: dec.classification,
      confidence: dec.confidence,
      reason: dec.reason,
      rule_source: dec.rule_source,
      client: clientName,
      period: periodName
    };
  });

  const definiteCorrections = allRowAudits.filter(r => r.classification === "definite_correction");
  const needsVerification = allRowAudits.filter(r => r.classification === "needs_verification");
  const correctRows = allRowAudits.filter(r => r.classification === "correct");

  // Reconcile check
  const reconciled = (allRowAudits.length === (definiteCorrections.length + needsVerification.length + correctRows.length));
  if (!reconciled) {
    console.error("Audit reconciliation error: sum of categories does not match total rows!", {
      total: allRowAudits.length,
      definite: definiteCorrections.length,
      needsVerification: needsVerification.length,
      correct: correctRows.length
    });
  }

  return {
    clientName,
    periodName,
    totalRows: allRowAudits.length,
    reconciled,
    definiteCorrections,
    needsVerification,
    correctRows,
    allRowAudits,
    summary: {
      totalReviewed: allRowAudits.length,
      definiteCount: definiteCorrections.length,
      verificationCount: needsVerification.length,
      correctCount: correctRows.length
    }
  };
}

// ─── 8. CLIPBOARD COPY UTILITIES ─────────────────────────────────────────────

/**
 * Generates TSV text for pasting directly into Excel:
 * Item \t Current Account Head \t Correct Account Head
 */
export function generateTsvCopy(rows, mode = "all_corrections") {
  let targetRows = rows;
  if (mode === "definite") {
    targetRows = rows.filter(r => r.classification === "definite_correction");
  } else if (mode === "verification") {
    targetRows = rows.filter(r => r.classification === "needs_verification");
  } else if (mode === "all_corrections") {
    targetRows = rows.filter(r => r.classification !== "correct");
  }

  const lines = [
    ["Item", "Current Account Head", "Recommended Account Head"].join("\t")
  ];

  targetRows.forEach(r => {
    lines.push([
      r.item || "",
      r.current_account_head || "",
      r.recommended_account_head || ""
    ].join("\t"));
  });

  return lines.join("\n");
}

// ─── 9. EXCEL EXPORT UTILITY ────────────────────────────────────────────────

export function exportAuditResultsToExcel({ auditResult, clientName, periodName }) {
  if (!auditResult) return;

  const wb = XLSX.utils.book_new();

  // Sheet 1: Definite Corrections
  const defData = [
    ["#", "Item Description", "Vendor", "Current Account Head", "Recommended Account Head", "Amount", "Confidence", "Reason", "Rule Source", "Client", "Period"],
    ...auditResult.definiteCorrections.map((r, i) => [
      i + 1,
      r.item,
      r.vendor,
      r.current_account_head,
      r.recommended_account_head,
      r.amount || 0,
      r.confidence ? `${Math.round(r.confidence * 100)}%` : "-",
      r.reason,
      r.rule_source,
      r.client,
      r.period
    ])
  ];
  const wsDef = XLSX.utils.aoa_to_sheet(defData);
  XLSX.utils.book_append_sheet(wb, wsDef, "Definite Corrections");

  // Sheet 2: Needs Verification
  const verData = [
    ["#", "Item Description", "Vendor", "Current Account Head", "Recommended Account Head", "Amount", "Confidence", "Reason", "Rule Source", "Client", "Period"],
    ...auditResult.needsVerification.map((r, i) => [
      i + 1,
      r.item,
      r.vendor,
      r.current_account_head,
      r.recommended_account_head,
      r.amount || 0,
      r.confidence ? `${Math.round(r.confidence * 100)}%` : "-",
      r.reason,
      r.rule_source,
      r.client,
      r.period
    ])
  ];
  const wsVer = XLSX.utils.aoa_to_sheet(verData);
  XLSX.utils.book_append_sheet(wb, wsVer, "Needs Verification");

  // Sheet 3: Summary & 100% Audit Trace
  const allData = [
    ["#", "Item Description", "Vendor", "Current Account Head", "Recommended Account Head", "Classification", "Amount", "Confidence", "Reason", "Rule Source", "Client", "Period"],
    ...auditResult.allRowAudits.map((r, i) => [
      i + 1,
      r.item,
      r.vendor,
      r.current_account_head,
      r.recommended_account_head,
      r.classification,
      r.amount || 0,
      r.confidence ? `${Math.round(r.confidence * 100)}%` : "-",
      r.reason,
      r.rule_source,
      r.client,
      r.period
    ])
  ];
  const wsAll = XLSX.utils.aoa_to_sheet(allData);
  XLSX.utils.book_append_sheet(wb, wsAll, "Full 100% Audit Log");

  const cleanClient = (clientName || "Client").replace(/[^a-zA-Z0-9_-]/g, "_");
  const cleanPeriod = (periodName || "Period").replace(/[^a-zA-Z0-9_-]/g, "_");
  const fileName = `${cleanClient}_Account_Head_Audit_${cleanPeriod}.xlsx`;

  XLSX.writeFile(wb, fileName);
}
