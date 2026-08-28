/**
 * knowledge.js
 * Pure-JS helpers for the client Knowledge Base.
 *
 * Priority order for classification:
 *   1. client_item_knowledge (verified)
 *   2. client_item_knowledge (historical_import)
 *   3. RESTAURANT_TAXONOMY rule engine  ← fallback only
 *   4. AI (Groq)                        ← unknown items only
 *   5. Human approval → save back to KB
 */

// ─── Normalisation ────────────────────────────────────────────────────────────

/**
 * Normalise an item name for KB lookup.
 * Strips brand suffixes, quantities, units, and special characters
 * so "MONIN WATERMELON 700ML" matches "MONIN WATERMELON".
 */
export function normalizeItemName(raw) {
  if (!raw) return "";
  return String(raw)
    .toLowerCase()
    // Remove quantities like 500ml, 1kg, 250g, 6pcs, 1ltr, etc.
    .replace(/\b\d+\s*(ml|ltr|litre|liter|l|kg|g|gm|gram|grams|pcs|pc|nos|no|unit|units|dozen|doz|box|boxes|pkt|packet|packets|btl|bottle|bottles|can|cans|bag|bags|sack|sacks)\b/gi, "")
    // Remove standalone numbers
    .replace(/\b\d+(\.\d+)?\b/g, "")
    // Remove special characters but keep spaces
    .replace(/[^a-z0-9\s]/g, " ")
    // Collapse whitespace
    .replace(/\s+/g, " ")
    .trim();
}

// ─── Knowledge Matching ───────────────────────────────────────────────────────

/**
 * Search the in-memory knowledge array for a match.
 * Returns the best matching knowledge item or null.
 *
 * Matching strategy (in priority order):
 *   1. Exact normalised match
 *   2. Starts-with match (item name starts with knowledge key)
 *   3. Contains match (knowledge key is contained in item name)
 */
export function matchItemToKnowledge(itemName, knowledgeItems) {
  if (!itemName || !knowledgeItems || knowledgeItems.length === 0) return null;

  const norm = normalizeItemName(itemName);
  if (!norm) return null;

  // 1. Exact match
  const exact = knowledgeItems.find((k) => k.item_name_norm === norm);
  if (exact) return { ...exact, matchType: "exact" };

  // 2. Starts-with match (longest match wins)
  let bestStartsWith = null;
  let bestLen = 0;
  for (const k of knowledgeItems) {
    if (norm.startsWith(k.item_name_norm) && k.item_name_norm.length > bestLen) {
      bestLen = k.item_name_norm.length;
      bestStartsWith = k;
    }
  }
  if (bestStartsWith) return { ...bestStartsWith, matchType: "starts_with" };

  // 3. Contains match (knowledge key appears in item name, longest wins)
  let bestContains = null;
  bestLen = 0;
  for (const k of knowledgeItems) {
    if (norm.includes(k.item_name_norm) && k.item_name_norm.length > bestLen) {
      bestLen = k.item_name_norm.length;
      bestContains = k;
    }
  }
  if (bestContains) return { ...bestContains, matchType: "contains" };

  return null;
}

// ─── Historical Sheet Parser ───────────────────────────────────────────────────

/**
 * Parse rows from a historical Items Sheet export into knowledge_items shape.
 *
 * @param {Array<Object>} rows  - Raw objects from XLSX parse (header row as keys)
 * @param {string} clientId
 * @param {string} authoritativeColumn - 'account' | 'purchase_account'
 * @returns {{ items: Array, conflicts: Array, skipped: number }}
 */
export function buildKnowledgeFromSheet(rows, clientId, authoritativeColumn = "account") {
  const items = [];
  const conflicts = [];
  let skipped = 0;

  // Normalise header keys from the sheet
  const normaliseKey = (k) =>
    String(k || "")
      .toLowerCase()
      .replace(/[^a-z0-9]/g, "_")
      .replace(/_+/g, "_")
      .replace(/^_|_$/g, "");

  for (const rawRow of rows) {
    // Re-key with normalised header names
    const row = {};
    for (const [k, v] of Object.entries(rawRow)) {
      row[normaliseKey(k)] = v;
    }

    // Try multiple column name variants
    const itemName =
      row.item_name || row.item || row.product_name || row.product || row.description || "";
    const account =
      row.account_name || row.account || row.expense_account || row.account_head || "";
    const purchaseAccount =
      row.purchase_account ||
      row.purchase_accounts ||
      row.purchase_account_name ||
      row.purchase_acct ||
      "";

    if (!itemName) {
      skipped++;
      continue;
    }

    const rawNorm = normalizeItemName(itemName);
    if (!rawNorm) {
      skipped++;
      continue;
    }

    // Determine authoritative account: auto-detect purchase_account if available, fallback to account
    const auth = purchaseAccount ? String(purchaseAccount).trim() : (account ? String(account).trim() : "");

    if (!auth) {
      skipped++;
      continue;
    }

    // Detect conflict
    const conflictFlag =
      account && purchaseAccount && account.trim() !== purchaseAccount.trim();

    const knowledgeItem = {
      client_id: clientId,
      item_name_raw: String(itemName).trim(),
      item_name_norm: rawNorm,
      account_head: String(auth).trim(),
      purchase_account: purchaseAccount ? String(purchaseAccount).trim() : null,
      conflict_flag: Boolean(conflictFlag),
      source: "historical_import",
      verified: true,
      confidence: 100,
      notes: conflictFlag
        ? `Account="${account}" vs PurchaseAccount="${purchaseAccount}"`
        : null,
    };

    if (conflictFlag) {
      conflicts.push(knowledgeItem);
    }
    items.push(knowledgeItem);
  }

  return { items, conflicts, skipped };
}

/**
 * Deduplicate knowledge items by item_name_norm.
 * If two rows have the same normalised name but different accounts → keep both
 * in conflicts array, pick the one from the authoritative column for the main list.
 */
export function deduplicateKnowledge(items) {
  const seen = new Map();
  const deduped = [];
  const duplicates = [];

  for (const item of items) {
    const key = item.item_name_norm;
    if (seen.has(key)) {
      const existing = seen.get(key);
      if (existing.account_head !== item.account_head) {
        duplicates.push({ existing, incoming: item });
      }
      // Skip — first occurrence wins
    } else {
      seen.set(key, item);
      deduped.push(item);
    }
  }

  return { deduped, duplicates };
}
