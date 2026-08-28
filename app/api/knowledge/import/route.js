import { NextResponse } from "next/server";
import { getServiceClient } from "@/app/lib/supabase";
import { buildKnowledgeFromSheet, deduplicateKnowledge } from "@/app/lib/knowledge";

/**
 * POST /api/knowledge/import
 *
 * Body:
 * {
 *   client_id: string,
 *   authoritative_column: 'account' | 'purchase_account',
 *   rows: Array<Object>  // raw objects from XLSX parse
 * }
 *
 * This is the ONE-TIME initialisation endpoint.
 * After it succeeds, clients.knowledge_initialized is set to true.
 * It will NOT be called again during normal monthly operations.
 */
export async function POST(request) {
  try {
    const { client_id, authoritative_column, rows } = await request.json();

    if (!client_id) {
      return NextResponse.json({ error: "client_id is required" }, { status: 400 });
    }
    if (!rows || !Array.isArray(rows) || rows.length === 0) {
      return NextResponse.json({ error: "rows array is required" }, { status: 400 });
    }

    const db = getServiceClient();

    // Verify client exists
    const { data: client, error: clientErr } = await db
      .from("clients")
      .select("id, knowledge_initialized")
      .eq("id", client_id)
      .single();

    if (clientErr || !client) {
      return NextResponse.json({ error: "Client not found" }, { status: 404 });
    }

    // Build knowledge items from sheet rows
    const auth = authoritative_column || "account";
    const { items, conflicts, skipped } = buildKnowledgeFromSheet(rows, client_id, auth);

    if (items.length === 0) {
      return NextResponse.json(
        { error: "No valid items found in the sheet. Check column names." },
        { status: 400 }
      );
    }

    // Deduplicate
    const { deduped, duplicates } = deduplicateKnowledge(items);

    // Upsert into client_item_knowledge
    // ON CONFLICT (client_id, item_name_norm) → update account_head
    const batchSize = 500;
    let insertedCount = 0;

    for (let i = 0; i < deduped.length; i += batchSize) {
      const batch = deduped.slice(i, i + batchSize);
      const { error: insertErr } = await db
        .from("client_item_knowledge")
        .upsert(batch, { onConflict: "client_id,item_name_norm" });

      if (insertErr) throw insertErr;
      insertedCount += batch.length;
    }

    // Also upsert account_heads from the unique account names
    const uniqueAccounts = [...new Set(deduped.map((i) => i.account_head).filter(Boolean))];
    if (uniqueAccounts.length > 0) {
      const accountRows = uniqueAccounts.map((name) => ({ client_id, name }));
      await db
        .from("account_heads")
        .upsert(accountRows, { onConflict: "client_id,name" });
    }

    // Mark client as initialized
    const { error: updateErr } = await db
      .from("clients")
      .update({
        knowledge_initialized: true,
        initialized_at: new Date().toISOString(),
        init_row_count: rows.length,
        init_item_count: insertedCount,
        authoritative_column: auth,
      })
      .eq("id", client_id);

    if (updateErr) throw updateErr;

    return NextResponse.json({
      success: true,
      inserted: insertedCount,
      conflicts: conflicts.length,
      duplicates: duplicates.length,
      skipped,
      accounts_imported: uniqueAccounts.length,
    });
  } catch (err) {
    console.error("[POST /api/knowledge/import]", err);
    return NextResponse.json({ error: err.message }, { status: 500 });
  }
}
