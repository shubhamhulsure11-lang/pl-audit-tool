import { NextResponse } from "next/server";
import { getServiceClient } from "@/app/lib/supabase";

/**
 * POST /api/classification/save
 *
 * Saves classification decisions from an audit session.
 * Called at end-of-audit when accountant approves/rejects items.
 *
 * Body: {
 *   period_id: string,
 *   client_id: string,
 *   decisions: Array<{
 *     item_name: string,
 *     vendor: string,
 *     detected_account: string,
 *     knowledge_match: boolean,
 *     knowledge_account: string | null,
 *     rule_match: boolean,
 *     rule_account: string | null,
 *     ai_called: boolean,
 *     ai_account: string | null,
 *     ai_confidence: number | null,
 *     final_account: string,
 *     is_misclassified: boolean,
 *     human_reviewed: boolean,
 *     human_approved: boolean | null,
 *     total_amount: number,
 *   }>
 * }
 */
export async function POST(request) {
  try {
    const { period_id, client_id, decisions } = await request.json();

    if (!period_id || !client_id || !Array.isArray(decisions)) {
      return NextResponse.json({ error: "period_id, client_id, decisions[] required" }, { status: 400 });
    }

    const db = getServiceClient();

    // Bulk insert classification history
    const rows = decisions.map((d) => ({
      client_id,
      period_id,
      item_name: d.item_name,
      vendor: d.vendor || null,
      detected_account: d.detected_account || null,
      knowledge_match: d.knowledge_match ?? false,
      knowledge_account: d.knowledge_account || null,
      rule_match: d.rule_match ?? false,
      rule_account: d.rule_account || null,
      ai_called: d.ai_called ?? false,
      ai_account: d.ai_account || null,
      ai_confidence: d.ai_confidence ?? null,
      final_account: d.final_account || null,
      is_misclassified: d.is_misclassified ?? false,
      human_reviewed: d.human_reviewed ?? false,
      human_approved: d.human_approved ?? null,
      saved_to_knowledge: false,
      total_amount: d.total_amount ?? 0,
    }));

    const { error: histErr } = await db
      .from("classification_history")
      .insert(rows);

    if (histErr) throw histErr;

    // For human_approved items — save to knowledge base
    const approved = decisions.filter(
      (d) => d.human_reviewed && d.human_approved && d.final_account
    );

    const { normalizeItemName } = await import("@/app/lib/knowledge");
    let savedToKb = 0;

    if (approved.length > 0) {
      const kbRows = approved.map((d) => ({
        client_id,
        item_name_raw: d.item_name,
        item_name_norm: normalizeItemName(d.item_name),
        account_head: d.final_account,
        source: "human_approved",
        verified: true,
        confidence: 100,
        notes: `Approved during period ${period_id}`,
      }));

      const { error: kbErr } = await db
        .from("client_item_knowledge")
        .upsert(kbRows, { onConflict: "client_id,item_name_norm" });

      if (kbErr) throw kbErr;
      savedToKb = kbRows.length;
    }

    // Mark period as completed
    await db
      .from("client_periods")
      .update({ status: "completed" })
      .eq("id", period_id);

    return NextResponse.json({
      success: true,
      decisions_saved: rows.length,
      saved_to_knowledge: savedToKb,
    });
  } catch (err) {
    console.error("[POST /api/classification/save]", err);
    return NextResponse.json({ error: err.message }, { status: 500 });
  }
}
