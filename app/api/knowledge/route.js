import { NextResponse } from "next/server";
import { getServiceClient } from "@/app/lib/supabase";

/**
 * GET /api/knowledge?client_id=xxx
 * Returns all knowledge items for a client
 */
export async function GET(request) {
  try {
    const { searchParams } = new URL(request.url);
    const client_id = searchParams.get("client_id");

    if (!client_id) {
      return NextResponse.json({ error: "client_id is required" }, { status: 400 });
    }

    const db = getServiceClient();
    const { data, error } = await db
      .from("client_item_knowledge")
      .select("id, item_name_raw, item_name_norm, account_head, purchase_account, conflict_flag, source, verified, confidence, notes, updated_at")
      .eq("client_id", client_id)
      .order("item_name_norm", { ascending: true });

    if (error) throw error;

    return NextResponse.json({ knowledge: data, count: data.length });
  } catch (err) {
    console.error("[GET /api/knowledge]", err);
    return NextResponse.json({ error: err.message }, { status: 500 });
  }
}

/**
 * POST /api/knowledge — save a single human-approved classification
 * Body: { client_id, item_name_raw, item_name_norm, account_head, source, notes }
 */
export async function POST(request) {
  try {
    const body = await request.json();
    const { client_id, item_name_raw, account_head } = body;

    if (!client_id || !item_name_raw || !account_head) {
      return NextResponse.json(
        { error: "client_id, item_name_raw, and account_head are required" },
        { status: 400 }
      );
    }

    const db = getServiceClient();

    // Import normalise logic server-side
    const { normalizeItemName } = await import("@/app/lib/knowledge");
    const item_name_norm = normalizeItemName(item_name_raw);

    const row = {
      client_id,
      item_name_raw: String(item_name_raw).trim(),
      item_name_norm,
      account_head: String(account_head).trim(),
      source: body.source || "human_approved",
      verified: true,
      confidence: body.confidence ?? 100,
      notes: body.notes || null,
    };

    const { data, error } = await db
      .from("client_item_knowledge")
      .upsert(row, { onConflict: "client_id,item_name_norm" })
      .select()
      .single();

    if (error) throw error;
    return NextResponse.json({ knowledge_item: data }, { status: 201 });
  } catch (err) {
    console.error("[POST /api/knowledge]", err);
    return NextResponse.json({ error: err.message }, { status: 500 });
  }
}
