import { NextResponse } from "next/server";
import { getServiceClient } from "@/app/lib/supabase";

/**
 * GET /api/periods?client_id=xxx
 * Returns all audit periods (monthly sessions) for a client
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
      .from("client_periods")
      .select("*")
      .eq("client_id", client_id)
      .order("created_at", { ascending: false });

    if (error) throw error;
    return NextResponse.json({ periods: data });
  } catch (err) {
    return NextResponse.json({ error: err.message }, { status: 500 });
  }
}

/**
 * POST /api/periods — create a new audit period
 * Body: { client_id, period_label, file_name, row_count }
 */
export async function POST(request) {
  try {
    const { client_id, period_label, file_name, row_count } = await request.json();

    if (!client_id || !period_label || !file_name) {
      return NextResponse.json(
        { error: "client_id, period_label, and file_name are required" },
        { status: 400 }
      );
    }

    const db = getServiceClient();
    const { data, error } = await db
      .from("client_periods")
      .insert({ client_id, period_label, file_name, row_count: row_count || 0 })
      .select()
      .single();

    if (error) throw error;
    return NextResponse.json({ period: data }, { status: 201 });
  } catch (err) {
    return NextResponse.json({ error: err.message }, { status: 500 });
  }
}
