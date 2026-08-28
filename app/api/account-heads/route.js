import { NextResponse } from "next/server";
import { getServiceClient } from "@/app/lib/supabase";

/**
 * GET /api/account-heads?client_id=xxx
 * Returns all account heads for a client
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
      .from("account_heads")
      .select("id, name, category, is_active")
      .eq("client_id", client_id)
      .eq("is_active", true)
      .order("name", { ascending: true });

    if (error) throw error;
    return NextResponse.json({ account_heads: data });
  } catch (err) {
    return NextResponse.json({ error: err.message }, { status: 500 });
  }
}

/**
 * POST /api/account-heads — add a new account head for a client
 */
export async function POST(request) {
  try {
    const { client_id, name, category } = await request.json();
    if (!client_id || !name) {
      return NextResponse.json({ error: "client_id and name are required" }, { status: 400 });
    }

    const db = getServiceClient();
    const { data, error } = await db
      .from("account_heads")
      .upsert({ client_id, name: name.trim(), category: category || null }, { onConflict: "client_id,name" })
      .select()
      .single();

    if (error) throw error;
    return NextResponse.json({ account_head: data }, { status: 201 });
  } catch (err) {
    return NextResponse.json({ error: err.message }, { status: 500 });
  }
}
