import { NextResponse } from "next/server";
import { getServiceClient } from "@/app/lib/supabase";

// GET /api/clients — list all clients
export async function GET() {
  try {
    const db = getServiceClient();
    const { data, error } = await db
      .from("clients")
      .select("*")
      .order("created_at", { ascending: false });

    if (error) throw error;
    return NextResponse.json({ clients: data });
  } catch (err) {
    console.error("[GET /api/clients]", err);
    return NextResponse.json({ error: err.message }, { status: 500 });
  }
}

// POST /api/clients — create a new client
export async function POST(request) {
  try {
    const { display_name } = await request.json();
    if (!display_name?.trim()) {
      return NextResponse.json({ error: "display_name is required" }, { status: 400 });
    }

    const db = getServiceClient();
    const { data, error } = await db
      .from("clients")
      .insert({ display_name: display_name.trim() })
      .select()
      .single();

    if (error) throw error;
    return NextResponse.json({ client: data }, { status: 201 });
  } catch (err) {
    console.error("[POST /api/clients]", err);
    return NextResponse.json({ error: err.message }, { status: 500 });
  }
}
