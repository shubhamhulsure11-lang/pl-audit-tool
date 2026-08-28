import { NextResponse } from "next/server";
import { getServiceClient } from "@/app/lib/supabase";

// GET /api/clients/[id] — get single client
export async function GET(request, { params }) {
  try {
    const db = getServiceClient();
    const { data, error } = await db
      .from("clients")
      .select("*")
      .eq("id", params.id)
      .single();

    if (error) throw error;
    if (!data) return NextResponse.json({ error: "Client not found" }, { status: 404 });
    return NextResponse.json({ client: data });
  } catch (err) {
    return NextResponse.json({ error: err.message }, { status: 500 });
  }
}

// PATCH /api/clients/[id] — update client fields
export async function PATCH(request, { params }) {
  try {
    const body = await request.json();
    const db = getServiceClient();
    const { data, error } = await db
      .from("clients")
      .update(body)
      .eq("id", params.id)
      .select()
      .single();

    if (error) throw error;
    return NextResponse.json({ client: data });
  } catch (err) {
    return NextResponse.json({ error: err.message }, { status: 500 });
  }
}

// DELETE /api/clients/[id] — remove client and all related data (cascade)
export async function DELETE(request, { params }) {
  try {
    const db = getServiceClient();
    const { error } = await db.from("clients").delete().eq("id", params.id);
    if (error) throw error;
    return NextResponse.json({ success: true });
  } catch (err) {
    return NextResponse.json({ error: err.message }, { status: 500 });
  }
}
