import { createClient } from "@supabase/supabase-js";

// Browser-safe client (anon key, respects RLS)
export function getSupabaseClient() {
  const url = process.env.NEXT_PUBLIC_SUPABASE_URL;
  const key = process.env.NEXT_PUBLIC_SUPABASE_ANON_KEY;
  if (!url || !key) throw new Error("Missing Supabase environment variables. Check .env.local");
  return createClient(url, key);
}

// Lazy singleton for client-side usage
let _supabase = null;
export const supabase = new Proxy({}, {
  get(_, prop) {
    if (!_supabase) _supabase = getSupabaseClient();
    return _supabase[prop];
  }
});

// Server-only admin client (service_role key, bypasses RLS)
// ONLY import this in /app/api/** routes — never in client components
export function getServiceClient() {
  const url = process.env.NEXT_PUBLIC_SUPABASE_URL;
  const anonKey = process.env.NEXT_PUBLIC_SUPABASE_ANON_KEY;
  const serviceKey = process.env.SUPABASE_SERVICE_ROLE_KEY;
  if (!url || !anonKey) throw new Error("Missing Supabase environment variables. Check .env.local");
  if (!serviceKey) throw new Error("Missing SUPABASE_SERVICE_ROLE_KEY — only available server-side");
  return createClient(url, serviceKey, {
    auth: { persistSession: false },
  });
}
