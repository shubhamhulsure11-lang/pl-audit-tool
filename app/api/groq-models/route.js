import { NextResponse } from "next/server";

// Non-text/audio models that cannot be used for text analysis
const EXCLUDED_PREFIXES = [
  "whisper-",
  "distil-whisper-",
  "canopylabs/",
  "playai-",
  "playht-"
];

const EXCLUDED_KEYWORDS = [
  "whisper",
  "tts",
  "embedding",
  "embed",
  "guard"
];

// Preferred priority order for restaurant accounting audits
const PREFERRED_ORDER = [
  "llama-3.1-8b-instant",
  "llama-3.3-70b-versatile",
  "llama-3.1-70b-versatile",
  "llama3-70b-8192",
  "llama3-8b-8192",
  "mixtral-8x7b-32768",
  "gemma2-9b-it",
  "gemma-7b-it",
  "openai/gpt-oss-20b",
  "openai/gpt-oss-120b"
];

export async function POST(req) {
  try {
    const { apiKey } = await req.json();
    if (!apiKey) return NextResponse.json({ error: "No API key provided." }, { status: 400 });

    const res = await fetch("https://api.groq.com/openai/v1/models", {
      headers: { Authorization: `Bearer ${apiKey}` },
      cache: "no-store"
    });

    if (res.status === 401) {
      return NextResponse.json({ error: "Invalid Groq API key. Please check your key at console.groq.com." }, { status: 401 });
    }
    if (!res.ok) {
      return NextResponse.json({ error: `Groq error (${res.status})` }, { status: res.status });
    }

    const data = await res.json();
    const rawList = Array.isArray(data.data) ? data.data : [];

    // Filter available models strictly to active text/chat models
    const filtered = rawList
      .filter(m => m.active !== false)
      .map(m => m.id)
      .filter(id => {
        const lower = id.toLowerCase();
        const hasBadPrefix = EXCLUDED_PREFIXES.some(p => lower.startsWith(p));
        const hasBadKeyword = EXCLUDED_KEYWORDS.some(k => lower.includes(k));
        return !hasBadPrefix && !hasBadKeyword;
      });

    // Sort according to preferred production priority
    filtered.sort((a, b) => {
      const aIdx = PREFERRED_ORDER.findIndex(p => a.toLowerCase().includes(p.toLowerCase()));
      const bIdx = PREFERRED_ORDER.findIndex(p => b.toLowerCase().includes(p.toLowerCase()));
      const aScore = aIdx >= 0 ? aIdx : 999;
      const bScore = bIdx >= 0 ? bIdx : 999;
      return aScore - bScore;
    });

    if (!filtered.length) {
      return NextResponse.json({ error: "No text/chat models found for this Groq key." }, { status: 400 });
    }

    return NextResponse.json({ models: filtered });
  } catch (err) {
    return NextResponse.json({ error: err.message || "Internal error" }, { status: 500 });
  }
}
