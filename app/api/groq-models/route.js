import { NextResponse } from "next/server";

// Standard production chat models only (NO reasoning models with <think> tokens, NO audio/embed/guard models)
const SAFE_PREFIXES = ["llama", "mixtral", "gemma"];
const EXCLUDE_PATTERNS = [
  "whisper", "tts", "distil", "embed", "vision", "playai", "playht",
  "guard", "canopylabs", "orpheus", "arabic", "preview", "speculative", "openai",
  "deepseek", "qwen", "r1", "reasoning", "think"
];

// Preferred priority order
const MODEL_PRIORITY = [
  "llama-3.1-8b-instant",
  "llama-3.3-70b-versatile",
  "llama-3.1-70b-versatile",
  "llama3-70b-8192",
  "llama3-8b-8192",
  "mixtral-8x7b-32768",
  "gemma2-9b-it",
  "gemma-7b-it"
];

export async function POST(req) {
  try {
    const { apiKey } = await req.json();
    if (!apiKey) return NextResponse.json({ error: "No API key provided." }, { status: 400 });

    const res = await fetch("https://api.groq.com/openai/v1/models", {
      headers: { Authorization: `Bearer ${apiKey}` }
    });

    if (res.status === 401) {
      return NextResponse.json({ error: "Invalid Groq API key. Please check your key at console.groq.com." }, { status: 401 });
    }
    if (!res.ok) {
      return NextResponse.json({ error: `Groq error (${res.status})` }, { status: res.status });
    }

    const data = await res.json();
    const rawList = data.data || [];

    // Filter available models strictly
    const filtered = rawList
      .map(m => m.id)
      .filter(id => {
        const lower = id.toLowerCase();
        const isSafe = SAFE_PREFIXES.some(p => lower.startsWith(p));
        const isExcluded = EXCLUDE_PATTERNS.some(p => lower.includes(p));
        return isSafe && !isExcluded;
      });

    // Sort according to preferred production priority
    filtered.sort((a, b) => {
      const aIdx = MODEL_PRIORITY.findIndex(p => a.toLowerCase().includes(p.toLowerCase()));
      const bIdx = MODEL_PRIORITY.findIndex(p => b.toLowerCase().includes(p.toLowerCase()));
      const aScore = aIdx >= 0 ? aIdx : 99;
      const bScore = bIdx >= 0 ? bIdx : 99;
      return aScore - bScore;
    });

    if (!filtered.length) {
      return NextResponse.json({ error: "No compatible production chat models found for this Groq key." }, { status: 400 });
    }

    return NextResponse.json({ models: filtered });
  } catch (err) {
    return NextResponse.json({ error: err.message || "Internal error" }, { status: 500 });
  }
}
