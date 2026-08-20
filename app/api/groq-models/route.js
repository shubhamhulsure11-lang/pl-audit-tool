import { NextResponse } from "next/server";

// Only return safe, standard chat-capable models
const SAFE_PREFIXES = ["llama", "mixtral", "gemma", "deepseek", "qwen"];
const EXCLUDE_PATTERNS = [
  "whisper", "tts", "distil", "embed", "vision", "playai", "playht",
  "guard", "canopylabs", "orpheus", "arabic", "preview", "speculative", "openai"
];

export async function POST(req) {
  try {
    const { apiKey } = await req.json();
    if (!apiKey) return NextResponse.json({ error: "No API key" }, { status: 400 });

    const res = await fetch("https://api.groq.com/openai/v1/models", {
      headers: { Authorization: `Bearer ${apiKey}` }
    });

    if (res.status === 401) {
      return NextResponse.json({ error: "Invalid API key. Check at console.groq.com." }, { status: 401 });
    }
    if (!res.ok) {
      return NextResponse.json({ error: `Groq error (${res.status})` }, { status: res.status });
    }

    const data = await res.json();
    const models = (data.data || [])
      .map(m => m.id)
      .filter(id => {
        const lower = id.toLowerCase();
        const isSafe = SAFE_PREFIXES.some(p => lower.startsWith(p));
        const isExcluded = EXCLUDE_PATTERNS.some(p => lower.includes(p));
        return isSafe && !isExcluded;
      })
      .sort((a, b) => {
        // Prefer instant/fast models first
        const aScore = a.includes("instant") ? 0 : a.includes("8b") ? 1 : a.includes("70b") ? 2 : 3;
        const bScore = b.includes("instant") ? 0 : b.includes("8b") ? 1 : b.includes("70b") ? 2 : 3;
        return aScore - bScore;
      });

    if (!models.length) {
      return NextResponse.json({
        error: "No compatible chat models found for this API key. Make sure your Groq account is active."
      }, { status: 400 });
    }

    return NextResponse.json({ models });
  } catch (err) {
    return NextResponse.json({ error: err.message || "Internal error" }, { status: 500 });
  }
}
