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
    const { apiKey, model, testOnly } = await req.json();

    if (!apiKey) {
      return NextResponse.json({ error: "Please enter a Groq API key." }, { status: 400 });
    }

    // ── 1. Fetch active models directly from Groq ────────────────────────────
    const modelsRes = await fetch("https://api.groq.com/openai/v1/models", {
      headers: { Authorization: `Bearer ${apiKey}` },
      cache: "no-store"
    });

    if (modelsRes.status === 401) {
      return NextResponse.json(
        { error: "Invalid Groq API key. Please check your key at console.groq.com/keys." },
        { status: 401 }
      );
    }

    if (!modelsRes.ok) {
      const errBody = await modelsRes.text();
      let errMsg = `Groq models request failed (${modelsRes.status})`;
      try { errMsg = JSON.parse(errBody)?.error?.message || errMsg; } catch { /* ignore */ }
      return NextResponse.json({ error: errMsg }, { status: modelsRes.status });
    }

    const modelsData = await modelsRes.json();
    const rawList = Array.isArray(modelsData.data) ? modelsData.data : [];

    if (!rawList.length) {
      return NextResponse.json(
        { error: "No models found for this Groq API key. Please verify your Groq account status." },
        { status: 400 }
      );
    }

    // Filter active text/chat models (exclude audio/whisper/guard)
    const availableModels = rawList
      .filter(m => m.active !== false)
      .map(m => m.id)
      .filter(id => {
        const lower = id.toLowerCase();
        const hasBadPrefix = EXCLUDED_PREFIXES.some(p => lower.startsWith(p));
        const hasBadKeyword = EXCLUDED_KEYWORDS.some(k => lower.includes(k));
        return !hasBadPrefix && !hasBadKeyword;
      });

    // Sort by preferred order
    availableModels.sort((a, b) => {
      const aIdx = PREFERRED_ORDER.findIndex(p => a.toLowerCase().includes(p.toLowerCase()));
      const bIdx = PREFERRED_ORDER.findIndex(p => b.toLowerCase().includes(p.toLowerCase()));
      const aScore = aIdx >= 0 ? aIdx : 999;
      const bScore = bIdx >= 0 ? bIdx : 999;
      return aScore - bScore;
    });

    if (!availableModels.length) {
      return NextResponse.json(
        { error: "No text/chat models found for this Groq key. Found: " + rawList.map(m => m.id).slice(0, 3).join(", ") },
        { status: 400 }
      );
    }

    // If caller only wants model discovery (no specific model requested for testing)
    if (!model) {
      return NextResponse.json({
        success: true,
        availableModels,
        defaultModel: availableModels[0]
      });
    }

    // ── 2. Verify selected model exists for this key ─────────────────────────
    if (!availableModels.includes(model)) {
      return NextResponse.json(
        {
          error: `Model '${model}' is not available for this API key. Available models: ${availableModels.slice(0, 4).join(", ")}`,
          availableModels
        },
        { status: 404 }
      );
    }

    // ── 3. Live test completion ping with the selected model ─────────────────
    const testRes = await fetch("https://api.groq.com/openai/v1/chat/completions", {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        Authorization: `Bearer ${apiKey}`
      },
      cache: "no-store",
      body: JSON.stringify({
        model,
        messages: [
          { role: "system", content: "You are an API validator. Reply with JSON." },
          { role: "user", content: "Ping. Return {\"status\":\"ok\"}." }
        ],
        temperature: 0.1,
        max_tokens: 30,
        response_format: { type: "json_object" }
      })
    });

    if (testRes.status === 429) {
      return NextResponse.json(
        { error: `Rate limit reached on model '${model}'. Please wait a moment or try another model.` },
        { status: 429 }
      );
    }

    if (!testRes.ok) {
      const errText = await testRes.text();
      let errMsg = `Live test failed for '${model}' (${testRes.status})`;
      try { errMsg = JSON.parse(errText)?.error?.message || errMsg; } catch { /* ignore */ }
      return NextResponse.json({ error: errMsg, availableModels }, { status: testRes.status });
    }

    // All 3 checks verified!
    return NextResponse.json({
      success: true,
      verified: true,
      model,
      availableModels,
      config: {
        provider: "groq",
        model,
        verified: true,
        verifiedAt: new Date().toISOString(),
        apiKeyConfigured: true
      }
    });

  } catch (err) {
    return NextResponse.json({ error: err.message || "Internal error verifying AI configuration." }, { status: 500 });
  }
}
