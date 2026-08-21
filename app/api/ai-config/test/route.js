import { NextResponse } from "next/server";

// Standard production chat models only (NO reasoning models with <think> tokens, NO audio/embed/guard models)
const SAFE_PREFIXES = ["llama", "mixtral", "gemma"];
const EXCLUDE_PATTERNS = [
  "whisper", "tts", "distil", "embed", "vision", "playai", "playht",
  "guard", "canopylabs", "orpheus", "arabic", "preview", "speculative", "openai",
  "deepseek", "qwen", "r1", "reasoning", "think"
];

// Preferred priority order for restaurant accounting data audit
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
    const { apiKey, model } = await req.json();

    if (!apiKey) {
      return NextResponse.json({ error: "Please provide a Groq API key." }, { status: 400 });
    }

    // ── Check 1: Validate API key and fetch actual active models from Groq ──
    const modelsRes = await fetch("https://api.groq.com/openai/v1/models", {
      headers: { Authorization: `Bearer ${apiKey}` }
    });

    if (modelsRes.status === 401) {
      return NextResponse.json(
        { error: "Invalid Groq API key. Please check your key at console.groq.com/keys." },
        { status: 401 }
      );
    }

    if (!modelsRes.ok) {
      const errBody = await modelsRes.text();
      let errMsg = `Groq error (${modelsRes.status})`;
      try { errMsg = JSON.parse(errBody)?.error?.message || errMsg; } catch { /* ignore */ }
      return NextResponse.json({ error: errMsg }, { status: modelsRes.status });
    }

    const modelsData = await modelsRes.json();
    const rawList = modelsData.data || [];

    if (!rawList.length) {
      return NextResponse.json(
        { error: "No models found for this Groq API key. Please verify your Groq account status." },
        { status: 400 }
      );
    }

    // Filter available models strictly to valid production chat models
    const availableModels = rawList
      .map(m => m.id)
      .filter(id => {
        const lower = id.toLowerCase();
        const isSafe = SAFE_PREFIXES.some(p => lower.startsWith(p));
        const isExcluded = EXCLUDE_PATTERNS.some(p => lower.includes(p));
        return isSafe && !isExcluded;
      });

    // Sort according to preferred production priority
    availableModels.sort((a, b) => {
      const aIdx = MODEL_PRIORITY.findIndex(p => a.toLowerCase().includes(p.toLowerCase()));
      const bIdx = MODEL_PRIORITY.findIndex(p => b.toLowerCase().includes(p.toLowerCase()));
      const aScore = aIdx >= 0 ? aIdx : 99;
      const bScore = bIdx >= 0 ? bIdx : 99;
      return aScore - bScore;
    });

    if (!availableModels.length) {
      return NextResponse.json(
        { error: "No compatible chat models found for this Groq key. Groq may be undergoing maintenance." },
        { status: 400 }
      );
    }

    // ── Check 2: Verify requested model exists in key's available models ──
    let targetModel = model;
    if (targetModel) {
      if (!availableModels.includes(targetModel)) {
        return NextResponse.json(
          {
            error: `Model '${targetModel}' is not available for this API key. Available models: ${availableModels.slice(0, 4).join(", ")}`,
            availableModels
          },
          { status: 404 }
        );
      }
    } else {
      // Pick top recommended model that actually exists
      targetModel = availableModels[0];
    }

    // ── Check 3: Actually perform a real tiny completion test with the selected model ──
    const testRes = await fetch("https://api.groq.com/openai/v1/chat/completions", {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        Authorization: `Bearer ${apiKey}`
      },
      body: JSON.stringify({
        model: targetModel,
        messages: [
          { role: "system", content: "You are an API verification validator. Return JSON only." },
          { role: "user", content: "Verify connection. Return {\"status\":\"ok\"}." }
        ],
        temperature: 0.1,
        max_tokens: 30,
        response_format: { type: "json_object" }
      })
    });

    if (testRes.status === 429) {
      return NextResponse.json(
        { error: `Rate limit hit on Groq during model test for '${targetModel}'. Please wait a moment or choose another model.` },
        { status: 429 }
      );
    }

    if (!testRes.ok) {
      const errText = await testRes.text();
      let errMsg = `Model test failed for '${targetModel}' (${testRes.status})`;
      try { errMsg = JSON.parse(errText)?.error?.message || errMsg; } catch { /* ignore */ }
      return NextResponse.json({ error: errMsg, availableModels }, { status: testRes.status });
    }

    // All 3 checks passed!
    return NextResponse.json({
      success: true,
      verified: true,
      model: targetModel,
      availableModels,
      config: {
        provider: "groq",
        model: targetModel,
        verified: true,
        verifiedAt: new Date().toISOString(),
        apiKeyConfigured: true
      }
    });

  } catch (err) {
    return NextResponse.json({ error: err.message || "Internal error verifying AI configuration." }, { status: 500 });
  }
}
