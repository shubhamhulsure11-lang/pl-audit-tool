import { NextResponse } from "next/server";

// ─── Compact system prompt (~500 tokens) ─────────────────────────────────────
// Designed to fit comfortably within llama-3.1-8b-instant's 6K TPM budget
// while giving the model full context for accurate product identification.
const SYSTEM_PROMPT = `You are an accounting auditor for a restaurant/hospitality business.

TASK: Independently identify the ACTUAL PURCHASED PRODUCT and determine the correct account head.

CRITICAL — NEVER CLASSIFY BY ISOLATED KEYWORDS:
A word in an item name may be a brand, flavour, colour, size, design, shape, or model number — NOT the actual product. Understand the complete item description before classifying.

KEYWORD TRAP EXAMPLES (learn these patterns):
• "DIP BOWL ROUND WHITE APPLE" → product=BOWL (apple=design, not fruit) → Cutlery/crockery
• "MONIN WATERMELON 700ML" → Monin=syrup brand, watermelon=flavour → Beverages
• "VANILLA 4LTR" from dairy vendor → vanilla flavouring/extract → Groceries, NOT dairy
• "BANANA LEAF" → serving/packing material for restaurant use → Packing material, NOT vegetables
• "ARCOROC WHEAT BEER GLASS" → product=GLASS (beer=glass type) → Cutlery/crockery
• "BACARDI CLASSIC WHITE RUM" → product=RUM (classic=variant) → Liquor, NOT cigarettes

RULES:
1. Vendor name alone does NOT determine category
2. Evaluate the current account head independently — it may be wrong
3. Evaluate the suggested account head independently — it may also be wrong
4. Choose the most specific valid account from the provided accounts list
5. Never invent a new account head — only use accounts from the provided list
6. If genuinely unclear, return classification_status "REVIEW_REQUIRED"

CLASSIFICATION_STATUS values (return exactly one):
CURRENT_CORRECT | CURRENT_INCORRECT | SUGGESTION_CORRECT | SUGGESTION_INCORRECT | BOTH_INCORRECT | BOTH_CORRECT | REVIEW_REQUIRED

VERDICT values (for current_verdict and suggested_verdict):
CORRECT | INCORRECT | UNCERTAIN

Return ONLY valid JSON, no markdown, no text outside JSON:
{"product_type":"","brand":"","flavour_or_variant":"","intended_use":"","current_verdict":"","suggested_verdict":"","classification_status":"","ai_final_account_head":"","ai_reason":"","confidence":0,"review_required":false,"review_note":""}`;

// ─── Helpers ──────────────────────────────────────────────────────────────────

function parseRetryAfter(errorMsg) {
  const m = String(errorMsg || "").match(/try again in (\d+(?:\.\d+)?)\s*s/i);
  return m ? Math.ceil(parseFloat(m[1])) : 30;
}

async function callGroq(apiKey, model, userContent) {
  return fetch("https://api.groq.com/openai/v1/chat/completions", {
    method: "POST",
    headers: { "Content-Type": "application/json", Authorization: `Bearer ${apiKey}` },
    body: JSON.stringify({
      model,
      messages: [
        { role: "system", content: SYSTEM_PROMPT },
        { role: "user", content: userContent }
      ],
      temperature: 0.1,
      max_tokens: 320,
      response_format: { type: "json_object" }   // JSON mode: guaranteed JSON output
    })
  });
}

function parseResult(rawContent) {
  try {
    return JSON.parse(rawContent || "{}");
  } catch {
    return null;
  }
}

// ─── Route ────────────────────────────────────────────────────────────────────

export async function POST(req) {
  try {
    const body = await req.json();
    const { singleItem, availableAccounts = [], apiKey, model, fallbackModel } = body;

    if (!apiKey) return NextResponse.json({ error: "No API key. Use ⚙️ AI Setup." }, { status: 400 });
    if (!singleItem) return NextResponse.json({ error: "No item provided." }, { status: 400 });

    const { item, vendor, actualAccount, suggestedAccount, matchedKeyword, total } = singleItem;

    // Compact JSON payload — sent as the user message
    const userContent = JSON.stringify({
      item: item || "",
      vendor: vendor || "",
      current: actualAccount || "",
      suggested: suggestedAccount || "",
      amount: total || 0,
      detection: matchedKeyword || "",
      accounts: (availableAccounts || []).slice(0, 14)
    });

    // ── Stage 1: Primary verified model ─────────────────────────────────────
    const PRIMARY_MODEL = model || "llama-3.1-8b-instant";
    const res1 = await callGroq(apiKey, PRIMARY_MODEL, userContent);

    // Handle rate limit from Stage 1
    if (res1.status === 429) {
      const errText = await res1.text();
      let errMsg = "";
      try { errMsg = JSON.parse(errText)?.error?.message || ""; } catch { /* ignore */ }
      return NextResponse.json({ error: "Rate limited", retryAfter: parseRetryAfter(errMsg) }, { status: 429 });
    }

    if (!res1.ok) {
      const errText = await res1.text();
      let errMsg = `Groq error (${res1.status})`;
      try { errMsg = JSON.parse(errText)?.error?.message || errMsg; } catch { /* ignore */ }
      return NextResponse.json({ error: errMsg }, { status: res1.status });
    }

    const data1 = await res1.json();
    let result = parseResult(data1.choices?.[0]?.message?.content);

    if (!result || !result.classification_status) {
      // JSON mode should prevent this, but handle gracefully
      return NextResponse.json({
        error: "AI returned an unexpected response. Try again or change model in ⚙️ AI Setup."
      }, { status: 422 });
    }

    let escalated = false;
    let usedModel = PRIMARY_MODEL;

    // ── Stage 2: Escalate to 70B if low confidence or uncertain (only if primary is not already 70B) ──
    const needsEscalation = result.review_required === true ||
      result.classification_status === "REVIEW_REQUIRED" ||
      (typeof result.confidence === "number" && result.confidence < 75);

    const ESCALATION_MODEL = fallbackModel || (PRIMARY_MODEL.includes("70b") ? null : "llama-3.3-70b-versatile");

    if (needsEscalation && ESCALATION_MODEL && ESCALATION_MODEL !== PRIMARY_MODEL) {
      try {
        const res2 = await callGroq(apiKey, ESCALATION_MODEL, userContent);
        if (res2.ok) {
          const data2 = await res2.json();
          const result2 = parseResult(data2.choices?.[0]?.message?.content);
          if (result2?.classification_status) {
            result = result2;
            escalated = true;
            usedModel = ESCALATION_MODEL;
          }
        }
        // If Stage 2 fails (rate limit, model not found etc.), silently use Stage 1 result
      } catch {
        // Stage 2 network error — proceed with Stage 1 result
      }
    }

    return NextResponse.json({ success: true, result, model: usedModel, escalated });
  } catch (err) {
    return NextResponse.json({ error: err.message || "Internal error" }, { status: 500 });
  }
}
