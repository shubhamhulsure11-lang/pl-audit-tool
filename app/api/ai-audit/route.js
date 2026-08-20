import { NextResponse } from "next/server";

function parseRetryAfter(errorMsg) {
  const m = String(errorMsg || "").match(/try again in (\d+(?:\.\d+)?)\s*s/i);
  return m ? Math.ceil(parseFloat(m[1])) : 30;
}

export async function POST(req) {
  try {
    const body = await req.json();
    const { singleItem, availableAccounts = [], apiKey, model } = body;

    if (!apiKey) return NextResponse.json({ error: "No API key. Use the ⚙️ AI Setup button to configure." }, { status: 400 });
    if (!model) return NextResponse.json({ error: "No model selected. Use the ⚙️ AI Setup button." }, { status: 400 });
    if (!singleItem) return NextResponse.json({ error: "No item provided." }, { status: 400 });

    const { item, vendor, actualAccount } = singleItem;

    // ── ULTRA-COMPACT PROMPT — target < 250 tokens total ──────────────────────
    // Only top 8 accounts to keep token count low
    const accountList = (availableAccounts || []).slice(0, 8).join(", ");

    const prompt = `Restaurant P&L auditor (India/Zoho Books). Check if ONE item is in the correct account.

Item: "${item}"
Vendor: "${vendor}"
Current account in Zoho: "${actualAccount}"
Accounts available: ${accountList || "Groceries, Poultry, Dairy, Beverages, Liquor Purchases, Cleaning and housekeeping, Packing materials"}

Quick rules:
- Knorr/broth/bouillon/seasoning powder → Groceries (NOT Poultry)
- Oat milk/almond milk/soy milk → Dairy
- Monin/Malas/syrup/crush → Beverages (NOT Fresh Fruits)
- Classic Connect/Marlboro/Gold Flake → Cigarette purchases
- Soap oil/dishwash/detergent → Cleaning (NOT Groceries)
- Tissue/napkin/foil/carry bag → Packing materials

Reply JSON only (no markdown, no explanation outside JSON):
{"ok":true} if correctly classified, OR
{"ok":false,"suggest":"exact account name from list","why":"one sentence"}`;

    const res = await fetch("https://api.groq.com/openai/v1/chat/completions", {
      method: "POST",
      headers: { "Content-Type": "application/json", Authorization: `Bearer ${apiKey}` },
      body: JSON.stringify({
        model,
        messages: [{ role: "user", content: prompt }],
        temperature: 0.1,
        max_tokens: 80
      })
    });

    if (res.status === 429) {
      const errText = await res.text();
      let errMsg = "";
      try { errMsg = JSON.parse(errText)?.error?.message || ""; } catch { /* ignore */ }
      const retryAfter = parseRetryAfter(errMsg);
      return NextResponse.json({ error: "Rate limited", retryAfter }, { status: 429 });
    }

    if (!res.ok) {
      const errText = await res.text();
      let errMsg = `Groq error (${res.status})`;
      try { errMsg = JSON.parse(errText)?.error?.message || errMsg; } catch { /* ignore */ }
      return NextResponse.json({ error: errMsg }, { status: res.status });
    }

    const data = await res.json();
    let text = (data.choices?.[0]?.message?.content || "").trim();

    // Strip markdown fences if AI added them despite instructions
    const fenceMatch = text.match(/```(?:json)?\s*([\s\S]*?)\s*```/);
    if (fenceMatch) text = fenceMatch[1].trim();

    // Strip leading/trailing non-JSON characters
    const jsonStart = text.indexOf("{");
    const jsonEnd = text.lastIndexOf("}");
    if (jsonStart >= 0 && jsonEnd > jsonStart) text = text.slice(jsonStart, jsonEnd + 1);

    try {
      const result = JSON.parse(text);
      return NextResponse.json({ success: true, result, model });
    } catch {
      // If we can't parse, treat as "ok" (don't show false positive)
      return NextResponse.json({ success: true, result: { ok: true }, model, parseWarning: text });
    }
  } catch (err) {
    return NextResponse.json({ error: err.message || "Internal error" }, { status: 500 });
  }
}
