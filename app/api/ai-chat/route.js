import { NextResponse } from "next/server";

function parseRetryAfter(errorMsg) {
  const m = String(errorMsg || "").match(/try again in (\d+(?:\.\d+)?)\s*s/i);
  return m ? Math.ceil(parseFloat(m[1])) : 30;
}

export async function POST(req) {
  try {
    const body = await req.json();
    const { message, history = [], availableAccounts = [], apiKey, model } = body;

    if (!apiKey) return NextResponse.json({ error: "No API key. Use the ⚙️ AI Setup button." }, { status: 400 });
    if (!model) return NextResponse.json({ error: "No model selected. Use the ⚙️ AI Setup button." }, { status: 400 });
    if (!message) return NextResponse.json({ error: "No message." }, { status: 400 });

    // Compact system prompt — top 8 accounts only, concise instructions
    const accountList = (availableAccounts || []).slice(0, 8).join(", ");
    const systemPrompt = `You are a restaurant accounting expert (India/Zoho Books). Identify products and suggest correct account heads.
Accounts: ${accountList || "Groceries, Poultry, Dairy, Beverages, Liquor Purchases, Cleaning, Packing materials"}
Key rules: broth/seasoning/knorr=Groceries; oat milk=Dairy; monin/syrup=Beverages; marlboro/classic connect=Cigarettes; soap/detergent=Cleaning.
Reply in 2-3 sentences max. End with "→ Account: [name]".`;

    const messages = [
      { role: "system", content: systemPrompt },
      // Keep only last 4 messages (2 exchanges) to limit token usage
      ...history.slice(-4),
      { role: "user", content: message }
    ];

    const res = await fetch("https://api.groq.com/openai/v1/chat/completions", {
      method: "POST",
      headers: { "Content-Type": "application/json", Authorization: `Bearer ${apiKey}` },
      body: JSON.stringify({
        model,
        messages,
        temperature: 0.3,
        max_tokens: 200  // Strict cap on response size
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
    const reply = data.choices?.[0]?.message?.content?.trim() || "Sorry, I couldn't process that.";
    return NextResponse.json({ reply, model });

  } catch (err) {
    return NextResponse.json({ error: err.message || "Internal error" }, { status: 500 });
  }
}
