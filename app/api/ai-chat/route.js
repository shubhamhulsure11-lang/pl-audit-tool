import { NextResponse } from "next/server";

function parseRetryAfter(errorMsg) {
  const m = String(errorMsg || "").match(/try again in (\d+(?:\.\d+)?)\s*s/i);
  return m ? Math.ceil(parseFloat(m[1])) : 30;
}

function cleanThinking(text) {
  if (!text) return "";
  // Remove closed <think>...</think> blocks
  let cleaned = text.replace(/<think>[\s\S]*?<\/think>/gi, "");
  // Remove unclosed <think>... blocks if model hit max tokens inside think
  cleaned = cleaned.replace(/<think>[\s\S]*/gi, "");
  return cleaned.trim();
}

export async function POST(req) {
  try {
    const body = await req.json();
    const { message, history = [], availableAccounts = [], apiKey, model = "llama-3.1-8b-instant" } = body;

    if (!apiKey) return NextResponse.json({ error: "No API key. Use ⚙️ AI Setup to configure." }, { status: 400 });
    if (!message) return NextResponse.json({ error: "No message." }, { status: 400 });

    const accountList = (availableAccounts || []).slice(0, 10).join(", ");
    const systemPrompt = `You are a restaurant accounting auditor for an Indian restaurant using Zoho Books.
Goal: Identify what product the user asks about and state the correct Account Head.

Available Account Heads from the uploaded sheet:
${accountList || "Groceries purchases, Dairy products purchases, Sea food purchases, Poultry and meat purchases, Beverages, Liquor purchases, Cleaning and housekeeping, Packaging & Disposables, Stationery"}

Key Accounting Rules:
- Vanilla essence, food color, knorr, seasonings, broth powders, coconut milk, cooking oil, sauces, rice, flour, spices → Groceries purchases
- Fresh milk, paneer, curd, butter, cream, cheese → Dairy products purchases
- Red bull, tonic, ginger ale, soda, juices, monin syrups, mineral water → Beverages (Non-alcoholic)
- Wine, beer, whisky, vodka, rum, gin, tequila → Liquor purchases (Alcoholic)
- Takeaway boxes, foil, paper bags, napkins, banana leaves for serving → Packaging & Disposables / Packing material
- Dishwash, floor cleaner, lizol, detergent, soap oil → Cleaning and housekeeping
- Raw chicken, mutton, lamb, eggs → Poultry and meat purchases
- Raw fish, prawns, shrimp, crab, salmon → Sea food purchases

Format your response in 2 concise sentences:
1. Explain what the product is.
2. End with: "→ Recommended Account: **[Exact Account Name]**"`;

    const messages = [
      { role: "system", content: systemPrompt },
      ...history.slice(-4).map(h => ({ role: h.role, content: cleanThinking(h.content) })),
      { role: "user", content: message }
    ];

    const res = await fetch("https://api.groq.com/openai/v1/chat/completions", {
      method: "POST",
      headers: { "Content-Type": "application/json", Authorization: `Bearer ${apiKey}` },
      body: JSON.stringify({
        model,
        messages,
        temperature: 0.1,
        max_tokens: 350
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
    let reply = (data.choices?.[0]?.message?.content || "").trim();
    reply = cleanThinking(reply);

    if (!reply) {
      reply = "Product identified. Please check available account heads in your sheet.";
    }

    return NextResponse.json({ reply, model });

  } catch (err) {
    return NextResponse.json({ error: err.message || "Internal error" }, { status: 500 });
  }
}
