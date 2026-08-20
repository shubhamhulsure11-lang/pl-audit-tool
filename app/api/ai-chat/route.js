import { NextResponse } from "next/server";

// Standard production Groq models (in preference order)
const ALLOWED_STANDARD_MODELS = [
  "llama-3.3-70b-versatile",
  "llama-3.1-8b-instant",
  "llama3-70b-8192",
  "llama3-8b-8192",
  "mixtral-8x7b-32768",
  "gemma2-9b-it"
];

async function pickBestGroqModel(apiKey) {
  try {
    const res = await fetch("https://api.groq.com/openai/v1/models", {
      headers: { Authorization: `Bearer ${apiKey}` }
    });
    if (!res.ok) return "llama-3.3-70b-versatile";
    const data = await res.json();
    const availableIds = new Set((data.data || []).map((m) => m.id.toLowerCase()));

    for (const model of ALLOWED_STANDARD_MODELS) {
      if (availableIds.has(model.toLowerCase())) return model;
    }

    const safeFallback = (data.data || [])
      .map(m => m.id)
      .find(id => {
        const lower = id.toLowerCase();
        return (lower.startsWith("llama") || lower.startsWith("gemma") || lower.startsWith("mixtral")) && !lower.includes("guard");
      });
    return safeFallback || "llama-3.3-70b-versatile";
  } catch {
    return "llama-3.3-70b-versatile";
  }
}

export async function POST(req) {
  try {
    const body = await req.json();
    const { message, history = [], availableAccounts = [], apiKey: customKey } = body;

    const apiKey = customKey || process.env.GROQ_API_KEY;
    if (!apiKey) {
      return NextResponse.json({ error: "No Groq API key provided." }, { status: 400 });
    }

    const model = await pickBestGroqModel(apiKey);

    const systemPrompt = `You are a restaurant accounting expert assistant for an Indian restaurant using Zoho Books.
Your job is to help identify what a product is and suggest the correct account head for it.

AVAILABLE ACCOUNT HEADS IN THIS SHEET: ${availableAccounts.length ? availableAccounts.join(", ") : "Not uploaded yet"}

RESTAURANT ACCOUNTING RULES (India):
- Groceries / Food Raw Materials: rice, dal, atta, spices, masala, ghee, oil, sauces, vinegar, dry fruits, knorr broth / powder, seasonings
- Dairy: milk, paneer, curd, butter, cream, cheese, khoya, buttermilk
- Poultry: raw fresh chicken, mutton, lamb, eggs, goat meat
- Sea food: fish (basa, surmai, pomfret, rawas, salmon), prawns, shrimp, crab, lobster, squid
- Beverages (NON-ALCOHOLIC only): red bull, tonic, soda, juices, monin syrups, malas, mineral water, tea, coffee
- Liquor Purchases (ALCOHOLIC only): wine, beer, whisky, vodka, rum, gin, tequila, brandy, champagne
- Cigarette purchases: classic connect, marlboro, gold flake, wills, ice burst
- Other Purchases: charcoal, coal, ice cubes, ice slabs, dry ice, skewers, toothpicks
- Packing materials / Packaging & Disposables (SAME CATEGORY): takeaway boxes, foil, paper bags, tissue, straws, disposable cutlery
- Cleaning and housekeeping: dishwash, detergent, floor cleaner, lizol, sanitizer, mops, garbage bags, soap oil
- Kitchen tools: utensils, crockery, glassware, kadai, tawa, hotelware, arcoroc, bar tools
- Stationery & Office: registers, KOT books, pens, POS rolls, thermal paper

When asked about an item, identify what it is, then state the correct account head from the AVAILABLE ACCOUNT HEADS list.
Be concise. Always conclude with the suggested account head in bold.`;

    const messages = [
      { role: "system", content: systemPrompt },
      ...history.slice(-8),
      { role: "user", content: message }
    ];

    const res = await fetch("https://api.groq.com/openai/v1/chat/completions", {
      method: "POST",
      headers: { "Content-Type": "application/json", Authorization: `Bearer ${apiKey}` },
      body: JSON.stringify({ model, messages, temperature: 0.3, max_tokens: 512 })
    });

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
