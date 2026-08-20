import { NextResponse } from "next/server";

// Preferred model order — all verified standard Groq models
const MODEL_PREFERENCES = [
  "llama-3.1-8b-instant",
  "llama-3.3-70b-versatile",
  "llama-3.1-70b-versatile",
  "llama3-8b-8192",
  "llama3-70b-8192",
  "mixtral-8x7b-32768",
  "gemma2-9b-it",
  "gemma-7b-it",
];

// Keywords that identify non-chat/unusable models
const EXCLUDE_PATTERNS = [
  "whisper", "tts", "distil", "embed", "vision",
  "playai", "playht", "guard", "canopylabs", "orpheus",
  "arabic", "preview", "speculative",
];

async function pickBestGroqModel(apiKey) {
  const listRes = await fetch("https://api.groq.com/openai/v1/models", {
    headers: { Authorization: `Bearer ${apiKey}` }
  });
  if (!listRes.ok) throw new Error(`Groq /models error (${listRes.status})`);

  const listData = await listRes.json();
  const available = (listData.data || [])
    .map(m => m.id)
    .filter(id => {
      const lower = id.toLowerCase();
      return !EXCLUDE_PATTERNS.some(pat => lower.includes(pat));
    });

  if (available.length === 0) throw new Error("No usable Groq chat model found on your account.");

  // Pick from preference list first
  for (const pref of MODEL_PREFERENCES) {
    const match = available.find(id => id.toLowerCase() === pref.toLowerCase());
    if (match) return match;
  }

  // Fallback: first available llama, then any safe model
  return available.find(id => id.toLowerCase().startsWith("llama")) || available[0];
}

export async function POST(req) {
  try {
    const body = await req.json();
    const { message, history = [], availableAccounts = [], apiKey: customKey } = body;

    const apiKey = customKey || process.env.GROQ_API_KEY;
    if (!apiKey) {
      return NextResponse.json({ error: "No Groq API key provided." }, { status: 400 });
    }

    let model;
    try {
      model = await pickBestGroqModel(apiKey);
    } catch (e) {
      return NextResponse.json({ error: `Model selection failed: ${e.message}` }, { status: 400 });
    }

    const systemPrompt = `You are a restaurant accounting expert assistant for an Indian restaurant using Zoho Books.
Your job is to help identify what a product is and suggest the correct account head for it.

AVAILABLE ACCOUNT HEADS IN THIS SHEET: ${availableAccounts.length ? availableAccounts.join(", ") : "Not uploaded yet"}

RESTAURANT ACCOUNTING RULES (India):
- Groceries / Food Raw Materials: rice, dal, atta, spices, masala, ghee, oil, sauces, vinegar, dry fruits, knorr broth / powder, seasonings, bouillon
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
