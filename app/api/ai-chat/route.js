import { NextResponse } from "next/server";

const EXCLUDE_PATTERNS = ["whisper", "tts", "distil", "embed", "vision", "playai", "playht", "guard"];
const MODEL_PREFERENCES = [
  "llama-3.1-8b-instant", "llama3-8b", "llama-3.1-8b",
  "llama-3.3-70b", "llama-3.1-70b", "llama3-70b",
  "llama", "mixtral", "gemma2", "gemma",
];

async function pickBestGroqModel(apiKey) {
  try {
    const res = await fetch("https://api.groq.com/openai/v1/models", {
      headers: { Authorization: `Bearer ${apiKey}` }
    });
    if (!res.ok) return null;
    const data = await res.json();
    const chatIds = (data.data || [])
      .map((m) => m.id)
      .filter((id) => !EXCLUDE_PATTERNS.some((pat) => id.toLowerCase().includes(pat)));
    for (const pref of MODEL_PREFERENCES) {
      const found = chatIds.find((id) => id.toLowerCase().includes(pref));
      if (found) return found;
    }
    return chatIds[0] || null;
  } catch { return null; }
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
    if (!model) {
      return NextResponse.json({ error: "No Groq model available for your API key. Verify at console.groq.com." }, { status: 400 });
    }

    const systemPrompt = `You are a restaurant accounting expert assistant for an Indian restaurant using Zoho Books.
Your job is to help identify what a product is and suggest the correct account head for it.

AVAILABLE ACCOUNT HEADS IN THIS SHEET: ${availableAccounts.length ? availableAccounts.join(", ") : "Not uploaded yet"}

RESTAURANT ACCOUNTING RULES (India):
- Groceries/Food: rice, dal, atta, spices, masala, ghee, oil, sauces, vinegar, dry fruits
- Dairy: milk, paneer, curd, butter, cream, cheese, khoya, buttermilk
- Poultry: chicken, mutton, lamb, eggs, goat meat
- Sea food: fish (basa, surmai, pomfret, rawas, salmon), prawns, shrimp, crab, lobster, squid
- Beverages (NON-ALCOHOLIC only): red bull, tonic, soda, juices, monin syrups, malas, mineral water, tea, coffee
- Liquor Purchases (ALCOHOLIC only): wine, beer, whisky, vodka, rum, gin, tequila, brandy, champagne
- Cigarette purchases: classic connect, marlboro, gold flake, wills, ice burst
- Other Purchases: charcoal, coal, ice cubes, ice slabs, dry ice, skewers, toothpicks
- Packing materials / Packaging & Disposables (SAME CATEGORY): takeaway boxes, foil, paper bags, tissue, straws, disposable cutlery
- Cleaning and housekeeping: dishwash, detergent, floor cleaner, lizol, sanitizer, mops, garbage bags
- Kitchen tools: utensils, crockery, glassware, kadai, tawa, hotelware, arcoroc, bar tools
- Stationery & Office: registers, KOT books, pens, POS rolls, thermal paper

When asked about an item, identify what it is, then state the correct account head from the AVAILABLE ACCOUNT HEADS list.
Be concise. Use emojis sparingly. Always end with the suggested account head name in bold.`;

    const messages = [
      { role: "system", content: systemPrompt },
      ...history.slice(-8), // Keep last 8 messages for context
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
