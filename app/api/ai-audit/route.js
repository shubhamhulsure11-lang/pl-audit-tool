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
    // Support both single-item mode and batch mode
    const { singleItem, items, availableAccounts = [], apiKey: customKey } = body;

    const apiKey = customKey || process.env.GROQ_API_KEY;
    if (!apiKey) {
      return NextResponse.json({ error: "No Groq API key provided." }, { status: 400 });
    }

    const model = await pickBestGroqModel(apiKey);
    if (!model) {
      return NextResponse.json({ error: "No Groq model available. Verify your key at console.groq.com." }, { status: 400 });
    }

    const accounts = availableAccounts.join(", ");

    // ─── SINGLE ITEM MODE ──────────────────────────────────────────────────────
    if (singleItem) {
      const { item, vendor, actualAccount } = singleItem;
      const prompt = `Restaurant item audit. Analyze this ONE item and determine if it is misclassified.

Item: "${item}"
Vendor: "${vendor}"
Current Account Head in Zoho: "${actualAccount}"
Available Account Heads: ${accounts}

Restaurant rules:
- Groceries: rice, dal, atta, spices, masala, ghee, cooking oil, sugar, sauces, vinegar, vanilla essence, food coloring
- Dairy: milk, paneer, curd, butter, cream, cheese, khoya, buttermilk (NOT ghee)
- Poultry: chicken, mutton, lamb, eggs, goat meat
- Sea food: fish (basa, surmai, pomfret, rawas, salmon, tuna), prawns, shrimp, crab, lobster, squid
- Beverages (NON-ALCOHOLIC): red bull, tonic water, soda, juices, monin syrups, malas, mineral water, tea, coffee, kokum
- Liquor Purchases (ALCOHOLIC ONLY): wine, beer, whisky, vodka, rum, gin, tequila, brandy, champagne
- Cigarette purchases: classic connect, classic bt, ice burst, marlboro, gold flake, wills
- Other Purchases: charcoal, coal, ice cubes, ice slabs, dry ice, skewers, toothpicks
- Packing materials / Packaging & Disposables (same): takeaway boxes, foil, paper bags, tissue, straws, disposable cutlery
- Cleaning and housekeeping: dishwash, detergent, lizol, phenyl, harpic, sanitizer, mops, garbage bags, soap oil
- Kitchen tools: utensils, crockery, glassware (arcoroc, pilsner), kadai, tawa, hotelware, bar tools, dip bowls
- Stationery & Office: registers, KOT books, pens, POS rolls, thermal paper, toners

If correctly classified, return: {"isMisclassified": false}
If misclassified, return:
{"isMisclassified": true, "suggestedAccount": "exact name from Available Account Heads", "webSummary": "what this product actually is (max 8 words)", "reason": "brief explanation"}

Return ONLY valid JSON. No markdown, no extra text.`;

      const res = await fetch("https://api.groq.com/openai/v1/chat/completions", {
        method: "POST",
        headers: { "Content-Type": "application/json", Authorization: `Bearer ${apiKey}` },
        body: JSON.stringify({
          model,
          messages: [{ role: "user", content: prompt }],
          temperature: 0.1,
          max_tokens: 256
        })
      });

      if (!res.ok) {
        const errText = await res.text();
        let errMsg = `Groq error (${res.status})`;
        try { errMsg = JSON.parse(errText)?.error?.message || errMsg; } catch { /* ignore */ }
        return NextResponse.json({ error: errMsg }, { status: res.status });
      }

      const data = await res.json();
      let text = (data.choices?.[0]?.message?.content || "").trim();
      const fence = text.match(/```(?:json)?\s*([\s\S]*?)\s*```/);
      if (fence) text = fence[1].trim();

      try {
        const result = JSON.parse(text);
        return NextResponse.json({ success: true, result, model });
      } catch {
        return NextResponse.json({ success: false, rawText: text, error: "Parse error" });
      }
    }

    // ─── BATCH MODE (legacy, max 10 items) ──────────────────────────────────────
    if (!items || !items.length) {
      return NextResponse.json({ error: "No items provided." }, { status: 400 });
    }

    const itemsToAudit = items.slice(0, 10).map(({ item, vendor, actualAccount }) => ({ item, vendor, actualAccount }));

    const RULES = `Rules:
1.Groceries: rice, dal, atta, spices, masala, cooking oil, sugar, sauces, vinegar, ghee, vanilla essence, food color
2.Dairy: milk, paneer, curd, butter, cream, cheese, khoya, buttermilk
3.Poultry: chicken, mutton, lamb, eggs
4.Sea food: basa, surmai, pomfret, rawas, salmon, prawns, shrimp, crab, lobster, squid
5.Beverages(NON-ALCOHOLIC): red bull, tonic, soda, juices, monin, malas, mineral water, tea, coffee
6.Liquor Purchases(ALCOHOLIC only): wine, beer, whisky, vodka, rum, gin, tequila, brandy, champagne
7.Cigarette purchases: classic connect, classic bt, marlboro, gold flake, wills, ice burst
8.Other Purchases: charcoal, coal, ice, dry ice, skewers, toothpicks
9.Packing materials / Packaging & Disposables(same): takeaway boxes, foil, bags, tissue, straws, disposable cutlery
10.Cleaning: dishwash, detergent, lizol, phenyl, harpic, sanitizer, mops, garbage bags, soap oil
11.Kitchen tools: crockery, glassware, arcoroc, pilsner, kadai, tawa, hotelware, dip bowl, bar tools
12.Stationery: registers, KOT books, pens, POS rolls, toners`;

    const prompt = `Restaurant accounting auditor.
${RULES}
Available accounts: ${accounts}
Items: ${JSON.stringify(itemsToAudit)}
Return JSON array of ONLY misclassified items ([] if all correct):
[{"item":"","vendor":"","actualAccount":"","suggestedAccount":"","webSummary":"","reason":""}]
Return ONLY raw JSON. No markdown.`;

    const res = await fetch("https://api.groq.com/openai/v1/chat/completions", {
      method: "POST",
      headers: { "Content-Type": "application/json", Authorization: `Bearer ${apiKey}` },
      body: JSON.stringify({
        model,
        messages: [{ role: "user", content: prompt }],
        temperature: 0.1,
        max_tokens: 1500
      })
    });

    if (!res.ok) {
      const errText = await res.text();
      let errMsg = `Groq error (${res.status})`;
      try { errMsg = JSON.parse(errText)?.error?.message || errMsg; } catch { /* ignore */ }
      return NextResponse.json({ error: errMsg }, { status: res.status });
    }

    const groqData = await res.json();
    const text = (groqData.choices?.[0]?.message?.content || "").trim();

    let jsonStr = text;
    const fenceMatch = text.match(/```(?:json)?\s*([\s\S]*?)\s*```/);
    if (fenceMatch) jsonStr = fenceMatch[1].trim();

    let parsed = [];
    try {
      parsed = JSON.parse(jsonStr);
      if (!Array.isArray(parsed)) parsed = parsed.misclassifications || parsed.items || [];
    } catch {
      return NextResponse.json({ rawText: text, error: "Failed to parse AI response", parsed: [] });
    }

    return NextResponse.json({ success: true, results: parsed, model });
  } catch (err) {
    return NextResponse.json({ error: err.message || "Internal error" }, { status: 500 });
  }
}
