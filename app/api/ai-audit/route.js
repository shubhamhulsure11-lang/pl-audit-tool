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

  // Fallback: first available safe model (prefer llama, then anything else)
  return available.find(id => id.toLowerCase().startsWith("llama")) || available[0];
}

export async function POST(req) {
  try {
    const body = await req.json();
    const { singleItem, items, availableAccounts = [], apiKey: customKey } = body;

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

    const accounts = availableAccounts.join(", ");

    // ─── SINGLE ITEM MODE ──────────────────────────────────────────────────────
    if (singleItem) {
      const { item, vendor, actualAccount } = singleItem;
      const prompt = `Restaurant accounting auditor. Analyze this ONE purchase item.

Item: "${item}"
Vendor: "${vendor}"
Current Account Head in Zoho: "${actualAccount}"
Available Account Heads: ${accounts}

Rules:
- Groceries / Food Raw Materials: rice, dal, atta, spices, masala, ghee, cooking oil, sugar, sauces, vinegar, vanilla essence, knorr, chicken broth powder, bouillon, seasonings
- Dairy: milk, paneer, curd, butter, cream, cheese, khoya, buttermilk (NOT ghee)
- Poultry: raw fresh chicken, mutton, lamb, eggs, goat meat (processed broth/powders are Groceries)
- Sea food: fish (basa, surmai, pomfret, rawas, salmon, tuna), prawns, shrimp, crab, lobster, squid
- Beverages (NON-ALCOHOLIC): red bull, tonic water, soda, juices, monin syrups, malas, mineral water, tea, coffee
- Liquor Purchases (ALCOHOLIC ONLY): wine, beer, whisky, vodka, rum, gin, tequila, brandy, champagne
- Cigarette purchases: classic connect, classic bt, ice burst, marlboro, gold flake, wills
- Other Purchases: charcoal, coal, ice cubes, ice slabs, dry ice, skewers, toothpicks
- Packing materials / Packaging & Disposables (same): takeaway boxes, foil, paper bags, tissue, straws, disposable cutlery
- Cleaning: dishwash, detergent, lizol, phenyl, harpic, sanitizer, mops, garbage bags, soap oil
- Kitchen tools: utensils, crockery, glassware (arcoroc, pilsner), kadai, tawa, hotelware, bar tools, dip bowls
- Stationery: registers, KOT books, pens, POS rolls, thermal paper, toners

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

    // ─── BATCH MODE (fallback) ──────────────────────────────────────────────────
    if (!items || !items.length) {
      return NextResponse.json({ error: "No items provided." }, { status: 400 });
    }

    const itemsToAudit = items.slice(0, 10).map(({ item, vendor, actualAccount }) => ({ item, vendor, actualAccount }));

    const RULES = `Rules:
1.Groceries: rice, dal, atta, spices, masala, cooking oil, sugar, sauces, vinegar, ghee, vanilla, knorr broth, seasoning powder
2.Dairy: milk, paneer, curd, butter, cream, cheese, khoya, buttermilk
3.Poultry: fresh raw chicken, mutton, lamb, eggs
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
