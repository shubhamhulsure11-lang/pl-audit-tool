import { NextResponse } from "next/server";

// Preferred model substrings in order of quality
const MODEL_PREFERENCES = [
  "llama-3.3",
  "llama-3.1-70b",
  "llama3-70b",
  "llama-3.1",
  "llama3",
  "mixtral",
  "gemma2",
  "gemma",
];

async function pickBestGroqModel(apiKey) {
  try {
    const res = await fetch("https://api.groq.com/openai/v1/models", {
      headers: { Authorization: `Bearer ${apiKey}` }
    });
    if (!res.ok) return null;
    const data = await res.json();
    const ids = (data.data || []).map((m) => m.id);

    for (const pref of MODEL_PREFERENCES) {
      const found = ids.find((id) => id.toLowerCase().includes(pref));
      if (found) return found;
    }
    // Fallback: first available model
    return ids[0] || null;
  } catch {
    return null;
  }
}

export async function POST(req) {
  try {
    const body = await req.json();
    const { items, availableAccounts, apiKey: customKey } = body;

    const apiKey = customKey || process.env.GROQ_API_KEY;
    if (!apiKey) {
      return NextResponse.json(
        { error: "No Groq API key provided. Please enter your Groq API key." },
        { status: 400 }
      );
    }

    if (!items || !items.length) {
      return NextResponse.json({ error: "No items provided for analysis." }, { status: 400 });
    }

    // Auto-discover the best available model for this API key
    const model = await pickBestGroqModel(apiKey);
    if (!model) {
      return NextResponse.json(
        { error: "Could not find any available Groq model for your API key. Please verify the key at console.groq.com." },
        { status: 400 }
      );
    }

    const itemsToAudit = items.slice(0, 30);

    const systemPrompt = `You are an expert restaurant accounting auditor in India. Your job is to identify misclassified purchase items in a restaurant's Zoho Books accounts.

RESTAURANT ACCOUNTING RULES:
1. Groceries: Food raw materials — rice, dal, wheat, atta, maida, sooji, spices, masala, cooking oil, sugar, salt, sauces, vinegar, dry fruits, pickles, flour. GHEE is always Groceries.
2. Dairy: Milk, paneer, curd, dahi, fresh cream, butter, cheese, khoya, mawa, buttermilk.
3. Poultry: Chicken, mutton, lamb, goat meat, eggs, quail.
4. Sea food purchases: Fish (Basa, Surmai, Pomfret, Rawas, Salmon, Tuna), Prawns (any size like 16/20, 21/25), Shrimp, Crab, Lobster, Squid, Octopus.
5. Beverages: Non-alcoholic ONLY — Red Bull, energy drinks, Tonic Water, Diet Coke, Sodas, Real Juices, Fruit Syrups like Monin/Malas, Packaged Mineral Water, Limca, Sprite, Thums Up, Tea, Coffee, Kokum, Sharbat.
6. Liquor Purchases: Alcoholic beverages ONLY — Wine, Beer, Whisky, Vodka, Rum, Gin, Tequila, Brandy, Champagne, Shiraz, Cabernet, Sauvignon, Scotch, Bourbon, Sake, Mead. NOT beverages.
7. Cigarette purchases: Cigarettes and tobacco — Classic Connect, Classic Regular, Classic Milds, Ice Burst, Marlboro, Gold Flake, Wills Navy Cut, Beedi, Hookah tobacco.
8. Other Purchases: Charcoal, Coal, Ice cubes, Ice slabs, Dry ice, wooden skewers, toothpicks.
9. Packing materials / Packaging & Disposables: Both names are the SAME category. Meal trays, takeaway boxes, paper bags, tissues, napkins, aluminium foil, cling wrap, straws, disposable plates/cutlery.
10. Cleaning and housekeeping: Dishwash liquid, detergents, floor cleaners (Lizol, Phenyl, Domex), Colin, Harpic, bleach, sanitizers, brooms, mops, garbage bags.
11. Kitchen tools: Utensils, kadai, tawa, knives, chopping boards, crockery, hotelware, cocktail shakers.
12. Stationery & Office: Registers, KOT books, bill books, pens, POS rolls, thermal paper, toners.

You MUST return ONLY valid JSON — no explanation, no markdown, no extra text. Just the raw JSON array.`;

    const userMessage = `Evaluate these items for misclassification.
AVAILABLE ACCOUNT HEADS: ${(availableAccounts || []).join(", ")}
ITEMS: ${JSON.stringify(itemsToAudit, null, 2)}

Return a JSON array of ONLY misclassified items. If nothing is wrong, return [].
Format: [{"item":"...","vendor":"...","actualAccount":"...","suggestedAccount":"one of the AVAILABLE ACCOUNT HEADS","isMisclassified":true,"webSummary":"what this product actually is","reason":"why the current account head is wrong"}]`;

    const groqRes = await fetch("https://api.groq.com/openai/v1/chat/completions", {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        Authorization: `Bearer ${apiKey}`
      },
      body: JSON.stringify({
        model,
        messages: [
          { role: "system", content: systemPrompt },
          { role: "user", content: userMessage }
        ],
        temperature: 0.1,
        max_tokens: 4096
      })
    });

    if (!groqRes.ok) {
      const errText = await groqRes.text();
      let errMsg = `Groq API error (${groqRes.status})`;
      try { errMsg = JSON.parse(errText)?.error?.message || errMsg; } catch { /* ignore */ }
      return NextResponse.json({ error: errMsg }, { status: groqRes.status });
    }

    const groqData = await groqRes.json();
    const text = (groqData.choices?.[0]?.message?.content || "").trim();

    let jsonStr = text;
    const fenceMatch = text.match(/```(?:json)?\s*([\s\S]*?)\s*```/);
    if (fenceMatch) jsonStr = fenceMatch[1].trim();

    let parsed = [];
    try {
      parsed = JSON.parse(jsonStr);
      if (!Array.isArray(parsed)) {
        parsed = parsed.misclassifications || parsed.items || parsed.results || [];
      }
    } catch {
      return NextResponse.json({ rawText: text, error: "Failed to parse AI response as JSON", parsed: [] });
    }

    return NextResponse.json({ success: true, results: parsed, modelUsed: model });
  } catch (err) {
    return NextResponse.json({ error: err.message || "Internal server error" }, { status: 500 });
  }
}
