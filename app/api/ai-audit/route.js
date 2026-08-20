import { NextResponse } from "next/server";

async function getAvailableModel(apiKey) {
  try {
    const res = await fetch(
      `https://generativelanguage.googleapis.com/v1beta/models?key=${apiKey}&pageSize=100`
    );
    if (!res.ok) return null;
    const data = await res.json();
    const models = (data.models || []).filter(
      (m) => m.supportedGenerationMethods?.includes("generateContent")
    );

    // Pick the best available model by preference
    const preferenceOrder = [
      "gemini-2.5-flash",
      "gemini-2.5-pro",
      "gemini-2.0-flash",
      "gemini-1.5-flash",
      "gemini-1.5-pro",
      "gemini-pro",
    ];

    for (const pref of preferenceOrder) {
      const found = models.find((m) => m.name?.includes(pref));
      if (found) return found.name.replace("models/", "");
    }

    // Fallback: just pick the first available model
    if (models[0]) return models[0].name.replace("models/", "");
    return null;
  } catch {
    return null;
  }
}

export async function POST(req) {
  try {
    const body = await req.json();
    const { items, availableAccounts, apiKey: customKey } = body;

    const apiKey = customKey || process.env.GEMINI_API_KEY;
    if (!apiKey) {
      return NextResponse.json(
        { error: "No Gemini API key provided. Please enter your Gemini API key." },
        { status: 400 }
      );
    }

    if (!items || !items.length) {
      return NextResponse.json({ error: "No items provided for analysis." }, { status: 400 });
    }

    // Step 1: Discover which model this API key actually supports
    const model = await getAvailableModel(apiKey);
    if (!model) {
      return NextResponse.json(
        { error: "Could not find any available Gemini model for your API key. Please check the key is valid and has the Generative Language API enabled." },
        { status: 400 }
      );
    }

    const itemsToAudit = items.slice(0, 30);

    const prompt = `You are an expert restaurant accounting auditor.
Analyze the following list of purchased items from a restaurant's Zoho Books export.
Identify what each product is in the real world using your knowledge.

RESTAURANT ACCOUNTING RULES:
1. Groceries: Food raw materials, rice, dal, wheat, atta, maida, sooji, spices, masala, oil, sugar, salt, sauces, vinegar, dry fruits. GHEE = Groceries.
2. Dairy: Milk, paneer, curd, dahi, fresh cream, butter, cheese, khoya, buttermilk.
3. Poultry: Chicken, mutton, lamb, goat meat, eggs.
4. Sea food purchases: Fish (Basa, Surmai, Pomfret, Rawas, Salmon), Prawns, Shrimp, Crab, Lobster, Squid.
5. Beverages: Non-alcoholic ONLY (Red Bull, Tonic Water, Sodas, Juices, Syrups like Monin/Malas, Mineral Water).
6. Liquor Purchases: Alcoholic beverages ONLY (Wine, Beer, Whisky, Vodka, Rum, Gin, Tequila, Brandy, Champagne, etc.).
7. Cigarette purchases: Cigarettes and tobacco (Classic Connect, Ice Burst, Marlboro, Gold Flake, Wills, etc.).
8. Other Purchases: Charcoal, Ice (cubes/slabs/dry ice), wooden skewers, toothpicks.
9. Packing materials / Packaging & Disposables: Takeaway boxes, paper bags, tissues, foil, cling wrap, straws, disposable cutlery.
10. Cleaning and housekeeping: Dishwash, detergents, floor cleaners, sanitizers, brooms, mops, garbage bags.
11. Kitchen tools: Utensils, kadai, tawa, knives, chopping boards, crockery, cocktail shakers.
12. Stationery & Office: Registers, KOT books, bill books, pens, POS rolls, thermal paper.

AVAILABLE ACCOUNT HEADS IN THIS SHEET:
${(availableAccounts || []).join(", ")}

ITEMS TO EVALUATE:
${JSON.stringify(itemsToAudit, null, 2)}

For each item, determine if it is MISCLASSIFIED based on the rules and available account heads.

Return ONLY a JSON array (no markdown fences) of misclassified items in this format:
[{"item":"...","vendor":"...","actualAccount":"...","suggestedAccount":"...","isMisclassified":true,"webSummary":"short product description","reason":"why it belongs in the suggested account"}]
If nothing is misclassified, return an empty array: []`;

    // Step 2: Try with googleSearch grounding first, then without
    const payloads = [
      {
        contents: [{ parts: [{ text: prompt }] }],
        tools: [{ googleSearch: {} }],
        generationConfig: { temperature: 0.1 }
      },
      {
        contents: [{ parts: [{ text: prompt }] }],
        generationConfig: { temperature: 0.1 }
      }
    ];

    let data = null;
    let lastError = "";

    for (const payload of payloads) {
      const response = await fetch(
        `https://generativelanguage.googleapis.com/v1beta/models/${model}:generateContent?key=${apiKey}`,
        {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify(payload)
        }
      );

      if (response.ok) {
        data = await response.json();
        break;
      }

      const errText = await response.text();
      try {
        lastError = JSON.parse(errText)?.error?.message || `HTTP ${response.status}`;
      } catch {
        lastError = `HTTP ${response.status}`;
      }
    }

    if (!data) {
      return NextResponse.json({ error: `Gemini API error: ${lastError}` }, { status: 400 });
    }

    const text = (data.candidates?.[0]?.content?.parts || [])
      .map((p) => p.text || "")
      .join("")
      .trim();

    let jsonStr = text;
    const fenceMatch = text.match(/```(?:json)?\s*([\s\S]*?)\s*```/);
    if (fenceMatch) jsonStr = fenceMatch[1].trim();

    // Sometimes the model wraps in an outer object
    if (jsonStr.startsWith("{")) {
      try {
        const obj = JSON.parse(jsonStr);
        jsonStr = JSON.stringify(obj.misclassifications || obj.items || obj.results || []);
      } catch { /* ignore */ }
    }

    let parsed = [];
    try {
      parsed = JSON.parse(jsonStr);
      if (!Array.isArray(parsed)) parsed = [];
    } catch {
      return NextResponse.json({ rawText: text, error: "Failed to parse AI response", parsed: [] });
    }

    return NextResponse.json({ success: true, results: parsed, modelUsed: model });
  } catch (err) {
    return NextResponse.json({ error: err.message || "Internal server error" }, { status: 500 });
  }
}
