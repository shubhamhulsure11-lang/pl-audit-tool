import { NextResponse } from "next/server";

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

    const itemsToAudit = items.slice(0, 30);

    const systemPrompt = `You are an expert restaurant accounting auditor in India. Your job is to identify misclassified purchase items in a restaurant's Zoho Books accounts.

RESTAURANT ACCOUNTING RULES:
1. Groceries: Food raw materials — rice, dal, wheat, atta, maida, sooji, spices, masala, cooking oil, sugar, salt, sauces, vinegar, dry fruits, pickles, flour. GHEE is always Groceries.
2. Dairy: Milk, paneer, curd, dahi, fresh cream, butter, cheese, khoya, mawa, buttermilk.
3. Poultry: Chicken, mutton, lamb, goat meat, eggs, quail.
4. Sea food purchases: Fish (Basa, Surmai, Pomfret, Rawas, Salmon, Tuna), Prawns (any size like 16/20, 21/25), Shrimp, Crab, Lobster, Squid, Octopus.
5. Beverages: Non-alcoholic ONLY — Red Bull, energy drinks, Tonic Water, Diet Coke, Sodas, Real Juices, Fruit Syrups like Monin/Malas, Packaged Mineral Water, Limca, Sprite, Thums Up, Tea, Coffee, Kokum, Sharbat.
6. Liquor Purchases: Alcoholic beverages ONLY — Wine, Beer, Whisky, Vodka, Rum, Gin, Tequila, Brandy, Champagne, Shiraz, Cabernet, Sauvignon, Scotch, Bourbon, Sake, Mead. NOT beverages.
7. Cigarette purchases: Cigarettes and tobacco — Classic Connect, Classic Regular, Classic Milds, Ice Burst, Marlboro, Gold Flake, Wills Navy Cut, Beedi, Hookah tobacco. NOT groceries or beverages.
8. Other Purchases: Charcoal, Coal, Ice cubes, Ice slabs, Dry ice, wooden skewers, toothpicks, match boxes, lighters.
9. Packing materials / Packaging & Disposables: Both names mean EXACTLY the same thing. Meal trays, takeaway boxes, paper bags, tissues, napkins, aluminium foil, cling wrap, straws, disposable plates, disposable cutlery, zip-lock bags, kraft boxes.
10. Cleaning and housekeeping: Dishwash liquid, detergents, floor cleaners (Lizol, Phenyl, Domex), Colin, Harpic, bleach, sanitizers, brooms, mops, wipers, scrubbers, garbage bags, gloves.
11. Kitchen tools: Katoris, utensils, kadai, tawa, knives, chopping boards, strainers, ladles, crockery, hotelware, cocktail shakers, jiggers, shot glasses, bar tools.
12. Stationery & Office: Registers, KOT books, bill books, pens, POS rolls, thermal paper, toners.

When you see an item name that looks like a product code, abbreviation, or SKU — reason about what it likely is based on the vendor name and context.

You MUST return ONLY valid JSON — no explanation, no markdown fences, no extra text. Just the raw JSON array.`;

    const userMessage = `Evaluate these items for misclassification.
AVAILABLE ACCOUNT HEADS: ${(availableAccounts || []).join(", ")}
ITEMS: ${JSON.stringify(itemsToAudit, null, 2)}

Return a JSON array of ONLY misclassified items. If nothing is wrong, return [].
Format: [{"item":"...","vendor":"...","actualAccount":"...","suggestedAccount":"one of the AVAILABLE ACCOUNT HEADS","isMisclassified":true,"webSummary":"what product this actually is","reason":"why this account head is wrong"}]`;

    const groqRes = await fetch("https://api.groq.com/openai/v1/chat/completions", {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        Authorization: `Bearer ${apiKey}`
      },
    // Try all free Groq models in order of quality until one works
    const groqModels = [
      "llama-3.3-70b-versatile",
      "llama3-70b-8192",
      "llama-3.1-70b-versatile",
      "llama-3.1-8b-instant",
      "llama3-8b-8192",
      "mixtral-8x7b-32768",
      "gemma2-9b-it",
      "gemma-7b-it"
    ];

    let groqData = null;
    let lastError = "";

    for (const model of groqModels) {
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

      if (groqRes.ok) {
        groqData = await groqRes.json();
        break;
      }

      const errText = await groqRes.text();
      try {
        const parsed = JSON.parse(errText);
        lastError = parsed?.error?.message || `HTTP ${groqRes.status}`;
      } catch { lastError = `HTTP ${groqRes.status}`; }

      // Only continue loop for model-not-found errors, not auth errors
      if (groqRes.status === 401 || groqRes.status === 403) break;
    }

    if (!groqData) {
      return NextResponse.json({ error: lastError || "No Groq model available for your account." }, { status: 400 });
    }

    const text = groqData.choices?.[0]?.message?.content?.trim() || "";

    let jsonStr = text;
    // Strip any accidental markdown fences
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

    return NextResponse.json({ success: true, results: parsed });
  } catch (err) {
    return NextResponse.json({ error: err.message || "Internal server error" }, { status: 500 });
  }
}
