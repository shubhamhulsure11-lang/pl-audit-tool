import { NextResponse } from "next/server";

export async function POST(req) {
  try {
    const body = await req.json();
    const { items, availableAccounts, apiKey: customKey } = body;

    const apiKey = customKey || process.env.GEMINI_API_KEY;
    if (!apiKey) {
      return NextResponse.json(
        { error: "No Gemini API key provided. Please enter your Gemini API key in the AI settings box." },
        { status: 400 }
      );
    }

    if (!items || !items.length) {
      return NextResponse.json({ error: "No items provided for analysis." }, { status: 400 });
    }

    // Limit batch size to 30 items per request for speed and response quality
    const itemsToAudit = items.slice(0, 30);

    const prompt = `You are an expert restaurant accounting auditor.
Analyze the following list of purchased items from a restaurant's Zoho Books export.
Use Google Search to look up unfamiliar product names, brand codes, packaging SKUs, or abbreviations to identify what the product actually is in the real world.

RESTAURANT ACCOUNTING RULES:
1. "Groceries" (or "Food & Groceries"): Food raw materials, rice, dal, wheat, atta, maida, sooji, spices, masala, cooking oil, sugar, salt, sauces, vinegar, dry fruits, etc.
   - GHEE is strictly Groceries (not Dairy).
2. "Dairy": Milk, paneer, curd, dahi, fresh cream, butter, cheese, khoya, mawa, buttermilk.
3. "Poultry": Chicken, mutton, lamb, goat meat, eggs, and general meats.
4. "Sea food purchases" (or "Sea food"): Fish (Basa, Surmai, Pomfret, Rawas, Salmon), Prawns, Shrimp, Crab, Lobster, Squid, etc.
5. "Beverages": Non-alcoholic drinks ONLY (Red Bull, Tonic Water, Diet Coke, Sodas, Real Juices, Fruit Syrups/Crushes like Monin/Malas, Packaged Mineral Water).
6. "Liquor Purchases": Alcoholic beverages ONLY (Wine, Beer, Whisky, Vodka, Rum, Gin, Tequila, Brandy, Champagne, Shiraz, Cabernet, etc.).
7. "Cigarette purchases": Cigarettes and smoking tobacco (Classic Connect, Ice Burst, Marlboro, Gold Flake, Wills, Lights, etc.).
8. "Other Purchases": Charcoal (wood charcoal, coal), Ice (ice cubes, ice slabs, dry ice), wooden skewers, toothpicks.
9. "Packing materials" / "Packaging & Disposables": Meal trays, containers, takeaway boxes, paper bags, tissues, paper napkins, aluminium foil, cling wrap, straws, disposable cutlery. Note: "Packing materials" and "Packaging & Disposables" are the EXACT SAME category.
10. "Cleaning and housekeeping": Dishwash, detergents, floor cleaners (Lizol, Phenyl), Colin, Harpic, bleach, sanitizers, brooms, mops, wipers, garbage bags.
11. "Kitchen tools": Katoris, utensils, kadai, tawa, knives, chopping boards, strainers, ladles, crockery, hotelware, cocktail shakers, jiggers.
12. "Stationery & Office": Registers, KOT books, bill books, pens, POS rolls, thermal paper, toners.

AVAILABLE ACCOUNT HEADS IN THIS SHEET:
${(availableAccounts || []).join(", ")}

ITEMS TO EVALUATE:
${JSON.stringify(itemsToAudit, null, 2)}

For each item:
1. Search the web to identify what it is.
2. Determine if it is correctly classified or MISCLASSIFIED based on the restaurant rules and available account heads.
3. If misclassified, determine the best suggested account from the AVAILABLE ACCOUNT HEADS.

Return a valid JSON array of objects with the following format:
[
  {
    "item": "Exact item name from input",
    "vendor": "Vendor name",
    "actualAccount": "Actual account name from input",
    "suggestedAccount": "Best matching account from AVAILABLE ACCOUNT HEADS",
    "isMisclassified": true,
    "webSummary": "Short 4-8 word description of what the product was identified as online",
    "reason": "Clear explanation of why it belongs in the suggested account"
  }
]
Only include items in the response that are MISCLASSIFIED (isMisclassified: true).
Return ONLY the raw JSON array. Do not wrap in markdown quotes if possible, or use standard \`\`\`json block.`;

    const response = await fetch(
      `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.0-flash:generateContent?key=${apiKey}`,
      {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          contents: [{ parts: [{ text: prompt }] }],
          tools: [{ googleSearch: {} }],
          generationConfig: {
            temperature: 0.1
          }
        })
      }
    );

    if (!response.ok) {
      const errText = await response.text();
      let errMsg = `Gemini API error (${response.status})`;
      try {
        const parsedErr = JSON.parse(errText);
        if (parsedErr?.error?.message) errMsg = parsedErr.error.message;
      } catch (e) {}
      return NextResponse.json({ error: errMsg }, { status: response.status });
    }

    const data = await response.json();
    const candidate = data.candidates?.[0];
    const text = candidate?.content?.parts?.map(p => p.text || "").join("") || "";

    let jsonStr = text.trim();
    const jsonMatch = jsonStr.match(/```(?:json)?\s*([\s\S]*?)\s*```/);
    if (jsonMatch) {
      jsonStr = jsonMatch[1].trim();
    }

    let parsed = [];
    try {
      parsed = JSON.parse(jsonStr);
      if (!Array.isArray(parsed)) {
        parsed = parsed.misclassifications || parsed.items || [];
      }
    } catch (e) {
      return NextResponse.json({
        rawText: text,
        error: "Failed to parse AI response as JSON",
        parsed: []
      });
    }

    return NextResponse.json({ success: true, results: parsed });
  } catch (err) {
    return NextResponse.json({ error: err.message || "Internal server error" }, { status: 500 });
  }
}
