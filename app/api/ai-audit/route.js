import { NextResponse } from "next/server";

// Exclude non-chat models (audio, tts, embedding, vision-only)
const EXCLUDE_PATTERNS = ["whisper", "tts", "distil", "embed", "vision", "playai", "playht", "guard"];

// Prefer fast models with higher free-tier TPM limits first
const MODEL_PREFERENCES = [
  "llama-3.1-8b-instant",
  "llama3-8b",
  "llama-3.1-8b",
  "llama-3.3-70b",
  "llama-3.1-70b",
  "llama3-70b",
  "llama",
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

    const chatIds = (data.data || [])
      .map((m) => m.id)
      .filter((id) => !EXCLUDE_PATTERNS.some((pat) => id.toLowerCase().includes(pat)));

    for (const pref of MODEL_PREFERENCES) {
      const found = chatIds.find((id) => id.toLowerCase().includes(pref));
      if (found) return found;
    }
    return chatIds[0] || null;
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

    const model = await pickBestGroqModel(apiKey);
    if (!model) {
      return NextResponse.json(
        { error: "Could not find any available Groq model. Please verify the key at console.groq.com." },
        { status: 400 }
      );
    }

    // Compact rules to minimize prompt tokens
    const RULES = `Rules:
1.Groceries:rice,dal,atta,maida,sooji,spices,masala,oil,sugar,salt,sauces,vinegar,dry fruits,ghee
2.Dairy:milk,paneer,curd,dahi,cream,butter,cheese,khoya,buttermilk
3.Poultry:chicken,mutton,lamb,goat,eggs
4.Sea food purchases:fish,basa,surmai,pomfret,rawas,salmon,tuna,prawns,shrimp,crab,lobster,squid
5.Beverages(non-alcoholic only):red bull,tonic,soda,juice,monin,malas,mineral water,limca,sprite,thums up,tea,coffee
6.Liquor Purchases(alcoholic only):wine,beer,whisky,vodka,rum,gin,tequila,brandy,champagne,scotch,bourbon
7.Cigarette purchases:classic,ice burst,marlboro,gold flake,wills,beedi,hookah tobacco
8.Other Purchases:charcoal,coal,ice cubes,ice slabs,dry ice,skewers,toothpicks
9.Packing materials/Packaging & Disposables(same category):trays,boxes,paper bags,tissues,foil,cling wrap,straws,disposable cutlery
10.Cleaning and housekeeping:dishwash,detergent,lizol,phenyl,harpic,bleach,sanitizer,brooms,mops,garbage bags
11.Kitchen tools:utensils,kadai,tawa,knives,chopping boards,crockery,cocktail shakers
12.Stationery & Office:registers,KOT books,bill books,pens,POS rolls,thermal paper`;

    // Send only 10 items per call to stay under free-tier TPM limit
    const itemsToAudit = items.slice(0, 10).map(({ item, vendor, actualAccount }) => ({ item, vendor, actualAccount }));
    const accounts = (availableAccounts || []).join(",");

    const prompt = `You are a restaurant accounting auditor in India.
${RULES}
Available account heads: ${accounts}
Items to check: ${JSON.stringify(itemsToAudit)}
Return ONLY a JSON array of misclassified items (empty array [] if all correct):
[{"item":"","vendor":"","actualAccount":"","suggestedAccount":"","webSummary":"","reason":""}]`;

    const groqRes = await fetch("https://api.groq.com/openai/v1/chat/completions", {
      method: "POST",
      headers: { "Content-Type": "application/json", Authorization: `Bearer ${apiKey}` },
      body: JSON.stringify({
        model,
        messages: [{ role: "user", content: prompt }],
        temperature: 0.1,
        max_tokens: 1500
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
