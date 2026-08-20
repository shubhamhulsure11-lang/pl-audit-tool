import { NextResponse } from "next/server";

function parseRetryAfter(errorMsg) {
  const m = String(errorMsg || "").match(/try again in (\d+(?:\.\d+)?)\s*s/i);
  return m ? Math.ceil(parseFloat(m[1])) : 30;
}

function cleanThinking(text) {
  if (!text) return "";
  let cleaned = text.replace(/<think>[\s\S]*?<\/think>/gi, "");
  cleaned = cleaned.replace(/<think>[\s\S]*/gi, "");
  return cleaned.trim();
}

export async function POST(req) {
  try {
    const body = await req.json();
    const { singleItem, availableAccounts = [], apiKey, model = "llama-3.1-8b-instant" } = body;

    if (!apiKey) return NextResponse.json({ error: "No API key. Use ⚙️ AI Setup." }, { status: 400 });
    if (!singleItem) return NextResponse.json({ error: "No item provided." }, { status: 400 });

    const { item, vendor, actualAccount } = singleItem;
    const accountList = (availableAccounts || []).slice(0, 10).join(", ");

    const prompt = `Restaurant accounting auditor. Verify if this purchase item is booked in the right Zoho account.

Item: "${item}"
Vendor: "${vendor}"
Current Account in Zoho: "${actualAccount}"
Available Account Heads: ${accountList || "Groceries purchases, Dairy products purchases, Sea food purchases, Poultry and meat purchases, Beverages, Liquor purchases, Cleaning and housekeeping, Packaging & Disposables, Stationery"}

Rules:
- Coconut milk / Coconut milk powder, vanilla, knorr, seasonings, broth powder, food colors, sauces, oil, sugar, spices → Groceries purchases
- Cream cheese, butter, paneer, curd, milk, fresh cream → Dairy products purchases
- Ginger ale, tonic water, soda, red bull, juices, syrups → Beverages (Non-alcoholic)
- Wine, beer, whisky, rum, gin, vodka → Liquor purchases (Alcoholic)
- Banana leaves (for serving/packing), takeaway boxes, paper bags, foil, napkins → Packaging & Disposables / Packing material
- Dishwash, floor cleaner, lizol, detergent, soap oil → Cleaning and housekeeping
- Raw chicken, mutton, lamb, eggs → Poultry and meat purchases
- Raw fish (basa, surmai, pomfret, salmon), prawns, shrimp, crab → Sea food purchases

Return JSON ONLY (no other words, no markdown):
If currently booked account is CORRECT: {"ok":true}
If WRONG: {"ok":false,"suggest":"exact name from Available Account Heads","why":"short reason"}`;

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

    if (res.status === 429) {
      const errText = await res.text();
      let errMsg = "";
      try { errMsg = JSON.parse(errText)?.error?.message || ""; } catch { /* ignore */ }
      const retryAfter = parseRetryAfter(errMsg);
      return NextResponse.json({ error: "Rate limited", retryAfter }, { status: 429 });
    }

    if (!res.ok) {
      const errText = await res.text();
      let errMsg = `Groq error (${res.status})`;
      try { errMsg = JSON.parse(errText)?.error?.message || errMsg; } catch { /* ignore */ }
      return NextResponse.json({ error: errMsg }, { status: res.status });
    }

    const data = await res.json();
    let text = (data.choices?.[0]?.message?.content || "").trim();
    text = cleanThinking(text);

    // Strip markdown fences
    const fenceMatch = text.match(/```(?:json)?\s*([\s\S]*?)\s*```/);
    if (fenceMatch) text = fenceMatch[1].trim();

    // Extract outermost JSON object
    const jsonStart = text.indexOf("{");
    const jsonEnd = text.lastIndexOf("}");
    if (jsonStart >= 0 && jsonEnd > jsonStart) {
      text = text.slice(jsonStart, jsonEnd + 1);
    }

    try {
      const result = JSON.parse(text);
      return NextResponse.json({ success: true, result, model });
    } catch {
      // Parse failed — do NOT fake success. Return error so UI shows ⚠️.
      const snippet = text ? text.slice(0, 120).replace(/\n/g, " ") : "(empty)";
      return NextResponse.json({
        error: `AI returned non-JSON output. Try a different model in ⚙️ AI Setup. Snippet: "${snippet}"`
      }, { status: 422 });
    }
  } catch (err) {
    return NextResponse.json({ error: err.message || "Internal error" }, { status: 500 });
  }
}
