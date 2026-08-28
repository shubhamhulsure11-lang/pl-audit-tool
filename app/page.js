"use client";
import { Fragment, useCallback, useEffect, useMemo, useState } from "react";
import * as XLSX from "xlsx";
import { matchItemToKnowledge, normalizeItemName } from "@/app/lib/knowledge";

const inr = new Intl.NumberFormat("en-IN", { style: "currency", currency: "INR", maximumFractionDigits: 0 });
const clean = (v) => String(v ?? "").trim().toLowerCase().replace(/[^a-z0-9]/g, "");
const n = (v) => typeof v === "number" ? v : Number(String(v ?? "").replace(/[,₹]/g, "")) || 0;
const show = (v) => inr.format(v || 0);
const col = (heads, choices) => choices.map(clean).map(x => heads.map(clean).indexOf(x)).find(x => x >= 0) ?? -1;

// ── TAXONOMY: fallback-only rule engine ─────────────────────────────────────
// NOTE: This is now SECONDARY. Client Knowledge Base is always checked first.
const RESTAURANT_TAXONOMY = [
  {
    id: "cigarettes",
    label: "Cigarette purchases",
    aliases: ["cigarette", "cigar", "tobacco", "smoke"],
    keywords: [
      "classic connect", "classic regular", "classic milds", "classic ultra milds",
      "cl ice burst", "ice burst", "classic bt", "classic ft",
      "marlboro", "marlboro lights", "marlboro red",
      "gold flake", "gold flake lights", "gold flake kings",
      "wills navy cut", "wills lights", "wills milds",
      "benson hedges", "dunhill", "esse", "parliament",
      "cigarette", "cigarettes", "cigar", "hookah tobacco", "shisha"
    ]
  },
  {
    id: "liquor",
    label: "Liquor Purchases",
    aliases: ["liquor", "alcohol", "spirit", "wine", "beer", "whisky", "whiskey", "bar purchase", "excise"],
    keywords: [
      "jacobs creek", "jacob creek", "shiraz", "cabernet", "chardonnay", "merlot", "sauvignon", "pinot noir",
      "red wine", "white wine", "rose wine", "sula wine", "fratelli", "grover", "chandon", "prosecco", "sparkling wine",
      "whisky", "whiskey", "scotch", "bourbon", "single malt",
      "glenfiddich", "black label", "red label", "double black", "chivas regal", "jack daniels",
      "jameson", "ballantines", "teachers", "100 pipers", "vat 69", "antiquity",
      "blenders pride", "royal challenge", "royal stag", "imperial blue", "oakton",
      "vodka", "absolut", "smirnoff", "magic moments", "grey goose", "ciroc", "belvedere",
      "rum", "old monk", "bacardi", "captain morgan",
      "bombay sapphire", "beefeater", "tanqueray", "gordons", "greater than gin", "hendricks", "dry gin", "gin bottle",
      "tequila", "jose cuervo", "patron tequila", "don julio", "camino real", "sauza",
      "brandy", "mansion house", "cognac", "hennessy",
      "beer", "kingfisher", "budweiser", "corona beer", "bira", "heineken", "carlsberg", "tuborg",
      "hoegaarden", "stella artois", "miller", "breezer", "craft beer", "draught beer",
      "baileys", "jagermeister", "kahlua", "cointreau", "campari", "aperol", "sambuca", "liqueur",
      "champagne"
    ]
  },
  {
    id: "seafood",
    label: "Sea food purchases",
    aliases: ["sea food", "seafood", "fish", "prawn", "crab", "marine"],
    keywords: [
      "16/20 prawns", "21/25 prawns", "26/30 prawns", "tiger prawn", "tiger prawns", "jumbo prawn", "jumbo prawns",
      "v. basa", "v basa", "basa fillet", "basa",
      "prawns", "prawn", "shrimp", "shrimps",
      "surmai", "kingfish", "pomfret", "white pomfret", "black pomfret",
      "rawas", "indian salmon", "salmon", "rohu", "katla", "tilapia",
      "tuna", "mackerel", "bangda", "hilsa", "bombil", "bombay duck",
      "squid", "calamari", "octopus", "mud crab", "fresh crab", "lobster", "clams", "oysters",
      "fish fillet", "fish curry cut"
    ]
  },
  {
    id: "poultry",
    label: "Poultry",
    aliases: ["poultry", "chicken", "mutton", "meat", "egg", "nonveg", "non-veg"],
    keywords: [
      "chicken breast", "chicken leg", "chicken drumstick", "chicken wings", "chicken keema",
      "chicken mince", "chicken curry cut", "chicken lollipop", "chicken liver", "boneless chicken",
      "broiler chicken", "country chicken", "desi chicken", "whole chicken", "chicken meat",
      "mutton keema", "mutton curry cut", "mutton chops", "mutton boti", "mutton brain",
      "lamb chops", "lamb shank", "goat meat", "mutton meat", "lamb meat",
      "brown eggs", "white eggs", "quail eggs", "egg tray", "eggs", "fresh egg",
      "beef", "pork", "bacon", "ham", "pepperoni"
    ]
  },
  {
    id: "dairy",
    label: "Dairy",
    aliases: ["dairy", "milk", "cream", "butter", "cheese", "paneer", "curd", "yogurt", "ghee"],
    keywords: [
      "amul butter", "amul cream", "amul cheese", "amul milk", "amul ghee",
      "milky mist", "govardhan ghee", "mother dairy", "nandini",
      "full cream milk", "toned milk", "double toned milk", "skimmed milk",
      "fresh cream", "whipping cream", "heavy cream",
      "mozzarella", "cheddar", "parmesan", "processed cheese", "cheese slice", "cheese block",
      "paneer", "fresh paneer", "malai paneer",
      "curd", "dahi", "yogurt", "greek yogurt",
      "ghee", "clarified butter", "white butter",
      "condensed milk", "evaporated milk", "khoa", "mawa", "rabdi"
    ]
  },
  {
    id: "beverages",
    label: "Beverages",
    aliases: ["beverage", "soft drink", "juice", "soda", "mocktail", "cold drink"],
    keywords: [
      "coca cola", "pepsi", "sprite", "fanta", "7up", "thums up", "limca", "maaza",
      "mirinda", "mountain dew", "red bull", "monster energy",
      "frooti", "slice mango", "real juice", "tropicana",
      "coconut water", "tender coconut", "coco", "kingcoconut",
      "cold coffee", "iced tea", "ice tea", "lemonade",
      "tonic water", "club soda", "soda water", "ginger beer", "ginger ale",
      "monin", "monin syrup", "torani", "da vinci syrup",
      "green tea", "herbal tea", "chai mix",
      "packaged water", "mineral water", "bisleri", "kinley", "aquafina"
    ]
  },
  {
    id: "vegetables",
    label: "Vegetables",
    aliases: ["vegetable", "sabzi", "tarkari", "greens", "fresh produce"],
    keywords: [
      "tomato", "onion", "potato", "garlic", "ginger", "carrot", "beans", "cabbage",
      "capsicum", "bell pepper", "brinjal", "eggplant", "cauliflower", "broccoli",
      "spinach", "palak", "methi", "fenugreek", "coriander", "curry leaves", "kadi patta",
      "mint leaves", "pudina", "green chilli", "red chilli", "chilli",
      "lady finger", "bhindi", "okra", "bitter gourd", "karela",
      "bottle gourd", "lauki", "ridge gourd", "turai", "ash gourd",
      "raw banana", "raw plantain", "yam", "suran", "colocasia", "arbi",
      "drumstick", "moringa", "cluster beans", "valor papdi",
      "spring onion", "leek", "celery", "baby corn", "sweet corn", "zucchini",
      "cherry tomato", "lettuce", "iceberg", "arugula", "kale",
      "mushroom", "button mushroom", "oyster mushroom", "portobello"
    ]
  },
  {
    id: "fruits",
    label: "Fruits",
    aliases: ["fruit", "seasonal fruit", "imported fruit"],
    keywords: [
      "apple", "banana", "mango", "papaya", "guava", "pineapple", "watermelon",
      "muskmelon", "cantaloupe", "grapes", "pomegranate", "chikoo", "sapodilla",
      "orange", "sweet lime", "mosambi", "lemon", "lime",
      "kiwi", "strawberry", "blueberry", "raspberry", "avocado",
      "pear", "peach", "plum", "cherry", "fig", "dates", "dragon fruit", "passion fruit"
    ]
  },
  {
    id: "groceries",
    label: "Groceries / Provisions",
    aliases: ["grocer", "grocery", "provision", "dry goods", "pantry", "raw material"],
    keywords: [
      "rice", "basmati rice", "ponni rice", "sona masoori", "jeera rice",
      "wheat flour", "maida", "atta", "semolina", "suji", "rawa",
      "dal", "lentils", "moong dal", "toor dal", "chana dal", "masoor dal", "urad dal",
      "oil", "sunflower oil", "groundnut oil", "canola oil", "palm oil", "olive oil",
      "coconut oil", "rice bran oil", "saffola", "fortune oil",
      "sugar", "jaggery", "honey",
      "salt", "rock salt", "pink salt",
      "vinegar", "apple cider vinegar",
      "soya sauce", "oyster sauce", "fish sauce", "worcestershire",
      "tomato ketchup", "tomato puree", "tomato paste",
      "corn flour", "corn starch", "arrowroot",
      "baking powder", "baking soda", "yeast",
      "vanilla essence", "food colour", "edible colour",
      "cashew", "almond", "pistachio", "walnut", "raisin", "sultana",
      "peanut", "groundnut",
      "black pepper", "cumin", "jeera", "coriander seed", "turmeric", "haldi",
      "garam masala", "biryani masala", "chaat masala", "kitchen king",
      "red chilli powder", "paprika", "kashmiri chilli",
      "cardamom", "cloves", "cinnamon", "bay leaf", "star anise", "mace", "nutmeg"
    ]
  },
  {
    id: "packaging",
    label: "Packaging / Disposables",
    aliases: ["packaging", "disposable", "takeaway", "parcel", "container", "wrapping"],
    keywords: [
      "food container", "meal box", "lunch box", "parcel box", "takeaway box",
      "foil container", "aluminium container", "aluminium foil",
      "paper cup", "plastic cup", "cold cup", "hot cup",
      "straw", "spoon", "fork", "knife", "plastic cutlery",
      "paper plate", "plastic plate", "banana leaf", "areca leaf",
      "tissue paper", "napkin", "paper napkin",
      "butter paper", "baking paper", "parchment",
      "cling wrap", "cling film", "shrink wrap",
      "ziplock bag", "plastic bag", "carry bag", "grocery bag",
      "paper bag", "kraft bag",
      "toothpick", "skewer", "cocktail pick",
      "cup lid", "container lid"
    ]
  },
  {
    id: "cleaning",
    label: "Cleaning / Housekeeping",
    aliases: ["cleaning", "housekeeping", "sanitation", "hygiene", "laundry"],
    keywords: [
      "vim bar", "vim liquid", "pril", "exo", "dishwash liquid", "dishwash bar",
      "harpic", "domex", "lizol", "toilet cleaner", "drain cleaner",
      "colin", "glass cleaner", "surface cleaner", "floor cleaner", "floor liquid",
      "dettol", "savlon", "antiseptic",
      "hand wash", "liquid hand wash", "hand sanitizer", "sanitizer",
      "soap bar", "bathing soap", "lifebuoy", "lux soap",
      "surf excel", "ariel", "tide", "rin", "washing powder", "detergent powder",
      "comfort", "fabric softener",
      "phenyl", "black phenyl", "white phenyl",
      "broom", "jhadu", "mop", "wiper", "scrubber", "sponge", "scotch brite",
      "gloves", "rubber gloves", "apron", "hair net", "cap",
      "garbage bag", "dustbin liner", "trash bag"
    ]
  },
  {
    id: "kitchen_tools",
    label: "Kitchen Tools / Equipment",
    aliases: ["equipment", "hotelware", "kitchen tool", "crockery", "cutlery", "utensil"],
    keywords: [
      "knife", "chef knife", "bread knife", "peeler", "grater", "chopping board",
      "pan", "kadai", "wok", "frying pan", "saute pan", "stock pot",
      "tongs", "ladle", "spatula", "skimmer", "whisk", "beater",
      "mixing bowl", "salad bowl", "serving bowl",
      "plate", "dinner plate", "side plate", "soup bowl",
      "glass", "tumbler", "mug", "cup", "saucer",
      "spoon", "dessert spoon", "soup spoon", "tablespoon",
      "fork", "salad fork", "dessert fork",
      "tray", "serving tray", "waiter tray",
      "bottle opener", "cork screw", "can opener",
      "thermometer", "kitchen thermometer", "probe"
    ]
  },
  {
    id: "stationery",
    label: "Stationery / Office Supplies",
    aliases: ["stationery", "office supply", "paper", "printing"],
    keywords: [
      "attendance register", "bill book", "kot book", "receipt book", "permanent marker", "ball pen", "gel pen",
      "pos roll", "billing roll", "thermal roll", "printer cartridge", "toner cartridge",
      "register", "notebook", "pencil", "marker", "stapler", "stapler pin",
      "punch machine", "brown tape", "cello tape", "scissor", "scissors",
      "stamp pad", "rubber band", "binder clip", "envelope", "a4 paper", "toner"
    ]
  }
];

function cleanWords(str) {
  return String(str || "").toLowerCase().replace(/[^a-z0-9\s]/g, " ").replace(/\s+/g, " ").trim();
}

function accountMatchesCategory(actualAccount, category) {
  const normAcc = String(actualAccount || "").toLowerCase();
  const isStaffOrExpense = /staff\s*(welfare|food|meal|expense)|entertainment|business\s*meal|petty\s*cash|general\s*expense/i.test(normAcc);
  if (isStaffOrExpense && ["seafood","poultry","groceries","beverages","dairy","vegetables","fruits","liquor"].includes(category.id)) return true;
  if (category.id === "groceries" && normAcc.includes("sea food")) return false;
  if (category.id === "groceries" && (normAcc.includes("poultry") || normAcc.includes("meat"))) return false;
  if (category.id === "groceries" && normAcc.includes("dairy")) return false;
  if (category.id === "beverages" && normAcc.includes("liquor")) return false;
  if (category.aliases.some(alias => normAcc.includes(alias))) return true;
  if (category.id === "packaging" && (normAcc.includes("pack") || normAcc.includes("clean") || normAcc.includes("housekeep") || normAcc.includes("disposab"))) return true;
  if (category.id === "cleaning" && (normAcc.includes("pack") || normAcc.includes("soap") || normAcc.includes("clean") || normAcc.includes("housekeep"))) return true;
  if (category.id === "kitchen_tools" && (normAcc.includes("hotelware") || normAcc.includes("equipment") || normAcc.includes("kitchen") || normAcc.includes("crockery") || normAcc.includes("cutlery"))) return true;
  if ((category.id === "vegetables" || category.id === "fruits") &&
      (normAcc.includes("vegetable") || normAcc.includes("fruit") || normAcc.includes("tarkari"))) return true;
  if (category.id === "groceries" && (normAcc.includes("grocer") || normAcc.includes("provision") || normAcc.includes("raw material"))) return true;
  return false;
}

const VETO_RULES = [
  { blockedCategoryId: "cigarettes", ifItemContains: ["rum", "vodka", "whisky", "whiskey", "gin", "tequila", "brandy", "wine", "beer", "scotch", "bourbon", "liqueur", "bacardi", "smirnoff", "absolut"] },
  { blockedCategoryId: "liquor", ifItemContains: ["ale", "ginger ale", "gin ale", "tonic", "syrup", "essence", "non alcoholic", "non-alcoholic", "mocktail", "vinegar", "cooking wine"] },
  { blockedCategoryId: "groceries", ifItemContains: ["soap", "detergent", "cleaner", "liquid soap", "dishwash", "hand wash"] },
  { blockedCategoryId: "fruits", ifItemContains: ["monin", "syrup", "crush", "malas", "cordial", "patra"] },
  { blockedCategoryId: "vegetables", ifItemContains: ["monin", "syrup", "crush", "malas", "cordial", "patra", "banana leaf", "banana leaves"] },
  { blockedCategoryId: "poultry", ifItemContains: ["broth", "powder", "seasoning", "cube", "bouillon", "knorr", "mix", "curry paste"] },
  { blockedCategoryId: "seafood", ifItemContains: ["broth", "powder", "seasoning", "cube", "bouillon", "knorr", "mix", "vinegar", "sauce", "cake mix"] },
  { blockedCategoryId: "dairy", ifItemContains: ["coconut milk", "coconut", "almond milk", "soy milk", "soya milk", "oat milk", "plant milk", "milk powder"] },
];

function isVetoed(categoryId, itemNorm) {
  return VETO_RULES.some(rule =>
    rule.blockedCategoryId === categoryId &&
    rule.ifItemContains.some(signal => {
      const escaped = signal.replace(/[-\/\\^$*+?.()|[\]{}]/g, "\\$&");
      return new RegExp(`(^|\\s)${escaped}(\\s|$)`, "i").test(itemNorm);
    })
  );
}

function classifyItem(itemName) {
  const norm = cleanWords(itemName);
  if (!norm) return null;
  let bestMatch = null;
  let maxKeywordLen = 0;
  for (const cat of RESTAURANT_TAXONOMY) {
    if (isVetoed(cat.id, norm)) continue;
    for (const kw of cat.keywords) {
      const normKw = cleanWords(kw);
      if (!normKw) continue;
      const escaped = normKw.replace(/[-\/\\^$*+?.()|[\]{}]/g, "\\$&");
      const wordRegex = new RegExp(`(^|\\s)${escaped}(\\s|$)`, "i");
      if (wordRegex.test(norm)) {
        if (normKw.length > maxKeywordLen) {
          maxKeywordLen = normKw.length;
          bestMatch = { category: cat, keyword: kw };
        }
      }
    }
  }
  return bestMatch;
}

// ─── KNOWLEDGE-FIRST MISCLASSIFICATION DETECTION ────────────────────────────
function detectMisclassifications(records, knowledgeItems = []) {
  if (!records || !records.length) return [];

  const sheetAccounts = [...new Set(records.map(r => r.account).filter(a => a && a !== "Unassigned Account"))];

  const findBestSheetAccountName = (category) => {
    const cid = category.id;
    if (cid === "groceries") { const m = sheetAccounts.find(a => /grocer/i.test(a)); if (m) return m; }
    if (cid === "seafood") { const m = sheetAccounts.find(a => /sea\s*food|fish|prawn/i.test(a)); if (m) return m; }
    if (cid === "poultry") { const m = sheetAccounts.find(a => /poultry|meat|chicken|mutton/i.test(a)); if (m) return m; }
    if (cid === "dairy") { const m = sheetAccounts.find(a => /dairy|milk/i.test(a)); if (m) return m; }
    if (cid === "beverages") { const m = sheetAccounts.find(a => /beverage|soft\s*drink/i.test(a) && !/liquor|alcohol|wine|beer/i.test(a)); if (m) return m; }
    if (cid === "liquor") { const m = sheetAccounts.find(a => /liquor|alcohol|wine|beer|spirit/i.test(a)); if (m) return m; }
    if (cid === "packaging") { const m = sheetAccounts.find(a => /pack|disposab/i.test(a)); if (m) return m; }
    if (cid === "cleaning") { const m = sheetAccounts.find(a => /clean|housekeep/i.test(a)); if (m) return m; }
    if (cid === "vegetables" || cid === "fruits") { const m = sheetAccounts.find(a => /vegetable|fruit/i.test(a)); if (m) return m; }
    if (cid === "cigarettes") { const m = sheetAccounts.find(a => /cigarette|tobacco|smoke/i.test(a)); if (m) return m; }
    const matched = sheetAccounts.find(acc => {
      const normA = acc.toLowerCase();
      if (cid === "groceries" && normA.includes("sea food")) return false;
      return category.aliases.some(alias => normA.includes(alias));
    });
    return matched || category.label;
  };

  const map = new Map();
  const hasVegAccount = sheetAccounts.some(a => /vegetable|fruit|sabzi|tarkari/i.test(a));

  records.forEach(r => {
    if (!r.item || !r.account || r.account === "Unassigned Account") return;

    // ── PRIORITY 1: Client Knowledge Base ─────────────────────────────────────
    if (knowledgeItems.length > 0) {
      const kbMatch = matchItemToKnowledge(r.item, knowledgeItems);
      if (kbMatch) {
        const expectedAcc = kbMatch.account_head;
        const normLedger = String(r.account).toLowerCase().trim();
        const normExpected = String(expectedAcc).toLowerCase().trim();

        // Check if the ledger account matches the knowledge account (flexible match)
        const isCorrect =
          normLedger === normExpected ||
          normLedger.includes(normExpected) ||
          normExpected.includes(normLedger);

        if (isCorrect) return;

        const key = `${r.item}:::${r.vendor}:::${r.account}:::kb`;
        if (!map.has(key)) {
          map.set(key, {
            item: r.item,
            vendor: r.vendor,
            actualAccount: r.account,
            suggestedAccount: expectedAcc,
            matchedKeyword: `KB match (${kbMatch.matchType})`,
            reason: `Knowledge Base: "${r.item}" should be under "${expectedAcc}"`,
            confidence: "High",
            source: "knowledge",
            knowledgeMatch: true,
            kbMatchType: kbMatch.matchType,
            count: 0,
            total: 0,
            rates: []
          });
        }
        const entry = map.get(key);
        entry.count += 1;
        entry.total += r.total;
        if (r.rate > 0) entry.rates.push(r.rate);
        return;
      }
    }

    // ── PRIORITY 2: RESTAURANT_TAXONOMY Rule Engine ───────────────────────────
    const match = classifyItem(r.item);
    if (!match) return;

    const expectedCat = match.category;
    const isCorrect = accountMatchesCategory(r.account, expectedCat);

    if (!isCorrect) {
      if (!hasVegAccount &&
          (expectedCat.id === "vegetables" || expectedCat.id === "fruits") &&
          /grocer|provision|raw material|food/i.test(r.account.toLowerCase())) {
        return;
      }

      const suggestedName = findBestSheetAccountName(expectedCat);
      const key = `${r.item}:::${r.vendor}:::${r.account}:::${expectedCat.id}`;
      if (!map.has(key)) {
        map.set(key, {
          item: r.item,
          vendor: r.vendor,
          actualAccount: r.account,
          suggestedAccount: suggestedName,
          matchedKeyword: match.keyword,
          reason: `Rule Engine: matched keyword "${match.keyword}"`,
          confidence: "High",
          source: "rule_engine",
          knowledgeMatch: false,
          count: 0,
          total: 0,
          rates: []
        });
      }
      const entry = map.get(key);
      entry.count += 1;
      entry.total += r.total;
      if (r.rate > 0) entry.rates.push(r.rate);
    }
  });

  return [...map.values()].sort((a, b) => b.total - a.total);
}

async function readFile(file) {
  const data = await file.arrayBuffer();
  const wb = XLSX.read(data, { type: "array", cellDates: true });
  const rows = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { header: 1, defval: "", raw: true });
  const at = rows.findIndex(r => r.map(clean).includes("vendorname") && r.map(clean).some(x => x === "billdate" || x === "date"));
  if (at < 0) throw new Error("I could not find the Zoho header row. Please use a Zoho purchase export containing Bill Date and Vendor Name.");
  const h = rows[at], i = {
    date: col(h, ["Bill Date", "Date", "Invoice Date"]),
    vendor: col(h, ["Vendor Name", "Vendor", "Supplier Name", "Supplier"]),
    bill: col(h, ["Bill Number", "Invoice Number", "Bill No", "Bill#", "Invoice#"]),
    account: col(h, ["Account Name", "Account", "Expense Account", "Account Head", "Chart of Accounts", "Expense Head", "Accounts"]),
    item: col(h, ["Item Name", "Item", "Product Name", "Product", "Description", "Item Description", "Items"]),
    qty: col(h, ["Quantity", "Qty", "Quantity Billed"]),
    rate: col(h, ["Rate", "Item Rate", "Price", "Unit Price"]),
    total: col(h, ["Item Total", "Line Item Total", "Total", "Amount", "Line Amount"]),
    branch: col(h, ["Branch Name", "Branch", "Location"])
  };
  if (i.vendor < 0 || i.date < 0) throw new Error("Bill Date and Vendor Name are required.");
  const get = (row, key) => i[key] >= 0 ? row[i[key]] : "";
  const records = rows.slice(at + 1).map((row, id) => {
    const qty = n(get(row, "qty")), rate = n(get(row, "rate"));
    const rawAcc = String(get(row, "account") || "").trim();
    return {
      id,
      date: String(get(row, "date") || ""),
      vendor: String(get(row, "vendor") || "").trim(),
      bill: String(get(row, "bill") || "").trim(),
      account: rawAcc || "Unassigned Account",
      item: String(get(row, "item") || "").trim(),
      qty,
      rate,
      total: n(get(row, "total")) || (qty * rate),
      branch: String(get(row, "branch") || "").trim()
    };
  }).filter(r => r.vendor && clean(r.vendor) !== "vendorname" && (r.item || r.total));
  return { name: file.name, records, hasAccountCol: i.account >= 0, colMap: i };
}

async function readFileAsObjects(file) {
  const data = await file.arrayBuffer();
  const wb = XLSX.read(data, { type: "array", cellDates: true });
  const rows = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { defval: "" });
  return { rows, name: file.name, headers: rows.length > 0 ? Object.keys(rows[0]) : [] };
}

function groups(rows, key) { return rows.reduce((m, r) => { const k = key(r); m.set(k, [...(m.get(k) || []), r]); return m; }, new Map()); }
function rollup(rows, field) { return [...groups(rows, r => clean(r[field]) || "unassigned")].map(([key, list]) => ({ key, label: list[0][field] || "Unassigned", total: list.reduce((s, r) => s + r.total, 0) })); }

function analyse(current, previous, thresholds, knowledgeItems = []) {
  const { vendor = 20, item = 25, price = 20 } = thresholds || {};
  const vendorCut = vendor / 100, itemCut = item / 100, priceMult = 1 + price / 100;
  const exact = groups(current.records, r => [clean(r.date), clean(r.vendor), clean(r.bill), clean(r.item), r.qty, r.rate, clean(r.branch)].join("|"));
  const near = groups(current.records, r => [clean(r.date), clean(r.vendor), clean(r.item), r.qty, r.rate, clean(r.branch)].join("|"));
  const duplicates = [];
  exact.forEach(rows => { if (rows.length > 1) duplicates.push({ kind: "Confirmed duplicate", risk: "Critical", rows, total: rows.reduce((s, r) => s + r.total, 0) }); });
  near.forEach(rows => { const bills = new Set(rows.map(r => clean(r.bill)).filter(Boolean)); if (rows.length > 1 && bills.size > 1) duplicates.push({ kind: "Same details, different bill no.", risk: "Review", rows, total: rows.reduce((s, r) => s + r.total, 0) }); });
  const compare = (field, cutoff) => {
    const prior = new Map(rollup(previous.records, field).map(x => [x.key, x]));
    return rollup(current.records, field).map(x => {
      const p = prior.get(x.key), old = p?.total || 0, diff = x.total - old;
      return { ...x, old, diff, pct: old ? diff / old : null, status: p ? "Changed" : "New" };
    }).filter(x => x.status === "New" || Math.abs(x.pct || 0) >= cutoff).sort((a, b) => Math.abs(b.diff) - Math.abs(a.diff));
  };
  const prices = [];
  groups(current.records.filter(r => r.item && r.qty && r.rate), r => clean(r.item)).forEach(rows => {
    const avg = rows.reduce((s, r) => s + r.total, 0) / rows.reduce((s, r) => s + r.qty, 0);
    rows.forEach(r => { if (r.rate > avg * priceMult) prices.push({ ...r, avg, pct: (r.rate - avg) / avg }); });
  });

  const misclassifications = detectMisclassifications(current.records, knowledgeItems);

  return { duplicates, vendors: compare("vendor", vendorCut), items: compare("item", itemCut), prices: prices.sort((a, b) => b.pct - a.pct), misclassifications };
}

// ─── UI PRIMITIVES ──────────────────────────────────────────────────────────
function Upload({ title, file, onChange, help }) {
  return <label className="upload"><input type="file" accept=".xlsx,.xls,.csv" onChange={e => e.target.files[0] && onChange(e.target.files[0])} /><span>↑</span><strong>{title}</strong><small>{file?.name || help}</small><em>{file ? "Replace file" : "Choose Excel or CSV"}</em></label>;
}
function Empty({ children }) { return <p className="empty">{children}</p>; }
function Table({ head, children }) { return <div className="table"><table><thead><tr>{head.map(x => <th key={x}>{x}</th>)}</tr></thead><tbody>{children}</tbody></table></div>; }
function ThresholdInput({ label, value, onChange, maxVal = 100 }) {
  return <label className="thresh-label">{label}<div className="thresh-wrap"><input className="thresh-input" type="number" min="0" max={maxVal} value={value} onChange={e => onChange(Math.max(0, Math.min(maxVal, e.target.value === "" ? 0 : Number(e.target.value) || 0)))} /><span className="thresh-pct">%</span></div></label>;
}

// ─── MISCLASSIFICATIONS VIEW ─────────────────────────────────────────────────
function MisclassificationsView({ items, current, onGoToPivot, sharedApiKey, sharedModel, onOpenSetup, sharedAccounts, clientId, onSaveToKnowledge }) {
  const [q, setQ] = useState("");
  const [rowAiState, setRowAiState] = useState({});
  const [aiResults, setAiResults] = useState([]);
  const [savedRows, setSavedRows] = useState(new Set());
  const [savingRows, setSavingRows] = useState(new Set());

  const askAiForRow = async (x) => {
    const rowKey = `${x.item}:::${x.vendor}`;
    if (!sharedApiKey || !sharedModel) { onOpenSetup(); return; }
    setRowAiState(prev => ({ ...prev, [rowKey]: { loading: true, result: null, error: null, countdown: 0 } }));
    try {
      const res = await fetch("/api/ai-audit", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          singleItem: { item: x.item, vendor: x.vendor, actualAccount: x.actualAccount, suggestedAccount: x.suggestedAccount, matchedKeyword: x.matchedKeyword, total: x.total },
          availableAccounts: sharedAccounts,
          apiKey: sharedApiKey,
          model: sharedModel
        })
      });
      const data = await res.json();
      if (res.status === 429) {
        const secs = data.retryAfter || 30;
        setRowAiState(prev => ({ ...prev, [rowKey]: { loading: false, result: null, error: `Rate limited. Retry in ${secs}s`, countdown: secs } }));
        let rem = secs;
        const iv = setInterval(() => {
          rem -= 1;
          if (rem <= 0) { clearInterval(iv); setRowAiState(prev => ({ ...prev, [rowKey]: { loading: false, result: null, error: null, countdown: 0 } })); }
          else { setRowAiState(prev => ({ ...prev, [rowKey]: { ...prev[rowKey], countdown: rem, error: `Retry in ${rem}s` } })); }
        }, 1000);
        return;
      }
      if (!res.ok || data.error) throw new Error(data.error || "AI error");
      const resObj = data.result || {};
      const status = resObj.classification_status || "";
      const isMis = status === "CURRENT_INCORRECT" || status === "BOTH_INCORRECT";
      const aiAccount = resObj.ai_final_account_head || "";
      setRowAiState(prev => ({
        ...prev,
        [rowKey]: {
          loading: false,
          result: {
            isMisclassified: isMis,
            suggestedAccount: isMis ? (aiAccount || x.suggestedAccount) : "",
            why: resObj.ai_reason || "",
            classificationStatus: status,
            currentVerdict: resObj.current_verdict || "",
            suggestedVerdict: resObj.suggested_verdict || "",
            confidence: typeof resObj.confidence === "number" ? resObj.confidence : null,
            reviewRequired: resObj.review_required || false,
            reviewNote: resObj.review_note || "",
            escalated: data.escalated || false,
            usedModel: data.model || ""
          },
          error: null,
          countdown: 0
        }
      }));
      if (isMis && aiAccount) {
        setAiResults(prev => {
          const existing = prev.findIndex(r => r.item === x.item && r.vendor === x.vendor);
          const entry = { ...x, isAi: true, suggestedAccount: aiAccount, matchedKeyword: resObj.ai_reason || "AI Verified", reason: resObj.ai_reason || x.reason };
          if (existing >= 0) { const next = [...prev]; next[existing] = entry; return next; }
          return [...prev, entry];
        });
      }
    } catch (e) {
      setRowAiState(prev => ({ ...prev, [rowKey]: { loading: false, result: null, error: e.message, countdown: 0 } }));
    }
  };

  const saveToKnowledge = async (x) => {
    if (!clientId) { alert("No client selected. Select a client to save to Knowledge Base."); return; }
    const rowKey = `${x.item}:::${x.vendor}`;
    const finalAccount = rowAiState[rowKey]?.result?.suggestedAccount || x.suggestedAccount;
    setSavingRows(prev => new Set([...prev, rowKey]));
    try {
      const res = await fetch("/api/knowledge", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ client_id: clientId, item_name_raw: x.item, account_head: finalAccount, source: "human_approved", notes: `Approved from audit` })
      });
      if (!res.ok) throw new Error("Failed to save");
      setSavedRows(prev => new Set([...prev, rowKey]));
      onSaveToKnowledge && onSaveToKnowledge({ item: x.item, account: finalAccount });
    } catch (e) {
      alert(`Save failed: ${e.message}`);
    } finally {
      setSavingRows(prev => { const n = new Set(prev); n.delete(rowKey); return n; });
    }
  };

  const combinedItems = useMemo(() => {
    const map = new Map();
    items.forEach(it => map.set(it.item, { ...it }));
    aiResults.forEach(it => { map.set(it.item, { ...it, isAi: true }); });
    return [...map.values()].sort((a, b) => b.total - a.total);
  }, [items, aiResults]);

  const filtered = useMemo(() => {
    const term = q.trim().toLowerCase();
    if (!term) return combinedItems;
    return combinedItems.filter(x =>
      x.item.toLowerCase().includes(term) ||
      x.vendor.toLowerCase().includes(term) ||
      x.actualAccount.toLowerCase().includes(term) ||
      x.suggestedAccount.toLowerCase().includes(term) ||
      (x.matchedKeyword && x.matchedKeyword.toLowerCase().includes(term))
    );
  }, [combinedItems, q]);

  const totalExposure = useMemo(() => combinedItems.reduce((s, x) => s + x.total, 0), [combinedItems]);
  const kbCount = combinedItems.filter(x => x.knowledgeMatch).length;
  const ruleCount = combinedItems.filter(x => !x.knowledgeMatch).length;

  return (
    <section className="panel">
      <div className="panelhead">
        <div>
          <p className="eyebrow">KNOWLEDGE-FIRST AUDIT ENGINE</p>
          <h2>Account Head Misclassifications</h2>
        </div>
        <div className="pivot-summary-badges">
          <b className="badge">{combinedItems.length} Flagged Items</b>
          {kbCount > 0 && <b className="badge" style={{ background: "#dbeafe", color: "#1d4ed8", border: "1px solid #bfdbfe" }}>📚 {kbCount} Knowledge Base</b>}
          {ruleCount > 0 && <b className="badge" style={{ background: "#fef3c7", color: "#92400e", border: "1px solid #fde68a" }}>🔬 {ruleCount} Rule Engine</b>}
          <b className="badge accent-badge">{show(totalExposure)} Total Exposure</b>
        </div>
      </div>

      <div className="misclass-info-banner">
        <span>{clientId ? "📚" : "🔬"}</span>
        <div style={{ flex: 1 }}>
          {clientId ? (
            <><strong>Knowledge-First Audit:</strong> Items matched from client Knowledge Base are flagged with 📚. Unknown items use <strong>Rule Engine</strong>, then <strong>🤖 AI</strong>. Save approved corrections to KB with 💾.</>
          ) : (
            <><strong>Rule Engine Only:</strong> No client selected. Select a client with an initialized Knowledge Base to enable knowledge-first classification.</>
          )}
        </div>
        <div style={{ display: "flex", gap: 8 }}>
          <button className="pivot-btn" onClick={onOpenSetup}>⚙️ AI Config</button>
          <button className="pivot-btn" onClick={onGoToPivot} style={{ whiteSpace: "nowrap" }}>View Pivot &rarr;</button>
        </div>
      </div>

      <div className="pivot-toolbar">
        <div className="pivot-search-wrap">
          <span>🔍</span>
          <input type="text" className="pivot-search" placeholder="Search flagged item, vendor, or account..." value={q} onChange={e => setQ(e.target.value)} />
          {q && <button className="pivot-clear-btn" onClick={() => setQ("")}>✕</button>}
        </div>
      </div>

      <Table head={["Item Description", "Vendor", "Current Account Head", "Correct Account Head", "Source / Reason", "Amount", "Actions"]}>
        {filtered.length ? filtered.map((x, i) => {
          const rowKey = `${x.item}:::${x.vendor}`;
          const rowAi = rowAiState[rowKey];
          const ai = rowAi?.result;
          const isSaved = savedRows.has(rowKey);
          const isSaving = savingRows.has(rowKey);

          const statusColors = {
            CURRENT_CORRECT: { bg: "#dcfce7", color: "#15803d", border: "#bbf7d0" },
            BOTH_CORRECT:    { bg: "#dcfce7", color: "#15803d", border: "#bbf7d0" },
            SUGGESTION_CORRECT: { bg: "#dbeafe", color: "#1d4ed8", border: "#bfdbfe" },
            CURRENT_INCORRECT:  { bg: "#fff7ed", color: "#c2410c", border: "#fed7aa" },
            SUGGESTION_INCORRECT: { bg: "#fff7ed", color: "#c2410c", border: "#fed7aa" },
            BOTH_INCORRECT:  { bg: "#fee2e2", color: "#b91c1c", border: "#fecaca" },
            REVIEW_REQUIRED: { bg: "#fef9c3", color: "#92400e", border: "#fde68a" },
          };
          const sc = ai?.classificationStatus ? (statusColors[ai.classificationStatus] || { bg: "#f3f4f6", color: "#374151", border: "#e5e7eb" }) : null;

          return (
            <tr key={i}>
              <td>
                <div style={{ fontWeight: 600, color: "var(--ink)", display: "flex", alignItems: "center", gap: 6 }}>
                  {x.knowledgeMatch && <span title="Flagged by Knowledge Base" style={{ fontSize: 13 }}>📚</span>}
                  {x.item}
                </div>
                {!ai && x.reason && <small style={{ color: "var(--muted)", display: "block", marginTop: "3px", fontSize: "11px" }}>{x.reason}</small>}
                {ai?.classificationStatus && sc && (
                  <span style={{ display: "inline-flex", alignItems: "center", gap: 4, marginTop: 5, padding: "2px 7px", borderRadius: 12, fontSize: "10.5px", fontWeight: 700, background: sc.bg, color: sc.color, border: `1px solid ${sc.border}`, letterSpacing: "0.02em" }}>
                    {ai.classificationStatus.replace(/_/g, " ")}
                    {ai.confidence != null && ` · ${ai.confidence}%`}
                    {ai.escalated && <span title="Escalated to 70B" style={{ marginLeft: 3, fontSize: "9px", opacity: 0.8 }}>70B↑</span>}
                  </span>
                )}
                {rowAi?.error && <small style={{ color: rowAi.countdown > 0 ? "#b45309" : "#dc2626", display: "block", marginTop: "3px", fontSize: "11px" }}>{rowAi.countdown > 0 ? `⏳ ${rowAi.error}` : `⚠ ${rowAi.error}`}</small>}
              </td>

              <td>{x.vendor}</td>

              <td>
                <span className="pill-actual-acc">{x.actualAccount}</span>
                {ai?.currentVerdict === "CORRECT" && <span style={{ display: "block", marginTop: 3, fontSize: "10px", color: "#16a34a", fontWeight: 700 }}>✓ Correct</span>}
                {ai?.currentVerdict === "INCORRECT" && <span style={{ display: "block", marginTop: 3, fontSize: "10px", color: "#dc2626", fontWeight: 700 }}>✗ Incorrect</span>}
                {ai?.currentVerdict === "UNCERTAIN" && <span style={{ display: "block", marginTop: 3, fontSize: "10px", color: "#d97706", fontWeight: 700 }}>? Uncertain</span>}
              </td>

              <td>
                {ai?.classificationStatus ? (
                  ai.isMisclassified && ai.suggestedAccount ? (
                    <>
                      <span className="pill-suggested-acc">🤖 {ai.suggestedAccount}</span>
                      {ai.suggestedVerdict === "CORRECT" && <span style={{ display: "block", marginTop: 3, fontSize: "10px", color: "#1d4ed8", fontWeight: 700 }}>✓ AI Confirmed</span>}
                    </>
                  ) : ai.reviewRequired ? (
                    <span style={{ fontSize: "11px", color: "#92400e", fontWeight: 600 }}>⚠ Review required</span>
                  ) : (
                    <span style={{ fontSize: "11px", color: "#15803d", fontWeight: 600 }}>✅ Correctly booked</span>
                  )
                ) : (
                  x.knowledgeMatch
                    ? <span className="pill-suggested-acc" style={{ background: "#dbeafe", borderColor: "#bfdbfe", color: "#1e40af" }}>📚 {x.suggestedAccount}</span>
                    : <span className="pill-suggested-acc">&rarr; {x.suggestedAccount}</span>
                )}
              </td>

              <td>
                {ai?.why ? (
                  <span className="pill-ai-badge">🤖 {ai.why}</span>
                ) : ai?.reviewNote ? (
                  <span className="pill-ai-badge" style={{ background: "#fef9c3", color: "#92400e", border: "1px solid #fde68a" }}>⚠ {ai.reviewNote}</span>
                ) : x.knowledgeMatch ? (
                  <span className="pill-ai-badge" style={{ background: "#dbeafe", color: "#1e40af", border: "1px solid #bfdbfe" }}>📚 Knowledge Base</span>
                ) : x.isAi ? (
                  <span className="pill-ai-badge">🤖 {x.matchedKeyword}</span>
                ) : (
                  <span className="rule-tag">matched "{x.matchedKeyword}"</span>
                )}
              </td>

              <td style={{ fontWeight: 700, fontFamily: "var(--font-mono, monospace)", whiteSpace: "nowrap" }}>{show(x.total)}</td>

              <td>
                <div style={{ display: "flex", flexDirection: "column", gap: 4 }}>
                  <button
                    className={`ask-ai-row-btn ${rowAi?.countdown > 0 ? "rate-limited" : ""}`}
                    onClick={() => askAiForRow(x)}
                    disabled={rowAi?.loading || rowAi?.countdown > 0}
                    title={rowAi?.countdown > 0 ? `Rate limited, retry in ${rowAi.countdown}s` : "AI second opinion"}
                  >
                    {rowAi?.loading ? "..." : rowAi?.countdown > 0 ? `${rowAi.countdown}s` : ai ? "🔄" : "🤖"}
                  </button>
                  {clientId && (
                    <button
                      className="ask-ai-row-btn"
                      onClick={() => saveToKnowledge(x)}
                      disabled={isSaved || isSaving}
                      title={isSaved ? "Saved to Knowledge Base" : "Save approved classification to Knowledge Base"}
                      style={{ background: isSaved ? "#dcfce7" : undefined, color: isSaved ? "#15803d" : undefined, fontSize: 11 }}
                    >
                      {isSaving ? "…" : isSaved ? "✓ Saved" : "💾 Save"}
                    </button>
                  )}
                </div>
              </td>
            </tr>
          );
        }) : (
          <tr>
            <td colSpan="7">
              <Empty>{combinedItems.length === 0 ? "🎉 No misclassifications detected! Use 🤖 Ask AI on any item for a second opinion." : "No items match your search."}</Empty>
            </td>
          </tr>
        )}
      </Table>
    </section>
  );
}

// ─── ACCOUNT PIVOT ───────────────────────────────────────────────────────────
function AccountPivot({ current }) {
  const [expandedAccs, setExpandedAccs] = useState(() => new Set());
  const [expandedVendors, setExpandedVendors] = useState(() => new Set());
  const [query, setQuery] = useState("");

  const pivotData = useMemo(() => {
    if (!current?.records) return { tree: [], totalSpend: 0, accCount: 0, vendorCount: 0 };
    const accMap = new Map();
    let totalSpend = 0;
    const allVendors = new Set();
    current.records.forEach(r => {
      totalSpend += r.total;
      allVendors.add(r.vendor);
      if (!accMap.has(r.account)) accMap.set(r.account, new Map());
      const vMap = accMap.get(r.account);
      if (!vMap.has(r.vendor)) vMap.set(r.vendor, []);
      vMap.get(r.vendor).push(r);
    });
    const q = query.trim().toLowerCase();
    const tree = [];
    accMap.forEach((vendorMap, accName) => {
      const vendors = [];
      let accSpend = 0;
      vendorMap.forEach((items, vName) => {
        const vSpend = items.reduce((s, r) => s + r.total, 0);
        accSpend += vSpend;
        const itemMap = new Map();
        items.forEach(it => {
          const key = it.item || "(No Item Name)";
          if (!itemMap.has(key)) itemMap.set(key, { name: key, qty: 0, total: 0, rates: [] });
          const entry = itemMap.get(key);
          entry.qty += it.qty;
          entry.total += it.total;
          if (it.rate > 0) entry.rates.push(it.rate);
        });
        const consolidatedItems = [...itemMap.values()].map(it => ({ name: it.name, qty: it.qty, total: it.total, rate: it.rates.length ? (it.total / (it.qty || 1)) : 0 })).sort((a, b) => b.total - a.total);
        const matchesQuery = !q || accName.toLowerCase().includes(q) || vName.toLowerCase().includes(q) || consolidatedItems.some(it => it.name.toLowerCase().includes(q));
        if (matchesQuery) vendors.push({ name: vName, spend: vSpend, items: consolidatedItems });
      });
      if (vendors.length > 0) { vendors.sort((a, b) => b.spend - a.spend); tree.push({ name: accName, spend: accSpend, vendors }); }
    });
    tree.sort((a, b) => b.spend - a.spend);
    return { tree, totalSpend, accCount: accMap.size, vendorCount: allVendors.size };
  }, [current, query]);

  const allAccKeys = useMemo(() => pivotData.tree.map(a => a.name), [pivotData.tree]);
  const allVendorKeys = useMemo(() => { const keys = []; pivotData.tree.forEach(a => a.vendors.forEach(v => keys.push(`${a.name}:::${v.name}`))); return keys; }, [pivotData.tree]);
  const toggleAcc = (name) => setExpandedAccs(prev => { const next = new Set(prev); if (next.has(name)) next.delete(name); else next.add(name); return next; });
  const toggleVendor = (key) => setExpandedVendors(prev => { const next = new Set(prev); if (next.has(key)) next.delete(key); else next.add(key); return next; });
  const expandAll = () => { setExpandedAccs(new Set(allAccKeys)); setExpandedVendors(new Set(allVendorKeys)); };
  const collapseAll = () => { setExpandedAccs(new Set()); setExpandedVendors(new Set()); };
  const isSearching = query.trim().length > 0;
  const isAccOpen = (name) => isSearching || expandedAccs.has(name);
  const isVendorOpen = (key) => isSearching || expandedVendors.has(key);

  return (
    <section className="panel">
      <div className="panelhead">
        <div><p className="eyebrow">CURRENT MONTH P&L HIERARCHY</p><h2>Account Heads Breakdown</h2></div>
        <div className="pivot-summary-badges">
          <b className="badge">{pivotData.accCount} Account Heads</b>
          <b className="badge">{pivotData.vendorCount} Vendors</b>
          <b className="badge accent-badge">{show(pivotData.totalSpend)} Total</b>
        </div>
      </div>
      {!current?.hasAccountCol && (
        <div className="pivot-warning">
          <span>ℹ️</span>
          <div><strong>No "Account" column found in export</strong><small>All items are grouped under "Unassigned Account". Include an "Account Name" or "Expense Account" column in your Zoho export for full P&L categorization.</small></div>
        </div>
      )}
      <div className="pivot-toolbar">
        <div className="pivot-search-wrap">
          <span>🔍</span>
          <input type="text" className="pivot-search" placeholder="Search account head, vendor, or item..." value={query} onChange={e => setQuery(e.target.value)} />
          {query && <button className="pivot-clear-btn" onClick={() => setQuery("")}>✕</button>}
        </div>
        <div className="pivot-actions">
          <button className="pivot-btn" onClick={expandAll}>Expand all</button>
          <button className="pivot-btn" onClick={collapseAll}>Collapse all</button>
        </div>
      </div>
      <div className="pivot-tree">
        {pivotData.tree.length === 0 ? (
          <Empty>No accounts match your search.</Empty>
        ) : (
          pivotData.tree.map(acc => {
            const accOpen = isAccOpen(acc.name);
            const accShare = pivotData.totalSpend ? ((acc.spend / pivotData.totalSpend) * 100).toFixed(1) : 0;
            return (
              <div className="pivot-acc-block" key={acc.name}>
                <div className="pivot-acc-header" onClick={() => toggleAcc(acc.name)}>
                  <div className="pivot-left">
                    <span className="pivot-toggle-icon">{accOpen ? "−" : "+"}</span>
                    <strong className="pivot-acc-title">{acc.name}</strong>
                    <span className="pivot-count-pill">{acc.vendors.length} {acc.vendors.length === 1 ? "vendor" : "vendors"}</span>
                  </div>
                  <div className="pivot-right">
                    <span className="pivot-share">{accShare}% of spend</span>
                    <b className="pivot-acc-spend">{show(acc.spend)}</b>
                  </div>
                </div>
                {accOpen && (
                  <div className="pivot-vendors-list">
                    {acc.vendors.map(v => {
                      const vKey = `${acc.name}:::${v.name}`;
                      const vOpen = isVendorOpen(vKey);
                      return (
                        <div className="pivot-vendor-block" key={v.name}>
                          <div className="pivot-vendor-header" onClick={() => toggleVendor(vKey)}>
                            <div className="pivot-left">
                              <span className="pivot-toggle-icon sub">{vOpen ? "−" : "+"}</span>
                              <strong className="pivot-vendor-title">{v.name}</strong>
                              <span className="pivot-count-pill sub">{v.items.length} {v.items.length === 1 ? "item" : "items"}</span>
                            </div>
                            <div className="pivot-right"><b className="pivot-vendor-spend">{show(v.spend)}</b></div>
                          </div>
                          {vOpen && (
                            <div className="pivot-items-table">
                              <table>
                                <thead><tr><th>Item Description</th><th style={{ textAlign: "right" }}>Quantity</th><th style={{ textAlign: "right" }}>Avg Rate</th><th style={{ textAlign: "right" }}>Total Amount</th></tr></thead>
                                <tbody>
                                  {v.items.map((it, idx) => (
                                    <tr key={idx}>
                                      <td className="pivot-item-name"><span className="bullet">•</span> {it.name}</td>
                                      <td style={{ textAlign: "right", fontFamily: "var(--font-mono, monospace)" }}>{it.qty > 0 ? it.qty.toLocaleString("en-IN") : "-"}</td>
                                      <td style={{ textAlign: "right", fontFamily: "var(--font-mono, monospace)" }}>{it.rate > 0 ? show(it.rate) : "-"}</td>
                                      <td style={{ textAlign: "right", fontWeight: "600", fontFamily: "var(--font-mono, monospace)" }}>{show(it.total)}</td>
                                    </tr>
                                  ))}
                                </tbody>
                              </table>
                            </div>
                          )}
                        </div>
                      );
                    })}
                  </div>
                )}
              </div>
            );
          })
        )}
      </div>
    </section>
  );
}

// ─── DATA PREVIEW ────────────────────────────────────────────────────────────
// ─── DATA PREVIEW ────────────────────────────────────────────────────────────
function DataPreview({ data }) {
  if (!data) return null;
  const { name, records, colMap } = data;
  const sample = records.slice(0, 20);
  const blankItem = records.filter(r => !r.item).length;
  const blankAccount = records.filter(r => !r.account || r.account === "Unassigned Account").length;
  return (
    <section className="panel" style={{ marginTop: 0 }}>
      <div className="panelhead">
        <div><p className="eyebrow">DATA VERIFICATION</p><h2>Parsed File Preview — {name}</h2></div>
        <div className="pivot-summary-badges">
          <b className="badge">{records.length} rows parsed</b>
          {blankAccount > 0 && <b className="badge" style={{ background: "#fef3c7", color: "#92400e" }}>⚠ {blankAccount} rows missing Account</b>}
          {blankItem > 0 && <b className="badge" style={{ background: "#fef3c7", color: "#92400e" }}>⚠ {blankItem} rows missing Item Name</b>}
        </div>
      </div>
      {colMap && (
        <div className="col-map-bar">
          {Object.entries(colMap).filter(([, v]) => v !== null).map(([k, v]) => (
            <span key={k} className="col-map-tag"><strong>{k}</strong> → Col {v}</span>
          ))}
        </div>
      )}
      <div className="table">
        <table>
          <thead>
            <tr>
              <th>#</th><th>Date</th><th>Vendor</th><th>Bill No.</th>
              <th style={{ background: blankAccount > 0 ? "#fef9c3" : undefined }}>Account Head</th>
              <th style={{ background: blankItem > 0 ? "#fef9c3" : undefined }}>Item Name</th>
              <th>Qty</th><th>Rate</th><th>Total</th>
            </tr>
          </thead>
          <tbody>
            {sample.map((r, i) => (
              <tr key={i}>
                <td style={{ color: "var(--muted)", fontSize: 11 }}>{i + 1}</td>
                <td>{r.date}</td>
                <td style={{ fontWeight: 500 }}>{r.vendor}</td>
                <td style={{ color: "var(--muted)" }}>{r.bill || "-"}</td>
                <td style={{ background: (!r.account || r.account === "Unassigned Account") ? "#fef3c7" : undefined }}>{r.account || <em style={{ color: "#ef4444" }}>Missing</em>}</td>
                <td style={{ background: !r.item ? "#fef3c7" : undefined }}>{r.item || <em style={{ color: "#ef4444" }}>Missing</em>}</td>
                <td>{r.qty || "-"}</td>
                <td>{r.rate ? show(r.rate) : "-"}</td>
                <td style={{ fontWeight: 600 }}>{r.total ? show(r.total) : "-"}</td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
      {records.length > 20 && <p style={{ padding: "8px 16px", color: "var(--muted)", fontSize: 12 }}>Showing first 20 of {records.length} rows.</p>}
    </section>
  );
}

// ─── AI SETUP MODAL ──────────────────────────────────────────────────────────
function AiSetupModal({ isOpen, onClose, apiKey, aiConfig, onSave }) {
  const [step, setStep] = useState("key");
  const [keyInput, setKeyInput] = useState(apiKey || "");
  const [selectedModel, setSelectedModel] = useState(aiConfig?.model || "");
  const [availableModels, setAvailableModels] = useState([]);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState("");

  useEffect(() => {
    setKeyInput(apiKey || "");
    setSelectedModel(aiConfig?.model || "");
    if (apiKey && aiConfig?.verified) setStep("configured");
    else setStep("key");
    setError("");
  }, [isOpen, apiKey, aiConfig]);

  const fetchModels = async (overrideKey) => {
    const key = overrideKey || keyInput.trim();
    if (!key) { setError("Please enter your Groq API key (starts with gsk_...)"); return; }
    setError(""); setLoading(true);
    try {
      const res = await fetch("/api/ai-config/test", { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify({ apiKey: key }) });
      const data = await res.json();
      if (!res.ok || data.error) throw new Error(data.error || "Failed to validate key with Groq");
      const models = data.availableModels || [];
      if (!models.length) throw new Error("No active chat models found for this Groq key.");
      setAvailableModels(models);
      const choice = (aiConfig?.model && models.includes(aiConfig.model)) ? aiConfig.model : (data.model || models[0]);
      setSelectedModel(choice);
      setStep("model");
    } catch (e) { setError(e.message); } finally { setLoading(false); }
  };

  const testAndSave = async () => {
    if (!selectedModel) { setError("Please choose a model"); return; }
    setLoading(true); setError(""); setStep("testing");
    try {
      const res = await fetch("/api/ai-config/test", { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify({ apiKey: keyInput.trim(), model: selectedModel }) });
      const data = await res.json();
      if (!res.ok || data.error) throw new Error(data.error || "Model verification failed");
      onSave(keyInput.trim(), data.config);
      setStep("configured");
    } catch (e) { setError(e.message); setStep("model"); } finally { setLoading(false); }
  };

  if (!isOpen) return null;
  return (
    <div className="setup-overlay">
      <div className="setup-modal">
        <div className="setup-modal-header">
          <strong>⚙️ AI Setup &amp; Configuration</strong>
          <button className="setup-close-btn" onClick={onClose}>✕</button>
        </div>
        <div className="setup-body">
          {step === "key" && (
            <>
              <p className="setup-desc">
                {aiConfig?.verified && !apiKey ? <span style={{ color: "#d97706", fontWeight: 600, display: "block", marginBottom: 6 }}>🔑 Session Expired: Please enter your Groq API key to activate AI for this session.</span> : <>Paste your free Groq API key from <a href="https://console.groq.com/keys" target="_blank" rel="noreferrer">console.groq.com/keys</a>.</>}
              </p>
              <label className="setup-label">Groq API Key</label>
              <input type="password" className="setup-input" placeholder="gsk_..." value={keyInput} onChange={e => setKeyInput(e.target.value)} onKeyDown={e => e.key === "Enter" && fetchModels()} />
              {error && <div className="setup-error"><span>⚠</span> {error}</div>}
              <div className="setup-actions">
                <button className="btn-cancel" onClick={onClose}>Cancel</button>
                <button className="btn-save" onClick={() => fetchModels()} disabled={loading || !keyInput.trim()}>{loading ? "Verifying Key..." : "Next: Choose Model →"}</button>
              </div>
            </>
          )}
          {step === "model" && (
            <>
              <p className="setup-desc">Choose the model from your active Groq account:</p>
              <label className="setup-label">Available Models for your Key</label>
              <select className="setup-select" value={selectedModel} onChange={e => setSelectedModel(e.target.value)}>
                {availableModels.map(m => <option key={m} value={m}>{m} {m.includes("instant") ? "⚡ (Recommended)" : m.includes("70b") ? "🧠 (Deep Accuracy)" : ""}</option>)}
              </select>
              <div className="model-info-box"><strong>Tip for Free Tier</strong> <code>llama-3.1-8b-instant</code> offers the best balance of speed and high free tokens-per-minute limits.</div>
              {error && <div className="setup-error"><span>⚠</span> {error}</div>}
              <div className="setup-actions">
                <button className="btn-cancel" onClick={() => setStep("key")}>← Change Key</button>
                <button className="btn-save" onClick={testAndSave} disabled={loading || !selectedModel}>{loading ? "Verifying..." : "Test & Save Config →"}</button>
              </div>
            </>
          )}
          {step === "testing" && (
            <div className="setup-testing"><span className="setup-spinner">⚙️</span><p>Performing live connection test with <strong>{selectedModel}</strong>...</p></div>
          )}
          {step === "configured" && (
            <div className="setup-success">
              <span className="setup-success-icon">✅</span>
              <h4>AI Configuration Verified</h4>
              <p>Active Model: <strong>{aiConfig?.model || selectedModel}</strong></p>
              <p style={{ fontSize: 11, color: "var(--muted)", marginTop: 4 }}>Provider: Groq · Verified Live</p>
              <div className="setup-actions" style={{ marginTop: 16 }}>
                <button className="btn-cancel" onClick={() => fetchModels(apiKey)}>Change Model</button>
                <button className="btn-save" onClick={onClose}>Done</button>
              </div>
            </div>
          )}
        </div>
      </div>
    </div>
  );
}

// ─── AI CHAT WIDGET ──────────────────────────────────────────────────────────
function AiChatWidget({ availableAccounts, apiKey, model, onOpenSetup }) {
  const [open, setOpen] = useState(false);
  const [input, setInput] = useState("");
  const [history, setHistory] = useState([]);
  const [loading, setLoading] = useState(false);
  const [chatCountdown, setChatCountdown] = useState(0);

  const sendMessage = async () => {
    const msg = input.trim();
    if (!msg || loading || chatCountdown > 0) return;
    if (!apiKey || !model) { onOpenSetup(); return; }
    const newHistory = [...history, { role: "user", content: msg }];
    setHistory(newHistory); setInput(""); setLoading(true);
    try {
      const res = await fetch("/api/ai-chat", { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify({ message: msg, history: history.slice(-4), availableAccounts, apiKey, model }) });
      const data = await res.json();
      if (res.status === 429) {
        const secs = data.retryAfter || 30;
        setChatCountdown(secs);
        setHistory([...newHistory, { role: "assistant", content: `⏳ Rate limit reached. Please wait ${secs}s before asking again.` }]);
        let rem = secs;
        const iv = setInterval(() => { rem -= 1; if (rem <= 0) { clearInterval(iv); setChatCountdown(0); } else setChatCountdown(rem); }, 1000);
        return;
      }
      if (data.error) throw new Error(data.error);
      setHistory([...newHistory, { role: "assistant", content: data.reply }]);
    } catch (e) { setHistory([...newHistory, { role: "assistant", content: `⚠️ Error: ${e.message}` }]); }
    finally { setLoading(false); }
  };

  return (
    <>
      <button className="chat-fab" onClick={() => setOpen(o => !o)} title="AI Product Assistant">
        {open ? "✕" : "🤖"}
        {!open && <span className="chat-fab-label">AI Assistant</span>}
      </button>
      {open && (
        <div className="chat-panel">
          <div className="chat-panel-header">
            <div><strong>🤖 AI Assistant</strong><small>{model ? `Using ${model}` : "Click ⚙️ to configure Groq"}</small></div>
            <div style={{ display: "flex", gap: 6 }}>
              <button className="chat-key-btn" onClick={onOpenSetup} title="Configure AI Key & Model">⚙️</button>
              <button className="chat-close-btn" onClick={() => setOpen(false)}>✕</button>
            </div>
          </div>
          <div className="chat-messages">
            {history.length === 0 && (
              <div className="chat-empty">
                <p>👋 Hi! Paste any item name or ask about restaurant accounting.</p>
                {!apiKey && <p style={{ marginTop: 10 }}><button className="pivot-btn" onClick={onOpenSetup} style={{ width: "100%", justifyContent: "center" }}>⚙️ Setup Groq API Key & Model</button></p>}
                <p style={{ marginTop: 10, fontSize: 12, fontWeight: 600 }}>Try asking:</p>
                <div className="chat-examples">
                  {["What is VANILLA 4LTR FD 368?", "Is Monin Watermelon a beverage?", "Which account for Classic Connect FTK?"].map(ex => (
                    <button key={ex} onClick={() => setInput(ex)} className="chat-example-chip">{ex}</button>
                  ))}
                </div>
              </div>
            )}
            {history.map((m, i) => <div key={i} className={`chat-msg chat-msg-${m.role}`}><div className="chat-bubble">{m.content}</div></div>)}
            {loading && <div className="chat-msg chat-msg-assistant"><div className="chat-bubble chat-typing">⠋ Thinking...</div></div>}
          </div>
          <div className="chat-input-row">
            <input type="text" className="chat-input" placeholder={chatCountdown > 0 ? `Wait ${chatCountdown}s...` : "Type item name..."} value={input} onChange={e => setInput(e.target.value)} onKeyDown={e => e.key === "Enter" && !e.shiftKey && sendMessage()} disabled={loading || chatCountdown > 0} />
            <button className="chat-send-btn" onClick={sendMessage} disabled={loading || !input.trim() || chatCountdown > 0}>{loading ? "…" : chatCountdown > 0 ? `${chatCountdown}` : "↑"}</button>
          </div>
        </div>
      )}
    </>
  );
}

// ─── CLIENT MANAGER SCREEN ───────────────────────────────────────────────────
function ClientManagerScreen({ clients, loading, onSelect, onCreated }) {
  const [showCreate, setShowCreate] = useState(false);
  const [newName, setNewName] = useState("");
  const [creating, setCreating] = useState(false);
  const [createError, setCreateError] = useState("");

  const createClient = async () => {
    if (!newName.trim()) return;
    setCreating(true); setCreateError("");
    try {
      const res = await fetch("/api/clients", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ display_name: newName.trim() })
      });
      const data = await res.json();
      if (!res.ok) throw new Error(data.error || "Failed to create client");
      onCreated(data.client);
      setNewName(""); setShowCreate(false);
    } catch (e) { setCreateError(e.message); }
    finally { setCreating(false); }
  };

  return (
    <section className="landing">
      <p className="eyebrow">CLIENT KNOWLEDGE ENGINE</p>
      <h2>Select a Client to begin the audit</h2>
      <p className="lead">Each client has their own persistent Knowledge Base. Select an existing client or create a new one.</p>

      {loading ? (
        <div style={{ textAlign: "center", padding: "40px 0", color: "var(--muted)" }}>Loading clients...</div>
      ) : (
        <>
          {clients.length === 0 && !showCreate && (
            <div style={{ textAlign: "center", padding: "32px 0" }}>
              <p style={{ color: "var(--muted)", marginBottom: 16 }}>No clients yet. Create your first client to get started.</p>
            </div>
          )}

          <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fill, minmax(280px, 1fr))", gap: 16, margin: "24px 0" }}>
            {clients.map(c => (
              <button
                key={c.id}
                onClick={() => onSelect(c)}
                style={{
                  background: "var(--surface)",
                  border: "1.5px solid var(--border)",
                  borderRadius: 12,
                  padding: "20px 24px",
                  cursor: "pointer",
                  textAlign: "left",
                  transition: "all 0.2s",
                  display: "flex",
                  flexDirection: "column",
                  gap: 10
                }}
                onMouseEnter={e => { e.currentTarget.style.borderColor = "var(--accent)"; e.currentTarget.style.boxShadow = "0 4px 20px rgba(0,0,0,0.12)"; }}
                onMouseLeave={e => { e.currentTarget.style.borderColor = "var(--border)"; e.currentTarget.style.boxShadow = "none"; }}
              >
                <div style={{ fontWeight: 700, fontSize: 17, color: "var(--ink)" }}>{c.display_name}</div>
                <div>
                  {c.knowledge_initialized ? (
                    <span style={{ display: "inline-flex", alignItems: "center", gap: 5, background: "#dcfce7", color: "#15803d", border: "1px solid #bbf7d0", borderRadius: 20, padding: "3px 10px", fontSize: 12, fontWeight: 700 }}>
                      ✅ KB Ready · {c.init_item_count?.toLocaleString("en-IN") || "?"} items
                    </span>
                  ) : (
                    <span style={{ display: "inline-flex", alignItems: "center", gap: 5, background: "#fef3c7", color: "#92400e", border: "1px solid #fde68a", borderRadius: 20, padding: "3px 10px", fontSize: 12, fontWeight: 700 }}>
                      ⚠️ Knowledge Not Initialized
                    </span>
                  )}
                </div>
                {c.initialized_at && (
                  <div style={{ fontSize: 11, color: "var(--muted)" }}>
                    Initialized {new Date(c.initialized_at).toLocaleDateString("en-IN", { day: "numeric", month: "short", year: "numeric" })}
                  </div>
                )}
              </button>
            ))}

            {showCreate && (
              <div style={{ background: "var(--surface)", border: "1.5px solid var(--accent)", borderRadius: 12, padding: "20px 24px" }}>
                <div style={{ fontWeight: 600, marginBottom: 10, color: "var(--ink)" }}>New Client</div>
                <input
                  type="text"
                  placeholder="Client / Restaurant name"
                  value={newName}
                  onChange={e => setNewName(e.target.value)}
                  onKeyDown={e => e.key === "Enter" && createClient()}
                  autoFocus
                  style={{ width: "100%", padding: "8px 12px", border: "1px solid var(--border)", borderRadius: 8, background: "var(--bg)", color: "var(--ink)", fontSize: 14, marginBottom: 8, boxSizing: "border-box" }}
                />
                {createError && <div style={{ color: "#dc2626", fontSize: 12, marginBottom: 8 }}>⚠ {createError}</div>}
                <div style={{ display: "flex", gap: 8 }}>
                  <button onClick={() => { setShowCreate(false); setNewName(""); setCreateError(""); }} style={{ flex: 1, padding: "7px 0", borderRadius: 8, border: "1px solid var(--border)", background: "transparent", color: "var(--muted)", cursor: "pointer" }}>Cancel</button>
                  <button onClick={createClient} disabled={creating || !newName.trim()} style={{ flex: 2, padding: "7px 0", borderRadius: 8, background: "var(--accent)", color: "#fff", border: "none", fontWeight: 600, cursor: "pointer", opacity: (creating || !newName.trim()) ? 0.6 : 1 }}>
                    {creating ? "Creating..." : "Create Client →"}
                  </button>
                </div>
              </div>
            )}
          </div>

          <button
            onClick={() => setShowCreate(true)}
            style={{ background: "var(--accent)", color: "#fff", border: "none", borderRadius: 10, padding: "12px 28px", fontSize: 15, fontWeight: 700, cursor: "pointer", display: showCreate ? "none" : "inline-flex", alignItems: "center", gap: 8 }}
          >
            + New Client
          </button>
        </>
      )}
    </section>
  );
}

// ─── KNOWLEDGE INITIALIZER ───────────────────────────────────────────────────
function KnowledgeInitializer({ client, onInitialized, onBack }) {
  const [file, setFile] = useState(null);
  const [importing, setImporting] = useState(false);
  const [error, setError] = useState("");
  const [result, setResult] = useState(null);

  const handleFile = async (f) => {
    setFile(f); setError(""); setResult(null);
    setImporting(true);
    try {
      const data = await readFileAsObjects(f);
      
      const res = await fetch("/api/knowledge/import", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ client_id: client.id, authoritative_column: "purchase_account", rows: data.rows })
      });
      const resData = await res.json();
      if (!res.ok) throw new Error(resData.error || "Import failed");
      
      setResult(resData);
      
      // Auto-redirect after 1.2 seconds so user sees the stats
      setTimeout(async () => {
        const clientRes = await fetch(`/api/clients/${client.id}`);
        const clientData = await clientRes.json();
        if (clientData.client) onInitialized(clientData.client);
      }, 1200);
      
    } catch (e) { 
      setError(e.message); 
      setImporting(false);
    }
  };

  return (
    <section className="landing">
      <button onClick={onBack} style={{ background: "none", border: "1px solid var(--border)", borderRadius: 8, padding: "6px 14px", cursor: "pointer", color: "var(--muted)", marginBottom: 24, fontSize: 13 }}>← Back to Clients</button>

      <p className="eyebrow">ONE-TIME SETUP</p>
      <h2>Initialize Knowledge Base — {client.display_name}</h2>
      <p className="lead" style={{ maxWidth: 620 }}>
        Upload this client's historical Items Sheet to teach the system their existing accounting history.
        <strong> This is a one-time step.</strong> After this, the system will automatically classify items from this knowledge base during monthly audits.
      </p>

      <div style={{ background: "#fef9c3", border: "1px solid #fde68a", borderRadius: 10, padding: "14px 18px", marginBottom: 24, maxWidth: 620 }}>
        <strong>⚠️ Important:</strong> You will not be asked to upload this file again. The knowledge base will grow automatically as your team approves new classifications during monthly audits.
      </div>

      {!result ? (
        <>
          <div className="uploads">
            <Upload title="Historical Items Sheet" help={importing ? "Processing and importing item list..." : "Upload the client's Zoho Items export"} file={file} onChange={handleFile} />
          </div>
          {importing && (
            <div style={{ textAlign: "center", padding: "20px 0", color: "var(--muted)" }}>
              <span className="setup-spinner" style={{ display: "inline-block", marginRight: 8 }}>⚙️</span>
              Learning historical accounts...
            </div>
          )}
          {error && <p style={{ color: "#dc2626", marginTop: 12 }}>⚠ {error}</p>}
        </>
      ) : (
        <div style={{ background: "var(--surface)", border: "1.5px solid #bbf7d0", borderRadius: 12, padding: "28px 32px", maxWidth: 520, textAlign: "center" }}>
          <div style={{ fontSize: 48, marginBottom: 16 }}>✅</div>
          <h3 style={{ margin: "0 0 8px", color: "var(--ink)" }}>Knowledge Base Initialized</h3>
          <p style={{ color: "var(--muted)", marginBottom: 20 }}>The system has learned from this client's historical data.</p>
          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 12, textAlign: "center", marginBottom: 24 }}>
            <div style={{ background: "var(--bg)", borderRadius: 8, padding: "12px" }}>
              <div style={{ fontSize: 24, fontWeight: 700, color: "var(--ink)" }}>{result.inserted?.toLocaleString("en-IN")}</div>
              <div style={{ fontSize: 12, color: "var(--muted)" }}>Items Imported</div>
            </div>
            <div style={{ background: "var(--bg)", borderRadius: 8, padding: "12px" }}>
              <div style={{ fontSize: 24, fontWeight: 700, color: "var(--ink)" }}>{result.accounts_imported}</div>
              <div style={{ fontSize: 12, color: "var(--muted)" }}>Account Heads</div>
            </div>
            {result.conflicts > 0 && (
              <div style={{ background: "#fef3c7", borderRadius: 8, padding: "12px" }}>
                <div style={{ fontSize: 24, fontWeight: 700, color: "#92400e" }}>{result.conflicts}</div>
                <div style={{ fontSize: 12, color: "#92400e" }}>Conflicting Items</div>
              </div>
            )}
            {result.skipped > 0 && (
              <div style={{ background: "var(--bg)", borderRadius: 8, padding: "12px" }}>
                <div style={{ fontSize: 24, fontWeight: 700, color: "var(--muted)" }}>{result.skipped}</div>
                <div style={{ fontSize: 12, color: "var(--muted)" }}>Skipped (no name)</div>
              </div>
            )}
          </div>
          <p style={{ fontSize: 12, color: "var(--muted)" }}>You will not be asked to upload this file again. Redirecting...</p>
        </div>
      )}
    </section>
  );
}

// ─── KNOWLEDGE STATUS BANNER ─────────────────────────────────────────────────
function KnowledgeStatusBanner({ client, knowledgeCount, onSwitch }) {
  if (!client) return null;
  return (
    <div style={{ display: "flex", alignItems: "center", gap: 10, background: client.knowledge_initialized ? "#f0fdf4" : "#fef9c3", border: `1px solid ${client.knowledge_initialized ? "#bbf7d0" : "#fde68a"}`, borderRadius: 10, padding: "7px 14px", fontSize: 13 }}>
      <span>{client.knowledge_initialized ? "📚" : "⚠️"}</span>
      <div>
        <strong style={{ color: "var(--ink)" }}>{client.display_name}</strong>
        <span style={{ color: "var(--muted)", marginLeft: 8 }}>
          {client.knowledge_initialized ? `${knowledgeCount.toLocaleString("en-IN")} KB items` : "Knowledge Not Initialized"}
        </span>
      </div>
      <button onClick={onSwitch} style={{ background: "none", border: "1px solid var(--border)", borderRadius: 6, padding: "3px 10px", cursor: "pointer", color: "var(--muted)", fontSize: 12, marginLeft: 4 }}>
        Switch Client
      </button>
    </div>
  );
}

// ─── HOME COMPONENT ──────────────────────────────────────────────────────────
export default function Home() {
  // ── Client & Knowledge State ─────────────────────────────────────────────
  const [clients, setClients] = useState([]);
  const [selectedClient, setSelectedClient] = useState(null);
  const [clientsLoading, setClientsLoading] = useState(true);
  const [knowledgeItems, setKnowledgeItems] = useState([]);
  const [knowledgeLoading, setKnowledgeLoading] = useState(false);

  // ── Audit State ──────────────────────────────────────────────────────────
  const [current, setCurrent] = useState(null);
  const [previous, setPrevious] = useState(null);
  const [error, setError] = useState("");
  const [tab, setTab] = useState("Overview");
  const [thresholds, setThresholds] = useState({ vendor: 20, item: 25, price: 20 });
  const [showAiSetup, setShowAiSetup] = useState(false);
  const [showDataPreview, setShowDataPreview] = useState(false);

  // ── AI Config ────────────────────────────────────────────────────────────
  const [aiConfig, setAiConfig] = useState(() => {
    if (typeof window === "undefined") return null;
    try { const raw = localStorage.getItem("ai_config"); if (raw) return JSON.parse(raw); } catch (e) {}
    return null;
  });
  const [apiKey, setApiKey] = useState(() => {
    if (typeof window === "undefined") return "";
    return sessionStorage.getItem("groq_api_key") || "";
  });
  const groqModel = aiConfig?.model || "";

  // ── Load clients on mount ────────────────────────────────────────────────
  const loadClients = useCallback(async () => {
    setClientsLoading(true);
    try {
      const res = await fetch("/api/clients");
      const data = await res.json();
      setClients(data.clients || []);
    } catch (e) { console.error("Failed to load clients:", e); }
    finally { setClientsLoading(false); }
  }, []);

  useEffect(() => { loadClients(); }, [loadClients]);

  // ── Load knowledge when client selected ──────────────────────────────────
  useEffect(() => {
    if (!selectedClient?.knowledge_initialized) { setKnowledgeItems([]); return; }
    setKnowledgeLoading(true);
    fetch(`/api/knowledge?client_id=${selectedClient.id}`)
      .then(r => r.json())
      .then(d => setKnowledgeItems(d.knowledge || []))
      .catch(e => console.error("Failed to load knowledge:", e))
      .finally(() => setKnowledgeLoading(false));
  }, [selectedClient]);

  // ── Screen routing ───────────────────────────────────────────────────────
  const screen = !selectedClient
    ? "client_manager"
    : !selectedClient.knowledge_initialized
    ? "knowledge_init"
    : "audit";

  // ── Handlers ─────────────────────────────────────────────────────────────
  const saveAiConfig = (key, config) => {
    setApiKey(key); setAiConfig(config);
    if (typeof window !== "undefined") {
      sessionStorage.setItem("groq_api_key", key);
      localStorage.setItem("ai_config", JSON.stringify(config));
      localStorage.removeItem("groq_api_key");
      localStorage.removeItem("groq_model");
    }
  };

  const setT = (k, v) => setThresholds(t => ({ ...t, [k]: v }));

  const result = useMemo(
    () => current && previous ? analyse(current, previous, thresholds, knowledgeItems) : null,
    [current, previous, thresholds, knowledgeItems]
  );

  const upload = async (file, setter) => { try { setError(""); setter(await readFile(file)); } catch (e) { setError(e.message); } };
  const spend = x => x?.records.reduce((s, r) => s + r.total, 0) || 0;
  const total = result && (result.duplicates.length + result.vendors.length + result.items.length + result.prices.length + result.misclassifications.length);
  const sharedAccounts = useMemo(() => current ? [...new Set(current.records.map(r => r.account).filter(a => a && a !== "Unassigned Account"))] : [], [current]);

  const handleClientSwitch = () => {
    setSelectedClient(null);
    setCurrent(null);
    setPrevious(null);
    setKnowledgeItems([]);
    setTab("Overview");
  };

  const handleKbSave = (item) => {
    setKnowledgeItems(prev => {
      const norm = normalizeItemName(item.item);
      const exists = prev.findIndex(k => k.item_name_norm === norm);
      const newItem = { item_name_raw: item.item, item_name_norm: norm, account_head: item.account, source: "human_approved", verified: true, confidence: 100 };
      if (exists >= 0) { const next = [...prev]; next[exists] = newItem; return next; }
      return [...prev, newItem];
    });
  };

  return (
    <main>
      <header>
        <div>
          <p className="eyebrow">FINANCE CONTROL CENTER</p>
          <h1>P&L Audit Desk</h1>
        </div>
        <div style={{ display: "flex", alignItems: "center", gap: 12, flexWrap: "wrap" }}>
          {selectedClient && (
            <KnowledgeStatusBanner
              client={selectedClient}
              knowledgeCount={knowledgeItems.length}
              onSwitch={handleClientSwitch}
            />
          )}
          {screen === "audit" && (
            <button
              className={`header-ai-btn ${aiConfig?.verified && apiKey ? "configured" : ""}`}
              onClick={() => setShowAiSetup(true)}
            >
              {aiConfig?.verified && apiKey
                ? `⚙️ AI: ${groqModel.replace("llama-", "").replace("-versatile", "").replace("-instant", "")}`
                : aiConfig?.verified && !apiKey
                ? `⚙️ AI: Re-enter Key ⚠️`
                : `⚙️ AI: Setup`}
            </button>
          )}
          <p className="privacy"><i /> Local analysis — files stay on your device</p>
        </div>
      </header>

      {/* ── Screen 1: No client selected ── */}
      {screen === "client_manager" && (
        <ClientManagerScreen
          clients={clients}
          loading={clientsLoading}
          onSelect={setSelectedClient}
          onCreated={(client) => {
            setClients(prev => [client, ...prev]);
            setSelectedClient(client);
          }}
        />
      )}

      {/* ── Screen 2: Client selected but KB not initialized ── */}
      {screen === "knowledge_init" && (
        <KnowledgeInitializer
          client={selectedClient}
          onInitialized={(updatedClient) => {
            setSelectedClient(updatedClient);
            setClients(prev => prev.map(c => c.id === updatedClient.id ? updatedClient : c));
          }}
          onBack={handleClientSwitch}
        />
      )}

      {/* ── Screen 3: Audit workspace ── */}
      {screen === "audit" && (
        <>
          {knowledgeLoading && (
            <div style={{ background: "#eff6ff", border: "1px solid #bfdbfe", borderRadius: 8, padding: "8px 16px", marginBottom: 8, fontSize: 13, color: "#1e40af" }}>
              📚 Loading Knowledge Base for {selectedClient.display_name}...
            </div>
          )}

          {!result ? (
            <section className="landing">
              <p className="eyebrow">ZOHO BOOKS PURCHASE &amp; P&L REVIEW</p>
              <h2>Find what the spreadsheet misses.</h2>
              <p className="lead">Compare two Zoho purchase exports to flag duplicate bills, account misclassifications, vendor movements, item variation, price exceptions, and explore your Account Head hierarchy.</p>

              {knowledgeItems.length > 0 && (
                <div style={{ background: "#eff6ff", border: "1px solid #bfdbfe", borderRadius: 8, padding: "10px 16px", marginBottom: 20, fontSize: 13, color: "#1e40af", display: "inline-flex", alignItems: "center", gap: 8 }}>
                  📚 <strong>{knowledgeItems.length.toLocaleString("en-IN")} knowledge items</strong> loaded — items will be classified from Knowledge Base first
                </div>
              )}

              <div className="chips">
                <b>Knowledge-first classification</b>
                <b>Misclassification alerts</b>
                <b>Account heads pivot</b>
                <b>Duplicate bills</b>
                <b>Vendor variation</b>
                <b>Item variation</b>
                <b>Price exceptions</b>
              </div>
              <div className="uploads">
                <Upload title="Current month" help="Upload the latest Zoho export" file={current} onChange={f => upload(f, setCurrent)} />
                <Upload title="Previous month" help="Upload the month to compare" file={previous} onChange={f => upload(f, setPrevious)} />
              </div>
              {error && <p className="error">{error}</p>}
              <p className="note">Required: Bill Date and Vendor Name. Recommended: Account Name, Bill Number, Item Name, Quantity, Rate, Branch, and Item Total for full audit.</p>
            </section>
          ) : (
            <>
              <section className="run">
                <div>
                  <strong>Analysis ready</strong>
                  <small>{current.name} vs {previous.name}</small>
                </div>
                <div style={{ display: "flex", gap: 8, alignItems: "center" }}>
                  <button className="ai-setup-run-btn" onClick={() => setShowAiSetup(true)}>⚙️ {groqModel ? `AI: ${groqModel}` : "Configure AI"}</button>
                  <button className="verify-data-btn" onClick={() => setShowDataPreview(true)}>🔎 Verify Data</button>
                  <button onClick={() => { setCurrent(null); setPrevious(null); setTab("Overview"); setShowDataPreview(false); }}>New review</button>
                </div>
              </section>

              <section className="thresholds">
                <p className="eyebrow">SENSITIVITY THRESHOLDS</p>
                <div className="thresh-row">
                  <ThresholdInput label="Vendor variation" value={thresholds.vendor} onChange={v => setT("vendor", v)} />
                  <ThresholdInput label="Item variation" value={thresholds.item} onChange={v => setT("item", v)} />
                  <ThresholdInput label="Price exception" value={thresholds.price} onChange={v => setT("price", v)} maxVal={1000} />
                  <button className="thresh-reset" onClick={() => setThresholds({ vendor: 20, item: 25, price: 20 })}>Reset to defaults</button>
                </div>
              </section>

              <nav>
                {["Overview", "Misclassifications", "Account heads", "Duplicate bills", "Vendor variation", "Purchase variation", "Price exceptions"].map(x => (
                  <button className={tab === x ? "active" : ""} onClick={() => setTab(x)} key={x}>
                    {x}
                    {x === "Misclassifications" && result.misclassifications.length > 0 && (
                      <span className="nav-badge-count">{result.misclassifications.length}</span>
                    )}
                  </button>
                ))}
              </nav>

              {tab === "Overview" && (
                <>
                  <section className="metrics">
                    <Card label="Audit findings" value={total} warm />
                    <Card label="Current-month spend" value={show(spend(current))} />
                    <Card label="Spend movement" value={show(spend(current) - spend(previous))} />
                    <Card label="Rows reviewed" value={current.records.length.toLocaleString("en-IN")} />
                  </section>
                  <section className="panel">
                    <div className="panelhead">
                      <div><p className="eyebrow">PRIORITY QUEUE</p><h2>What to review first</h2></div>
                      <b className="badge">{result.misclassifications.length + result.duplicates.length} critical issues</b>
                    </div>
                    {result.misclassifications.length || result.duplicates.length || result.vendors.length || result.prices.length ? (
                      <div className="queue">
                        {[
                          ...result.misclassifications.slice(0, 3).map(x => ({ title: `Wrong Account: ${x.item}`, detail: `Booked in "${x.actualAccount}" → Should be "${x.suggestedAccount}"`, value: x.total, red: true })),
                          ...result.duplicates.slice(0, 3).map(x => ({ title: x.kind, detail: `${x.rows[0].vendor} - ${x.rows.length} matching lines`, value: x.total, red: x.risk === "Critical" })),
                          ...result.vendors.slice(0, 2).map(x => ({ title: `${x.status} vendor`, detail: x.label, value: x.diff })),
                          ...result.prices.slice(0, 2).map(x => ({ title: "High item price", detail: `${x.item} - ${x.vendor}`, value: x.total }))
                        ].slice(0, 8).map((x, i) => (
                          <div className="queueitem" key={i}>
                            <i className={x.red ? "red" : "amber"} />
                            <div><strong>{x.title}</strong><small>{x.detail}</small></div>
                            <b>{show(x.value)}</b>
                          </div>
                        ))}
                      </div>
                    ) : <Empty>No material flags found.</Empty>}
                  </section>
                </>
              )}

              {tab === "Misclassifications" && (
                <MisclassificationsView
                  items={result.misclassifications}
                  current={current}
                  onGoToPivot={() => setTab("Account heads")}
                  sharedApiKey={apiKey}
                  sharedModel={groqModel}
                  onOpenSetup={() => setShowAiSetup(true)}
                  sharedAccounts={sharedAccounts}
                  clientId={selectedClient?.id}
                  onSaveToKnowledge={handleKbSave}
                />
              )}

              {tab === "Account heads" && <AccountPivot current={current} />}

              {tab === "Duplicate bills" && (
                <section className="panel">
                  <Panel title="Duplicate bill patterns" badge={`${result.duplicates.length} findings`} />
                  <Table head={["Classification", "Vendor", "Item", "Matching lines", "Exposure"]}>
                    {result.duplicates.length ? result.duplicates.map((x, i) => (
                      <Fragment key={i}>
                        <tr>
                          <td><b className={x.risk === "Critical" ? "pill critical" : "pill"}>{x.kind}</b></td>
                          <td>{x.rows[0].vendor}</td>
                          <td>{x.rows[0].item || "-"}</td>
                          <td>{x.rows.length}</td>
                          <td>{show(x.total)}</td>
                        </tr>
                        <tr>
                          <td colSpan="5" style={{ padding: "0 0 10px 0", background: "var(--surface)" }}>
                            <table style={{ width: "100%", borderCollapse: "collapse", fontSize: "0.78rem" }}>
                              <thead>
                                <tr style={{ background: "var(--border)" }}>
                                  <th style={{ padding: "5px 14px", textAlign: "left", fontWeight: 600, color: "var(--muted)", letterSpacing: "0.05em" }}>Bill No.</th>
                                  <th style={{ padding: "5px 14px", textAlign: "left", fontWeight: 600, color: "var(--muted)", letterSpacing: "0.05em" }}>Date</th>
                                  <th style={{ padding: "5px 14px", textAlign: "right", fontWeight: 600, color: "var(--muted)", letterSpacing: "0.05em" }}>Amount</th>
                                </tr>
                              </thead>
                              <tbody>
                                {x.rows.map((r, j) => (
                                  <tr key={j} style={{ borderBottom: "1px solid var(--border)" }}>
                                    <td style={{ padding: "5px 14px", color: "var(--text)" }}>{r.bill || "—"}</td>
                                    <td style={{ padding: "5px 14px", color: "var(--text)" }}>{r.date || "—"}</td>
                                    <td style={{ padding: "5px 14px", textAlign: "right", color: x.risk === "Critical" ? "var(--red, #c0392b)" : "var(--text)", fontWeight: 500 }}>{show(r.total)}</td>
                                  </tr>
                                ))}
                              </tbody>
                            </table>
                          </td>
                        </tr>
                      </Fragment>
                    )) : <tr><td colSpan="5"><Empty>No duplicate patterns found.</Empty></td></tr>}
                  </Table>
                  {result.duplicates.length > 0 && <CopyDuplicatesButton duplicates={result.duplicates} />}
                </section>
              )}

              {tab === "Vendor variation" && <Changes title="Vendor variation" rows={result.vendors} field="Vendor" threshold={thresholds.vendor} />}
              {tab === "Purchase variation" && <Changes title="Purchase variation" rows={result.items} field="Item" threshold={thresholds.item} />}

              {tab === "Price exceptions" && (
                <section className="panel">
                  <Panel title="Items purchased above weighted average" badge={`${thresholds.price}%+ above average`} />
                  <Table head={["Item", "Vendor", "Rate paid", "Weighted average", "Variance"]}>
                    {result.prices.length ? result.prices.map((x, i) => (
                      <tr key={i}>
                        <td>{x.item}</td><td>{x.vendor}</td><td>{show(x.rate)}</td><td>{show(x.avg)}</td>
                        <td className="bad">+{(x.pct * 100).toFixed(0)}%</td>
                      </tr>
                    )) : <tr><td colSpan="5"><Empty>No price exceptions found.</Empty></td></tr>}
                  </Table>
                </section>
              )}
            </>
          )}

          {/* Data preview modal */}
          {current && showDataPreview && (
            <div style={{ position: "fixed", inset: 0, background: "rgba(0,0,0,0.5)", zIndex: 200, overflowY: "auto", padding: "40px 16px" }}>
              <div style={{ maxWidth: 1100, margin: "0 auto", background: "var(--surface)", borderRadius: 12, overflow: "hidden" }}>
                <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", padding: "16px 20px", borderBottom: "1px solid var(--border)" }}>
                  <strong>Data Verification Preview</strong>
                  <button className="pivot-btn" onClick={() => setShowDataPreview(false)}>Close ✕</button>
                </div>
                <DataPreview data={current} />
              </div>
            </div>
          )}

          <AiSetupModal isOpen={showAiSetup} onClose={() => setShowAiSetup(false)} apiKey={apiKey} aiConfig={aiConfig} onSave={saveAiConfig} />
          {current && <AiChatWidget availableAccounts={sharedAccounts} apiKey={apiKey} model={groqModel} onOpenSetup={() => setShowAiSetup(true)} />}
        </>
      )}
    </main>
  );
}

function Card({ label, value, warm }) { return <article className={warm ? "card warm" : "card"}><small>{label}</small><strong>{value}</strong></article>; }
function Panel({ title, badge }) { return <div className="panelhead"><div><p className="eyebrow">AUDIT REVIEW</p><h2>{title}</h2></div><b className="badge">{badge}</b></div>; }
function Changes({ title, rows, field, threshold }) {
  return (
    <section className="panel">
      <Panel title={title} badge={`${threshold}%+ change or new entry`} />
      <Table head={[field, "Current month", "Previous month", "Change", "Status"]}>
        {rows.length ? rows.map((x, i) => (
          <tr key={i}>
            <td>{x.label}</td><td>{show(x.total)}</td><td>{show(x.old)}</td>
            <td className={x.diff >= 0 ? "bad" : "good"}>{x.diff >= 0 ? "+" : ""}{show(x.diff)}</td>
            <td><b className="pill">{x.status}</b></td>
          </tr>
        )) : <tr><td colSpan="5"><Empty>No material movements found.</Empty></td></tr>}
      </Table>
    </section>
  );
}

function CopyDuplicatesButton({ duplicates }) {
  const [copied, setCopied] = useState(false);

  function buildText() {
    const colW = [12, 20, 22, 26, 16, 14];
    const pad = (s, w) => String(s ?? "—").padEnd(w).slice(0, w);
    const lines = [];
    lines.push("DUPLICATE BILL REPORT");
    lines.push("=".repeat(90));
    lines.push("");

    duplicates.forEach((x, idx) => {
      lines.push(`${idx + 1}. ${x.kind.toUpperCase()}   (${x.rows.length} matching lines  |  Total Exposure: ${show(x.total)})`);
      lines.push("-".repeat(90));
      const header = pad("Date", colW[0]) + pad("Bill No.", colW[1]) + pad("Vendor Name", colW[2]) + pad("Item", colW[3]) + pad("Qty", colW[4]) + pad("Amount", colW[5]);
      lines.push(header);
      lines.push("-".repeat(90));
      x.rows.forEach(r => {
        lines.push(
          pad(r.date || "—", colW[0]) +
          pad(r.bill || "—", colW[1]) +
          pad(r.vendor || "—", colW[2]) +
          pad(r.item || "—", colW[3]) +
          pad(r.qty != null && r.qty !== 0 ? r.qty : "—", colW[4]) +
          pad(show(r.total), colW[5])
        );
      });
      lines.push("");
    });

    lines.push("=".repeat(90));
    lines.push(`Generated on ${new Date().toLocaleString("en-IN")}`);
    return lines.join("\n");
  }

  function handleCopy() {
    navigator.clipboard.writeText(buildText()).then(() => {
      setCopied(true);
      setTimeout(() => setCopied(false), 2500);
    });
  }

  return (
    <div style={{ display: "flex", justifyContent: "flex-end", padding: "12px 0 2px 0" }}>
      <button
        onClick={handleCopy}
        style={{
          display: "inline-flex", alignItems: "center", gap: 7,
          padding: "8px 18px", borderRadius: 8, border: "1px solid var(--border)",
          background: copied ? "var(--green, #27ae60)" : "var(--surface)",
          color: copied ? "#fff" : "var(--text)",
          fontWeight: 600, fontSize: "0.82rem", cursor: "pointer",
          transition: "background 0.25s, color 0.25s",
        }}
      >
        {copied ? (
          <><span>✓</span> Copied!</>
        ) : (
          <><span style={{ fontSize: "1rem" }}>📋</span> Copy for Accountant</>
        )}
      </button>
    </div>
  );
}
