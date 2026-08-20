"use client";
import { useEffect, useMemo, useState } from "react";
import * as XLSX from "xlsx";

const inr = new Intl.NumberFormat("en-IN", { style: "currency", currency: "INR", maximumFractionDigits: 0 });
const clean = (v) => String(v ?? "").trim().toLowerCase().replace(/[^a-z0-9]/g, "");
const n = (v) => typeof v === "number" ? v : Number(String(v ?? "").replace(/[,₹]/g, "")) || 0;
const show = (v) => inr.format(v || 0);
const col = (heads, choices) => choices.map(clean).map(x => heads.map(clean).indexOf(x)).find(x => x >= 0) ?? -1;

// ── TAXONOMY: Only specific brands + unambiguous multi-word phrases ──────────
// RULE: No generic single words (oil, slice, wheat, apple, etc.) — they cause false positives.
// Single words are only allowed when they are UNAMBIGUOUS in a restaurant context (e.g. "prawns", "paneer", "mutton").
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
      "gin", "bombay sapphire", "beefeater", "tanqueray", "gordons", "greater than gin", "hendricks",
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
      "squid", "calamari", "octopus", "crab", "mud crab", "lobster", "clams", "oysters",
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
      "broiler chicken", "country chicken", "desi chicken", "whole chicken", "chicken",
      "mutton keema", "mutton curry cut", "mutton chops", "mutton boti", "mutton brain",
      "lamb chops", "lamb shank", "goat meat", "mutton", "lamb",
      "brown eggs", "white eggs", "quail eggs", "egg tray", "eggs", "egg",
      "beef", "pork", "bacon", "ham", "pepperoni"
    ]
  },
  {
    id: "beverages",
    label: "Beverages",
    aliases: ["beverage", "soft drink", "drinks", "cold drink", "non alcoholic"],
    keywords: [
      "real apple juice", "real mango juice", "real orange juice", "real cranberry", "real pineapple",
      "real litchi", "real juice", "tropicana", "minute maid", "raw pressery", "paper boat",
      "red bull", "redbull", "monster energy", "sting energy",
      "tonic water", "schweppes tonic", "ginger ale", "club soda", "soda water", "lehar soda", "kinley soda",
      "diet coke", "coke zero", "coca cola", "thums up", "limca", "sprite", "fanta", "mirinda", "mountain dew", "pepsi",
      "bisleri", "kinley water", "aquafina", "vedica", "himalayan water", "packaged water", "mineral water",
      "monin", "malas", "malass",
      "watermelon syrup", "strawberry crush", "blueberry crush", "kiwi crush", "litchi crush",
      "orange crush", "pineapple crush", "mojito syrup", "grenadine", "blue curacao",
      "fruit syrup", "mocktail syrup",
      "frooti", "maaza", "appy fizz"
    ]
  },
  {
    id: "groceries",
    label: "Groceries",
    aliases: ["grocer", "food and grocery", "food & grocery", "food purchase", "provision", "dry good", "raw material", "staple", "spices", "grain", "ingredients"],
    keywords: [
      "sona masoori", "basmati rice", "kolam rice", "ponni rice", "rice", "basmati", "kolam",
      "wheat flour", "atta", "maida", "sooji", "semolina", "besan", "cornflour", "corn flour",
      "cornstarch", "custard powder", "bread crumbs", "breadcrumbs",
      "toor dal", "tur dal", "moong dal", "mung dal", "urad dal", "chana dal",
      "kabuli chana", "rajma", "soya chunks", "poha", "vermicelli", "sevai", "noodles", "pasta", "macaroni", "spaghetti", "croutons", "dal", "lentil", "lentils",
      "garam masala", "chaat masala", "biryani masala", "kitchen king masala", "pav bhaji masala", "sambar powder", "rasam powder", "coriander powder", "chilli powder", "turmeric powder", "mustard seed", "fenugreek seed", "methi seed", "bay leaf", "tej patta", "black pepper", "kali mirch", "white pepper", "kasuri methi", "red chilli powder", "kashmiri chilli", "degi mirch", "chili flakes",
      "spices", "masala", "haldi", "turmeric", "jeera", "cumin", "dhania", "rai", "sarson", "saunf", "fennel", "cardamom", "elaichi", "laung", "cinnamon", "dalchini", "star anise", "nutmeg", "jaiphal", "saffron", "kesar", "ajwain", "kalonji", "hing", "asafoetida", "oregano", "thyme", "rosemary", "paprika",
      "sunflower oil", "mustard oil", "groundnut oil", "peanut oil", "sesame oil", "til oil", "olive oil", "canola oil", "soybean oil", "palm oil", "vanaspati", "dalda", "refined oil", "cooking oil",
      "desi ghee", "cow ghee", "buffalo ghee", "amul ghee", "ghee",
      "brown sugar", "jaggery", "gur", "honey", "rock salt", "black salt", "sendha namak", "pink salt", "baking soda", "baking powder", "yeast", "citric acid", "ajinomoto", "msg", "sugar", "salt",
      "tomato ketchup", "red chilli sauce", "green chilli sauce", "soya sauce", "dark soy", "white vinegar", "apple cider vinegar", "sriracha", "tabasco", "schezwan sauce", "mayonnaise", "mayo", "mustard paste", "salsa", "peri peri sauce", "ketchup", "vinegar",
      "pickle", "achaar", "murabba", "chutney", "papad", "appalam", "tamarind", "imli", "desiccated coconut", "black olive", "green olive", "stuffed olive", "olive slice",
      "cashew", "kaju", "badam", "almond", "kismis", "raisin", "pista", "pistachio", "walnut", "akhrot", "dates", "khajoor", "melon seeds", "poppy seeds", "khus khus",
      "cocoa powder", "cooking chocolate", "chocolate chips", "vanilla essence", "vanilla extract", "vanilla 4ltr", "food color", "food colouring",
      "knorr chicken broth", "chicken broth", "chicken powder", "knorr", "broth powder", "bouillon", "seasoning powder"
    ]
  },
  {
    id: "dairy",
    label: "Dairy",
    aliases: ["dairy", "milk product"],
    keywords: [
      "toned milk", "full cream milk", "cow milk", "buffalo milk", "cottage cheese", "hung curd", "fresh cream", "amul cream", "sour cream", "whipped cream", "salted butter", "unsalted butter", "table butter", "mozzarella", "cheddar cheese", "cheese slice", "cheese block", "cream cheese",
      "milk", "paneer", "curd", "dahi", "yogurt", "yoghurt", "cream", "butter", "cheese", "khoya", "mawa", "buttermilk", "chaas", "lassi"
    ]
  },
  {
    id: "other_purchases",
    label: "Other Purchases",
    aliases: ["other purchase", "other purchases", "misc", "miscellaneous", "general purchase", "other expense"],
    keywords: [
      "wood charcoal", "charcoal", "coal",
      "ice slab", "ice slabs", "ice cube", "ice cubes", "crushed ice", "dry ice",
      "wooden skewers", "bamboo skewers", "skewers", "toothpick", "toothpicks", "birthday candles", "matchbox"
    ]
  },
  {
    id: "kitchen_tools",
    label: "Kitchen tools",
    aliases: ["kitchen tool", "kitchen tools", "utensil", "utensils", "crockery", "cutlery", "hotelware", "equipment", "bar tool"],
    keywords: [
      "dal katori", "katori", "dip bowl", "soup bowl", "dip bowl round", "mixing bowl", "salad bowl",
      "kadai", "fry pan", "sauce pan", "tawa", "dosa tawa", "pressure cooker", "patila", "strainer", "colander", "ladle", "karchi", "chef knife", "chopping knife", "peeler", "grater", "chopping board", "cutting board", "tongs", "chimta", "whisk", "rolling pin", "belan", "chakla", "baking tray", "sizzler plate",
      "plate ceramic", "bowl ceramic", "crockery", "hotelware", "glassware", "arcoroc", "pilsner glass", "pilsner", "beer glass", "wine glass", "shot glass", "whisky glass", "water glass",
      "cocktail shaker", "bar strainer", "jigger", "peg measurer", "muddler", "bar spoon", "pourer", "corkscrew", "bottle opener"
    ]
  },
  {
    id: "cleaning",
    label: "Cleaning and housekeeping",
    aliases: ["clean", "housekeep", "sanit", "hygiene", "detergent", "soap"],
    keywords: [
      "dishwash bar", "dishwash liquid", "dishwash", "vim bar", "vim liquid", "vim", "exo", "pril", "surf excel", "surf", "ariel", "tide", "rin", "detergent powder", "liquid detergent", "detergent",
      "soap oil", "soap oil thick", "liquid soap", "bar soap",
      "floor cleaner", "lizol", "phenyl", "white phenyl", "colin", "glass cleaner", "harpic", "toilet cleaner", "bathroom cleaner", "drain cleaner", "caustic soda",
      "bleaching powder", "bleach", "disinfectant", "hand sanitizer", "sanitizer", "hand wash", "lifebuoy", "dettol", "savlon",
      "phool jhadu", "coconut broom", "broom", "jhadu", "mop", "floor wiper", "wiper", "duster", "cleaning cloth", "microfiber cloth", "sponge", "steel scrubber", "green scrubber", "scrubber", "scotch brite", "garbage bag", "trash bag", "dustbin cover", "dust pan", "rubber gloves", "room freshener", "odonil"
    ]
  },
  {
    id: "packaging",
    label: "Packaging & Disposables",
    aliases: ["packag", "packing", "pack material", "disposab", "takeaway", "parcel"],
    keywords: [
      "meal tray", "meal box", "500ml container", "750ml container", "1000ml container", "aluminium container", "foil container", "burger box", "pizza box", "cake box", "sweet box", "plastic container", "food container",
      "kraft paper bag", "non woven bag", "d-cut bag", "paper bag", "carry bag", "zip lock", "polythene bag",
      "paper plate", "disposable plate", "paper cup", "plastic glass", "disposable glass",
      "paper napkin", "cocktail napkin", "tissue paper", "tissue roll", "kitchen roll", "toilet roll", "tissue", "napkin",
      "aluminium foil", "silver foil", "cling wrap", "cling film", "butter paper", "parchment paper",
      "paper straw", "plastic straw", "straw", "wooden spoon", "plastic spoon", "wooden fork", "plastic fork", "disposable cutlery", "chopstick", "chopsticks"
    ]
  },
  {
    id: "vegetables",
    label: "Fresh Vegetables",
    aliases: ["vegetable", "fresh veg", "veggie", "greens", "sabzi", "tarkari"],
    keywords: [
      "thai red chilli", "thai chilli", "green chilli", "hari mirch", "bell pepper", "shimla mirch", "capsicum", "green peas", "matar", "coriander leaves", "fresh dhania", "mint leaves", "pudina", "curry leaves", "kadi patta", "spring onion", "button mushroom", "baby corn", "sweet corn", "lady finger", "bottle gourd", "bitter gourd", "raw banana",
      "onion", "pyaz", "potato", "aloo", "tomato", "tamatar", "ginger", "adrak", "garlic", "lahsun", "carrot", "gajar", "beans", "cabbage", "cauliflower", "broccoli", "spinach", "palak", "methi", "lettuce", "iceberg", "cucumber", "kheera", "beetroot", "radish", "mooli", "leek", "celery", "zucchini", "mushroom", "bhindi", "okra", "brinjal", "baingan", "eggplant", "lauki", "karela", "pumpkin", "kaddu", "lemon", "nimbu", "drumstick"
    ]
  },
  {
    id: "fruits",
    label: "Fresh Fruits",
    aliases: ["fruit", "fresh fruit"],
    keywords: [
      "sweet lime", "fresh mango", "blueberry fresh", "fresh strawberry",
      "fresh apple", "seb", "banana", "kela", "fresh orange", "santra", "mosambi", "pomegranate", "anar", "fresh watermelon", "tarbooj", "muskmelon", "kharbuj", "papaya", "papita", "fresh pineapple", "ananas", "aam", "grapes", "angoor", "kiwi", "guava", "amrood", "pear", "chikoo", "dragonfruit", "plum", "peach", "cherry"
    ]
  },
  {
    id: "stationery",
    label: "Stationery & Office",
    aliases: ["station", "office supplies", "printing", "paper"],
    keywords: [
      "attendance register", "bill book", "kot book", "receipt book", "permanent marker", "ball pen", "gel pen", "pos roll", "billing roll", "thermal roll", "printer cartridge", "toner cartridge",
      "register", "notebook", "pencil", "marker", "stapler", "stapler pin", "punch machine", "brown tape", "cello tape", "scissor", "scissors", "stamp pad", "rubber band", "binder clip", "envelope", "a4 paper", "toner"
    ]
  }
];

function cleanWords(str) {
  return String(str || "").toLowerCase().replace(/[^a-z0-9\s]/g, " ").replace(/\s+/g, " ").trim();
}

// Check if actual account string matches category aliases or related categories
function accountMatchesCategory(actualAccount, category) {
  const normAcc = String(actualAccount || "").toLowerCase();
  
  // Direct alias match
  if (category.aliases.some(alias => normAcc.includes(alias))) return true;

  // Cross-category allowances:
  // e.g. Packing materials vs Packaging & Disposables (same category)
  if (category.id === "packaging" && (normAcc.includes("pack") || normAcc.includes("clean") || normAcc.includes("housekeep"))) return true;
  if (category.id === "cleaning" && (normAcc.includes("pack") || normAcc.includes("soap"))) return true;

  // e.g. Kitchen utensils in cutlery, crockery, hotelware, equipment or kitchen
  if (category.id === "kitchen_tools" && (normAcc.includes("hotelware") || normAcc.includes("equipment") || normAcc.includes("kitchen") || normAcc.includes("crockery") || normAcc.includes("cutlery"))) return true;

  // e.g. Vegetables & Fruits under Food / Groceries / Combined accounts
  if ((category.id === "vegetables" || category.id === "fruits") && (normAcc.includes("grocer") || normAcc.includes("food") || normAcc.includes("vegetable") || normAcc.includes("fruit") || normAcc.includes("tarkari"))) return true;
  if (category.id === "groceries" && (normAcc.includes("food") || normAcc.includes("provision") || normAcc.includes("raw material"))) return true;

  return false;
}

// Veto rules: if item contains these signals, ignore certain category matches
const VETO_RULES = [
  // Alcohol keywords block cigarette classification
  { blockedCategoryId: "cigarettes", ifItemContains: ["rum", "vodka", "whisky", "whiskey", "gin", "tequila", "brandy", "wine", "beer", "scotch", "bourbon", "liqueur", "bacardi", "smirnoff", "absolut"] },
  // "soap" or "detergent" in item name blocks grocery oil match
  { blockedCategoryId: "groceries", ifItemContains: ["soap", "detergent", "cleaner", "liquid soap", "dishwash"] },
  // "monin" or "syrup" blocks fruit classification
  { blockedCategoryId: "fruits", ifItemContains: ["monin", "syrup", "crush", "malas", "cordial"] },
  { blockedCategoryId: "vegetables", ifItemContains: ["monin", "syrup", "crush", "malas", "cordial"] },
  // "broth", "powder", "seasoning", "cube", "bouillon", "knorr" in item blocks poultry (it's a grocery seasoning)
  { blockedCategoryId: "poultry", ifItemContains: ["broth", "powder", "seasoning", "cube", "bouillon", "knorr"] },
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

// Find expected category for an item description using strict word-boundary matching
function classifyItem(itemName) {
  const norm = cleanWords(itemName);
  if (!norm) return null;

  let bestMatch = null;
  let maxKeywordLen = 0;

  for (const cat of RESTAURANT_TAXONOMY) {
    // Skip this category if a veto rule blocks it for this item
    if (isVetoed(cat.id, norm)) continue;

    for (const kw of cat.keywords) {
      const normKw = cleanWords(kw);
      if (!normKw) continue;

      // Strict matching: keyword must match as complete whole words
      // e.g. "classic bt" will NOT match "Bacardi Classic" since "classic" alone is no longer a keyword
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

// Auditor: Detect items booked under wrong account heads
function detectMisclassifications(records) {
  if (!records || !records.length) return [];
  
  // Find all distinct account names in this sheet to suggest the exact sheet account name when possible
  const sheetAccounts = [...new Set(records.map(r => r.account).filter(a => a && a !== "Unassigned Account"))];
  const findBestSheetAccountName = (category) => {
    const matched = sheetAccounts.find(acc => category.aliases.some(alias => acc.toLowerCase().includes(alias)));
    return matched || category.label;
  };

  const map = new Map();

  records.forEach(r => {
    if (!r.item || !r.account || r.account === "Unassigned Account") return;
    const match = classifyItem(r.item);
    if (!match) return; // Unclassified items are not flagged

    const expectedCat = match.category;
    // Check if actual account head matches expected category
    const isCorrect = accountMatchesCategory(r.account, expectedCat);

    if (!isCorrect) {
      const suggestedName = findBestSheetAccountName(expectedCat);
      const key = `${r.item}:::${r.vendor}:::${r.account}:::${expectedCat.id}`;
      if (!map.has(key)) {
        map.set(key, {
          item: r.item,
          vendor: r.vendor,
          actualAccount: r.account,
          suggestedAccount: suggestedName,
          matchedKeyword: match.keyword,
          confidence: "High",
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

function groups(rows, key) { return rows.reduce((m, r) => { const k = key(r); m.set(k, [...(m.get(k) || []), r]); return m; }, new Map()); }
function rollup(rows, field) { return [...groups(rows, r => clean(r[field]) || "unassigned")].map(([key, list]) => ({ key, label: list[0][field] || "Unassigned", total: list.reduce((s, r) => s + r.total, 0) })); }

function analyse(current, previous, thresholds) {
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

  const misclassifications = detectMisclassifications(current.records);

  return {
    duplicates,
    vendors: compare("vendor", vendorCut),
    items: compare("item", itemCut),
    prices: prices.sort((a, b) => b.pct - a.pct),
    misclassifications
  };
}

function Upload({ title, file, onChange, help }) {
  return <label className="upload"><input type="file" accept=".xlsx,.xls,.csv" onChange={e => e.target.files[0] && onChange(e.target.files[0])} /><span>↑</span><strong>{title}</strong><small>{file?.name || help}</small><em>{file ? "Replace file" : "Choose Excel or CSV"}</em></label>;
}
function Empty({ children }) { return <p className="empty">{children}</p>; }
function Table({ head, children }) { return <div className="table"><table><thead><tr>{head.map(x => <th key={x}>{x}</th>)}</tr></thead><tbody>{children}</tbody></table></div>; }
function ThresholdInput({ label, value, onChange, maxVal = 100 }) {
  return <label className="thresh-label">{label}<div className="thresh-wrap"><input className="thresh-input" type="number" min="0" max={maxVal} value={value} onChange={e => onChange(Math.max(0, Math.min(maxVal, e.target.value === "" ? 0 : Number(e.target.value) || 0)))} /><span className="thresh-pct">%</span></div></label>;
}

function MisclassificationsView({ items, current, onGoToPivot, sharedApiKey, sharedModel, onOpenSetup, sharedAccounts }) {
  const [q, setQ] = useState("");
  const [rowAiState, setRowAiState] = useState({}); // { [itemKey]: { loading, result, error, countdown } }
  const [aiResults, setAiResults] = useState([]);

  const askAiForRow = async (x) => {
    const rowKey = `${x.item}:::${x.vendor}`;
    if (!sharedApiKey || !sharedModel) {
      onOpenSetup();
      return;
    }
    setRowAiState(prev => ({ ...prev, [rowKey]: { loading: true, result: null, error: null, countdown: 0 } }));
    try {
      const res = await fetch("/api/ai-audit", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          singleItem: { item: x.item, vendor: x.vendor, actualAccount: x.actualAccount },
          availableAccounts: sharedAccounts,
          apiKey: sharedApiKey,
          model: sharedModel
        })
      });
      const data = await res.json();

      if (res.status === 429) {
        const secs = data.retryAfter || 30;
        setRowAiState(prev => ({
          ...prev,
          [rowKey]: { loading: false, result: null, error: `Rate limited. Retry in ${secs}s`, countdown: secs }
        }));
        // Countdown timer
        let rem = secs;
        const iv = setInterval(() => {
          rem -= 1;
          if (rem <= 0) {
            clearInterval(iv);
            setRowAiState(prev => ({
              ...prev,
              [rowKey]: { loading: false, result: null, error: null, countdown: 0 }
            }));
          } else {
            setRowAiState(prev => ({
              ...prev,
              [rowKey]: { ...prev[rowKey], countdown: rem, error: `Retry in ${rem}s` }
            }));
          }
        }, 1000);
        return;
      }

      if (!res.ok || data.error) throw new Error(data.error || "AI error");

      const resObj = data.result || {};
      const isMis = resObj.ok === false;
      setRowAiState(prev => ({
        ...prev,
        [rowKey]: {
          loading: false,
          result: { isMisclassified: isMis, suggestedAccount: resObj.suggest, why: resObj.why },
          error: null,
          countdown: 0
        }
      }));

      if (isMis && resObj.suggest) {
        setAiResults(prev => {
          const existing = prev.findIndex(r => r.item === x.item && r.vendor === x.vendor);
          const entry = {
            ...x,
            isAi: true,
            suggestedAccount: resObj.suggest,
            matchedKeyword: resObj.why || "AI Verified",
            reason: resObj.why || x.reason
          };
          if (existing >= 0) { const next = [...prev]; next[existing] = entry; return next; }
          return [...prev, entry];
        });
      }
    } catch (e) {
      setRowAiState(prev => ({ ...prev, [rowKey]: { loading: false, result: null, error: e.message, countdown: 0 } }));
    }
  };

  // Combine local and AI results
  const combinedItems = useMemo(() => {
    const map = new Map();
    items.forEach(it => map.set(it.item, { ...it, source: "Rule Engine" }));
    aiResults.forEach(it => {
      map.set(it.item, { ...it, source: "AI Web Search" });
    });
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
      (x.matchedKeyword && x.matchedKeyword.toLowerCase().includes(term)) ||
      (x.reason && x.reason.toLowerCase().includes(term))
    );
  }, [combinedItems, q]);

  const totalExposure = useMemo(() => combinedItems.reduce((s, x) => s + x.total, 0), [combinedItems]);

  return (
    <section className="panel">
      <div className="panelhead">
        <div>
          <p className="eyebrow">RESTAURANT AUDIT — RULE ENGINE + AI LAYER</p>
          <h2>Account Head Misclassifications</h2>
        </div>
        <div className="pivot-summary-badges">
          <b className="badge">{combinedItems.length} Flagged Items</b>
          <b className="badge accent-badge">{show(totalExposure)} Total Exposure</b>
        </div>
      </div>

      <div className="misclass-info-banner">
        <span>🔬</span>
        <div style={{ flex: 1 }}>
          <strong>Dual-Layer Audit:</strong> High-confidence rule engine runs instantly. Click <strong>🤖 Ask AI</strong> on any row for an instant AI second-opinion using <strong>{sharedModel || "your configured Groq model"}</strong>.
        </div>
        <div style={{ display: "flex", gap: 8 }}>
          <button className="pivot-btn" onClick={onOpenSetup}>⚙️ AI Config</button>
          <button className="pivot-btn" onClick={onGoToPivot} style={{ whiteSpace: "nowrap" }}>View Pivot &rarr;</button>
        </div>
      </div>

      <div className="pivot-toolbar">
        <div className="pivot-search-wrap">
          <span>🔍</span>
          <input
            type="text"
            className="pivot-search"
            placeholder="Search flagged item, vendor, or account..."
            value={q}
            onChange={e => setQ(e.target.value)}
          />
          {q && <button className="pivot-clear-btn" onClick={() => setQ("")}>✕</button>}
        </div>
      </div>

      <Table head={["Item Description", "Vendor", "Current Account Head", "Suggested Account Head", "Detection / AI Reason", "Amount", "AI Check"]}>
        {filtered.length ? filtered.map((x, i) => {
          const rowKey = `${x.item}:::${x.vendor}`;
          const rowAi = rowAiState[rowKey];
          return (
            <tr key={i}>
              <td>
                <div style={{ fontWeight: 600, color: "var(--ink)" }}>{x.item}</div>
                {x.reason && <small style={{ color: "var(--muted)", display: "block", marginTop: "3px", fontSize: "11px" }}>{x.reason}</small>}
                {rowAi?.result && !rowAi.result.isMisclassified && (
                  <small style={{ color: "#16a34a", display: "block", marginTop: "3px", fontSize: "11px", fontWeight: 600 }}>✅ AI: Correctly booked</small>
                )}
                {rowAi?.error && (
                  <small style={{ color: rowAi.countdown > 0 ? "#b45309" : "#dc2626", display: "block", marginTop: "3px", fontSize: "11px" }}>
                    {rowAi.countdown > 0 ? `⏳ ${rowAi.error}` : `⚠ ${rowAi.error}`}
                  </small>
                )}
              </td>
              <td>{x.vendor}</td>
              <td><span className="pill-actual-acc">{x.actualAccount}</span></td>
              <td>
                {rowAi?.result?.isMisclassified && rowAi.result.suggestedAccount ? (
                  <span className="pill-suggested-acc">🤖 {rowAi.result.suggestedAccount}</span>
                ) : (
                  <span className="pill-suggested-acc">&rarr; {x.suggestedAccount}</span>
                )}
              </td>
              <td>
                {rowAi?.result?.why ? (
                  <span className="pill-ai-badge">🤖 {rowAi.result.why}</span>
                ) : x.isAi ? (
                  <span className="pill-ai-badge">🤖 {x.matchedKeyword}</span>
                ) : (
                  <span className="rule-tag">matched "{x.matchedKeyword}"</span>
                )}
              </td>
              <td style={{ fontWeight: 700, fontFamily: "var(--font-mono, monospace)", whiteSpace: "nowrap" }}>{show(x.total)}</td>
              <td>
                <button
                  className={`ask-ai-row-btn ${rowAi?.countdown > 0 ? "rate-limited" : ""}`}
                  onClick={() => askAiForRow(x)}
                  disabled={rowAi?.loading || rowAi?.countdown > 0}
                  title={rowAi?.countdown > 0 ? `Rate limited, retry in ${rowAi.countdown}s` : `Ask AI (${sharedModel || "Setup AI"})`}
                >
                  {rowAi?.loading ? "..." : rowAi?.countdown > 0 ? `${rowAi.countdown}s` : "🤖"}
                </button>
              </td>
            </tr>
          );
        }) : (
          <tr>
            <td colSpan="7">
              <Empty>{combinedItems.length === 0 ? "🎉 No misclassifications detected by rule engine! Use 🤖 Ask AI buttons on any item to double-check with AI." : "No items match your search."}</Empty>
            </td>
          </tr>
        )}
      </Table>
    </section>
  );
}

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
      if (!accMap.has(r.account)) {
        accMap.set(r.account, new Map());
      }
      const vMap = accMap.get(r.account);
      if (!vMap.has(r.vendor)) {
        vMap.set(r.vendor, []);
      }
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
          if (!itemMap.has(key)) {
            itemMap.set(key, { name: key, qty: 0, total: 0, rates: [] });
          }
          const entry = itemMap.get(key);
          entry.qty += it.qty;
          entry.total += it.total;
          if (it.rate > 0) entry.rates.push(it.rate);
        });

        const consolidatedItems = [...itemMap.values()].map(it => ({
          name: it.name,
          qty: it.qty,
          total: it.total,
          rate: it.rates.length ? (it.total / (it.qty || 1)) : 0
        })).sort((a, b) => b.total - a.total);

        const matchesQuery = !q ||
          accName.toLowerCase().includes(q) ||
          vName.toLowerCase().includes(q) ||
          consolidatedItems.some(it => it.name.toLowerCase().includes(q));

        if (matchesQuery) {
          vendors.push({ name: vName, spend: vSpend, items: consolidatedItems });
        }
      });

      if (vendors.length > 0) {
        vendors.sort((a, b) => b.spend - a.spend);
        tree.push({ name: accName, spend: accSpend, vendors });
      }
    });

    tree.sort((a, b) => b.spend - a.spend);
    return { tree, totalSpend, accCount: accMap.size, vendorCount: allVendors.size };
  }, [current, query]);

  const allAccKeys = useMemo(() => pivotData.tree.map(a => a.name), [pivotData.tree]);
  const allVendorKeys = useMemo(() => {
    const keys = [];
    pivotData.tree.forEach(a => a.vendors.forEach(v => keys.push(`${a.name}:::${v.name}`)));
    return keys;
  }, [pivotData.tree]);

  const toggleAcc = (name) => {
    setExpandedAccs(prev => {
      const next = new Set(prev);
      if (next.has(name)) next.delete(name); else next.add(name);
      return next;
    });
  };

  const toggleVendor = (key) => {
    setExpandedVendors(prev => {
      const next = new Set(prev);
      if (next.has(key)) next.delete(key); else next.add(key);
      return next;
    });
  };

  const expandAll = () => {
    setExpandedAccs(new Set(allAccKeys));
    setExpandedVendors(new Set(allVendorKeys));
  };

  const collapseAll = () => {
    setExpandedAccs(new Set());
    setExpandedVendors(new Set());
  };

  const isSearching = query.trim().length > 0;
  const isAccOpen = (name) => isSearching || expandedAccs.has(name);
  const isVendorOpen = (key) => isSearching || expandedVendors.has(key);

  return (
    <section className="panel">
      <div className="panelhead">
        <div>
          <p className="eyebrow">CURRENT MONTH P&L HIERARCHY</p>
          <h2>Account Heads Breakdown</h2>
        </div>
        <div className="pivot-summary-badges">
          <b className="badge">{pivotData.accCount} Account Heads</b>
          <b className="badge">{pivotData.vendorCount} Vendors</b>
          <b className="badge accent-badge">{show(pivotData.totalSpend)} Total</b>
        </div>
      </div>

      {!current?.hasAccountCol && (
        <div className="pivot-warning">
          <span>ℹ️</span>
          <div>
            <strong>No "Account" column found in export</strong>
            <small>All items are grouped under "Unassigned Account". Include an "Account Name" or "Expense Account" column in your Zoho export for full P&L categorization.</small>
          </div>
        </div>
      )}

      <div className="pivot-toolbar">
        <div className="pivot-search-wrap">
          <span>🔍</span>
          <input
            type="text"
            className="pivot-search"
            placeholder="Search account head, vendor, or item..."
            value={query}
            onChange={e => setQuery(e.target.value)}
          />
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
                            <div className="pivot-right">
                              <b className="pivot-vendor-spend">{show(v.spend)}</b>
                            </div>
                          </div>

                          {vOpen && (
                            <div className="pivot-items-table">
                              <table>
                                <thead>
                                  <tr>
                                    <th>Item Description</th>
                                    <th style={{ textAlign: "right" }}>Quantity</th>
                                    <th style={{ textAlign: "right" }}>Avg Rate</th>
                                    <th style={{ textAlign: "right" }}>Total Amount</th>
                                  </tr>
                                </thead>
                                <tbody>
                                  {v.items.map((it, idx) => (
                                    <tr key={idx}>
                                      <td className="pivot-item-name">
                                        <span className="bullet">•</span> {it.name}
                                      </td>
                                      <td style={{ textAlign: "right", fontFamily: "var(--font-mono, monospace)" }}>
                                        {it.qty > 0 ? it.qty.toLocaleString("en-IN") : "-"}
                                      </td>
                                      <td style={{ textAlign: "right", fontFamily: "var(--font-mono, monospace)" }}>
                                        {it.rate > 0 ? show(it.rate) : "-"}
                                      </td>
                                      <td style={{ textAlign: "right", fontWeight: "600", fontFamily: "var(--font-mono, monospace)" }}>
                                        {show(it.total)}
                                      </td>
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

// ─── DATA PREVIEW COMPONENT ─────────────────────────────────────────────────
function DataPreview({ data }) {
  if (!data) return null;
  const { name, records, colMap } = data;
  const sample = records.slice(0, 20);
  const blankItem = records.filter(r => !r.item).length;
  const blankAccount = records.filter(r => !r.account || r.account === "Unassigned Account").length;
  return (
    <section className="panel" style={{ marginTop: 0 }}>
      <div className="panelhead">
        <div>
          <p className="eyebrow">DATA VERIFICATION</p>
          <h2>Parsed File Preview — {name}</h2>
        </div>
        <div className="pivot-summary-badges">
          <b className="badge">{records.length} rows parsed</b>
          {blankAccount > 0 && <b className="badge" style={{ background: "#fef3c7", color: "#92400e" }}>⚠ {blankAccount} rows missing Account</b>}
          {blankItem > 0 && <b className="badge" style={{ background: "#fef3c7", color: "#92400e" }}>⚠ {blankItem} rows missing Item Name</b>}
        </div>
      </div>
      {colMap && (
        <div className="col-map-bar">
          {Object.entries(colMap).filter(([, v]) => v !== null).map(([k, v]) => (
            <span key={k} className="col-map-tag">
              <strong>{k}</strong> → Col {v}
            </span>
          ))}
        </div>
      )}
      <div className="table">
        <table>
          <thead>
            <tr>
              <th>#</th>
              <th>Date</th>
              <th>Vendor</th>
              <th>Bill No.</th>
              <th style={{ background: blankAccount > 0 ? "#fef9c3" : undefined }}>Account Head</th>
              <th style={{ background: blankItem > 0 ? "#fef9c3" : undefined }}>Item Name</th>
              <th>Qty</th>
              <th>Rate</th>
              <th>Total</th>
            </tr>
          </thead>
          <tbody>
            {sample.map((r, i) => (
              <tr key={i}>
                <td style={{ color: "var(--muted)", fontSize: 11 }}>{i + 1}</td>
                <td>{r.date}</td>
                <td style={{ fontWeight: 500 }}>{r.vendor}</td>
                <td style={{ color: "var(--muted)" }}>{r.bill || "-"}</td>
                <td style={{ background: (!r.account || r.account === "Unassigned Account") ? "#fef3c7" : undefined }}>
                  {r.account || <em style={{ color: "#ef4444" }}>Missing</em>}
                </td>
                <td style={{ background: !r.item ? "#fef3c7" : undefined }}>
                  {r.item || <em style={{ color: "#ef4444" }}>Missing</em>}
                </td>
                <td>{r.qty || "-"}</td>
                <td>{r.rate ? show(r.rate) : "-"}</td>
                <td style={{ fontWeight: 600 }}>{r.total ? show(r.total) : "-"}</td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
      {records.length > 20 && <p style={{ padding: "8px 16px", color: "var(--muted)", fontSize: 12 }}>Showing first 20 of {records.length} rows. Scroll to Audit tabs for full analysis.</p>}
    </section>
  );
}

// ─── AI SETUP MODAL ──────────────────────────────────────────────────────────
function AiSetupModal({ isOpen, onClose, apiKey, model, onSave }) {
  const [step, setStep] = useState("key"); // "key" | "model" | "testing" | "configured"
  const [keyInput, setKeyInput] = useState(apiKey || "");
  const [selectedModel, setSelectedModel] = useState(model || "");
  const [availableModels, setAvailableModels] = useState([]);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState("");

  useEffect(() => {
    setKeyInput(apiKey || "");
    setSelectedModel(model || "");
    if (apiKey && model) setStep("configured");
    else setStep("key");
    setError("");
  }, [isOpen, apiKey, model]);

  const fetchModels = async (overrideKey) => {
    const key = overrideKey || keyInput.trim();
    if (!key) { setError("Please enter your Groq API key (starts with gsk_...)"); return; }
    setError("");
    setLoading(true);
    try {
      const res = await fetch("/api/groq-models", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ apiKey: key })
      });
      const data = await res.json();
      if (!res.ok || data.error) throw new Error(data.error || "Failed to connect to Groq");
      const models = data.models || [];
      if (!models.length) throw new Error("No chat models found for this key.");
      setAvailableModels(models);
      const choice = (model && models.includes(model)) ? model : models[0];
      setSelectedModel(choice);
      setStep("model");
    } catch (e) {
      setError(e.message);
    } finally {
      setLoading(false);
    }
  };

  const testAndSave = async () => {
    if (!selectedModel) { setError("Please choose a model"); return; }
    setLoading(true);
    setError("");
    setStep("testing");
    try {
      const res = await fetch("/api/ai-audit", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          singleItem: { item: "Amul Butter 500g", vendor: "Amul Store", actualAccount: "Dairy" },
          availableAccounts: ["Dairy", "Groceries"],
          apiKey: keyInput.trim(),
          model: selectedModel
        })
      });
      const data = await res.json();
      if (res.status === 429) {
        throw new Error(`Rate limit hit on Groq (${data.retryAfter || 30}s). Try 'llama-3.1-8b-instant' which has a higher quota.`);
      }
      if (!res.ok || data.error) throw new Error(data.error || "Model test failed");
      onSave(keyInput.trim(), selectedModel);
      setStep("configured");
    } catch (e) {
      setError(e.message);
      setStep("model");
    } finally {
      setLoading(false);
    }
  };

  if (!isOpen) return null;

  return (
    <div className="setup-overlay">
      <div className="setup-modal">
        <div className="setup-modal-header">
          <strong>⚙️ AI Setup & Configuration</strong>
          <button className="setup-close-btn" onClick={onClose}>✕</button>
        </div>

        <div className="setup-body">
          {step === "key" && (
            <>
              <p className="setup-desc">
                Paste your free Groq API key. If you need one, create it for free at <a href="https://console.groq.com/keys" target="_blank" rel="noreferrer">console.groq.com/keys</a>.
              </p>
              <label className="setup-label">Groq API Key</label>
              <input
                type="password"
                className="setup-input"
                placeholder="gsk_..."
                value={keyInput}
                onChange={e => setKeyInput(e.target.value)}
                onKeyDown={e => e.key === "Enter" && fetchModels()}
              />
              {error && <div className="setup-error"><span>⚠</span> {error}</div>}
              <div className="setup-actions">
                <button className="btn-cancel" onClick={onClose}>Cancel</button>
                <button className="btn-save" onClick={() => fetchModels()} disabled={loading || !keyInput.trim()}>
                  {loading ? "Checking..." : "Next: Choose Model →"}
                </button>
              </div>
            </>
          )}

          {step === "model" && (
            <>
              <p className="setup-desc">
                Choose the model from your active Groq account to use for audit reviews and chat:
              </p>
              <label className="setup-label">Available Models for your Key</label>
              <select
                className="setup-select"
                value={selectedModel}
                onChange={e => setSelectedModel(e.target.value)}
              >
                {availableModels.map(m => (
                  <option key={m} value={m}>
                    {m} {m.includes("instant") ? "⚡ (Recommended — Fast & High Quota)" : m.includes("70b") ? "🧠 (Deep Accuracy)" : ""}
                  </option>
                ))}
              </select>

              <div className="model-info-box">
                <strong>Tip for Free Tier</strong>
                <code>llama-3.1-8b-instant</code> offers the best performance with high free tokens-per-minute limits.
              </div>

              {error && <div className="setup-error"><span>⚠</span> {error}</div>}

              <div className="setup-actions">
                <button className="btn-cancel" onClick={() => setStep("key")}>← Change Key</button>
                <button className="btn-save" onClick={testAndSave} disabled={loading || !selectedModel}>
                  {loading ? "Testing..." : "Test & Save Config →"}
                </button>
              </div>
            </>
          )}

          {step === "testing" && (
            <div className="setup-testing">
              <span className="setup-spinner">⚙️</span>
              <p>Testing connection with <strong>{selectedModel}</strong>...</p>
            </div>
          )}

          {step === "configured" && (
            <div className="setup-success">
              <span className="setup-success-icon">✅</span>
              <h4>AI Configured & Ready</h4>
              <p>Active Model: <strong>{model || selectedModel}</strong></p>
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
  const messagesEndRef = typeof window !== "undefined" ? { current: null } : null;

  const sendMessage = async () => {
    const msg = input.trim();
    if (!msg || loading || chatCountdown > 0) return;
    if (!apiKey || !model) {
      onOpenSetup();
      return;
    }

    const newHistory = [...history, { role: "user", content: msg }];
    setHistory(newHistory);
    setInput("");
    setLoading(true);

    try {
      const res = await fetch("/api/ai-chat", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          message: msg,
          history: history.slice(-4),
          availableAccounts,
          apiKey,
          model
        })
      });
      const data = await res.json();

      if (res.status === 429) {
        const secs = data.retryAfter || 30;
        setChatCountdown(secs);
        setHistory([...newHistory, { role: "assistant", content: `⏳ Rate limit reached. Please wait ${secs}s before asking again.` }]);
        let rem = secs;
        const iv = setInterval(() => {
          rem -= 1;
          if (rem <= 0) { clearInterval(iv); setChatCountdown(0); }
          else setChatCountdown(rem);
        }, 1000);
        return;
      }

      if (data.error) throw new Error(data.error);
      setHistory([...newHistory, { role: "assistant", content: data.reply }]);
    } catch (e) {
      setHistory([...newHistory, { role: "assistant", content: `⚠️ Error: ${e.message}` }]);
    } finally {
      setLoading(false);
    }
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
            <div>
              <strong>🤖 AI Assistant</strong>
              <small>{model ? `Using ${model}` : "Click ⚙️ to configure Groq"}</small>
            </div>
            <div style={{ display: "flex", gap: 6 }}>
              <button className="chat-key-btn" onClick={onOpenSetup} title="Configure AI Key & Model">⚙️</button>
              <button className="chat-close-btn" onClick={() => setOpen(false)}>✕</button>
            </div>
          </div>

          <div className="chat-messages">
            {history.length === 0 && (
              <div className="chat-empty">
                <p>👋 Hi! Paste any item name or ask about restaurant accounting.</p>
                {!apiKey && (
                  <p style={{ marginTop: 10 }}>
                    <button className="pivot-btn" onClick={onOpenSetup} style={{ width: "100%", justifyContent: "center" }}>
                      ⚙️ Setup Groq API Key & Model
                    </button>
                  </p>
                )}
                <p style={{ marginTop: 10, fontSize: 12, fontWeight: 600 }}>Try asking:</p>
                <div className="chat-examples">
                  {["What is VANILLA 4LTR FD 368?", "Is Monin Watermelon a beverage?", "Which account for Classic Connect FTK?"].map(ex => (
                    <button key={ex} onClick={() => setInput(ex)} className="chat-example-chip">{ex}</button>
                  ))}
                </div>
              </div>
            )}
            {history.map((m, i) => (
              <div key={i} className={`chat-msg chat-msg-${m.role}`}>
                <div className="chat-bubble">{m.content}</div>
              </div>
            ))}
            {loading && (
              <div className="chat-msg chat-msg-assistant">
                <div className="chat-bubble chat-typing">⠋ Thinking...</div>
              </div>
            )}
            <div ref={messagesEndRef} />
          </div>

          <div className="chat-input-row">
            <input
              type="text"
              className="chat-input"
              placeholder={chatCountdown > 0 ? `Wait ${chatCountdown}s...` : "Type item name..."}
              value={input}
              onChange={e => setInput(e.target.value)}
              onKeyDown={e => e.key === "Enter" && !e.shiftKey && sendMessage()}
              disabled={loading || chatCountdown > 0}
            />
            <button className="chat-send-btn" onClick={sendMessage} disabled={loading || !input.trim() || chatCountdown > 0}>
              {loading ? "…" : chatCountdown > 0 ? `${chatCountdown}` : "↑"}
            </button>
          </div>
        </div>
      )}
    </>
  );
}

export default function Home() {
  const [current, setCurrent] = useState(null), [previous, setPrevious] = useState(null), [error, setError] = useState(""), [tab, setTab] = useState("Overview");
  const [thresholds, setThresholds] = useState({ vendor: 20, item: 25, price: 20 });
  const [apiKey, setApiKey] = useState(() => (typeof window !== "undefined" ? localStorage.getItem("groq_api_key") || "" : ""));
  const [groqModel, setGroqModel] = useState(() => (typeof window !== "undefined" ? localStorage.getItem("groq_model") || "" : ""));
  const [showAiSetup, setShowAiSetup] = useState(false);
  const [showDataPreview, setShowDataPreview] = useState(false);

  const saveAiConfig = (key, model) => {
    setApiKey(key);
    setGroqModel(model);
    if (typeof window !== "undefined") {
      localStorage.setItem("groq_api_key", key);
      localStorage.setItem("groq_model", model);
    }
  };

  const setT = (k, v) => setThresholds(t => ({ ...t, [k]: v }));
  const result = useMemo(() => current && previous ? analyse(current, previous, thresholds) : null, [current, previous, thresholds]);
  const upload = async (file, setter) => { try { setError(""); setter(await readFile(file)); } catch (e) { setError(e.message); } };
  const spend = x => x?.records.reduce((s, r) => s + r.total, 0) || 0;
  const total = result && (result.duplicates.length + result.vendors.length + result.items.length + result.prices.length + result.misclassifications.length);
  const sharedAccounts = useMemo(() => current ? [...new Set(current.records.map(r => r.account).filter(a => a && a !== "Unassigned Account"))] : [], [current]);

  return (
    <main>
      <header>
        <div>
          <p className="eyebrow">FINANCE CONTROL CENTER</p>
          <h1>P&L Audit Desk</h1>
        </div>
        <div style={{ display: "flex", alignItems: "center", gap: 12 }}>
          <button className={`header-ai-btn ${groqModel ? "configured" : ""}`} onClick={() => setShowAiSetup(true)}>
            ⚙️ AI: {groqModel ? groqModel.replace("llama-", "").replace("-versatile", "").replace("-instant", "") : "Setup"}
          </button>
          <p className="privacy"><i /> Local analysis - files stay on your device</p>
        </div>
      </header>

      {!result ? (
        <section className="landing">
          <p className="eyebrow">ZOHO BOOKS PURCHASE & P&L REVIEW</p>
          <h2>Find what the spreadsheet misses.</h2>
          <p className="lead">Compare two Zoho purchase exports to flag duplicate bills, account misclassifications, vendor movements, item variation, price exceptions, and explore your Account Head hierarchy.</p>
          <div className="chips">
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
              <button className="ai-setup-run-btn" onClick={() => setShowAiSetup(true)}>
                ⚙️ {groqModel ? `AI: ${groqModel}` : "Configure AI"}
              </button>
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
                  <div>
                    <p className="eyebrow">PRIORITY QUEUE</p>
                    <h2>What to review first</h2>
                  </div>
                  <b className="badge">
                    {result.misclassifications.length + result.duplicates.length} critical issues
                  </b>
                </div>
                {result.misclassifications.length || result.duplicates.length || result.vendors.length || result.prices.length ? (
                  <div className="queue">
                    {[
                      ...result.misclassifications.slice(0, 3).map(x => ({
                        title: `Wrong Account: ${x.item}`,
                        detail: `Booked in "${x.actualAccount}" → Should be "${x.suggestedAccount}"`,
                        value: x.total,
                        red: true
                      })),
                      ...result.duplicates.slice(0, 3).map(x => ({
                        title: x.kind,
                        detail: `${x.rows[0].vendor} - ${x.rows.length} matching lines`,
                        value: x.total,
                        red: x.risk === "Critical"
                      })),
                      ...result.vendors.slice(0, 2).map(x => ({
                        title: `${x.status} vendor`,
                        detail: x.label,
                        value: x.diff
                      })),
                      ...result.prices.slice(0, 2).map(x => ({
                        title: "High item price",
                        detail: `${x.item} - ${x.vendor}`,
                        value: x.total
                      }))
                    ].slice(0, 8).map((x, i) => (
                      <div className="queueitem" key={i}>
                        <i className={x.red ? "red" : "amber"} />
                        <div><strong>{x.title}</strong><small>{x.detail}</small></div>
                        <b>{show(x.value)}</b>
                      </div>
                    ))}
                  </div>
                ) : (
                  <Empty>No material flags found.</Empty>
                )}
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
            />
          )}

          {tab === "Account heads" && <AccountPivot current={current} />}

          {tab === "Duplicate bills" && (
            <section className="panel">
              <Panel title="Duplicate bill patterns" badge={`${result.duplicates.length} findings`} />
              <Table head={["Classification", "Vendor", "Item", "Matching lines", "Exposure"]}>
                {result.duplicates.length ? result.duplicates.map((x, i) => (
                  <tr key={i}>
                    <td><b className={x.risk === "Critical" ? "pill critical" : "pill"}>{x.kind}</b></td>
                    <td>{x.rows[0].vendor}</td>
                    <td>{x.rows[0].item || "-"}</td>
                    <td>{x.rows.length}</td>
                    <td>{show(x.total)}</td>
                  </tr>
                )) : <tr><td colSpan="5"><Empty>No duplicate patterns found.</Empty></td></tr>}
              </Table>
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
                    <td>{x.item}</td>
                    <td>{x.vendor}</td>
                    <td>{show(x.rate)}</td>
                    <td>{show(x.avg)}</td>
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

      {/* AI Setup & Configuration Modal */}
      <AiSetupModal
        isOpen={showAiSetup}
        onClose={() => setShowAiSetup(false)}
        apiKey={apiKey}
        model={groqModel}
        onSave={saveAiConfig}
      />

      {/* AI Chat Widget — always visible when file loaded */}
      {current && (
        <AiChatWidget
          availableAccounts={sharedAccounts}
          apiKey={apiKey}
          model={groqModel}
          onOpenSetup={() => setShowAiSetup(true)}
        />
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
            <td>{x.label}</td>
            <td>{show(x.total)}</td>
            <td>{show(x.old)}</td>
            <td className={x.diff >= 0 ? "bad" : "good"}>{x.diff >= 0 ? "+" : ""}{show(x.diff)}</td>
            <td><b className="pill">{x.status}</b></td>
          </tr>
        )) : <tr><td colSpan="5"><Empty>No material movements found.</Empty></td></tr>}
      </Table>
    </section>
  );
}

