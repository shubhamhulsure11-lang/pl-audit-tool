"use client";
import { Fragment, useCallback, useEffect, useMemo, useState } from "react";
import * as XLSX from "xlsx";
import JSZip from "jszip";
import {
  extractZipRecursively,
  classifyDeliverableFiles,
  auditZomatoReconciliation,
  auditSwiggyReconciliation,
} from "./lib/reconciliation";
import {
  runZomatoRecon,
  runSwiggyRecon,
  runDineoutRecon,
  runZomatoPayRecon,
  runPaytmRecon,
  runPosCleaner,
  exportReconWorkbook,
  exportFullReconWorkbook,
  getCashflowRowDefs,
  getProfitRowDefs,
} from "./lib/reconEngines";

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
    id: "beverages",
    label: "Beverages",
    aliases: ["beverage", "soft drink", "drinks", "cold drink", "non alcoholic"],
    keywords: [
      "real apple juice", "real mango juice", "real orange juice", "real cranberry", "real pineapple",
      "real litchi", "real juice", "tropicana", "minute maid", "raw pressery", "paper boat",
      "red bull", "redbull", "monster energy", "sting energy",
      "tonic water", "schweppes tonic", "ginger ale", "gin ale", "schw gin ale", "club soda", "soda water", "lehar soda", "kinley soda",
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
    aliases: ["grocer", "grocery", "provision", "dry good", "raw material", "staple", "spices", "grain", "ingredients"],
    keywords: [
      "sona masoori", "basmati rice", "kolam rice", "ponni rice", "rice", "basmati", "kolam",
      "wheat flour", "atta", "maida", "sooji", "semolina", "besan", "cornflour", "corn flour",
      "cornstarch", "custard powder", "bread crumbs", "breadcrumbs",
      "toor dal", "tur dal", "moong dal", "mung dal", "urad dal", "chana dal",
      "kabuli chana", "rajma", "soya chunks", "poha", "vermiceli", "sevai", "noodles", "pasta", "macaroni", "spaghetti", "croutons", "dal", "lentil", "lentils",
      "garam masala", "chaat masala", "biryani masala", "kitchen king masala", "pav bhaji masala", "sambar powder", "rasam powder", "coriander powder", "chilli powder", "turmeric powder", "mustard seed", "fenugreek seed", "methi seed", "bay leaf", "tej patta", "black pepper", "kali mirch", "white pepper", "kasuri methi", "red chilli powder", "kashmiri chilli", "degi mirch", "chili flakes",
      "spices", "masala", "haldi", "turmeric", "jeera", "cumin", "dhania", "rai", "sarson", "saunf", "fennel", "cardamom", "elaichi", "laung", "cinnamon", "dalchini", "star anise", "nutmeg", "jaiphal", "saffron", "kesar", "ajwain", "kalonji", "hing", "asafoetida", "oregano", "thyme", "rosemary", "paprika",
      "sunflower oil", "mustard oil", "groundnut oil", "peanut oil", "sesame oil", "til oil", "olive oil", "canola oil", "soybean oil", "palm oil", "vanaspati", "dalda", "refined oil", "cooking oil",
      "desi ghee", "cow ghee", "buffalo ghee", "amul ghee", "ghee",
      "brown sugar", "jaggery", "gur", "honey", "rock salt", "black salt", "sendha namak", "pink salt", "baking soda", "baking powder", "yeast", "citric acid", "ajinomoto", "msg", "sugar", "salt",
      "tomato ketchup", "red chilli sauce", "green chilli sauce", "soya sauce", "dark soy", "white vinegar", "apple cider vinegar", "sriracha", "tabasco", "schezwan sauce", "mayonnaise", "mayo", "mustard paste", "salsa", "peri peri sauce", "ketchup", "vinegar", "wine vinegar", "cooking wine", "shao hsing",
      "pickle", "achaar", "murabba", "chutney", "papad", "appalam", "tamarind", "imli", "desiccated coconut", "black olive", "green olive", "stuffed olive", "olive slice",
      "coconut milk powder", "coconut milk", "coconut cream", "maggi coconut",
      "cashew", "kaju", "badam", "almond", "kismis", "raisin", "pista", "pistachio", "walnut", "akhrot", "dates", "khajoor", "melon seeds", "poppy seeds", "khus khus",
      "cocoa powder", "cooking chocolate", "chocolate chips", "vanilla essence", "vanilla extract", "vanilla 4ltr", "food color", "food colouring",
      "knorr chicken broth", "chicken broth", "chicken powder", "knorr", "broth powder", "bouillon", "seasoning powder", "crab cake mix"
    ]
  },
  {
    id: "dairy",
    label: "Dairy",
    aliases: ["dairy", "milk product"],
    keywords: [
      "toned milk", "full cream milk", "cow milk", "buffalo milk", "cottage cheese", "hung curd", "fresh cream", "amul cream", "sour cream", "whipped cream", "salted butter", "unsalted butter", "table butter", "mozzarella", "cheddar cheese", "cheese slice", "cheese block", "cream cheese",
      "fresh milk", "paneer", "curd", "dahi", "yogurt", "yoghurt", "cream", "butter", "cheese", "khoya", "mawa", "buttermilk", "chaas", "lassi"
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
      "paper straw", "plastic straw", "straw", "wooden spoon", "plastic spoon", "wooden fork", "plastic fork", "disposable cutlery", "chopstick", "chopsticks",
      "banana leaf", "banana leaves", "patra", "serving leaf"
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
      "fresh apple", "seb", "fresh banana", "kela", "fresh orange", "santra", "mosambi", "pomegranate", "anar", "fresh watermelon", "tarbooj", "muskmelon", "kharbuj", "papaya", "papita", "fresh pineapple", "ananas", "aam", "grapes", "angoor", "kiwi", "guava", "amrood", "pear", "chikoo", "dragonfruit", "plum", "peach", "cherry"
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
  
  // STAFF WELFARE / ENTERTAINMENT — any food item (raw or cooked) bought under
  // these accounts is a legitimate expense. Never flag them as mis-classified.
  const isStaffOrExpense = /staff\s*(welfare|food|meal|expense)|entertainment|business\s*meal|petty\s*cash|general\s*expense/i.test(normAcc);
  if (isStaffOrExpense && ["seafood","poultry","groceries","beverages","dairy","vegetables","fruits","liquor"].includes(category.id)) return true;

  // Direct specific exclusions — stop wrong cross-category matches
  if (category.id === "groceries" && normAcc.includes("sea food")) return false;
  if (category.id === "groceries" && (normAcc.includes("poultry") || normAcc.includes("meat"))) return false;
  if (category.id === "groceries" && normAcc.includes("dairy")) return false;
  if (category.id === "beverages" && normAcc.includes("liquor")) return false;

  // Direct alias match
  if (category.aliases.some(alias => normAcc.includes(alias))) return true;

  // Cross-category allowances:
  if (category.id === "packaging" && (normAcc.includes("pack") || normAcc.includes("clean") || normAcc.includes("housekeep") || normAcc.includes("disposab"))) return true;
  if (category.id === "cleaning" && (normAcc.includes("pack") || normAcc.includes("soap") || normAcc.includes("clean") || normAcc.includes("housekeep"))) return true;
  if (category.id === "kitchen_tools" && (normAcc.includes("hotelware") || normAcc.includes("equipment") || normAcc.includes("kitchen") || normAcc.includes("crockery") || normAcc.includes("cutlery"))) return true;
  // Vegetables & fruits ONLY accepted in groceries/food if caller explicitly opts in (opts.allowVegInGrocery)
  // This is set to true in detectMisclassifications when the file has NO separate vegetable account.
  if ((category.id === "vegetables" || category.id === "fruits") &&
      (normAcc.includes("vegetable") || normAcc.includes("fruit") || normAcc.includes("tarkari"))) return true;
  if (category.id === "groceries" && (normAcc.includes("grocer") || normAcc.includes("provision") || normAcc.includes("raw material"))) return true;

  return false;
}

// Veto rules: if item contains these signals, ignore certain category matches
const VETO_RULES = [
  // Alcohol keywords block cigarette classification
  { blockedCategoryId: "cigarettes", ifItemContains: ["rum", "vodka", "whisky", "whiskey", "gin", "tequila", "brandy", "wine", "beer", "scotch", "bourbon", "liqueur", "bacardi", "smirnoff", "absolut"] },
  // Non-alcoholic mixers block liquor classification
  { blockedCategoryId: "liquor", ifItemContains: ["ale", "ginger ale", "gin ale", "tonic", "syrup", "essence", "non alcoholic", "non-alcoholic", "mocktail", "vinegar", "cooking wine"] },
  // Soap / cleaning blocks grocery
  { blockedCategoryId: "groceries", ifItemContains: ["soap", "detergent", "cleaner", "liquid soap", "dishwash", "hand wash"] },
  // Syrups & cordials block fruit & vegetable — but NOT "leaf"/"leaves"/"patta" because
  // curry leaves, coriander leaves, mint leaves, kadi patta are VALID vegetables.
  // Only banana leaves / serving leaves / patra (lotus leaf) should be blocked (they're packing).
  { blockedCategoryId: "fruits", ifItemContains: ["monin", "syrup", "crush", "malas", "cordial", "patra"] },
  { blockedCategoryId: "vegetables", ifItemContains: ["monin", "syrup", "crush", "malas", "cordial", "patra", "banana leaf", "banana leaves"] },
  // Processed broth / seasonings / mixes block raw poultry & seafood
  { blockedCategoryId: "poultry", ifItemContains: ["broth", "powder", "seasoning", "cube", "bouillon", "knorr", "mix", "curry paste"] },
  { blockedCategoryId: "seafood", ifItemContains: ["broth", "powder", "seasoning", "cube", "bouillon", "knorr", "mix", "vinegar", "sauce", "cake mix"] },
  // Plant-based milks block dairy
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

// Find expected category for an item description using strict word-boundary matching
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

// Auditor: Detect items booked under wrong account heads
function detectMisclassifications(records) {
  if (!records || !records.length) return [];
  
  const sheetAccounts = [...new Set(records.map(r => r.account).filter(a => a && a !== "Unassigned Account"))];
  
  const findBestSheetAccountName = (category) => {
    const cid = category.id;
    if (cid === "groceries") {
      const match = sheetAccounts.find(a => /grocer/i.test(a));
      if (match) return match;
    }
    if (cid === "seafood") {
      const match = sheetAccounts.find(a => /sea\s*food|fish|prawn/i.test(a));
      if (match) return match;
    }
    if (cid === "poultry") {
      const match = sheetAccounts.find(a => /poultry|meat|chicken|mutton/i.test(a));
      if (match) return match;
    }
    if (cid === "dairy") {
      const match = sheetAccounts.find(a => /dairy|milk/i.test(a));
      if (match) return match;
    }
    if (cid === "beverages") {
      const match = sheetAccounts.find(a => /beverage|soft\s*drink/i.test(a) && !/liquor|alcohol|wine|beer/i.test(a));
      if (match) return match;
    }
    if (cid === "liquor") {
      const match = sheetAccounts.find(a => /liquor|alcohol|wine|beer|spirit/i.test(a));
      if (match) return match;
    }
    if (cid === "packaging") {
      const match = sheetAccounts.find(a => /pack|disposab/i.test(a));
      if (match) return match;
    }
    if (cid === "cleaning") {
      const match = sheetAccounts.find(a => /clean|housekeep/i.test(a));
      if (match) return match;
    }
    if (cid === "vegetables" || cid === "fruits") {
      const match = sheetAccounts.find(a => /vegetable|fruit/i.test(a));
      if (match) return match;
    }
    if (cid === "cigarettes") {
      const match = sheetAccounts.find(a => /cigarette|tobacco|smoke/i.test(a));
      if (match) return match;
    }

    const matched = sheetAccounts.find(acc => {
      const normA = acc.toLowerCase();
      if (cid === "groceries" && normA.includes("sea food")) return false;
      return category.aliases.some(alias => normA.includes(alias));
    });
    return matched || category.label;
  };

  const map = new Map();

  // Context: does this file have a dedicated vegetables/fruits account?
  const hasVegAccount = sheetAccounts.some(a => /vegetable|fruit|sabzi|tarkari/i.test(a));

  records.forEach(r => {
    if (!r.item || !r.account || r.account === "Unassigned Account") return;
    const match = classifyItem(r.item);
    if (!match) return; // Unclassified items are not flagged

    const expectedCat = match.category;
    // Check if actual account head matches expected category
    const isCorrect = accountMatchesCategory(r.account, expectedCat);

    if (!isCorrect) {
      // Smart skip: if no separate veg/fruit account exists, booking them in Groceries/Food is acceptable
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
          singleItem: {
            item: x.item,
            vendor: x.vendor,
            actualAccount: x.actualAccount,
            suggestedAccount: x.suggestedAccount,   // Rule engine suggestion
            matchedKeyword: x.matchedKeyword,         // Detection reason
            total: x.total
          },
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
      const status = resObj.classification_status || "";
      // Item is misclassified when AI says current account is wrong
      const isMis = status === "CURRENT_INCORRECT" || status === "BOTH_INCORRECT";
      // AI-confirmed account: prefer ai_final_account_head, fallback to rule engine suggestion
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
          const entry = {
            ...x,
            isAi: true,
            suggestedAccount: aiAccount,
            matchedKeyword: resObj.ai_reason || "AI Verified",
            reason: resObj.ai_reason || x.reason
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
          <strong>Dual-Layer Audit:</strong> Rule engine flags instantly. Click <strong>🤖 Ask AI</strong> for AI second-opinion — uses <strong>llama-3.1-8b-instant</strong> (fast) and auto-escalates to <strong>llama-3.3-70b-versatile</strong> (shown as 70B↑) when confidence is below 75% or review is needed.
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

      <Table head={["Item Description", "Vendor", "Current Account Head", "AI Final Account", "Detection / AI Reason", "Amount", "AI Check"]}>
        {filtered.length ? filtered.map((x, i) => {
          const rowKey = `${x.item}:::${x.vendor}`;
          const rowAi = rowAiState[rowKey];
          const ai = rowAi?.result;

          // Chip color by classification_status
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
                <div style={{ fontWeight: 600, color: "var(--ink)" }}>{x.item}</div>
                {/* Rule engine detection reason (before AI runs) */}
                {!ai && x.reason && <small style={{ color: "var(--muted)", display: "block", marginTop: "3px", fontSize: "11px" }}>{x.reason}</small>}
                {/* AI classification status chip + confidence */}
                {ai?.classificationStatus && sc && (
                  <span style={{
                    display: "inline-flex", alignItems: "center", gap: 4, marginTop: 5,
                    padding: "2px 7px", borderRadius: 12, fontSize: "10.5px", fontWeight: 700,
                    background: sc.bg, color: sc.color, border: `1px solid ${sc.border}`, letterSpacing: "0.02em"
                  }}>
                    {ai.classificationStatus.replace(/_/g, " ")}
                    {ai.confidence != null && ` · ${ai.confidence}%`}
                    {ai.escalated && <span title="Escalated to llama-3.3-70b-versatile for higher confidence" style={{ marginLeft: 3, fontSize: "9px", opacity: 0.8 }}>70B↑</span>}
                  </span>
                )}
                {/* Error / rate limit */}
                {rowAi?.error && (
                  <small style={{ color: rowAi.countdown > 0 ? "#b45309" : "#dc2626", display: "block", marginTop: "3px", fontSize: "11px" }}>
                    {rowAi.countdown > 0 ? `⏳ ${rowAi.error}` : `⚠ ${rowAi.error}`}
                  </small>
                )}
              </td>

              <td>{x.vendor}</td>

              {/* Current Account + AI verdict on it */}
              <td>
                <span className="pill-actual-acc">{x.actualAccount}</span>
                {ai?.currentVerdict === "CORRECT" && (
                  <span style={{ display: "block", marginTop: 3, fontSize: "10px", color: "#16a34a", fontWeight: 700 }}>✓ Correct</span>
                )}
                {ai?.currentVerdict === "INCORRECT" && (
                  <span style={{ display: "block", marginTop: 3, fontSize: "10px", color: "#dc2626", fontWeight: 700 }}>✗ Incorrect</span>
                )}
                {ai?.currentVerdict === "UNCERTAIN" && (
                  <span style={{ display: "block", marginTop: 3, fontSize: "10px", color: "#d97706", fontWeight: 700 }}>? Uncertain</span>
                )}
              </td>

              {/* AI final account head (overrides rule engine suggestion when AI has responded) */}
              <td>
                {ai?.classificationStatus ? (
                  ai.isMisclassified && ai.suggestedAccount ? (
                    <>
                      <span className="pill-suggested-acc">🤖 {ai.suggestedAccount}</span>
                      {ai.suggestedVerdict === "CORRECT" && (
                        <span style={{ display: "block", marginTop: 3, fontSize: "10px", color: "#1d4ed8", fontWeight: 700 }}>✓ AI Confirmed</span>
                      )}
                    </>
                  ) : ai.reviewRequired ? (
                    <span style={{ fontSize: "11px", color: "#92400e", fontWeight: 600 }}>⚠ Review required</span>
                  ) : (
                    <span style={{ fontSize: "11px", color: "#15803d", fontWeight: 600 }}>✅ Correctly booked</span>
                  )
                ) : (
                  <span className="pill-suggested-acc">&rarr; {x.suggestedAccount}</span>
                )}
              </td>

              {/* Detection reason or AI explanation */}
              <td>
                {ai?.why ? (
                  <span className="pill-ai-badge">🤖 {ai.why}</span>
                ) : ai?.reviewNote ? (
                  <span className="pill-ai-badge" style={{ background: "#fef9c3", color: "#92400e", border: "1px solid #fde68a" }}>⚠ {ai.reviewNote}</span>
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
                  title={
                    rowAi?.countdown > 0
                      ? `Rate limited, retry in ${rowAi.countdown}s`
                      : ai?.escalated
                      ? `Checked by llama-3.3-70b-versatile (escalated)`
                      : `AI second opinion (8B → 70B auto-escalate)`
                  }
                >
                  {rowAi?.loading ? "..." : rowAi?.countdown > 0 ? `${rowAi.countdown}s` : ai ? "🔄" : "🤖"}
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

  const exportCSV = () => {
    if (!pivotData.tree.length) return;

    const headers = ["Account Head", "Vendor", "Item Description", "Quantity", "Avg Rate", "Total Amount"];
    const rows = [];

    pivotData.tree.forEach(acc => {
      acc.vendors.forEach(v => {
        v.items.forEach(it => {
          rows.push([
            acc.name,
            v.name,
            it.name,
            it.qty || 0,
            it.rate ? Number(it.rate.toFixed(2)) : 0,
            it.total || 0
          ]);
        });
      });
    });

    const csvContent = [
      headers.map(h => `"${h.replace(/"/g, '""')}"`).join(","),
      ...rows.map(row => row.map(val => {
        const strVal = String(val ?? "");
        return `"${strVal.replace(/"/g, '""')}"`;
      }).join(","))
    ].join("\n");

    const blob = new Blob([csvContent], { type: "text/csv;charset=utf-8;" });
    const url = URL.createObjectURL(blob);
    const link = document.createElement("a");
    link.setAttribute("href", url);
    const clientPrefix = current?.name ? current.name.replace(/\.[^/.]+$/, "") : "audit";
    const fileName = `${clientPrefix}_account_head_pivot.csv`;
    link.setAttribute("download", fileName);
    link.style.visibility = "hidden";
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
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
          <button className="pivot-btn" onClick={exportCSV}>Export CSV</button>
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

// ─── AI SETUP MODAL (Single-source Verification) ───────────────────────────
function AiSetupModal({ isOpen, onClose, apiKey, aiConfig, onSave }) {
  const [step, setStep] = useState("key"); // "key" | "model" | "testing" | "configured"
  const [keyInput, setKeyInput] = useState(apiKey || "");
  const [selectedModel, setSelectedModel] = useState(aiConfig?.model || "");
  const [availableModels, setAvailableModels] = useState([]);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState("");

  useEffect(() => {
    setKeyInput(apiKey || "");
    setSelectedModel(aiConfig?.model || "");
    if (apiKey && aiConfig?.verified) {
      setStep("configured");
    } else {
      setStep("key");
    }
    setError("");
  }, [isOpen, apiKey, aiConfig]);

  const fetchModels = async (overrideKey) => {
    const key = overrideKey || keyInput.trim();
    if (!key) { setError("Please enter your Groq API key (starts with gsk_...)"); return; }
    setError("");
    setLoading(true);
    try {
      const res = await fetch("/api/ai-config/test", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ apiKey: key })
      });
      const data = await res.json();
      if (!res.ok || data.error) throw new Error(data.error || "Failed to validate key with Groq");
      const models = data.availableModels || [];
      if (!models.length) throw new Error("No active chat models found for this Groq key.");
      setAvailableModels(models);
      const choice = (aiConfig?.model && models.includes(aiConfig.model)) ? aiConfig.model : (data.model || models[0]);
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
      const res = await fetch("/api/ai-config/test", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          apiKey: keyInput.trim(),
          model: selectedModel
        })
      });
      const data = await res.json();
      if (!res.ok || data.error) throw new Error(data.error || "Model verification failed");
      
      onSave(keyInput.trim(), data.config);
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
                {aiConfig?.verified && !apiKey ? (
                  <span style={{ color: "#d97706", fontWeight: 600, display: "block", marginBottom: 6 }}>
                    🔑 Session Expired: Please enter your Groq API key to activate AI for this session.
                  </span>
                ) : (
                  <>Paste your free Groq API key. If you need one, create it for free at <a href="https://console.groq.com/keys" target="_blank" rel="noreferrer">console.groq.com/keys</a>.</>
                )}
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
                  {loading ? "Verifying Key..." : "Next: Choose Model →"}
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
                <code>llama-3.1-8b-instant</code> offers the best balance of speed and high free tokens-per-minute limits.
              </div>

              {error && <div className="setup-error"><span>⚠</span> {error}</div>}

              <div className="setup-actions">
                <button className="btn-cancel" onClick={() => setStep("key")}>← Change Key</button>
                <button className="btn-save" onClick={testAndSave} disabled={loading || !selectedModel}>
                  {loading ? "Verifying..." : "Test & Save Config →"}
                </button>
              </div>
            </>
          )}

          {step === "testing" && (
            <div className="setup-testing">
              <span className="setup-spinner">⚙️</span>
              <p>Performing live connection test with <strong>{selectedModel}</strong>...</p>
            </div>
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
  const [mainTab, setMainTab] = useState("purchase");
  const [current, setCurrent] = useState(null), [previous, setPrevious] = useState(null), [error, setError] = useState(""), [tab, setTab] = useState("Overview");
  const [thresholds, setThresholds] = useState({ vendor: 20, item: 25, price: 20 });
  const [aiConfig, setAiConfig] = useState(() => {
    if (typeof window === "undefined") return null;
    try {
      const raw = localStorage.getItem("ai_config");
      if (raw) return JSON.parse(raw);
    } catch (e) {}
    return null;
  });
  const [apiKey, setApiKey] = useState(() => {
    if (typeof window === "undefined") return "";
    return sessionStorage.getItem("groq_api_key") || "";
  });
  const groqModel = aiConfig?.model || "";
  const [showAiSetup, setShowAiSetup] = useState(false);
  const [showDataPreview, setShowDataPreview] = useState(false);

  const saveAiConfig = (key, config) => {
    setApiKey(key);
    setAiConfig(config);
    if (typeof window !== "undefined") {
      sessionStorage.setItem("groq_api_key", key);
      localStorage.setItem("ai_config", JSON.stringify(config));
      localStorage.removeItem("groq_api_key");
      localStorage.removeItem("groq_model");
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
          <p className="privacy"><i /> Local analysis - files stay on your device</p>
        </div>
      </header>

      <div style={{ display: "flex", gap: 0, borderBottom: "1px solid var(--line)", marginBottom: 24 }}>
        {[
          { id: "purchase", label: "Purchase Audit", sub: "Zoho P&L" },
          { id: "sales",    label: "Sales Reconciliation", sub: "POS & Aggregators" },
        ].map(t => (
          <button
            key={t.id}
            onClick={() => setMainTab(t.id)}
            style={{
              background: "none", border: 0,
              padding: "14px 24px",
              cursor: "pointer",
              textAlign: "left",
              borderBottom: mainTab === t.id ? "2px solid var(--forest)" : "2px solid transparent",
              transition: "0.15s",
            }}
          >
            <span style={{ display: "block", fontWeight: 700, fontSize: "0.9rem", color: mainTab === t.id ? "var(--forest)" : "var(--muted)" }}>{t.label}</span>
            <span style={{ display: "block", fontSize: "0.72rem", color: "var(--muted)", marginTop: 2 }}>{t.sub}</span>
          </button>
        ))}
      </div>

      {mainTab === "purchase" ? (
        <>
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
                        <td>{x.item}</td>
                        <td>{x.vendor}</td>
                        <td>{show(x.rate)}</td>
                        <td>{show(x.avg)}</td>
                        <td className="bad">+{(x.pct * 100).toFixed(0)}%</td>
                      </tr>
                    )) : <tr><td colSpan="5"><Empty>No price exceptions found.</Empty></td></tr>}
                  </Table>
                  {result.prices.length > 0 && <CopyPriceExceptionsButton prices={result.prices} />}
                </section>
              )}
            </>
          )}
        </>
      ) : (
        <SalesReconciliationView />
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
        aiConfig={aiConfig}
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

// Renders one of the shared row-def sets (Cashflow / Profit statement) as an
// on-screen table: weeks as columns, line items as rows, styled by "kind".
function LineItemBreakdownTable({ title, subtitle, report, rowDefs }) {
  const rowStyle = (kind) => {
    if (kind === "subtotal") return { fontWeight: 700, background: "#f7fbf8", borderTop: "1px solid var(--line)", borderBottom: "1px solid var(--line)" };
    if (kind === "total") return { fontWeight: 700, background: "#f0f6f2", borderTop: "2px solid var(--forest)", borderBottom: "2px solid var(--forest)" };
    return { borderBottom: "1px solid var(--line)" };
  };
  const valueColor = (kind, val) => {
    if (kind === "variance") return val === 0 ? "var(--muted)" : "#c0392b";
    if (kind === "less") return "#c0392b";
    if (kind === "add") return "#1a6f3b";
    if (kind === "total") return "var(--forest)";
    return "inherit";
  };
  const fmtCell = (val, kind) => {
    if (kind === "variance") return val === 0 ? "0" : show(val);
    return show(val);
  };

  return (
    <section className="panel" style={{ background: "#fff", borderRadius: 12, border: "1px solid var(--line)", overflow: "hidden" }}>
      <div style={{ padding: "20px 24px", borderBottom: "1px solid var(--line)" }}>
        <p className="eyebrow">{subtitle}</p>
        <h2 style={{ margin: 0, fontSize: "1.3rem" }}>{title}</h2>
      </div>
      <div style={{ overflowX: "auto" }}>
        <table style={{ width: "100%", borderCollapse: "collapse", fontSize: "0.85rem" }}>
          <thead>
            <tr style={{ background: "#f7fbf8", borderBottom: "2px solid var(--line)" }}>
              <th style={{ padding: "12px 16px", textAlign: "left" }}>Details</th>
              {report.weeks.map((w) => (
                <th key={w.weekNum} style={{ padding: "12px 16px", textAlign: "right" }}>{w.label}</th>
              ))}
              <th style={{ padding: "12px 16px", textAlign: "right", fontWeight: 700 }}>Total</th>
            </tr>
          </thead>
          <tbody>
            {rowDefs.map(([label, getter, kind], idx) => (
              <tr key={idx} style={rowStyle(kind)}>
                <td style={{ padding: "12px 16px", fontWeight: kind === "subtotal" || kind === "total" ? 700 : 500 }}>{label}</td>
                {report.weeks.map((w) => {
                  const val = getter(w);
                  return (
                    <td key={w.weekNum} style={{ padding: "12px 16px", textAlign: "right", color: valueColor(kind, val) }}>
                      {fmtCell(val, kind)}
                    </td>
                  );
                })}
                <td style={{ padding: "12px 16px", textAlign: "right", fontWeight: 700, color: valueColor(kind, getter(report.total)) }}>
                  {fmtCell(getter(report.total), kind)}
                </td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </section>
  );
}

// Discrepancies section: expected payout vs bank actual, per week.
function DiscrepancyTable({ report }) {
  return (
    <section className="panel" style={{ background: "#fff", borderRadius: 12, border: "1px solid var(--line)", overflow: "hidden" }}>
      <div style={{ padding: "20px 24px", borderBottom: "1px solid var(--line)" }}>
        <p className="eyebrow">DISCREPANCIES</p>
        <h2 style={{ margin: 0, fontSize: "1.3rem" }}>Expected Payout vs Bank Actual</h2>
        {!report.hasBank && (
          <small style={{ display: "block", marginTop: 6, color: "var(--muted)" }}>
            No bank statement uploaded — every week shows as unmatched. Upload one to check actual settlement vs expected payout.
          </small>
        )}
      </div>
      <div style={{ overflowX: "auto" }}>
        <table style={{ width: "100%", borderCollapse: "collapse", fontSize: "0.85rem" }}>
          <thead>
            <tr style={{ background: "#f7fbf8", borderBottom: "2px solid var(--line)" }}>
              <th style={{ padding: "12px 16px", textAlign: "left" }}>Week / Period</th>
              <th style={{ padding: "12px 16px", textAlign: "right" }}>Expected Payout</th>
              <th style={{ padding: "12px 16px", textAlign: "right" }}>Bank Actual</th>
              <th style={{ padding: "12px 16px", textAlign: "right" }}>Variance</th>
              <th style={{ padding: "12px 16px", textAlign: "left" }}>UTR</th>
              <th style={{ padding: "12px 16px", textAlign: "center" }}>Status</th>
            </tr>
          </thead>
          <tbody>
            {report.weeks.map((w) => {
              const expected = w.expectedPayout ?? w.expectedReceipt;
              return (
                <tr key={w.weekNum} style={{ borderBottom: "1px solid var(--line)", background: w.bankMatched ? "transparent" : "#fff8f8" }}>
                  <td style={{ padding: "12px 16px", fontWeight: 700 }}>{w.label}</td>
                  <td style={{ padding: "12px 16px", textAlign: "right" }}>{show(expected)}</td>
                  <td style={{ padding: "12px 16px", textAlign: "right" }}>{show(w.bankActual)}</td>
                  <td style={{ padding: "12px 16px", textAlign: "right", fontWeight: 700, color: w.bankDiff === 0 ? "var(--muted)" : "#c0392b" }}>
                    {w.bankDiff === 0 ? "0" : show(w.bankDiff)}
                  </td>
                  <td style={{ padding: "12px 16px", fontSize: "12px", color: "var(--muted)" }}>{w.utr || "—"}</td>
                  <td style={{ padding: "12px 16px", textAlign: "center" }}>
                    <span style={{ display: "inline-block", padding: "3px 10px", borderRadius: 20, fontSize: "11px", fontWeight: 700, background: w.bankMatched ? "#edf8f0" : "#feeceb", color: w.bankMatched ? "#1a6f3b" : "#a43024" }}>
                      {w.bankMatched ? "MATCHED" : "DISCREPANCY"}
                    </span>
                  </td>
                </tr>
              );
            })}
            <tr style={{ background: "#f0f6f2", fontWeight: 700, borderTop: "2px solid var(--forest)" }}>
              <td style={{ padding: "14px 16px" }}>{report.total.label}</td>
              <td style={{ padding: "14px 16px", textAlign: "right" }}>{show(report.total.expectedPayout ?? report.total.expectedReceipt)}</td>
              <td style={{ padding: "14px 16px", textAlign: "right" }}>{show(report.total.bankActual)}</td>
              <td style={{ padding: "14px 16px", textAlign: "right" }}>{show(report.total.bankDiff)}</td>
              <td style={{ padding: "14px 16px" }}>—</td>
              <td style={{ padding: "14px 16px", textAlign: "center" }}>
                <span style={{ display: "inline-block", padding: "3px 10px", borderRadius: 20, fontSize: "11px", fontWeight: 700, background: report.total.bankMatched ? "#edf8f0" : "#feeceb", color: report.total.bankMatched ? "#1a6f3b" : "#a43024" }}>
                  {report.total.bankMatched ? "ALL MATCHED" : "FLAGGED"}
                </span>
              </td>
            </tr>
          </tbody>
        </table>
      </div>
    </section>
  );
}

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
      {rows.length > 0 && (
        <CopyChangesButton title={title} rows={rows} field={field} threshold={threshold} />
      )}
    </section>
  );
}

function CopyDuplicatesButton({ duplicates }) {
  const [copied, setCopied] = useState(false);

  function handleCopy() {
    const cleanTsvVal = (val) => String(val ?? "").replace(/\t/g, " ").replace(/\r?\n/g, " ");
    const rows = [];
    duplicates.forEach(x => {
      x.rows.forEach(r => {
        const branch = cleanTsvVal(r.branch);
        const date = cleanTsvVal(r.date);
        const category = "";
        const type = "";
        const supplier = cleanTsvVal(r.vendor);
        const error = cleanTsvVal(x.kind);
        const review = cleanTsvVal(`Bill No: ${r.bill || "—"} | Item: ${r.item || "—"} | Qty: ${r.qty || "—"} | Amt: ${show(r.total)}`);

        rows.push([branch, date, category, type, supplier, error, review].join("\t"));
      });
    });

    navigator.clipboard.writeText(rows.join("\n")).then(() => {
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
        {copied ? (<><span>✓</span> Copied!</>) : (<><span style={{ fontSize: "1rem" }}>📋</span> Copy for Spreadsheet</>)}
      </button>
    </div>
  );
}

function CopyChangesButton({ title, rows, field, threshold }) {
  const [copied, setCopied] = useState(false);

  function handleCopy() {
    const cleanTsvVal = (val) => String(val ?? "").replace(/\t/g, " ").replace(/\r?\n/g, " ");
    const tsvRows = [];
    rows.forEach(x => {
      const branch = "";
      const date = "";
      const category = "";
      const type = "";
      const supplier = cleanTsvVal(field === "Vendor" ? x.label : "");
      const error = cleanTsvVal(title);
      const diffStr = (x.diff >= 0 ? "+" : "") + show(x.diff);
      const review = cleanTsvVal(field === "Vendor"
        ? `Current: ${show(x.total)} | Previous: ${show(x.old)} | Change: ${diffStr} | Status: ${x.status}`
        : `Item: ${x.label} | Current: ${show(x.total)} | Previous: ${show(x.old)} | Change: ${diffStr} | Status: ${x.status}`
      );

      tsvRows.push([branch, date, category, type, supplier, error, review].join("\t"));
    });

    navigator.clipboard.writeText(tsvRows.join("\n")).then(() => {
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
          <><span style={{ fontSize: "1rem" }}>📋</span> Copy for Spreadsheet</>
        )}
      </button>
    </div>
  );
}

function CopyPriceExceptionsButton({ prices }) {
  const [copied, setCopied] = useState(false);

  function handleCopy() {
    const cleanTsvVal = (val) => String(val ?? "").replace(/\t/g, " ").replace(/\r?\n/g, " ");
    const rows = [];
    prices.forEach(x => {
      const branch = cleanTsvVal(x.branch);
      const date = cleanTsvVal(x.date);
      const category = "";
      const type = "";
      const supplier = cleanTsvVal(x.vendor);
      const error = "Price exception";
      const varStr = `+${(x.pct * 100).toFixed(0)}%`;
      const review = cleanTsvVal(`Item: ${x.item} | Rate paid: ${show(x.rate)} | Weighted avg: ${show(x.avg)} | Variance: ${varStr}`);

      rows.push([branch, date, category, type, supplier, error, review].join("\t"));
    });

    navigator.clipboard.writeText(rows.join("\n")).then(() => {
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
          <><span style={{ fontSize: "1rem" }}>📋</span> Copy for Spreadsheet</>
        )}
      </button>
    </div>
  );
}

function SalesReconciliationView() {
  const [platform, setPlatform] = useState("zomato"); // zomato | swiggy | dineout | zpay | paytm | pos | zip
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState("");
  const [clientName, setClientName] = useState("Kailash Parbat");
  const [month, setMonth] = useState("July");
  const [zomatoMode, setZomatoMode] = useState("weekly"); // weekly | consolidated
  
  // Week range states
  const [fStart, setFStart] = useState(1);
  const [fEnd, setFEnd] = useState(7);
  const [lStart, setLStart] = useState(29);
  const [lEnd, setLEnd] = useState(31);

  // Uploaded files states
  const [invoiceFiles, setInvoiceFiles] = useState([]);
  const [bankFile, setBankFile] = useState(null);
  const [adsFiles, setAdsFiles] = useState([]);
  const [posFile, setPosFile] = useState(null);
  const [zipFile, setZipFile] = useState(null);

  // Result reports
  const [report, setReport] = useState(null);
  const [posReport, setPosReport] = useState(null);
  const [zipReport, setZipReport] = useState(null);
  const [copySuccess, setCopySuccess] = useState("");

  const months = [
    "January", "February", "March", "April", "May", "June",
    "July", "August", "September", "October", "November", "December"
  ];

  const resetAll = () => {
    setReport(null);
    setPosReport(null);
    setZipReport(null);
    setError("");
    setInvoiceFiles([]);
    setBankFile(null);
    setAdsFiles([]);
    setPosFile(null);
    setZipFile(null);
  };

  const handleExecute = async () => {
    setLoading(true);
    setError("");
    setReport(null);
    setPosReport(null);
    setZipReport(null);

    try {
      if (platform === "pos") {
        if (!posFile) throw new Error("Please upload a POS report file (.xlsx / .xls)");
        const res = await runPosCleaner({
          file: posFile,
          firstWeekStart: fStart,
          firstWeekEnd: fEnd,
          lastWeekStart: lStart,
          lastWeekEnd: lEnd,
        });
        setPosReport(res);
      } else if (platform === "zip") {
        if (!zipFile) throw new Error("Please upload a deliverables .zip archive");
        const extracted = await extractZipRecursively(zipFile);
        if (extracted.length === 0) throw new Error("No files found inside the uploaded zip archive.");
        const classified = classifyDeliverableFiles(extracted);
        const zAudit = auditZomatoReconciliation(classified.zomatoRaw, classified.pos, classified.zomatoSummaries);
        const sAudit = auditSwiggyReconciliation(classified.swiggyRaw, classified.pos, classified.swiggySummaries);
        setZipReport({ classified, zomato: zAudit, swiggy: sAudit });
      } else {
        if (invoiceFiles.length === 0) {
          throw new Error(`Please upload at least one ${platform.toUpperCase()} invoice/settlement file`);
        }

        let res;
        const opts = {
          files: invoiceFiles,
          bankFile,
          clientName: clientName || "Client",
          month,
          firstWeekStart: fStart,
          firstWeekEnd: fEnd,
          lastWeekStart: lStart,
          lastWeekEnd: lEnd,
        };

        if (platform === "zomato") {
          res = await runZomatoRecon({ ...opts, mode: zomatoMode });
        } else if (platform === "swiggy") {
          res = await runSwiggyRecon(opts);
        } else if (platform === "dineout") {
          res = await runDineoutRecon(opts);
        } else if (platform === "zpay") {
          res = await runZomatoPayRecon({ ...opts, adsFiles });
        } else if (platform === "paytm") {
          res = await runPaytmRecon(opts);
        }

        setReport(res);
      }
    } catch (err) {
      console.error(err);
      setError(err.message || "An error occurred during reconciliation processing.");
    } finally {
      setLoading(false);
    }
  };

  const copyTsv = () => {
    if (!report) return;
    const lines = [];
    lines.push([`Client: ${report.clientName}`, `Platform: ${report.platform}`, `Month: ${report.month}`].join("\t"));
    lines.push("");

    if (report.platform === "Zomato" || report.platform === "Swiggy") {
      lines.push(["Week / Period", "Orders", "Gross Sales", "Discounts", "GST", "Commission", "Other Deductions", "Expected Payout", "Bank Actual", "Variance", "Status"].join("\t"));
      report.weeks.forEach(w => {
        lines.push([
          w.label,
          w.orders,
          w.grossSales || w.totalIncome,
          w.discounts,
          w.gstCollected,
          w.commission,
          w.otherDeductions || w.taxesAndDeductions,
          w.expectedPayout,
          w.bankActual,
          w.bankDiff,
          w.bankMatched ? "Match" : "Discrepancy"
        ].join("\t"));
      });
      lines.push([
        report.total.label,
        report.total.orders,
        report.total.grossSales || report.total.totalIncome,
        report.total.discounts,
        report.total.gstCollected,
        report.total.commission,
        report.total.otherDeductions || report.total.taxesAndDeductions,
        report.total.expectedPayout,
        report.total.bankActual,
        report.total.bankDiff,
        report.total.bankMatched ? "All Matched" : "Discrepancy"
      ].join("\t"));
    } else {
      lines.push(["Week / Period", "Sales Excl. GST", "Discounts", "GST 5%", "Sales Incl. GST", "Commission", "Expected Payout", "Bank Actual", "Variance", "Status"].join("\t"));
      report.weeks.forEach(w => {
        lines.push([
          w.label,
          w.salesExclGst || w.salesExclGstBefore,
          w.discounts,
          w.gst5Pct,
          w.salesInclGst || w.salesAfterFailed,
          w.commission || w.commissionInclGst,
          w.expectedPayout || w.expectedReceipt,
          w.bankActual,
          w.bankDiff,
          w.bankMatched ? "Match" : "Discrepancy"
        ].join("\t"));
      });
      lines.push([
        report.total.label,
        report.total.salesExclGst || report.total.salesExclGstBefore,
        report.total.discounts,
        report.total.gst5Pct,
        report.total.salesInclGst || report.total.salesAfterFailed,
        report.total.commission || report.total.commissionInclGst,
        report.total.expectedPayout || report.total.expectedReceipt,
        report.total.bankActual,
        report.total.bankDiff,
        report.total.bankMatched ? "All Matched" : "Discrepancy"
      ].join("\t"));
    }

    navigator.clipboard.writeText(lines.join("\n")).then(() => {
      setCopySuccess("tsv");
      setTimeout(() => setCopySuccess(""), 2500);
    });
  };

  const copyNotes = () => {
    if (!report) return;
    const lines = [
      `=======================================================`,
      `  ${report.platform.toUpperCase()} RECONCILIATION AUDIT SUMMARY`,
      `  Client: ${report.clientName} | Period: ${report.month}`,
      `  Generated on: ${new Date().toLocaleString("en-IN")}`,
      `=======================================================`,
      "",
    ];

    report.weeks.forEach((w) => {
      const isDisc = !w.bankMatched && report.hasBank;
      lines.push(`Period: ${w.label}  ${isDisc ? "⚠️ DISCREPANCY" : "✅ OK"}`);
      lines.push(`- Gross Sales: ${show(w.grossSales || w.totalIncome || w.salesInclGst)}`);
      lines.push(`- Platform Commission: ${show(w.commission || w.commissionInclGst)}`);
      lines.push(`- Net Payout Expected: ${show(w.expectedPayout || w.expectedReceipt)}`);
      if (report.hasBank) {
        lines.push(`- Bank Actual Receipt: ${show(w.bankActual)} (Diff: ${show(w.bankDiff)})`);
      }
      if (w.utr && w.utr !== "—") lines.push(`- Bank UTR / CTR: ${w.utr}`);
      lines.push(`-------------------------------------------------------`);
    });

    lines.push(`TOTAL SUMMARY:`);
    lines.push(`Total Sales: ${show(report.total.grossSales || report.total.totalIncome || report.total.salesInclGst)}`);
    lines.push(`Total Commission: ${show(report.total.commission || report.total.commissionInclGst)}`);
    lines.push(`Total Expected Net: ${show(report.total.expectedPayout || report.total.expectedReceipt)}`);
    if (report.hasBank) {
      lines.push(`Total Bank Actual: ${show(report.total.bankActual)} (Variance: ${show(report.total.bankDiff)})`);
    }

    navigator.clipboard.writeText(lines.join("\n")).then(() => {
      setCopySuccess("notes");
      setTimeout(() => setCopySuccess(""), 2500);
    });
  };

  const tabs = [
    { id: "zomato", label: "Zomato", icon: "🔴", desc: "Weekly & Consolidated Payouts" },
    { id: "swiggy", label: "Swiggy", icon: "🟠", desc: "Order Level & Deductions" },
    { id: "dineout", label: "Swiggy Dineout", icon: "🍽️", desc: "Dineout Settlement Recon" },
    { id: "zpay", label: "Zomato Pay", icon: "💳", desc: "Transactions & Ads Recon" },
    { id: "paytm", label: "Paytm", icon: "🔷", desc: "MDR & Settlement Recon" },
    { id: "pos", label: "POS Extractor", icon: "📊", desc: "Petpooja / Posist Sales Channel" },
    { id: "zip", label: "Deliverables Zip", icon: "📦", desc: "Full Auto Audit Archive" },
  ];

  return (
    <div style={{ display: "flex", flexDirection: "column", gap: 24, paddingBottom: 60 }}>
      {/* Sub-Tabs Selector */}
      <div style={{ display: "flex", flexWrap: "wrap", gap: 8, background: "var(--surface)", padding: 6, borderRadius: 12, border: "1px solid var(--line)" }}>
        {tabs.map((t) => (
          <button
            key={t.id}
            onClick={() => { setPlatform(t.id); resetAll(); }}
            style={{
              display: "flex",
              alignItems: "center",
              gap: 8,
              padding: "10px 18px",
              borderRadius: 8,
              border: platform === t.id ? "1px solid var(--forest)" : "1px solid transparent",
              background: platform === t.id ? "var(--forest)" : "transparent",
              color: platform === t.id ? "#fff" : "var(--text)",
              fontWeight: 600,
              fontSize: "0.84rem",
              cursor: "pointer",
              transition: "0.15s",
            }}
          >
            <span>{t.icon}</span>
            <span>{t.label}</span>
          </button>
        ))}
      </div>

      {/* Loading View */}
      {loading && (
        <div style={{ display: "flex", flexDirection: "column", alignItems: "center", justifyContent: "center", minHeight: 340, background: "#fff", borderRadius: 12, border: "1px solid var(--line)", padding: 40, gap: 16 }}>
          <div style={{ width: 48, height: 48, border: "4px solid var(--lime)", borderTop: "4px solid var(--forest)", borderRadius: "50%", animation: "spin 0.9s linear infinite" }} />
          <strong style={{ color: "var(--forest)", fontSize: "1.1rem" }}>Executing {platform.toUpperCase()} Reconciliation...</strong>
          <small style={{ color: "var(--muted)" }}>Parsing multi-week orders, computing GST formulas, matching bank settlements & flagging variances</small>
          <style>{`@keyframes spin { 0% { transform: rotate(0deg); } 100% { transform: rotate(360deg); } }`}</style>
        </div>
      )}

      {/* Report View (Visible directly on screen!) */}
      {!loading && report && (
        <div style={{ display: "flex", flexDirection: "column", gap: 24 }}>
          {/* Header Bar */}
          <section className="run" style={{ background: "var(--forest)", color: "#fff", display: "flex", justifyContent: "space-between", alignItems: "center", borderRadius: 12, padding: "18px 24px" }}>
            <div>
              <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
                <strong style={{ fontSize: "1.2rem" }}>{report.clientName} — {report.platform} Reconciliation</strong>
                <span style={{ background: "rgba(255,255,255,0.2)", padding: "2px 8px", borderRadius: 4, fontSize: "0.75rem" }}>{report.month}</span>
                {report.mode && <span style={{ background: "var(--lime)", color: "var(--forest)", padding: "2px 8px", borderRadius: 4, fontSize: "0.75rem", fontWeight: 700 }}>{report.mode.toUpperCase()}</span>}
              </div>
              <small style={{ display: "block", opacity: 0.8, marginTop: 4 }}>
                {report.weeks.length} Weeks Audited · {report.filesCount} Source Invoices Processed · {report.hasBank ? "Bank Statement Verified" : "No Bank Statement Attached"}
              </small>
            </div>
            <div style={{ display: "flex", gap: 10 }}>
              <button onClick={() => exportReconWorkbook(report)} style={{ background: "rgba(255,255,255,0.15)", border: "1px solid rgba(255,255,255,0.3)", color: "#fff", padding: "8px 14px", borderRadius: 6, cursor: "pointer", fontWeight: 600, fontSize: "0.82rem" }}>
                📥 Quick Export (.xlsx)
              </button>
              <button onClick={() => exportFullReconWorkbook(report)} style={{ background: "var(--lime)", border: "1px solid var(--lime)", color: "var(--forest)", padding: "8px 14px", borderRadius: 6, cursor: "pointer", fontWeight: 700, fontSize: "0.82rem" }}>
                📊 Full Report (Summary / Cashflow / Profit / Discrepancies)
              </button>
              <button onClick={resetAll} style={{ background: "transparent", border: "1px solid rgba(255,255,255,0.4)", color: "#fff", padding: "8px 14px", borderRadius: 6, cursor: "pointer", fontWeight: 600, fontSize: "0.82rem" }}>
                🔄 Start New Recon
              </button>
            </div>
          </section>

          {/* Metric Summary Cards */}
          <section className="metrics" style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(200px, 1fr))", gap: 14 }}>
            <Card label="Total Gross Sales" value={show(report.total.grossSales || report.total.totalIncome || report.total.salesInclGst)} />
            <Card label="Platform Commission" value={show(report.total.commission || report.total.commissionInclGst)} />
            <Card label="Expected Net Payout" value={show(report.total.expectedPayout || report.total.expectedReceipt)} />
            {report.hasBank ? (
              <Card
                label="Bank Actual Variance"
                value={show(report.total.bankDiff)}
                warm={report.total.bankDiff !== 0}
              />
            ) : (
              <Card label="Total Orders" value={report.total.orders || "—"} />
            )}
          </section>

          {/* Comprehensive On-Screen Table */}
          <section className="panel" style={{ background: "#fff", borderRadius: 12, border: "1px solid var(--line)", overflow: "hidden" }}>
            <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", padding: "20px 24px", borderBottom: "1px solid var(--line)" }}>
              <div>
                <p className="eyebrow">AUDITOR RECONCILIATION SPREADSHEET</p>
                <h2 style={{ margin: 0, fontSize: "1.3rem" }}>Multi-Week Financial Breakdown</h2>
              </div>
              <div style={{ display: "flex", gap: 10 }}>
                <button className="pivot-btn" onClick={copyTsv}>
                  {copySuccess === "tsv" ? "✓ Copied Table!" : "📋 Copy Table for Sheets"}
                </button>
                <button className="pivot-btn" onClick={copyNotes}>
                  {copySuccess === "notes" ? "✓ Copied Notes!" : "📋 Copy Discrepancy Notes"}
                </button>
              </div>
            </div>

            <div style={{ overflowX: "auto" }}>
              <table style={{ width: "100%", borderCollapse: "collapse", fontSize: "0.85rem" }}>
                <thead>
                  <tr style={{ background: "#f7fbf8", borderBottom: "2px solid var(--line)" }}>
                    <th style={{ padding: "12px 16px", textAlign: "left" }}>Week / Period</th>
                    {report.platform === "Zomato" || report.platform === "Swiggy" ? (
                      <>
                        <th style={{ padding: "12px 16px", textAlign: "right" }}>Orders</th>
                        <th style={{ padding: "12px 16px", textAlign: "right" }}>Gross Sales</th>
                        <th style={{ padding: "12px 16px", textAlign: "right" }}>Discounts</th>
                        <th style={{ padding: "12px 16px", textAlign: "right" }}>GST Collected</th>
                        <th style={{ padding: "12px 16px", textAlign: "right" }}>Commission</th>
                        <th style={{ padding: "12px 16px", textAlign: "right" }}>Other Deductions</th>
                        <th style={{ padding: "12px 16px", textAlign: "right", fontWeight: 700 }}>Expected Payout</th>
                        {report.hasBank && <th style={{ padding: "12px 16px", textAlign: "right" }}>Bank Actual</th>}
                        {report.hasBank && <th style={{ padding: "12px 16px", textAlign: "right" }}>Variance</th>}
                        <th style={{ padding: "12px 16px", textAlign: "center" }}>UTR / Status</th>
                      </>
                    ) : (
                      <>
                        <th style={{ padding: "12px 16px", textAlign: "right" }}>Sales (Excl. GST)</th>
                        <th style={{ padding: "12px 16px", textAlign: "right" }}>Discounts</th>
                        <th style={{ padding: "12px 16px", textAlign: "right" }}>GST 5%</th>
                        <th style={{ padding: "12px 16px", textAlign: "right" }}>Sales (Incl. GST)</th>
                        <th style={{ padding: "12px 16px", textAlign: "right" }}>Commission</th>
                        <th style={{ padding: "12px 16px", textAlign: "right", fontWeight: 700 }}>Expected Net</th>
                        {report.hasBank && <th style={{ padding: "12px 16px", textAlign: "right" }}>Bank Actual</th>}
                        {report.hasBank && <th style={{ padding: "12px 16px", textAlign: "right" }}>Variance</th>}
                        <th style={{ padding: "12px 16px", textAlign: "center" }}>Status</th>
                      </>
                    )}
                  </tr>
                </thead>
                <tbody>
                  {report.weeks.map((w, idx) => {
                    const hasVar = report.hasBank && !w.bankMatched;
                    return (
                      <tr key={idx} style={{ borderBottom: "1px solid var(--line)", background: hasVar ? "#fff8f8" : "transparent" }}>
                        <td style={{ padding: "14px 16px", fontWeight: 700 }}>{w.label}</td>
                        {report.platform === "Zomato" || report.platform === "Swiggy" ? (
                          <>
                            <td style={{ padding: "14px 16px", textAlign: "right", color: "var(--muted)" }}>{w.orders}</td>
                            <td style={{ padding: "14px 16px", textAlign: "right", fontWeight: 600 }}>{show(w.grossSales || w.totalIncome)}</td>
                            <td style={{ padding: "14px 16px", textAlign: "right", color: "#c0392b" }}>-{show(w.discounts)}</td>
                            <td style={{ padding: "14px 16px", textAlign: "right" }}>{show(w.gstCollected)}</td>
                            <td style={{ padding: "14px 16px", textAlign: "right", color: "#c0392b" }}>-{show(w.commission)}</td>
                            <td style={{ padding: "14px 16px", textAlign: "right", color: "var(--muted)" }}>-{show(w.otherDeductions || w.taxesAndDeductions)}</td>
                            <td style={{ padding: "14px 16px", textAlign: "right", fontWeight: 700, color: "var(--forest)" }}>{show(w.expectedPayout)}</td>
                            {report.hasBank && <td style={{ padding: "14px 16px", textAlign: "right" }}>{show(w.bankActual)}</td>}
                            {report.hasBank && (
                              <td style={{ padding: "14px 16px", textAlign: "right", fontWeight: 700, color: w.bankDiff === 0 ? "var(--muted)" : "#c0392b" }}>
                                {w.bankDiff === 0 ? "0" : show(w.bankDiff)}
                              </td>
                            )}
                            <td style={{ padding: "14px 16px", textAlign: "center" }}>
                              {report.hasBank ? (
                                <span style={{ display: "inline-block", padding: "3px 10px", borderRadius: 20, fontSize: "11px", fontWeight: 700, background: w.bankMatched ? "#edf8f0" : "#feeceb", color: w.bankMatched ? "#1a6f3b" : "#a43024" }}>
                                  {w.bankMatched ? "Match" : "Discrepancy"}
                                </span>
                              ) : (
                                <span style={{ fontSize: "11px", color: "var(--muted)" }}>{w.utr || "—"}</span>
                              )}
                            </td>
                          </>
                        ) : (
                          <>
                            <td style={{ padding: "14px 16px", textAlign: "right" }}>{show(w.salesExclGst || w.salesExclGstBefore)}</td>
                            <td style={{ padding: "14px 16px", textAlign: "right", color: "#c0392b" }}>-{show(w.discounts)}</td>
                            <td style={{ padding: "14px 16px", textAlign: "right" }}>{show(w.gst5Pct)}</td>
                            <td style={{ padding: "14px 16px", textAlign: "right", fontWeight: 600 }}>{show(w.salesInclGst || w.salesAfterFailed)}</td>
                            <td style={{ padding: "14px 16px", textAlign: "right", color: "#c0392b" }}>-{show(w.commission || w.commissionInclGst)}</td>
                            <td style={{ padding: "14px 16px", textAlign: "right", fontWeight: 700, color: "var(--forest)" }}>{show(w.expectedPayout || w.expectedReceipt)}</td>
                            {report.hasBank && <td style={{ padding: "14px 16px", textAlign: "right" }}>{show(w.bankActual)}</td>}
                            {report.hasBank && (
                              <td style={{ padding: "14px 16px", textAlign: "right", fontWeight: 700, color: w.bankDiff === 0 ? "var(--muted)" : "#c0392b" }}>
                                {w.bankDiff === 0 ? "0" : show(w.bankDiff)}
                              </td>
                            )}
                            <td style={{ padding: "14px 16px", textAlign: "center" }}>
                              <span style={{ display: "inline-block", padding: "3px 10px", borderRadius: 20, fontSize: "11px", fontWeight: 700, background: w.bankMatched ? "#edf8f0" : "#feeceb", color: w.bankMatched ? "#1a6f3b" : "#a43024" }}>
                                {w.bankMatched ? "Match" : "Discrepancy"}
                              </span>
                            </td>
                          </>
                        )}
                      </tr>
                    );
                  })}

                  {/* Total Row */}
                  <tr style={{ background: "#f0f6f2", fontWeight: 700, borderTop: "2px solid var(--forest)" }}>
                    <td style={{ padding: "16px" }}>{report.total.label}</td>
                    {report.platform === "Zomato" || report.platform === "Swiggy" ? (
                      <>
                        <td style={{ padding: "16px", textAlign: "right" }}>{report.total.orders}</td>
                        <td style={{ padding: "16px", textAlign: "right" }}>{show(report.total.grossSales || report.total.totalIncome)}</td>
                        <td style={{ padding: "16px", textAlign: "right", color: "#c0392b" }}>-{show(report.total.discounts)}</td>
                        <td style={{ padding: "16px", textAlign: "right" }}>{show(report.total.gstCollected)}</td>
                        <td style={{ padding: "16px", textAlign: "right", color: "#c0392b" }}>-{show(report.total.commission)}</td>
                        <td style={{ padding: "16px", textAlign: "right" }}>-{show(report.total.otherDeductions || report.total.taxesAndDeductions)}</td>
                        <td style={{ padding: "16px", textAlign: "right", color: "var(--forest)" }}>{show(report.total.expectedPayout)}</td>
                        {report.hasBank && <td style={{ padding: "16px", textAlign: "right" }}>{show(report.total.bankActual)}</td>}
                        {report.hasBank && <td style={{ padding: "16px", textAlign: "right" }}>{show(report.total.bankDiff)}</td>}
                        <td style={{ padding: "16px", textAlign: "center" }}>
                          {report.hasBank && (
                            <span style={{ display: "inline-block", padding: "3px 10px", borderRadius: 20, fontSize: "11px", fontWeight: 700, background: report.total.bankMatched ? "#edf8f0" : "#feeceb", color: report.total.bankMatched ? "#1a6f3b" : "#a43024" }}>
                              {report.total.bankMatched ? "ALL MATCHED" : "DISCREPANCY"}
                            </span>
                          )}
                        </td>
                      </>
                    ) : (
                      <>
                        <td style={{ padding: "16px", textAlign: "right" }}>{show(report.total.salesExclGst || report.total.salesExclGstBefore)}</td>
                        <td style={{ padding: "16px", textAlign: "right", color: "#c0392b" }}>-{show(report.total.discounts)}</td>
                        <td style={{ padding: "16px", textAlign: "right" }}>{show(report.total.gst5Pct)}</td>
                        <td style={{ padding: "16px", textAlign: "right" }}>{show(report.total.salesInclGst || report.total.salesAfterFailed)}</td>
                        <td style={{ padding: "16px", textAlign: "right", color: "#c0392b" }}>-{show(report.total.commission || report.total.commissionInclGst)}</td>
                        <td style={{ padding: "16px", textAlign: "right", color: "var(--forest)" }}>{show(report.total.expectedPayout || report.total.expectedReceipt)}</td>
                        {report.hasBank && <td style={{ padding: "16px", textAlign: "right" }}>{show(report.total.bankActual)}</td>}
                        {report.hasBank && <td style={{ padding: "16px", textAlign: "right" }}>{show(report.total.bankDiff)}</td>}
                        <td style={{ padding: "16px", textAlign: "center" }}>
                          <span style={{ display: "inline-block", padding: "3px 10px", borderRadius: 20, fontSize: "11px", fontWeight: 700, background: report.total.bankMatched ? "#edf8f0" : "#feeceb", color: report.total.bankMatched ? "#1a6f3b" : "#a43024" }}>
                            {report.total.bankMatched ? "ALL MATCHED" : "DISCREPANCY"}
                          </span>
                        </td>
                      </>
                    )}
                  </tr>
                </tbody>
              </table>
            </div>
          </section>

          {/* Cashflow Statement */}
          <LineItemBreakdownTable
            title={`${report.platform} Cash Flow Statement`}
            subtitle="LINE-BY-LINE CASHFLOW"
            report={report}
            rowDefs={getCashflowRowDefs(report)}
          />

          {/* Profit Statement */}
          <LineItemBreakdownTable
            title={`${report.platform} Profit Statement`}
            subtitle="PROFIT & TAX ADJUSTMENTS"
            report={report}
            rowDefs={getProfitRowDefs(report)}
          />

          {/* Discrepancies */}
          <DiscrepancyTable report={report} />
        </div>
      )}

      {/* POS Extractor Report */}
      {!loading && posReport && (
        <div style={{ display: "flex", flexDirection: "column", gap: 24 }}>
          <section className="run" style={{ background: "var(--forest)", color: "#fff", display: "flex", justifyContent: "space-between", alignItems: "center", borderRadius: 12, padding: "18px 24px" }}>
            <div>
              <strong style={{ fontSize: "1.2rem" }}>POS Sales Channel Breakdown</strong>
              <small style={{ display: "block", opacity: 0.8, marginTop: 4 }}>File: {posReport.fileName} · Total Extracted: {show(posReport.grandTotal)}</small>
            </div>
            <button onClick={resetAll} style={{ background: "transparent", border: "1px solid rgba(255,255,255,0.4)", color: "#fff", padding: "8px 14px", borderRadius: 6, cursor: "pointer", fontWeight: 600 }}>
              🔄 New POS Extraction
            </button>
          </section>

          <section className="panel" style={{ background: "#fff", borderRadius: 12, border: "1px solid var(--line)", overflow: "hidden" }}>
            <div style={{ padding: "20px 24px", borderBottom: "1px solid var(--line)" }}>
              <p className="eyebrow">PAYMENT TYPE & WEEKLY MATRIX</p>
              <h2 style={{ margin: 0 }}>Channel Sales Matrix</h2>
            </div>
            <div style={{ overflowX: "auto" }}>
              <table style={{ width: "100%", borderCollapse: "collapse", fontSize: "0.85rem" }}>
                <thead>
                  <tr style={{ background: "#f7fbf8", borderBottom: "2px solid var(--line)" }}>
                    <th style={{ padding: "12px 16px", textAlign: "left" }}>Payment Channel</th>
                    {posReport.weekRanges.map((w) => (
                      <th key={w.weekNum} style={{ padding: "12px 16px", textAlign: "right" }}>{w.label}</th>
                    ))}
                    <th style={{ padding: "12px 16px", textAlign: "right", fontWeight: 700 }}>Total Sales</th>
                    <th style={{ padding: "12px 16px", textAlign: "right" }}>Orders</th>
                  </tr>
                </thead>
                <tbody>
                  {posReport.channels.map((ch, idx) => (
                    <tr key={idx} style={{ borderBottom: "1px solid var(--line)" }}>
                      <td style={{ padding: "14px 16px", fontWeight: 700 }}>{ch.channel}</td>
                      {ch.weeks.map((w, wIdx) => (
                        <td key={wIdx} style={{ padding: "14px 16px", textAlign: "right" }}>{show(w.amount)}</td>
                      ))}
                      <td style={{ padding: "14px 16px", textAlign: "right", fontWeight: 700, color: "var(--forest)" }}>{show(ch.total)}</td>
                      <td style={{ padding: "14px 16px", textAlign: "right", color: "var(--muted)" }}>{ch.totalOrders}</td>
                    </tr>
                  ))}
                  <tr style={{ background: "#f0f6f2", fontWeight: 700, borderTop: "2px solid var(--forest)" }}>
                    <td style={{ padding: "16px" }}>Grand Total</td>
                    {posReport.weekRanges.map((w, idx) => {
                      const weekSum = posReport.channels.reduce((s, c) => s + (c.weeks[idx]?.amount || 0), 0);
                      return <td key={idx} style={{ padding: "16px", textAlign: "right" }}>{show(weekSum)}</td>;
                    })}
                    <td style={{ padding: "16px", textAlign: "right", color: "var(--forest)" }}>{show(posReport.grandTotal)}</td>
                    <td style={{ padding: "16px", textAlign: "right" }}>{posReport.grandOrders}</td>
                  </tr>
                </tbody>
              </table>
            </div>
          </section>
        </div>
      )}

      {/* Deliverables Zip Report */}
      {!loading && zipReport && (
        <div style={{ display: "flex", flexDirection: "column", gap: 24 }}>
          <section className="run" style={{ background: "var(--forest)", color: "#fff", display: "flex", justifyContent: "space-between", alignItems: "center", borderRadius: 12, padding: "18px 24px" }}>
            <div>
              <strong style={{ fontSize: "1.2rem" }}>Deliverables Zip Auto-Audit Completed</strong>
              <small style={{ display: "block", opacity: 0.8, marginTop: 4 }}>
                {zipReport.classified.swiggyRaw.length} Swiggy · {zipReport.classified.zomatoRaw.length} Zomato · {zipReport.classified.pos.length} POS Files Cross-Checked
              </small>
            </div>
            <button onClick={resetAll} style={{ background: "transparent", border: "1px solid rgba(255,255,255,0.4)", color: "#fff", padding: "8px 14px", borderRadius: 6, cursor: "pointer", fontWeight: 600 }}>
              🔄 Audit Another Zip
            </button>
          </section>

          {/* Swiggy Audit Summary */}
          {zipReport.swiggy && zipReport.swiggy.weeks.length > 0 && (
            <section className="panel" style={{ background: "#fff", borderRadius: 12, border: "1px solid var(--line)", overflow: "hidden" }}>
              <div style={{ padding: "20px 24px", borderBottom: "1px solid var(--line)" }}>
                <p className="eyebrow">SWIGGY RECONCILIATION AUDIT</p>
                <h2 style={{ margin: 0 }}>Swiggy Accountant vs Calculated</h2>
              </div>
              <div style={{ overflowX: "auto" }}>
                <table style={{ width: "100%", borderCollapse: "collapse", fontSize: "0.85rem" }}>
                  <thead>
                    <tr style={{ background: "#f7fbf8", borderBottom: "2px solid var(--line)" }}>
                      <th style={{ padding: "12px 16px", textAlign: "left" }}>Week</th>
                      <th style={{ padding: "12px 16px", textAlign: "right" }}>Accountant Sales</th>
                      <th style={{ padding: "12px 16px", textAlign: "right" }}>Calculated Sales</th>
                      <th style={{ padding: "12px 16px", textAlign: "right" }}>Accountant Comm.</th>
                      <th style={{ padding: "12px 16px", textAlign: "right" }}>Calculated Comm.</th>
                      <th style={{ padding: "12px 16px", textAlign: "right" }}>Accountant Payout</th>
                      <th style={{ padding: "12px 16px", textAlign: "right" }}>Calculated Payout</th>
                      <th style={{ padding: "12px 16px", textAlign: "center" }}>Status</th>
                    </tr>
                  </thead>
                  <tbody>
                    {zipReport.swiggy.weeks.map((w, idx) => {
                      const bad = w.discrepancy.sales !== 0 || w.discrepancy.commission !== 0 || w.discrepancy.payout !== 0;
                      return (
                        <tr key={idx} style={{ borderBottom: "1px solid var(--line)", background: bad ? "#fff8f8" : "transparent" }}>
                          <td style={{ padding: "14px 16px", fontWeight: 700 }}>{w.label}</td>
                          <td style={{ padding: "14px 16px", textAlign: "right" }}>{show(w.accountant.sales)}</td>
                          <td style={{ padding: "14px 16px", textAlign: "right", color: w.discrepancy.sales !== 0 ? "#c0392b" : "inherit" }}>{show(w.calculated.sales)}</td>
                          <td style={{ padding: "14px 16px", textAlign: "right" }}>{show(w.accountant.commission)}</td>
                          <td style={{ padding: "14px 16px", textAlign: "right", color: w.discrepancy.commission !== 0 ? "#c0392b" : "inherit" }}>{show(w.calculated.commission)}</td>
                          <td style={{ padding: "14px 16px", textAlign: "right" }}>{show(w.accountant.payout)}</td>
                          <td style={{ padding: "14px 16px", textAlign: "right", color: w.discrepancy.payout !== 0 ? "#c0392b" : "inherit" }}>{show(w.calculated.payout)}</td>
                          <td style={{ padding: "14px 16px", textAlign: "center" }}>
                            <span style={{ display: "inline-block", padding: "3px 10px", borderRadius: 20, fontSize: "11px", fontWeight: 700, background: bad ? "#feeceb" : "#edf8f0", color: bad ? "#a43024" : "#1a6f3b" }}>
                              {bad ? "Discrepancy" : "Match"}
                            </span>
                          </td>
                        </tr>
                      );
                    })}
                  </tbody>
                </table>
              </div>
            </section>
          )}

          {/* Zomato Audit Summary */}
          {zipReport.zomato && zipReport.zomato.weeks.length > 0 && (
            <section className="panel" style={{ background: "#fff", borderRadius: 12, border: "1px solid var(--line)", overflow: "hidden" }}>
              <div style={{ padding: "20px 24px", borderBottom: "1px solid var(--line)" }}>
                <p className="eyebrow">ZOMATO RECONCILIATION AUDIT</p>
                <h2 style={{ margin: 0 }}>Zomato Accountant vs Calculated</h2>
              </div>
              <div style={{ overflowX: "auto" }}>
                <table style={{ width: "100%", borderCollapse: "collapse", fontSize: "0.85rem" }}>
                  <thead>
                    <tr style={{ background: "#f7fbf8", borderBottom: "2px solid var(--line)" }}>
                      <th style={{ padding: "12px 16px", textAlign: "left" }}>Week</th>
                      <th style={{ padding: "12px 16px", textAlign: "right" }}>Accountant Sales</th>
                      <th style={{ padding: "12px 16px", textAlign: "right" }}>Calculated Sales</th>
                      <th style={{ padding: "12px 16px", textAlign: "right" }}>Accountant Comm.</th>
                      <th style={{ padding: "12px 16px", textAlign: "right" }}>Calculated Comm.</th>
                      <th style={{ padding: "12px 16px", textAlign: "right" }}>Accountant Payout</th>
                      <th style={{ padding: "12px 16px", textAlign: "right" }}>Calculated Payout</th>
                      <th style={{ padding: "12px 16px", textAlign: "center" }}>Status</th>
                    </tr>
                  </thead>
                  <tbody>
                    {zipReport.zomato.weeks.map((w, idx) => {
                      const bad = w.discrepancy.sales !== 0 || w.discrepancy.commission !== 0 || w.discrepancy.payout !== 0;
                      return (
                        <tr key={idx} style={{ borderBottom: "1px solid var(--line)", background: bad ? "#fff8f8" : "transparent" }}>
                          <td style={{ padding: "14px 16px", fontWeight: 700 }}>{w.label}</td>
                          <td style={{ padding: "14px 16px", textAlign: "right" }}>{show(w.accountant.sales)}</td>
                          <td style={{ padding: "14px 16px", textAlign: "right", color: w.discrepancy.sales !== 0 ? "#c0392b" : "inherit" }}>{show(w.calculated.sales)}</td>
                          <td style={{ padding: "14px 16px", textAlign: "right" }}>{show(w.accountant.commission)}</td>
                          <td style={{ padding: "14px 16px", textAlign: "right", color: w.discrepancy.commission !== 0 ? "#c0392b" : "inherit" }}>{show(w.calculated.commission)}</td>
                          <td style={{ padding: "14px 16px", textAlign: "right" }}>{show(w.accountant.payout)}</td>
                          <td style={{ padding: "14px 16px", textAlign: "right", color: w.discrepancy.payout !== 0 ? "#c0392b" : "inherit" }}>{show(w.calculated.payout)}</td>
                          <td style={{ padding: "14px 16px", textAlign: "center" }}>
                            <span style={{ display: "inline-block", padding: "3px 10px", borderRadius: 20, fontSize: "11px", fontWeight: 700, background: bad ? "#feeceb" : "#edf8f0", color: bad ? "#a43024" : "#1a6f3b" }}>
                              {bad ? "Discrepancy" : "Match"}
                            </span>
                          </td>
                        </tr>
                      );
                    })}
                  </tbody>
                </table>
              </div>
            </section>
          )}
        </div>
      )}

      {/* Input Form (When no report is active) */}
      {!loading && !report && !posReport && !zipReport && (
        <div style={{ background: "#fff", borderRadius: 12, border: "1px solid var(--line)", padding: 32 }}>
          <div style={{ marginBottom: 28 }}>
            <p className="eyebrow" style={{ color: "var(--forest)", fontWeight: 700, letterSpacing: "0.1em" }}>
              {platform === "pos" ? "POS SALES CHANNEL EXTRACTOR" : platform === "zip" ? "DELIVERABLES RECONCILIATION ARCHIVE" : `${platform.toUpperCase()} AUTOMATED RECONCILIATION`}
            </p>
            <h2 style={{ fontSize: "1.6rem", margin: "4px 0 8px" }}>
              {platform === "zomato" && "Zomato Payout & Settlement Audit"}
              {platform === "swiggy" && "Swiggy Order Level Reconciliation"}
              {platform === "dineout" && "Swiggy Dineout Settlement Audit"}
              {platform === "zpay" && "Zomato Pay & Ads Reconciliation"}
              {platform === "paytm" && "Paytm Settlement & MDR Audit"}
              {platform === "pos" && "Petpooja / Posist Channel Breakdown"}
              {platform === "zip" && "Auto-Audit Entire Deliverables Folder"}
            </h2>
            <p style={{ color: "var(--muted)", fontSize: "0.9rem", lineHeight: 1.6, margin: 0 }}>
              Upload your raw transaction files and bank statement. All formulas and figures are computed instantly on-screen so you can review and compare without needing to download.
            </p>
          </div>

          {/* Form Grid */}
          <div style={{ display: "flex", flexDirection: "column", gap: 24 }}>
            {/* Row 1: Client Name, Month, Mode */}
            {platform !== "pos" && platform !== "zip" && (
              <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(220px, 1fr))", gap: 16 }}>
                <div>
                  <label style={{ display: "block", fontSize: "0.78rem", fontWeight: 700, color: "var(--muted)", marginBottom: 6 }}>CLIENT / RESTAURANT NAME</label>
                  <input
                    type="text"
                    value={clientName}
                    onChange={(e) => setClientName(e.target.value)}
                    placeholder="e.g. Kailash Parbat"
                    style={{ width: "100%", padding: "10px 14px", borderRadius: 8, border: "1px solid var(--line)", fontSize: "0.9rem" }}
                  />
                </div>
                <div>
                  <label style={{ display: "block", fontSize: "0.78rem", fontWeight: 700, color: "var(--muted)", marginBottom: 6 }}>MONTH</label>
                  <select
                    value={month}
                    onChange={(e) => setMonth(e.target.value)}
                    style={{ width: "100%", padding: "10px 14px", borderRadius: 8, border: "1px solid var(--line)", fontSize: "0.9rem", background: "#fff" }}
                  >
                    {months.map((m) => (
                      <option key={m} value={m}>{m}</option>
                    ))}
                  </select>
                </div>
                {platform === "zomato" && (
                  <div>
                    <label style={{ display: "block", fontSize: "0.78rem", fontWeight: 700, color: "var(--muted)", marginBottom: 6 }}>RECONCILIATION MODE</label>
                    <div style={{ display: "flex", gap: 8 }}>
                      <button
                        type="button"
                        onClick={() => setZomatoMode("weekly")}
                        style={{
                          flex: 1, padding: "10px", borderRadius: 8, fontSize: "0.82rem", fontWeight: 700, cursor: "pointer",
                          background: zomatoMode === "weekly" ? "var(--forest)" : "#f5f7f5",
                          color: zomatoMode === "weekly" ? "#fff" : "var(--text)",
                          border: "1px solid var(--line)",
                        }}
                      >
                        WEEKLY
                      </button>
                      <button
                        type="button"
                        onClick={() => setZomatoMode("consolidated")}
                        style={{
                          flex: 1, padding: "10px", borderRadius: 8, fontSize: "0.82rem", fontWeight: 700, cursor: "pointer",
                          background: zomatoMode === "consolidated" ? "var(--forest)" : "#f5f7f5",
                          color: zomatoMode === "consolidated" ? "#fff" : "var(--text)",
                          border: "1px solid var(--line)",
                        }}
                      >
                        CONSOLIDATED
                      </button>
                    </div>
                  </div>
                )}
              </div>
            )}

            {/* Row 2: Week Ranges */}
            {platform !== "zip" && (
              <div style={{ background: "#fbfcfb", border: "1px solid var(--line)", borderRadius: 10, padding: 18 }}>
                <div style={{ fontSize: "0.8rem", fontWeight: 700, color: "var(--forest)", marginBottom: 12 }}>
                  📅 WEEK DATE RANGES (CUSTOMIZE IF NEEDED)
                </div>
                <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(180px, 1fr))", gap: 16 }}>
                  <div>
                    <label style={{ display: "block", fontSize: "0.74rem", color: "var(--muted)", marginBottom: 4 }}>First Week Start (Day)</label>
                    <input type="number" min="1" max="31" value={fStart} onChange={(e) => setFStart(Number(e.target.value))} style={{ width: "100%", padding: "8px 12px", borderRadius: 6, border: "1px solid var(--line)" }} />
                  </div>
                  <div>
                    <label style={{ display: "block", fontSize: "0.74rem", color: "var(--muted)", marginBottom: 4 }}>First Week End (Day)</label>
                    <input type="number" min="1" max="31" value={fEnd} onChange={(e) => setFEnd(Number(e.target.value))} style={{ width: "100%", padding: "8px 12px", borderRadius: 6, border: "1px solid var(--line)" }} />
                  </div>
                  <div>
                    <label style={{ display: "block", fontSize: "0.74rem", color: "var(--muted)", marginBottom: 4 }}>Last Week Start (Day)</label>
                    <input type="number" min="1" max="31" value={lStart} onChange={(e) => setLStart(Number(e.target.value))} style={{ width: "100%", padding: "8px 12px", borderRadius: 6, border: "1px solid var(--line)" }} />
                  </div>
                  <div>
                    <label style={{ display: "block", fontSize: "0.74rem", color: "var(--muted)", marginBottom: 4 }}>Last Week End (Day)</label>
                    <input type="number" min="1" max="31" value={lEnd} onChange={(e) => setLEnd(Number(e.target.value))} style={{ width: "100%", padding: "8px 12px", borderRadius: 6, border: "1px solid var(--line)" }} />
                  </div>
                </div>
              </div>
            )}

            {/* Row 3: Upload Dropzones */}
            <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(280px, 1fr))", gap: 18 }}>
              {/* Primary file input */}
              {platform === "pos" ? (
                <div>
                  <label style={{ display: "block", fontSize: "0.78rem", fontWeight: 700, color: "var(--muted)", marginBottom: 6 }}>POS SALES REPORT (.xlsx / .xls / .csv)</label>
                  <label className="upload" style={{ minHeight: 160, border: "2px dashed var(--line)", background: posFile ? "#edf8f0" : undefined }}>
                    <input type="file" accept=".xlsx,.xls,.csv" onChange={(e) => setPosFile(e.target.files[0])} />
                    <span style={{ fontSize: "1.3rem" }}>📊</span>
                    <strong>{posFile ? posFile.name : "Attach POS Report"}</strong>
                    <small>Petpooja or Posist payment wise export</small>
                  </label>
                </div>
              ) : platform === "zip" ? (
                <div>
                  <label style={{ display: "block", fontSize: "0.78rem", fontWeight: 700, color: "var(--muted)", marginBottom: 6 }}>DELIVERABLES ARCHIVE (.zip)</label>
                  <label className="upload" style={{ minHeight: 160, border: "2px dashed var(--line)", background: zipFile ? "#edf8f0" : undefined }}>
                    <input type="file" accept=".zip" onChange={(e) => setZipFile(e.target.files[0])} />
                    <span style={{ fontSize: "1.3rem" }}>📦</span>
                    <strong>{zipFile ? zipFile.name : "Attach Deliverables Zip"}</strong>
                    <small>Auto-extracts POS, Swiggy, Zomato & summaries</small>
                  </label>
                </div>
              ) : (
                <div>
                  <label style={{ display: "block", fontSize: "0.78rem", fontWeight: 700, color: "var(--muted)", marginBottom: 6 }}>
                    {platform.toUpperCase()} WEEKLY INVOICES (.xlsx / .xls / .csv)
                  </label>
                  <label className="upload" style={{ minHeight: 160, border: "2px dashed var(--line)", background: invoiceFiles.length > 0 ? "#edf8f0" : undefined }}>
                    <input
                      type="file"
                      multiple
                      accept=".xlsx,.xls,.csv"
                      onChange={(e) => {
                        if (e.target.files) setInvoiceFiles(Array.from(e.target.files));
                      }}
                    />
                    <span style={{ fontSize: "1.3rem" }}>📄</span>
                    <strong>{invoiceFiles.length > 0 ? `${invoiceFiles.length} Invoices Attached` : `Drop ${platform} Invoices`}</strong>
                    <small>Select all weekly settlement/invoice files</small>
                  </label>
                  {invoiceFiles.length > 0 && (
                    <div style={{ marginTop: 8, display: "flex", flexWrap: "wrap", gap: 6 }}>
                      {invoiceFiles.map((f, i) => (
                        <span key={i} style={{ background: "#f0f2f0", padding: "3px 8px", borderRadius: 4, fontSize: "0.74rem", color: "var(--text)" }}>
                          ✓ {f.name}
                        </span>
                      ))}
                    </div>
                  )}
                </div>
              )}

              {/* Optional Bank Statement input */}
              {platform !== "pos" && platform !== "zip" && (
                <div>
                  <label style={{ display: "block", fontSize: "0.78rem", fontWeight: 700, color: "var(--muted)", marginBottom: 6 }}>
                    BANK STATEMENT <span style={{ color: "var(--forest)" }}>(OPTIONAL FOR VERIFICATION)</span>
                  </label>
                  <label className="upload" style={{ minHeight: 160, border: "2px dashed var(--line)", background: bankFile ? "#edf8f0" : undefined }}>
                    <input type="file" accept=".xlsx,.xls,.csv" onChange={(e) => setBankFile(e.target.files[0])} />
                    <span style={{ fontSize: "1.3rem" }}>🏦</span>
                    <strong>{bankFile ? bankFile.name : "Attach Bank Statement"}</strong>
                    <small>Verifies actual deposit amounts vs expected payout</small>
                  </label>
                  {bankFile && (
                    <button
                      type="button"
                      onClick={() => setBankFile(null)}
                      style={{ marginTop: 6, background: "none", border: 0, color: "#c0392b", fontSize: "0.75rem", cursor: "pointer" }}
                    >
                      ✕ Remove bank file
                    </button>
                  )}
                </div>
              )}

              {/* Zomato Pay Ads File */}
              {platform === "zpay" && (
                <div>
                  <label style={{ display: "block", fontSize: "0.78rem", fontWeight: 700, color: "var(--muted)", marginBottom: 6 }}>
                    ZOMATO PAY ADS REPORT <span style={{ color: "var(--forest)" }}>(OPTIONAL)</span>
                  </label>
                  <label className="upload" style={{ minHeight: 160, border: "2px dashed var(--line)", background: adsFiles.length > 0 ? "#edf8f0" : undefined }}>
                    <input type="file" multiple accept=".xlsx,.xls,.csv" onChange={(e) => e.target.files && setAdsFiles(Array.from(e.target.files))} />
                    <span style={{ fontSize: "1.3rem" }}>📢</span>
                    <strong>{adsFiles.length > 0 ? `${adsFiles.length} Ads Files Attached` : "Attach Zpay Ads Export"}</strong>
                    <small>Deducts marketing & ad campaign spend</small>
                  </label>
                </div>
              )}
            </div>

            {error && (
              <p className="error" style={{ color: "var(--red)", marginTop: 8, padding: "10px 14px", background: "#feeceb", borderRadius: 8 }}>
                {error}
              </p>
            )}

            {/* Submit Button */}
            <div style={{ marginTop: 8 }}>
              <button
                type="button"
                onClick={handleExecute}
                style={{
                  width: "100%",
                  padding: "16px 24px",
                  borderRadius: 10,
                  border: 0,
                  background: "var(--forest)",
                  color: "#fff",
                  fontWeight: 700,
                  fontSize: "1rem",
                  letterSpacing: "0.05em",
                  cursor: "pointer",
                  boxShadow: "0 4px 14px rgba(20,63,53,0.25)",
                  transition: "0.2s",
                }}
              >
                EXECUTE {platform.toUpperCase()} RECONCILIATION
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}
