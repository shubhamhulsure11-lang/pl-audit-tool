"use client";
import { useMemo, useState } from "react";
import * as XLSX from "xlsx";

const inr = new Intl.NumberFormat("en-IN", { style: "currency", currency: "INR", maximumFractionDigits: 0 });
const clean = (v) => String(v ?? "").trim().toLowerCase().replace(/[^a-z0-9]/g, "");
const n = (v) => typeof v === "number" ? v : Number(String(v ?? "").replace(/[,₹]/g, "")) || 0;
const show = (v) => inr.format(v || 0);
const col = (heads, choices) => choices.map(clean).map(x => heads.map(clean).indexOf(x)).find(x => x >= 0) ?? -1;

// Restaurant-specific Account Head Taxonomy & Keyword Rules
const RESTAURANT_TAXONOMY = [
  {
    id: "groceries",
    label: "Groceries & Dry Goods",
    aliases: ["grocer", "provision", "dry good", "food and grocery", "food & grocery", "raw material", "kitchen raw", "staple", "spices", "grain"],
    keywords: [
      "rice", "basmati", "kolam", "sona masoori", "wheat", "atta", "maida", "sooji", "semolina", "besan", "cornflour", "corn flour", "cornstarch", "custard powder",
      "dal", "lentil", "toor", "tur dal", "moong", "mung", "urad", "chana", "kabuli", "rajma", "soya", "poha", "vermicelli", "sevai", "noodles", "pasta", "macaroni", "spaghetti", "crouton", "breadcrumbs",
      "spices", "masala", "garam masala", "chaat masala", "biryani masala", "kitchen king", "turmeric", "haldi", "jeera", "cumin", "dhania", "coriander seed", "coriander powder", "mustard seed", "rai", "sarson", "fenugreek", "methi seed", "fennel", "saunf", "cardamom", "elaichi", "clove", "laung", "cinnamon", "dalchini", "star anise", "bay leaf", "tej patta", "black pepper", "kali mirch", "white pepper", "nutmeg", "saffron", "kesar", "ajwain", "kalonji", "hing", "asafoetida", "kasuri methi", "red chilli", "kashmiri chilli", "degi mirch", "paprika", "chili flakes", "oregano", "thyme", "rosemary",
      "oil", "refined oil", "sunflower oil", "mustard oil", "groundnut oil", "peanut oil", "sesame oil", "til oil", "olive oil", "canola oil", "soybean oil", "palm oil", "vanaspati", "dalda",
      "ghee", "desi ghee", "cow ghee", "buffalo ghee", "amul ghee", // Rule: Ghee is Groceries
      "sugar", "brown sugar", "jaggery", "gur", "honey", "salt", "rock salt", "black salt", "sendha namak", "pink salt", "baking soda", "baking powder", "yeast", "citric acid", "ajinomoto", "msg",
      "sauce", "ketchup", "tomato ketchup", "chilli sauce", "soya sauce", "dark soy", "vinegar", "sriracha", "tabasco", "schezwan", "mayonnaise", "mayo", "mustard paste", "salsa", "peri peri", "pickle", "achaar", "murabba", "chutney", "papad", "appalam",
      "tamarind", "imli", "desiccated coconut", "coconut dry",
      "kaju", "cashew", "badam", "almond", "kismis", "raisin", "pista", "pistachio", "walnut", "akhrot", "dates", "khajoor", "melon seeds", "poppy seeds", "khus khus",
      "cocoa powder", "cooking chocolate", "chocolate chips", "vanilla essence", "food color"
    ]
  },
  {
    id: "dairy",
    label: "Dairy",
    aliases: ["dairy", "milk product"],
    keywords: [
      "milk", "toned milk", "full cream milk", "cow milk", "buffalo milk", "paneer", "cottage cheese", "curd", "dahi", "yogurt", "yoghurt", "hung curd", "fresh cream", "amul cream", "sour cream", "whipped cream", "butter", "salted butter", "unsalted butter", "table butter", "cheese", "mozzarella", "cheddar", "cheese slice", "cheese block", "parmesan", "feta", "gouda", "cream cheese", "khoya", "mawa", "buttermilk", "chaas", "lassi"
    ]
  },
  {
    id: "poultry_eggs",
    label: "Eggs & Poultry",
    aliases: ["poultry", "egg", "chicken"],
    keywords: [
      "egg", "eggs", "brown eggs", "white eggs", "quail egg", "chicken", "broiler", "country chicken", "desi chicken", "boneless chicken", "chicken breast", "chicken leg", "chicken drumstick", "chicken wings", "chicken keema", "chicken mince", "chicken curry cut", "whole chicken", "chicken lollipop", "chicken liver"
    ]
  },
  {
    id: "meat_seafood",
    label: "Meat & Seafood",
    aliases: ["meat", "mutton", "seafood", "fish", "pork", "beef", "lamb", "prawn", "crab"],
    keywords: [
      "mutton", "lamb", "goat meat", "mutton keema", "mutton chops", "beef", "pork", "bacon", "ham", "sausage", "pepperoni", "fish", "surmai", "kingfish", "pomfret", "rawas", "salmon", "rohu", "katla", "tilapia", "basa", "basa fillet", "tuna", "mackerel", "bangda", "hilsa", "prawns", "prawn", "shrimp", "tiger prawn", "jumbo prawn", "crab", "lobster", "squid", "calamari", "octopus", "clams", "oysters"
    ]
  },
  {
    id: "vegetables",
    label: "Fresh Vegetables",
    aliases: ["vegetable", "fresh veg", "veggie", "greens", "sabzi"],
    keywords: [
      "onion", "pyaz", "potato", "aloo", "tomato", "tamatar", "ginger", "adrak", "garlic", "lahsun", "green chilli", "hari mirch", "capsicum", "bell pepper", "shimla mirch", "carrot", "gajar", "beans", "green peas", "matar", "cabbage", "cauliflower", "gobhi", "broccoli", "spinach", "palak", "methi leaves", "coriander leaves", "fresh dhania", "mint leaves", "pudina", "curry leaves", "kadi patta", "lettuce", "iceberg", "cucumber", "kheera", "beetroot", "radish", "mooli", "spring onion", "leek", "celery", "zucchini", "mushroom", "button mushroom", "baby corn", "sweet corn", "bhindi", "lady finger", "okra", "brinjal", "baingan", "eggplant", "bottle gourd", "lauki", "bitter gourd", "karela", "pumpkin", "kaddu", "lemon fresh", "nimbu fresh", "raw banana", "drumstick"
    ]
  },
  {
    id: "fruits",
    label: "Fresh Fruits",
    aliases: ["fruit", "fresh fruit"],
    keywords: [
      "apple", "seb", "banana", "kela", "orange", "santra", "mosambi", "sweet lime", "pomegranate", "anar", "watermelon", "tarbooj", "muskmelon", "kharbuj", "papaya", "papita", "pineapple", "ananas", "mango fresh", "aam", "grapes", "angoor", "strawberry", "kiwi", "guava", "amrood", "pear", "chikoo", "blueberry fresh", "dragonfruit", "plum", "peach", "cherry"
    ]
  },
  {
    id: "beverages",
    label: "Beverages",
    aliases: ["beverage", "drink", "bar purchase", "liquor", "soft drink", "mocktail", "cocktail"],
    keywords: [
      "cola", "coca cola", "coke", "diet coke", "pepsi", "7up", "sprite", "thums up", "limca", "fanta", "mirinda", "mountain dew",
      "redbull", "red bull", "monster energy", "sting", "tonic water", "ginger ale", "gingerale", "club soda", "soda water", "lehar soda", "kinley soda",
      "juice", "real juice", "real apple", "real mango", "real orange", "real cranberry", "real pineapple", "real litchi", "tropicana", "minute maid", "frooti", "maaza", "slice", "appy", "raw pressery", "paper boat",
      "packaged water", "mineral water", "bisleri", "kinley", "aquafina", "vedica", "himalayan",
      "syrup", "monin", "malas", "malass", "blue curacao", "grenadine", "mojito syrup", "peach syrup", "watermelon syrup", "crush", "strawberry crush", "blueberry crush", "kiwi crush", "litchi crush", "orange crush", "pineapple crush",
      "beer", "kingfisher", "budweiser", "carlsberg", "bira", "corona", "heineken", "tuborg", "breezer",
      "whisky", "whiskey", "scotch", "bourbon", "rum", "old monk", "bacardi", "vodka", "absolut", "smirnoff", "magic moments", "gin", "bombay sapphire", "tequila", "brandy", "wine", "sula", "champagne"
    ]
  },
  {
    id: "fuel_gas",
    label: "Fuel & Gas",
    aliases: ["fuel", "gas", "lpg", "cng", "diesel", "petrol"],
    keywords: [
      "lpg", "commercial cylinder", "lpg cylinder", "19kg cylinder", "bharat gas", "indane", "hp gas", "png", "diesel", "petrol", "kerosene", "firewood"
    ]
  },
  {
    id: "other_purchases",
    label: "Other Purchases",
    aliases: ["other purchase", "misc", "miscellaneous", "general purchase", "kitchen consumable", "other expense"],
    keywords: [
      "charcoal", "wood charcoal", "coal", // Rule: Charcoal is Other Purchases
      "ice", "ice cube", "ice cubes", "ice slab", "crushed ice", "dry ice", // Rule: Ice is Other Purchases
      "wooden skewers", "bamboo skewers", "skewers", "toothpick", "toothpicks", "birthday candle", "lighter", "matchbox"
    ]
  },
  {
    id: "packaging",
    label: "Packaging & Disposables",
    aliases: ["packag", "disposab", "takeaway", "parcel", "container"],
    keywords: [
      "container", "plastic container", "meal tray", "meal box", "500ml container", "750ml container", "1000ml container", "aluminium container", "foil container", "burger box", "pizza box", "cake box", "sweet box",
      "paper bag", "carry bag", "kraft paper bag", "non woven bag", "d-cut bag", "zip lock", "polythene",
      "paper plate", "disposable plate", "paper cup", "plastic glass", "disposable glass",
      "tissue", "paper napkin", "cocktail napkin", "tissue roll", "kitchen roll", "toilet roll",
      "aluminium foil", "silver foil", "cling wrap", "cling film", "butter paper", "parchment paper",
      "paper straw", "plastic straw", "wooden spoon", "plastic spoon", "wooden fork", "plastic fork", "disposable cutlery", "chopstick"
    ]
  },
  {
    id: "cleaning",
    label: "Cleaning & Housekeeping",
    aliases: ["clean", "housekeep", "sanit", "hygiene", "detergent"],
    keywords: [
      "dishwash", "dishwash bar", "dishwash liquid", "vim", "exo", "pril", "surf", "ariel", "tide", "rin", "detergent",
      "floor cleaner", "lizol", "phenyl", "colin", "glass cleaner", "harpic", "toilet cleaner", "bathroom cleaner", "drain cleaner", "caustic soda",
      "bleach", "bleaching powder", "disinfectant", "sanitizer", "hand wash", "lifebuoy", "dettol", "savlon", "liquid soap",
      "broom", "jhadu", "mop", "floor wiper", "duster", "cleaning cloth", "microfiber cloth", "sponge", "scrubber", "scotch brite", "garbage bag", "trash bag", "dustbin cover", "dust pan", "rubber gloves", "room freshener", "odonil"
    ]
  },
  {
    id: "kitchen_tools",
    label: "Kitchen & Bar Equipment",
    aliases: ["utensil", "kitchen tool", "crockery", "cutlery", "equipment", "bar tool"],
    keywords: [
      "kadai", "fry pan", "sauce pan", "tawa", "cooker", "pressure cooker", "patila", "strainer", "colander", "ladle", "karchi", "chef knife", "chopping knife", "peeler", "grater", "chopping board", "tongs", "chimta", "whisk", "rolling pin", "belan", "chakla", "mixing bowl", "baking tray", "sizzler plate",
      "cocktail shaker", "bar strainer", "jigger", "muddler", "bar spoon", "pourer", "corkscrew", "bottle opener"
    ]
  },
  {
    id: "stationery",
    label: "Stationery & Office",
    aliases: ["station", "office supplies", "printing", "paper"],
    keywords: [
      "pen", "ball pen", "gel pen", "pencil", "marker", "permanent marker", "notebook", "register", "attendance register", "bill book", "kot book", "receipt book",
      "stapler", "stapler pin", "punch machine", "brown tape", "cello tape", "scissor", "stamp pad", "rubber band", "binder clip", "file", "folder", "envelope", "a4 paper", "pos roll", "billing roll", "printer cartridge", "toner"
    ]
  }
];

// Helper to clean item string for token search
function tokenizeItem(str) {
  return String(str || "").toLowerCase().replace(/[^a-z0-9\s]/g, " ").replace(/\s+/g, " ").trim();
}

// Check if actual account string matches category aliases
function accountMatchesCategory(actualAccount, category) {
  const normAcc = String(actualAccount || "").toLowerCase();
  return category.aliases.some(alias => normAcc.includes(alias));
}

// Find expected category for an item description
function classifyItem(itemName) {
  const norm = tokenizeItem(itemName);
  if (!norm) return null;

  let bestMatch = null;
  let maxKeywordLen = 0;

  for (const cat of RESTAURANT_TAXONOMY) {
    for (const kw of cat.keywords) {
      const normKw = tokenizeItem(kw);
      if (!normKw) continue;

      // Check exact word or substring match
      const regex = new RegExp(`(^|\\s)${normKw.replace(/[-\/\\^$*+?.()|[\]{}]/g, "\\$&")}(\\s|$)`, "i");
      const matched = regex.test(norm) || (normKw.length >= 4 && norm.includes(normKw));

      if (matched && normKw.length > maxKeywordLen) {
        maxKeywordLen = normKw.length;
        bestMatch = { category: cat, keyword: kw };
      }
    }
  }
  return bestMatch;
}

// Auditor: Detect items booked under wrong account heads
function detectMisclassifications(records) {
  if (!records || !records.length) return [];
  const map = new Map();

  records.forEach(r => {
    if (!r.item || !r.account || r.account === "Unassigned Account") return;
    const match = classifyItem(r.item);
    if (!match) return; // Unclassified items are not flagged

    const expectedCat = match.category;
    // Check if actual account head matches expected category
    const isCorrect = accountMatchesCategory(r.account, expectedCat);

    if (!isCorrect) {
      const key = `${r.item}:::${r.vendor}:::${r.account}:::${expectedCat.id}`;
      if (!map.has(key)) {
        map.set(key, {
          item: r.item,
          vendor: r.vendor,
          actualAccount: r.account,
          suggestedAccount: expectedCat.label,
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
    date: col(h, ["Bill Date", "Date"]),
    vendor: col(h, ["Vendor Name", "Vendor"]),
    bill: col(h, ["Bill Number", "Invoice Number"]),
    account: col(h, ["Account Name", "Account", "Expense Account", "Account Head", "Chart of Accounts"]),
    item: col(h, ["Item Name", "Item"]),
    qty: col(h, ["Quantity", "Qty"]),
    rate: col(h, ["Rate", "Item Rate"]),
    total: col(h, ["Item Total", "Line Item Total", "Total"]),
    branch: col(h, ["Branch Name", "Branch"])
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
  return { name: file.name, records, hasAccountCol: i.account >= 0 };
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

function MisclassificationsView({ items, onGoToPivot }) {
  const [q, setQ] = useState("");
  const filtered = useMemo(() => {
    const term = q.trim().toLowerCase();
    if (!term) return items;
    return items.filter(x =>
      x.item.toLowerCase().includes(term) ||
      x.vendor.toLowerCase().includes(term) ||
      x.actualAccount.toLowerCase().includes(term) ||
      x.suggestedAccount.toLowerCase().includes(term) ||
      x.matchedKeyword.toLowerCase().includes(term)
    );
  }, [items, q]);

  const totalExposure = useMemo(() => items.reduce((s, x) => s + x.total, 0), [items]);

  return (
    <section className="panel">
      <div className="panelhead">
        <div>
          <p className="eyebrow">RESTAURANT TAXONOMY AUDIT</p>
          <h2>Account Head Misclassifications</h2>
        </div>
        <div className="pivot-summary-badges">
          <b className="badge">{items.length} Flagged Items</b>
          <b className="badge accent-badge">{show(totalExposure)} Total Exposure</b>
        </div>
      </div>

      <div className="misclass-info-banner">
        <span>💡</span>
        <div>
          <strong>How this works:</strong> Items are matched against restaurant accounting rules (e.g., <em>Rice/Ghee &rarr; Groceries</em>, <em>Milk/Paneer &rarr; Dairy</em>, <em>Charcoal/Ice &rarr; Other Purchases</em>). Items placed in conflicting account heads are flagged below.
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
        {onGoToPivot && (
          <button className="pivot-btn" onClick={onGoToPivot}>
            View Account Pivot &rarr;
          </button>
        )}
      </div>

      <Table head={["Item Description", "Vendor", "Current Account Head", "Suggested Account Head", "Matched Rule", "Amount Exposure"]}>
        {filtered.length ? filtered.map((x, i) => (
          <tr key={i}>
            <td style={{ fontWeight: 600, color: "var(--ink)" }}>{x.item}</td>
            <td>{x.vendor}</td>
            <td><span className="pill-actual-acc">{x.actualAccount}</span></td>
            <td><span className="pill-suggested-acc">&rarr; {x.suggestedAccount}</span></td>
            <td><span className="rule-tag">matched "{x.matchedKeyword}"</span></td>
            <td style={{ fontWeight: 700, fontFamily: "var(--font-mono, monospace)" }}>{show(x.total)}</td>
          </tr>
        )) : (
          <tr>
            <td colSpan="6">
              <Empty>{items.length === 0 ? "🎉 No account head misclassifications detected! All items match their restaurant categories." : "No items match your search filter."}</Empty>
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

export default function Home() {
  const [current, setCurrent] = useState(null), [previous, setPrevious] = useState(null), [error, setError] = useState(""), [tab, setTab] = useState("Overview");
  const [thresholds, setThresholds] = useState({ vendor: 20, item: 25, price: 20 });
  const setT = (k, v) => setThresholds(t => ({ ...t, [k]: v }));
  const result = useMemo(() => current && previous ? analyse(current, previous, thresholds) : null, [current, previous, thresholds]);
  const upload = async (file, setter) => { try { setError(""); setter(await readFile(file)); } catch (e) { setError(e.message); } };
  const spend = x => x?.records.reduce((s, r) => s + r.total, 0) || 0;
  const total = result && (result.duplicates.length + result.vendors.length + result.items.length + result.prices.length + result.misclassifications.length);

  return (
    <main>
      <header>
        <div>
          <p className="eyebrow">FINANCE CONTROL CENTER</p>
          <h1>P&L Audit Desk</h1>
        </div>
        <p className="privacy"><i /> Local analysis - files stay on your device</p>
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
              <small>{current.name} compared with {previous.name}</small>
            </div>
            <button onClick={() => { setCurrent(null); setPrevious(null); setTab("Overview"); }}>New review</button>
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
            <MisclassificationsView items={result.misclassifications} onGoToPivot={() => setTab("Account heads")} />
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

