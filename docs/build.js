// SPDX-License-Identifier: Apache-2.0
/**
 * build.js — 電商品牌分析報告產生器（PptxGenJS，塔圖主題）
 * ------------------------------------------------------------------
 * 讀 deck_data.json（Claude 依 Commerce MCP 資料組成），產出精簡版報告：主體 14–17 頁，另加 A-2 續頁最多 2 頁。
 * 欄位契約、逐頁規格見 report-format-spec.md §E、§F。
 *
 *   npm install pptxgenjs
 *   node build.js --input deck_data.json --output report.pptx
 *
 * 設計重點：
 *   - 先決定要出哪些頁（plan），總頁數一開始就確定，頁碼直接蓋 n / N
 *   - 不用章節分隔頁；只有 A-2 引用商品清單會開續頁（每頁 10 筆），其他頁超過列數上限只取前 N 列並警告
 *   - 欄位名錯誤（整欄不存在、主鍵空白）一律中止，避免交出整欄空白的表格
 *   - 價格頁（03-2、03-3、03-4）只放正規通路售價；頁尾預設文字可用 page_sources 依頁覆蓋（02-1、03-4 除外）
 *   - 商品名稱（03-3、04-1、A-2）完整列出、最多兩行，有 url 就加商品頁超連結
 */

const fs = require("fs");
const pptxgen = require("pptxgenjs");
const T = require("./tagtoo_pptx.js");
const C = T.COLORS;

// ── 版面常數（吋）────────────────────────────────────────────────────────────
const X0 = 0.35, W0 = 9.30, BODY_TOP = 1.10, BODY_BOTTOM = 4.80;
const ROW_H = 0.50, ROW_H_APPX = 0.33, APPX_PER_PAGE = 10;
const TBL_HEADER = "873222", TBL_ALT = "F9F3F0", TBL_BORDER = "E2E8F0", CLIENT_BG = "FFF5F3";
const MUTED = "B2B2B2";
const COMP_COLORS = ["8C8C8C", "B5B5B5", "D4D4D4"];
// 02-2 賣家類型是類別，用不同色相（官方＝品牌紅）；四色都配白字（對比 4.2 以上），相鄰兩色色差已用配色檢查工具驗過
const SELLER_COLORS = ["873222", "2A78D6", "A86F00", "167F5A"];
// 04-2 銷量級距是順序，同一色相由淺到深、拉開間距；「未顯示級距」用深灰，表示不知道，不是最低一級
const BAND_COLORS = ["6E6E6E", "D97A62", "A8432F", "74220F", "3A0C06"];
const DEFAULT_TIERS = ["未顯示級距", "未達 1 千件", "1 千件以上", "1 萬件以上", "10 萬件以上"];
const TITLE_MAX = 24, MAX_COLS = 5;
const SELLER_TYPES = ["官方賣場", "連鎖與平台直營", "代購", "一般賣家"];
const PLATFORMS = ["momo", "pchome", "shopee"];
const PLATFORM_LABEL = { momo: "momo", pchome: "PChome", shopee: "蝦皮" };

// ── CLI ───────────────────────────────────────────────────────────────────────
function arg(name, def) {
  const i = process.argv.indexOf(`--${name}`);
  return i >= 0 ? process.argv[i + 1] : def;
}
const INPUT = arg("input");
const OUTPUT = arg("output", "commerce_report.pptx");
if (!INPUT) { console.error("ERROR: 需要 --input deck_data.json"); process.exit(1); }
const data = JSON.parse(fs.readFileSync(INPUT, "utf8"));

const warnings = [];
const warn = (m) => { warnings.push(m); console.warn(`⚠ ${m}`); };
const fail = (m) => { console.error(`🛑 ${m}\n   見 report-format-spec.md §F。`); process.exit(2); };

// ── 格式化 ────────────────────────────────────────────────────────────────────
const isBlank = (v) => v == null || v === "";
const dash = (v) => (isBlank(v) ? "—" : String(v));
const num = (v, d = 0) => {
  if (isBlank(v)) return "—";
  const x = Number(v);
  return Number.isFinite(x) ? x.toLocaleString("en-US", { minimumFractionDigits: d, maximumFractionDigits: d }) : dash(v);
};
const price = (v) => (isBlank(v) ? "—" : `${num(v)} 元`);
const trunc = (s, n) => {
  if (isBlank(s)) return "—";
  const t = String(s);
  return [...t].length > n ? [...t].slice(0, n).join("") + "…" : t;
};
// 商品名稱完整列出，超過欄寬兩行才截斷：中文字算 1 格、英數字算 0.55 格，留一成餘裕給換行位置
const charW = (ch) => (/[\u0000-\u00ff]/.test(ch) ? 0.55 : 1);
function fitLines(s, colW, fontPt, lines = 2) {
  if (isBlank(s)) return "—";
  const budget = ((colW - 0.2) / (fontPt / 72)) * lines * 0.9;
  let w = 0, out = "";
  for (const ch of String(s)) {
    if (w + charW(ch) > budget) return out + "…";
    w += charW(ch); out += ch;
  }
  return out;
}
function linkCell(c, url) {
  if (!isBlank(url)) c.options.hyperlink = { url, tooltip: "開啟商品頁" };
  return c;
}

// ── 契約檢查 ──────────────────────────────────────────────────────────────────
// rows 不存在或為空 → 回傳 []；主鍵空白或必要欄位整欄不存在 → 中止；部分列缺欄 → 警告。
function assertRows(name, rows, keys, cap) {
  if (!Array.isArray(rows) || !rows.length) return [];
  const [primary] = keys;
  const blank = rows.map((r, i) => (isBlank(r[primary]) ? i : -1)).filter((i) => i >= 0);
  if (blank.length) {
    fail(`[${name}] 主鍵「${primary}」在第 ${blank.join(", ")} 列是空的，那些列會畫成空白。` +
      `\n   第 0 列的欄位：${JSON.stringify(Object.keys(rows[0]))}` +
      (blank.length === rows.length ? "\n   每一列主鍵都空白，多半是欄位名整組寫錯。" : ""));
  }
  for (const k of keys.slice(1)) {
    const missing = rows.map((r, i) => (k in r ? -1 : i)).filter((i) => i >= 0);
    if (missing.length === rows.length) fail(`[${name}] 每一列都沒有「${k}」欄位，是否改名了？整欄真的沒有值請明寫 null。`);
    missing.forEach((i) => warn(`[${name}] 第 ${i} 列沒有「${k}」，顯示「—」`));
  }
  if (cap && rows.length > cap) {
    warn(`[${name}] 有 ${rows.length} 列，超過上限 ${cap}，只取前 ${cap} 列`);
    return rows.slice(0, cap);
  }
  return rows;
}
function assertCount(name, arr, min, max) {
  const n = Array.isArray(arr) ? arr.length : 0;
  if (n < min || n > max) fail(`[${name}] 需要 ${min === max ? `剛好 ${min}` : `${min}–${max}`} 筆，目前 ${n} 筆`);
  return arr;
}
function assertSoldBand(name, rows) {
  rows.forEach((r, i) => {
    if (!isBlank(r.sold_band) && /^[\s\d,.]+$/.test(String(r.sold_band))) {
      fail(`[${name}] 第 ${i} 列的 sold_band「${r.sold_band}」是純數字；momo 銷量是級距，要寫成「N 件以上」`);
    }
  });
}
function assertNoShopeeRating(name, rows) {
  rows.forEach((r, i) => {
    if (r.platform === "shopee" && (!isBlank(r.rating) || !isBlank(r.review_count))) {
      fail(`[${name}] 第 ${i} 列是蝦皮商品卻帶有評分或評論數；蝦皮評論數與商品評分都不引用`);
    }
  });
}
// 銷量文字寫法統一：「3 千件以上」「1.5 萬件以上」，不寫「3,000 件以上」
function checkBandText(name, texts) {
  texts.forEach((t, i) => {
    if (!isBlank(t) && /\d,\d{3}\s*件/.test(String(t))) warn(`[${name}] 第 ${i} 筆「${t}」請改寫成「N 千件以上」「N 萬件以上」`);
  });
}
function assertEnum(name, rows, key, allowed) {
  rows.forEach((r, i) => {
    if (!isBlank(r[key]) && !allowed.includes(r[key])) fail(`[${name}] 第 ${i} 列的 ${key}「${r[key]}」不在允許值 ${allowed.join("／")} 之內`);
  });
}

// ── 讀資料並檢查 ──────────────────────────────────────────────────────────────
const client = data.client || "品牌";
const mode = data.mode === "single" ? "single" : "compare";
const fetched = data.fetched_at || "";
const BRAND = String(data.brand_primary_color || "873222").replace("#", "");
const brands = Array.isArray(data.brands) ? data.brands : [];
const clientNames = new Set(brands.filter((b) => b.is_client).map((b) => b.name));
if (!clientNames.size) clientNames.add(client);
const isClient = (name) => clientNames.has(name);
const brandOrder = brands.length ? brands.map((b) => b.name) : [client];
const sortByBrand = (rows) => [...rows].sort((a, b) => {
  const ia = brandOrder.indexOf(a.brand), ib = brandOrder.indexOf(b.brand);
  return (ia < 0 ? 99 : ia) - (ib < 0 ? 99 : ib);
});

const D = {
  kpis: assertCount("summary_kpis", data.summary_kpis, 3, 3),
  findings: assertCount("summary_findings", data.summary_findings, 3, 3),
  matrix: assertRows("channel_matrix", data.channel_matrix?.rows, ["brand", "platform", "sampled", "kept", "removed"], 12),
  sellers: assertRows("shopee_sellers", data.shopee_sellers?.rows, ["brand", "seller_type", "shop_count"], 16),
  assortment: assertRows("assortment", data.assortment?.rows, ["brand", "line", "spec", "price_min"], 6),
  unit: assertRows("unit_price", data.unit_price?.rows, ["label", "n", "median"], 4),
  compare: assertRows("price_compare", data.price_compare?.rows, ["item", "momo", "pchome", "shopee"], 6),
  discount: assertRows("discount", data.discount?.rows, ["brand", "n_discounted", "n_total", "avg_pct"], 4),
  top: assertRows("momo_top", data.momo_top?.rows, ["rank", "brand", "title", "price", "sold_band", "rating", "review_count"], 6),
  band: assertRows("sold_band", data.sold_band?.rows, ["brand", "tier", "count"], 20),
  quotes: assertRows("review_quotes", data.review_quotes?.rows, ["quote", "brand", "rating"], 6),
  claims: assertRows("title_claims", data.title_claims?.rows, ["claim", "brand", "count", "example"], 6),
  insights: assertCount("insights", data.insights, 3, 3),
  actions: assertCount("actions", data.actions, 3, 3),
  scope: assertCount("scope_items", data.scope_items, 1, 5),
  notes: assertCount("appendix_notes", data.appendix_notes, 1, 5),
  appx: assertRows("appendix_products", data.appendix_products?.rows, ["idx", "brand", "platform", "title", "price", "sold_band"], 30),
};
const TIERS = data.sold_band?.tiers || DEFAULT_TIERS;
if (TIERS.length > BAND_COLORS.length) fail(`[sold_band] tiers 有 ${TIERS.length} 段，最多 ${BAND_COLORS.length} 段`);
assertEnum("channel_matrix", D.matrix, "platform", PLATFORMS);
assertEnum("appendix_products", D.appx, "platform", PLATFORMS);
assertEnum("shopee_sellers", D.sellers, "seller_type", SELLER_TYPES);
[["momo_top", D.top], ["appendix_products", D.appx]].forEach(([n, rows]) => {
  const miss = rows.filter((r) => isBlank(r.url)).length;
  if (miss) warn(`[${n}] ${miss} 列沒有 url，商品名不會有連結（url 從 items.json 複製）`);
});
assertEnum("sold_band", D.band, "tier", TIERS);
assertSoldBand("momo_top", D.top);
assertSoldBand("appendix_products", D.appx);
assertNoShopeeRating("momo_top", D.top);
assertNoShopeeRating("appendix_products", D.appx);
checkBandText("momo_top.sold_band", D.top.map((r) => r.sold_band));
checkBandText("appendix_products.sold_band", D.appx.map((r) => r.sold_band));
checkBandText("summary_kpis", D.kpis.map((k) => `${k.value ?? ""} ${k.sub ?? ""}`));
checkBandText("summary_findings", D.findings);

// 每單位售價：沒有可換算商品的品牌不畫成 0 元長條，改在圖下註明
const unitDropped = D.unit.filter((r) => isBlank(r.median) || !(Number(r.n) > 0));
D.unit = D.unit.filter((r) => !unitDropped.includes(r));
unitDropped.forEach((r) => warn(`[unit_price] ${r.label} 沒有可換算的正規通路商品，不畫長條，改在圖下註明`));
// 客戶可換算的正規通路商品少於 3 筆，整頁不出（規則見 project-instructions「數字取用規則」）
const UNIT_MIN = 3;
const clientUnitN = mode === "single"
  ? D.unit.reduce((s, r) => s + (Number(r.n) || 0), 0)
  : D.unit.filter((r) => isClient(r.label)).reduce((s, r) => s + (Number(r.n) || 0), 0);
// 比價單位是「每件」時，每件售價就是售價中位數，和 03-1、03-3 重複，整頁不出
const unitPerItem = /^每?件$/.test(String(data.unit_price?.unit || "").trim());
const unitOk = clientUnitN >= UNIT_MIN && !unitPerItem;
if (!D.unit.length) console.log("· 略過 03-2 每單位售價比較（無資料）");
else if (unitPerItem) console.log("· 略過 03-2 每單位售價比較（比價單位是每件，和 03-1、03-3 重複）");
else if (!unitOk) console.log(`· 略過 03-2 每單位售價比較（${client} 可換算的正規通路商品 ${clientUnitN} 筆，少於 ${UNIT_MIN} 筆）`);

// 01-1「有上架的平台」小字一律依通路矩陣寫出哪幾個平台，不用 Claude 寫的版本
const platformSub = (() => {
  const c = [...clientNames][0];
  const rows = D.matrix.filter((r) => r.brand === c);
  if (!rows.length) return null;
  const have = PLATFORMS.filter((p) => rows.some((r) => r.platform === p && Number(r.kept) > 0)).map((p) => PLATFORM_LABEL[p]);
  const miss = PLATFORMS.filter((p) => !have.includes(PLATFORM_LABEL[p])).map((p) => PLATFORM_LABEL[p]);
  if (!have.length) return "三個平台都查無自家商品";
  return miss.length ? `有：${have.join("、")}｜${miss.join("、")} 查無` : "momo、PChome、蝦皮都有自家商品";
})();
D.kpis.forEach((k) => { if (/平台/.test(String(k.label || "")) && platformSub) k.sub = platformSub; });

const discountTotal = D.discount.reduce((s, r) => s + (Number(r.n_discounted) || 0), 0);
const sampleOf = (p) => Math.max(0, ...D.matrix.filter((r) => r.platform === p).map((r) => Number(r.sampled) || 0)) || "N";

// ── 頁面規劃 ──────────────────────────────────────────────────────────────────
// required: true → 沒資料仍產出並寫「本次無資料」；false → 沒資料整頁略過
const PAGES = [
  { ch: 0, id: "01-1", name: "重點摘要", required: true, has: true, render: pSummary,
    sub: "本次分析的三個關鍵數字與發現" },
  { ch: 1, id: "02-1", name: "三平台上架矩陣", required: true, has: D.matrix.length > 0, render: pMatrix,
    sub: "各品牌在三個平台有沒有自家商品" },
  { ch: 1, id: "02-2", name: "蝦皮賣家分布", required: false, has: D.sellers.length > 0, render: pSellers,
    sub: "誰在蝦皮賣這個品牌" },
  { ch: 2, id: "03-1", name: "產品線與規格", required: true, has: D.assortment.length > 0, render: pAssortment,
    sub: "每條產品線賣哪些規格與組合" },
  { ch: 2, id: "03-2", name: "每單位售價比較", required: false, has: D.unit.length > 0 && unitOk, render: pUnitPrice,
    sub: "換算成同一單位後，誰貴誰便宜" },
  { ch: 2, id: "03-3", name: "跨平台同品比價", required: true, has: D.compare.length > 0, render: pCompare,
    sub: "同一個商品在哪個平台最便宜" },
  { ch: 2, id: "03-4", name: "折扣深度", required: false, has: mode !== "single" && D.discount.length > 1 && discountTotal >= 5, render: pDiscount,
    sub: "各品牌標示原價的商品平均打幾折" },
  { ch: 3, id: "04-1", name: "momo 主力商品", required: true, has: D.top.length > 0, render: pTop,
    sub: "momo 上銷量級距最高的商品" },
  { ch: 3, id: "04-2", name: "momo 銷量級距分布", required: true, has: D.band.length > 0, render: pBand,
    sub: "各品牌的商品落在哪個銷量級距" },
  // 04-3 目前不產出：台灣蝦皮評論工具實測連續失敗（2026-10-28），修好後 review_quotes 有資料就會自動出現
  { ch: 3, id: "04-3", name: "蝦皮買家評論摘錄", required: false, has: D.quotes.length > 0, render: pQuotes,
    sub: "買家實際在意的點" },
  { ch: 4, id: "05-1", name: "商品標題訴求詞", required: true, has: D.claims.length > 0, render: pClaims,
    sub: "各品牌在商品標題上強調什麼" },
  { ch: 4, id: "05-2", name: "競爭洞察", required: true, has: true, render: pInsights,
    sub: "從電商資料看到的競爭局勢" },
  { ch: 4, id: "05-3", name: "行動建議", required: true, has: true, render: pActions,
    sub: "依本次發現建議的下一步" },
  { ch: 5, id: "A-1", name: "研究範圍與資料說明", required: true, has: true, render: pScope,
    sub: "本報告的研究對象、抽樣方式與資料限制" },
  { ch: 5, id: "A-2", name: "引用商品清單", required: true, has: D.appx.length > 0, render: pAppendix,
    sub: "本報告各頁引用的商品" },
];
const CHAPTERS = [
  { no: "01", title: "摘要" },
  { no: "02", title: "通路布局" },
  { no: "03", title: "商品與價格" },
  { no: "04", title: "熱度與口碑" },
  { no: "05", title: "賣點與建議" },
  { no: "A", title: "附錄" },
];
// 議程說明依實際產出的頁面組成，略過的頁（例：03-2、04-3）不會出現在議程上
const AGENDA_PART = {
  "01-1": "三個關鍵數字與發現", "02-1": "三平台上架狀況", "02-2": "蝦皮賣家",
  "03-1": "產品線與規格", "03-2": "每單位售價", "03-3": "跨平台比價", "03-4": "折扣深度",
  "04-1": "momo 主力商品", "04-2": "銷量級距", "04-3": "買家評論",
  "05-1": "標題訴求", "05-2": "競爭洞察", "05-3": "行動建議",
  "A-1": "研究範圍與資料說明", "A-2": "引用商品",
};
const plan = PAGES.filter((p) => p.required || p.has);
// 頁代號 → 實際頁碼（封面 1、議程 2，內容頁從 3 開始）；簡報上一律顯示頁碼，不顯示頁代號
const PAGE_NO = Object.fromEntries(plan.map((p, i) => [p.id, i + 3]));
const PAGE_CODE = /(?<![\w-])(0[1-5]-[1-4]|A-[12])(?![\w-])/g;
function evidenceText(raw) {
  const codes = String(raw).match(PAGE_CODE) || [];
  if (!codes.length) return String(raw);
  const nums = codes.filter((c) => PAGE_NO[c]).map((c) => PAGE_NO[c]);
  codes.filter((c) => !PAGE_NO[c]).forEach((c) => warn(`[insights] 依據頁 ${c} 本次沒有產出，已從「見第 N 頁」移除`));
  return nums.length ? `見第 ${[...new Set(nums)].join("、")} 頁` : "";
}
// 除了 insights[].evidence（會自動換成頁碼），客戶看得到的文字不可出現頁代號
(function checkPageCodes(v, path) {
  if (typeof v === "string") { if (PAGE_CODE.test(v)) warn(`[${path}] 出現頁代號「${v.match(PAGE_CODE).join("、")}」，客戶看不懂，請改寫`); PAGE_CODE.lastIndex = 0; return; }
  if (Array.isArray(v)) return v.forEach((x, i) => checkPageCodes(x, `${path}[${i}]`));
  if (v && typeof v === "object") Object.entries(v).forEach(([k, x]) => {
    if (k === "evidence" || k.startsWith("_")) return;
    checkPageCodes(x, path ? `${path}.${k}` : k);
  });
})(data, "");
const skipWhy = (p) => (p.id === "03-4" && mode === "single" ? "單品牌模式不出" : p.id === "03-4" && D.discount.length ? "有原價的商品少於 5 筆" : "無資料");
PAGES.filter((p) => !plan.includes(p) && p.id !== "03-2").forEach((p) => console.log(`· 略過 ${p.id} ${p.name}（${skipWhy(p)}）`));
const APPX_PAGES = plan.some((p) => p.id === "A-2") ? Math.max(1, Math.ceil(D.appx.length / APPX_PER_PAGE)) : 0;
const BASE_TOTAL = plan.length + 3;   // 封面、議程、結尾
const TOTAL = BASE_TOTAL + Math.max(0, APPX_PAGES - 1);   // A-2 續頁排在最後，前面各頁頁碼不受影響
if (BASE_TOTAL < 14 || BASE_TOTAL > 17) warn(`總頁數 ${BASE_TOTAL}（不含 A-2 續頁）不在 14–17 頁：${plan.map((p) => p.id).join("、")}`);

// ── 共用元件 ──────────────────────────────────────────────────────────────────
function hdr(text) { return { text, options: { bold: true, color: "FFFFFF", fill: { color: TBL_HEADER }, align: "center" } }; }
function cell(text, i, client, extra = {}) {
  return { text: dash(text), options: { fill: { color: client ? CLIENT_BG : (i % 2 ? TBL_ALT : "FFFFFF") }, align: "center", ...extra } };
}
function table(s, header, rows, colW, opts = {}) {
  if (header.length > MAX_COLS) warn(`表格 ${header.length} 欄，超過 ${MAX_COLS} 欄`);
  const rowH = opts.rowH || ROW_H;
  s.addTable([header.map(hdr), ...rows], {
    x: X0, y: BODY_TOP, w: W0, colW, rowH, valign: "middle", autoPage: false,
    fontFace: T.FONT, fontSize: opts.fontSize || 10, color: C.dark,
    border: { type: "solid", color: TBL_BORDER, pt: 0.5 },
  });
}
function caption(s, text) {
  s.addText(text, { x: X0, y: 4.52, w: W0, h: 0.24, fontFace: T.FONT, fontSize: 8.5, color: C.gray, align: "center" });
}
const barColors = (labels) => {
  let k = 0;
  return labels.map((l) => (isClient(l) ? BRAND : COMP_COLORS[Math.min(k++, COMP_COLORS.length - 1)]));
};
const CHART_BASE = {
  catAxisLabelFontFace: T.CJK_FONT, catAxisLabelFontSize: 10, catAxisLabelColor: C.dark,
  valAxisLabelFontFace: T.LATIN_FONT, valAxisLabelFontSize: 8, valAxisLabelColor: C.gray,
  valAxisMinVal: 0, valGridLine: { color: "EEEEEE", size: 0.5 },
  dataLabelFontFace: T.CJK_FONT, dataLabelFontSize: 9, dataLabelColor: C.dark,
  legendFontFace: T.CJK_FONT, legendFontSize: 9,
};
function stackedByBrand(s, rows, segKey, segments, valKey, colors = SELLER_COLORS, h = 3.35) {
  const names = sortByBrand(rows).map((r) => r.brand).filter((v, i, a) => a.indexOf(v) === i);
  const val = (b, seg) => { const r = rows.find((x) => x.brand === b && x[segKey] === seg); return r ? Number(r[valKey]) || 0 : 0; };
  if (names.length === 1) {
    // 單一品牌：直條，一根一段
    s.addChart(pptx.ChartType.bar, [{ name: names[0], labels: segments, values: segments.map((g) => val(names[0], g)) }], {
      ...CHART_BASE, x: X0, y: BODY_TOP, w: W0, h, barDir: "col", chartColors: colors,
      showLegend: false, showValue: true, dataLabelFormatCode: "0", valAxisLabelFormatCode: "0",
    });
    return;
  }
  // 橫條由下往上排，反轉讓客戶出現在最上方
  const labels = [...names].reverse();
  s.addChart(pptx.ChartType.bar, segments.map((g) => ({ name: g, labels, values: labels.map((b) => val(b, g)) })), {
    ...CHART_BASE, x: X0, y: BODY_TOP, w: W0, h, barDir: "bar", barGrouping: "stacked", chartColors: colors,
    showLegend: true, legendPos: "b", showValue: true, dataLabelFormatCode: "0;;;", valAxisLabelFormatCode: "0",
    dataLabelColor: "FFFFFF",
  });
}

// ── 各頁 ──────────────────────────────────────────────────────────────────────
function pSummary(ctx) {
  const s = slideOf(ctx);
  const cw = 2.90, gap = 0.30;
  D.kpis.forEach((k, i) => {
    const x = X0 + i * (cw + gap);
    s.addShape("roundRect", { x, y: BODY_TOP, w: cw, h: 1.10, fill: { color: TBL_ALT }, line: { type: "none" }, rectRadius: 0.08 });
    s.addText(dash(k.label), { x: x + 0.15, y: BODY_TOP + 0.08, w: cw - 0.3, h: 0.26, fontFace: T.FONT, fontSize: 9.5, color: C.gray });
    s.addText(dash(k.value), { x: x + 0.15, y: BODY_TOP + 0.34, w: cw - 0.3, h: 0.44, fontFace: T.font(String(k.value)), fontSize: 22, bold: true, color: BRAND });
    if (!isBlank(k.sub)) s.addText(String(k.sub), { x: x + 0.15, y: BODY_TOP + 0.78, w: cw - 0.3, h: 0.24, fontFace: T.FONT, fontSize: 8.5, color: C.gray });
  });
  D.findings.forEach((f, i) => {
    if ([...String(f)].length > 40) warn(`[summary_findings] 第 ${i} 句超過 40 字`);
    const y = 2.55 + i * 0.72;
    s.addText(String(i + 1), { x: X0, y, w: 0.4, h: 0.4, fontFace: T.LATIN_FONT, fontSize: 16, bold: true, color: BRAND });
    s.addText(String(f), { x: X0 + 0.5, y, w: W0 - 0.5, h: 0.5, fontFace: T.FONT, fontSize: 12, color: C.dark, valign: "top" });
  });
}

function pMatrix(ctx) {
  const n = D.matrix.map((r) => Number(r.sampled) || 0).filter((v) => v > 0);
  const range = !n.length ? "" : Math.min(...n) === Math.max(...n) ? `各 ${Math.max(...n)} 筆` : `${Math.min(...n)}–${Math.max(...n)} 筆`;
  const s = slideOf({ ...ctx, ds: FIXED(ctx, `各格＝該平台搜尋結果中屬於該品牌的商品數，含第三方轉賣${range ? `；各平台搜尋結果 ${range}` : ""}`) });
  const names = sortByBrand(D.matrix).map((r) => r.brand).filter((v, i, a) => a.indexOf(v) === i);
  const rows = names.map((b, i) => [
    { text: b, options: { bold: isClient(b), color: isClient(b) ? BRAND : C.dark, fill: { color: isClient(b) ? CLIENT_BG : (i % 2 ? TBL_ALT : "FFFFFF") }, align: "center" } },
    ...PLATFORMS.map((p) => {
      const r = D.matrix.find((x) => x.brand === b && x.platform === p);
      if (!r) return cell("—", i, isClient(b));
      const k = Number(r.kept);
      return k > 0 ? cell(`${num(k)} 筆`, i, isClient(b)) : cell("查無自家商品", i, isClient(b), { color: MUTED });
    }),
  ]);
  table(s, ["品牌", "momo", "PChome", "蝦皮"], rows, [2.70, 2.20, 2.20, 2.20]);
}

function pSellers(ctx) {
  const s = slideOf({ ...ctx, ds: DS(ctx, `依賣場名稱歸類；同一賣場算一個賣家｜蝦皮搜尋前 ${sampleOf("shopee")} 筆`) });
  stackedByBrand(s, D.sellers, "seller_type", SELLER_TYPES, "shop_count", SELLER_COLORS);
  // momo、PChome 沒有可分四類的賣家資料，改列第三方轉賣筆數（analyze.py 的 channel_matrix.resale）
  if (D.matrix.some((r) => "resale" in r)) {
    const part = (p) => {
      const rows = sortByBrand(D.matrix.filter((r) => r.platform === p && Number(r.kept) > 0));
      return rows.length ? `${PLATFORM_LABEL[p]} ${rows.map((r) => `${r.brand} ${num(r.resale)} 筆`).join("、")}` : `${PLATFORM_LABEL[p]} 查無商品`;
    };
    caption(s, `momo、PChome 沒有可分類的賣家資料，第三方轉賣另計：${part("momo")}｜${part("pchome")}`);
  }
}

function pAssortment(ctx) {
  const s = slideOf({ ...ctx, ds: DS(ctx, "依商品標題整理｜momo、PChome、蝦皮搜尋結果") });
  const rows = sortByBrand(D.assortment).map((r, i) => {
    const c = isClient(r.brand);
    return [cell(r.brand, i, c, { bold: c, color: c ? BRAND : C.dark }), cell(trunc(r.line, 14), i, c),
      cell(trunc(r.spec, 20), i, c), cell(price(r.price_min), i, c)];
  });
  table(s, ["品牌", "產品線", "規格與組合", "最低售價"], rows, [1.70, 2.60, 3.40, 1.60]);
}

function pUnitPrice(ctx) {
  const unit = data.unit_price?.unit || "每單位";
  const s = slideOf({ ...ctx, ds: DS(ctx, "依商品標題容量換算；只計正規通路，換算不出的不列入｜樣本數見各品牌下方") });
  const rows = [...D.unit];
  const mm = data.unit_price?.market_median;
  if (mode === "single" && !isBlank(mm)) rows.push({ label: "同品類市場中位數", n: null, median: mm, _ref: true });
  // 各品牌下方寫樣本數、平台分布各一行（analyze.py 的 by_platform），讀者不用對照圖說
  const sampleLine = (r) => {
    if (r._ref) return "";
    const bp = r.by_platform ? PLATFORMS.filter((p) => Number(r.by_platform[p]) > 0).map((p) => `${PLATFORM_LABEL[p]} ${num(r.by_platform[p])}`) : [];
    return `\n${num(r.n)} 筆${bp.length ? `\n${bp.join("・")}` : ""}`;
  };
  const labels = rows.map((r) => r.label);
  const axisLabels = rows.map((r) => `${r.label}${sampleLine(r)}`);
  const colors = rows.map((r, i) => (r._ref ? "D4D4D4" : (mode === "single" ? BRAND : barColors(labels)[i])));
  s.addChart(pptx.ChartType.bar, [{ name: `元／${unit}`, labels: axisLabels, values: rows.map((r) => Number(r.median) || 0) }], {
    ...CHART_BASE, x: X0, y: BODY_TOP, w: W0, h: 3.35, barDir: "col", chartColors: colors, barGapWidthPct: 80, catAxisLabelFontSize: 9,
    showLegend: false, showValue: true, dataLabelFormatCode: "0.00\"元\"", valAxisLabelFormatCode: "0.00",
    showValAxisTitle: true, valAxisTitle: `元／${unit}`, valAxisTitleFontFace: T.CJK_FONT, valAxisTitleFontSize: 9, valAxisTitleColor: C.gray,
  });
  const dropped = unitDropped.length ? `；${unitDropped.map((r) => r.label).join("、")}無可換算商品` : "";
  // 每粒、每包這類計數單位：各品牌每粒含量與劑型不同，提醒只看價格帶
  const countUnit = /^每?(粒|顆|錠|包|片|入|條)$/.test(String(unit).trim()) ? `；各品牌${unit}含量不同，僅供價格帶參考` : "";
  const scope = data.category_keywords?.length ? `標題含「${data.category_keywords.join("」或「")}」、` : "";
  caption(s, `只計${scope}算得出容量的正規通路商品` + dropped + countUnit);
}

// 最便宜由三個價格自行判斷，不讀資料裡的 cheapest：同價寫「同價」，只有一個平台有價寫「僅一平台」
function cheapestOf(r) {
  const priced = PLATFORMS.filter((p) => !isBlank(r[p]) && Number.isFinite(Number(r[p])));
  if (priced.length < 2) return { label: priced.length ? "僅一平台" : "—", win: [] };
  const min = Math.min(...priced.map((p) => Number(r[p])));
  const win = priced.filter((p) => Number(r[p]) === min);
  return { label: win.length === priced.length ? "同價" : win.map((p) => PLATFORM_LABEL[p]).join("、"), win: win.length === priced.length ? [] : win };
}
function pCompare(ctx) {
  const s = slideOf({ ...ctx, ds: DS(ctx, "同品牌、同容量與件數，系列相同、品名相近的正規通路商品；同平台多筆取品名最接近者") });
  const rows = D.compare.map((r, i) => {
    const c = cheapestOf(r);
    return [
      cell(fitLines(r.item, 3.30, 10), i, false, { align: "left" }),
      ...PLATFORMS.map((p) => cell(price(r[p]), i, false, { bold: c.win.includes(p), color: c.win.includes(p) ? BRAND : C.dark })),
      cell(c.label, i, false, { bold: true }),
    ];
  });
  table(s, ["品項", "momo", "PChome", "蝦皮", "最便宜"], rows, [3.30, 1.50, 1.50, 1.50, 1.50], { rowH: 0.46 });
  caption(s, "只列正規通路售價（品牌官方、平台自營、連鎖通路）；「—」表示該平台沒有正規通路售價");
}

function pDiscount(ctx) {
  const s = slideOf({ ...ctx, ds: FIXED(ctx, "折扣率＝1 − 售價 ÷ 平台標示原價，只計正規通路；依擷取當日價格計算，會隨每日價格與檔期活動變動") });
  // 直條，客戶在最左；各品牌下方三行：品牌、正規通路件數、納入計算件數（2026-11-02 PowerPoint 實測多行標籤正常）
  const rows = sortByBrand(D.discount);
  const colors = barColors(rows.map((r) => r.brand));
  const labels = rows.map((r) => `${r.brand}\n正規通路 ${num(r.n_total)} 件\n納入計算 ${num(r.n_discounted)} 件`);
  s.addChart(pptx.ChartType.bar, [{ name: "平均折扣率", labels, values: rows.map((r) => Number(r.avg_pct) || 0) }], {
    ...CHART_BASE, x: X0, y: BODY_TOP, w: W0, h: 3.35, barDir: "col", chartColors: colors, barGapWidthPct: 80, catAxisLabelFontSize: 9,
    showLegend: false, showValue: true, dataLabelFormatCode: "0\"%\"", valAxisLabelFormatCode: "0\"%\"",
  });
  caption(s, "＊折扣率要有平台標示原價才算得出來，只有標示原價的商品納入計算");
}

function pTop(ctx) {
  const s = slideOf({ ...ctx, ds: DS(ctx, "銷量為 momo 顯示的累計級距，屬第三方估計") });
  const withImg = D.top.some((r) => !isBlank(r.image_path));
  const rows = D.top.map((r, i) => {
    const c = isClient(r.brand);
    const name = fitLines(withImg ? `${r.brand}｜${r.title}` : r.title, 4.40, 9);
    const titleCell = linkCell(cell(name, i, c, { align: "left", fontSize: 9 }), r.url);
    const rating = isBlank(r.rating) ? "—" : `${num(r.rating, 1)}（${num(r.review_count)} 則）`;
    return [withImg ? cell("", i, c) : cell(r.brand, i, c, { bold: c, color: c ? BRAND : C.dark }),
      titleCell, cell(price(r.price), i, c), cell(r.sold_band, i, c), cell(rating, i, c)];
  });
  table(s, ["品牌", "商品名", "售價", "銷量級距", "評分"], rows, [1.10, 4.40, 0.90, 1.30, 1.60]);
  if (withImg) {
    D.top.forEach((r, i) => {
      const box = { x: X0 + 0.32, y: BODY_TOP + ROW_H * (i + 1) + 0.025, w: 0.45, h: 0.45 };
      if (isBlank(r.image_path) || !T.coverImage(s, { path: r.image_path }, box)) {
        s.addShape("rect", { ...box, fill: { color: "EEEEEE" }, line: { color: "CCCCCC", width: 0.5 } });
      }
    });
  }
}

function pBand(ctx) {
  const s = slideOf({ ...ctx, ds: DS(ctx, "momo 顯示的累計銷量級距，屬第三方估計；非期間銷量") });
  stackedByBrand(s, D.band, "tier", TIERS, "count", BAND_COLORS);
}

function pQuotes(ctx) {
  const s = slideOf({ ...ctx, ds: DS(ctx, `蝦皮評論抽樣 ${data.review_quotes?.sample_size ?? "N"} 則；抽樣可能偏向高分評論，僅供質性參考`) });
  const rows = sortByBrand(D.quotes).map((r, i) => {
    const c = isClient(r.brand);
    return [cell(r.brand, i, c, { bold: c, color: c ? BRAND : C.dark }), cell(trunc(r.quote, 40), i, c, { align: "left" }),
      cell(isBlank(r.rating) ? "—" : `${r.rating} 星`, i, c)];
  });
  table(s, ["品牌", "評論摘錄", "評分"], rows, [1.60, 6.50, 1.20]);
}

function pClaims(ctx) {
  const s = slideOf({ ...ctx, ds: DS(ctx, "依商品標題文字歸納｜三平台搜尋結果") });
  const rows = D.claims.map((r, i) => {
    const c = isClient(r.brand);
    return [cell(r.claim, i, c, { bold: true }), cell(r.brand, i, c, { color: c ? BRAND : C.dark }),
      cell(num(r.count), i, c), cell(trunc(r.example, 20), i, c, { align: "left" })];
  });
  table(s, ["訴求", "品牌", "出現次數", "範例"], rows, [1.90, 1.70, 1.30, 4.40]);
}

function threeBlocks(s, items, withEvidence) {
  items.forEach((it, i) => {
    const y = BODY_TOP + 0.05 + i * 1.20;
    if ([...String(it.title || "")].length > 20) warn(`[${withEvidence ? "insights" : "actions"}] 第 ${i} 則標題超過 20 字`);
    if ([...String(it.body || "")].length > 60) warn(`[${withEvidence ? "insights" : "actions"}] 第 ${i} 則說明超過 60 字`);
    s.addShape("rect", { x: X0, y, w: 0.06, h: 1.0, fill: { color: BRAND }, line: { type: "none" } });
    s.addText(dash(it.title), { x: X0 + 0.2, y, w: withEvidence ? 7.2 : 9.0, h: 0.34, fontFace: T.FONT, fontSize: 13, bold: true, color: C.dark });
    s.addText(dash(it.body), { x: X0 + 0.2, y: y + 0.36, w: withEvidence ? 7.2 : 9.0, h: 0.62, fontFace: T.FONT, fontSize: 11, color: C.dark, valign: "top" });
    const ev = withEvidence && !isBlank(it.evidence) ? evidenceText(it.evidence) : "";
    if (ev) {
      s.addText(ev, { x: 7.85, y, w: 1.8, h: 0.34, fontFace: T.FONT, fontSize: 9, color: C.gray, align: "right" });
    }
  });
}
function pInsights(ctx) { threeBlocks(slideOf(ctx), D.insights, true); }
function pActions(ctx) { threeBlocks(slideOf(ctx), D.actions, false); }

function pScope(ctx) {
  const s = slideOf(ctx);
  const col = (x, head, items) => {
    s.addText(head, { x, y: BODY_TOP, w: 4.5, h: 0.32, fontFace: T.FONT, fontSize: 12, bold: true, color: BRAND });
    s.addText(items.map((t) => ({ text: String(t), options: { bullet: true, breakLine: true } })), {
      x, y: BODY_TOP + 0.40, w: 4.5, h: 3.25, fontFace: T.FONT, fontSize: 9.5, color: C.dark, valign: "top", paraSpaceAfter: 6,
    });
  };
  col(X0, "研究範圍", D.scope);
  col(X0 + 4.80, "資料說明", D.notes);
}

// 每頁 10 筆，超過就開續頁（標題加「（續）」）；商品名點了開商品頁
function pAppendix(ctx) {
  for (let k = 0; k < APPX_PAGES; k++) {
    if (k > 0) page += 1;
    const s = slideOf({ ...ctx, title: k ? `${ctx.title}（續）` : ctx.title,
      ds: DS(ctx, "本報告各頁引用的商品，點商品名開啟商品頁；各平台抽樣與剔除筆數見研究範圍頁") });
    const rows = D.appx.slice(k * APPX_PER_PAGE, (k + 1) * APPX_PER_PAGE).map((r, i) => {
      const c = isClient(r.brand);
      return [cell(r.brand, i, c), cell(PLATFORM_LABEL[r.platform] || r.platform, i, c),
        linkCell(cell(fitLines(r.title, 4.80, 8), i, c, { align: "left" }), r.url),
        cell(price(r.price), i, c), cell(r.sold_band, i, c)];
    });
    table(s, ["品牌", "平台", "商品名", "售價", "銷量級距"], rows, [1.20, 0.80, 4.80, 1.00, 1.50], { rowH: ROW_H_APPX, fontSize: 8 });
  }
}

// ── 簡報 ──────────────────────────────────────────────────────────────────────
const pptx = new pptxgen();
T.init(pptx);
const FOOTER = `${client} × Tagtoo 電商品牌分析`;
const SRC = (s) => `${s}｜擷取日 ${fetched}`;
// 頁尾：deck_data.page_sources[頁代號] 有寫就用它（實際來源和預設不同時，例如 05-1 只統計 momo）
const PAGE_SOURCES = data.page_sources || {};
const DS = (ctx, def) => SRC(isBlank(PAGE_SOURCES[ctx.id]) ? def : PAGE_SOURCES[ctx.id]);
// 02-1、03-4 的頁尾是固定說明，不接受 page_sources 覆蓋
const FIXED = (ctx, def) => {
  if (!isBlank(PAGE_SOURCES[ctx.id])) warn(`[page_sources] ${ctx.id} 的頁尾是固定說明，忽略 page_sources 的寫法`);
  return SRC(def);
};
let page = 0;

// 封面資訊行：分析品項、比較品牌（比較模式）、分析來源、擷取日，由上往下排在分隔線下方
const cover = T.addCover(pptx, { mainTitle: client, subtitleEN: "E-commerce Brand Analysis", reportTitle: "電商品牌分析報告" });
const rivals = brands.filter((b) => !b.is_client).map((b) => b.momo_tag && b.momo_tag !== b.name && !String(b.momo_tag).includes(b.name) ? `${b.momo_tag} ${b.name}` : (b.momo_tag || b.name));
if (isBlank(data.category)) warn("[category] 沒有比較品項，封面只寫分析來源與擷取日");
[
  !isBlank(data.category) && `分析品項：${data.category}`,
  mode !== "single" && rivals.length && `比較品牌：${rivals.join("、")}`,
  "分析來源：momo · PChome · 蝦皮",
  `資料擷取日 ${fetched}`,
].filter(Boolean).forEach((t, i) => cover.addText(t, {
  x: 0.5, y: 3.65 + i * 0.35, w: 6.5, h: 0.3, fontFace: T.font(t), fontSize: 11, color: "FFFFFF",
}));
page += 1;
T.addAgenda(pptx, {
  pageNum: ++page, totalPages: TOTAL,
  items: CHAPTERS.map((c, i) => ({ ...c, desc: plan.filter((p) => p.ch === i).map((p) => AGENDA_PART[p.id]).join("、") }))
    .filter((c) => c.desc),
});

const pageTitles = data.page_titles || {};
for (const p of plan) {
  let title = pageTitles[p.id];
  if (isBlank(title)) { warn(`[page_titles] ${p.id} 沒有結論句標題，改用頁名「${p.name}」`); title = p.name; }
  if ([...String(title)].length > TITLE_MAX) warn(`[page_titles] ${p.id} 標題 ${[...String(title)].length} 字，超過 ${TITLE_MAX} 字`);
  page += 1;
  const ctx = { id: p.id, title, sub: p.sub, ds: DS({ id: p.id }, "資料來源：Tagtoo Commerce MCP（momo／PChome／蝦皮公開頁）") };
  if (!p.has) {
    const s = slideOf(ctx);
    s.addText(`本次無資料：${p.name}所需的資料沒有取得，原因見附錄「研究範圍與資料說明」。`, {
      x: X0, y: 2.4, w: W0, h: 0.6, fontFace: T.FONT, fontSize: 12, color: C.gray, align: "center",
    });
    continue;
  }
  p.render(ctx);
}
T.addClosing(pptx, { subtitle: `${client} 電商品牌分析報告`, poweredBy: "Powered by Tagtoo Agency Harness", pageNum: ++page, totalPages: TOTAL });

function slideOf(ctx) {
  return T.addContent(pptx, { title: ctx.title, subtitle: ctx.sub, footerCenter: FOOTER, dataSource: ctx.ds, pageNum: page, totalPages: TOTAL });
}

// ── 輸出 ──────────────────────────────────────────────────────────────────────
pptx.writeFile({ fileName: OUTPUT }).then(() => {
  console.log(`✓ ${OUTPUT}（${TOTAL} 頁：${plan.map((p) => p.id).join("、")}）`);
  if (warnings.length) console.log(`  警告 ${warnings.length} 則，見上方 ⚠`);
});
