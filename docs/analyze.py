#!/usr/bin/env python3
# SPDX-License-Identifier: Apache-2.0
"""
analyze.py — 品牌比對、通路分類、統計，產出摘要與簡報共用的數字
-----------------------------------------------------------------
讀手抄的 raw/compact.json（格式見 project-instructions.md「資料蒐集」）與 brands.json；
也讀得懂完整工具回傳格式的 raw/_index.json 加各檔（舊版存檔方式），依 project-instructions.md 的
「品牌比對與剔除規則」「數字取用規則」「通路分類規則」處理，輸出：

  items.json       每筆商品一列（平台、品牌、通路、售價、級距、換算量……），A-2 與人工判斷用
  deck_parts.json  可直接放進 deck_data.json 的區塊：client、brands、summary_kpis（第 1 筆上架平台）、
                   channel_matrix、shopee_sellers、unit_price、price_compare、discount、momo_top、
                   sold_band、appendix_products
  stats.json       摘要確認用的指標（上架平台數、售價中位數、momo 最高級距……）
  summary_draft.md 摘要確認的版面，指標表與固定段落已填好，Claude 只補〔〕裡的結論、發現與建議

畫面上只印統計結果與需要人工確認的少數商品，不印整份清單。
摘要與簡報都讀這三個檔，不重新查詢，數字才會一致。

用法：python3 analyze.py --raw raw --brands brands.json --out .
      印摘要前最後一次執行加 --lock：在輸出資料夾留下鎖定標記 .summary_locked，之後數字不可再變。
      已鎖定時再執行會印 🛑 並停止；只有對方在摘要確認時要求修改，才加 --unlock 重跑（重跑後維持鎖定）。

brands.json 範例：
{
  "unit": "每毫升",                       // 每克、每毫升、每粒、每件
  "g_ml_same": true,                      // 選填：液體類（洗衣精、洗髮精、沐浴乳）克與毫升混標時，1 克視同 1 毫升
  "category": "洗髮精",                   // 比較品項：各品牌搜尋詞都是「品牌 品項」，封面也會寫出來
  "category_keywords": ["洗髮"],          // 選填：標題要含其中一個字詞才算這個品項；不含的商品剔除，各頁都不計
  "brands": [
    { "name": "森歐黎漾", "client": true,
      "query": "SAHOLEA 森歐黎漾 洗髮精",   // 三平台用的搜尋詞（要和實際查詢一字不差）
      "momo_tag": "SAHOLEA 森歐黎漾",       // momo 標題【】裡的品牌標記（完全相符才算；有兩種寫法時給清單）
      "keywords": ["森歐黎漾", "SAHOLEA"],  // 蝦皮、PChome 標題比對用
      "pchome_lines": ["淨平衡", "極緞光"]   // 選填：PChome 標題常不寫品牌，用產品線名稱比對
    }
  ],
  "unit_overrides": { "momo:12730446": 2400, "pchome:DDBX7J-A900HEJTB": null }   // 選填：人工指定換算量，null＝不列入
}
"""
import argparse, json, os, re, statistics as st, sys

PLATFORMS = ("momo", "pchome", "shopee")
TIERS = ["未顯示級距", "未達 1 千件", "1 千件以上", "1 萬件以上", "10 萬件以上"]
SELLER_TYPES = ["官方賣場", "連鎖與平台直營", "代購", "一般賣家"]
CHAINS = ("康是美", "屈臣氏", "Watsons", "watsons", "寶雅", "POYA", "家樂福", "全聯", "愛買", "大潤發", "萬家福",
          "樂家康", "日藥本舖", "松本清", "美廉社", "躍獅", "大樹藥局", "丁丁藥局", "杏一", "燦坤", "全國電子",
          "順發", "誠品", "博客來", "Costco 官方")
PROXY = ("代購", "好市多", "Costco", "COSTCO", "costco", "代買", "代訂")


def norm(s):
    return re.sub(r"\s+", "", str(s or "")).lower()


def has_any(text, words):
    t = norm(text)
    return any(norm(w) and norm(w) in t for w in words)


# ── 銷量級距 ──────────────────────────────────────────────────────────────────
def sold_value(p):
    u = p.get("units_sold") or {}
    if u.get("available") and u.get("value") is not None:
        return float(u["value"])
    raw = str(u.get("raw") or "")
    m = re.search(r">\s*([\d,.]+)\s*(萬)?", raw)
    if not m:
        return None
    v = float(m.group(1).replace(",", ""))
    return v * 10000 if m.group(2) else v


def band_text(v):
    if v is None:
        return None
    if v >= 10000:
        return f"{v / 10000:g} 萬件以上"
    if v >= 1000:
        return f"{v / 1000:g} 千件以上"
    return f"{int(v)} 件以上"


def tier(v):
    if v is None:
        return TIERS[0]
    return TIERS[1] if v < 1000 else TIERS[2] if v < 10000 else TIERS[3] if v < 100000 else TIERS[4]


# ── 每單位換算 ────────────────────────────────────────────────────────────────
UNIT_RE = {
    "每毫升": (r"(\d+(?:\.\d+)?)\s*(ml|毫升|公升|l)(?![a-wyz])", {"ml": 1, "毫升": 1, "公升": 1000, "l": 1000}),
    "每克": (r"(?<![mμ])(\d+(?:\.\d+)?)\s*(kg|公斤|g|公克|克)(?![a-wyz])", {"kg": 1000, "公斤": 1000, "g": 1, "公克": 1, "克": 1}),
    "每粒": (r"(\d+)\s*(粒|顆|錠|膠囊)", {"粒": 1, "顆": 1, "錠": 1, "膠囊": 1}),
}
CN = {"一": 1, "二": 2, "兩": 2, "雙": 2, "三": 3, "四": 4, "五": 5, "六": 6, "十": 10}


def multiplier(t):
    m = re.search(r"共\s*(\d+)\s*(入|瓶|罐|盒|包|袋|組|件|條)", t)
    if m:
        return int(m.group(1))
    m = re.search(r"[x×*＊]\s*(\d+)(?!\s*(ml|g|毫升|公克|克|粒|顆))", t, re.I)
    if m:
        return int(m.group(1))
    m = re.search(r"(?<![\d.])(\d+)\s*(入|瓶組|罐組|盒組|包組)", t)
    if m:
        return int(m.group(1))
    m = re.search(r"[(（]\s*(\d+)\s*(罐|盒|包|袋|瓶)\s*[)）]", t)
    if m:
        return int(m.group(1))
    m = re.search(r"([一二兩雙三四五六十])\s*(入|瓶|罐|盒)組?", t)
    if m:
        return CN[m.group(1)]
    return 1


def size_patterns(unit, g_ml_same):
    """回傳 [(容量正規式, 換算表)]；液體類可把克與毫升視為相同（1 克＝1 毫升）"""
    if g_ml_same and unit in ("每毫升", "每克"):
        return [UNIT_RE["每毫升"], UNIT_RE["每克"]]
    return [UNIT_RE[unit]]


def unit_qty(title, unit, g_ml_same=False):
    """回傳 (總量, 說明, 單件容量)；換算不出時總量為 None"""
    if unit not in UNIT_RE:
        return None, "不換算", None
    t = title.replace("ＭＬ", "ml").replace("Ｇ", "g")
    pats = size_patterns(unit, g_ml_same)
    found, rest = [], t   # found：(容量, 緊接在容量後的件數)；rest：拿掉「容量×件數」後剩下的標題
    for pat, scale in pats:
        for m in re.finditer(pat + r"(?:\s*[x×*＊]\s*(\d+))?", t, re.I):
            found.append((float(m.group(1)) * scale[m.group(2).lower()], int(m.group(3)) if m.group(3) else 1))
            rest = rest.replace(m.group(0), " ", 1)
    if not found:
        return None, "標題沒有容量", None
    if re.search(r"任選|多款|款式可選", t) and re.search(r"[+＋/／]", t):
        return None, "多款任選，無法判斷", None
    other = {"每毫升": "每克", "每克": "每毫升"}.get(unit)
    if other and not g_ml_same and re.search(r"[+＋]", t) and re.search(UNIT_RE[other][0], t, re.I):
        return None, "組合含其他單位，無法判斷", None
    # 外層件數從剩下的標題找：「黑芝麻糊x2袋(30gx12入/袋)」＝30 克×12 入×2 袋；「(650gx4入/盒)三盒組」＝650×4×3
    outer = multiplier(rest)
    if len(found) > 1 and re.search(r"[+＋]", t):
        # 組合包：「1800ml×1 瓶＋1700ml×2 包」→ 每段容量×件數加總；「700g+100g（4罐）」再乘外層件數
        total = sum(sz * k for sz, k in found) * outer
        return total, "＋".join(f"{sz:g}×{k}" for sz, k in found) + (f"，共 {outer} 組" if outer > 1 else ""), max(sz for sz, _ in found)
    sizes = {sz for sz, _ in found}
    if len(sizes) > 1:
        return None, "多種容量，無法判斷", None
    s = sizes.pop()
    k = found[0][1] * outer
    return s * k, f"{s:g}×{k}", s


# ── 通路分類 ──────────────────────────────────────────────────────────────────
# 官方賣場只認賣場名含該品牌名稱或關鍵字的店；店名有「官方」「旗艦」但不含品牌名的是經銷商或第三方，
# 品牌自營店也常不寫「官方」（例：御復珍 ufood、耕沅GengYuanOil）（2026-11-02 馬玉山V2 實測）。
# 蝦皮直營是平台自營，屬授權通路，跟連鎖通路同一類。
def seller_type(p, b):
    name = (p.get("shop") or {}).get("name") or ""
    if "蝦皮直營" in name:
        return "連鎖與平台直營"
    if any(w in name for w in PROXY):
        return "代購"
    if has_any(name, b["keywords"] + [b["name"]]):
        return "官方賣場"
    if any(c in name for c in CHAINS):
        return "連鎖與平台直營"
    return "一般賣家"


def momo_tags(title):
    return re.findall(r"【([^】]+)】", title or "")


def classify(platform, p, b):
    """回傳 (是否該品牌, 通路, 蝦皮賣家類別)"""
    title = p.get("title") or ""
    if platform == "momo":
        tags = momo_tags(title)
        want = b["momo_tag"] if isinstance(b["momo_tag"], list) else [b["momo_tag"]]
        if any(norm(w) == norm(t) for w in want for t in tags):
            return True, "正規", None
        if has_any(title, b["keywords"]):
            return True, "轉賣", None
        return False, None, None
    if platform == "pchome":
        if has_any(title, b["keywords"]) or has_any(title, b.get("pchome_lines", [])):
            shop = (p.get("shop") or {}).get("name") or ""
            return True, ("正規" if shop.startswith("PChome") else "轉賣"), None
        return False, None, None
    # shopee：標題或網址含品牌名
    if has_any(title + " " + (p.get("url") or ""), b["keywords"] + [b["name"]]):
        stype = seller_type(p, b)
        return True, ("正規" if stype in ("官方賣場", "連鎖與平台直營") else "轉賣"), stype
    return False, None, None


# ── 商品網址 ──────────────────────────────────────────────────────────────────
# momo 第三方商品（TP 開頭）走 /product/ 會 404，要用 /TP/前 9 碼/goodsDetail/ 路徑（2026-10-31 實測）；
# 蝦皮短網址 -i.賣場.商品 會開到「找不到頁面」，改用 /product/賣場/商品；沒有賣場編號就不給網址
def product_url(plat, pid, shop_id=None, url=None):
    pid = str(pid or "")
    if not pid:
        return None
    if plat == "momo":
        return (f"https://www.momoshop.com.tw/TP/{pid[:9]}/goodsDetail/{pid}" if pid.startswith("TP")
                else f"https://www.momoshop.com.tw/product/{pid}")
    if plat == "pchome":
        return f"https://24h.pchome.com.tw/prod/{pid}"
    if url and "-i." in url and not url.startswith("https://shopee.tw/-i."):
        return url
    return f"https://shopee.tw/product/{shop_id}/{pid}" if shop_id not in (None, "", "None") else None


# ── 主流程 ────────────────────────────────────────────────────────────────────
COMPACT_FIELDS = {
    "momo": ("product_id", "title", "price", "original", "rating", "review_count", "sold"),
    "pchome": ("product_id", "title", "price", "shop"),
    "shopee": ("product_id", "title", "price", "original", "shop", "shop_id"),
}


def from_compact(plat, rows):
    """精簡格式的一列 → 和工具回傳相同結構的商品"""
    out = []
    for r in rows:
        f = dict(zip(COMPACT_FIELDS[plat], r))
        sold = f.get("sold")
        out.append({
            "product_id": str(f["product_id"]), "title": f["title"],
            "url": None,
            "price": {"amount": f.get("price"), "original": f.get("original")},
            "rating": {"score": f.get("rating"), "count": f.get("review_count")},
            "units_sold": {"value": sold, "available": sold is not None},
            "shop": {"name": f.get("shop") or ("PChome" if plat == "pchome" else None), "id": f.get("shop_id")},
        })
    return out


def load_raw(raw_dir):
    idx_path = os.path.join(raw_dir, "_index.json")
    compact = os.path.join(raw_dir, "compact.json")
    if not os.path.exists(idx_path) and os.path.exists(compact):
        # 手抄的精簡格式（格式見 project-instructions.md「資料蒐集」）
        data = {}
        for key, rows in json.load(open(compact, encoding="utf-8")).items():
            plat, query = key.replace("\\|", "|").split("|", 1)
            if plat not in COMPACT_FIELDS:
                sys.exit(f"🛑 compact.json 的「{key}」平台名稱要是 momo、pchome 或 shopee")
            data[(plat, norm(query))] = {"products": from_compact(plat, rows)}
        return data
    if not os.path.exists(idx_path):
        sys.exit(f"🛑 {raw_dir} 裡沒有 compact.json，先照 project-instructions.md「資料蒐集」第 1 步手抄搜尋結果")
    index = json.load(open(idx_path, encoding="utf-8"))
    data = {}
    for r in index:
        plat = r["tool"].split("_")[0]
        data[(plat, norm(r["query"]))] = json.load(open(os.path.join(raw_dir, r["file"]), encoding="utf-8"))
    return data


UNIT_MIN = 3      # 客戶可換算的正規通路商品少於這個數，每單位售價頁不出（和 build.js 一致）


def fmt_price(v, unit_price=False):
    if v is None:
        return "—"
    return f"{v:g} 元" if unit_price else f"{v:,.0f} 元"


def summary_draft(brands, stats, unit, unit_comparable):
    """摘要確認的版面：指標表、段落標題、自動處理、回覆方式都填好，Claude 只補〔〕裡的文字"""
    names = [b["name"] for b in brands]
    client = next((b["name"] for b in brands if b.get("client")), names[0])
    st_ = [stats[n] for n in names]
    use_unit = unit_comparable and (stats[client]["unit_price_n"] or 0) >= UNIT_MIN
    rows = [("有上架的平台", [s["platform_kpi"]["value"] + (f"（{'、'.join(s['platform_kpi']['missing'])} 查無）"
                                                         if s["platform_kpi"]["missing"] else "") for s in st_])]
    if use_unit:
        rows.append((f"{unit}售價中位數", [fmt_price(s["unit_price_median"], True) for s in st_]))
    else:
        rows.append(("售價中位數", [fmt_price(s["price_median"]) for s in st_]))
    rows.append(("momo 最高銷量級距", [s["momo_max_band"] or "—" for s in st_]))
    off_cat = sum(s["off_category"] for s in st_)
    removed = sum(s["removed"] for s in st_) - off_cat
    category = stats["_unit"].get("category")
    resale = sum(s["resale"] for s in st_)
    third = (stats["_unit"]["skip_note"] or stats["_unit"]["equiv_note"]
             or ("買家評論頁本次不產出"))
    lines = [
        "### 📊 分析摘要",
        "**〔一句話結論，30 字內〕**",
        "",
        "| 指標 | " + " | ".join(names) + " |",
        "|---|" + "---|" * len(names),
        *["| " + label + " | " + " | ".join(vals) + " |" for label, vals in rows],
        "",
        "**🔍 三個發現**",
        "1. 〔一行，40 字內〕", "2. 〔一行，40 字內〕", "3. 〔一行，40 字內〕",
        "",
        "**🎯 三個建議**",
        "1. 〔一行，40 字內〕", "2. 〔一行，40 字內〕", "3. 〔一行，40 字內〕",
        "",
        "**🛠️ 自動處理**",
        *([f"- 排除 {removed} 筆其他品牌商品與重複頁面"] if removed else []),
        *([f"- 排除 {off_cat} 筆非{category}商品（同品牌的其他品項）"] if off_cat else []),
        *([f"- 第三方轉賣 {resale} 筆只放在通路頁，不進價格比較"] if resale else []),
        f"- {third}",
        "",
        "✅ 沒問題 → 回「**確認**」，我就產出簡報",
        "✏️ 要調整 → 直接寫要改哪一點",
    ]
    return "\n".join(lines) + "\n"


SIZE_GAP_MAX = 3   # 各品牌常見單件容量的最大／最小超過這個倍數，每單位售價頁不出
PLATFORM_LABEL = {"momo": "momo", "pchome": "PChome", "shopee": "蝦皮"}


def platform_kpi(matrix, brand):
    """上架平台數與是哪幾個平台；摘要確認表與簡報第 1 個指標共用"""
    rows = [m for m in matrix if m["brand"] == brand]
    have = [PLATFORM_LABEL[m["platform"]] for m in rows if m["kept"] > 0]
    miss = [PLATFORM_LABEL[m["platform"]] for m in rows if m["kept"] == 0]
    if not have:
        sub = "三個平台都查無自家商品"
    elif not miss:
        sub = "momo、PChome、蝦皮都有自家商品"
    else:
        sub = f"有：{'、'.join(have)}｜{'、'.join(miss)} 查無"
    return {"label": "有上架的平台", "value": f"{len(have)}／3", "sub": sub, "missing": miss}


# ── 跨平台同品比價 ────────────────────────────────────────────────────────────
# 同品牌、同總容量的正規通路商品，標題相似度達門檻才算同一品項；同平台有多筆符合時取最低售價
SIM_MIN = 0.25
COMPARE_MAX = 6
NOISE = ("官方", "公司貨", "正品", "現貨", "免運", "限定", "獨家", "momo", "pchome", "蝦皮", "直營", "旗艦")


def bigrams(title, b):
    t = re.sub(r"【[^】]*】", "", str(title or "")).lower()
    for w in list(b["keywords"]) + [b["name"]] + list(NOISE):
        t = t.replace(str(w).lower(), "")
    t = re.sub(r"\d+(\.\d+)?\s*(ml|毫升|公升|l|kg|公斤|g|公克|克|入|包|瓶|罐|盒|組|顆|粒)?", "", t)
    t = re.sub(r"[^a-z\u4e00-\u9fff]", "", t)
    return {t[i:i + 2] for i in range(len(t) - 1)}


def sim(a, b_, brand):
    x, y = bigrams(a["title"], brand), bigrams(b_["title"], brand)
    return len(x & y) / len(x | y) if x and y else 0


def display_name(its, b):
    """品項名稱：取最短的標題（去掉【】標記）；沒寫品牌就補上"""
    names = sorted((re.sub(r"\s+", " ", re.sub(r"【[^】]*】", "", i["title"] or "")).strip() for i in its), key=len)
    n = names[0] if names and names[0] else (its[0]["title"] or "")
    return n if has_any(n, list(b["keywords"]) + [b["name"]]) else f"{b['name']} {n}"


def compare_rows(items, brands):
    rows, used = [], []
    for b in brands:
        pool = [i for i in items if i["brand"] == b["name"] and i["channel"] == "正規" and i["qty"] and i["price"]]
        for qty in sorted({i["qty"] for i in pool}):
            group = [i for i in pool if i["qty"] == qty]
            while len({i["platform"] for i in group}) >= 2:
                best = None
                for anchor in sorted(group, key=lambda i: (PLATFORMS.index(i["platform"]), -(i["sold_value"] or 0))):
                    pick = {}
                    for pl in PLATFORMS:
                        cands = [i for i in group if i["platform"] == pl and (i is anchor or sim(anchor, i, b) >= SIM_MIN)]
                        if cands:
                            pick[pl] = min(cands, key=lambda i: i["price"])
                    if best is None or len(pick) > len(best):
                        best = pick
                if len(best) < 2:
                    break
                its = list(best.values())
                rows.append({"item": display_name(its, b), **{pl: (best[pl]["price"] if pl in best else None) for pl in PLATFORMS},
                             "_client": bool(b.get("client")), "_sold": max(i["sold_value"] or 0 for i in its), "_items": its})
                group = [i for i in group if not any(i is j for j in its)]
    rows.sort(key=lambda r: (not r["_client"], -sum(r[pl] is not None for pl in PLATFORMS), -r["_sold"]))
    rows = rows[:COMPARE_MAX]
    for r in rows:
        used += r.pop("_items")
        r.pop("_client"), r.pop("_sold")
    return rows, used


def pct(a, b):
    return round(a / b * 100) if b else None


def med(xs):
    return round(st.median(xs), 2) if xs else None


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--raw", default="raw")
    ap.add_argument("--brands", default="brands.json")
    ap.add_argument("--out", default=".")
    ap.add_argument("--lock", action="store_true", help="摘要前最後一次執行：跑完後鎖定數字")
    ap.add_argument("--unlock", action="store_true", help="對方要求修改摘要時才用：允許重跑，跑完維持鎖定")
    a = ap.parse_args()

    lock_path = os.path.join(a.out, ".summary_locked")
    was_locked = os.path.exists(lock_path)
    if was_locked and not a.unlock:
        print("🛑 摘要已確認，數字已鎖定：不重跑分析、不改 brands.json，簡報照已確認的 deck_parts.json 與 stats.json 產出。"
              "只有對方在摘要確認時要求修改，才加 --unlock 重跑。")
        sys.exit(2)

    cfg = json.load(open(a.brands, encoding="utf-8"))
    unit = cfg.get("unit", "每件")
    overrides = cfg.get("unit_overrides", {})
    g_ml_same = bool(cfg.get("g_ml_same"))
    category = cfg.get("category")
    cat_words = cfg.get("category_keywords") or []
    brands = cfg["brands"]
    if isinstance(brands, dict):   # 也接受以品牌名當鍵的寫法
        brands = [{"name": k, **v} for k, v in brands.items()]
    for b in brands:
        b.setdefault("keywords", [b["name"]])
        b.setdefault("momo_tag", b["name"])
    raw = load_raw(a.raw)

    items, matrix, warns = [], [], []
    for b in brands:
        for plat in PLATFORMS:
            q = (b.get("queries") or {}).get(plat, b["query"])
            obj = raw.get((plat, norm(q)))
            if obj is None:
                warns.append(f"{b['name']} 在 {plat} 沒有找到查詢「{q}」的資料")
                matrix.append({"brand": b["name"], "platform": plat, "sampled": 0, "kept": 0, "removed": 0, "off_category": 0, "resale": 0,
                               "is_client": bool(b.get("client"))})
                continue
            prods = obj.get("products") or []
            seen, kept, removed, dup, resale, off_cat = set(), 0, 0, 0, 0, 0
            for p in prods:
                ok, channel, stype = classify(plat, p, b)
                if not ok:
                    removed += 1
                    continue
                # 同品牌但不是這次比較的品項（例：比燕麥片時的杏仁粉、玉米片），各頁都不計
                if cat_words and not has_any(p.get("title"), cat_words):
                    off_cat += 1
                    continue
                price = (p.get("price") or {}).get("amount")
                key = (norm(p.get("title")), price)
                if key in seen:
                    dup += 1
                    continue
                seen.add(key)
                kept += 1
                resale += channel == "轉賣"
                orig = (p.get("price") or {}).get("original")
                sv = sold_value(p) if plat == "momo" else None
                oid = f"{plat}:{p.get('product_id')}"
                if oid in overrides:
                    qty, qnote, size = overrides[oid], "人工指定", None
                else:
                    qty, qnote, size = unit_qty(p.get("title") or "", unit, g_ml_same)
                rating = p.get("rating") or {}
                items.append({
                    "platform": plat, "brand": b["name"], "client": bool(b.get("client")),
                    "product_id": p.get("product_id"), "title": p.get("title"),
                    "url": product_url(plat, p.get("product_id"), (p.get("shop") or {}).get("id"), p.get("url")),
                    "channel": channel, "seller_type": stype,
                    "shop": (p.get("shop") or {}).get("name"), "shop_id": (p.get("shop") or {}).get("id"),
                    "price": price, "original": orig,
                    "discount_pct": round((1 - price / orig) * 100, 1) if price and orig and orig > price else None,
                    "sold_value": sv, "sold_band": band_text(sv), "sold_tier": tier(sv) if plat == "momo" else None,
                    # 評分、評論數只有 momo 可用；蝦皮是賣場層級，不引用
                    "rating": rating.get("score") if plat == "momo" else None,
                    "review_count": rating.get("count") if plat == "momo" else None,
                    "qty": qty, "qty_note": qnote, "size": size,
                    "unit_price": round(price / qty, 3) if price and qty and channel == "正規" else None,
                })
            matrix.append({"brand": b["name"], "platform": plat, "sampled": len(prods), "kept": kept,
                           "removed": removed + dup + off_cat, "off_category": off_cat, "resale": resale,
                           "is_client": bool(b.get("client"))})

    # ── 區塊 ──
    def of(brand, plat=None, channel=None):
        return [i for i in items if i["brand"] == brand and (plat is None or i["platform"] == plat)
                and (channel is None or i["channel"] == channel)]

    # 各品牌常見的單件容量差太多（例：40ml 精華乳對 473ml 身體乳），每單位售價不能直接比較，整頁不出
    sizes = {b["name"]: med([i["size"] for i in of(b["name"], channel="正規") if i["size"]]) for b in brands}
    known = {n: v for n, v in sizes.items() if v}
    size_gap = round(max(known.values()) / min(known.values()), 1) if len(known) > 1 else 1
    unit_comparable = unit in UNIT_RE and size_gap <= SIZE_GAP_MAX

    sellers, unit_rows, disc_rows, band_rows, stats = [], [], [], [], {}
    for b in brands:
        n = b["name"]
        shops = {}
        for i in of(n, "shopee"):
            shops.setdefault(i["seller_type"], set()).add(i["shop_id"] or i["shop"])
        for t in SELLER_TYPES:
            if shops.get(t):
                sellers.append({"brand": n, "seller_type": t, "shop_count": len(shops[t])})
        reg = of(n, channel="正規")
        ups = [i["unit_price"] for i in reg if i["unit_price"]]
        if ups and unit_comparable:
            unit_rows.append({"label": n, "n": len(ups), "median": med(ups),
                              "by_platform": {pl: sum(1 for i in reg if i["platform"] == pl and i["unit_price"]) for pl in PLATFORMS},
                              **({"is_client": True} if b.get("client") else {})})
        dis = [i["discount_pct"] for i in reg if i["discount_pct"]]
        disc_rows.append({"brand": n, "n_discounted": len(dis), "n_total": len(reg),
                          "avg_pct": round(sum(dis) / len(dis)) if dis else 0})
        momo = of(n, "momo")
        for t in TIERS:
            c = sum(1 for i in momo if i["sold_tier"] == t)
            if c:
                band_rows.append({"brand": n, "tier": t, "count": c})
        vals = [i["sold_value"] for i in momo if i["sold_value"] is not None]
        stats[n] = {
            "client": bool(b.get("client")),
            "platforms_listed": sum(1 for m in matrix if m["brand"] == n and m["kept"] > 0),
            "platforms_regular": len({i["platform"] for i in reg}),
            "price_median": med([i["price"] for i in reg if i["price"]]),
            "unit_price_median": med(ups), "unit_price_n": len(ups),
            "unit_unparsed": sum(1 for i in reg if i["qty"] is None),
            "momo_max_band": band_text(max(vals)) if vals else None,
            "momo_10man_count": sum(1 for v in vals if v >= 100000),
            "discount_avg_pct": round(sum(dis) / len(dis), 1) if dis else None,
            "kept": len(of(n)), "regular": len(reg), "resale": len(of(n, channel="轉賣")),
            "removed": sum(m["removed"] for m in matrix if m["brand"] == n),
            "off_category": sum(m["off_category"] for m in matrix if m["brand"] == n),
            "typical_size": sizes.get(n),
            # 比例一律由這裡算，摘要直接引用，不心算
            "resale_pct": pct(len(of(n, channel="轉賣")), len(of(n))),
            "by_platform": {pl: {"kept": len(of(n, pl)), "regular": len(of(n, pl, "正規")), "resale": len(of(n, pl, "轉賣")),
                                 "resale_pct": pct(len(of(n, pl, "轉賣")), len(of(n, pl)))} for pl in PLATFORMS},
        }

    top = sorted([i for i in items if i["platform"] == "momo" and i["channel"] == "正規"],
                 key=lambda i: (-(i["sold_value"] or 0), -(i["review_count"] or 0)))[:6]
    clients = [b["name"] for b in brands if b.get("client")]
    client = clients[0] if clients else brands[0]["name"]
    compare, compare_items = compare_rows(items, brands)
    # 附錄：04-1 與 03-3 引用到的商品，依品牌、平台排序；其他頁引用的商品由 Claude 從 items.json 補
    cited, appx = top + compare_items, []
    for i in sorted(cited, key=lambda i: ([b["name"] for b in brands].index(i["brand"]), PLATFORMS.index(i["platform"]))):
        if any(a["platform"] == i["platform"] and a["product_id"] == i["product_id"] for a in appx):
            continue
        appx.append(i)
    for n in stats:
        stats[n]["platform_kpi"] = platform_kpi(matrix, n)
    parts = {
        # 品牌名單與客戶名一起輸出，簡報各頁的品牌名才會和這裡一致
        "client": client,
        "category": category,
        "category_keywords": cat_words,
        "brands": [{"name": b["name"], "is_client": bool(b.get("client")),
                    "momo_tag": b["momo_tag"] if isinstance(b["momo_tag"], str) else b["momo_tag"][0]} for b in brands],
        # 第 1 個指標固定是上架平台，Claude 再補 2 個
        "summary_kpis": [{k: v for k, v in stats[client]["platform_kpi"].items() if k != "missing"}],
        "channel_matrix": {"rows": matrix},
        "shopee_sellers": {"rows": sellers},
        "unit_price": {"unit": unit, "market_median": None, "rows": unit_rows},
        "price_compare": {"rows": compare},
        "discount": {"rows": disc_rows},
        "momo_top": {"rows": [{"rank": k + 1, "brand": i["brand"], "title": i["title"], "price": i["price"],
                               "sold_band": i["sold_band"], "rating": i["rating"], "review_count": i["review_count"],
                               "url": i["url"]} for k, i in enumerate(top)]},
        "sold_band": {"tiers": TIERS, "rows": band_rows},
        "appendix_products": {"rows": [{"idx": k + 1, "brand": i["brand"], "platform": i["platform"], "title": i["title"],
                                        "price": i["price"], "sold_band": i["sold_band"], "url": i["url"]}
                                       for k, i in enumerate(appx)]},
    }

    os.makedirs(a.out, exist_ok=True)
    stats["_unit"] = {"unit": unit, "comparable": unit_comparable, "category": category, "category_keywords": cat_words, "size_gap": size_gap, "typical_sizes": sizes,
                      "g_ml_same": g_ml_same,
                      "equiv_note": "液體商品以 1 克視同 1 毫升換算" if g_ml_same and unit in ("每毫升", "每克") else None,
                      "skip_note": None if unit_comparable or unit not in UNIT_RE else
                      f"各品牌常見容量相差 {size_gap:g} 倍，{unit}售價不宜直接比較，本次不出每單位售價頁"}
    for fn, obj in (("items.json", items), ("deck_parts.json", parts), ("stats.json", stats)):
        with open(os.path.join(a.out, fn), "w", encoding="utf-8") as f:
            json.dump(obj, f, ensure_ascii=False, indent=1)
    with open(os.path.join(a.out, "summary_draft.md"), "w", encoding="utf-8") as f:
        f.write(summary_draft(brands, stats, unit, unit_comparable))
    if a.lock or was_locked:
        with open(lock_path, "w", encoding="utf-8") as f:
            f.write("摘要數字已鎖定\n")

    # ── 只印統計 ──
    print(f"比價單位：{unit}；商品 {len(items)} 筆（已剔除他牌與重複頁面）")
    for n, s in stats.items():
        if n.startswith("_"):
            continue
        print(f"■ {n}{'（客戶）' if s['client'] else ''}：上架 {s['platforms_listed']}／3 平台，正規 {s['regular']}、轉賣 {s['resale']}、剔除 {s['removed']}")
        print(f"   售價中位數 {s['price_median']} 元｜{unit}中位數 {s['unit_price_median']}（{s['unit_price_n']} 筆，換算不出 {s['unit_unparsed']} 筆）"
              f"｜momo 最高 {s['momo_max_band']}（10 萬件以上 {s['momo_10man_count']} 件）｜平均折扣 {s['discount_avg_pct']}%")
        bp = s["by_platform"]
        print("   轉賣占比：" + "、".join(f"{pl} {bp[pl]['resale']}／{bp[pl]['kept']}（{bp[pl]['resale_pct']}%）" for pl in PLATFORMS if bp[pl]["kept"])
              + f"；全部 {s['resale_pct']}%")
    print(f"跨平台同品比價：自動配對 {len(compare)} 列" + ("（少於 3 列，可用 compare_prices 結果補，頁尾要寫明）" if len(compare) < 3 else ""))
    if unit in UNIT_RE and not unit_comparable:
        print(f"⚠ {stats['_unit']['skip_note']}：" + "、".join(f"{n} {v:g}" for n, v in known.items()))
    for m in matrix:
        if m["sampled"] and m["kept"] == 0:
            print(f"⚠ {m['brand']} 在 {m['platform']} 查無自家商品（抽樣 {m['sampled']} 筆全數剔除）")
    if unit in UNIT_RE:
        bad = [i for i in items if i["client"] and i["channel"] == "正規" and i["qty"] is None][:8]
        if bad:
            print("需要人工確認換算量的客戶商品（可寫進 brands.json 的 unit_overrides）：")
            for i in bad:
                print(f"   {i['platform']}:{i['product_id']}  {i['qty_note']}  {str(i['title'])[:40]}")
    for w in warns:
        print("⚠ " + w)


if __name__ == "__main__":
    main()
