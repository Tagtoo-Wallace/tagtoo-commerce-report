#!/usr/bin/env python3
# SPDX-License-Identifier: Apache-2.0
"""
make_deck.py — 把已確認的摘要與已鎖定的數字合成 deck_data.json
-----------------------------------------------------------------
步驟 ④：Claude 把 summary_draft.md 的〔〕填好、存成 summary.md，先跑
        python3 make_deck.py --check-summary
        沒有 🛑 才把 summary.md 全文照貼給對方確認。
步驟 ⑤：Claude 只寫 deck_extra.json（需要判斷的欄位，格式見 deck_extra.example.json），再跑
        python3 make_deck.py
        產出 deck_data.json，接著執行 build.js。

合併規則：
  - 數字區塊一律取自 deck_parts.json（analyze.py 已鎖定），deck_extra.json 不可覆蓋。
  - summary_findings、insights[].body 取摘要「三個發現」原文；actions[].body 取「三個建議」原文。
  - summary_kpis 三張卡取自 stats.json，和摘要指標表同一組數字。
  - scope_items、appendix_notes 由本程式依 stats.json 產出（appendix_notes 可在 deck_extra.json 改寫）。
  - deck_extra.json 的 appendix_add 只寫商品代號（「平台:商品編號」），標題、售價、級距、網址由 items.json 帶入。
所有問題一次列出：🛑 要修正後重跑，⚠ 逐則確認是否可接受。

用法（在工作資料夾執行）：python3 <pkg>/docs/make_deck.py [--dir .] [--check-summary]
"""
import argparse, datetime, json, os, re, sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from analyze import UNIT_MIN, fmt_price, PLATFORMS, PLATFORM_LABEL  # noqa: E402

TITLE_MAX, LINE_MAX, ITEM_TITLE_MAX, ITEM_BODY_MAX = 24, 40, 20, 60
PAGE_CODE = re.compile(r"\b(0[1-5]-\d|A-[12])\b")
# deck_extra.json 只能有這些鍵；數字區塊與摘要原文不在其中
EXTRA_KEYS = {"page_titles", "page_sources", "assortment", "title_claims", "insight_titles",
              "action_titles", "appendix_add", "appendix_notes", "fetched_at"}
LOCKED_KEYS = {"client", "brands", "category", "category_keywords", "summary_kpis", "summary_findings",
               "channel_matrix", "shopee_sellers", "unit_price", "price_compare", "discount", "momo_top",
               "sold_band", "appendix_products", "insights", "actions", "scope_items", "mode"}

errors, warns = [], []
stop = errors.append
warn = warns.append


def clen(s):
    return len(str(s or ""))


def load(path, required=True):
    if not os.path.exists(path):
        if required:
            stop(f"找不到 {os.path.basename(path)}")
        return None
    try:
        return json.load(open(path, encoding="utf-8"))
    except json.JSONDecodeError as e:
        stop(f"{os.path.basename(path)} 不是合法的 JSON：{e}")
        return None


def parse_summary(path):
    """讀 summary.md：一句話結論、三個發現、三個建議"""
    if not os.path.exists(path):
        stop("找不到 summary.md：步驟 ④ 要先把填好的摘要存成 summary.md")
        return None
    text = open(path, encoding="utf-8").read()
    if "〔" in text or "〕" in text:
        stop("summary.md 還有〔〕沒填：" + "、".join(sorted(set(re.findall(r"〔[^〕]*〕", text)))[:5]))
    lines = [l.rstrip() for l in text.splitlines()]
    out, sec = {"conclusion": None, "findings": [], "actions": []}, None
    for l in lines:
        s = l.strip()
        if "三個發現" in s:
            sec = "findings"
            continue
        if "三個建議" in s:
            sec = "actions"
            continue
        if "自動處理" in s or s.startswith("✅") or s.startswith("✏️"):
            sec = None
            continue
        m = re.match(r"^\*\*(.+)\*\*$", s)
        if m and out["conclusion"] is None and sec is None and not s.startswith("**🔍"):
            out["conclusion"] = m.group(1).strip()
            continue
        m = re.match(r"^\d+[.、．]\s*(.+)$", s)
        if m and sec:
            out[sec].append(m.group(1).replace("**", "").strip())
    for key, name in (("findings", "三個發現"), ("actions", "三個建議")):
        n = len(out[key])
        if n != 3:
            stop(f"summary.md 的「{name}」要剛好 3 則，讀到 {n} 則")
        for i, t in enumerate(out[key], 1):
            if clen(t) > LINE_MAX:
                warn(f"{name}第 {i} 則 {clen(t)} 字，超過 {LINE_MAX} 字（簡報第 3 頁會擠）")
            if PAGE_CODE.search(t):
                stop(f"{name}第 {i} 則出現頁代號「{PAGE_CODE.search(t).group(0)}」，摘要不寫頁代號")
    if not out["conclusion"]:
        stop("summary.md 讀不到一句話結論（「### 📊 分析摘要」下一行的粗體句）")
    return out


def kpi_cards(parts, stats, client, rivals):
    """三張數字卡，和摘要指標表同一組數字；不另算新數字"""
    u = stats["_unit"]
    unit = u["unit"]
    use_unit = u["comparable"] and (stats[client]["unit_price_n"] or 0) >= UNIT_MIN
    field, label = (("unit_price_median", f"{unit}售價中位數") if use_unit else ("price_median", "售價中位數"))
    fmt = (lambda v: fmt_price(v, use_unit))
    sub = None
    vals = [(r, stats[r][field]) for r in rivals if stats[r][field] is not None]
    if len(vals) == 1:
        sub = f"{vals[0][0]} {fmt(vals[0][1])}"
    elif len(vals) > 1:
        lo, hi = min(v for _, v in vals), max(v for _, v in vals)
        sub = f"競品 {fmt(lo)}" if lo == hi else f"競品 {fmt(lo)[:-2]}–{fmt(hi)}"
    elif use_unit:
        sub = "依標題容量換算"
    band = stats[client]["momo_max_band"]
    return [parts["summary_kpis"][0],
            {"label": label, "value": fmt(stats[client][field]), **({"sub": sub} if sub else {})},
            {"label": "momo 最高銷量級距", "value": band or "—", **({"sub": "累計、第三方估計"} if band else {})}]


def scope_items(parts, stats, client, rivals, fetched):
    u = stats["_unit"]
    cat = parts.get("category")
    rows = parts["channel_matrix"]["rows"]
    per = []
    for pl in PLATFORMS:
        rs = [r for r in rows if r["platform"] == pl]
        per.append(f"{PLATFORM_LABEL[pl]} {sum(r['sampled'] for r in rs)}／{sum(r['removed'] for r in rs)}／{sum(r['kept'] for r in rs)}")
    return [
        f"研究品牌：{client}；比較競品：{'、'.join(rivals)}" if rivals else f"研究品牌：{client}（不比較競品）",
        f"搜尋字：「品牌名 {cat}」，台灣市場" if cat else "搜尋字：品牌名，台灣市場",
        "抽樣／剔除／留下筆數：" + "、".join(per),
        "比價單位：每件（本品類以件計價）" if u["unit"] == "每件" else f"比價單位：{u['unit']}，依商品標題容量換算",
        f"資料擷取日：{fetched}",
    ]


def appendix_notes(stats, client):
    u = stats["_unit"]
    unit = u["unit"]
    if unit == "每件":
        unit_note = "本品類以件計價，不另比較每單位售價"
    elif u.get("skip_note"):
        unit_note = u["skip_note"]
    elif (stats[client]["unit_price_n"] or 0) < UNIT_MIN:
        unit_note = f"{client}的商品標題多未標示容量，本次不比較{unit}售價"
    else:
        unit_note = f"{unit}售價依標題容量換算，換算不出的商品不列入" + (f"；{u['equiv_note']}" if u.get("equiv_note") else "")
    return [
        "所有數字為擷取當下的公開頁面，沒有時間軸，不代表期間銷量",
        "momo 銷量為平台顯示的累計級距，屬第三方估計",
        "價格頁只計正規通路（品牌官方、平台自營、連鎖通路），不含第三方轉賣",
        unit_note,
        "蝦皮評論本次未取得；蝦皮評論數與商品評分不引用",
    ]


def fetched_date(d):
    p = os.path.join(d, "raw", "compact.json")
    ts = os.path.getmtime(p) if os.path.exists(p) else datetime.datetime.now().timestamp()
    tw = datetime.timezone(datetime.timedelta(hours=8))
    return datetime.datetime.fromtimestamp(ts, tw).strftime("%Y-%m-%d")


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--dir", default=".", help="工作資料夾（analyze.py 的輸出位置）")
    ap.add_argument("--check-summary", action="store_true", help="步驟 ④：只檢查 summary.md，不產出 deck_data.json")
    a = ap.parse_args()
    d = a.dir

    if not os.path.exists(os.path.join(d, ".summary_locked")):
        stop("數字還沒鎖定：印摘要前最後一次分析要加 --lock")
    summ = parse_summary(os.path.join(d, "summary.md"))

    if a.check_summary:
        return finish(None, d, summ)

    parts = load(os.path.join(d, "deck_parts.json"))
    stats = load(os.path.join(d, "stats.json"))
    items = load(os.path.join(d, "items.json"))
    extra = load(os.path.join(d, "deck_extra.json"))
    if errors or not (parts and stats and items and extra is not None and summ):
        return finish(None, d, summ)

    bad = sorted(set(extra) - EXTRA_KEYS)
    for k in bad:
        stop(f"deck_extra.json 不可有「{k}」" + ("：這是已鎖定的數字或已確認的摘要，由本程式帶入" if k in LOCKED_KEYS else "：不認得的欄位"))

    client = parts["client"]
    rivals = [b["name"] for b in parts["brands"] if not b["is_client"]]
    fetched = extra.get("fetched_at") or fetched_date(d)

    # 頁標題：沒寫 01-1 時，一句話結論在 24 字內就沿用
    titles = dict(extra.get("page_titles") or {})
    if not titles.get("01-1"):
        if summ["conclusion"] and clen(summ["conclusion"]) <= TITLE_MAX:
            titles["01-1"] = summ["conclusion"]
        else:
            stop(f"page_titles 沒有 01-1，一句話結論 {clen(summ['conclusion'])} 字超過 {TITLE_MAX} 字不能直接沿用，請另寫")
    for k, t in titles.items():
        if clen(t) > TITLE_MAX:
            stop(f"page_titles {k} 標題 {clen(t)} 字，超過 {TITLE_MAX} 字：「{t}」")
        if PAGE_CODE.search(str(t)):
            stop(f"page_titles {k} 標題出現頁代號，請改寫")

    it = extra.get("insight_titles") or []
    at = extra.get("action_titles") or []
    if len(it) != 3:
        stop(f"insight_titles 要剛好 3 則（title、evidence），目前 {len(it)} 則")
    if len(at) != 3:
        stop(f"action_titles 要剛好 3 則，目前 {len(at)} 則")
    insights, actions = [], []
    for i, (x, body) in enumerate(zip(it, summ["findings"]), 1):
        x = x if isinstance(x, dict) else {"title": x}
        if clen(x.get("title")) > ITEM_TITLE_MAX:
            stop(f"insight_titles 第 {i} 則標題 {clen(x.get('title'))} 字，超過 {ITEM_TITLE_MAX} 字")
        if not x.get("evidence"):
            stop(f"insight_titles 第 {i} 則沒有 evidence（寫依據頁代號，例「見 02-1、03-2」）")
        insights.append({"title": x.get("title"), "body": body, "evidence": x.get("evidence")})
    for i, (t, body) in enumerate(zip(at, summ["actions"]), 1):
        t = t.get("title") if isinstance(t, dict) else t
        if clen(t) > ITEM_TITLE_MAX:
            stop(f"action_titles 第 {i} 則標題 {clen(t)} 字，超過 {ITEM_TITLE_MAX} 字")
        actions.append({"title": t, "body": body})

    # 03-1 最低售價要是該品牌真的出現過的正規通路售價
    assortment = extra.get("assortment") or []
    for i, r in enumerate(assortment):
        prices = {x["price"] for x in items if x["brand"] == r.get("brand") and x["channel"] == "正規"}
        if r.get("price_min") is not None and r.get("price_min") not in prices:
            warn(f"assortment 第 {i} 列 {r.get('brand')}「{r.get('line')}」最低售價 {r.get('price_min')} 不在該品牌正規通路售價裡，請對照 items.json")

    # A-2：analyze.py 已列 04-1、03-3 引用的商品；其他頁引用的只給代號，由 items.json 帶入
    appx = [dict(r) for r in parts["appendix_products"]["rows"]]
    idx = {f"{x['platform']}:{x['product_id']}": x for x in items}
    for code in extra.get("appendix_add") or []:
        x = idx.get(str(code))
        if not x:
            stop(f"appendix_add「{code}」在 items.json 找不到（寫法：平台:商品編號，例 momo:12730446）")
            continue
        if any(r["platform"] == x["platform"] and r["title"] == x["title"] for r in appx):
            continue
        appx.append({"idx": 0, "brand": x["brand"], "platform": x["platform"], "title": x["title"],
                     "price": x["price"], "sold_band": x["sold_band"], "url": x["url"]})
    for k, r in enumerate(appx, 1):
        r["idx"] = k

    notes = extra.get("appendix_notes") or appendix_notes(stats, client)
    if not 1 <= len(notes) <= 5:
        stop(f"appendix_notes 要 1–5 條，目前 {len(notes)} 條")

    deck = {
        "client": client,
        "mode": "compare" if rivals else "single",
        "fetched_at": fetched,
        "category": parts.get("category"),
        "category_keywords": parts.get("category_keywords"),
        "brand_primary_color": "873222",
        "brands": parts["brands"],
        "page_titles": titles,
        "page_sources": extra.get("page_sources") or {},
        "summary_kpis": kpi_cards(parts, stats, client, rivals),
        "summary_findings": summ["findings"],
        "channel_matrix": parts["channel_matrix"],
        "shopee_sellers": parts["shopee_sellers"],
        "assortment": {"rows": assortment},
        "unit_price": parts["unit_price"],
        "price_compare": parts["price_compare"],
        "discount": parts["discount"],
        "momo_top": parts["momo_top"],
        "sold_band": parts["sold_band"],
        "review_quotes": {"sample_size": None, "rows": []},
        "title_claims": {"rows": extra.get("title_claims") or []},
        "insights": insights,
        "actions": actions,
        "scope_items": scope_items(parts, stats, client, rivals, fetched),
        "appendix_notes": notes,
        "appendix_products": {"rows": appx},
    }
    if not assortment:
        warn("assortment 是空的，產品線頁會寫「本次無資料」")
    if not deck["title_claims"]["rows"]:
        warn("title_claims 是空的，標題訴求頁會寫「本次無資料」")
    return finish(deck, d, summ)


def finish(deck, d, summ):
    for m in warns:
        print(f"⚠ {m}")
    for m in errors:
        print(f"🛑 {m}")
    if errors:
        print(f"🛑 共 {len(errors)} 個問題，修正後重跑；deck_data.json 沒有產出")
        sys.exit(2)
    if deck is None:
        print(f"✓ summary.md 檢查通過：結論 1 句、發現 {len(summ['findings'])} 則、建議 {len(summ['actions'])} 則")
        return
    out = os.path.join(d, "deck_data.json")
    with open(out, "w", encoding="utf-8") as f:
        json.dump(deck, f, ensure_ascii=False, indent=1)
    print(f"✓ deck_data.json：發現與建議沿用 summary.md 原文；引用商品 {len(deck['appendix_products']['rows'])} 筆")


if __name__ == "__main__":
    main()
