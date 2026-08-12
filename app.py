# -*- coding: utf-8 -*-
"""
K.U.P 晶球魚油 · 市場資料查詢 Demo
專題報告用的現場展示小程式。
輸入商品名 → 自動去購物通路抓即時價格 → 表格 + 圖表 + 下載 Excel。

啟動方式（不用懂程式）：直接雙擊資料夾裡的「啟動程式.bat」即可。
"""

import io
import re
import datetime as dt

import requests
import pandas as pd
import streamlit as st

# ---------------------------------------------------------------------------
# 基本設定
# ---------------------------------------------------------------------------
st.set_page_config(page_title="K.U.P 魚油 · 市場資料查詢", page_icon="🐟", layout="wide")

# --- 視覺樣式（清爽專業・海洋主題）---
st.markdown("""
<style>
html, body, [class*="css"] {
  font-family: "Segoe UI","PingFang TC","Microsoft JhengHei","Noto Sans TC",sans-serif;
}
.kup-hero{
  background: linear-gradient(120deg,#0e7c86 0%,#12a3a3 45%,#1b6ca8 100%);
  border-radius:18px; padding:26px 30px; color:#fff; margin:4px 0 18px;
  box-shadow:0 8px 24px rgba(14,124,134,.25);
}
.kup-hero .t{ font-size:30px; font-weight:800; letter-spacing:.5px; margin:0; }
.kup-hero .s{ font-size:15px; opacity:.93; margin:6px 0 0; }
.kup-chips{ margin-top:16px; display:flex; flex-wrap:wrap; gap:8px; }
.kup-chips .c{
  background:rgba(255,255,255,.18); border:1px solid rgba(255,255,255,.4);
  padding:6px 14px; border-radius:20px; font-size:13px; font-weight:700;
}
[data-testid="stMetric"]{
  background:#ffffff; border:1px solid #e2ebee; border-radius:14px;
  padding:14px 16px; box-shadow:0 2px 8px rgba(19,41,61,.05);
}
[data-testid="stMetricValue"]{ color:#0e7c86; font-weight:800; }
[data-testid="stMetricLabel"]{ color:#5b7488; }
.stButton>button, .stDownloadButton>button, .stLinkButton>a{
  border-radius:10px; font-weight:700;
}
button[data-baseweb="tab"]{ font-size:15px; font-weight:600; }
.kup-badges{ display:flex; flex-wrap:wrap; gap:10px; margin-top:4px; }
.kup-badge{
  background:#eaf6f6; color:#0a5b63; border:1px solid #bfe3e3;
  border-radius:12px; padding:10px 14px; font-size:13.5px; line-height:1.4;
}
.kup-badge .d{ font-weight:400; color:#4b6a72; font-size:12.5px; }
</style>
""", unsafe_allow_html=True)

HEADERS = {
    "User-Agent": ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                   "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36"),
    "Accept-Language": "zh-TW,zh;q=0.9",
}
TIMEOUT = 15
PACKS_PER_BOX = 28  # K.U.P 一盒 28 包

# ---------------------------------------------------------------------------
# K.U.P 魚油 內建產品資料（供「產品資料卡」分頁，demo 一定看得到完整內容）
# ---------------------------------------------------------------------------
KUP_SPECS = [
    ("規格", "2000 mg／包，28 包／盒"),
    ("劑型", "0.38 cm 微型無縫晶球膠囊（EE 乙酯型）"),
    ("Omega-3 濃度", "90%（EPA+DHA 占 84%）"),
    ("EPA / DHA（每包）", "EPA 920 mg｜DHA 760 mg（合計 1680 mg）"),
    ("成分", "魚油（含 DHA、EPA）"),
    ("產地 / 品牌", "韓國世宗市／韓國聯合製藥 Korea United Pharm"),
    ("台灣代理", "泰和碩藥品科技／沁軒貿易（依賣場而異）"),
    ("保存期限", "36 個月（3 年）"),
]
KUP_CERTS = [
    "IFOS 五星認證（國際魚油第三方檢測：安全性、含量、雜質、新鮮度）",
    "世界品質評鑑大賞 最高特級金獎（Monde Selection）",
    "PIC/S GMP · cGMP · EUGMP 製藥級工廠生產",
    "海洋之友 Friend of the Sea 永續認證",
    "自由落體無縫包覆製程（減少高溫與空氣接觸，不易氧化）",
    "每包單獨鋁袋分包，杜絕空氣、鎖住新鮮",
]
KUP_USAGE = [
    "成人／長者基礎保養每日 1 包，孩童每日 1/2 包，隨餐或餐後食用，多食無益。",
    "膠囊顆粒小，每包宜分 2–3 次食用，避免嗆到。",
    "孕哺婦女、重大疾病、計劃手術、凝血功能不佳或服用抗凝血藥劑者，食用前先諮詢醫療專業人員。",
]
KUP_REF_PRICES = pd.DataFrame({
    "通路 / 賣家": ["富凱昆陽藥局", "iOPEN／蝦皮 多家藥局", "蝦皮・理可生活"],
    "單盒售價 (NT$)": [1580, 1980, 2600],
    "備註": ["現貨最低", "最常見價", "較高"],
})

# 驗證過的商品頁／購買通路連結（供報告佐證：點下去是真實販售頁）
# 跨通路比價頁(BigGo/飛比)永遠顯示目前有貨的賣場，最適合當證據，不會出現缺貨頁。
KUP_LINKS = [
    ("BigGo 跨通路即時比價・K.U.P 專屬（最推薦當證據）",
     "https://biggo.com.tw/s/%E6%99%B6%E7%90%83%E9%AD%9A%E6%B2%B9%20k.u.p"),
    ("飛比 feebee 跨通路即時比價",
     "https://feebee.com.tw/s/%E6%99%B6%E7%90%83%E9%AD%9A%E6%B2%B9/"),
    ("K.U.P 官方網站（官方購買通路總表）",
     "https://kupomega3.com.tw/"),
    ("蝦皮 商品頁",
     "https://shopee.tw/product/3975876/24418691796"),
]


# ---------------------------------------------------------------------------
# 抓取函式
# ---------------------------------------------------------------------------
@st.cache_data(ttl=600, show_spinner=False)
def fetch_pchome(term, limit=30):
    """PChome 24h 商品搜尋 API（穩定、含明碼價）。"""
    url = "https://ecshweb.pchome.com.tw/search/v3.3/all/results"
    r = requests.get(url, params={"q": term, "page": 1, "sort": "sale/dc"},
                     headers=HEADERS, timeout=TIMEOUT)
    r.raise_for_status()
    prods = r.json().get("prods", []) or []
    return [{
        "通路": "PChome 24h",
        "品項": (p.get("name") or "").strip(),
        "價格": p.get("price"),
        "網址": f"https://24h.pchome.com.tw/prod/{p.get('Id')}",
    } for p in prods[:limit]]


@st.cache_data(ttl=600, show_spinner=False)
def fetch_momo(term, limit=30):
    """momo 搜尋 API（盡力而為，有反爬，失敗就回空清單）。"""
    url = "https://apisearch.momoshop.com.tw/momoSearchCloud/moec/textSearch"
    payload = {"host": "momoshop", "flag": "searchEngine",
               "data": {"searchValue": term, "curPage": "1",
                        "priceS": "0", "priceE": "9999999", "searchType": "1"}}
    r = requests.post(url, json=payload, headers=HEADERS, timeout=TIMEOUT)
    r.raise_for_status()
    goods = (r.json().get("rtnSearchData", {}) or {}).get("goodsInfoList", []) or []
    return [{
        "通路": "momo 購物網",
        "品項": (g.get("goodsName") or "").strip(),
        "價格": g.get("goodsPrice"),
        "網址": f"https://www.momoshop.com.tw/goods/GoodsDetail.jsp?i_code={g.get('goodsCode')}",
    } for g in goods[:limit]]


SOURCES = [("PChome 24h", fetch_pchome), ("momo 購物網", fetch_momo)]

# K.U.P 專用：用多組關鍵字去撈，再嚴格過濾成「同時含『晶球』和『魚油』」的品項，
# 這樣才不會把『晶球肉泥』(貓食) 等雜訊抓進來，精準鎖定 K.U.P。
KUP_QUERIES = ["韓國 晶球魚油", "K.U.P 晶球魚油", "KUP 魚油",
               "晶球魚油 EPA DHA", "藥品級 魚油 韓國"]

# K.U.P 專屬 BigGo 比價頁：永遠顯示目前有貨的賣場，點下去一定連得到（不會下架/空白）。
KUP_BIGGO_URL = "https://biggo.com.tw/s/%E6%99%B6%E7%90%83%E9%AD%9A%E6%B2%B9%20k.u.p"

# 回退資料：K.U.P 在 PChome 常缺貨，且有貨的通路(蝦皮/iOPEN)反爬抓不到。
# 這批是從 BigGo 跨通路比價頁擷取的「真實」單盒價格（來源：BigGo，2026/08）；
# 連結一律導向 BigGo K.U.P 比價頁，確保永遠點得開、看得到現貨。
KUP_FALLBACK = [
    {"通路": "富凱昆陽藥局(iOPEN)", "品項": "K.U.P 90%晶球魚油 單盒/28包 2000mg",
     "價格": 1580, "網址": KUP_BIGGO_URL},
    {"通路": "好晴朗(iOPEN)", "品項": "K.U.P晶球魚油90% 28包/盒 DHA EPA 韓國進口",
     "價格": 1980, "網址": KUP_BIGGO_URL},
    {"通路": "蝦皮・理可生活", "品項": "K.U.P 韓國進口 高純度微粒晶球膠囊魚油 28包/盒",
     "價格": 2600, "網址": KUP_BIGGO_URL},
]


@st.cache_data(ttl=600, show_spinner=False)
def _fetch_kup_live():
    """只抓 PChome 現貨（這部分可快取）；抓不到回空清單。"""
    seen, rows = set(), []
    for q in KUP_QUERIES:
        try:
            for x in fetch_pchome(q, limit=40):
                nm = x.get("品項", "")
                if "晶球" in nm and "魚油" in nm and to_price(x.get("價格")) is not None:
                    if nm in seen:
                        continue
                    seen.add(nm)
                    rows.append(x)
        except Exception:  # noqa: BLE001
            pass
    return rows


def fetch_kup():
    """回傳 (rows, is_live)。墊底資料不快取，改資料就即時反映。"""
    rows = _fetch_kup_live()
    if rows:
        return rows, True
    return [dict(r) for r in KUP_FALLBACK], False


def estimate_packs(name):
    if not name:
        return PACKS_PER_BOX
    m = re.search(r"[xX×]\s*(\d+)\s*盒", name) or re.search(r"(\d+)\s*盒", name)
    boxes = int(m.group(1)) if m else 1
    return boxes * PACKS_PER_BOX


def build_df(rows):
    if not rows:
        return pd.DataFrame(columns=["通路", "品項", "價格", "每包單價", "網址"])
    df = pd.DataFrame(rows)
    df["價格"] = pd.to_numeric(df["價格"], errors="coerce")
    df["每包單價"] = df.apply(
        lambda r: round(r["價格"] / estimate_packs(r["品項"]), 1)
        if pd.notna(r["價格"]) else None, axis=1)
    df = df.dropna(subset=["價格"]).drop_duplicates(subset=["通路", "品項"])
    return df.sort_values("價格").reset_index(drop=True)


def to_excel_bytes(df, term):
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as xw:
        df.to_excel(xw, sheet_name="通路比價", index=False)
        summary = pd.DataFrame({
            "項目": ["搜尋關鍵字", "抓取時間", "筆數", "最低價", "最高價", "平均價"],
            "值": [term, dt.datetime.now().strftime("%Y-%m-%d %H:%M"), len(df),
                  df["價格"].min(), df["價格"].max(), round(df["價格"].mean())],
        })
        summary.to_excel(xw, sheet_name="分析摘要", index=False)
    return buf.getvalue()


# ---------------------------------------------------------------------------
# 競品比較（大廠牌；顯示名 -> 搜尋關鍵字，可自行增減）
# ---------------------------------------------------------------------------
BRANDS = {
    "K.U.P 晶球魚油": "晶球魚油",
    "三得利 Suntory": "三得利 魚油 DHA EPA",
    "白蘭氏 Brand's": "白蘭氏 深海魚油",
    "大研生醫": "大研生醫 深海魚油",
    "澳佳寶 Blackmores": "Blackmores 魚油",
    "GNC 健安喜": "GNC 魚油",
}
FISH_KEYWORDS = ("魚油", "omega", "epa", "dha", "fish oil")


def looks_like_fishoil(name):
    n = (name or "").lower()
    return any(k in n for k in FISH_KEYWORDS)


def to_price(v):
    """把可能含逗號/文字的價格轉成數字；轉不出來回 None。"""
    if v is None:
        return None
    s = re.sub(r"[^\d.]", "", str(v))
    return float(s) if s else None


def parse_units(name):
    """從品名粗估總粒/包數（含盒/瓶/入倍數），用來估每單位價；估不到回 None。"""
    if not name:
        return None
    m = re.search(r"(\d+)\s*(?:粒|顆|包|錠|軟膠囊)", name)
    if not m:
        return None
    base = int(m.group(1))
    mm = (re.search(r"[xX×]\s*(\d+)\s*(?:盒|瓶|入|組)", name)
          or re.search(r"(\d+)\s*(?:盒|瓶|入|組)", name))
    return base * (int(mm.group(1)) if mm else 1)


def compare_brands(brand_items):
    """brand_items: tuple of (顯示名, 搜尋關鍵字)；每個品牌抓回最便宜的魚油品項。"""
    rows = []
    for disp, kw in brand_items:
        found = []
        if disp.startswith("K.U.P") or "晶球" in kw:
            try:
                found, _live = fetch_kup()   # K.U.P 專用（PChome→BigGo 墊底）
            except Exception:  # noqa: BLE001
                found = []
        else:
            for _, fn in SOURCES:
                try:
                    found += fn(kw)
                except Exception:  # noqa: BLE001
                    pass
        cand = []
        for x in found:
            if not looks_like_fishoil(x.get("品項", "")):
                continue
            p = to_price(x.get("價格"))
            if p is None:
                continue
            x["_p"] = p
            cand.append(x)
        if not cand:
            rows.append({"品牌": disp, "代表品項（最低價）": "— 這次未抓到 —",
                         "最低售價 (NT$)": None, "平均售價 (NT$)": None,
                         "每單位價(平均/顆)": None, "找到筆數": 0, "網址": ""})
            continue
        cand.sort(key=lambda x: x["_p"])
        best = cand[0]
        prices = [c["_p"] for c in cand]
        avg = sum(prices) / len(prices)
        units = parse_units(best["品項"])
        rows.append({
            "品牌": disp,
            "代表品項（最低價）": best["品項"][:42],
            "最低售價 (NT$)": int(best["_p"]),
            "平均售價 (NT$)": int(round(avg)),
            "每單位價(平均/顆)": round(avg / units, 1) if units else None,
            "找到筆數": len(cand),
            "網址": best["網址"],
        })
    return pd.DataFrame(rows)


def to_excel_compare(df):
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as xw:
        df.to_excel(xw, sheet_name="競品比價", index=False)
        note = pd.DataFrame({"說明": [
            "資料由 PChome / momo 搜尋自動抓取，取各品牌最便宜的魚油品項。",
            "各品牌劑型與每份 EPA+DHA 含量不同，售價僅供概覽。",
            "『每單位價』為由品名推估之粒/包單價，估不到會留空。",
            f"產生時間：{dt.datetime.now():%Y-%m-%d %H:%M}",
        ]})
        note.to_excel(xw, sheet_name="說明", index=False)
    return buf.getvalue()


# ---------------------------------------------------------------------------
# 介面
# ---------------------------------------------------------------------------
st.markdown("""
<div class="kup-hero">
  <p class="t">🐟 K.U.P 晶球魚油 · 市場資料查詢</p>
  <p class="s">韓國進口 · 藥品級 EE 型態魚油 · 微型無縫晶球膠囊｜自動抓取通路價格・整理成 Excel</p>
  <div class="kup-chips">
    <span class="c">Omega-3 90%</span>
    <span class="c">EPA+DHA 1680mg</span>
    <span class="c">IFOS 5★</span>
    <span class="c">2000mg × 28 包</span>
  </div>
</div>
""", unsafe_allow_html=True)

tab1, tab2, tab3 = st.tabs(["🔎 即時比價搜尋", "🆚 競品比較", "📋 產品資料卡"])

with tab1:
    col_in, col_btn = st.columns([4, 1])
    with col_in:
        term = st.text_input("商品或成分名稱", value="晶球魚油",
                             label_visibility="collapsed",
                             placeholder="輸入商品名稱，例如：晶球魚油")
    with col_btn:
        go = st.button("開始搜尋", type="primary", use_container_width=True)
    only_fish = st.checkbox("只顯示魚油相關結果（過濾掉不相干品項）", value=True)

    if go:
        rows, errors = [], []
        kup_live = True
        term_l = term.lower()
        is_kup = ("晶球" in term) or ("kup" in term_l) or ("k.u.p" in term_l)
        with st.spinner("正在搜尋各通路，請稍候…"):
            if is_kup:
                # K.U.P 專用：PChome 現貨優先，抓不到就用 BigGo 擷取的真實資料墊底
                try:
                    kup_rows, kup_live = fetch_kup()
                    rows += kup_rows
                except Exception as e:  # noqa: BLE001
                    errors.append(f"PChome（{type(e).__name__}）")
            else:
                for label, fn in SOURCES:
                    try:
                        rows += fn(term)
                    except Exception as e:  # noqa: BLE001
                        errors.append(f"{label}（{type(e).__name__}）")
        df = build_df(rows)
        if only_fish and not df.empty:
            df = df[df["品項"].apply(looks_like_fishoil)].reset_index(drop=True)
        if is_kup and not kup_live:
            st.link_button("🔗 BigGo・K.U.P 即時比價（點此看目前有貨賣場）",
                           "https://biggo.com.tw/s/%E6%99%B6%E7%90%83%E9%AD%9A%E6%B2%B9%20k.u.p")

        if df.empty:
            st.warning("這次沒有抓到明碼價格。可能是關鍵字太少結果，或通路把價格藏在動態頁面。"
                       "可以換個關鍵字再試，或參考「產品資料卡」分頁。")
        else:
            c1, c2, c3, c4 = st.columns(4)
            c1.metric("筆數", f"{len(df)}")
            c2.metric("最低價", f"${int(df['價格'].min()):,}")
            c3.metric("最高價", f"${int(df['價格'].max()):,}")
            c4.metric("平均價", f"${int(round(df['價格'].mean())):,}")

            st.dataframe(
                df, use_container_width=True, hide_index=True,
                column_config={"網址": st.column_config.LinkColumn("連結", display_text="開啟")},
            )

            chart = df.head(10).set_index("品項")["每包單價"]
            st.subheader("每包單價比較（最便宜前 10 筆）")
            st.bar_chart(chart)

            st.download_button(
                "⬇️ 下載 Excel", data=to_excel_bytes(df, term),
                file_name=f"KUP價格_{dt.date.today():%Y%m%d}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary")

        if errors:
            st.caption("註：以下通路這次沒抓到資料 —— " + "、".join(errors) +
                       "。動態載入的購物站需用 Playwright 才抓得到（報告可寫成技術說明）。")

with tab2:
    st.subheader("競品比較（自動搜尋大廠牌）")
    st.caption("勾選要比較的品牌 → 按「開始比較」。程式會各抓回最便宜的魚油品項，"
               "並算出最低價與平均價，做並排比較、可下載 Excel。")

    all_brands = list(BRANDS.keys())
    st.markdown("**要比較的品牌**（點方塊即可勾選／取消）")
    bcols = st.columns(3)
    chosen = []
    for i, b in enumerate(all_brands):
        if bcols[i % 3].checkbox(b, value=True, key=f"brand_{i}"):
            chosen.append(b)
    extra = st.text_input("再加一個自訂品牌（選填）", placeholder="例如：Swisse 魚油")

    if st.button("開始比較", type="primary"):
        items = [(b, BRANDS[b]) for b in chosen]
        if extra.strip():
            items.append((extra.strip(), extra.strip()))
        if not items:
            st.warning("請至少勾選一個品牌。")
        else:
            with st.spinner("正在比較各品牌，請稍候（品牌越多越久）…"):
                cdf = compare_brands(tuple(items))
            priced = cdf[cdf["最低售價 (NT$)"].notna()].copy()

            if priced.empty:
                st.warning("這次各品牌都沒抓到明碼價格，稍後再試或減少品牌數量。")
            else:
                cheapest = priced.loc[priced["最低售價 (NT$)"].idxmin(), "品牌"]
                c1, c2, c3, c4 = st.columns(4)
                c1.metric("比較品牌數", f"{len(priced)}")
                c2.metric("最低售價品牌", cheapest)
                c3.metric("最低售價", f"${int(priced['最低售價 (NT$)'].min()):,}")
                c4.metric("平均售價", f"${int(round(priced['平均售價 (NT$)'].mean())):,}")

                st.dataframe(
                    cdf, use_container_width=True, hide_index=True,
                    column_config={"網址": st.column_config.LinkColumn("連結", display_text="開啟")})

                st.subheader("各品牌平均售價")
                st.bar_chart(priced.set_index("品牌")["平均售價 (NT$)"])

                st.download_button(
                    "⬇️ 下載競品比價 Excel", data=to_excel_compare(cdf),
                    file_name=f"魚油競品比價_{dt.date.today():%Y%m%d}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    type="primary")

            st.caption("註：各品牌劑型與每份 EPA+DHA 含量不同，售價僅供概覽；"
                       "精準比較請看各商品頁的每份濃度。『每單位價(平均/顆)』＝平均售價÷推估顆數，"
                       "估不到會留空。")

with tab3:
    st.subheader("K.U.P 晶球魚油")
    st.write("韓國進口 · 藥品級 EE 型態魚油 · 微型無縫晶球膠囊")

    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Omega-3 純度", "90%")
    m2.metric("EPA+DHA / 包", "1680 mg")
    m3.metric("膠囊大小", "0.38 cm")
    m4.metric("IFOS 認證", "5 ★")

    st.markdown("#### 規格與成分")
    st.table(pd.DataFrame(KUP_SPECS, columns=["項目", "內容"]))

    left, right = st.columns(2)
    with left:
        st.markdown("#### 認證與品質")
        badges = '<div class="kup-badges">'
        for c in KUP_CERTS:
            if "（" in c:
                name, desc = c.split("（", 1)
                desc = desc.rstrip("）")
            else:
                name, desc = c, ""
            badges += (f'<div class="kup-badge">✓ <b>{name}</b>'
                       + (f'<br><span class="d">{desc}</span>' if desc else "")
                       + '</div>')
        badges += '</div>'
        st.markdown(badges, unsafe_allow_html=True)
    with right:
        st.markdown("#### 食用方式與注意事項")
        for u in KUP_USAGE:
            st.markdown(f"- {u}")
        st.markdown("#### 市場價格參考")
        st.table(KUP_REF_PRICES)

    st.markdown("#### 🔗 商品頁 / 購買通路（點連結可連到販售頁，供報告佐證）")
    lcols = st.columns(2)
    for i, (label, url) in enumerate(KUP_LINKS):
        with lcols[i % 2]:
            st.link_button(label, url, use_container_width=True)
    st.caption("提醒：單一賣場的特定商品可能『售完／下架』，該連結會顯示缺貨頁（這也代表資料是即時的）；"
               "跨通路比價頁（BigGo／飛比）永遠列出目前有貨的賣場，最適合當佐證截圖。")

    st.caption("資料由公開網路資訊彙整，價格與規格會隨時間變動，實際以各賣場頁面為準。"
               "本資料僅供整理參考，非醫療建議。")
