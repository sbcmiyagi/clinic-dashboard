import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import calendar
from datetime import date
from pathlib import Path

# Googleスプレッドシート（公開URL）またはローカルファイルを自動判定
SHEET_ID = "1nYlBXMdibPZf08RmSCUCKmOHrClTl8VEHBEeNTNvgd4"
GSHEET_URL = f"https://docs.google.com/spreadsheets/d/{SHEET_ID}/export?format=xlsx"
LOCAL_PATH = Path.home() / "Documents" / "クリニックDB" / "院情報一覧_カウント自動化.xlsx"
EXCEL_PATH = GSHEET_URL  # Streamlit Cloud ではGoogleスプレッドシートから読み込む
ORANGE_TWIST_COUNT = 24

# ── 色定義（元Excelと同じ） ──────────────────────────
C_HEADER_TITLE = "#D9EAD3"   # 薄緑：セクション見出し
C_HEADER_COL   = "#2C3E50"   # 濃紺：列ヘッダー
C_TOTAL        = "#FFF2CC"   # 薄黄：合計行
C_ORANGE       = "#E67E22"   # 橙：OrangeTwist・IR広報
C_BLUE         = "#2980B9"   # 青：既存グループ合計・全拠点合計
C_WHITE        = "#FFFFFF"
C_TEXT_LIGHT   = "#FFFFFF"

TARGET_BRANDS = [
    "湘南美容クリニック", "湘南美容クリニック（ﾍﾞﾄﾅﾑ）", "湘南歯科クリニック", "湘南AGAクリニック", "湘南美容皮フ科",
    "湘南皮膚科クリニック", "SBC NEO Skin Clinic", "肌の青空クリニック",
    "湘南内科皮フ科クリニック", "湘南美容皮フ科内科クリニック", "イテウォン",
    "湘南メディカル記念病院", "新宿近視クリニック", "西新宿整形外科",
    "SBC横浜駅前整形外科", "六本木レディース", "神奈川レディース", "リッツ美容外科",
    "リゼクリニック", "ゴリラクリニック", "JUNCLINIC", "SBC東京医療大学付属クリニック",
    "SBC東京接骨院", "SBC BODY ARCHI",
    "SkinGo!", "Wen & Weng Family Clinic", "The Chelsea Clinic",
    "Chelsea Aesthetics", "Chelsea Asthetics", "Gangnam Laser Clinic", "Rochor Centre Clinic"
]
EXCLUDE_BRANDS_FOR_PR = ["SBC東京医療大学付属クリニック", "SBC東京接骨院", "SBC BODY ARCHI"]
OVERSEAS_KEYWORDS = ["DS", "DSS", "RCC", "WWFC", "WWMG", "SkinGo", "Chelsea", "ﾍﾞﾄﾅﾑ", "Vietnam", "Shoubikai"]
TARGET_HOUJIN_ORDER = [
    "医療法人 湘美会", "医療法人社団 孝和会", "医療法人社団 樹慶会", "医療法人社団 愛恵会",
    "医療法人社団 風林会", "医療法人社団 菜寿会", "医療法人社団 孝仁会",
    "一般社団法人美央斗会", "一般社団法人MASA", "健美会",
    "医療法人社団 リッツ美容外科", "株式会社 湘南美容クリニック",
    "DS", "DSS", "DSS(FC)", "RCC", "WWFC", "WWMG",
    "学校法人 SBC東京医療大学", "医療法人 きびたき会", "SBC東京接骨院", "株式会社 MG"
]
EXCLUDE_HOUJIN_FROM_COUNT = [
    "学校法人 SBC東京医療大学", "㈻SBC東京医療大学附属", "医療法人 きびたき会",
    "医療法人社団 百花会", "SBC東京接骨院", "株式会社 MG"
]


st.set_page_config(page_title="クリニック数ダッシュボード", layout="wide")

st.markdown("""
<style>
/* テーブル全体の文字サイズを1pt大きく（デフォルト14px → 15px相当） */
div[data-testid="stDataFrame"] iframe {
    font-size: 15px !important;
}
/* dataframeセルのフォントサイズ */
.dvn-scroller, .cell-wrap-text, [data-testid="glideDataEditor"] {
    font-size: 15px !important;
}
</style>
""", unsafe_allow_html=True)

def to_ts(v):
    if isinstance(v, pd.Timestamp): return v
    if pd.isna(v): return None
    try: return pd.Timestamp(v)
    except: return None

@st.cache_data(ttl=3600)
def load_master():
    df = pd.read_excel(EXCEL_PATH, sheet_name="クリニック一覧")
    for c in ["開院日", "MA日", "移転拡張日", "業態転換日", "閉院日"]:
        if c in df.columns:
            df[c] = pd.to_datetime(df[c], errors="coerce")
    excl = "猶予期限（業態転換後の院が開院するまでの期間）"
    def parse_limit(v):
        if isinstance(v, pd.Timestamp) and not pd.isna(v): return v.strftime("%Y/%m")
        if pd.isna(v): return ""
        return str(v).strip()
    if excl in df.columns:
        df[excl] = df[excl].apply(parse_limit)
    return df


def check_active(row, target):
    ma = to_ts(row.get("MA日")); op = to_ts(row.get("開院日"))
    base = ma if ma is not None else op
    if base is None: return False
    d_conv = to_ts(row.get("業態転換日")); d_close = to_ts(row.get("閉院日"))
    # 業態転換日がある場合はそちらを終了日とする（閉院日は参考情報のみ）
    final_end = d_conv if d_conv is not None else d_close
    base_only = pd.Timestamp(base.year, base.month, base.day)
    if base_only <= target and (final_end is None or final_end > target): return True
    return False

def get_region(row):
    region_val = row.get("海外／国内", "")
    r = "" if pd.isna(region_val) else str(region_val).strip()
    if not r:
        b = str(row.get("ブランド","") or ""); h = str(row.get("法人名","") or "")
        return "海外" if any(kw in b or kw in h for kw in OVERSEAS_KEYWORDS) else "国内"
    return r

def get_brand(row):
    b = str(row.get("ブランド","") or "").strip()
    return "Chelsea Aesthetics" if b == "Chelsea Asthetics" else b

def month_end_ts(y, m): return pd.Timestamp(y, m, calendar.monthrange(y,m)[1], 23,59,59)
def month_start_ts(y, m): return pd.Timestamp(y, m, 1, 0, 0, 0)

def aggregate(df, me, ms):
    r1, r2, r3 = {}, {}, {}
    for _, row in df.iterrows():
        brand = get_brand(row); gyoutai = str(row.get("業態","") or "").strip()
        houjin = str(row.get("法人名","") or "個人/その他").strip()
        region = get_region(row)
        if not brand or not gyoutai: continue
        if brand not in TARGET_BRANDS: continue
        is_pr = brand not in EXCLUDE_BRANDS_FOR_PR
        if check_active(row, me):
            k1 = brand+"|"+gyoutai
            r1.setdefault(k1, {"all":0,"pr":0}); r1[k1]["all"]+=1
            if is_pr: r1[k1]["pr"]+=1
            k2 = region+"|"+houjin
            r2.setdefault(k2, {"all":0,"pr":0}); r2[k2]["all"]+=1
            if is_pr: r2[k2]["pr"]+=1
        if check_active(row, ms):
            if houjin in TARGET_HOUJIN_ORDER:
                r3.setdefault(houjin, {"all":0})
                if houjin not in EXCLUDE_HOUJIN_FROM_COUNT: r3[houjin]["all"]+=1
    return r1, r2, r3

def style_header(df, title_cols):
    """DataFrameにHTML形式でヘッダー色を適用して返す"""
    return df

def make_brand_df(r1):
    rows = []
    for brand in TARGET_BRANDS:
        keys = sorted([k for k in r1 if k.startswith(brand+"|")])
        for k in keys:
            p = k.split("|",1)
            rows.append({"ブランド":p[0],"業態":p[1],"広報・IR用":r1[k]["pr"],"全拠点(ALL)":r1[k]["all"]})
    df = pd.DataFrame(rows)
    sum_pr = df["広報・IR用"].sum() if not df.empty else 0
    sum_all = df["全拠点(ALL)"].sum() if not df.empty else 0
    totals = [
        {"ブランド":"集計内訳合計（店舗数のみ）","業態":"","広報・IR用":sum_pr,"全拠点(ALL)":sum_all},
        {"ブランド":"海外(OrangeTwist)加算","業態":"","広報・IR用":ORANGE_TWIST_COUNT,"全拠点(ALL)":"－"},
        {"ブランド":"最終報告数値","業態":"","広報・IR用":sum_pr+ORANGE_TWIST_COUNT,"全拠点(ALL)":sum_all},
    ]
    return pd.concat([df, pd.DataFrame(totals)], ignore_index=True), sum_pr, sum_all

def make_region_dfs(r2):
    dom, ovs = [], []
    for region in ["国内","海外"]:
        for k in sorted([k for k in r2 if k.startswith(region+"|")]):
            p=k.split("|",1); e={"法人名":p[1],"広報・IR用":r2[k]["pr"],"全拠点(ALL)":r2[k]["all"]}
            (dom if region=="国内" else ovs).append(e)
    return pd.DataFrame(dom), pd.DataFrame(ovs)

def make_houjin_df(r3):
    rows=[]; total=0
    for h in TARGET_HOUJIN_ORDER:
        cnt=r3.get(h,{}).get("all",0); rows.append({"法人名":h,"全拠点(ALL)":cnt}); total+=cnt
    rows.append({"法人名":"最終報告数値(合計)","全拠点(ALL)":total})
    return pd.DataFrame(rows)

def apply_style(df, total_rows=None, orange_rows=None, blue_rows=None):
    """行ごとに色を適用するスタイラーを返す"""
    total_rows = total_rows or []
    orange_rows = orange_rows or []
    blue_rows = blue_rows or []
    def row_style(row):
        idx = row.name
        if idx in orange_rows:
            return [f"background-color:{C_ORANGE};color:white;font-weight:bold"] * len(row)
        if idx in blue_rows:
            return [f"background-color:{C_BLUE};color:white;font-weight:bold"] * len(row)
        if idx in total_rows:
            return [f"background-color:{C_TOTAL};font-weight:bold"] * len(row)
        return [""] * len(row)
    return df.style.apply(row_style, axis=1)

def col_label(brand, gyoutai):
    return f"{brand}({gyoutai})" if gyoutai else brand

@st.cache_data(ttl=3600)
def load_brand_settings():
    """Excelの「ブランド設定」シートからブランド順・設定を読み込む。"""
    try:
        # まず全データを読み込み、ヘッダー行を自動検出する
        raw = pd.read_excel(EXCEL_PATH, sheet_name="ブランド設定", header=None)

        # 「ブランド名」という文字が入っている行をヘッダーとして使う
        header_row = None
        for i, row in raw.iterrows():
            if any("ブランド名" in str(v) for v in row.values):
                header_row = i
                break

        if header_row is None:
            return [], [], [], "「ブランド設定」シートにヘッダー行（ブランド名）が見つかりません。"

        df = pd.read_excel(EXCEL_PATH, sheet_name="ブランド設定", header=header_row)
        df = df.dropna(subset=["ブランド名"])
        df = df[df["ブランド名"].astype(str).str.strip() != ""]

        # 表示順列があれば並び替え
        if "表示順" in df.columns:
            df = df.sort_values("表示順")
        df = df.reset_index(drop=True)

        cols, existing_flags, exclude_pr_brands = [], [], []
        for _, row in df.iterrows():
            brand    = str(row.get("ブランド名", "") or "").strip()
            gyoutai_val = row.get("業態", "")
            gyoutai  = str(gyoutai_val).strip() if pd.notna(gyoutai_val) and str(gyoutai_val).strip() not in ("", "nan") else None
            existing_val = row.get("既存グループ合計", "")
            existing = str(existing_val).strip() if pd.notna(existing_val) else ""
            exclude_val = row.get("広報IR除外", "")
            exclude  = str(exclude_val).strip() if pd.notna(exclude_val) else ""
            if brand:
                cols.append((brand, gyoutai))
                existing_flags.append(existing == "○")
                if exclude == "○":
                    exclude_pr_brands.append(brand)

        return cols, existing_flags, exclude_pr_brands, None  # None = エラーなし

    except Exception as e:
        return [], [], [], str(e)

@st.cache_data(ttl=3600)
def build_jikei_table(_df_master):
    """クリニック一覧 ＋ ブランド設定シートから時系列推移を自動計算する"""
    # ブランド設定シートを読み込む
    brand_cols, existing_flags, _, _err = load_brand_settings()
    # 設定シートが空の場合はデータなしで返す
    if not brand_cols:
        return pd.DataFrame(columns=["年月"])

    today = date.today()
    y, m = 2021, 1
    records = []
    while (y, m) <= (today.year, today.month):
        me = month_end_ts(y, m)
        ms = month_start_ts(y, m)
        r1, _, _ = aggregate(_df_master, me, ms)

        row = {"年月": f"{y}/{m:02d}"}
        existing_total = 0
        all_total = 0
        pr_total = 0

        for i, (brand, gyoutai) in enumerate(brand_cols):
            if gyoutai:
                keys = [k for k in r1 if k == f"{brand}|{gyoutai}"]
            else:
                keys = [k for k in r1 if k.startswith(f"{brand}|")]
            cnt_all = sum(r1[k]["all"] for k in keys)
            cnt_pr  = sum(r1[k]["pr"]  for k in keys)
            label = col_label(brand, gyoutai)
            row[label] = cnt_all
            all_total += cnt_all
            pr_total  += cnt_pr
            if existing_flags[i]:
                existing_total += cnt_all

        row["既存グループ合計"] = existing_total
        row["全拠点合計(ALL)"]  = all_total
        row["OrangeTwist"]      = ORANGE_TWIST_COUNT
        row["IR・広報用"]       = pr_total + ORANGE_TWIST_COUNT
        records.append(row)
        m += 1
        if m > 12:
            m = 1; y += 1
    return pd.DataFrame(records)

# 時系列表の列ごとの色分け定義
JIKEI_COL_COLORS = {
    "既存グループ合計": ("既存グループ合計", C_BLUE,   "white"),
    "全拠点合計(ALL)":  ("全拠点合計(ALL)", C_BLUE,   "white"),
    "OrangeTwist":      ("OrangeTwist",     C_ORANGE, "white"),
    "IR・広報用":       ("IR・広報用",      C_ORANGE, "white"),
}

# ── ページ設定 ──────────────────────────────────────
df_master = load_master()
today = date.today()

# ブランド設定シートから広報IR除外リストとTARGET_BRANDSを動的に取得
_brand_cols_cfg, _, _exclude_pr_from_cfg, _settings_err = load_brand_settings()
if _settings_err:
    st.sidebar.warning(f"⚠️ ブランド設定の読み込みエラー:\n{_settings_err}")
if _exclude_pr_from_cfg:
    EXCLUDE_BRANDS_FOR_PR = _exclude_pr_from_cfg
if _brand_cols_cfg:
    TARGET_BRANDS = list(dict.fromkeys([b for b, _ in _brand_cols_cfg]))

st.sidebar.header("表示設定")
year_options = list(range(2021, today.year+1))
sel_year = st.sidebar.selectbox("年", year_options, index=len(year_options)-1)
sel_month = st.sidebar.selectbox("月", list(range(1,13)), index=today.month-1)
if st.sidebar.button("データを再読込"):
    st.cache_data.clear(); st.rerun()
st.sidebar.caption(f"データ元: 院情報一覧_カウント自動化.xlsx")

me = month_end_ts(sel_year, sel_month)
ms = month_start_ts(sel_year, sel_month)
with st.spinner("集計中..."): r1, r2, r3 = aggregate(df_master, me, ms)

brand_df, sum_pr, sum_all = make_brand_df(r1)
n_total = len(brand_df)

st.title("クリニック数ダッシュボード")
st.subheader(f"{sel_year}年{sel_month}月末 時点")
c1,c2,c3 = st.columns(3)
c1.metric("IR・広報用（OrangeTwist含む）", f"{sum_pr+ORANGE_TWIST_COUNT:,} 院")
c2.metric("全拠点(ALL)", f"{sum_all:,} 院")
c3.metric("OrangeTwist（別管理）", f"{ORANGE_TWIST_COUNT} 院")
st.markdown("---")

tab1, tab2, tab3, tab4, tab5 = st.tabs(["📋 月次内訳表", "📈 時系列推移", "📊 グラフ", "🏥 開院・閉院履歴", "🔵 業態転換履歴"])

# ══════════════════════════════════════════════
# Tab1: 月次内訳表_詳細（3表）
# ══════════════════════════════════════════════
with tab1:
    st.markdown(f"#### {sel_year}/{sel_month:02d}末 月次内訳表")
    col_t1, col_sp, col_t2, col_sp2, col_t3 = st.columns([4, 0.2, 4, 0.2, 2.5])

    # ── 表1: ブランド×業態 ──
    with col_t1:
        st.markdown(
            f'<div style="background:{C_HEADER_TITLE};padding:6px;text-align:center;font-weight:bold;border-radius:4px 4px 0 0">'
            f'{sel_year}/{sel_month:02d}末 ブランド×業態</div>',
            unsafe_allow_html=True
        )
        total_idx = list(range(n_total-3, n_total))       # 集計3行
        orange_idx = [n_total-2, n_total-1]               # OrangeTwist・最終
        styled = apply_style(brand_df, total_rows=[n_total-3], orange_rows=[n_total-1], blue_rows=[])
        # OrangeTwist行だけ橙
        def brand_row_style(row):
            i = row.name
            if i == n_total-1:   return [f"background-color:{C_ORANGE};color:white;font-weight:bold"]*len(row)
            if i == n_total-2:   return [f"background-color:{C_ORANGE};color:white"]*len(row)
            if i == n_total-3:   return [f"background-color:{C_TOTAL};font-weight:bold"]*len(row)
            return [""]*len(row)
        st.dataframe(brand_df.style.apply(brand_row_style, axis=1), hide_index=True, height="content")

    # ── 表2: 国内/海外×法人 ──
    with col_t2:
        dom_df, ovs_df = make_region_dfs(r2)
        st.markdown(
            f'<div style="background:{C_HEADER_TITLE};padding:6px;text-align:center;font-weight:bold;border-radius:4px 4px 0 0">'
            f'{sel_year}/{sel_month:02d}末 海外/国内×法人</div>',
            unsafe_allow_html=True
        )
        # 国内・海外を結合して1表に
        dom_sum_pr = dom_df["広報・IR用"].sum() if not dom_df.empty else 0
        dom_sum_all = dom_df["全拠点(ALL)"].sum() if not dom_df.empty else 0
        ovs_sum_pr = ovs_df["広報・IR用"].sum() if not ovs_df.empty else 0
        ovs_sum_all = ovs_df["全拠点(ALL)"].sum() if not ovs_df.empty else 0
        total_pr2 = dom_sum_pr + ovs_sum_pr
        total_all2 = dom_sum_all + ovs_sum_all

        # 国内・海外それぞれの小計
        dom_sub_pr  = dom_df["広報・IR用"].sum() if not dom_df.empty else 0
        dom_sub_all = dom_df["全拠点(ALL)"].sum() if not dom_df.empty else 0
        ovs_sub_pr  = ovs_df["広報・IR用"].sum() if not ovs_df.empty else 0
        ovs_sub_all = ovs_df["全拠点(ALL)"].sum() if not ovs_df.empty else 0

        sep_dom  = pd.DataFrame([{"法人名":"── 国内 ──","広報・IR用":"","全拠点(ALL)":""}])
        sub_dom  = pd.DataFrame([{"法人名":f"　国内 小計","広報・IR用":dom_sub_pr,"全拠点(ALL)":dom_sub_all}])
        sep_ovs  = pd.DataFrame([{"法人名":"── 海外 ──","広報・IR用":"","全拠点(ALL)":""}])
        sub_ovs  = pd.DataFrame([{"法人名":f"　海外 小計","広報・IR用":ovs_sub_pr,"全拠点(ALL)":ovs_sub_all}])
        tot_row  = pd.DataFrame([
            {"法人名":"集計内訳合計（店舗数のみ）","広報・IR用":total_pr2,"全拠点(ALL)":total_all2},
            {"法人名":"海外(OrangeTwist)加算","広報・IR用":ORANGE_TWIST_COUNT,"全拠点(ALL)":"－"},
            {"法人名":"最終報告数値","広報・IR用":total_pr2+ORANGE_TWIST_COUNT,"全拠点(ALL)":total_all2},
        ])
        region_df = pd.concat(
            [sep_dom, dom_df, sub_dom, sep_ovs, ovs_df, sub_ovs, tot_row],
            ignore_index=True
        )
        n2 = len(region_df)
        # 各行のインデックスを特定
        idx_sep_dom = 0
        idx_sub_dom = 1 + len(dom_df)
        idx_sep_ovs = idx_sub_dom + 1
        idx_sub_ovs = idx_sep_ovs + 1 + len(ovs_df)
        def region_row_style(row):
            i = row.name
            if i in [idx_sep_dom, idx_sep_ovs]:
                return [f"background-color:{C_HEADER_COL};color:white;font-weight:bold"]*len(row)
            if i in [idx_sub_dom, idx_sub_ovs]:
                return [f"background-color:#D6EAF8;font-weight:bold;color:#1A5276"]*len(row)
            if i == n2-1:
                return [f"background-color:{C_ORANGE};color:white;font-weight:bold"]*len(row)
            if i == n2-2:
                return [f"background-color:{C_ORANGE};color:white"]*len(row)
            if i == n2-3:
                return [f"background-color:{C_TOTAL};font-weight:bold"]*len(row)
            return [""]*len(row)
        st.dataframe(region_df.style.apply(region_row_style, axis=1), hide_index=True, height="content")

    # ── 表3: 法人別 ──
    with col_t3:
        houjin_df = make_houjin_df(r3)
        nh = len(houjin_df)
        st.markdown(
            f'<div style="background:{C_HEADER_TITLE};padding:6px;text-align:center;font-weight:bold;border-radius:4px 4px 0 0">'
            f'{sel_year}/{sel_month:02d}初 法人別内訳</div>',
            unsafe_allow_html=True
        )
        def houjin_row_style(row):
            if row.name == nh-1: return [f"background-color:{C_TOTAL};font-weight:bold"]*len(row)
            return [""]*len(row)
        st.dataframe(houjin_df.style.apply(houjin_row_style, axis=1), hide_index=True, height="content")

# ══════════════════════════════════════════════
# Tab2: 時系列推移（クリニック一覧から自動計算）
# ══════════════════════════════════════════════
with tab2:
    st.markdown("#### 時系列推移（月末時点）")

    with st.spinner("月次データを集計中..."):
        jikei_df = build_jikei_table(df_master)

    # 列ごとの色スタイル
    def jikei_col_style(col):
        if col.name in ["OrangeTwist", "IR・広報用"]:
            return ["background-color:#FEF0E3; color:#C0392B; font-weight:bold"] * len(col)
        if col.name in ["既存グループ合計", "全拠点合計(ALL)"]:
            return ["background-color:#EBF5FB; color:#1A5276; font-weight:bold"] * len(col)
        return [""] * len(col)

    st.caption("※ 「クリニック一覧」シートから自動計算しています")
    st.dataframe(
        jikei_df.style.apply(jikei_col_style, axis=0),
        hide_index=True,
        height=500,
        use_container_width=True,
    )

# ══════════════════════════════════════════════
# Tab3: グラフ
# ══════════════════════════════════════════════
with tab3:
    st.markdown(f"#### {sel_year}年{sel_month}月末 ブランド別院数")
    plot_df = brand_df[brand_df["ブランド"].isin(TARGET_BRANDS)].copy()
    if not plot_df.empty:
        fig = px.bar(plot_df, x="全拠点(ALL)", y="ブランド", orientation="h", height=600,
                     color_discrete_sequence=[C_BLUE])
        fig.update_layout(yaxis=dict(autorange="reversed"))
        st.plotly_chart(fig, use_container_width=True)

    st.markdown("#### ブランド別 月次推移")
    brand_cols = [c for c in jikei_df.columns if c not in
                  ["年月","既存グループ合計","全拠点合計(ALL)","OrangeTwist","IR・広報用"]]
    default_brands = [c for c in ["湘南美容クリニック(美容外科・皮膚科)","リゼクリニック","ゴリラクリニック"] if c in brand_cols]
    sel = st.multiselect("ブランドを選択", brand_cols,
                         default=default_brands if default_brands else brand_cols[:3])
    if sel:
        melt = jikei_df[["年月"]+sel].melt(id_vars="年月", var_name="ブランド", value_name="院数")
        fig2 = px.line(melt, x="年月", y="院数", color="ブランド", markers=True, height=400)
        fig2.update_layout(xaxis=dict(tickangle=45))
        st.plotly_chart(fig2, use_container_width=True)

# ══════════════════════════════════════════════
# Tab4: 開院・閉院履歴
# ══════════════════════════════════════════════
with tab4:
    st.markdown("#### 開院・閉院 年別・月別履歴")

    # 年の範囲をデータから自動決定
    all_dates = pd.concat([
        df_master["開院日"].dropna(),
        df_master["MA日"].dropna(),
        df_master["閉院日"].dropna(),
        df_master["業態転換日"].dropna(),
    ])
    data_max_year = int(all_dates.dt.year.max()) if not all_dates.empty else today.year
    hist_years = list(range(2025, max(data_max_year, today.year) + 1))

    def get_base_date(row):
        ma = to_ts(row.get("MA日")); op = to_ts(row.get("開院日"))
        return ma if ma is not None else op

    def build_history(year):
        monthly = {}
        for month in range(1, 13):
            ms = pd.Timestamp(year, month, 1)
            me = pd.Timestamp(year, month, calendar.monthrange(year, month)[1])
            opened_rows, closed_rows, convert_rows = [], [], []
            for _, row in df_master.iterrows():
                brand = get_brand(row)
                if not brand or brand not in TARGET_BRANDS:
                    continue
                base    = get_base_date(row)
                d_conv  = to_ts(row.get("業態転換日"))
                d_close = to_ts(row.get("閉院日"))

                # ── 開院（業態転換後の院には印をつける） ──
                tenkan_mae = str(row.get("転換前業態","") or "").strip()
                is_converted = tenkan_mae not in ("", "nan")
                if base is not None:
                    base_only = pd.Timestamp(base.year, base.month, base.day)
                    if ms <= base_only <= me:
                        opened_rows.append({
                            "種別": "🔵 業態転換" if is_converted else "🟢 新規開院",
                            "開院日": base_only.strftime("%Y/%m/%d"),
                            "院名": row.get("正式名称",""),
                            "ブランド": brand,
                            "業態": str(row.get("業態","") or ""),
                            "法人名": str(row.get("法人名","") or ""),
                            "国内／海外": get_region(row),
                        })

                # ── 業態転換（業態転換日がある院はすべてここに表示） ──
                if d_conv is not None:
                    conv_only = pd.Timestamp(d_conv.year, d_conv.month, d_conv.day)
                    if ms <= conv_only <= me:
                        # 転換前業態：I列が空なら業態列（N列）を自動で使用
                        before_gyoutai = str(row.get("転換前業態","") or "").strip()
                        if not before_gyoutai or before_gyoutai == "nan":
                            before_gyoutai = str(row.get("業態","") or "").strip()
                        after_name     = str(row.get("転換後院名","") or "").strip()
                        after_gyoutai  = str(row.get("転換後業態","") or "").strip()
                        convert_rows.append({
                            "転換前院名":   row.get("正式名称",""),
                            "転換前ブランド": brand,
                            "転換前業態":   before_gyoutai if before_gyoutai else "―",
                            "　":          "⇒",
                            "転換後院名":   after_name    if after_name    else "―",
                            "転換後業態":   after_gyoutai if after_gyoutai else "―",
                            "法人名":      str(row.get("法人名","") or ""),
                            "国内／海外":   get_region(row),
                            "業態転換日":   conv_only.strftime("%Y/%m/%d"),
                        })

                # ── 閉院（閉院日または業態転換日が当月内の院をすべて表示） ──
                # 表示に使う日付：閉院日があればそちら、なければ業態転換日
                close_ref = d_close if d_close is not None else d_conv
                if close_ref is not None:
                    close_only = pd.Timestamp(close_ref.year, close_ref.month, close_ref.day)
                    if ms <= close_only <= me:
                        closed_rows.append({
                            "種別": "🔵 業態転換" if d_conv is not None else "🔴 閉院",
                            "閉院日": close_only.strftime("%Y/%m/%d"),
                            "院名": row.get("正式名称",""),
                            "ブランド": brand,
                            "業態": str(row.get("業態","") or ""),
                            "法人名": str(row.get("法人名","") or ""),
                            "国内／海外": get_region(row),
                        })

            monthly[month] = {"opened": opened_rows, "closed": closed_rows, "convert": convert_rows}
        return monthly

    # 年ごとに表示
    for hist_year in hist_years:
        with st.spinner(f"{hist_year}年のデータを集計中..."):
            monthly_data = build_history(hist_year)

        year_total_open  = sum(len(v["opened"]) for v in monthly_data.values())
        year_total_close = sum(len(v["closed"]) for v in monthly_data.values())

        # 年ヘッダー
        st.markdown(
            f'<div style="background:{C_HEADER_COL};color:white;padding:10px 16px;'
            f'font-weight:bold;font-size:18px;border-radius:6px;margin-top:24px">'
            f'📅 {hist_year}年　｜　'
            f'🟢 開院 {year_total_open} 院　　'
            f'🔴 閉院 {year_total_close} 院'
            f'</div>',
            unsafe_allow_html=True
        )
        st.markdown("")

        # 月ごとのプルダウン（開院・閉院と業態転換を別々に表示）
        for month in range(1, 13):
            opened  = monthly_data[month]["opened"]
            closed  = monthly_data[month]["closed"]
            convert = monthly_data[month]["convert"]
            n_open    = len(opened)
            n_close   = len(closed)
            n_convert = len(convert)
            if n_open == 0 and n_close == 0 and n_convert == 0:
                continue

            # ── プルダウン①：開院・閉院 ──
            if n_open > 0 or n_close > 0:
                label1 = f"{hist_year}年{month}月　🟢 開院 {n_open}院　　🔴 閉院 {n_close}院"
                with st.expander(label1, expanded=False):
                    col_open, col_close = st.columns(2)
                    with col_open:
                        st.markdown(
                            f'<div style="background:#D5F5E3;padding:5px 10px;font-weight:bold;'
                            f'border-left:4px solid #27AE60;margin-bottom:6px">'
                            f'🟢 開院　{n_open}院</div>', unsafe_allow_html=True
                        )
                        if n_open > 0:
                            open_df = pd.DataFrame(opened)[["種別","開院日","院名","ブランド","業態","法人名","国内／海外"]]
                            st.dataframe(
                                open_df.style.apply(
                                    lambda r: ["background-color:#F0FFF4"]*len(r), axis=1),
                                hide_index=True, height="content", use_container_width=True
                            )
                        else:
                            st.caption("開院なし")
                    with col_close:
                        st.markdown(
                            f'<div style="background:#FADBD8;padding:5px 10px;font-weight:bold;'
                            f'border-left:4px solid #E74C3C;margin-bottom:6px">'
                            f'🔴 閉院　{n_close}院</div>', unsafe_allow_html=True
                        )
                        if n_close > 0:
                            close_df = pd.DataFrame(closed)[["種別","閉院日","院名","ブランド","業態","法人名","国内／海外"]]
                            st.dataframe(
                                close_df.style.apply(
                                    lambda r: ["background-color:#FFF5F5"]*len(r), axis=1),
                                hide_index=True, height="content", use_container_width=True
                            )
                        else:
                            st.caption("閉院なし")

# ══════════════════════════════════════════════
# Tab5: 業態転換履歴
# ══════════════════════════════════════════════
with tab5:
    st.markdown("#### 業態転換 年別・月別履歴")

    for hist_year in hist_years:
        with st.spinner(f"{hist_year}年のデータを集計中..."):
            monthly_data_conv = build_history(hist_year)

        year_total_convert = sum(len(v["convert"]) for v in monthly_data_conv.values())
        if year_total_convert == 0:
            continue

        # 年ヘッダー
        st.markdown(
            f'<div style="background:{C_BLUE};color:white;padding:10px 16px;'
            f'font-weight:bold;font-size:18px;border-radius:6px;margin-top:24px">'
            f'📅 {hist_year}年　｜　🔵 業態転換 {year_total_convert} 院'
            f'</div>',
            unsafe_allow_html=True
        )
        st.markdown("")

        for month in range(1, 13):
            convert = monthly_data_conv[month]["convert"]
            n_convert = len(convert)
            if n_convert == 0:
                continue

            label = f"{hist_year}年{month}月　🔵 業態転換 {n_convert}院"
            with st.expander(label, expanded=False):
                st.markdown(
                    f'<div style="background:#D6EAF8;padding:5px 10px;font-weight:bold;'
                    f'border-left:4px solid #2980B9;margin-bottom:6px">'
                    f'🔵 業態転換　{n_convert}院</div>', unsafe_allow_html=True
                )
                st.dataframe(
                    pd.DataFrame(convert).style.apply(
                        lambda r: ["background-color:#EBF5FB"]*len(r), axis=1),
                    hide_index=True, height="content", use_container_width=True
                )
