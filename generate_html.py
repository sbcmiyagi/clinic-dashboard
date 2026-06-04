"""
クリニック数ダッシュボード HTMLレポート生成スクリプト
実行すると index.html を生成します
"""
import json
import pandas as pd
import plotly.express as px
import calendar
from datetime import date
from pathlib import Path

# ── ファイルパス ──────────────────────────────────────
BOX_PATH  = Path(r"C:\Users\宮城杏奈\Box\総合企画部_特殊案件\その他\院情報一覧カウント\院情報一覧_カウント自動化.xlsx")
LOCAL_PATH = Path.home() / "Documents" / "クリニックDB" / "院情報一覧_カウント自動化.xlsx"
OUTPUT_HTML = Path(__file__).parent / "index.html"

ORANGE_TWIST_COUNT = 24
C_HEADER = "#2C3E50"
C_TOTAL  = "#FFF2CC"
C_ORANGE = "#E67E22"
C_BLUE   = "#2980B9"
C_GREEN  = "#D9EAD3"

TARGET_BRANDS = [
    "湘南美容クリニック","湘南美容クリニック（ﾍﾞﾄﾅﾑ）","湘南歯科クリニック","湘南AGAクリニック","湘南美容皮フ科",
    "湘南皮膚科クリニック","SBC NEO Skin Clinic","肌の青空クリニック","湘南内科皮フ科クリニック",
    "湘南美容皮フ科内科クリニック","イテウォン","湘南メディカル記念病院","新宿近視クリニック",
    "西新宿整形外科","SBC横浜駅前整形外科","六本木レディース","神奈川レディース","リッツ美容外科",
    "リゼクリニック","ゴリラクリニック","JUNCLINIC","SBC東京医療大学付属クリニック",
    "SBC東京接骨院","SBC BODY ARCHI","SkinGo!","Wen & Weng Family Clinic",
    "The Chelsea Clinic","Chelsea Aesthetics","Chelsea Asthetics","Gangnam Laser Clinic","Rochor Centre Clinic"
]
EXCLUDE_PR = ["SBC東京医療大学付属クリニック","SBC東京接骨院","SBC BODY ARCHI"]
OVERSEAS_KW = ["DS","DSS","RCC","WWFC","WWMG","SkinGo","Chelsea","ﾍﾞﾄﾅﾑ","Vietnam","Shoubikai"]
HOUJIN_ORDER = [
    "医療法人 湘美会","医療法人社団 孝和会","医療法人社団 樹慶会","医療法人社団 愛恵会",
    "医療法人社団 風林会","医療法人社団 菜寿会","医療法人社団 孝仁会",
    "一般社団法人美央斗会","一般社団法人MASA","健美会",
    "医療法人社団 リッツ美容外科","株式会社 湘南美容クリニック",
    "DS","DSS","DSS(FC)","RCC","WWFC","WWMG",
    "学校法人 SBC東京医療大学","医療法人 きびたき会","SBC東京接骨院","株式会社 MG"
]
EXCLUDE_HOUJIN = ["学校法人 SBC東京医療大学","㈻SBC東京医療大学附属","医療法人 きびたき会","医療法人社団 百花会","SBC東京接骨院","株式会社 MG"]

HOUJIN_GROUPS = [
    ("①6医療法人合計", ["医療法人 湘美会","医療法人社団 孝和会","医療法人社団 菜寿会","医療法人社団 愛恵会","医療法人社団 樹慶会","医療法人社団 リッツ美容外科","一般社団法人MASA","健美会","法人無し（個人開設）","個人/その他"]),
    ("②3医療法人合計", ["医療法人社団 風林会","医療法人 きびたき会","医療法人社団 百花会"]),
    ("③医療法人社団十二会", ["医療法人社団十二会"]),
    ("④2医療法人合計", ["医療法人社団美咲会","一般社団法人美央斗会"]),
    ("⑤㈻SBC東京医療大学附属", ["㈻SBC東京医療大学附属","学校法人 SBC東京医療大学"]),
    ("⑥株式会社SBC湘南接骨院", ["株式会社SBC湘南接骨院"]),
    ("⑦SBCメディカルグループ株式会社", ["SBCメディカルグループ株式会社","株式会社 MG"]),
    ("⑧Shoubikai Medical Vietnam Co., Ltd.", ["Shoubikai Medical Vietnam Co., Ltd."]),
    ("⑨WWMG", ["WWMG"]),("⑩DS", ["DS"]),("⑪DSS", ["DSS"]),("⑫DSS(FC)", ["DSS(FC)"]),("⑬WWFC", ["WWFC"]),("⑭RCC", ["RCC"]),
]

JIKEI_BRAND_COLS = [
    ("湘南美容クリニック","美容外科・皮膚科"),("湘南美容クリニック","オンライン"),("湘南美容クリニック","スキンLab"),
    ("湘南歯科クリニック",None),("湘南AGAクリニック",None),("湘南美容皮フ科",None),("湘南皮膚科クリニック",None),
    ("SBC NEO Skin Clinic",None),("肌の青空クリニック",None),("湘南内科皮フ科クリニック",None),
    ("湘南美容皮フ科内科クリニック",None),("イテウォン",None),("湘南メディカル記念病院",None),
    ("新宿近視クリニック",None),("西新宿整形外科",None),("SBC横浜駅前整形外科",None),
    ("六本木レディース",None),("神奈川レディース",None),("リッツ美容外科",None),
    ("リゼクリニック",None),("ゴリラクリニック",None),("JUNCLINIC",None),
    ("SBC東京医療大学付属クリニック",None),("SBC東京接骨院",None),("SBC BODY ARCHI",None),
    ("湘南美容クリニック（ﾍﾞﾄﾅﾑ）",None),("The Chelsea Clinic",None),("Chelsea Aesthetics",None),
    ("Gangnam Laser Clinic",None),("SkinGo!",None),("Wen & Weng Family Clinic",None),("Rochor Centre Clinic",None),
]
EXISTING_GROUP_END = 19
LABEL_IR = "IR・広報用（除：Holdingsへの収益貢献なし、含：OrangeTwist）"
LABEL_ALL = "全拠点"

REGION_HOUJIN_ORDER = {
    "国内": [
        "医療法人 湘美会","医療法人社団 孝和会","医療法人社団 菜寿会","医療法人社団 愛恵会",
        "医療法人社団 樹慶会","医療法人社団 リッツ美容外科","一般社団法人MASA","健美会","法人無し（個人開設）",
        "医療法人社団 風林会","医療法人 きびたき会","医療法人社団 百花会",
        "医療法人社団十二会","医療法人社団美咲会","一般社団法人美央斗会",
        "㈻SBC東京医療大学附属","株式会社SBC湘南接骨院","株式会社ボディアーキ・ジャパン"
    ],
    "海外": [
        "Shoubikai Medical Vietnam Co., Ltd.","WWMG","DS","DSS","DSS(FC)","WWFC","RCC"
    ]
}


def get_path():
    if BOX_PATH.exists(): return BOX_PATH
    if LOCAL_PATH.exists(): return LOCAL_PATH
    raise FileNotFoundError("Excelファイルが見つかりません")

def to_ts(v):
    if isinstance(v, pd.Timestamp): return v
    if pd.isna(v): return None
    try: return pd.Timestamp(v)
    except: return None

def load_data():
    path = get_path()
    df = pd.read_excel(path, sheet_name="クリニック一覧")
    for c in ["開院日","MA日","移転拡張日","業態転換日","閉院日"]:
        if c in df.columns:
            df[c] = pd.to_datetime(df[c], errors="coerce")
    def parse_limit(v):
        if isinstance(v, pd.Timestamp) and not pd.isna(v): return v.strftime("%Y/%m")
        if pd.isna(v): return ""
        return str(v).strip()
    excl = "猶予期限（業態転換後の院が開院するまでの期間）"
    if excl in df.columns:
        df[excl] = df[excl].apply(parse_limit)
    return df

def load_brand_settings():
    try:
        path = get_path()
        raw = pd.read_excel(path, sheet_name="ブランド設定", header=None)
        header_row = None
        for i, row in raw.iterrows():
            if any("ブランド名" in str(v) for v in row.values):
                header_row = i; break
        if header_row is None: return [], [], []
        df = pd.read_excel(path, sheet_name="ブランド設定", header=header_row)
        df = df.dropna(subset=["ブランド名"])
        if "表示順" in df.columns: df = df.sort_values("表示順")
        cols, flags, excl_pr = [], [], []
        for _, row in df.iterrows():
            b = str(row.get("ブランド名","") or "").strip()
            g_val = row.get("業態","")
            g = str(g_val).strip() if pd.notna(g_val) and str(g_val).strip() not in ("","nan") else None
            ex = str(row.get("既存グループ合計","") or "").strip()
            ep = str(row.get("広報IR除外","") or "").strip()
            if b:
                cols.append((b,g)); flags.append(ex=="○")
                if ep=="○": excl_pr.append(b)
        return cols, flags, excl_pr
    except: return [], [], []

def get_base_date(row):
    ma = to_ts(row.get("MA日")); op = to_ts(row.get("開院日"))
    return ma if ma is not None else op

def get_brand(row):
    b = str(row.get("ブランド","") or "").strip()
    return "Chelsea Aesthetics" if b=="Chelsea Asthetics" else b

def get_region(row):
    rv = row.get("海外／国内","")
    r = "" if pd.isna(rv) else str(rv).strip()
    if not r:
        b = str(row.get("ブランド","") or ""); h = str(row.get("法人名","") or "")
        return "海外" if any(kw in b or kw in h for kw in OVERSEAS_KW) else "国内"
    return r

def check_active(row, target):
    ma = to_ts(row.get("MA日")); op = to_ts(row.get("開院日"))
    base = ma if ma is not None else op
    if base is None: return False
    d_conv = to_ts(row.get("業態転換日")); d_close = to_ts(row.get("閉院日"))
    final_end = d_conv if d_conv is not None else d_close
    base_only = pd.Timestamp(base.year, base.month, base.day)
    return base_only <= target and (final_end is None or final_end > target)

def month_end_ts(y, m): return pd.Timestamp(y, m, calendar.monthrange(y,m)[1], 23,59,59)
def month_start_ts(y, m): return pd.Timestamp(y, m, 1, 0, 0, 0)

def aggregate(df, me, ms, target_brands, exclude_pr):
    r1, r2, r3 = {}, {}, {}
    for _, row in df.iterrows():
        brand = get_brand(row); gyoutai = str(row.get("業態","") or "").strip()
        houjin = str(row.get("法人名","") or "個人/その他").strip()
        region = get_region(row)
        if not brand or not gyoutai or brand not in target_brands: continue
        is_pr = brand not in exclude_pr
        if check_active(row, me):
            k1 = brand+"|"+gyoutai
            r1.setdefault(k1,{"all":0,"pr":0}); r1[k1]["all"]+=1
            if is_pr: r1[k1]["pr"]+=1
            k2 = region+"|"+houjin
            r2.setdefault(k2,{"all":0,"pr":0}); r2[k2]["all"]+=1
            if is_pr: r2[k2]["pr"]+=1
        if check_active(row, ms):
            if houjin in HOUJIN_ORDER:
                r3.setdefault(houjin,{"all":0})
                if houjin not in EXCLUDE_HOUJIN: r3[houjin]["all"]+=1
    return r1, r2, r3

def df_to_html_table(df, highlight_last=True, orange_last=False):
    rows_html = ""
    for i, (_, row) in enumerate(df.iterrows()):
        is_last = i == len(df) - 1
        is_second_last = i == len(df) - 2
        is_third_last = i == len(df) - 3
        if orange_last and is_last:
            style = f'background:{C_ORANGE};color:white;font-weight:bold'
        elif orange_last and is_second_last:
            style = f'background:{C_ORANGE};color:white'
        elif highlight_last and is_last:
            style = f'background:{C_TOTAL};font-weight:bold'
        elif highlight_last and (is_second_last or is_third_last) and orange_last:
            style = f'background:{C_TOTAL};font-weight:bold'
        else:
            style = 'background:white'
        cells = "".join(f'<td style="padding:6px 10px;border:1px solid #ddd">{v}</td>' for v in row)
        rows_html += f'<tr style="{style}">{cells}</tr>'
    headers = "".join(f'<th style="padding:8px 10px;background:{C_HEADER};color:white;border:1px solid #555">{c}</th>' for c in df.columns)
    return f'<table style="border-collapse:collapse;width:100%;font-size:14px"><thead><tr>{headers}</tr></thead><tbody>{rows_html}</tbody></table>'

def build_history(df, target_brands, exclude_pr):
    """年別・月別の開院・閉院・業態転換データを構築"""
    today = date.today()
    all_dates = pd.concat([df["開院日"].dropna(), df["MA日"].dropna(),
                           df["閉院日"].dropna(), df["業態転換日"].dropna()])
    data_max_year = int(all_dates.dt.year.max()) if not all_dates.empty else today.year
    hist_years = list(range(2025, max(data_max_year, today.year) + 1))
    result = {}
    for year in hist_years:
        monthly = {}
        for month in range(1, 13):
            ms = pd.Timestamp(year, month, 1)
            me = pd.Timestamp(year, month, calendar.monthrange(year, month)[1])
            opened, closed, convert = [], [], []
            for _, row in df.iterrows():
                brand = get_brand(row)
                if not brand or brand not in target_brands: continue
                base = get_base_date(row)
                d_conv = to_ts(row.get("業態転換日")); d_close = to_ts(row.get("閉院日"))
                tenkan_mae = str(row.get("転換前業態","") or "").strip()
                is_converted = tenkan_mae not in ("","nan")
                if base is not None:
                    base_only = pd.Timestamp(base.year, base.month, base.day)
                    if ms <= base_only <= me:
                        opened.append({
                            "種別": "🔵 業態転換" if is_converted else "🟢 新規開院",
                            "開院日": base_only.strftime("%Y/%m/%d"),
                            "院名": row.get("正式名称",""), "ブランド": brand,
                            "業態": str(row.get("業態","") or ""),
                            "法人名": str(row.get("法人名","") or ""),
                            "国内／海外": get_region(row),
                        })
                if d_conv is not None:
                    conv_only = pd.Timestamp(d_conv.year, d_conv.month, d_conv.day)
                    if ms <= conv_only <= me:
                        before = str(row.get("転換前業態","") or "").strip()
                        if not before or before=="nan": before = str(row.get("業態","") or "").strip()
                        after_n = str(row.get("転換後院名","") or "").strip()
                        after_g = str(row.get("転換後業態","") or "").strip()
                        convert.append({
                            "転換前院名": row.get("正式名称",""), "転換前ブランド": brand,
                            "転換前業態": before if before else "―", "　": "→",
                            "転換後院名": after_n if after_n else "―",
                            "転換後業態": after_g if after_g else "―",
                            "法人名": str(row.get("法人名","") or ""),
                            "国内／海外": get_region(row),
                            "業態転換日": conv_only.strftime("%Y/%m/%d"),
                        })
                close_ref = d_close if d_close is not None else d_conv
                if close_ref is not None:
                    close_only = pd.Timestamp(close_ref.year, close_ref.month, close_ref.day)
                    if ms <= close_only <= me:
                        closed.append({
                            "種別": "🔵 業態転換" if d_conv is not None else "🔴 閉院",
                            "閉院日": close_only.strftime("%Y/%m/%d"),
                            "院名": row.get("正式名称",""), "ブランド": brand,
                            "業態": str(row.get("業態","") or ""),
                            "法人名": str(row.get("法人名","") or ""),
                            "国内／海外": get_region(row),
                        })
            monthly[month] = {"opened": opened, "closed": closed, "convert": convert}
        result[year] = monthly
    return result, hist_years

def build_history_html(history, hist_years, mode="openclose", prefix=""):
    """開院・閉院または業態転換のアコーディオンHTMLを生成"""
    html = ""
    uid = 0
    for year in hist_years:
        monthly = history[year]
        if mode == "openclose":
            ytotal_o = sum(len(v["opened"]) for v in monthly.values())
            ytotal_c = sum(len(v["closed"]) for v in monthly.values())
            if ytotal_o == 0 and ytotal_c == 0: continue
            year_label = f"📅 {year}年　｜　🟢 開院 {ytotal_o} 院　　🔴 閉院 {ytotal_c} 院"
            color = C_HEADER
        else:
            ytotal_k = sum(len(v["convert"]) for v in monthly.values())
            if ytotal_k == 0: continue
            year_label = f"📅 {year}年　｜　🔵 業態転換 {ytotal_k} 院"
            color = C_BLUE

        html += f'<div style="background:{color};color:white;padding:10px 16px;font-weight:bold;font-size:16px;border-radius:6px;margin-top:20px">{year_label}</div>'

        for month in range(1, 13):
            opened  = monthly[month]["opened"]
            closed  = monthly[month]["closed"]
            convert = monthly[month]["convert"]
            uid += 1

            if mode == "openclose":
                n_o = len(opened); n_c = len(closed)
                if n_o == 0 and n_c == 0: continue
                label = f"{year}年{month}月　🟢 開院 {n_o}院　　🔴 閉院 {n_c}院"
                inner = ""
                if n_o > 0:
                    inner += f'<div style="border-left:4px solid #27AE60;background:#D5F5E3;padding:5px 10px;font-weight:bold;margin-bottom:6px">🟢 開院　{n_o}院</div>'
                    inner += df_to_html_table(pd.DataFrame(opened), highlight_last=False)
                    inner += "<br>"
                if n_c > 0:
                    inner += f'<div style="border-left:4px solid #E74C3C;background:#FADBD8;padding:5px 10px;font-weight:bold;margin-bottom:6px">🔴 閉院　{n_c}院</div>'
                    inner += df_to_html_table(pd.DataFrame(closed), highlight_last=False)
            else:
                n_k = len(convert)
                if n_k == 0: continue
                label = f"{year}年{month}月　🔵 業態転換 {n_k}院"
                inner = f'<div style="border-left:4px solid {C_BLUE};background:#D6EAF8;padding:5px 10px;font-weight:bold;margin-bottom:6px">🔵 業態転換　{n_k}院</div>'
                inner += df_to_html_table(pd.DataFrame(convert), highlight_last=False)

            html += f"""
<div style="margin-top:8px;border:1px solid #ddd;border-radius:6px;overflow:hidden">
  <div onclick="toggleAcc('{prefix}acc{uid}','{prefix}arr{uid}')" style="padding:10px 16px;cursor:pointer;background:#f8f9fa;display:flex;align-items:center;gap:8px">
    <span id="{prefix}arr{uid}" style="font-size:12px">▶</span> {label}
  </div>
  <div id="{prefix}acc{uid}" style="display:none;padding:16px">
    {inner}
  </div>
</div>"""
    return html


def aggregate_with_houjin(df, target_brands, exclude_pr, me, ms):
    """
    Returns:
      r1: {brand|gyoutai: {all, pr}}  (month end)
      r2: {region|houjin: {all, pr}}  (month end)
      r3: {houjin: {all}}             (month start)
      houjin_brand: {group_name: {brand_label: count}}  (month end)
    """
    r1, r2, r3 = {}, {}, {}
    # houjin_brand: group_name -> brand_label -> count
    houjin_brand = {gname: {} for gname, _ in HOUJIN_GROUPS}

    for _, row in df.iterrows():
        brand = get_brand(row)
        gyoutai = str(row.get("業態", "") or "").strip()
        houjin = str(row.get("法人名", "") or "個人/その他").strip()
        region = get_region(row)
        if not brand or not gyoutai or brand not in target_brands:
            continue
        is_pr = brand not in exclude_pr

        if check_active(row, me):
            k1 = brand + "|" + gyoutai
            r1.setdefault(k1, {"all": 0, "pr": 0})
            r1[k1]["all"] += 1
            if is_pr:
                r1[k1]["pr"] += 1

            k2 = region + "|" + houjin
            r2.setdefault(k2, {"all": 0, "pr": 0})
            r2[k2]["all"] += 1
            if is_pr:
                r2[k2]["pr"] += 1

            # 法人グループ × ブランド
            for gname, members in HOUJIN_GROUPS:
                if houjin in members:
                    # build brand label same as JIKEI_BRAND_COLS
                    for (bc_brand, bc_gyoutai) in JIKEI_BRAND_COLS:
                        if bc_brand == brand:
                            if bc_gyoutai is None or bc_gyoutai == gyoutai:
                                blabel = f"{bc_brand}({bc_gyoutai})" if bc_gyoutai else bc_brand
                                houjin_brand[gname][blabel] = houjin_brand[gname].get(blabel, 0) + 1
                                break

        if check_active(row, ms):
            if houjin in HOUJIN_ORDER:
                r3.setdefault(houjin, {"all": 0})
                if houjin not in EXCLUDE_HOUJIN:
                    r3[houjin]["all"] += 1

    return r1, r2, r3, houjin_brand


def build_all_monthly_data(df, target_brands, exclude_pr):
    """
    2021/01 から今月まで各月の集計データを返す
    Returns dict keyed by "YYYY/MM"
    """
    today = date.today()
    sy, sm = 2021, 1
    all_data = {}

    while (sy, sm) <= (today.year, today.month):
        ym_key = f"{sy}/{sm:02d}"
        me = month_end_ts(sy, sm)
        ms_ts = month_start_ts(sy, sm)

        r1, r2, r3, houjin_brand = aggregate_with_houjin(df, target_brands, exclude_pr, me, ms_ts)

        # brand_rows
        brand_rows = []
        for (brand, gyoutai) in JIKEI_BRAND_COLS:
            if gyoutai is not None:
                keys = [k for k in r1 if k == f"{brand}|{gyoutai}"]
            else:
                keys = [k for k in r1 if k.startswith(f"{brand}|")]
            cnt_all = sum(r1[k]["all"] for k in keys)
            cnt_pr = sum(r1[k]["pr"] for k in keys)
            label = f"{brand}({gyoutai})" if gyoutai else brand
            brand_rows.append({"label": label, "pr": cnt_pr, "all": cnt_all})

        sum_pr = sum(row["pr"] for row in brand_rows)
        sum_all = sum(row["all"] for row in brand_rows)

        # existing group sum (first EXISTING_GROUP_END brand cols)
        existing_sum = sum(brand_rows[i]["all"] for i in range(min(EXISTING_GROUP_END, len(brand_rows))))

        # IR広報用
        ir_sum = sum_pr + ORANGE_TWIST_COUNT

        # jikei_row: brand counts + summary columns
        jikei_row = {}
        for r in brand_rows:
            jikei_row[r["label"]] = r["all"]
        jikei_row["既存G合計"] = existing_sum
        jikei_row[LABEL_ALL] = sum_all
        jikei_row["OrangeTwist"] = ORANGE_TWIST_COUNT
        jikei_row[LABEL_IR] = ir_sum
        # per houjin group totals
        for gname, members in HOUJIN_GROUPS:
            cnt = 0
            for _, row in df.iterrows():
                brand = get_brand(row)
                houjin = str(row.get("法人名", "") or "個人/その他").strip()
                if not brand or brand not in target_brands:
                    continue
                if houjin not in members:
                    continue
                if check_active(row, me):
                    cnt += 1
            jikei_row[gname] = cnt
        # houjin_brand per group
        for gname in houjin_brand:
            for blabel, cnt in houjin_brand[gname].items():
                jikei_row[f"{gname}|{blabel}"] = cnt

        # region_dom / region_ovs
        region_dom = []
        region_ovs = []
        for region_label in ["国内", "海外"]:
            for k in sorted([k for k in r2 if k.startswith(region_label + "|")]):
                parts = k.split("|", 1)
                entry = {"法人名": parts[1], "pr": r2[k]["pr"], "all": r2[k]["all"]}
                if region_label == "国内":
                    region_dom.append(entry)
                else:
                    region_ovs.append(entry)

        # houjin_rows
        houjin_rows = []
        for h in HOUJIN_ORDER:
            cnt = r3.get(h, {}).get("all", 0)
            houjin_rows.append({"法人名": h, "all": cnt})

        all_data[ym_key] = {
            "brand_rows": brand_rows,
            "sum_pr": sum_pr,
            "sum_all": sum_all,
            "jikei_row": jikei_row,
            "region_dom": region_dom,
            "region_ovs": region_ovs,
            "houjin_rows": houjin_rows,
        }

        sm += 1
        if sm > 12:
            sm = 1
            sy += 1

    return all_data


def build_timeseries_html(all_data):
    """時系列推移テーブルHTMLを生成（法人グループ×ブランドの階層ヘッダー付き）"""
    if not all_data:
        return "<p>データなし</p>"

    brand_labels = [f"{b}({g})" if g else b for b, g in JIKEI_BRAND_COLS]

    # 第1パス: 各法人グループで出現するブランドを収集
    # jikei_row には "{gname}|{blabel}" 形式のキーが入っている
    group_brands = {}  # gname -> ordered list of blabels that appear
    for gname, _ in HOUJIN_GROUPS:
        seen = set()
        ordered = []
        for ym, rec in all_data.items():
            jr = rec.get("jikei_row", {})
            for k, v in jr.items():
                if k.startswith(f"{gname}|") and v and v > 0:
                    blabel = k[len(gname)+1:]
                    if blabel not in seen:
                        seen.add(blabel)
                        ordered.append(blabel)
        # JIKEIの順序で並べ直す
        jikei_order = [f"{b}({g})" if g else b for b, g in JIKEI_BRAND_COLS]
        sorted_blabels = [bl for bl in jikei_order if bl in seen]
        # 未定義のものは末尾
        for bl in ordered:
            if bl not in sorted_blabels:
                sorted_blabels.append(bl)
        group_brands[gname] = sorted_blabels

    # スタイル定義
    def th_attr(colspan=1, rowspan=1, bg=C_HEADER):
        s = f'padding:6px 8px;background:{bg};color:white;border:1px solid #444;white-space:nowrap;font-size:11px;position:sticky;top:0;z-index:2'
        attrs = f'style="{s}"'
        if colspan > 1: attrs += f' colspan="{colspan}"'
        if rowspan > 1: attrs += f' rowspan="{rowspan}"'
        return attrs

    def th_sticky_attr(rowspan=2):
        s = f'padding:6px 8px;background:{C_HEADER};color:white;border:1px solid #444;white-space:nowrap;font-size:11px;position:sticky;top:0;left:0;z-index:3'
        return f'style="{s}" rowspan="{rowspan}"'

    # 第1行ヘッダー: 年月 | ブランド別(colspan=n_brand) | 既存G合計 | OrangeTwist | LABEL_IR | LABEL_ALL | 各法人グループ(colspan=サブブランド数+1)
    n_brand = len(brand_labels)
    row1 = f'<th {th_sticky_attr()}>年月</th>'
    row1 += f'<th {th_attr(colspan=n_brand)}>ブランド別</th>'
    row1 += f'<th {th_attr(rowspan=2, bg="#1E8449")}>既存G合計</th>'
    row1 += f'<th {th_attr(rowspan=2, bg=C_ORANGE)}>OrangeTwist</th>'
    row1 += f'<th {th_attr(rowspan=2, bg=C_ORANGE)}>{LABEL_IR}</th>'
    row1 += f'<th {th_attr(rowspan=2)}>{LABEL_ALL}</th>'
    for gname, _ in HOUJIN_GROUPS:
        sub_count = len(group_brands[gname])
        colspan = sub_count + 1  # サブブランド列 + 合計列
        row1 += f'<th {th_attr(colspan=colspan, bg=C_BLUE)}>{gname}</th>'

    # 第2行ヘッダー: ブランド名 | 各法人グループのサブブランド + 合計
    row2 = ""
    for label in brand_labels:
        row2 += f'<th {th_attr()}>{label}</th>'
    for gname, _ in HOUJIN_GROUPS:
        for blabel in group_brands[gname]:
            row2 += f'<th {th_attr(bg=C_BLUE)}>{blabel}</th>'
        row2 += f'<th {th_attr(bg=C_BLUE)}>合計</th>'

    thead = f"<thead><tr>{row1}</tr><tr>{row2}</tr></thead>"

    # tbody
    months_sorted = sorted(all_data.keys())
    tbody_rows = ""
    for ym in months_sorted:
        rec = all_data[ym]
        is_dec = ym.endswith("/12")
        row_bg = "#EBF5FB" if is_dec else "white"

        td_base = f'padding:5px 8px;border:1px solid #ddd;text-align:right;font-size:12px;background:{row_bg}'
        td_green = f'padding:5px 8px;border:1px solid #ddd;text-align:right;font-size:12px;background:#D5F5E3;font-weight:bold'
        td_orange = f'padding:5px 8px;border:1px solid #ddd;text-align:right;font-size:12px;background:#FAD7A0;font-weight:bold'
        td_sticky = f'padding:5px 8px;border:1px solid #ddd;text-align:center;font-size:12px;background:{row_bg};position:sticky;left:0;z-index:1;font-weight:bold;white-space:nowrap'

        cells = f'<td style="{td_sticky}">{ym}</td>'

        jr = rec.get("jikei_row", {})
        for label in brand_labels:
            cells += f'<td style="{td_base}">{jr.get(label, 0)}</td>'

        cells += f'<td style="{td_green}">{jr.get("既存G合計", 0)}</td>'
        cells += f'<td style="{td_orange}">{jr.get("OrangeTwist", ORANGE_TWIST_COUNT)}</td>'
        cells += f'<td style="{td_orange}">{jr.get(LABEL_IR, 0)}</td>'
        cells += f'<td style="{td_base}">{jr.get(LABEL_ALL, 0)}</td>'

        for gname, _ in HOUJIN_GROUPS:
            for blabel in group_brands[gname]:
                key = f"{gname}|{blabel}"
                cells += f'<td style="{td_base}">{jr.get(key, 0)}</td>'
            cells += f'<td style="{td_base}">{jr.get(gname, 0)}</td>'

        tbody_rows += f'<tr>{cells}</tr>'

    tbody = f"<tbody>{tbody_rows}</tbody>"
    table = f'<table style="border-collapse:collapse;font-size:12px;min-width:100%">{thead}{tbody}</table>'
    return f'<div style="overflow-x:auto;max-height:70vh;overflow-y:auto;border:1px solid #ddd;border-radius:4px">{table}</div>'


def generate():
    print("データを読み込み中...")
    df = load_data()
    brand_cols, existing_flags, exclude_pr = load_brand_settings()
    if not brand_cols:
        brand_cols = [(b,None) for b in TARGET_BRANDS]
        existing_flags = [True]*19 + [False]*(len(brand_cols)-19)
        exclude_pr = EXCLUDE_PR

    today = date.today()
    y, m = today.year, today.month
    me = month_end_ts(y, m)
    ms = month_start_ts(y, m)
    r1, r2, r3 = aggregate(df, me, ms, [b for b,_ in brand_cols], exclude_pr)

    sum_pr = sum(v["pr"] for v in r1.values())
    sum_all = sum(v["all"] for v in r1.values())

    # ── ブランド×業態表 ──
    brand_rows_disp = []
    for brand, gyoutai in brand_cols:
        if gyoutai: keys=[k for k in r1 if k==f"{brand}|{gyoutai}"]
        else: keys=[k for k in r1 if k.startswith(f"{brand}|")]
        cnt_all=sum(r1[k]["all"] for k in keys); cnt_pr=sum(r1[k]["pr"] for k in keys)
        label = f"{brand}({gyoutai})" if gyoutai else brand
        brand_rows_disp.append({"ブランド":label,"広報・IR用":cnt_pr,"全拠点":cnt_all})
    brand_df = pd.DataFrame(brand_rows_disp)
    total_row = pd.DataFrame([
        {"ブランド":"集計内訳合計","広報・IR用":sum_pr,"全拠点":sum_all},
        {"ブランド":"海外(OrangeTwist)加算","広報・IR用":ORANGE_TWIST_COUNT,"全拠点":"－"},
        {"ブランド":"最終報告数値","広報・IR用":sum_pr+ORANGE_TWIST_COUNT,"全拠点":sum_all},
    ])
    brand_df = pd.concat([brand_df, total_row], ignore_index=True)

    # ブランド×業態表HTML（数値右寄せ、pr≠all で黄色ハイライト）
    def brand_table_html(df_in):
        rows_html = ""
        n = len(df_in)
        for i, (_, row) in enumerate(df_in.iterrows()):
            is_last = i == n - 1
            is_second_last = i == n - 2
            is_third_last = i == n - 3
            pr_val = row.get("広報・IR用", "")
            all_val = row.get("全拠点", "")
            highlight = (str(pr_val) != str(all_val) and str(all_val) != "－" and i < n - 3)
            if i >= n - 3:
                if is_last:
                    # 最終報告数値 → オレンジ
                    row_bg = C_ORANGE
                    row_color = "white"
                    row_fw = "bold"
                elif is_second_last:
                    # 海外(OrangeTwist)加算 → 黄色
                    row_bg = "#FFF9C4"
                    row_color = "#333"
                    row_fw = "normal"
                else:
                    # 集計内訳合計 → オレンジ
                    row_bg = C_ORANGE
                    row_color = "white"
                    row_fw = "bold"
            elif highlight:
                row_bg = "#FFFDE7"
                row_color = "#333"
                row_fw = "normal"
            else:
                row_bg = "white"
                row_color = "#333"
                row_fw = "normal"
            cells = ""
            for ci, (col, val) in enumerate(row.items()):
                align = "right" if ci > 0 else "left"
                cells += f'<td style="padding:6px 10px;border:1px solid #ddd;text-align:{align};background:{row_bg};color:{row_color};font-weight:{row_fw}">{val}</td>'
            rows_html += f"<tr>{cells}</tr>"
        headers = "".join(f'<th style="padding:8px 10px;background:{C_HEADER};color:white;border:1px solid #555;text-align:left">{c}</th>' for c in df_in.columns)
        return f'<table style="border-collapse:collapse;width:100%;font-size:14px"><thead><tr>{headers}</tr></thead><tbody>{rows_html}</tbody></table>'

    brand_table = brand_table_html(brand_df)

    # ── 国内/海外×法人表 ──
    # r2 から 国内/海外 それぞれのデータを取得
    raw_dom, raw_ovs = {}, {}
    for k, v in r2.items():
        if k.startswith("国内|"):
            houjin_name = k.split("|", 1)[1]
            raw_dom[houjin_name] = {"pr": v["pr"], "all": v["all"]}
        elif k.startswith("海外|"):
            houjin_name = k.split("|", 1)[1]
            raw_ovs[houjin_name] = {"pr": v["pr"], "all": v["all"]}

    def build_region_table_html(raw_dom_data, raw_ovs_data):
        """国内/海外×法人の統合テーブルHTMLを生成"""
        C_YELLOW = "#FFF9C4"
        C_LIGHT_YELLOW = "#FFFDE7"
        td_style = 'padding:6px 10px;border:1px solid #ddd'
        th_style = f'padding:8px 10px;background:{C_HEADER};color:white;border:1px solid #555'

        rows_html = ""

        def make_row(cells_data, bg, color="inherit", fw="normal"):
            cells = ""
            for ci, val in enumerate(cells_data):
                align = "right" if ci > 0 else "left"
                cells += f'<td style="{td_style};text-align:{align};background:{bg};color:{color};font-weight:{fw}">{val}</td>'
            return f"<tr>{cells}</tr>"

        # 国内行
        dom_order = REGION_HOUJIN_ORDER["国内"]
        dom_seen = set()
        dom_rows_ordered = []
        for h in dom_order:
            if h in raw_dom_data:
                dom_rows_ordered.append(h)
                dom_seen.add(h)
        # 未定義の法人を末尾に追加
        for h in sorted(raw_dom_data.keys()):
            if h not in dom_seen:
                dom_rows_ordered.append(h)

        dom_sum_pr = 0
        dom_sum_all = 0
        for h in dom_rows_ordered:
            v = raw_dom_data[h]
            pr_val = v["pr"]; all_val = v["all"]
            dom_sum_pr += pr_val; dom_sum_all += all_val
            highlight = str(pr_val) != str(all_val)
            bg = C_LIGHT_YELLOW if highlight else "white"
            rows_html += make_row([h, pr_val, all_val], bg)

        # 国内 小計
        rows_html += make_row(["国内 小計", dom_sum_pr, dom_sum_all], C_ORANGE, "white", "bold")

        # 海外行
        ovs_order = REGION_HOUJIN_ORDER["海外"]
        ovs_seen = set()
        ovs_rows_ordered = []
        for h in ovs_order:
            if h in raw_ovs_data:
                ovs_rows_ordered.append(h)
                ovs_seen.add(h)
        for h in sorted(raw_ovs_data.keys()):
            if h not in ovs_seen:
                ovs_rows_ordered.append(h)

        ovs_sum_pr = 0
        ovs_sum_all = 0
        # OrangeTwist判定：海外法人のうちOrangeTwist系を除いた合計のため、ここでは全海外を合計
        for h in ovs_rows_ordered:
            v = raw_ovs_data[h]
            pr_val = v["pr"]; all_val = v["all"]
            ovs_sum_pr += pr_val; ovs_sum_all += all_val
            highlight = str(pr_val) != str(all_val)
            bg = C_LIGHT_YELLOW if highlight else "white"
            rows_html += make_row([h, pr_val, all_val], bg)

        # 海外 小計
        rows_html += make_row(["海外 小計", ovs_sum_pr, ovs_sum_all], C_ORANGE, "white", "bold")

        # 合計（OrangeTwist除く）= 全拠点合計（OrangeTwistはsum_prに含まれていないため広報用合計を使用）
        total_pr = dom_sum_pr + ovs_sum_pr
        total_all = dom_sum_all + ovs_sum_all
        rows_html += make_row(["合計（OrangeTwist除く）", total_pr, total_all], C_ORANGE, "white", "bold")

        # OrangeTwist加算
        rows_html += make_row(["OrangeTwist加算", ORANGE_TWIST_COUNT, "－"], C_YELLOW, "#333", "normal")

        # 最終報告数値
        rows_html += make_row(["最終報告数値", total_pr + ORANGE_TWIST_COUNT, total_all], C_ORANGE, "white", "bold")

        headers = "".join(f'<th style="{th_style};text-align:left">{c}</th>' for c in ["法人名", "広報・IR用", "全拠点"])
        return f'<table style="border-collapse:collapse;width:100%;font-size:14px"><thead><tr>{headers}</tr></thead><tbody>{rows_html}</tbody></table>'

    region_table_html_str = build_region_table_html(raw_dom, raw_ovs)

    # 後方互換のためdom_df/ovs_dfも保持（使用箇所は後でregion_table_html_strに差し替え）
    dom_tmp, ovs_tmp = [], []
    for k in sorted([k for k in r2 if k.startswith("国内|")]):
        p=k.split("|",1); dom_tmp.append({"法人名":p[1],"広報・IR用":r2[k]["pr"],"全拠点":r2[k]["all"]})
    for k in sorted([k for k in r2 if k.startswith("海外|")]):
        p=k.split("|",1); ovs_tmp.append({"法人名":p[1],"広報・IR用":r2[k]["pr"],"全拠点":r2[k]["all"]})
    dom_df=pd.DataFrame(dom_tmp); ovs_df=pd.DataFrame(ovs_tmp)

    # ── 法人別表 ──
    houjin_rows=[]; total_h=0
    for h in HOUJIN_ORDER:
        cnt=r3.get(h,{}).get("all",0); houjin_rows.append({"法人名":h,"全拠点":cnt}); total_h+=cnt
    houjin_rows.append({"法人名":"合計","全拠点":total_h})
    houjin_df=pd.DataFrame(houjin_rows)

    # ── 時系列データ構築 ──
    print("時系列推移データを集計中（2021/01〜現在）...")
    all_monthly_data = build_all_monthly_data(df, [b for b,_ in brand_cols], exclude_pr)
    timeseries_html = build_timeseries_html(all_monthly_data)

    # JSON埋め込み用（シングルクォートをエスケープ）
    json_str = json.dumps(all_monthly_data, ensure_ascii=False).replace("'", "\\'")

    # ── ブランド別棒グラフ ──
    plot_df = brand_df[brand_df["ブランド"].isin([f"{b}({g})" if g else b for b,g in brand_cols])].copy()
    fig2=px.bar(plot_df,x="全拠点",y="ブランド",orientation="h",height=550,color_discrete_sequence=[C_BLUE])
    fig2.update_layout(yaxis=dict(autorange="reversed"),margin=dict(t=30))
    chart2_html = fig2.to_html(full_html=False, include_plotlyjs="cdn")

    # ── 開院・閉院・業態転換履歴 ──
    print("開院・閉院・業態転換履歴を集計中...")
    history, hist_years = build_history(df, [b for b,_ in brand_cols], exclude_pr)
    openclose_html = build_history_html(history, hist_years, mode="openclose", prefix="oc")
    convert_html   = build_history_html(history, hist_years, mode="convert",   prefix="cv")

    # 期間セレクターの選択肢生成
    year_options = "".join(f'<option value="{yr}">{yr}</option>' for yr in range(2021, today.year+1))
    month_options = "".join(f'<option value="{mo:02d}">{mo}月</option>' for mo in range(1, 13))

    brand_labels_js = [f"{b}({g})" if g else b for b, g in JIKEI_BRAND_COLS]
    brand_labels_json = json.dumps(brand_labels_js, ensure_ascii=False)

    report_date = f"{y}年{m}月末"
    html = f"""<!DOCTYPE html>
<html lang="ja">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>クリニック数ダッシュボード {report_date}</title>
<style>
  body{{font-family:'Meiryo',sans-serif;margin:0;background:#f5f5f5;color:#333}}
  .header{{background:{C_HEADER};color:white;padding:20px 30px}}
  .header h1{{margin:0;font-size:22px}}
  .header p{{margin:5px 0 0;opacity:.7;font-size:13px}}
  .kpi{{display:flex;gap:16px;padding:20px 30px;background:white;border-bottom:1px solid #ddd}}
  .kpi-card{{flex:1;background:#f8f9fa;border-radius:8px;padding:16px;text-align:center;border-left:4px solid {C_BLUE}}}
  .kpi-card.orange{{border-left-color:{C_ORANGE}}}
  .kpi-label{{font-size:12px;color:#666;margin-bottom:4px}}
  .kpi-value{{font-size:28px;font-weight:bold;color:{C_HEADER}}}
  .period-bar{{background:white;padding:12px 30px;border-bottom:1px solid #ddd;display:flex;align-items:center;gap:12px;flex-wrap:wrap}}
  .period-bar select{{padding:4px 8px;border:1px solid #ccc;border-radius:4px;font-size:13px}}
  .period-bar button{{padding:6px 16px;background:{C_BLUE};color:white;border:none;border-radius:4px;cursor:pointer;font-size:13px}}
  .period-bar button:hover{{background:#1a6fa8}}
  .tabs{{display:flex;background:white;border-bottom:2px solid #ddd;padding:0 30px;flex-wrap:wrap}}
  .tab{{padding:12px 20px;cursor:pointer;border-bottom:3px solid transparent;font-size:14px;color:#666}}
  .tab.active{{border-bottom-color:{C_BLUE};color:{C_BLUE};font-weight:bold}}
  .content{{display:none;padding:24px 30px}}
  .content.active{{display:block}}
  .section-title{{background:{C_GREEN};padding:8px 14px;font-weight:bold;border-radius:4px;margin-bottom:12px}}
  .grid-2{{display:grid;grid-template-columns:1fr 1fr;gap:20px}}
  .box{{background:white;border-radius:8px;padding:16px;box-shadow:0 1px 4px rgba(0,0,0,.08)}}
  .updated{{text-align:right;font-size:12px;color:#999;padding:8px 30px}}
  .compare-table td, .compare-table th{{padding:5px 10px;border:1px solid #ddd;font-size:13px}}
  .compare-table th{{background:{C_HEADER};color:white}}
  .compare-table .neg{{color:#c0392b}}
  .compare-table .pos{{color:#27ae60}}
</style>
</head>
<body>
<div class="header">
  <h1>クリニック数ダッシュボード</h1>
  <p>{report_date} 時点</p>
</div>

<div class="kpi">
  <div class="kpi-card orange">
    <div class="kpi-label">IR・広報用（除：Holdingsへの収益貢献なし、含：OrangeTwist）</div>
    <div class="kpi-value">{sum_pr+ORANGE_TWIST_COUNT:,} 院</div>
  </div>
  <div class="kpi-card">
    <div class="kpi-label">全拠点</div>
    <div class="kpi-value">{sum_all:,} 院</div>
  </div>
  <div class="kpi-card">
    <div class="kpi-label">OrangeTwist（別管理）</div>
    <div class="kpi-value">{ORANGE_TWIST_COUNT} 院</div>
  </div>
</div>

<div class="period-bar">
  <span>表示期間：</span>
  <select id="startYear">{year_options}</select>年
  <select id="startMonth">{month_options}</select>月 〜
  <select id="endYear">{year_options}</select>年
  <select id="endMonth">{month_options}</select>月
  <button onclick="filterPeriod()">適用</button>
  <span id="periodStatus" style="font-size:12px;color:#666"></span>
</div>

<div class="tabs">
  <div class="tab active" onclick="showTab('brand',this)">📋 ブランド×業態</div>
  <div class="tab" onclick="showTab('region',this)">🌏 国内／海外×法人</div>
  <div class="tab" onclick="showTab('houjin',this)">🏢 法人別（月初時点／経理用）</div>
  <div class="tab" onclick="showTab('trend',this)">📈 時系列推移</div>
  <div class="tab" onclick="showTab('history',this)">🏥 開院・閉院履歴</div>
  <div class="tab" onclick="showTab('convert',this)">🔵 業態転換履歴</div>
</div>

<div id="brand" class="content active">
  <div class="box">
    <div class="section-title" id="brandTableTitle">{report_date} ブランド×業態</div>
    <div id="brandTableContainer">
      {brand_table}
    </div>
  </div>
  <div class="box" style="margin-top:20px">
    <div class="section-title">期間比較（開始月 vs 終了月）</div>
    <div id="compareTableContainer"><p style="color:#999;font-size:13px">期間を選択して「適用」を押すと比較表が表示されます。</p></div>
  </div>
  <div class="box" style="margin-top:20px">{chart2_html}</div>
</div>

<div id="region" class="content">
  <div class="box">
    <div class="section-title" id="regionTableTitle">{report_date} 国内／海外×法人</div>
    {region_table_html_str}
  </div>
</div>

<div id="houjin" class="content">
  <div class="box">
    <div class="section-title">{y}年{m}月初 法人別内訳（月初時点／経理用）</div>
    {df_to_html_table(houjin_df)}
  </div>
</div>

<div id="trend" class="content">
  <div class="box">
    <div class="section-title">時系列推移 2021/01〜{y}/{m:02d}</div>
    <div id="trendTableWrapper">
      {timeseries_html}
    </div>
  </div>
</div>

<div id="history" class="content">
  <h3 style="margin-top:0">開院・閉院 年別・月別履歴</h3>
  <div id="historyContainer">
    {openclose_html}
  </div>
</div>

<div id="convert" class="content">
  <h3 style="margin-top:0">業態転換 年別・月別履歴</h3>
  {convert_html}
</div>

<div class="updated">最終更新: {today.strftime('%Y/%m/%d')}</div>

<script>
const ALL_DATA = JSON.parse('{json_str}');
const BRAND_LABELS = {brand_labels_json};
const LABEL_IR = {json.dumps(LABEL_IR, ensure_ascii=False)};
const LABEL_ALL = {json.dumps(LABEL_ALL, ensure_ascii=False)};
const ORANGE_TWIST_COUNT = {ORANGE_TWIST_COUNT};

function showTab(id, el) {{
  document.querySelectorAll('.content').forEach(e=>e.classList.remove('active'));
  document.querySelectorAll('.tab').forEach(e=>e.classList.remove('active'));
  document.getElementById(id).classList.add('active');
  if(el) el.classList.add('active');
}}

function toggleAcc(accId, arrId) {{
  var el=document.getElementById(accId);
  var arr=document.getElementById(arrId);
  if(el.style.display==='none'){{el.style.display='block';arr.textContent='▼';}}
  else{{el.style.display='none';arr.textContent='▶';}}
}}

// 初期値：最新月を終了月に、1年前を開始月に
(function initPeriod() {{
  const keys = Object.keys(ALL_DATA).sort();
  if (!keys.length) return;
  const last = keys[keys.length - 1];
  const [ey, em] = last.split('/');
  // 1年前
  let sy = parseInt(ey) - 1, sm = parseInt(em);
  if (sy < 2021) {{ sy = 2021; sm = 1; }}
  document.getElementById('startYear').value = sy;
  document.getElementById('startMonth').value = String(sm).padStart(2,'0');
  document.getElementById('endYear').value = ey;
  document.getElementById('endMonth').value = em;
}})();

function filterPeriod() {{
  const sy = document.getElementById('startYear').value;
  const sm = document.getElementById('startMonth').value;
  const ey = document.getElementById('endYear').value;
  const em = document.getElementById('endMonth').value;
  const startKey = sy + '/' + sm;
  const endKey   = ey + '/' + em;

  if (startKey > endKey) {{
    alert('開始月が終了月より後になっています。');
    return;
  }}

  document.getElementById('periodStatus').textContent = '（' + startKey + ' 〜 ' + endKey + '）';

  // ── 時系列テーブルの行フィルタ ──
  const trendTable = document.querySelector('#trendTableWrapper table');
  if (trendTable) {{
    const rows = trendTable.querySelectorAll('tbody tr');
    rows.forEach(row => {{
      const ymCell = row.querySelector('td');
      if (!ymCell) return;
      const ym = ymCell.textContent.trim();
      row.style.display = (ym >= startKey && ym <= endKey) ? '' : 'none';
    }});
  }}

  // ── 履歴アコーディオンフィルタ ──
  filterHistory(startKey, endKey);

  // ── ブランド比較表 ──
  buildCompareTable(startKey, endKey);
}}

function filterHistory(startKey, endKey) {{
  // 開院・閉院タブとの連携はアコーディオンのdata属性がないため年月テキストで判断
  const histContainer = document.getElementById('historyContainer');
  if (!histContainer) return;
  const accItems = histContainer.querySelectorAll('div[style*="border:1px solid #ddd"]');
  accItems.forEach(item => {{
    const labelEl = item.querySelector('div[onclick]');
    if (!labelEl) return;
    const text = labelEl.textContent;
    // "2025年3月" のような表記から年月を取る
    const match = text.match(/(\\d{{4}})年(\\d{{1,2}})月/);
    if (!match) return;
    const ym = match[1] + '/' + String(match[2]).padStart(2,'0');
    item.style.display = (ym >= startKey && ym <= endKey) ? '' : 'none';
  }});
}}

function buildCompareTable(startKey, endKey) {{
  const startData = ALL_DATA[startKey];
  const endData = ALL_DATA[endKey];
  const container = document.getElementById('compareTableContainer');

  if (!startData || !endData) {{
    container.innerHTML = '<p style="color:#c0392b;font-size:13px">選択した期間のデータが存在しません（' + startKey + ' または ' + endKey + '）</p>';
    return;
  }}

  const startRows = startData.brand_rows;
  const endRows = endData.brand_rows;

  let html = '<table class="compare-table" style="border-collapse:collapse;width:100%">';
  html += '<thead><tr>';
  html += '<th>ブランド</th>';
  html += '<th>' + startKey + '（全拠点）</th>';
  html += '<th>' + endKey + '（全拠点）</th>';
  html += '<th>増減（全拠点）</th>';
  html += '<th>' + startKey + '（IR用）</th>';
  html += '<th>' + endKey + '（IR用）</th>';
  html += '<th>増減（IR用）</th>';
  html += '</tr></thead><tbody>';

  for (let i = 0; i < endRows.length; i++) {{
    const er = endRows[i];
    const sr = startRows[i] || {{label: er.label, all: 0, pr: 0}};
    const diffAll = er.all - sr.all;
    const diffPr  = er.pr  - sr.pr;
    const diffAllCls = diffAll > 0 ? 'pos' : diffAll < 0 ? 'neg' : '';
    const diffPrCls  = diffPr  > 0 ? 'pos' : diffPr  < 0 ? 'neg' : '';
    const diffAllStr = diffAll > 0 ? '+' + diffAll : String(diffAll);
    const diffPrStr  = diffPr  > 0 ? '+' + diffPr  : String(diffPr);
    html += '<tr>';
    html += '<td style="text-align:left">' + er.label + '</td>';
    html += '<td style="text-align:right">' + sr.all + '</td>';
    html += '<td style="text-align:right">' + er.all + '</td>';
    html += '<td style="text-align:right" class="' + diffAllCls + '">' + diffAllStr + '</td>';
    html += '<td style="text-align:right">' + sr.pr + '</td>';
    html += '<td style="text-align:right">' + er.pr + '</td>';
    html += '<td style="text-align:right" class="' + diffPrCls + '">' + diffPrStr + '</td>';
    html += '</tr>';
  }}

  // 合計行
  const sAll = startData.sum_all, eAll = endData.sum_all;
  const sPr  = startData.sum_pr,  ePr  = endData.sum_pr;
  const dAll = eAll - sAll, dPr = ePr - sPr;
  const dAllStr = dAll >= 0 ? '+' + dAll : String(dAll);
  const dPrStr  = dPr  >= 0 ? '+' + dPr  : String(dPr);
  html += '<tr style="background:#FFF2CC;font-weight:bold">';
  html += '<td>合計</td>';
  html += '<td style="text-align:right">' + sAll + '</td>';
  html += '<td style="text-align:right">' + eAll + '</td>';
  html += '<td style="text-align:right;' + (dAll>=0?'color:#27ae60':'color:#c0392b') + '">' + dAllStr + '</td>';
  html += '<td style="text-align:right">' + sPr + '</td>';
  html += '<td style="text-align:right">' + ePr + '</td>';
  html += '<td style="text-align:right;' + (dPr>=0?'color:#27ae60':'color:#c0392b') + '">' + dPrStr + '</td>';
  html += '</tr>';

  // IR広報用（OrangeTwist込み）
  const sIR = sPr + ORANGE_TWIST_COUNT, eIR = ePr + ORANGE_TWIST_COUNT;
  const dIR = eIR - sIR;
  const dIRStr = dIR >= 0 ? '+' + dIR : String(dIR);
  html += '<tr style="background:#FAD7A0;font-weight:bold">';
  html += '<td>IR・広報用（除：Holdingsへの収益貢献なし、含：OrangeTwist）</td>';
  html += '<td style="text-align:right">' + sIR + '</td>';
  html += '<td style="text-align:right">' + eIR + '</td>';
  html += '<td style="text-align:right;' + (dIR>=0?'color:#27ae60':'color:#c0392b') + '">' + dIRStr + '</td>';
  html += '<td colspan="3" style="text-align:center;color:#888">（参考）</td>';
  html += '</tr>';

  html += '</tbody></table>';
  container.innerHTML = html;
}}
</script>
</body>
</html>"""

    OUTPUT_HTML.write_text(html, encoding="utf-8")
    print(f"HTMLレポートを生成しました: {OUTPUT_HTML}")
    return True

if __name__ == "__main__":
    generate()
    print("完了！")
