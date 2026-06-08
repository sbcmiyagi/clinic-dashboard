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
from itertools import groupby

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

# 法人別集計から特定ブランドを除外する設定
HOUJIN_BRAND_EXCLUSIONS = {
    "医療法人社団 樹慶会":       {"brands": ["神奈川レディース"], "note": "※神奈川レディース・神奈川ウィメンズを除く"},
    "医療法人社団 リッツ美容外科": {"brands": ["リッツ美容外科"],  "note": "※リッツ美容外科を除く"},
}

HOUJIN_GROUPS = [
    ("①6医療法人合計", ["医療法人 湘美会","医療法人社団 孝和会","医療法人社団 菜寿会","医療法人社団 愛恵会","医療法人社団 樹慶会","医療法人社団 リッツ美容外科","一般社団法人MASA","健美会","法人無し（個人開設）","個人/その他"]),
    ("②3医療法人合計", ["医療法人社団 風林会","医療法人 きびたき会","医療法人社団 百花会"]),
    ("③医療法人社団十二会", ["医療法人社団十二会"]),
    ("④2医療法人合計", ["医療法人社団美咲会","一般社団法人美央斗会"]),
    ("⑤㈻SBC東京医療大学附属", ["㈻SBC東京医療大学附属","学校法人 SBC東京医療大学"]),
    ("⑥株式会社SBC湘南接骨院", ["株式会社SBC湘南接骨院"]),
    ("⑦株式会社ボディアーキ・ジャパン", ["株式会社ボディアーキ・ジャパン","SBCメディカルグループ株式会社","株式会社 MG"]),
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

def load_houjin_settings():
    """Excelの「法人設定」シートから法人の並び順を読み込む。シートがない場合はデフォルト値を返す。"""
    default = {
        "国内": [
            "医療法人 湘美会","医療法人社団 孝和会","医療法人社団 菜寿会","医療法人社団 愛恵会",
            "医療法人社団 樹慶会","医療法人社団 リッツ美容外科","一般社団法人MASA","健美会","法人無し（個人開設）",
            "医療法人社団 風林会","医療法人 きびたき会","医療法人社団 百花会",
            "医療法人社団十二会","医療法人社団美咲会","一般社団法人美央斗会",
            "㈻SBC東京医療大学附属","株式会社SBC湘南接骨院","株式会社ボディアーキ・ジャパン"
        ],
        "海外": ["Shoubikai Medical Vietnam Co., Ltd.","WWMG","DS","DSS","DSS(FC)","WWFC","RCC"]
    }
    try:
        path = get_path()
        raw = pd.read_excel(path, sheet_name="法人設定", header=None)
        # ヘッダー行を自動検出
        header_row = None
        for i, row in raw.iterrows():
            if any("法人名" in str(v) for v in row.values):
                header_row = i; break
        if header_row is None:
            return default
        df = pd.read_excel(path, sheet_name="法人設定", header=header_row)
        df = df.dropna(subset=["法人名"])
        df["法人名"] = df["法人名"].astype(str).str.strip()
        if "表示順" in df.columns:
            df = df.sort_values("表示順")
        result = {"国内": [], "海外": []}
        for _, row in df.iterrows():
            name   = str(row.get("法人名","") or "").strip()
            region = str(row.get("国内／海外","") or "").strip()
            if name and region in result:
                result[region].append(name)
        return result if (result["国内"] or result["海外"]) else default
    except Exception:
        return default

# 法人設定はgenerate()内で動的読み込みするためここでは初期化のみ
REGION_HOUJIN_ORDER = {"国内": [], "海外": []}

DOCTOR_FILE_PATH = Path(r"C:\Users\宮城杏奈\Box\総合企画部_特殊案件\その他\院情報一覧カウント\ドクター人事通達 自動集計.xlsx")
PAST_DIRECTOR_FILE_PATH = Path(r"C:\Users\宮城杏奈\Downloads\過去分院長変更履歴.xlsx")


def load_doctor_data():
    """ドクター人事通達ファイルを読み込む"""
    try:
        if not DOCTOR_FILE_PATH.exists():
            return pd.DataFrame()
        df = pd.read_excel(DOCTOR_FILE_PATH, sheet_name="全体", header=1)
        df["付日"] = pd.to_datetime(df["付日"], errors="coerce")
        df = df.dropna(subset=["付日"])
        df = df.sort_values("付日").reset_index(drop=True)
        return df
    except Exception:
        return pd.DataFrame()

def load_past_director_data():
    """過去分院長変更履歴から院IDごとの院長変更履歴を構築
    Returns: (result, current_info)
      result: {clinic_id: [(date, director_name), ...]}
      current_info: {clinic_id: (current_director, start_date_str)} D列の着任日
    """
    try:
        if not PAST_DIRECTOR_FILE_PATH.exists():
            return {}, {}

        # Read raw without header to get exact positions
        raw = pd.read_excel(PAST_DIRECTOR_FILE_PATH, header=None)
        # Row 1 (index 1) is the actual column header
        # Data starts from row 2 (index 2)

        result = {}
        current_info = {}  # {clinic_id: (director_name, "YYYY/MM")}

        for row_idx in range(2, len(raw)):
            row = raw.iloc[row_idx]

            clinic_id_val = row.iloc[0]  # TWE院ID = column A
            if pd.isna(clinic_id_val): continue
            try:
                clinic_id = int(float(str(clinic_id_val)))
            except: continue

            events = []  # list of (date, director_name)

            # Opening director (column AK = index 36)
            opening_dir = str(row.iloc[36] if len(row) > 36 else "").strip()
            if opening_dir and opening_dir not in ("-", "", "nan"):
                events.append((pd.Timestamp("2000-01-01"), opening_dir))

            # Changes 1-10: columns G(6), H(7), I(8) for 1回目, J(9),K(10),L(11) for 2回目, etc.
            for i in range(10):
                base_col = 6 + i * 3  # G=6 for 1st change
                if base_col + 2 >= len(row): break

                change_date_val = row.iloc[base_col]       # 交代日
                new_dir_val     = row.iloc[base_col + 2]   # 交代分院長 (who came in)

                change_date = to_ts(change_date_val)
                new_dir = str(new_dir_val or "").strip() if pd.notna(new_dir_val) else ""

                if change_date and new_dir and new_dir not in ("-", "", "nan"):
                    events.append((change_date, new_dir))

            # Current director with start date (columns C=2, D=3)
            current_dir = str(row.iloc[2] if len(row) > 2 else "").strip()
            current_start_val = row.iloc[3] if len(row) > 3 else None
            current_start = to_ts(current_start_val)
            if current_dir and current_dir not in ("-", "", "nan") and current_start:
                events.append((current_start, current_dir))
                # D列の着任日を current_info に保存（YYYY/MM形式）
                current_info[clinic_id] = (
                    current_dir,
                    f"{current_start.year}/{current_start.month:02d}"
                )

            # Sort and deduplicate
            if events:
                seen = {}
                for dt, name in events:
                    seen[dt] = name
                events = sorted(seen.items(), key=lambda x: x[0])
                result[clinic_id] = events

        return result, current_info
    except Exception:
        return {}, {}


def extract_doctor_name(raw_name):
    """氏名（よみ）列から漢字名を抽出（読み仮名を除く）"""
    if pd.isna(raw_name): return ""
    s = str(raw_name).strip()
    # Remove content in parentheses (reading)
    import re
    s = re.sub(r'[（(][^）)]*[）)]', '', s).strip()
    return s

def parse_director_from_detail(detail):
    """
    異動後/詳細から院長情報を抽出
    Returns (clinic_name, True) if 院長 found, else (None, False)
    """
    if pd.isna(detail): return None, False
    s = str(detail)
    # Split by 兼 to handle multiple roles
    parts = s.split("兼")
    for part in parts:
        part = part.strip()
        if "・" in part:
            idx = part.index("・")
            clinic = part[:idx].strip()
            role = part[idx+1:].strip()
            if "院長" in role and "副院長" not in role:
                return clinic, True
    return None, False

def build_director_pivot(doctor_df, clinic_df, past_data):
    """院長履歴のピボットデータを構築（過去データ+人事通達データを統合）"""
    today = date.today()

    # Determine start date: 2021/01
    months = []
    y, m = 2021, 1
    while (y, m) <= (today.year, today.month):
        months.append(f"{y}/{m:02d}")
        m += 1
        if m > 12: m = 1; y += 1

    # TWE表記 → 正式名称のマッピング（ブランド別候補対応）
    twe_to_fullname = {}   # {twe: 正式名称} 院ID昇順で先着優先（SBC湘南美容が最初に来るため正しいブランドを選択）
    twe_to_clinics  = {}   # {twe: [(正式名称, ブランド)]} ブランド優先検索用
    clinic_to_brand = {}   # {正式名称: ブランド} 逆引き用
    for _, row in clinic_df.iterrows():
        name  = str(row.get("正式名称", "") or "").strip()
        twe   = str(row.get("TWE表記",  "") or "").strip()
        brand = str(row.get("ブランド",  "") or "").strip()
        if name and twe and twe not in ("nan", ""):
            if twe not in twe_to_fullname:   # 最初のものを優先（院ID昇順→SBC湘南美容が先）
                twe_to_fullname[twe] = name
            twe_to_clinics.setdefault(twe, []).append((name, brand))
        if name and brand and brand not in ("nan", ""):
            clinic_to_brand[name] = brand

    # Build clinic ID to name mapping
    # 業態転換がある院は「転換前の院名」を優先（歴史データのマッピングに使う）
    id_to_name = {}
    id_to_old_name = {}  # 転換前院名（業態転換日あり）
    id_to_new_name = {}  # 転換後院名
    for _, row in clinic_df.iterrows():
        cid = row.get("院ID")
        name = str(row.get("正式名称", "") or "").strip()
        d_conv = to_ts(row.get("業態転換日"))
        tenkan_mae = str(row.get("転換前業態", "") or "").strip()
        if pd.notna(cid) and name:
            try:
                cid_int = int(float(str(cid)))
                if d_conv is not None:
                    # 業態転換日あり → 転換元の院
                    id_to_old_name[cid_int] = name
                elif tenkan_mae and tenkan_mae not in ("", "nan"):
                    # 転換前業態あり → 転換後の院
                    id_to_new_name[cid_int] = name
                # id_to_name: 転換元があればそれを優先
                if cid_int not in id_to_name or d_conv is not None:
                    id_to_name[cid_int] = name
            except: pass

    # Build past director lookup: {clinic_name: [(date, director), ...]}
    # 転換元院名にマッピングし、転換後も同じデータを引き継ぐ
    past_by_name = {}
    for cid, events in past_data.items():
        if cid in id_to_name:
            old_name = id_to_old_name.get(cid) or id_to_name[cid]
            past_by_name[old_name] = events
            # 転換後の院にも同じ履歴を引き継ぐ（転換日以降の分を継続表示）
            new_name = id_to_new_name.get(cid)
            if new_name and new_name not in past_by_name:
                past_by_name[new_name] = events

    # Build monthly events index from doctor_df (2025/09+)
    events_by_month = {}
    if not doctor_df.empty:
        for _, row in doctor_df.iterrows():
            date_val = row["付日"]
            if pd.isna(date_val): continue
            mk = f"{date_val.year}/{date_val.month:02d}"
            events_by_month.setdefault(mk, []).append(row)

    # Build monthly snapshots
    monthly_states = {}
    state = {}  # {clinic_name: doctor_name} for recent data
    doctor_clinic = {}  # {doctor: clinic} for recent data
    promotion_events  = {}  # {month_key: [{doctor, from, to, kubun}]}
    resignation_events = {}  # {month_key: [{doctor, clinic}]}

    for month_key in months:
        yr, mo = int(month_key[:4]), int(month_key[5:])
        month_end_dt = pd.Timestamp(yr, mo, calendar.monthrange(yr, mo)[1])

        # Start with past data snapshot for this month
        snapshot = {}
        for clinic_name, events in past_by_name.items():
            director = ""
            for dt, name in events:
                if dt <= month_end_dt:
                    director = name
                else:
                    break
            if director and director not in ("-", ""):
                snapshot[clinic_name] = director

        # From recent data (2025/09+) - process events and override past data
        if month_key >= "2025/09":
            if month_key in events_by_month:
                for row in events_by_month[month_key]:
                    kubun = str(row.get("区分", "") or "").strip()
                    detail = str(row.get("異動後/詳細", "") or "")
                    doctor = extract_doctor_name(row.get("氏名（よみ）", ""))

                    clinic, is_dir = parse_director_from_detail(detail)
                    # 略称（TWE表記）を正式名称に変換（ブランド別候補対応・「○○院」→「○○」でも試行）
                    if clinic:
                        _ck = clinic if clinic in twe_to_clinics else (clinic[:-1] if clinic.endswith('院') else clinic)
                        if _ck in twe_to_clinics:
                            clinic = twe_to_clinics[_ck][0][0]  # 院ID最小（SBC湘南美容）を使用

                    if kubun == "退職":
                        # 退職クリニックを特定して記録（玉突き起点として使用）
                        resign_clinic = None
                        resign_raw = str(row.get("異動前/現所属", "") or "").strip()
                        if resign_raw and resign_raw not in ("nan", ""):
                            rc_key = resign_raw if resign_raw in twe_to_clinics else \
                                     (resign_raw[:-1] if resign_raw.endswith('院') else resign_raw)
                            resign_clinic = twe_to_clinics[rc_key][0][0] if rc_key in twe_to_clinics else resign_raw
                        if not resign_clinic and doctor in doctor_clinic:
                            resign_clinic = doctor_clinic[doctor]
                        if not resign_clinic:
                            for cl, dr in state.items():
                                if dr == doctor:
                                    resign_clinic = cl
                                    break
                        if resign_clinic:
                            resignation_events.setdefault(month_key, []).append(
                                {"doctor": doctor, "clinic": resign_clinic})
                        # 既存処理
                        if doctor in doctor_clinic:
                            old_clinic = doctor_clinic.pop(doctor)
                            if state.get(old_clinic) == doctor:
                                state[old_clinic] = ""
                    elif is_dir and clinic:
                        # 昇格・新任（前月に院長職なし）の場合は昇格イベントとして記録
                        if doctor not in doctor_clinic:
                            from_place_raw = str(row.get("異動前/現所属", "") or "").strip()
                            if from_place_raw and from_place_raw not in ("nan", ""):
                                # 着任先のブランドを参照して異動元クリニックを正確に特定
                                to_brand = clinic_to_brand.get(clinic, "")
                                fp_key = from_place_raw
                                if fp_key not in twe_to_clinics and fp_key.endswith('院'):
                                    fp_key = fp_key[:-1]
                                if fp_key in twe_to_clinics:
                                    cands = twe_to_clinics[fp_key]
                                    if len(cands) == 1:
                                        from_place = cands[0][0]
                                    elif to_brand:
                                        # 着任先と同じブランドを優先
                                        same = [n for n, b in cands if b == to_brand]
                                        from_place = same[0] if same else cands[0][0]
                                    else:
                                        from_place = cands[0][0]
                                else:
                                    from_place = from_place_raw
                            else:
                                from_place = "（新任）"
                            promotion_events.setdefault(month_key, []).append({
                                "doctor": doctor, "from": from_place, "to": clinic,
                                "kubun": kubun
                            })
                        if doctor in doctor_clinic:
                            old_clinic = doctor_clinic[doctor]
                            if state.get(old_clinic) == doctor:
                                state[old_clinic] = ""
                        prev = state.get(clinic, "")
                        if prev and prev != doctor and doctor_clinic.get(prev) == clinic:
                            del doctor_clinic[prev]
                        state[clinic] = doctor
                        doctor_clinic[doctor] = clinic
                    elif kubun in ("異動", "昇格") and not is_dir:
                        if doctor in doctor_clinic:
                            old_clinic = doctor_clinic.pop(doctor)
                            if state.get(old_clinic) == doctor:
                                state[old_clinic] = ""

            # Override snapshot with recent state
            for clinic_name, director in state.items():
                if director:
                    snapshot[clinic_name] = director

        monthly_states[month_key] = snapshot

    # ── 院長履歴スナップショット比較による退職・昇格の自動補完 ──
    # 人事通達データのない月（2025/08以前）や未記入の退職を補う
    for i, month_key in enumerate(months):
        if i == 0: continue
        prev_month = months[i - 1]
        prev_state = monthly_states.get(prev_month, {})
        cur_state  = monthly_states.get(month_key,  {})

        # 院長→クリニック の逆引きマップ
        prev_dr = {d: c for c, d in prev_state.items() if d and d not in ('-', '', 'nan')}
        cur_dr  = {d: c for c, d in cur_state.items()  if d and d not in ('-', '', 'nan')}

        # 前月は院長だったが今月はどこにも院長として現れない人 → 退職と判定
        for doctor in set(prev_dr) - set(cur_dr):
            old_clinic = prev_dr[doctor]
            new_doctor = cur_state.get(old_clinic, "")
            # そのクリニックに新しい院長が就任している場合のみ記録（空席のままは除外）
            if new_doctor and new_doctor != doctor:
                if not any(e["doctor"] == doctor
                           for e in resignation_events.get(month_key, [])):
                    resignation_events.setdefault(month_key, []).append(
                        {"doctor": doctor, "clinic": old_clinic})

        # 今月初めて院長になった人（前月は院長ではなかった）→ 昇格と判定
        for doctor in set(cur_dr) - set(prev_dr):
            new_clinic = cur_dr[doctor]
            if not any(e["doctor"] == doctor
                       for e in promotion_events.get(month_key, [])):
                promotion_events.setdefault(month_key, []).append(
                    {"doctor": doctor, "from": "（昇格）", "to": new_clinic, "kubun": "昇格"})

    return monthly_states, months, promotion_events, resignation_events

def build_director_html(doctor_df, clinic_df, brand_cols):
    """院長履歴ピボットテーブルHTMLを生成"""
    past_data, past_current_info = load_past_director_data()

    monthly_states, months, _, __ = build_director_pivot(doctor_df, clinic_df, past_data)
    if not months:
        return '<p style="color:#999">データがありません</p>'

    # Get active clinics (use クリニック一覧)
    target_brands = [b for b, _ in brand_cols]
    active_clinics = []
    seen = set()
    for _, row in clinic_df.iterrows():
        brand = get_brand(row)
        if brand not in target_brands: continue
        # Only include clinics that appear in director data
        name = str(row.get("正式名称","") or "").strip()
        if name and name not in seen:
            seen.add(name)
            active_clinics.append(name)

    # Find clinics that actually have director data
    clinics_with_data = set()
    for m_state in monthly_states.values():
        for clinic, doctor in m_state.items():
            if doctor:
                clinics_with_data.add(clinic)

    if not clinics_with_data:
        return '<p style="color:#999">院長データが見つかりません。ファイルの内容を確認してください。</p>'

    # 院IDをキーにした院名リストを作成（院ID順でソート）
    clinic_id_map   = {}  # {正式名称: 院ID}
    clinic_abbr_map = {}  # {正式名称: 略式院名}
    clinic_active_map = {} # {正式名称: True/False（現存中か）}
    for _, row in clinic_df.iterrows():
        name = str(row.get("正式名称","") or "").strip()
        cid  = row.get("院ID")
        abbr = str(row.get("略式院名","") or "").strip()
        flag = str(row.get("開院フラグ","") or "").strip()
        if name and pd.notna(cid):
            try:
                clinic_id_map[name] = int(float(str(cid)))
            except: pass
        if name and abbr and abbr not in ("nan",""):
            clinic_abbr_map[name] = abbr
        if name:
            clinic_active_map[name] = (flag == "開院")

    # 業態転換ペアを特定: {旧院名: (新院名, 業態転換年月)}
    conversion_map = {}    # 旧院名 → (新院名, 転換年月)
    converted_new = set()  # 転換後の院名（新）
    for _, row in clinic_df.iterrows():
        d_conv = to_ts(row.get("業態転換日"))
        new_name = str(row.get("転換後院名","") or "").strip()
        old_name = str(row.get("正式名称","") or "").strip()
        if d_conv and new_name and old_name:
            conv_month = f"{d_conv.year}/{d_conv.month:02d}"
            old_id = clinic_id_map.get(old_name)
            new_id = clinic_id_map.get(new_name)
            same_id = (old_id is not None and new_id is not None and old_id == new_id)
            conversion_map[old_name] = (new_name, conv_month, same_id, old_id, new_id)
            converted_new.add(new_name)

    # 院IDがある院は院ID順、ない院は末尾に名前順
    clinics_sorted = sorted(
        clinics_with_data,
        key=lambda c: (clinic_id_map.get(c, 999999), c)
    )
    # 転換後の院は独立表示しない（旧院の直後に表示するため除外）
    clinics_sorted_main = [c for c in clinics_sorted if c not in converted_new]

    # Build HTML table
    th_style = f'padding:6px 10px;border:1px solid #555;background:{C_HEADER};color:white;white-space:nowrap;font-size:12px'

    # 左4列を固定（各列のleft位置を積算で指定）
    W_ID   = 44   # 院ID列幅
    W_NAME = 184  # 院名列幅
    W_DIR  = 104  # 現院長列幅
    W_DATE = 74   # 就任時期列幅
    sticky = 'position:sticky;top:0;z-index:5;background:{C_HEADER}'
    W_FLAG = 50   # 状態列幅
    headers  = f'<th style="{th_style};position:sticky;top:0;left:0px;z-index:5;min-width:{W_ID}px">院ID</th>'
    headers += f'<th style="{th_style};position:sticky;top:0;left:{W_ID}px;z-index:5;min-width:{W_FLAG}px">状態</th>'
    headers += f'<th style="{th_style};position:sticky;top:0;left:{W_ID+W_FLAG}px;z-index:5;min-width:{W_NAME}px">院名</th>'
    headers += f'<th style="{th_style};position:sticky;top:0;left:{W_ID+W_FLAG+W_NAME}px;z-index:5;min-width:{W_DIR}px">現院長</th>'
    headers += f'<th style="{th_style};position:sticky;top:0;left:{W_ID+W_FLAG+W_NAME+W_DIR}px;z-index:5;min-width:{W_DATE}px">就任時期</th>'
    for mk in months:
        headers += f'<th style="{th_style};position:sticky;top:0;z-index:4;min-width:80px">{mk}</th>'

    # 院ID → clinic_name の逆引きマップ（past_current_infoをclinic名で使うため）
    id_to_clinic_name = {v: k for k, v in clinic_id_map.items()}
    # clinic名 → D列の着任日 マップ
    clinic_exact_start = {}
    for cid, (dir_name, start_str) in past_current_info.items():
        clinic_name = id_to_clinic_name.get(cid, "")
        if clinic_name:
            clinic_exact_start[clinic_name] = (dir_name, start_str)

    def get_current_director_info(clinic, limit_before=None, limit_from=None):
        """現院長と就任時期を返す（D列の着任日を優先使用）"""
        current = ""
        start_month = ""
        prev = ""
        for mk in months:
            if limit_before and mk >= limit_before: continue
            if limit_from and mk < limit_from: continue
            doc = monthly_states.get(mk, {}).get(clinic, "")
            if doc and doc != prev:
                current = doc
                start_month = mk
            elif not doc:
                pass
            prev = doc if doc else prev

        # D列の着任日が存在し院長名が一致する場合はそちらを優先
        if current and clinic in clinic_exact_start:
            exact_dir, exact_start = clinic_exact_start[clinic]
            if exact_dir == current:
                start_month = exact_start

        return current, start_month

    def make_clinic_row(clinic, limit_before=None, limit_from=None, row_bg="white"):
        """1院分の行HTMLを生成。limit_before=この月より後は空白、limit_from=この月より前は空白"""
        prev_doctor = ""
        cid_val = clinic_id_map.get(clinic, "")
        # ① 現院長・就任時期
        cur_dir, cur_start = get_current_director_info(clinic, limit_before, limit_from)
        # 状態フラグ（開院中/閉院・転換）
        is_active = clinic_active_map.get(clinic, True)
        status_txt = "開院中" if is_active else "閉院/転換"
        status_bg  = "#D5F5E3" if is_active else "#FADBD8"
        status_col = "#1E8449" if is_active else "#922B21"
        # 左列をsticky固定（left位置を積算）
        cells  = f'<td style="padding:5px 8px;border:1px solid #ddd;position:sticky;left:0px;background:{row_bg};font-size:12px;text-align:right;color:#666;z-index:1">{cid_val}</td>'
        cells += f'<td style="padding:3px 6px;border:1px solid #ddd;position:sticky;left:{W_ID}px;background:{status_bg};font-size:10px;text-align:center;color:{status_col};font-weight:bold;z-index:1">{status_txt}</td>'
        cells += f'<td style="padding:5px 10px;border:1px solid #ddd;position:sticky;left:{W_ID+W_FLAG}px;background:{row_bg};font-size:12px;white-space:nowrap;color:#333;z-index:1">{clinic}</td>'
        cells += f'<td style="padding:5px 10px;border:1px solid #ddd;position:sticky;left:{W_ID+W_FLAG+W_NAME}px;background:{row_bg};font-size:12px;font-weight:bold;color:#2C3E50;z-index:1">{cur_dir}</td>'
        cells += f'<td style="padding:5px 8px;border:1px solid #ddd;position:sticky;left:{W_ID+W_FLAG+W_NAME+W_DIR}px;background:{row_bg};font-size:11px;color:#555;text-align:center;z-index:1">{cur_start}</td>'
        for mk in months:
            # 業態転換による表示範囲制限
            if limit_before and mk >= limit_before:
                cells += f'<td style="padding:5px 8px;border:1px solid #ddd;background:#f0f0f0;font-size:12px"></td>'
                continue
            if limit_from and mk < limit_from:
                cells += f'<td style="padding:5px 8px;border:1px solid #ddd;background:#f0f0f0;font-size:12px"></td>'
                continue
            doctor = monthly_states.get(mk, {}).get(clinic, "")
            changed = doctor and doctor != prev_doctor and prev_doctor != ""
            if not doctor:
                bg = "#f8f9fa"; color = "#ccc"; txt = "―"
            elif changed:
                bg = "#FFF9C4"; color = "#333"; txt = doctor
            else:
                bg = row_bg; color = "#333"; txt = doctor
            cells += f'<td style="padding:5px 8px;border:1px solid #ddd;background:{bg};color:{color};font-size:12px;white-space:nowrap;text-align:center">{txt}</td>'
            if doctor: prev_doctor = doctor
        return f"<tr>{cells}</tr>"

    # ── ビュー1: 院ID順 ──
    rows_id = ""
    for clinic in clinics_sorted_main:
        if clinic in conversion_map:
            new_name, conv_month, same_id, old_id, new_id = conversion_map[clinic]
            rows_id += make_clinic_row(clinic, limit_before=conv_month, row_bg="#D7BDE2")
            if new_name in clinics_with_data:
                rows_id += make_clinic_row(new_name, limit_from=conv_month, row_bg="#F3E5F5")
        else:
            rows_id += make_clinic_row(clinic)

    # ── ビュー2: 業態転換グループ順 ──
    rows_group = ""
    shown_in_group = set()
    # 転換ペアをまとめて表示（旧院のIDでソート）
    conv_pairs = sorted(
        [(old, info) for old, info in conversion_map.items() if old in clinics_with_data or info[0] in clinics_with_data],
        key=lambda x: (clinic_id_map.get(x[0], 999999), x[0])
    )
    for old_clinic, (new_name, conv_month, same_id, old_id, new_id) in conv_pairs:
        # グループヘッダー
        id_badge = (
            f'🔵 同ID（{old_id}）' if same_id
            else f'🔴 ID変更（{old_id} → {new_id}）'
        )
        rows_group += (
            f'<tr><td colspan="{len(months)+4}" style="padding:4px 12px;'
            f'background:#4A235A;color:white;font-size:11px;font-weight:bold">'
            f'業態転換 {conv_month}　{id_badge}</td></tr>'
        )
        rows_group += make_clinic_row(old_clinic, limit_before=conv_month, row_bg="#D7BDE2")
        if new_name in clinics_with_data:
            rows_group += make_clinic_row(new_name, limit_from=conv_month, row_bg="#F3E5F5")
        shown_in_group.add(old_clinic)
        shown_in_group.add(new_name)

    # 転換に関係ない院を院ID順で追加
    for clinic in clinics_sorted:
        if clinic not in shown_in_group and clinic in clinics_with_data:
            rows_group += make_clinic_row(clinic)

    legend = '''<div style="display:flex;gap:12px;margin-bottom:8px;font-size:12px;align-items:center;flex-wrap:wrap">
      <span>凡例：</span>
      <span style="background:#FFF9C4;padding:2px 8px;border:1px solid #ddd">院長交代</span>
      <span style="background:#D7BDE2;padding:2px 8px;border:1px solid #ddd;color:#6c3483">業態転換前</span>
      <span style="background:#F3E5F5;padding:2px 8px;border:1px solid #ddd;color:#6c3483">業態転換後</span>
      <span style="background:#f0f0f0;padding:2px 8px;border:1px solid #ddd;color:#999">対象外期間</span>
      <span style="color:#ccc;padding:2px 8px;border:1px solid #ddd">―　記録なし</span>
      &nbsp;│&nbsp;
      <span>🔵 同ID引き継ぎ</span>
      <span>🔴 ID変更</span>
    </div>'''

    toggle_btn = '''<div style="margin-bottom:8px;display:flex;gap:8px">
      <button id="btnViewId" onclick="switchView('id')"
        style="padding:5px 14px;border-radius:16px;border:none;cursor:pointer;font-size:13px;background:#2C3E50;color:white;font-weight:bold">
        院ID順
      </button>
      <button id="btnViewGroup" onclick="switchView('group')"
        style="padding:5px 14px;border-radius:16px;border:none;cursor:pointer;font-size:13px;background:#ddd;color:#333">
        業態転換グループ順
      </button>
    </div>
    <script>
    function switchView(mode) {
      document.getElementById('dirViewId').style.display    = mode==='id'    ? '' : 'none';
      document.getElementById('dirViewGroup').style.display = mode==='group' ? '' : 'none';
      document.getElementById('btnViewId').style.background    = mode==='id'    ? '#2C3E50' : '#ddd';
      document.getElementById('btnViewId').style.color         = mode==='id'    ? 'white'   : '#333';
      document.getElementById('btnViewGroup').style.background = mode==='group' ? '#2C3E50' : '#ddd';
      document.getElementById('btnViewGroup').style.color      = mode==='group' ? 'white'   : '#333';
    }
    </script>'''

    table = f'''
    {legend}
    {toggle_btn}
    <div id="dirViewId" style="overflow-x:auto;overflow-y:auto;max-height:70vh">
    <table style="border-collapse:collapse;font-size:12px">
    <thead><tr>{headers}</tr></thead>
    <tbody>{rows_id}</tbody>
    </table>
    </div>
    <div id="dirViewGroup" style="display:none;overflow-x:auto;overflow-y:auto;max-height:70vh">
    <table style="border-collapse:collapse;font-size:12px">
    <thead><tr>{headers}</tr></thead>
    <tbody>{rows_group}</tbody>
    </table>
    </div>'''

    return table


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

def load_orangetwist_settings():
    """ブランド設定シートからOrangeTwistの院数・開始年月を読み込む"""
    default = {"count": 24, "start": "2025/12"}
    try:
        path = get_path()
        raw = pd.read_excel(path, sheet_name="ブランド設定", header=None)
        for i, row in raw.iterrows():
            for j, val in enumerate(row):
                if str(val).strip() == "OrangeTwist":
                    cnt_val   = raw.iloc[i, j+1] if j+1 < raw.shape[1] else 24
                    start_val = raw.iloc[i, j+2] if j+2 < raw.shape[1] else "2025/12"
                    cnt = int(cnt_val) if pd.notna(cnt_val) else 24
                    if isinstance(start_val, pd.Timestamp):
                        start = start_val.strftime("%Y/%m")
                    else:
                        start = str(start_val).strip() if pd.notna(start_val) else "2025/12"
                    return {"count": cnt, "start": start}
        return default
    except Exception:
        return default

def get_orangetwist_count(month_key, ot_settings):
    """指定年月のOrangeTwist院数を返す（開始年月より前は0）"""
    return ot_settings["count"] if month_key >= ot_settings["start"] else 0


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
        if check_active(row, me):
            if houjin in HOUJIN_ORDER:
                r3.setdefault(houjin,{"all":0})
                if houjin not in EXCLUDE_HOUJIN: r3[houjin]["all"]+=1
    return r1, r2, r3

def df_to_html_table(df, highlight_last=True, orange_last=False, right_align_nums=False):
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
        if right_align_nums:
            cells = "".join(
                f'<td style="padding:6px 10px;border:1px solid #ddd;text-align:{"right" if ci > 0 else "left"}">{v}</td>'
                for ci, v in enumerate(row)
            )
        else:
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

        if check_active(row, me):
            if houjin in HOUJIN_ORDER:
                r3.setdefault(houjin, {"all": 0})
                if houjin not in EXCLUDE_HOUJIN:
                    r3[houjin]["all"] += 1

    return r1, r2, r3, houjin_brand


def build_all_monthly_data(df, target_brands, exclude_pr, brand_cols=None, existing_flags=None, ot_settings=None):
    """
    2021/01 から今月まで各月の集計データを返す
    brand_cols: ブランド設定シートから読んだ[(brand, gyoutai)] のリスト
    existing_flags: 各ブランドが既存G合計に含まれるかのboolリスト
    Returns dict keyed by "YYYY/MM"
    """
    # ブランド設定が未指定の場合はJIKEI_BRAND_COLSをフォールバックとして使用
    if not brand_cols:
        brand_cols = JIKEI_BRAND_COLS
        existing_flags = [True] * EXISTING_GROUP_END + [False] * (len(JIKEI_BRAND_COLS) - EXISTING_GROUP_END)
    if not existing_flags:
        existing_flags = [False] * len(brand_cols)

    today = date.today()
    sy, sm = 2021, 1
    all_data = {}

    while (sy, sm) <= (today.year, today.month):
        ym_key = f"{sy}/{sm:02d}"
        me = month_end_ts(sy, sm)
        ms_ts = month_start_ts(sy, sm)

        r1, r2, r3, houjin_brand = aggregate_with_houjin(df, target_brands, exclude_pr, me, ms_ts)

        # brand_rows（ブランド設定の順序で構築）
        brand_rows = []
        for (brand, gyoutai) in brand_cols:
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

        # 既存G合計：ブランド設定の「既存グループ合計＝○」のブランドのみ合算
        existing_sum = sum(
            brand_rows[i]["all"]
            for i in range(len(brand_rows))
            if i < len(existing_flags) and existing_flags[i]
        )

        # IR広報用（OrangeTwistは開始年月以降のみ加算）
        ot_count = get_orangetwist_count(ym_key, ot_settings) if ot_settings else ORANGE_TWIST_COUNT
        ir_sum = sum_pr + ot_count

        # jikei_row: brand counts + summary columns
        jikei_row = {}
        for r in brand_rows:
            jikei_row[r["label"]] = r["all"]
        jikei_row["既存G合計"] = existing_sum
        jikei_row[LABEL_ALL] = sum_all
        jikei_row["IR・広報用（除：Holdingsへの収益貢献なし）"] = sum_pr   # 広報IR用の除外前ベース
        jikei_row["OrangeTwist"] = ot_count
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
        # houjin group × brand × gyoutai counts
        for gname, members in HOUJIN_GROUPS:
            for _, row in df.iterrows():
                if check_active(row, me):
                    houjin = str(row.get("法人名", "") or "").strip()
                    if houjin in members:
                        brand = get_brand(row)
                        gyoutai = str(row.get("業態", "") or "").strip()
                        key = f"{gname}|{brand}|{gyoutai}"
                        jikei_row[key] = jikei_row.get(key, 0) + 1

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


def build_timeseries_html(all_data, brand_cols=None):
    """時系列推移テーブルHTMLを生成（法人グループ×ブランド×業態の3段階層ヘッダー付き）"""
    if not all_data:
        return "<p>データなし</p>"
    if not brand_cols:
        brand_cols = JIKEI_BRAND_COLS

    months = sorted(all_data.keys())

    # 各法人グループで実際に出現する(brand, gyoutai)の組み合わせを収集
    group_cols = {}
    for gname, members in HOUJIN_GROUPS:
        combos_seen = set()
        for month_data in all_data.values():
            jikei = month_data.get("jikei_row", {})
            for key, val in jikei.items():
                if key.startswith(f"{gname}|") and val and val > 0:
                    parts = key.split("|", 2)
                    if len(parts) == 3:
                        combos_seen.add((parts[1], parts[2]))
        # JIKEI_BRAND_COLS の順序で並べ直す
        jikei_order = []
        for brand, gyoutai in JIKEI_BRAND_COLS:
            g = gyoutai or ""
            if (brand, g) in combos_seen:
                jikei_order.append((brand, g))
        for combo in sorted(combos_seen):
            if combo not in jikei_order:
                jikei_order.append(combo)
        group_cols[gname] = jikei_order

    # ── ヘッダー行1：法人グループ名（構成法人内訳）+ 集計列 ──
    header1 = '<tr style="background:#2C3E50;color:white;font-size:11px;position:sticky;top:0;z-index:3">'
    header1 += '<th rowspan="3" style="position:sticky;left:0;z-index:4;background:#2C3E50;min-width:70px;padding:4px 6px">年月</th>'
    # 法人グループ列（ブランド削除）
    for gname, members in HOUJIN_GROUPS:
        combos = group_cols.get(gname, [])
        if not combos:
            continue
        # ⑦グループは「株式会社ボディアーキ・ジャパン」のみ表示
        if "ボディアーキ" in gname:
            member_label = "株式会社ボディアーキ・ジャパン"
        else:
            member_label = " / ".join(members)
        header1 += f'<th colspan="{len(combos)}" style="padding:4px 6px;border:1px solid #555;text-align:center;font-size:10px">{member_label}</th>'
    # 集計列（末尾）
    header1 += '<th rowspan="3" style="padding:4px 6px;border:1px solid #555;background:#1a5276;min-width:60px">既存G合計</th>'
    header1 += '<th rowspan="3" style="padding:4px 6px;border:1px solid #555;background:#1a5276;min-width:60px">全拠点</th>'
    header1 += '<th rowspan="3" style="padding:4px 6px;border:1px solid #555;background:#8e44ad;color:white;min-width:70px;white-space:normal;text-align:center">IR・広報用（除：Holdingsへの収益貢献なし）</th>'
    header1 += '<th rowspan="3" style="padding:4px 6px;border:1px solid #555;background:#d35400;min-width:60px">OrangeTwist</th>'
    header1 += f'<th rowspan="3" style="padding:4px 6px;border:1px solid #555;background:#d35400;min-width:70px;white-space:normal;text-align:center">{LABEL_IR}</th>'
    header1 += '</tr>'

    # ── ヘッダー行2：ブランド名 ──
    header2 = '<tr style="background:#2C3E50;color:white;font-size:11px;position:sticky;top:24px;z-index:3">'
    for gname, members in HOUJIN_GROUPS:
        combos = group_cols.get(gname, [])
        if not combos:
            continue
        grouped = []
        for brand, grp in groupby(combos, key=lambda x: x[0]):
            grouped.append((brand, len(list(grp))))
        for brand, cnt in grouped:
            header2 += f'<th colspan="{cnt}" style="padding:4px 6px;border:1px solid #555;text-align:center">{brand}</th>'
    header2 += '</tr>'

    # ── ヘッダー行3：業態名 ──
    header3 = '<tr style="background:#2C3E50;color:white;font-size:11px;position:sticky;top:48px;z-index:3">'
    for gname, members in HOUJIN_GROUPS:
        combos = group_cols.get(gname, [])
        for brand, gyoutai in combos:
            header3 += f'<th style="padding:4px 6px;border:1px solid #555;min-width:50px">{gyoutai if gyoutai else "―"}</th>'
    header3 += '</tr>'

    # ── データ行 ──
    rows_html = ""
    for month in months:
        month_data = all_data[month]
        jikei = month_data.get("jikei_row", {})
        is_dec = month.endswith("/12")
        bg = "#EBF5FB" if is_dec else "white"
        fw = "bold" if is_dec else "normal"

        row = f'<tr style="background:{bg};font-size:12px">'
        row += f'<td style="position:sticky;left:0;background:{bg};padding:4px 6px;border:1px solid #ddd;font-weight:{fw}">{month}</td>'

        # 法人グループ×ブランド×業態
        for gname, members in HOUJIN_GROUPS:
            combos = group_cols.get(gname, [])
            for brand, gyoutai in combos:
                key = f"{gname}|{brand}|{gyoutai}"
                val = jikei.get(key, 0)
                row += f'<td style="padding:4px 6px;border:1px solid #ddd;text-align:right">{val if val else ""}</td>'

        # 集計列（末尾）
        existing  = jikei.get("既存G合計", 0)
        all_total = jikei.get("全拠点", 0)
        excl_base = jikei.get("IR・広報用（除：Holdingsへの収益貢献なし）", 0)
        ot = jikei.get("OrangeTwist", ORANGE_TWIST_COUNT)
        ir = jikei.get(LABEL_IR, 0)
        row += f'<td style="padding:4px 6px;border:1px solid #ddd;text-align:right;background:#EBF5FB;color:#1a5276;font-weight:bold">{existing}</td>'
        row += f'<td style="padding:4px 6px;border:1px solid #ddd;text-align:right;background:#EBF5FB;color:#1a5276;font-weight:bold">{all_total}</td>'
        row += f'<td style="padding:4px 6px;border:1px solid #ddd;text-align:right;background:#E8DAEF;color:#6c3483;font-weight:bold">{excl_base}</td>'
        row += f'<td style="padding:4px 6px;border:1px solid #ddd;text-align:right;background:#FEF0E3;color:#d35400">{ot}</td>'
        row += f'<td style="padding:4px 6px;border:1px solid #ddd;text-align:right;background:#FEF0E3;color:#d35400;font-weight:bold">{ir}</td>'

        row += '</tr>'
        rows_html += row

    table = f'''
    <div style="overflow-x:auto;overflow-y:auto;max-height:70vh">
    <table style="border-collapse:collapse;font-size:12px;white-space:nowrap">
    <thead>{header1}{header2}{header3}</thead>
    <tbody>{rows_html}</tbody>
    </table>
    </div>'''
    return table


def build_doctor_movement_data(monthly_states, months):
    """先生ごとの異動データを構築"""
    # {doctor: {month: clinic}}
    doctor_monthly = {}
    for mk in months:
        state = monthly_states.get(mk, {})
        for clinic, doctor in state.items():
            if doctor and doctor not in ("-", ""):
                doctor_monthly.setdefault(doctor, {})[mk] = clinic

    # 連続する同一院をスティントにまとめる
    doctor_stints = {}
    all_clinics = set()
    for doctor, mc in doctor_monthly.items():
        sorted_months = sorted(mc.keys())
        stints = []
        if sorted_months:
            cur_clinic = mc[sorted_months[0]]
            cur_start  = sorted_months[0]
            cur_end    = sorted_months[0]
            for mk in sorted_months[1:]:
                if mc[mk] == cur_clinic:
                    cur_end = mk
                else:
                    stints.append({"clinic": cur_clinic, "start": cur_start, "end": cur_end})
                    all_clinics.add(cur_clinic)
                    cur_clinic = mc[mk]
                    cur_start  = mk
                    cur_end    = mk
            stints.append({"clinic": cur_clinic, "start": cur_start, "end": cur_end})
            all_clinics.add(cur_clinic)
        doctor_stints[doctor] = stints

    return doctor_stints, sorted(all_clinics), months


def generate():
    global REGION_HOUJIN_ORDER
    print("データを読み込み中...")
    df = load_data()
    # 法人設定シートから並び順を読み込む
    REGION_HOUJIN_ORDER = load_houjin_settings()
    ot_settings = load_orangetwist_settings()
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

    # ── 法人別表（特定法人から特定ブランドを除外） ──
    # 除外対象の院数を先に計算
    excl_counts = {}
    for houjin_name, excl_info in HOUJIN_BRAND_EXCLUSIONS.items():
        cnt = 0
        for _, row in df.iterrows():
            brand = get_brand(row)
            houjin = str(row.get("法人名","") or "").strip()
            if houjin == houjin_name and brand in excl_info["brands"]:
                if check_active(row, me):
                    cnt += 1
        excl_counts[houjin_name] = cnt

    # 法人別表に表示する法人（この6法人のみ）
    HOUJIN_DISPLAY = [
        "医療法人 湘美会",
        "医療法人社団 孝和会",
        "医療法人社団 菜寿会",
        "医療法人社団 愛恵会",
        "医療法人社団 樹慶会",
        "医療法人社団 リッツ美容外科",
    ]
    houjin_rows=[]; total_h=0
    for h in HOUJIN_DISPLAY:
        cnt = r3.get(h,{}).get("all",0)
        excl = excl_counts.get(h, 0)
        adj_cnt = cnt - excl
        note = HOUJIN_BRAND_EXCLUSIONS[h]["note"] if h in HOUJIN_BRAND_EXCLUSIONS else ""
        row_label = f"{h}　{note}" if note else h
        houjin_rows.append({"法人名": row_label, "全拠点": adj_cnt})
        total_h += adj_cnt
    houjin_rows.append({"法人名":"合計","全拠点":total_h})
    houjin_df=pd.DataFrame(houjin_rows)

    # ── 時系列データ構築 ──
    print("時系列推移データを集計中（2021/01〜現在）...")
    # ブランド設定シートの並び順・既存G合計フラグを時系列にも反映
    all_monthly_data = build_all_monthly_data(
        df, [b for b,_ in brand_cols], exclude_pr,
        brand_cols=brand_cols, existing_flags=existing_flags,
        ot_settings=ot_settings
    )
    timeseries_html = build_timeseries_html(all_monthly_data, brand_cols=brand_cols)

    # JSON埋め込み用（シングルクォートをエスケープ）
    json_str = json.dumps(all_monthly_data, ensure_ascii=False).replace("'", "\\'")

    # クリニック一覧をJSON用に変換（全院）
    clinic_records = []
    for _, row in df.iterrows():
        brand = get_brand(row)
        if not brand or brand not in [b for b,_ in brand_cols]:
            continue
        def fmt_date(v):
            ts = to_ts(v)
            return ts.strftime("%Y-%m-%d") if ts else None

        ma = to_ts(row.get("MA日"))
        op = to_ts(row.get("開院日"))
        base = ma if ma is not None else op

        clinic_records.append({
            "id": str(row.get("院ID","") or ""),
            "name": str(row.get("正式名称","") or ""),
            "brand": brand,
            "gyoutai": str(row.get("業態","") or ""),
            "houjin": str(row.get("法人名","") or ""),
            "region": get_region(row),
            "open_date": fmt_date(base),
            "conv_date": fmt_date(row.get("業態転換日")),
            "close_date": fmt_date(row.get("閉院日")),
        })
    clinic_json = json.dumps(clinic_records, ensure_ascii=False)
    clinic_json_escaped = clinic_json.replace("'", "\\'")

    # ブランド設定の順序でブランド一覧を作成（指示⑥）
    brand_cols_order = [b for b, _ in brand_cols]
    seen_brands = set()
    unique_brands = []
    for b in brand_cols_order:
        if b not in seen_brands and any(r["brand"] == b for r in clinic_records):
            unique_brands.append(b)
            seen_brands.add(b)
    # 設定にないブランドは末尾へ
    for r in clinic_records:
        if r["brand"] not in seen_brands:
            unique_brands.append(r["brand"])
            seen_brands.add(r["brand"])
    brands_json = json.dumps(unique_brands, ensure_ascii=False)

    # 法人設定の順序をJavaScript用にJSON化（国内→海外の順）
    houjin_order_list = REGION_HOUJIN_ORDER.get("国内", []) + REGION_HOUJIN_ORDER.get("海外", [])
    houjin_order_json = json.dumps(houjin_order_list, ensure_ascii=False)

    # ── ブランド別棒グラフ ──
    plot_df = brand_df[brand_df["ブランド"].isin([f"{b}({g})" if g else b for b,g in brand_cols])].copy()
    fig2=px.bar(plot_df,x="全拠点",y="ブランド",orientation="h",height=550,color_discrete_sequence=[C_BLUE])
    fig2.update_layout(yaxis=dict(autorange="reversed"),margin=dict(t=30))
    chart2_html = fig2.to_html(full_html=False, include_plotlyjs="cdn")

    # ── 院長履歴 ──
    doctor_df = load_doctor_data()
    director_html = build_director_html(doctor_df, df, brand_cols)

    # ── 先生の異動データ ──
    past_data_mv, _ = load_past_director_data()
    monthly_states_mv, months_mv, promotion_events, resignation_events = build_director_pivot(doctor_df, df, past_data_mv)
    doctor_stints, all_clinics_list, _ = build_doctor_movement_data(monthly_states_mv, months_mv)
    doctor_stints_json     = json.dumps(doctor_stints, ensure_ascii=False).replace("'", "\\'")
    all_doctors_json       = json.dumps(sorted(doctor_stints.keys()), ensure_ascii=False)
    all_clinics_json       = json.dumps(all_clinics_list, ensure_ascii=False)
    months_mv_json         = json.dumps(months_mv, ensure_ascii=False)
    promotion_events_json   = json.dumps(promotion_events,  ensure_ascii=False).replace("'", "\\'")
    resignation_events_json = json.dumps(resignation_events, ensure_ascii=False).replace("'", "\\'")

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
  .tabs{{display:flex;background:#f0f2f5;padding:10px 20px;flex-wrap:wrap;gap:8px;border-bottom:1px solid #ddd}}
  .tab{{padding:7px 16px;cursor:pointer;border-radius:20px;font-size:13px;font-weight:bold;
        color:white;opacity:0.6;transition:opacity 0.2s,box-shadow 0.2s;white-space:nowrap}}
  .tab:hover{{opacity:0.85}}
  .tab.active{{opacity:1;box-shadow:0 2px 8px rgba(0,0,0,0.25)}}
  .tab-brand{{background:#2980B9}}
  .tab-region{{background:#27AE60}}
  .tab-houjin{{background:#8E44AD}}
  .tab-trend{{background:#E67E22}}
  .tab-history{{background:#E74C3C}}
  .tab-convert{{background:#16A085}}
  .tab-snapshot{{background:#2C3E50}}
  .tab-director{{background:#1A5276}}
  .tab-movement{{background:#B7950B}}
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
    <div class="kpi-value">{sum_pr + get_orangetwist_count(f"{y}/{m:02d}", ot_settings):,} 院</div>
  </div>
  <div class="kpi-card">
    <div class="kpi-label">全拠点</div>
    <div class="kpi-value">{sum_all:,} 院</div>
  </div>
</div>

<div class="tabs">
  <div class="tab tab-brand active" onclick="showTab('brand',this)">📋 ブランド別 集計（月末時点）</div>
  <div class="tab tab-region" onclick="showTab('region',this)">🌏 地域・法人別 集計（月末時点）</div>
  <div class="tab tab-houjin" onclick="showTab('houjin',this)">🏢 法人別集計（フィー計算）</div>
  <div class="tab tab-trend" onclick="showTab('trend',this)">📅 年月別 院数サマリー</div>
  <div class="tab tab-history" onclick="showTab('history',this)">🏥 開院・閉院履歴</div>
  <div class="tab tab-convert" onclick="showTab('convert',this)">🔵 業態転換履歴</div>
  <div class="tab tab-snapshot" onclick="showTab('snapshot',this)">🏥 在院一覧（月末時点）</div>
  <div class="tab tab-director" onclick="showTab('director',this)">👨‍⚕️ 院長履歴</div>
  <div class="tab tab-movement" onclick="showTab('movement',this)">🔀 先生の異動履歴</div>
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
    <div class="period-bar" style="background:#f8f9fa;border:1px solid #ddd;border-radius:6px;padding:10px 16px;margin-bottom:12px;display:flex;align-items:center;gap:10px;flex-wrap:wrap">
      <span style="font-weight:bold;font-size:13px">表示期間：</span>
      <select id="startYear">{year_options}</select>年
      <select id="startMonth">{month_options}</select>月 〜
      <select id="endYear">{year_options}</select>年
      <select id="endMonth">{month_options}</select>月
      <button onclick="filterPeriod()" style="padding:5px 14px;background:{C_BLUE};color:white;border:none;border-radius:4px;cursor:pointer;font-size:13px">適用</button>
      <span id="periodStatus" style="font-size:12px;color:#666"></span>
    </div>
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
    <div class="section-title">{y}年{m}月末 法人別集計（フィー計算）</div>
    {df_to_html_table(houjin_df, right_align_nums=True)}
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

<div id="snapshot" class="content">
  <div class="box">
    <div class="section-title">在院一覧（月末時点）</div>
    <div style="display:flex;gap:12px;align-items:flex-end;flex-wrap:wrap;margin-bottom:12px;padding:12px;background:#f8f9fa;border-radius:6px;border:1px solid #ddd">
      <div>
        <div style="font-size:12px;color:#666;margin-bottom:4px">月末時点</div>
        <select id="snapYear" style="padding:4px 8px;border:1px solid #ccc;border-radius:4px"></select>年
        <select id="snapMonth" style="padding:4px 8px;border:1px solid #ccc;border-radius:4px"></select>月末
      </div>
      <div>
        <div style="font-size:12px;color:#666;margin-bottom:4px">ブランド</div>
        <select id="snapBrand" style="padding:4px 8px;border:1px solid #ccc;border-radius:4px;min-width:150px">
          <option value="">（全て）</option>
        </select>
      </div>
      <div>
        <div style="font-size:12px;color:#666;margin-bottom:4px">国内／海外</div>
        <select id="snapRegion" style="padding:4px 8px;border:1px solid #ccc;border-radius:4px">
          <option value="">（全て）</option>
          <option value="国内">国内</option>
          <option value="海外">海外</option>
        </select>
      </div>
      <div>
        <div style="font-size:12px;color:#666;margin-bottom:4px">法人名（部分一致）</div>
        <input id="snapHoujin" type="text" placeholder="例：孝和会" style="padding:4px 8px;border:1px solid #ccc;border-radius:4px;min-width:150px">
      </div>
      <button onclick="runSnapshot()" style="padding:6px 16px;background:#2980B9;color:white;border:none;border-radius:4px;cursor:pointer;font-size:13px">検索</button>
      <button onclick="downloadCSV()" style="padding:6px 16px;background:#27AE60;color:white;border:none;border-radius:4px;cursor:pointer;font-size:13px">📥 CSV ダウンロード</button>
    </div>
    <div style="display:flex;gap:8px;margin-bottom:12px;align-items:center">
      <span style="font-size:13px;color:#666">グループ表示：</span>
      <button id="grpBrand" onclick="setGroupBy('brand')"
        style="padding:5px 14px;border-radius:16px;border:none;cursor:pointer;font-size:13px;background:#2980B9;color:white;font-weight:bold">
        ブランド別
      </button>
      <button id="grpHoujin" onclick="setGroupBy('houjin')"
        style="padding:5px 14px;border-radius:16px;border:none;cursor:pointer;font-size:13px;background:#ddd;color:#333">
        法人別
      </button>
    </div>
    <div id="snapshotResult"><p style="color:#999;font-size:13px">月末と条件を選択して「検索」を押してください。</p></div>
  </div>
</div>

<div id="director" class="content">
  <div class="box">
    <div class="section-title">院長履歴（月末時点）</div>
    <p style="font-size:12px;color:#666;margin-bottom:8px">※過去分院長変更履歴（2021年〜）および2025年9月以降の人事通達データを統合表示。黄色セルは院長交代。</p>
    {director_html}
  </div>
</div>

<div id="movement" class="content">
  <div class="box">
    <div class="section-title">先生の異動履歴</div>

    <!-- ビュー切り替え -->
    <div style="margin-bottom:12px;display:flex;gap:8px;flex-wrap:wrap">
      <button id="mvBtnA" onclick="mvSwitch('A')"
        style="padding:5px 14px;border-radius:16px;border:none;cursor:pointer;font-size:13px;background:#B7950B;color:white;font-weight:bold">
        A：先生タイムライン
      </button>
      <button id="mvBtnB" onclick="mvSwitch('B')"
        style="padding:5px 14px;border-radius:16px;border:none;cursor:pointer;font-size:13px;background:#ddd;color:#333">
        B：玉突き人事検出
      </button>
      <button id="mvBtnC" onclick="mvSwitch('C')"
        style="padding:5px 14px;border-radius:16px;border:none;cursor:pointer;font-size:13px;background:#ddd;color:#333">
        C：月次配属マップ
      </button>
    </div>

    <!-- Case A: 先生タイムライン（複数選択対応） -->
    <div id="mvViewA">
      <div style="display:flex;gap:16px;flex-wrap:wrap;align-items:flex-start">
        <!-- 左：先生選択パネル -->
        <div style="flex:0 0 240px">
          <div style="font-size:13px;font-weight:bold;margin-bottom:6px">先生を選択（複数可）</div>
          <input id="mvDoctorInput" type="text" placeholder="名前で絞り込み..."
            style="width:100%;padding:5px 8px;border:1px solid #ccc;border-radius:4px;font-size:13px;box-sizing:border-box;margin-bottom:4px"
            oninput="mvFilterDoctors()">
          <div style="display:flex;gap:6px;align-items:center;margin-bottom:4px;font-size:12px">
            <button onclick="mvSelectAll(true)"
              style="padding:2px 8px;border:1px solid #ccc;border-radius:4px;cursor:pointer;font-size:11px;background:#f0f0f0">全選択</button>
            <button onclick="mvSelectAll(false)"
              style="padding:2px 8px;border:1px solid #ccc;border-radius:4px;cursor:pointer;font-size:11px;background:#f0f0f0">全解除</button>
            <span id="mvSelectedCount" style="color:#B7950B;font-weight:bold">0名選択中</span>
          </div>
          <div id="mvDoctorList"
            style="height:300px;overflow-y:auto;border:1px solid #ccc;border-radius:4px;padding:4px;background:white">
          </div>
        </div>
        <!-- 右：ガントチャート -->
        <div style="flex:1;min-width:0">
          <button onclick="mvDrawGantt()"
            style="margin-bottom:10px;padding:6px 18px;background:#B7950B;color:white;border:none;border-radius:4px;cursor:pointer;font-size:13px;font-weight:bold">
            📊 ガントチャートを表示
          </button>
          <div id="mvGanttArea" style="overflow-x:auto"></div>
        </div>
      </div>
    </div>

    <!-- Case B: 玉突き人事検出 -->
    <div id="mvViewB" style="display:none">
      <div style="display:flex;gap:10px;align-items:center;margin-bottom:12px;flex-wrap:wrap">
        <span style="font-size:13px">月を選択：</span>
        <select id="mvChainMonth"
          style="padding:5px 10px;border:1px solid #ccc;border-radius:4px;font-size:13px">
        </select>
        <button onclick="mvDetectChains()"
          style="padding:5px 16px;background:#E74C3C;color:white;border:none;border-radius:4px;cursor:pointer;font-size:13px;font-weight:bold">
          🔀 玉突き検出
        </button>
        <span style="font-size:12px;color:#888">先生Aが院Xから院Yへ→同月に先生Bが院Xへ着任、の連鎖を検出します</span>
      </div>
      <div id="mvChainArea"></div>
    </div>

    <!-- Case C: 月次配属マップ -->
    <div id="mvViewC" style="display:none">
      <div style="display:flex;gap:10px;align-items:center;margin-bottom:12px;flex-wrap:wrap">
        <span style="font-size:13px">月を選択：</span>
        <select id="mvMonthSelect"
          style="padding:5px 10px;border:1px solid #ccc;border-radius:4px;font-size:13px"
          onchange="mvDrawMonthMap()">
        </select>
        <span style="font-size:12px;color:#888">🟡 前月から異動あり</span>
      </div>
      <div id="mvMonthMapArea" style="overflow-x:auto;max-height:70vh;overflow-y:auto"></div>
    </div>
  </div>
</div>

<div class="updated">最終更新: {today.strftime('%Y/%m/%d')}</div>

<script>
const ALL_DATA = JSON.parse('{json_str}');
const DOCTOR_STINTS = JSON.parse('{doctor_stints_json}');
const ALL_DOCTORS   = {all_doctors_json};
const ALL_CLINICS   = {all_clinics_json};
const MONTHS_LIST   = {months_mv_json};
const PROMOTION_EVENTS   = JSON.parse('{promotion_events_json}');
const RESIGNATION_EVENTS = JSON.parse('{resignation_events_json}');
const BRAND_LABELS = {brand_labels_json};
const CLINIC_DATA = JSON.parse('{clinic_json_escaped}');
const UNIQUE_BRANDS = {brands_json};
const HOUJIN_ORDER = {houjin_order_json};
const LABEL_IR = {json.dumps(LABEL_IR, ensure_ascii=False)};
const LABEL_ALL = {json.dumps(LABEL_ALL, ensure_ascii=False)};
const ORANGE_TWIST_COUNT = {ot_settings["count"]};
const OT_START = "{ot_settings["start"]}";
function getOT(monthKey) {{ return monthKey >= OT_START ? ORANGE_TWIST_COUNT : 0; }}

function showTab(id, el) {{
  document.querySelectorAll('.content').forEach(e=>e.classList.remove('active'));
  document.querySelectorAll('.tab').forEach(e=>e.classList.remove('active'));
  document.getElementById(id).classList.add('active');
  if(el) el.classList.add('active');
}}

// ── 先生の異動履歴 ──────────────────────────────
// クリニック別カラー（最大20色）
const CLINIC_COLORS = [
  '#2980B9','#27AE60','#E67E22','#8E44AD','#E74C3C',
  '#16A085','#D35400','#2C3E50','#F39C12','#1ABC9C',
  '#7D3C98','#1F618D','#148F77','#B7950B','#922B21',
  '#117A65','#784212','#1A5276','#196F3D','#6C3483'
];
const clinicColorMap = {{}};
function getClinicColor(clinic) {{
  if (!clinicColorMap[clinic]) {{
    const idx = Object.keys(clinicColorMap).length % CLINIC_COLORS.length;
    clinicColorMap[clinic] = CLINIC_COLORS[idx];
  }}
  return clinicColorMap[clinic];
}}

function mvSwitch(mode) {{
  ['A','B','C'].forEach(m => {{
    document.getElementById('mvView'+m).style.display = mode===m ? '' : 'none';
    const btn = document.getElementById('mvBtn'+m);
    btn.style.background = mode===m ? '#B7950B' : '#ddd';
    btn.style.color      = mode===m ? 'white'   : '#333';
    btn.style.fontWeight = mode===m ? 'bold'     : 'normal';
  }});
  if (mode==='C') mvDrawMonthMap();
}}

// ── 初期化 ──
(function initMvDoctors() {{
  // Case A: チェックボックスリスト
  const listDiv = document.getElementById('mvDoctorList');
  ALL_DOCTORS.forEach(d => {{
    const label = document.createElement('label');
    label.style.cssText = 'display:flex;align-items:center;gap:5px;padding:3px 5px;cursor:pointer;font-size:13px;border-radius:3px';
    label.onmouseover = ()=>label.style.background='#f0f4f8';
    label.onmouseout  = ()=>label.style.background='';
    const cb = document.createElement('input');
    cb.type = 'checkbox'; cb.value = d;
    cb.onchange = updateMvSelectedCount;
    label.appendChild(cb);
    label.appendChild(document.createTextNode(d));
    listDiv.appendChild(label);
  }});

  // Case B / C: 月セレクター（新しい月順）
  ['mvChainMonth','mvMonthSelect'].forEach(id => {{
    const sel = document.getElementById(id);
    MONTHS_LIST.slice().reverse().forEach(mk => {{
      const opt = document.createElement('option');
      opt.value = mk; opt.textContent = mk;
      sel.appendChild(opt);
    }});
  }});
  // Case C: onchange
  document.getElementById('mvMonthSelect').onchange = mvDrawMonthMap;
}})();

// Case A: 絞り込み・全選択
function mvFilterDoctors() {{
  const q = document.getElementById('mvDoctorInput').value.toLowerCase();
  document.querySelectorAll('#mvDoctorList label').forEach(label => {{
    const name = label.textContent.toLowerCase();
    label.style.display = (!q || name.includes(q)) ? '' : 'none';
  }});
}}
function mvSelectAll(checked) {{
  document.querySelectorAll('#mvDoctorList label').forEach(label => {{
    if (label.style.display !== 'none') {{
      label.querySelector('input').checked = checked;
    }}
  }});
  updateMvSelectedCount();
}}
function updateMvSelectedCount() {{
  const n = document.querySelectorAll('#mvDoctorList input:checked').length;
  document.getElementById('mvSelectedCount').textContent = n + '名選択中';
}}

// Case A: ガントチャート（複数先生対応）
function mvDrawGantt() {{
  const selected = [...document.querySelectorAll('#mvDoctorList input:checked')].map(cb=>cb.value);
  const area = document.getElementById('mvGanttArea');
  if (!selected.length) {{
    area.innerHTML = '<p style="color:#999;font-size:13px">先生をチェックして「表示」を押してください</p>';
    return;
  }}

  const allMonths = MONTHS_LIST;
  const totalCols = allMonths.length;

  // ヘッダー行（年/月）
  let headerRow = '<tr>';
  headerRow += '<th style="min-width:140px;max-width:160px;padding:4px 8px;border:1px solid #ccc;background:#2C3E50;color:white;position:sticky;left:0;z-index:3;font-size:12px;text-align:left">先生名</th>';
  allMonths.forEach(mk => {{
    const yr = mk.split('/')[0];
    const mo = mk.split('/')[1];
    const isJan = mo==='01';
    const bg = isJan ? '#2C3E50' : '#ECF0F1';
    const col = isJan ? 'white' : '#555';
    const label = isJan ? yr : mo;
    headerRow += `<th style="min-width:36px;width:36px;padding:2px 0;text-align:center;font-size:9px;background:${{bg}};color:${{col}};border:1px solid #ccc">${{label}}</th>`;
  }});
  headerRow += '</tr>';

  // 先生ごとの行
  let bodyRows = '';
  const allUsedClinics = new Set();
  selected.forEach(doctor => {{
    const stints = DOCTOR_STINTS[doctor] || [];
    // 月→クリニック マップ
    const mc = {{}};
    stints.forEach(s => {{
      allMonths.forEach(mk => {{
        if (mk >= s.start && mk <= s.end) mc[mk] = s.clinic;
      }});
      allUsedClinics.add(s.clinic);
    }});

    bodyRows += '<tr>';
    bodyRows += `<td style="padding:4px 8px;border:1px solid #ddd;font-weight:bold;font-size:12px;white-space:nowrap;background:white;position:sticky;left:0;z-index:1">${{doctor}}</td>`;
    let prevClinic = null;
    allMonths.forEach(mk => {{
      const clinic = mc[mk] || null;
      if (clinic) {{
        const color = getClinicColor(clinic);
        const borderLeft = (prevClinic && prevClinic !== clinic)
          ? 'border-left:3px solid #E74C3C;'
          : `border-left:1px solid ${{color}};`;
        bodyRows += `<td style="background:${{color}};${{borderLeft}}border-top:1px solid ${{color}};border-right:1px solid ${{color}};border-bottom:1px solid ${{color}};min-width:36px;width:36px;padding:0" title="${{clinic}} (${{mk}})"></td>`;
      }} else {{
        bodyRows += `<td style="background:#f5f5f5;border:1px solid #e8e8e8;min-width:36px;width:36px;padding:0"></td>`;
      }}
      prevClinic = clinic;
    }});
    bodyRows += '</tr>';
  }});

  let html = `<div style="overflow-x:auto;overflow-y:auto;max-height:70vh">
    <table style="border-collapse:collapse;font-size:12px">
    <thead>${{headerRow}}</thead>
    <tbody>${{bodyRows}}</tbody>
    </table></div>`;

  // 凡例
  html += '<div style="display:flex;gap:6px;flex-wrap:wrap;margin-top:8px;font-size:11px;align-items:center">';
  html += '<span style="font-size:12px;color:#666">凡例：</span>';
  allUsedClinics.forEach(c => {{
    html += `<span style="background:${{getClinicColor(c)}};color:white;padding:2px 6px;border-radius:3px">${{c}}</span>`;
  }});
  html += '<span style="margin-left:8px;font-size:11px;color:#E74C3C">│ 赤い左境界線 = 院の変更</span>';
  html += '</div>';

  area.innerHTML = html;
}}

// ── Case B: 玉突き人事検出 ──
function mvDetectChains() {{
  const monthKey = document.getElementById('mvChainMonth').value;
  const area = document.getElementById('mvChainArea');
  if (!monthKey) return;

  const monthIdx = MONTHS_LIST.indexOf(monthKey);
  const prevMonth = monthIdx > 0 ? MONTHS_LIST[monthIdx - 1] : null;
  if (!prevMonth) {{
    area.innerHTML = '<p style="color:#999;font-size:13px">最初の月は前月データがないため検出できません</p>';
    return;
  }}

  // 今月・前月の配属マップ
  const curMap = {{}}, prevMap = {{}};
  Object.entries(DOCTOR_STINTS).forEach(([doc, stints]) => {{
    stints.forEach(s => {{
      if (s.start <= monthKey  && s.end >= monthKey)  curMap[doc]  = s.clinic;
      if (s.start <= prevMonth && s.end >= prevMonth) prevMap[doc] = s.clinic;
    }});
  }});

  // 異動リストを構築
  const moves = [];
  const allDocs = new Set([...Object.keys(curMap), ...Object.keys(prevMap)]);
  allDocs.forEach(doc => {{
    const from = prevMap[doc] || null;
    const to   = curMap[doc]  || null;
    if (from && to && from !== to) moves.push({{doctor:doc, from, to}});
  }});

  if (!moves.length) {{
    area.innerHTML = `<div style="padding:16px;background:#D5F5E3;border-radius:8px;color:#1E8449">
      <b>${{monthKey}} は院長の異動が検出されませんでした</b>
    </div>`;
    return;
  }}

  // グラフ構築：院 → [出ていった人の異動]
  const outEdge = {{}};  // fromClinic → 異動リスト
  const inEdge  = {{}};  // toClinic   → 異動リスト
  moves.forEach(m => {{
    outEdge[m.from] = outEdge[m.from] || []; outEdge[m.from].push(m);
    inEdge [m.to]   = inEdge [m.to]   || []; inEdge [m.to].push(m);
  }});

  // 昇格・新任データ（Python側で収集）
  const promoThisMonth = PROMOTION_EVENTS[monthKey] || [];
  const promoByClinic = {{}};  // to_clinic → promo info
  promoThisMonth.forEach(p => {{ promoByClinic[p.to] = p; }});

  // ── チェーン抽出 ──
  // 退職データがあれば退職起点、なければグラフ起点にフォールバック
  const resignations = RESIGNATION_EVENTS[monthKey] || [];
  const chains = [];  // 構造: resign + chainの配列
  const used = new Set();

  if (resignations.length > 0) {{
    // 【退職起点】退職で空いた院を起点にinEdgeを逆引きしてチェーン構築
    resignations.forEach(resign => {{
      const vacantClinic = resign.clinic;
      // 退職院に誰かが入ってきたか確認
      const firstFillers = (inEdge[vacantClinic] || []).filter(m => !used.has(m.doctor));
      if (!firstFillers.length) return;

      const chain = [];
      let cur = vacantClinic;
      const visited = new Set();

      while (cur && !visited.has(cur)) {{
        visited.add(cur);
        const movesIn = (inEdge[cur] || []).filter(m => !used.has(m.doctor));
        if (movesIn.length) {{
          const move = movesIn[0];
          chain.push(move);
          used.add(move.doctor);
          cur = move.from;  // この先生の出身院を次のループで調べる
        }} else {{
          // 通常異動なし → 昇格・新任か確認
          const promo = promoByClinic[cur];
          if (promo && !used.has(promo.doctor)) {{
            chain.push({{ doctor: promo.doctor, from: promo.from, to: cur, isPromotion: true }});
            used.add(promo.doctor);
          }}
          break;
        }}
      }}

      if (chain.length >= 1) chains.push({{ resign, chain }});
    }});
  }} else {{
    // 【グラフ起点フォールバック】前月から入がない院を起点に従来方式で検出
    moves.forEach(startMove => {{
      if (used.has(startMove.doctor)) return;
      if (inEdge[startMove.from] && inEdge[startMove.from].length > 0) return;

      const chain = [startMove];
      used.add(startMove.doctor);
      let cur = startMove.to;
      let safety = 0;
      while (cur && safety < 20) {{
        const nexts = (outEdge[cur] || []).filter(m => !used.has(m.doctor));
        if (!nexts.length) break;
        const next = nexts[0];
        chain.push(next);
        used.add(next.doctor);
        cur = next.to;
        safety++;
      }}
      // 昇格を先頭に追加
      const startClinic = chain[0].from;
      if (promoByClinic[startClinic]) {{
        const promo = promoByClinic[startClinic];
        chain.unshift({{ doctor: promo.doctor, from: promo.from, to: startClinic, isPromotion: true }});
      }}
      if (chain.length >= 2) chains.push({{ resign: null, chain }});
    }});
  }}

  // ── 結果表示 ──
  let html = '';

  if (!chains.length) {{
    html += `<div style="padding:12px 16px;background:#FEF9E7;border:1px solid #F1C40F;border-radius:8px;margin-bottom:16px">
      <b style="color:#7D6608">${{monthKey}}：玉突き人事は検出されませんでした</b>
      <span style="font-size:12px;color:#888;margin-left:8px">（異動件数: ${{moves.length}}件）</span>
    </div>`;
  }} else {{
    html += `<div style="padding:10px 16px;background:#FDFEFE;border:1px solid #B7950B;border-radius:8px;margin-bottom:16px">
      <b style="color:#7D6608;font-size:14px">🔀 ${{monthKey}} に検出された玉突き人事：${{chains.length}}件</b>
    </div>`;
    chains.forEach((chainObj, i) => {{
      const {{resign, chain}} = chainObj;
      const totalLen = chain.length + (resign ? 1 : 0);
      html += `<div style="background:#FFF9C4;border:2px solid #B7950B;border-radius:10px;padding:14px 16px;margin-bottom:14px">`;
      html += `<div style="font-weight:bold;color:#6E2F00;margin-bottom:10px;font-size:13px">
        🔀 玉突き #${{i+1}}（${{totalLen}}連鎖）${{resign ? `<span style="color:#C0392B;margin-left:8px">起点：${{resign.doctor}}退職</span>` : ''}}
      </div>`;
      html += `<div style="display:flex;align-items:flex-start;flex-wrap:wrap;gap:6px">`;

      // 【退職起点の場合】退職ブロックを左端に表示
      if (resign) {{
        html += `<div style="background:#FDEDEC;border:2px solid #C0392B;border-radius:8px;padding:8px 12px;text-align:center;min-width:110px">
          <div style="font-size:10px;font-weight:bold;color:#C0392B;margin-bottom:4px">🚪 退職（起点）</div>
          <div style="font-size:13px;font-weight:bold;color:#2C3E50">${{resign.doctor}}</div>
          <div style="font-size:10px;color:#888;margin-top:2px">${{resign.clinic}}</div>
        </div>`;
        // チェーン各ノード（退職院←新院長←出身院 の順）
        chain.forEach(move => {{
          html += `<div style="font-size:18px;color:#C0392B;font-weight:bold;align-self:center">→</div>`;
          const doctorBg = move.isPromotion ? '#27AE60' : '#2C3E50';
          const fromLabel = move.isPromotion
            ? `<span style="color:#27AE60;font-size:10px">⬆ 昇格就任</span>`
            : `<span style="font-size:10px;color:#888">${{move.from}}より</span>`;
          html += `<div style="text-align:center;min-width:110px">
            <div style="background:${{getClinicColor(move.to)}};color:white;border-radius:6px;padding:6px 10px;font-size:12px;font-weight:bold">${{move.to}}</div>
            <div style="margin-top:4px;display:flex;flex-direction:column;align-items:center;gap:2px">
              <div style="background:${{doctorBg}};color:white;border-radius:10px;padding:2px 10px;font-size:11px">${{move.doctor}}</div>
              <div>${{fromLabel}}</div>
            </div>
          </div>`;
        }});
      }} else {{
        // 【グラフ起点フォールバック】従来の表示（昇格元 → ... → 空席）
        const firstMove = chain[0];
        if (firstMove.isPromotion) {{
          html += `<div style="background:#D5F5E3;border:2px solid #27AE60;border-radius:6px;padding:6px 12px;font-size:12px;font-weight:bold;color:#1E8449;text-align:center">
            ${{firstMove.from}}<br><span style="font-size:10px;color:#27AE60">（昇格元）</span>
          </div>`;
        }} else {{
          html += `<div style="background:#ECF0F1;border:2px solid #BDC3C7;border-radius:6px;padding:6px 12px;font-size:12px;font-weight:bold;color:#555;text-align:center">
            ${{firstMove.from}}<br><span style="font-size:10px;color:#999">（空きが生じる）</span>
          </div>`;
        }}
        chain.forEach(move => {{
          html += `<div style="font-size:20px;color:#B7950B;font-weight:bold;align-self:center">→</div>`;
          const doctorBg = move.isPromotion ? '#27AE60' : '#2C3E50';
          html += `<div style="text-align:center">
            <div style="background:${{doctorBg}};color:white;border-radius:12px;padding:2px 10px;font-size:11px;white-space:nowrap;margin-bottom:3px">${{move.doctor}}${{move.isPromotion?' ⬆':''}}</div>
            <div style="background:${{getClinicColor(move.to)}};color:white;border-radius:6px;padding:6px 10px;font-size:12px;font-weight:bold;white-space:nowrap">${{move.to}}</div>
          </div>`;
        }});
      }}  // end else (fallback)
      html += `</div></div>`;
    }});
  }}

  // すべての異動一覧
  const chainDoctors = new Set(chains.flatMap(obj => obj.chain.map(m => m.doctor)));
  chains.forEach(obj => {{ if (obj.resign) chainDoctors.add(obj.resign.doctor); }});
  const otherMoves = moves.filter(m => !chainDoctors.has(m.doctor));
  html += `<div style="margin-top:16px">`;
  html += `<b style="font-size:13px">${{monthKey}} の全異動一覧（${{moves.length}}件）</b>`;
  html += mvBuildMovesTable(moves, chainDoctors);
  html += `</div>`;

  area.innerHTML = html;
}}

function mvBuildMovesTable(moves, highlightDoctors) {{
  let html = '<table style="border-collapse:collapse;font-size:13px;width:100%;margin-top:8px">';
  html += '<thead><tr style="background:#2C3E50;color:white">';
  html += '<th style="padding:6px 12px;border:1px solid #555">先生名</th>';
  html += '<th style="padding:6px 12px;border:1px solid #555">異動前の院</th>';
  html += '<th style="padding:6px 12px;border:1px solid #555;width:24px"></th>';
  html += '<th style="padding:6px 12px;border:1px solid #555">異動後の院</th>';
  html += '</tr></thead><tbody>';
  moves.forEach((m, i) => {{
    const isChain = highlightDoctors && highlightDoctors.has(m.doctor);
    const bg = isChain ? '#FFF9C4' : (i%2===0 ? 'white' : '#f8f9fa');
    html += `<tr style="background:${{bg}}">`;
    html += `<td style="padding:5px 12px;border:1px solid #ddd;font-weight:bold">${{isChain?'🔀 ':''}}${{m.doctor}}</td>`;
    html += `<td style="padding:5px 12px;border:1px solid #ddd;color:#666">${{m.from}}</td>`;
    html += `<td style="padding:5px 12px;border:1px solid #ddd;text-align:center;color:#B7950B;font-weight:bold">→</td>`;
    html += `<td style="padding:5px 12px;border:1px solid #ddd">${{m.to}}</td>`;
    html += '</tr>';
  }});
  html += '</tbody></table>';
  return html;
}}

// Case C: 月次配属マップ
function mvDrawMonthMap() {{
  const selMonth = document.getElementById('mvMonthSelect').value;
  const area = document.getElementById('mvMonthMapArea');
  if (!selMonth) return;

  const monthIdx = MONTHS_LIST.indexOf(selMonth);
  const prevMonth = monthIdx > 0 ? MONTHS_LIST[monthIdx - 1] : null;

  const curMap  = {{}};
  const prevMap = {{}};
  Object.entries(DOCTOR_STINTS).forEach(([doctor, stints]) => {{
    stints.forEach(s => {{
      if (s.start <= selMonth && s.end >= selMonth) curMap[doctor]  = s.clinic;
      if (prevMonth && s.start <= prevMonth && s.end >= prevMonth)  prevMap[doctor] = s.clinic;
    }});
  }});

  const rows = Object.entries(curMap).sort((a,b) => a[1].localeCompare(b[1],'ja'));
  let html = `<p style="font-size:13px;color:#666;margin-bottom:8px"><b>${{selMonth}} 時点の配属一覧（${{rows.length}}名）</b></p>`;
  html += '<table style="border-collapse:collapse;font-size:13px;width:100%">';
  html += `<thead><tr style="background:#2C3E50;color:white">
    <th style="padding:6px 12px;border:1px solid #555;min-width:120px">先生名</th>
    <th style="padding:6px 12px;border:1px solid #555;min-width:180px">今月の院</th>
    <th style="padding:6px 12px;border:1px solid #555;min-width:180px">前月の院</th>
    <th style="padding:6px 12px;border:1px solid #555">変化</th>
  </tr></thead><tbody>`;

  rows.forEach(([doctor, clinic]) => {{
    const prev = prevMap[doctor] || '―';
    const moved = prev !== '―' && prev !== clinic;
    const bg = moved ? '#FFF9C4' : 'white';
    const change = moved ? `${{prev}} → ${{clinic}}` : '変化なし';
    const changeColor = moved ? '#E67E22' : '#999';
    html += `<tr style="background:${{bg}}">
      <td style="padding:5px 12px;border:1px solid #ddd;font-weight:bold">${{doctor}}</td>
      <td style="padding:5px 12px;border:1px solid #ddd">${{clinic}}</td>
      <td style="padding:5px 12px;border:1px solid #ddd;color:#666">${{prev}}</td>
      <td style="padding:5px 12px;border:1px solid #ddd;color:${{changeColor}};font-weight:${{moved?'bold':'normal'}}">${{change}}</td>
    </tr>`;
  }});
  html += '</tbody></table>';
  area.innerHTML = html;
}}
// ─────────────────────────────────────────────────

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

  // 列順：左=IR・広報用、右=全拠点
  let html = '<table class="compare-table" style="border-collapse:collapse;width:100%">';
  html += '<thead><tr>';
  html += '<th rowspan="2" style="vertical-align:middle">ブランド</th>';
  html += '<th colspan="3" style="text-align:center;background:#d35400">IR・広報用</th>';
  html += '<th colspan="3" style="text-align:center;background:#1a5276">全拠点</th>';
  html += '</tr><tr>';
  html += '<th style="background:#d35400">' + startKey + '</th>';
  html += '<th style="background:#d35400">' + endKey + '</th>';
  html += '<th style="background:#d35400">増減</th>';
  html += '<th style="background:#1a5276">' + startKey + '</th>';
  html += '<th style="background:#1a5276">' + endKey + '</th>';
  html += '<th style="background:#1a5276">増減</th>';
  html += '</tr></thead><tbody>';

  function diffStr(d) {{ return d > 0 ? '+' + d : String(d); }}
  function diffStyle(d) {{ return d > 0 ? 'color:#27ae60' : d < 0 ? 'color:#c0392b' : ''; }}

  // ブランド行
  for (let i = 0; i < endRows.length; i++) {{
    const er = endRows[i];
    const sr = startRows[i] || {{label: er.label, all: 0, pr: 0}};
    const dPr  = er.pr  - sr.pr;
    const dAll = er.all - sr.all;
    html += '<tr>';
    html += '<td style="text-align:left">' + er.label + '</td>';
    html += '<td style="text-align:right">' + sr.pr  + '</td>';
    html += '<td style="text-align:right">' + er.pr  + '</td>';
    html += '<td style="text-align:right;' + diffStyle(dPr)  + '">' + diffStr(dPr)  + '</td>';
    html += '<td style="text-align:right">' + sr.all + '</td>';
    html += '<td style="text-align:right">' + er.all + '</td>';
    html += '<td style="text-align:right;' + diffStyle(dAll) + '">' + diffStr(dAll) + '</td>';
    html += '</tr>';
  }}

  // OrangeTwist行（開始年月を考慮）
  const startOT = getOT(startKey);
  const endOT   = getOT(endKey);
  const diffOT  = endOT - startOT;
  html += '<tr style="background:#FFF9C4">';
  html += '<td style="text-align:left">OrangeTwist</td>';
  html += '<td style="text-align:right">' + startOT + '</td>';
  html += '<td style="text-align:right">' + endOT   + '</td>';
  html += '<td style="text-align:right;' + diffStyle(diffOT) + '">' + diffStr(diffOT) + '</td>';
  html += '<td style="text-align:right;color:#888">別管理</td>';
  html += '<td style="text-align:right;color:#888">別管理</td>';
  html += '<td style="text-align:right;color:#888">－</td>';
  html += '</tr>';

  // 合計行（IR・広報用＝ブランド合計＋OrangeTwist）
  const sPr  = startData.sum_pr,  ePr  = endData.sum_pr;
  const sAll = startData.sum_all, eAll = endData.sum_all;
  const sIR  = sPr + startOT, eIR = ePr + endOT;  // OrangeTwistは開始年月を考慮
  const dIR  = eIR  - sIR,  dAll = eAll - sAll;
  html += '<tr style="background:#E67E22;color:white;font-weight:bold">';
  html += '<td style="text-align:left">合計</td>';
  html += '<td style="text-align:right">' + sIR  + '</td>';
  html += '<td style="text-align:right">' + eIR  + '</td>';
  html += '<td style="text-align:right;' + diffStyle(dIR)  + '">' + diffStr(dIR)  + '</td>';
  html += '<td style="text-align:right">' + sAll + '</td>';
  html += '<td style="text-align:right">' + eAll + '</td>';
  html += '<td style="text-align:right;' + diffStyle(dAll) + '">' + diffStr(dAll) + '</td>';
  html += '</tr>';

  html += '</tbody></table>';
  container.innerHTML = html;
}}

// Populate brand dropdown
(function initSnapshot() {{
  const keys = Object.keys(ALL_DATA).sort();
  const snapYearEl = document.getElementById('snapYear');
  const snapMonthEl = document.getElementById('snapMonth');
  const snapBrandEl = document.getElementById('snapBrand');

  // Year/month options
  const years = [...new Set(keys.map(k => k.split('/')[0]))];
  years.forEach(y => snapYearEl.innerHTML += `<option value="${{y}}">${{y}}</option>`);
  for (let m = 1; m <= 12; m++) {{
    snapMonthEl.innerHTML += `<option value="${{String(m).padStart(2,'0')}}">${{m}}月</option>`;
  }}
  // Default to latest month
  if (keys.length) {{
    const [ey, em] = keys[keys.length-1].split('/');
    snapYearEl.value = ey;
    snapMonthEl.value = em;
  }}
  // Brand options
  UNIQUE_BRANDS.forEach(b => snapBrandEl.innerHTML += `<option value="${{b}}">${{b}}</option>`);
}})();

function checkActiveJS(clinic, targetDateStr) {{
  if (!clinic.open_date) return false;
  const base = new Date(clinic.open_date);
  const target = new Date(targetDateStr + 'T23:59:59');
  if (base > target) return false;
  const finalEnd = clinic.conv_date || clinic.close_date;
  if (finalEnd) {{
    const endDate = new Date(finalEnd);
    if (endDate <= target) return false;
  }}
  return true;
}}

let lastSnapshotData = [];
let currentGroupBy = 'brand';

function setGroupBy(mode) {{
  currentGroupBy = mode;
  document.getElementById('grpBrand').style.background  = mode==='brand'  ? '#2980B9' : '#ddd';
  document.getElementById('grpBrand').style.color       = mode==='brand'  ? 'white'   : '#333';
  document.getElementById('grpHoujin').style.background = mode==='houjin' ? '#8E44AD' : '#ddd';
  document.getElementById('grpHoujin').style.color      = mode==='houjin' ? 'white'   : '#333';
  if (lastSnapshotData.length) renderSnapshot(lastSnapshotData);
}}

function runSnapshot() {{
  const year = document.getElementById('snapYear').value;
  const month = document.getElementById('snapMonth').value;
  const brand = document.getElementById('snapBrand').value;
  const region = document.getElementById('snapRegion').value;
  const houjin = document.getElementById('snapHoujin').value.trim();

  // Last day of selected month
  const lastDay = new Date(parseInt(year), parseInt(month), 0).getDate();
  const targetDate = `${{year}}-${{month}}-${{String(lastDay).padStart(2,'0')}}`;

  // Filter clinics
  const filtered = CLINIC_DATA.filter(c => {{
    if (!checkActiveJS(c, targetDate)) return false;
    if (brand && c.brand !== brand) return false;
    if (region && c.region !== region) return false;
    if (houjin && !c.houjin.includes(houjin)) return false;
    return true;
  }});

  lastSnapshotData = filtered;

  const container = document.getElementById('snapshotResult');
  if (!filtered.length) {{
    container.innerHTML = '<p style="color:#c0392b;font-size:13px">該当する院がありません。</p>';
    return;
  }}
  renderSnapshot(filtered);
}}

function renderSnapshot(filtered) {{
  const container = document.getElementById('snapshotResult');
  const year  = document.getElementById('snapYear').value;
  const month = document.getElementById('snapMonth').value;
  const isByHoujin = currentGroupBy === 'houjin';

  // グループキーと表示名
  const getKey  = c => isByHoujin ? (c.houjin || '（法人名なし）') : c.brand;
  const hdrBg   = isByHoujin ? '#8E44AD' : '#2C3E50';

  // グループ化
  const grouped = {{}};
  filtered.forEach(c => {{
    const k = getKey(c);
    if (!grouped[k]) grouped[k] = [];
    grouped[k].push(c);
  }});

  // ソート
  const sorted = Object.entries(grouped).sort((a,b) => {{
    if (!isByHoujin) {{
      const ia = UNIQUE_BRANDS.indexOf(a[0]);
      const ib = UNIQUE_BRANDS.indexOf(b[0]);
      if (ia === -1 && ib === -1) return a[0].localeCompare(b[0], 'ja');
      if (ia === -1) return 1;
      if (ib === -1) return -1;
      return ia - ib;
    }}
    // 法人設定シートの順序で並べる
    const ia = HOUJIN_ORDER.indexOf(a[0]);
    const ib = HOUJIN_ORDER.indexOf(b[0]);
    if (ia === -1 && ib === -1) return a[0].localeCompare(b[0], 'ja');
    if (ia === -1) return 1;
    if (ib === -1) return -1;
    return ia - ib;
  }});

  let html = `<p style="font-size:13px;color:#666;margin-bottom:8px"><b>${{year}}年${{parseInt(month)}}月末時点：${{filtered.length}}院</b></p>`;

  sorted.forEach(([key, clinics], idx) => {{
    const uid = 'snap' + currentGroupBy + idx;
    html += `<div style="margin-bottom:6px;border:1px solid #ddd;border-radius:6px;overflow:hidden">`;
    html += `<div onclick="toggleAcc('${{uid}}a','${{uid}}r')" style="padding:8px 14px;cursor:pointer;background:#f8f9fa;display:flex;justify-content:space-between;align-items:center">`;
    html += `<span><span id="${{uid}}r" style="font-size:11px">▶</span> <b>${{key}}</b>（${{clinics.length}}院）</span></div>`;
    html += `<div id="${{uid}}a" style="display:none;overflow-x:auto">`;
    html += `<table style="border-collapse:collapse;width:100%;font-size:13px">`;
    html += `<thead><tr style="background:${{hdrBg}};color:white">`;
    html += `<th style="padding:6px 10px;border:1px solid #555">院ID</th>`;
    html += `<th style="padding:6px 10px;border:1px solid #555">院名</th>`;
    html += `<th style="padding:6px 10px;border:1px solid #555">ブランド</th>`;
    html += `<th style="padding:6px 10px;border:1px solid #555">業態</th>`;
    if (!isByHoujin) html += `<th style="padding:6px 10px;border:1px solid #555">法人名</th>`;
    html += `<th style="padding:6px 10px;border:1px solid #555">国内／海外</th>`;
    html += `<th style="padding:6px 10px;border:1px solid #555">開院日</th>`;
    html += `</tr></thead><tbody>`;
    clinics.forEach((c, ci) => {{
      const bg = ci % 2 === 0 ? 'white' : '#f8f9fa';
      html += `<tr style="background:${{bg}}">`;
      html += `<td style="padding:5px 10px;border:1px solid #ddd;text-align:right;color:#666">${{c.id || ''}}</td>`;
      html += `<td style="padding:5px 10px;border:1px solid #ddd">${{c.name}}</td>`;
      html += `<td style="padding:5px 10px;border:1px solid #ddd">${{c.brand}}</td>`;
      html += `<td style="padding:5px 10px;border:1px solid #ddd">${{c.gyoutai}}</td>`;
      if (!isByHoujin) html += `<td style="padding:5px 10px;border:1px solid #ddd">${{c.houjin}}</td>`;
      html += `<td style="padding:5px 10px;border:1px solid #ddd;text-align:center">${{c.region}}</td>`;
      html += `<td style="padding:5px 10px;border:1px solid #ddd">${{c.open_date || ''}}</td>`;
      html += `</tr>`;
    }});
    html += `</tbody></table></div></div>`;
  }});

  container.innerHTML = html;
}}

function downloadCSV() {{
  if (!lastSnapshotData.length) {{
    alert('先に「検索」を実行してください。');
    return;
  }}
  const year = document.getElementById('snapYear').value;
  const month = document.getElementById('snapMonth').value;
  const headers = ['院ID','院名','ブランド','業態','法人名','国内／海外','開院日'];
  const rows = lastSnapshotData.map(c => [
    c.id || '', c.name, c.brand, c.gyoutai, c.houjin, c.region, c.open_date || ''
  ]);
  const csvContent = [headers, ...rows].map(r => r.map(v => `"${{String(v).replace(/"/g,'""')}}"`).join(',')).join('\\n');
  const blob = new Blob(['\\uFEFF' + csvContent], {{type:'text/csv;charset=utf-8;'}});
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  a.download = `月次断面_${{year}}${{month}}.csv`;
  a.click();
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
