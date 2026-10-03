# -*- coding: utf-8 -*-
"""
永豐 x DAC 人資考績 AI 專案 — 假資料生成腳本

產出 4 張表,對應 docx 定義的欄位:
  1. employees            員工基本資料
  2. performance_history  考績歷史
  3. review_360           360評鑑
  4. rotation_history     輪調經歷

設計邏輯:
- 以「員工基本資料」為錨點(500 筆),其餘三張表都用「員編」關聯回去
- 考績歷史 / 360評鑑 的年度數量,依員工的「集團年資」決定(年資不夠久的新人不會有 2023 年的考績)
- 輪調經歷用來反推出員工現在的職等/職稱/部門,確保跟員工基本資料一致,而不是兩邊各自隨機、兜不起來
- 考績等第刻意做出「主管打分風格」差異(每個員工指定一位主管,主管有鬆/嚴偏好),
  之後要做「標準落差」分析或 demo 時,這個分佈差異已經先埋好了
- 考績評語 / 表現敘述用片語庫組合而成,長度盡量落在規格要求的 100-150 字 / 140 字上下

執行: python generate_fake_data.py
輸出: output/永豐考績假資料.xlsx (四個工作表) + output/ 下的四個 csv
"""

import random
import math
from datetime import date
import pandas as pd
from faker import Faker

SEED = 42
random.seed(SEED)
fake = Faker("zh_TW")
Faker.seed(SEED)

# --- 單部門測試用設定 -------------------------------------------------------
# DEPARTMENT_FILTER 設 None = 照舊從所有部門隨機抽;
# 設成一個部門名稱(例如 "法令遵循處")= 全部員工的「現職單位/職位」都固定是這個部門,
# 其餘表(考績歷史、360評鑑、輪調經歷)跟著只產這個部門的人,方便先拿小規模資料測試介面。
DEPARTMENT_FILTER = "法令遵循處"
N_EMPLOYEES = 10 if DEPARTMENT_FILTER else 500
OUT_DIR = "output"

# ---------------------------------------------------------------------------
# 參考清單 / 對照表
# ---------------------------------------------------------------------------

DEPARTMENTS = [
    "法令遵循處", "教育發展部", "人力資源處", "風險管理處", "資訊處",
    "財富管理處", "授信管理處", "客服中心", "總務處", "稽核處",
    "數位金融處", "企業金融處", "個人金融處", "財務會計處", "法務處",
    "投資銀行處", "國際業務處", "通路業務處",
]

OTHER_COMPANIES = ["中信銀行", "台新銀行", "國泰世華銀行", "富邦銀行", "玉山銀行", "兆豐銀行", "第一銀行"]
HOME_COMPANY = "永豐銀行"

# 職等 -> 職稱 對照(職等越高,職稱越資深)
GRADE_TITLE_BANDS = [
    (1, 2, "工讀"),
    (3, 4, "專員"),
    (5, 6, "副理"),
    (7, 8, "經理"),
    (9, 10, "資深經理"),
    (11, 12, "協理"),
    (13, 14, "資深協理"),
]

WORK_NATURE_BY_TITLE = {
    "工讀": ["一般行員"],
    "專員": ["一般行員", "專業人員"],
    "副理": ["專業人員", "法遵主管"],
    "經理": ["專業人員", "高級專業人員", "法遵主管"],
    "資深經理": ["高級專業人員", "法遵主管"],
    "協理": ["法遵主管", "高級專業人員"],
    "資深協理": ["法遵主管"],
}

ROTATION_REASONS_FIRST = ["到職", "初任"]
ROTATION_REASONS_LATER = ["升等", "內部輪調", "職務調整"]
ROTATION_MONTHS_DAYS = [(1, 1), (6, 1), (7, 1), (12, 1)]  # 都是月初

SCHOOLS = ["台灣大學", "政治大學", "清華大學", "交通大學", "成功大學", "中央大學",
           "中山大學", "輔仁大學", "東吳大學", "淡江大學", "中興大學", "台北大學"]
MAJORS = ["財務金融系", "會計系", "經濟系", "企業管理系", "資訊管理系", "法律系",
          "風險管理與保險系", "統計系", "國際企業系", "財稅系"]
GRAD_SCHOOLS = ["管理學院碩士班", "財務金融研究所", "會計研究所", "法律研究所", "高階經營管理碩士班(EMBA)"]

GRADE_LEVELS = ["特優", "優", "甲上", "甲"]
GRADE_SCORE = {"特優": 4, "優": 3, "甲上": 2, "甲": 1}  # 給抽樣用的相對分數

# 主管打分風格(之後要做「標準落差」分析的伏筆):
#   -1 = 偏嚴格(等第偏低), 0 = 中性, 1 = 偏寬鬆(等第偏高)
MANAGER_STYLES = ["嚴格", "中性", "中性", "中性", "寬鬆"]  # 中性機率較高,較符合常態

STRENGTH_PHRASES = [
    "對專案推進展現高度主動性", "能有效協調跨部門資源", "在法規遵循細節上掌握度高",
    "溝通表達清晰,能快速取得共識", "對數字敏感度佳,報表產出準確度高",
    "面對臨時交辦任務應變能力強", "能主動發掘流程中的風險並提出改善建議",
    "團隊合作意願高,樂於協助同仁", "對客戶/內部需求回應速度快",
    "在壓力下仍能維持穩定的產出品質",
]
WEAKNESS_PHRASES = [
    "在時間管理上仍有進步空間", "跨部門溝通的主動性可以再加強",
    "對新系統/新制度的熟悉速度稍慢", "簡報與文件呈現的精簡度有待加強",
    "授權與培養下屬的意識可以更積極", "面對模糊不確定的任務時,決策速度偏慢",
]
RECOMMENDATION_PHRASES = [
    "建議持續在現職累積歷練,並給予更大範圍的專案磨練", "建議安排進階管理課程以強化領導能力",
    "建議列入後續晉升評估名單觀察", "建議持續保持目前表現,維持穩定貢獻",
    "建議於下年度加強時間規劃與跨部門協作訓練",
]

PERF_OPENING_BY_GRADE = {
    "特優": "{name}本年度整體表現優異,",
    "優": "{name}本年度表現良好,",
    "甲上": "{name}本年度表現穩定,",
    "甲": "{name}本年度表現符合基本要求,",
}

# ---------------------------------------------------------------------------
# 工具函式
# ---------------------------------------------------------------------------

def pick_grade_title(grade):
    for lo, hi, title in GRADE_TITLE_BANDS:
        if lo <= grade <= hi:
            return title
    return GRADE_TITLE_BANDS[-1][2]


def weighted_grade_level():
    """製造一個偏向中段(甲上/優)的常態分佈,特優最少、甲次少。"""
    return random.choices(GRADE_LEVELS, weights=[8, 27, 40, 25], k=1)[0]


def apply_manager_style(base_level, style):
    """主管風格會把等第往上或往下微調一格,模擬『標準落差』。"""
    idx = GRADE_LEVELS.index(base_level)  # 0=特優(最高) ... 3=甲(最低)
    if style == "嚴格" and random.random() < 0.45:
        idx = min(idx + 1, len(GRADE_LEVELS) - 1)
    elif style == "寬鬆" and random.random() < 0.45:
        idx = max(idx - 1, 0)
    return GRADE_LEVELS[idx]


CLOSING_FILLERS = [
    "整體而言對團隊貢獻穩定,值得肯定", "後續仍會持續追蹤其成長幅度",
    "與跨部門協作時展現一定的專業度", "在既有職掌範圍內均能如期完成交辦事項",
    "對部門整體績效有實質貢獻",
]


def _grow_to_range(parts_pool, prefix, min_len, max_len, max_parts=6):
    """從片語庫逐句疊加,直到落在 [min_len, max_len] 區間,避免生硬截斷。"""
    pool = parts_pool.copy()
    random.shuffle(pool)
    text = prefix
    used = 0
    while len(text) < min_len and used < max_parts and pool:
        text += pool.pop() + "。"
        used += 1
    if len(text) > max_len:
        # 超過上限時,從句號邊界截斷,不要切在字中間
        cut = text.rfind("。", 0, max_len)
        text = text[:cut + 1] if cut > 0 else text[:max_len]
    return text


def build_comment(name, grade_level):
    strengths = random.sample(STRENGTH_PHRASES, k=2)
    weakness = random.choice(WEAKNESS_PHRASES)
    recommendation = random.choice(RECOMMENDATION_PHRASES)
    prefix = (
        f"{name}於本考核年度{strengths[0]},{strengths[1]}。"
        f"惟{weakness}。{recommendation}。"
    )
    extra_pool = [p for p in (STRENGTH_PHRASES + CLOSING_FILLERS) if p not in strengths]
    return _grow_to_range(extra_pool, prefix, min_len=100, max_len=150)


def build_perf_summary(name, landing_score, grade_level):
    opening = PERF_OPENING_BY_GRADE[grade_level].format(name=name)
    body_parts = random.sample(STRENGTH_PHRASES, k=2) + [random.choice(WEAKNESS_PHRASES)]
    body = "、".join(body_parts[:-1]) + f"；但{body_parts[-1]}。"
    tail = random.choice(RECOMMENDATION_PHRASES) + "。"
    prefix = f"{name}（{landing_score}）該員{opening[len(name):]}{body}{tail}"
    extra_pool = [p for p in (STRENGTH_PHRASES + CLOSING_FILLERS) if p not in body_parts]
    return _grow_to_range(extra_pool, prefix, min_len=130, max_len=150)


def random_rotation_date(start_year, end_year):
    year = random.randint(start_year, end_year)
    month, day = random.choice(ROTATION_MONTHS_DAYS)
    return date(year, month, day)


# ---------------------------------------------------------------------------
# 1) 員工基本資料
# ---------------------------------------------------------------------------

employees = []
CURRENT_YEAR = 2025

for i in range(1, N_EMPLOYEES + 1):
    emp_id = f"E{i:05d}"[:6] if len(f"E{i:05d}") >= 6 else f"E{i:05d}"
    emp_id = f"{i:06d}"  # 6 碼文字代碼

    group_tenure = max(1, int(random.gauss(8, 4)))
    group_tenure = min(group_tenure, 35)
    dept_tenure = random.randint(1, group_tenure)

    age = max(24, min(60, 24 + group_tenure + random.randint(0, 8)))

    # 職等與年資正相關,但保留隨機性
    grade = min(14, max(1, int(3 + group_tenure * 0.6 + random.randint(-2, 2))))
    title = pick_grade_title(grade)

    dept = DEPARTMENT_FILTER if DEPARTMENT_FILTER else random.choice(DEPARTMENTS)

    school = random.choice(SCHOOLS)
    major = random.choice(MAJORS)
    edu = f"{school}{major}"
    if random.random() < 0.35:
        edu += f"；{random.choice(SCHOOLS)}{random.choice(GRAD_SCHOOLS)}"

    if random.random() < 0.4:
        prev_company = random.choice(OTHER_COMPANIES)
        prev_title = random.choice(["專員", "副理", "經理"])
        prev_years = random.randint(1, 5)
        prev_exp = f"{prev_company}{prev_title}（{prev_years}年）"
    else:
        prev_exp = "-"

    gender = random.choice(["male", "female"])
    name = fake.name_male() if gender == "male" else fake.name_female()

    manager_style = random.choice(MANAGER_STYLES)

    employees.append({
        "員編": emp_id,
        "姓名": name,
        "年齡": age,
        "現職公司": HOME_COMPANY,
        "現職單位/職位": dept,
        "職等": grade,
        "職稱": title,
        "集團年資": group_tenure,
        "部門年資": dept_tenure,
        "學經歷": edu,
        "前職經歷": prev_exp,
        # 以下兩欄不在原始規格的「介面欄位」裡,但生成其他表需要,先保留在工作表後面方便對照;
        # 如果不需要可以在匯出前刪除
        "_manager_style": manager_style,
        "_group_tenure": group_tenure,
    })

df_employees_full = pd.DataFrame(employees)
df_employees = df_employees_full.drop(columns=["_manager_style", "_group_tenure"])

# ---------------------------------------------------------------------------
# 2) 考績歷史(依集團年資決定有幾年資料,最多到 2023-2025)
# ---------------------------------------------------------------------------

performance_rows = []
for _, emp in df_employees_full.iterrows():
    years_available = min(3, emp["_group_tenure"])
    years = list(range(CURRENT_YEAR - years_available + 1, CURRENT_YEAR + 1))
    for year in years:
        base_level = weighted_grade_level()
        final_level = apply_manager_style(base_level, emp["_manager_style"])
        comment = build_comment(emp["姓名"], final_level)
        landing_score = random.randint(55, 98)
        summary = build_perf_summary(emp["姓名"], landing_score, final_level)
        performance_rows.append({
            "員編": emp["員編"],
            "年度": year,
            "考績等第": final_level,
            "考績評語": comment,
            "表現敘述": summary,
        })

df_performance = pd.DataFrame(performance_rows)

# ---------------------------------------------------------------------------
# 3) 360評鑑(跟考績歷史同樣的年度,主管職才有「部屬平均」)
# ---------------------------------------------------------------------------

review_rows = []
MANAGER_TITLES = {"經理", "資深經理", "協理", "資深協理"}

for _, emp in df_employees_full.iterrows():
    years_available = min(3, emp["_group_tenure"])
    years = list(range(CURRENT_YEAR - years_available + 1, CURRENT_YEAR + 1))
    is_manager = emp["職稱"] in MANAGER_TITLES
    for year in years:
        self_score = round(random.uniform(6.5, 9.5), 1)
        peer_score = round(random.uniform(max(5.0, self_score - 1.5), min(10.0, self_score + 1.0)), 1)
        sub_score = round(random.uniform(max(5.0, self_score - 1.0), min(10.0, self_score + 1.2)), 1) if is_manager else None
        others = [peer_score] + ([sub_score] if sub_score is not None else [])
        other_avg = round(sum(others) / len(others), 1)
        diff = round(other_avg - self_score, 1)
        review_rows.append({
            "員編": emp["員編"],
            "年度": year,
            "個人自評": self_score,
            "同儕平均": peer_score,
            "部屬平均": sub_score if sub_score is not None else "",
            "他評平均": other_avg,
            "自評與他評差異": diff,
        })

df_review = pd.DataFrame(review_rows)

# ---------------------------------------------------------------------------
# 4) 輪調經歷(由到職往現在推進,最後一筆要對上員工目前的職等/部門)
# ---------------------------------------------------------------------------

rotation_rows = []
for _, emp in df_employees_full.iterrows():
    tenure = emp["_group_tenure"]
    start_year = CURRENT_YEAR - tenure
    n_events = min(4, max(1, tenure // 3 + 1))

    # 第一筆:到職或初任,可能來自其他銀行
    if random.random() < 0.3:
        first_company = random.choice(OTHER_COMPANIES)
        first_reason = "到職"
    else:
        first_company = HOME_COMPANY
        first_reason = "初任"

    start_grade = max(1, emp["職等"] - min(6, n_events))
    current_grade = start_grade
    event_year_span = max(1, tenure // max(1, n_events))

    for idx in range(n_events):
        year = min(CURRENT_YEAR, start_year + idx * event_year_span)
        event_date = random_rotation_date(year, year)
        if idx == 0:
            reason = first_reason
            company = first_company
        else:
            reason = random.choice(ROTATION_REASONS_LATER)
            company = HOME_COMPANY
            current_grade = min(emp["職等"], current_grade + random.randint(1, 2))

        if idx == n_events - 1:
            current_grade = emp["職等"]
            dept = emp["現職單位/職位"]
        else:
            dept = random.choice(DEPARTMENTS)

        title = pick_grade_title(current_grade)
        work_nature = random.choice(WORK_NATURE_BY_TITLE[title])

        rotation_rows.append({
            "員編": emp["員編"],
            "日期": event_date.isoformat(),
            "異動原因": reason,
            "公司": company,
            "部門": dept,
            "職等": current_grade,
            "職稱": title,
            "工作性質": work_nature,
        })

df_rotation = pd.DataFrame(rotation_rows)
df_rotation = df_rotation.sort_values(["員編", "日期"]).reset_index(drop=True)

# ---------------------------------------------------------------------------
# 輸出
# ---------------------------------------------------------------------------

import os
os.makedirs(OUT_DIR, exist_ok=True)

# 員編是「6碼文字代碼」(例如 000007),存成一般 CSV 字串即可。
# 注意:之後若用 pandas 讀回這些 CSV,要指定 dtype={"員編": str},
# 不然 pandas 會自動把它當成整數,吃掉開頭的 0(000007 會變成 7)。
df_employees.to_csv(f"{OUT_DIR}/employees.csv", index=False, encoding="utf-8-sig")
df_performance.to_csv(f"{OUT_DIR}/performance_history.csv", index=False, encoding="utf-8-sig")
df_review.to_csv(f"{OUT_DIR}/review_360.csv", index=False, encoding="utf-8-sig")
df_rotation.to_csv(f"{OUT_DIR}/rotation_history.csv", index=False, encoding="utf-8-sig")

xlsx_name = f"永豐考績假資料_{DEPARTMENT_FILTER}.xlsx" if DEPARTMENT_FILTER else "永豐考績假資料.xlsx"
with pd.ExcelWriter(f"{OUT_DIR}/{xlsx_name}", engine="openpyxl") as writer:
    df_employees.to_excel(writer, sheet_name="員工基本資料", index=False)
    df_performance.to_excel(writer, sheet_name="考績歷史", index=False)
    df_review.to_excel(writer, sheet_name="360評鑑", index=False)
    df_rotation.to_excel(writer, sheet_name="輪調經歷", index=False)

    # 把「員編」欄位強制設成文字格式,避免 Excel 自動轉成數字吃掉前導零
    from openpyxl.styles import numbers
    for sheet_name, df in [
        ("員工基本資料", df_employees), ("考績歷史", df_performance),
        ("360評鑑", df_review), ("輪調經歷", df_rotation),
    ]:
        ws = writer.sheets[sheet_name]
        col_idx = list(df.columns).index("員編") + 1  # openpyxl 是 1-based
        for row in range(2, len(df) + 2):
            ws.cell(row=row, column=col_idx).number_format = numbers.FORMAT_TEXT

print(f"輸出檔案: {OUT_DIR}/{xlsx_name}")
if DEPARTMENT_FILTER:
    print(f"部門篩選: 僅 {DEPARTMENT_FILTER}, {N_EMPLOYEES} 人")
print("員工基本資料:", df_employees.shape)
print("考績歷史:", df_performance.shape)
print("360評鑑:", df_review.shape)
print("輪調經歷:", df_rotation.shape)
print("主管風格分佈(用於製造標準落差,未輸出到介面欄位):")
print(df_employees_full["_manager_style"].value_counts())
