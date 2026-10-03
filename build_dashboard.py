# -*- coding: utf-8 -*-
"""
build_dashboard.py — 把假資料組進「考績評定平台」網頁樣板，產生可以直接打開的 HTML。

使用方式：
    1. 先跑 generate_fake_data.py，產生 output/ 底下的 4 個 csv
       （employees.csv / performance_history.csv / review_360.csv / rotation_history.csv）
    2. 再跑這支腳本：
           python build_dashboard.py
       預設會讀 output/ 裡的 4 個 csv，輸出 考績評定平台.html 到目前資料夾。

    如果資料放在別的路徑，或想換輸出檔名，可以加參數：
           python build_dashboard.py --data-dir output --out 考績評定平台.html

重要：
    - dashboard_template.html 是網頁的「骨架」（登入頁、考績清單、主要考績作業、
      員工基本資料、歷史考績圖表，全部的版面/樣式/互動邏輯都在裡面），
      裡面有一段 /* __EMPLOYEE_DATA_BLOCK__ ... */ 佔位註解，
      這支腳本只會把假資料轉成 JS 格式塞進那個佔位處，不會動版面本身。
    - 如果要改版面、改欄位顯示方式，要改 dashboard_template.html；
      如果只是換一批假資料（人數、部門不同），只要重新跑這支腳本就好，不用找人重做網頁。
"""
import argparse
import csv
import json
import sys
from pathlib import Path

STATUS_CYCLE = ["未完成", "草稿", "已送出", "待修改", "已完成"]
PLACEHOLDER = "/* __EMPLOYEE_DATA_BLOCK__：由 build_dashboard.py 自動產生並填入，請勿手動編輯這段 */\n"

REQUIRED_FILES = ["employees.csv", "performance_history.csv", "review_360.csv", "rotation_history.csv"]


def read_csv(path):
    with open(path, encoding="utf-8-sig") as f:
        return list(csv.DictReader(f))


def build_data_block(data_dir: Path) -> str:
    employees = read_csv(data_dir / "employees.csv")
    perf_rows = read_csv(data_dir / "performance_history.csv")
    review_rows = read_csv(data_dir / "review_360.csv")
    rotation_rows = read_csv(data_dir / "rotation_history.csv")

    # 1) 考績清單頁用的員工清單；考績狀態依序循環分配（未完成/草稿/已送出/待修改/已完成）
    employees_out = []
    for i, e in enumerate(employees):
        employees_out.append({
            "id": e["員編"],
            "name": e["姓名"],
            "title": e["職稱"],
            "dept": e["現職單位/職位"],
            "status": STATUS_CYCLE[i % len(STATUS_CYCLE)],
        })

    # 2) 考績歷史，依員編分組，年度新到舊排序
    perf_by_emp = {}
    for r in perf_rows:
        perf_by_emp.setdefault(r["員編"], []).append({
            "year": int(r["年度"]),
            "grade": r["考績等第"],
            "comment": r["考績評語"],
            "summary": r["表現敘述"],
        })
    for emp_id in perf_by_emp:
        perf_by_emp[emp_id].sort(key=lambda x: -x["year"])

    # 3) 360 評鑑，依員編分組
    review_by_emp = {}
    for r in review_rows:
        review_by_emp.setdefault(r["員編"], []).append({
            "year": int(r["年度"]),
            "self": float(r["個人自評"]),
            "peer": float(r["同儕平均"]),
            "sub": float(r["部屬平均"]) if r["部屬平均"] not in ("", None) else None,
            "others": float(r["他評平均"]),
            "diff": float(r["自評與他評差異"]),
        })
    for emp_id in review_by_emp:
        review_by_emp[emp_id].sort(key=lambda x: -x["year"])

    # 4) 員工基本資料（主要考績作業頁、員工基本資料頁都會用到）
    profile_by_emp = {
        e["員編"]: {
            "age": int(e["年齡"]),
            "company": e["現職公司"],
            "dept": e["現職單位/職位"],
            "grade": int(e["職等"]),
            "title": e["職稱"],
            "groupTenure": int(e["集團年資"]),
            "deptTenure": int(e["部門年資"]),
            "edu": e["學經歷"],
            "prevExp": e["前職經歷"],
        }
        for e in employees
    }

    # 5) 輪調經歷，依員編分組，日期由舊到新排序
    rotation_by_emp = {}
    for r in rotation_rows:
        rotation_by_emp.setdefault(r["員編"], []).append({
            "date": r["日期"],
            "reason": r["異動原因"],
            "company": r["公司"],
            "dept": r["部門"],
            "grade": int(r["職等"]),
            "title": r["職稱"],
            "workNature": r["工作性質"],
        })
    for emp_id in rotation_by_emp:
        rotation_by_emp[emp_id].sort(key=lambda x: x["date"])

    parts = [
        "const EMPLOYEES = " + json.dumps(employees_out, ensure_ascii=False, indent=2) + ";",
        "const PROFILE_BY_EMP = " + json.dumps(profile_by_emp, ensure_ascii=False, indent=2) + ";",
        "const PERF_BY_EMP = " + json.dumps(perf_by_emp, ensure_ascii=False, indent=2) + ";",
        "const REVIEW_BY_EMP = " + json.dumps(review_by_emp, ensure_ascii=False, indent=2) + ";",
        "const ROTATION_BY_EMP = " + json.dumps(rotation_by_emp, ensure_ascii=False, indent=2) + ";",
    ]
    return "\n\n".join(parts) + "\n\n"


def main():
    ap = argparse.ArgumentParser(description="把假資料組進考績工作台 HTML 樣板")
    ap.add_argument("--data-dir", default="output", help="放 4 個 csv 的資料夾（預設 output）")
    ap.add_argument("--template", default="dashboard_template.html", help="HTML 樣板檔案")
    ap.add_argument("--out", default="考績評定平台.html", help="輸出的 HTML 檔名")
    args = ap.parse_args()

    data_dir = Path(args.data_dir)
    template_path = Path(args.template)

    missing = [name for name in REQUIRED_FILES if not (data_dir / name).exists()]
    if missing:
        print(f"錯誤：在「{data_dir}」裡找不到這些檔案：{', '.join(missing)}")
        print("請先執行 generate_fake_data.py 產生假資料，或用 --data-dir 指到正確的資料夾。")
        sys.exit(1)

    if not template_path.exists():
        print(f"錯誤：找不到樣板檔案「{template_path}」。")
        sys.exit(1)

    template = template_path.read_text(encoding="utf-8")
    if PLACEHOLDER not in template:
        print("錯誤：樣板裡找不到資料佔位區塊，樣板可能被改動過，無法自動組裝。")
        sys.exit(1)

    data_block = build_data_block(data_dir)
    output_html = template.replace(PLACEHOLDER, data_block)

    out_path = Path(args.out)
    out_path.write_text(output_html, encoding="utf-8")
    print(f"完成：已輸出 {out_path}（讀取自 {data_dir}/ 底下的 4 個 csv）")


if __name__ == "__main__":
    main()
