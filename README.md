# 考績評定平台（HR Review AI）

這是一個社團專案，與永豐銀行（Bank SinoPac）合作，目標是打造一套 **AI 輔助的人資考績評核系統**。目前階段先用假資料做出可互動的網頁原型（prototype），用來驗證介面流程與後續跟銀行討論的方向。

> ⚠️ 本專案目前使用的所有員工資料皆為**程式產生的假資料**（姓名、考績、評語等皆非真實員工資訊），僅供介面展示與內部討論使用。

---

## 專案內容

目前原型涵蓋以下頁面：

- **登入頁**
- **考績清單**：依考績狀態（未完成／草稿／已送出／待修改／已完成）篩選、搜尋員工
- **主要考績作業**：單一員工的考績填寫工作區，含歷年考績分數、AI 摘要/對話區（目前為 UI 示意，尚未接真實 AI）
- **員工基本資料**：基本資料、職等/年資，以及輪調經歷時間軸
- **歷史考績圖表**：歷年考績等第與 360 評鑑趨勢圖（Chart.js）
- **考績狀態維護**：待開發（目前為空白頁面）

---

## 專案結構

```
hr-review-ai/
├── generate_fake_data.py        # 產生假資料（員工/考績歷史/360評鑑/輪調經歷）
├── build_dashboard.py           # 把假資料組進網頁樣板，產生最終 HTML
├── dashboard_template.html      # 網頁樣板（版面、樣式、互動邏輯），含資料佔位區塊
├── 考績評定平台.html             # 由上面兩支程式產出的最終網頁，雙擊即可在瀏覽器打開
├── output/                      # generate_fake_data.py 產生的假資料
│   ├── employees.csv            # 員工基本資料
│   ├── performance_history.csv  # 考績歷史
│   ├── review_360.csv           # 360 評鑑
│   ├── rotation_history.csv     # 輪調經歷
│   └── *.xlsx                   # 整合成單一 Excel 檔（供人工檢視）
├── requirements.txt             # Python 套件需求
└── .gitignore
```

---

## 環境設定

建議使用獨立的虛擬環境（conda 或 venv 皆可）。

```bash
conda create -n hr-review-ai python=3.11
conda activate hr-review-ai
pip install -r requirements.txt
```

---

## 使用方式

### 1. 產生假資料

```bash
python generate_fake_data.py
```

會在 `output/` 底下產生 4 個 csv，以及一個整合的 Excel 檔。預設可在腳本內調整：

- `N_EMPLOYEES`：產生筆數
- `DEPARTMENT_FILTER`：若只想產生單一部門的資料（例如 `"法令遵循處"`），設定此變數即可；設為 `None` 則產生全公司各部門資料。

### 2. 組裝成網頁

```bash
python build_dashboard.py
```

預設會讀取 `output/` 底下的 4 個 csv，並依據 `dashboard_template.html` 的版面，輸出 `考績評定平台.html`。

若資料或樣板放在別的路徑，可加參數：

```bash
python build_dashboard.py --data-dir output --template dashboard_template.html --out 考績評定平台.html
```

### 3. 打開網頁

直接用瀏覽器開啟 `考績評定平台.html` 即可（需要網路連線以載入 Chart.js CDN）。

> 💡 之後只要重新產生假資料（人數、部門不同），**重新跑一次 `build_dashboard.py` 就好**，不需要手動改網頁檔案。如果要改版面、樣式或互動邏輯，才需要去改 `dashboard_template.html`。

---

## 開發狀態

- [x] 登入頁
- [x] 考績清單
- [x] 主要考績作業（AI 區塊為 UI 示意）
- [x] 員工基本資料
- [x] 歷史考績圖表
- [ ] 考績狀態維護

---

## 加入協作

本 repo 為 Public，但未開放協作者（Collaborator）直接 push 權限。若要貢獻修改，請用 **Fork + Pull Request** 的流程：

1. 到這個 repo 頁面，點右上角 **Fork**，把專案複製一份到自己的 GitHub 帳號下
2. clone 自己的 fork 到本機：
   ```bash
   git clone <你自己 fork 的網址>
   cd hr-review-ai
   ```
3. 修改檔案，commit 後 push 回自己的 fork：
   ```bash
   git add .
   git commit -m "說明這次改了什麼"
   git push
   ```
4. 回到 GitHub 網頁，對**原始 repo** 發起 **Pull Request（PR）**，簡單描述改了什麼
5. 專案負責人（repo 擁有者）檢查後再決定 Merge

> 💡 開始改之前，建議先同步最新版本，避免自己的 fork 落後太多：
> ```bash
> git remote add upstream <原始 repo 網址>   # 第一次設定，只要做一次
> git fetch upstream
> git merge upstream/main
> ```

---

## 免責聲明

本專案所有員工姓名、考績紀錄、評語等資料皆由程式隨機生成，與任何真實人物或真實考績結果無關，僅用於專案介面展示與內部討論。
