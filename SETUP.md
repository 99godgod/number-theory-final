# 下一台電腦完整使用步驟

本文件記錄如何在新電腦 clone 並使用本 repo（含 LaTeX 編譯、Python 投影片生成、git 同步流程）。

## Step 1:安裝 Git

| 作業系統 | 安裝方式 |
|---|---|
| **Windows** | 到 https://git-scm.com/download/win 下載安裝（內建 Git Credential Manager）|
| **macOS** | 終端機執行 `xcode-select --install`，或 `brew install git` |
| **Linux (Ubuntu)** | `sudo apt update && sudo apt install git` |

安裝後驗證:
```bash
git --version
```

## Step 2:Clone Repo（首次）

打開終端機 / PowerShell，切到想存放的位置（例如桌面）:
```bash
cd Desktop
git clone https://github.com/99godgod/number-theory-final.git
cd number-theory-final
```

**認證**（Private repo 必經）:
- **Windows / macOS**:會自動跳出瀏覽器要求登入 GitHub（Git Credential Manager），登入一次後永久免登入。
- **Linux**:可能需要 PAT (Personal Access Token)。流程:GitHub → Settings → Developer settings → Personal access tokens → 產生 token (勾 `repo` 範圍) → 把 token 當作密碼貼上。

## Step 3:設定 Git User（首次，每台電腦只需一次）

```bash
git config --global user.name "Chung-Chin Huang"
git config --global user.email "99godgod@gmail.com"
```

## Step 4:安裝編譯軟體（依需求）

### 只是要看 PDF
不需要任何額外軟體，雙擊 `.pdf` 即可。

### 要編輯 / 重新編譯 `.tex`
- **Windows**:安裝 MiKTeX (https://miktex.org/) → 自動下載缺少的套件
- **macOS**:安裝 MacTeX (https://www.tug.org/mactex/)
- **Linux**:`sudo apt install texlive-xetex texlive-lang-chinese texlive-fonts-extra`
- **字型**:需確保系統有 `Microsoft JhengHei`（Windows 內建）；其他平台可改 `Noto Sans CJK TC` 或 `PingFang TC`
- **編譯指令**:
  ```bash
  xelatex summary.tex
  xelatex presentation_script_v2.tex
  ```

### 要執行 `make_ppt_v2.py`
- 安裝 Python 3.10+（https://www.python.org/downloads/）
- 安裝套件:
  ```bash
  pip install python-pptx
  ```
- 執行:
  ```bash
  python make_ppt_v2.py
  ```

### 要編輯 `.pptx`
- PowerPoint、Keynote、或 LibreOffice Impress

## Step 5:日常使用流程

### 每次開始工作前（先抓最新版）
```bash
cd path/to/number-theory-final
git pull
```

### 修改後存檔
```bash
git add .
git commit -m "修改摘要的標題排版"
git push
```

### 常用查詢
```bash
git status              # 看哪些檔案被改過
git log --oneline       # 看 commit 歷史
git diff <檔名>         # 看某檔案改了什麼
```

## Step 6:回原本電腦時

原本那台若要拿到另一台改過的內容:
```bash
cd "D:/成大碩士班/114-2/數論（一）/期末報告"
git pull
```

## 一頁速查卡

```bash
# 首次（每台電腦執行一次）
git clone https://github.com/99godgod/number-theory-final.git
git config --global user.name "Chung-Chin Huang"
git config --global user.email "99godgod@gmail.com"

# 每次開始
git pull

# 每次結束
git add .
git commit -m "說明這次改了什麼"
git push
```

## 注意事項

- **永遠先 `git pull` 再開始改**:否則容易產生衝突
- **commit 訊息寫清楚**:未來自己看歷史會感謝自己
- **不要在兩台同時改同一個檔案**:若不慎發生，git 會提示 conflict，需手動編輯保留要的版本
- **大檔案**:若未來要加入超大檔案（> 100 MB），需使用 Git LFS（`git lfs install` + `git lfs track "*.pdf"`）
- **Private repo 轉 Public**:GitHub repo 頁面 → Settings → Danger Zone → Change repository visibility

## 衝突 (Conflict) 處理速記

若 `git pull` 出現 conflict:
1. `git status` 看哪些檔案有衝突
2. 用編輯器打開檔案，會看到類似:
   ```
   <<<<<<< HEAD
   這台電腦改的內容
   =======
   另一台電腦改的內容
   >>>>>>> origin/main
   ```
3. 編輯保留想要的版本，刪掉 `<<<<<<<` / `=======` / `>>>>>>>` 標記
4. `git add <檔名>`
5. `git commit -m "resolve conflict"`
6. `git push`

## 還原操作

| 想做的事 | 指令 |
|---|---|
| 還沒 add，丟棄修改 | `git checkout -- <檔名>` |
| 已 add 還沒 commit，取消 stage | `git reset HEAD <檔名>` |
| 已 commit 但還沒 push，回到上一個 commit | `git reset --soft HEAD^` |
| 已 push 想看舊版內容 | `git show <commit-hash>:<檔名>` |
