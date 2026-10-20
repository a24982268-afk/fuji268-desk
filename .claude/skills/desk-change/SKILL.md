---
name: desk-change
description: 修改蹦火仔接團台（index.html、apps-script/Code.gs）或預約網站（fuji268 repo）的標準流程：開分支、相容舊後端、用模擬後端與瀏覽器測試、開 PR 合併、需要時產出最新 Code.gs 給店家部署。Use whenever changing the 接團台 / 預約頁 / LINE 機器人 code in this project.
---

# 修改接團台

先讀 repo 根目錄的 `CLAUDE.md`（系統架構、資料欄位、必守規則）。以下是每次修改的流程。

## 1. 確認需求

- 使用者多半用截圖＋一句話描述。看不懂、或有會影響資料結構的選擇（例如「每項各自勾選」還是「依序累積」）時，用 AskUserQuestion 問 2–4 個選擇題再做，選項描述用店家聽得懂的話。
- 回覆一律繁體中文、口語、短句；不要丟程式術語。

## 2. 開分支

```bash
git fetch origin main
git checkout -B <指定的開發分支> origin/main
```
上一個 PR 已合併時，從最新 `main` 重開同名分支（必要時 force push），不要疊在已合併的歷史上。

## 3. 修改的原則

- **後端（`apps-script/Code.gs`）**
  - 新欄位：加進 `SHEETS` 的 `cols`（數字欄同時加進 `NUM`）。舊試算表會自動補欄，不用遷移。
  - 新的寫入動作：加在 `writeOp_`，這樣 `batch` 也會支援；不要在每個動作裡自己 `loadAll()`。
  - 會被前端判斷「後端是否新版」的東西，放進 `loadAll()` 回傳值（例如新集合，或 `features` 旗標）。
- **前端（`index.html`）**
  - 依後端旗標／欄位決定新功能是否出現；舊後端時顯示提示並退回舊行為（參考 `hasSail`、`hasTrips`、`useBatch`、`templates`）。
  - 寫入用 `write()`／`db.*`／`upd()`，傳入本地更新函式讓畫面立即變化。
  - 用現有的樣式類別（`card`、`gbox`、`pill okp/lack`、`ghost`、`primary`、`acts`），CSS 變數支援深色模式。
  - 危險動作用 `armed` 兩段式確認。
- 狀態相關一律用 `baseStatus(r)`、`hasMark(r, …)`、`markData(r, …)`，不要直接比對 `r.status === "導覽員已排定"` 這類舊值。

## 4. 測試（每次都要跑）

工具在 `scripts/`。瀏覽器測試需要 playwright：沙盒裡 Chromium 已預裝在 `/opt/pw-browsers`，
在暫存目錄 `npm install playwright@1`，執行時設 `NODE_PATH=<那個目錄>/node_modules`。

- 後端：`node scripts/gas_sim.js apps-script/Code.gs`
  會先語法檢查，再跑冒煙測試（初始化、batch 新增／更新、行程合併、LINE 測試模式、名額查詢）。
  自己的情境：`const sim = require('./scripts/gas_sim.js').load('apps-script/Code.gs'); sim.run("tripOp_({...})")`。
- 前端：`node scripts/desk_mock.js --data data.json --view 導覽員安排 --shot out.png [--width 390] [--latency 1500]`
  用假後端（支援 load、batch、add/update/delete、trip）開 `index.html`，印出 `pageerror` 並截圖；
  也可 `require('./scripts/desk_mock.js').open({...})` 拿到 page 自己點按。
  - 舊後端相容：把資料裡的新欄位或 `features` 拿掉再跑一次，確認不會壞。
  - 用 Read 看截圖確認版面（手機寬 390px）。
- 產生的 Word/Excel：可用 `soffice --headless --convert-to pdf` 再 `pdftoppm` 轉圖檢查。

## 5. 上線

1. `git commit`（訊息中文，說明改了什麼），`git push -u origin <分支>`。
2. 開 PR（中文標題與內容：內容、測試、上線注意）。使用者同意後合併；修正剛上線功能的小錯誤可直接合併後告知。
3. **有改 `Code.gs`**：從 `main` 取出最新整份，語法檢查後用 SendUserFile 傳給使用者，附上部署步驟：
   - 貼進 Apps Script 取代全部 → 儲存
   - 部署 → 管理部署作業 → 編輯 `AKfycbyl_kkW…` → 版本選「新版本」→ 部署
   - 新增 Google 權限時，在編輯器執行一次「測試LINE連線」或「初始化」完成授權
4. 回覆使用者：改了什麼（用他們看得到的畫面描述）、要不要更新後端、怎麼確認成功。

## 常見坑

- `ScriptApp.getService().getUrl()` 從選單執行會回傳測試部署網址，LINE 打不進去；不要拿來給使用者當 Webhook。
- `UrlFetchApp` 需要「連線到外部服務」授權；新增會用到外部連線的功能時，提醒在編輯器跑一次函式授權。
- 不要寫「等後端回來再一張張存」的流程，Google 每次寫入加讀取要 2–5 秒；用一次送出的動作（參考 `tripOp_`、`batch`）。
- 15 秒同步會重畫畫面：正在編輯的輸入框要保留草稿與焦點。
