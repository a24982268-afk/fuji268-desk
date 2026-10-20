# 蹦火仔 接團台 — 給 Claude 的專案說明

富吉漁創（金山蹦火仔）內部用的接團工具。使用者是店家同事，**介面與回覆一律用繁體中文**，多數時候用手機操作。
要修改這個系統時，先載入 skill `desk-change`（`.claude/skills/desk-change/SKILL.md`）照流程做。

## 系統組成

| 部分 | 位置 | 怎麼上線 |
|---|---|---|
| 接團台網頁 | 本 repo `index.html`（單檔，原生 JS，無建置） | 合併到 `main` → GitHub Pages 自動更新：https://a24982268-afk.github.io/fuji268-desk/ |
| 後端 | 本 repo `apps-script/Code.gs`（Google Apps Script，綁在店家的「蹦火仔接團台」試算表） | **不會自動更新**：店家要把整份 `Code.gs` 貼到 Apps Script → 部署 → 管理部署作業 → 編輯 `AKfycbyl_kkW…` 那個部署 → 新版本 |
| 預約頁、旅客名單填寫頁 | repo `a24982268-afk/fuji268`（`index.html`、`pax.html`） | 合併到 `main` → https://a24982268-afk.github.io/fuji268/ |
| LINE 機器人 | `Code.gs` 的 webhook（`doPost` 帶 `?line=鑰匙`） | LINE 官方帳號 @FUJI268，Webhook 指向上面那個部署 + `?line=…` |

網頁裡的 `API_URL`（`…/AKfycbyl_kkW…/exec`）是正式部署的網址，三個網頁共用。只有一個正式部署，不要叫店家「新增部署作業」（網址會變）。

## 資料（試算表工作表）

`Code.gs` 的 `SHEETS` 定義欄位（key ↔ 中文標題）。欄位依標題比對，新增的欄位會自動補在最右邊，所以**加欄位不用動舊資料**。

- 團單（bookings）：狀態只有 待確認／已成單／取消；進度在 `marks`（導覽員已排定、已通知導覽員、已完成，各自勾選）。
  舊資料的狀態可能是「導覽員已排定」等字，前後端都視為已成單＋對應進度（`baseStatus`／`marksOf`／`marksOf_`）。
  `trip` 連到行程；`code` 團號；`lineUser`、`inbox`、`unread`、`confirm` 是 LINE 相關；`passengers` 一行一位「姓名｜證號｜生日｜電話」。
- 行程（trips）：同一天幾張團單合併、共用導覽員。合併／移出／解散／排導覽員都走後端 `trip` 動作一次完成。
- 船隻（boats）、船班（sailings）：一天一艘船一班；當日船位＝有開、沒停航的船上限加總，只算要搭船的遊程（遊程「搭船」欄，空白依名稱判斷）。
- 導覽員（guides）：`off` 請假日。遊程（tours）。設定、訊息範本、自動回覆、LINE 訊息（紀錄）。

Script Properties（機密，只在 Apps Script 裡）：`PW_HASH SALT SECRET SHEET_ID LINE_TOKEN LINE_KEY LINE_TESTERS LINE_LIVE PAX_KEY CODE_SEQ`。

## 後端 API（`Code.gs`）

- `doPost` JSON：`login`；`paxSubmit`（不需登入，靠名單金鑰）；其餘要 `token`：`load`、`manifest`、`linePush`、`lineConfirm`，
  以及寫入類 `add/update/delete/settings/saveTemplate/saveReplies/trip`——這些都經 `writeOp_`，前端平常用 `batch` 一次送多個。
- `doGet`：`?avail=1`（預約頁名額，只回 open/few/full/stop/none）、`?pax=…&k=…`（名單頁讀資料）。
- `loadAll()` 回傳全部資料並帶 `features: { batch: 1 }`；前端用「有沒有某個欄位／旗標」判斷後端是不是新版（例如 `hasSail`、`hasTrips`、`useBatch`、`templates`）。

## 一定要遵守

1. **前端要能在舊後端上運作**：網頁會比 `Code.gs` 先上線。新功能依後端回傳的旗標或欄位決定要不要出現，缺的時候顯示「後端還沒更新」並退回舊行為，不能壞掉或默默丟資料。
2. **改到 `Code.gs` 就要提醒店家部署**，並把最新整份 `Code.gs` 用檔案傳給使用者（他們用複製貼上更新）。
3. **機密不外流**：LINE token、`?line=` 鑰匙、Channel secret、名單連結金鑰、旅客證號不要寫進程式、commit 或回覆；使用者截圖露出時提醒更換。
4. **手機優先**：390px 寬也要好用；按鈕文字中文、簡短。危險動作（刪除、解散、傳送 LINE）要按兩次確認（`armed` 模式）。
5. **儲存體驗**：寫入先在畫面上立即生效（`write(action, extra, localFn)`），背景送出；不要做「等伺服器回來才更新畫面」的流程。
6. 試算表寫入一律經 `putRow`（會擋公式注入 `=`、依標題對欄）。
7. LINE 機器人預設「測試模式」（只回應 `LINE_TESTERS` 名單），`LINE_LIVE=1` 才對所有人；推播會用到 LINE 每月額度，能用 reply 就用 reply。

## 前端重點（`index.html`）

- 分頁：團單、船班、遊客資料（名冊、Word、Excel 匯出）、導覽員安排（行程）、遊程、設定（導覽員名冊、船隻、訊息範本）。
- `render()` 每 15 秒同步可能重畫：使用者正在打字的區塊要保留內容與游標（看 `paxDraft`、`tplDraft`、`renderTemplates` 的寫法）。
- Excel、Word（`.docx`）都在瀏覽器端用內建的 `zip()` 產生，不經後端。報到點名表不能含證號、生日月日、電話。

## 測試方式

用 skill 裡的工具：`node .claude/skills/desk-change/scripts/gas_sim.js` 模擬後端；`desk_mock.js` 用假後端開接團台並截圖。
本專案沒有單元測試框架，**每次修改都要實際跑過**：後端改動用模擬器、前端改動用瀏覽器（手機寬 390px），看 `pageerror` 為空、截圖確認版面。
