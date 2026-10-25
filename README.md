# IE 人力需求小程式

IE 部門每月月底計算未來六個月人力需求的小工具。一步一步做，目前完成第 1 到第 10 步（第 10 步 RMA人力需求還有幾列沒填）。

名詞：**行** = 直的（A、B、C…），**列** = 橫的（1、2、3…）。

## 怎麼開啟

- **最簡單**：點兩下 `index.html`，用 Chrome 或 Edge 開啟。
- **要安裝成 App、離線使用**：點兩下 `啟動小程式.command`，瀏覽器會開啟 `http://localhost:8765`。用完把跳出來的終端機視窗關掉。

第一次使用需要連上網路，小程式會載入 ExcelJS（讀寫 Excel 的元件）。
要在不能連外部網站的電腦上用，請看 `vendor/README.txt`。

## 第 1 步：篩選 MPS 報表

1. 上傳 MPS 的 Excel（只支援 `.xlsx`；舊的 `.xls` 請先在 Excel 另存成 `.xlsx`）。
2. 選取條件，兩種方式二選一（點哪一種的輸入框就會切到那一種）：
   - **依 A 行（原本的方式）**：輸入 A 行等於多少。可以輸入多個值，例如 `1, 3`。
   - **指定某一行，等於或不等於某些值**：輸入哪一行（例如 `F`）、選「等於」或「不等於」、輸入值（例如 `EOL, 停產`）。選「不等於」時：
     - 「排除這一行空白的列」（預設勾選、會記住）：那一行是空白的列不要，用來拿掉分類之間的「x」「TTL (…)」小計列。被排除的列會逐列列出（列號＋內容），取消勾選的那一列會保留；也可以按「全部排除」「全部保留」。逐列的勾選不記住，因為每個月列號不同。
     - 不勾時，那一行是空白的列會保留。
     - 不論勾不勾，要保留的行全都空白的列（空行）都不會留。
   符合的列才會留下，從第 5 列開始比對。報表裡那一行出現過的值會列在下面，點一下就加入或移除。
3. 輸入要保留的行，例如 `B, C, U, Z, AE`，連續的行可寫 `U:Z`；也可以直接點預覽表格上方的行號。
4. 按「匯出 Excel」，得到 `原檔名_篩選.xlsx`。

匯出的檔案：

- 第 1–4 列表頭原樣保留，包含格式、行寬、列高、凍結窗格。
- 合併儲存格如果只保留了其中幾行，會縮成只合併那幾行；大標題如果在沒保留的行，會移到左上角。
- 公式會換成計算結果的數值（篩掉列和行之後，原本的公式參照會錯位）。
- 所有儲存格的填滿色彩都會拿掉（變透明），包括表頭。原報表有些列塗了底色，留著的話會蓋過後面標的前三大淺橘色。字型、框線、數字格式照舊。
- 選的方式和輸入的條件都會記住（共用設定），下個月打開還在。

報表版面不同時（表頭不是 4 列，或選取條件不在 A 行），在第 3 項下面的「報表版面不一樣時」調整。

## 第 2 步：匯總

算出 MPS 各月的預計總產量（Total）、前三大料號的合計（Top 3）、其餘的合計（Non-Top 3）。

1. 第 1 步的 Excel：剛做完第 1 步的話會直接沿用；也可以上傳之前第 1 步匯出的檔案。
2. 確認當月 WIP 在哪一行、各月產量在哪幾行（預設是第 4 行和第 5 到 10 行；第一個是當月）。
3. 按「匯出 Excel」，得到 `…_篩選_匯總.xlsx`。

結果（以 WIP 在 D 行、當月在 E 行、資料在第 5 到 74 列為例）：

| 列 | 當月（要加 WIP） | 其他月份 |
| --- | --- | --- |
| Total | `=SUM(D5:D74)+SUM(E5:E74)` | `=SUM(F5:F74)` |
| Top 3 | `=SUMPRODUCT(LARGE(D5:D74+E5:E74,{1,2,3}))` | `=LARGE(F5:F74,1)+LARGE(F5:F74,2)+LARGE(F5:F74,3)` |
| Non-Top 3 | `=E76-E77` | `=F76-F77` |

- 位置：最後一列資料之後空一列開始，共三列，對齊各月的行；文字標籤在第一個月左邊那一行（通常就是 WIP 那一行）。
- 當月一律先加 WIP：總產量、前三大、標示都是用「WIP + 當月產量」來算。
- Total 那一列是淺綠色底（`#B5E48C`）；標籤加上各月圍成一塊，外框粗線、裡面細格線。
- 各月前三大的儲存格用條件式格式標成淺橘色，數值是 0 的不算。
  當月的規則是 `=AND($D5+E5>0,SUMPRODUCT(--($D$5:$D$74+E$5:E$74>$D5+E5))<3)`（「WIP + 當月」比它大的不到 3 個），
  其他月份是 `=AND(F5>0,F5>=LARGE(F$5:F$74,3))`。
- 料號固定看第 1 行（用來判斷資料到第幾列）。
- 這三列和第 6 步的需求工時匯總在同一列上，一個在左邊（產量）、一個在右邊（工時），互不影響。

## 第 3 步：匯入 P003 標準工時

把 P003 報表（公司所有產品型號的標準工時）放進第 1 步完成的 Excel，成為同一個檔案裡名為 `P003` 的工作表。

1. 第 2 步的 Excel：剛做完第 2 步的話會直接沿用，不用先匯出；也可以上傳之前第 2 步匯出的檔案。跳過第 2 步的話，會沿用第 1 步篩選好的結果。
2. 上傳 P003 報表（`.xlsx`）。檔案裡有多個工作表時可以選要用哪一個。
3. 按「匯出 Excel」，得到 `…_匯總_含P003.xlsx`，裡面有原本的 MPS 工作表和 `P003` 工作表。

P003 工作表是整張照抄：值、格式、合併儲存格、行寬、列高、凍結窗格、自動篩選都會帶過去，公式換成數值。
上傳的檔案裡如果已經有 `P003` 工作表，會換成這次上傳的。

## 第 4 步：帶入標準工時 CT.

拿 MPS 工作表的料號到 P003 工作表找同一個料號，把那一列的 Labor ÷ Base Qty 帶回 MPS。
P003 的 Labor 是做 Base Qty 個的工時（Base Qty 有 1 也有 100），除過之後才是單件工時。

1. 第 3 步的 Excel：剛做完第 3 步的話會直接沿用；也可以上傳之前第 3 步匯出的檔案（要有 MPS 和 P003 兩個工作表）。
2. 選 MPS 的料號在哪一行。
3. 選 P003 的料號、Labor、Base Qty 各在哪一行。
4. 按「匯出 Excel」，得到 `…_含P003_CT.xlsx`。

結果：

- CT. 放在 MPS 工作表「原本最後一行之後空一行」的位置，表頭是 `CT.`。
- 填入的是公式，例如 `='P003'!M5/'P003'!K5`（Labor ÷ Base Qty），在 Excel 點儲存格就能看到數值是從 P003 哪兩格算出來的。
- 同一個料號在 P003 出現很多列時，取最上面那一列（和 VLOOKUP 一樣）；畫面上會列出各列算出來的 CT. 不一樣的料號。這一點之後可能再調整。
- Base Qty 是空白或 0 的料號，Excel 會顯示 `#DIV/0!`，畫面上會列出是哪些料號。
- P003 找不到的料號，CT. 留空白並標上淺紅色底，畫面上會列出是哪些料號。
- 料號比對不分大小寫、會去掉頭尾空白；`0012` 和 `12` 視為不同。
- 對已經有 `CT.` 的檔案重做這一步，會覆蓋原本那一行，不會再多一行。

## 第 5 步：計算需求工時

把每個料號各月的計畫產量乘上該列的 CT.，算出各月的需求工時。

前提是第 1 步保留的行照這個順序：前三行是料號、型號、生產方式，第 4 行是當月 WIP，
第 5 到第 10 行是當月和往後五個月的計畫產量。順序不同時可以在畫面上改。

1. 第 4 步的 Excel：剛做完第 4 步的話會直接沿用；也可以上傳之前第 4 步匯出的檔案。
2. 確認當月 WIP 在哪一行、各月產量在哪幾行（預設是第 4 行和第 5 到 10 行）。這組設定和第 2 步共用，在其中一步改了，另一步也會跟著變。CT. 會自動找表頭是 `CT.` 的那一行。
3. 輸入各月的表頭。第一個（當月）用下拉選單選 Jan～Dec，再按最後一個框後面的「自動帶出後面月份」，後面幾個月會依序填好（Dec 之後接 Jan）；後面的月份也可以自己改。程式不會去 MPS 的表頭找，月份也不會記住或上傳。
4. 按「匯出 Excel」，得到 `…_CT_需求工時.xlsx`。

結果：

- 從 MPS 工作表「原本最後一行之後空一行」開始，一個月一行。
- 每格都是公式：當月是 `=(D5+E5)*L5`（WIP + 當月產量，再乘 CT.），其餘各月是 `=F5*L5`。
- 表頭：第 2 列是跨所有月份置中的「需求工時」；第 3、4 列合併，寫你輸入的月份。
- 數值只顯示整數。儲存格裡仍保留小數，加總用的是完整的數值，所以合計可能和「把畫面上的整數加起來」差 1。
- 沒有 CT. 的料號一樣會填公式，結果是 0，並標上淺紅色底；畫面上會列出是哪些料號。
- 數值的單位和 CT. 相同（P003 的 Labor 單位是分鐘，所以這裡是分鐘），不換算。
- 對已經有「需求工時」的檔案重做這一步，會覆蓋原本的位置。

## 第 6 步：需求工時匯總

在需求工時下面（最後一列資料之後空一列）加三列，每格都是公式：

| 列 | 公式（以 N 行、資料在第 5 到 74 列為例） |
| --- | --- |
| Production 總需求工時 | `=SUM(N5:N74)` |
| Top 3 總需求工時 | `=LARGE(N5:N74,1)+LARGE(N5:N74,2)+LARGE(N5:N74,3)` |
| Others 總需求工時 | `=N76-N77`（Production − Top 3） |

- 文字標籤放在需求工時左邊那一行（原本空著的那一行），置中；那一行會加寬到放得下標籤。
- 標籤加上各月共七行、三列圍成一塊：外框粗線、裡面細格線。「Production 總需求工時」那一列是淺藍色底（`#A2D2FF`）。
- 每個月前三大的儲存格會標成淺橘色，數值是 0 的不算。用的是 Excel 的「條件式格式」，
  規則是 `=AND(N5>0,N5>=LARGE(N$5:N$74,3))`（大於 0，而且不小於該月第三大的數值），
  所以之後手動補 CT.、數值變了，標示和這三列的數字都會跟著變。
- 排名是一個料號（一列）算一個。
- 對已經有匯總的檔案重做這一步，會覆蓋原本那三列。

## 第 7 步：RMA 需求工時

前面算的是正常生產（Production）的需求工時。這一步在匯總下面空一列，加上重工（RMA）的需求工時。

在畫面上一列輸入一個項目，列數可以自己增減。每個項目輸入名稱、CT.、各月的重工數量。

| 位置（以匯總在第 76 到 78 列為例） | 內容 |
| --- | --- |
| 第 80 列 | `CT.`（標題）、`RMA總需求工時`、各月 `=SUM(N81:N83)`（底下各項目的合計） |
| 第 81 列起，一個項目一列 | CT.（數字）、項目名稱、各月 `=120*L81`（重工數量 × 左邊的 CT.） |

- CT. 放在項目名稱左邊那一格。數量直接寫在公式裡；之後在 Excel 改 CT.，工時會跟著變。
- 外觀和匯總那一塊一樣：外框粗線、裡面細格線；第一列的標籤置中，項目名稱靠左。「RMA總需求工時」那一列是淺藍色底（`#A2D2FF`），左邊的 `CT.` 標題不上色。
- 項目名稱和 CT. 會記在這台電腦的瀏覽器裡，下次打開還在；數量每次都要重新輸入。
- 對已經有 RMA 的檔案重做這一步，會把原本那一塊換掉。

## 第 8 步：匯入人力報表

把人力報表（未來六個月的人力預計）放進前面做好的 Excel，成為同一個檔案裡名為 `人力` 的工作表，
並確認之後總表要抓的資料在哪裡。

1. 第 7 步的 Excel：剛做完第 7 步的話會直接沿用；也可以上傳之前第 7 步匯出的檔案。跳過第 7 步的話，會沿用第 6 步的結果。
2. 上傳人力報表（`.xlsx`）。檔案裡有多個工作表時可以選要用哪一個（有叫「人力」的會先選它）。
3. 確認六個項目各在第幾列：Online、Online（生產）、Online（重工）、Off line（直接）、Off line（間接）、TTL（On + Off直接）；
   以及當月（第一個月）在哪一行。六個月是從當月那一行開始往右連續六行。
4. 按「匯出 Excel」，得到 `…_RMA_人力.xlsx`。

- 選單裡每一列會顯示「第 N 列　這一列的文字」。程式會先依文字帶入（例如有 `Online` 和 `生產` 的那一列就是 Online（生產）），請對照預覽確認，不對就改。
- 當月那一行：選好的各列裡的數字剛好落在六個連續的行時才會自動帶入；報表連過去的月份也列出來時，程式不知道哪一行是當月，要自己選。
- 預覽表格裡螢光的格子，就是之後總表要抓的資料；下面的核對表列出六個項目 × 六個月讀到的數值，空白或不是數字的格子標成淺紅色。
- 兩個項目選了同一列、或有空白的格子，會提醒但不擋匯出；六個項目和當月都要選好才能匯出。
- 六個項目都選好時，各項目的列會記在這台電腦的瀏覽器裡；下個月列的文字對不上時，就用記住的列當預設值。
- `人力` 工作表是整張照抄（值、格式、合併儲存格、行寬、列高），公式換成數值。上傳的檔案裡如果已經有 `人力` 工作表，會換成這次上傳的。
- 確認的位置放在 `IEApp.state.withManpower`（`rows`：各項目的列號，`firstMonthCol`：當月的行號，`monthCount`：6），給之後的總表用。

## 第 9 步：總表

在同一個檔案裡新增 `總表` 工作表，建立表格的樣式、格式、外框格線，並用公式填入已經做好的列；其他列之後再補上。

1. 第 8 步的 Excel：剛做完第 8 步的話會直接沿用；也可以上傳之前第 8 步匯出的檔案。
2. 按「匯出 Excel」，得到 `…_人力_總表.xlsx`。

版面（B 到 J 行、第 2 到 24 列）：

| 位置 | 內容 |
| --- | --- |
| 第 2 列 B:C | 法定加班上限46hr/人（紅底白字） |
| 第 3 列 | 表頭，深藍底（`#366092`）白字：B:C 合併畫斜線；D 到 I 是六個月；J 是備註 |
| 第 4 列 | MPS (Production)，淺藍底（`#DCE6F1`） |
| 第 5–11 列 | 需求工時（Top 3、Non-Top 3、重工）、平日／假日工作天數、工作時間、加班時間；C 行是單位 |
| 第 12–14 列 | 綜合效率%（Top-3、Non-Top 3 機種）、出勤率%（B:C 合併） |
| 第 15–18 列 | 需求人力：On line、Off line_直接、Off line_間接、小計（小計淺綠底 `#EBF1DE`） |
| 第 19–22 列 | 現有人力(早+夜)：同上，小計(on+off) |
| 第 23 列 | 人力需求差異(現有-需求) |
| 第 24 列 | Loading Rate%(含加班)，淺橘底（`#FFCC99`），上面是粗線 |

已經填好的列（都是公式，參照 MPS 工作表；用 MPS 裡的標籤文字找出是哪一列，只顯示整數）：

| 總表的列 | 抓 MPS 工作表的哪一列 | 例子 |
| --- | --- | --- |
| MPS (Production) | 第 2 步匯總的 `Total`（當月已含 WIP） | `='MPS'!E76` |
| 需求工時(Top 3) | 第 6 步的 `Top 3 總需求工時` | `='MPS'!N77` |
| 需求工時(Non-Top 3) | 第 6 步的 `Others 總需求工時` | `='MPS'!N78` |
| 需求工時（重工） | 第 7 步的 `RMA總需求工時` | `='MPS'!N80` |

在畫面上輸入的列，以及由它們算出來的列：

| 總表的列 | 內容 |
| --- | --- |
| 平日工作天數 | 各月自己輸入（每個月不同，不會記住） |
| 假日工作天數 | 公式 `=(46-D8*2)/8`：（法定加班上限 46 小時 − 平日工作天數 × 每天加班 2 小時）÷ 每天 8 小時 |
| 工作時間 | 各月自己輸入，預設 480 分 |
| 加班時間 | 輸入每日加班幾小時、加班出勤人數 %，填公式 `=2*60*80%`；備註欄寫「每日加班 2 小時 × 60 × 加班出勤人數 80%」 |
| 綜合效率%（Top-3 機種） | 各月自己輸入，預設 80% |
| 綜合效率%（Non-Top 3 機種） | 各月自己輸入，預設 95% |
| 出勤率% | 各月自己輸入，預設 95% |
| 需求人力 On line | 公式 `=D5/(D8+D9)/(D10+D11)/D12/D14+D6/(D8+D9)/(D10+D11)/D13/D14+D7/D8/D10/D13/D14`（Top 3 + Non-Top 3 + 重工，見下方說明），顯示到小數一位 |
| 需求人力 Off line_直接 | 輸入文書、物料、材製損、包裝、領班的人數，填公式 `=2+3+1+4+1`；備註寫「文書 2、物料 3、…」 |
| 需求人力 Off line_間接 | 輸入物料、領班的人數，同上 |
| 需求人力 小計 | 公式 `=D15+D16`（On line + Off line_直接）；備註寫「不含Off line_間接」 |
| 現有人力 On line | 公式抓「人力」工作表第 8 步確認的 Online 那一列，例如 `='人力'!D4` |
| 現有人力 Off line_直接 | 抓「人力」工作表 Off line（直接）那一列，例如 `='人力'!D7` |
| 現有人力 Off line_間接 | 抓「人力」工作表 Off line（間接）那一列，例如 `='人力'!D8` |
| 現有人力 小計(on+off) | 公式 `=D19+D20`；備註寫「不含Off line_間接」 |
| 人力需求差異(現有-需求) | 公式 `=D22-D18`（現有人力小計 − 需求人力小計），顯示到小數一位 |
| Loading Rate%(含加班) | 公式 `=D22/D18`（現有人力小計 ÷ 需求人力小計），顯示成 % |

- 工作時間、綜合效率、出勤率、每日加班小時、加班出勤人數 % 會記在這台電腦的瀏覽器裡，下次打開還是上次輸入的值。
- 需求人力 On line 的三段：
  - Top 3：需求工時(Top 3) ÷（平日 + 假日工作天數）÷（工作時間 + 加班時間）÷ 綜合效率%（Top-3）÷ 出勤率%
  - Non-Top 3：需求工時(Non-Top 3) ÷（平日 + 假日工作天數）÷（工作時間 + 加班時間）÷ 綜合效率%（Non-Top 3）÷ 出勤率%
  - 重工：需求工時（重工）÷ 平日工作天數 ÷ 工作時間 ÷ 綜合效率%（Non-Top 3）÷ 出勤率%
  平日工作天數沒填的月份會是 `#DIV/0!`。
- 現有人力的六個月從第 8 步選的「當月」那一行開始往右六行。剛做完第 8 步就用它確認的位置；上傳第 8 步匯出的檔案時，用第 8 步上次確認、記在瀏覽器裡的位置（畫面上會標明）。
- Off line 的人數每個月一樣，也會記住上次輸入的值；人員種類在 `js/excel-io.js` 的 `REPORT_OFFLINE`。
- 百分比輸入 80 或 80% 都可以，Excel 裡存成 0.8，顯示成 80%。
- 平日工作天數沒填時，假日工作天數照 0 天算（畫面會提醒）；加班的兩個數字都填了才會填加班時間。

- 標籤只在 MPS 資料開始的那一列以下找（避免找到表頭裡的字）；同一個標籤出現很多次時取最下面那一個。
- 數值取標籤右邊的格子，由左往右六個。找不到標籤（例如沒做第 7 步）或不到六個月時，畫面會提醒，那幾格留空白。

- 字型全部是微軟正黑體（表格外的空白格子也是）。
- 外框粗線、裡面細線；第 15 列以下的字是深藍色（`#000080`）。
- C 行加寬到放得下「Off line_直接」「小計(on+off)」；太窄的話字會跑進 D 行（第一個月），看起來像跨欄置中。
- 月份表頭用公式參照 MPS 工作表「需求工時」的月份（第 5 步輸入的），例如 `='MPS'!N3`；找不到時留空白。
- 版面的設定都在 `js/excel-io.js` 的 `REPORT`、`REPORT_COLOR`、`REPORT_ROWS`，各列的 key（例如 `hoursTop3`）之後填內容時用。
- 對已經有 `總表` 的檔案重做這一步，會把總表重新建立。

### 第二張表：YO 加班上限

在第一張表（第 2–24 列）下面空三列，第 28–50 列再放一張一樣的表：

- 紅色格子寫「YO加班上限78hr/人」。78 是預設值，可以在畫面上的「YO 加班上限」改（會記住），紅色格子和公式一起變。
- 假日工作天數：`=(78-D34*2)/8`（D34 是第二張表的平日工作天數）。
- 其他輸入和抓來的數值都參照第一張表，例如 D34 `=D8`、D30 `=D4`、月份表頭 D29 `=D3`；上面改了下面會跟著變。
- 需求人力 On line、兩個小計、人力需求差異、Loading Rate 用第二張表自己的數字算（因為假日工作天數不同，結果也不同）。

## 第 10 步：RMA人力需求

在同一個檔案裡新增 `RMA人力需求` 工作表，建立表格的樣式、格式、外框格線，並用公式填入已經做好的列。

1. 第 9 步的 Excel：剛做完第 9 步的話會直接沿用；也可以上傳之前第 9 步匯出的檔案。
2. 確認右上角的 Cut-off 日期（預設今天）和版次（預設 A，會記住上次輸入的值）。
3. 按「匯出 Excel」，得到 `…_總表_RMA人力.xlsx`。

版面（B 到 J 行、第 3 到 13 列）：

| 位置 | 內容 |
| --- | --- |
| 第 3 列 | B:C「法定加班上限46hr/人」（紅底白字置中）；J「Cut-off: 2026/9/29Ver:A」 |
| 第 4 列 | 表頭，深藍底白字：B:C 合併畫斜線；D 到 I 是六個月（參照總表的月份，例如 `='總表'!D3`）；J 是備註（靠左） |
| 第 5 列 | MPS (PCS)，淺藍底 |
| 第 6–10 列 | 需求工時、平日／假日工作天數、工作時間、加班時間；C 行是單位 |
| 第 11–12 列 | 綜合效率%、出勤率%（B:C 合併） |
| 第 13 列 | 需求人力 On line，淺綠底 |

已經填好的列（都是公式）：

| 列 | 內容（D 行為例） |
| --- | --- |
| MPS (PCS) | 第 7 步各 RMA 項目當月數量的合計，例如 `=120+300` |
| 需求工時 | `='MPS'!N80`（MPS 工作表「RMA總需求工時」那一列） |
| 平日工作天數 | `='總表'!D8`（第 9 步輸入的） |
| 工作時間 | `='總表'!D10` |
| 綜合效率% | `='總表'!D13`（Non-Top 3 機種） |
| 出勤率% | `='總表'!D14` |
| 需求人力 On line | `=D6/(D7+D8)/(D9+D10)/D11/D12`：需求工時 ÷（平日 + 假日工作天數）÷（工作時間 + 加班時間）÷ 綜合效率% ÷ 出勤率%，顯示到小數點第二位 |

- RMA 的數量沒有單獨放在儲存格裡，而是寫在 MPS 工作表 RMA 項目的公式裡（`=120*L81`），程式從這些公式讀出數量。所以上傳第 9 步的檔案也能算。
- 假日工作天數、加班時間之後再補；還沒填的時候當成 0 算。

- 外框粗線、裡面細線，字型是微軟正黑體。版面設定在 `js/excel-io.js` 的 `RMA_LAYOUT`、`RMA_ROWS`。
- 對已經有 `RMA人力需求` 的檔案重做這一步，會重新建立。

### 已知的 ExcelJS 問題：整張表變成粗體＋斜體＋刪除線

P003 這類系統匯出的報表，字型設定會把「沒有斜體、沒有刪除線」明寫成 `<i val="0">`、`<strike val="0">`。
ExcelJS 4.x 只看到標籤就當成「有」，結果匯出後整張表被畫上刪除線。
`js/excel-io.js` 的 `fixFontFlags()` 會在讀檔前先把這些標籤拿掉，所有上傳的檔案都會經過這道修正。

## 檔案

```
index.html                  主畫面
js/config.js                專案代號、Supabase 網址和公開鑰匙（SUPABASE_PUBLIC_URL、SUPABASE_PUBLISHABLE_KEY）
js/cloud.js                 登入與共用設定（Supabase）
manifest.webmanifest        安裝成 App 用的資訊
sw.js                       離線快取
css/style.css               樣式
js/app.js                   外框：步驟清單、切換步驟、載入 ExcelJS
js/excel-io.js              所有 Excel 讀寫與排版（ExcelJS）
js/ui.js                    各步驟共用的畫面元件：上傳區、試算表預覽、下載，以及 WIP／月份行的共用設定
js/steps/step1-mps-filter.js    第 1 步　篩選 MPS 報表
js/steps/step2-mps-totals.js    第 2 步　匯總
js/steps/step3-p003-import.js   第 3 步　匯入 P003 標準工時
js/steps/step4-ct-lookup.js     第 4 步　帶入標準工時 CT.
js/steps/step5-demand-hours.js  第 5 步　計算需求工時
js/steps/step6-summary.js       第 6 步　需求工時匯總
js/steps/step7-rma.js           第 7 步　RMA 需求工時
js/steps/step8-manpower.js      第 8 步　匯入人力報表
js/steps/step9-report.js        第 9 步　總表
js/steps/step10-rma-manpower.js 第 10 步　RMA人力需求
vendor/supabase.js          supabase-js 2.117.3（連 Supabase 用，必要）
vendor/                     ExcelJS 本機副本（可選）
icons/                      圖示
啟動小程式.command           在 Mac 上啟動
.gitignore                  不進 git 的檔案（.DS_Store、.env）
```

新增步驟時：在 `js/steps/` 加一個檔案、在 `index.html` 加一行 `<script>`、
在 `sw.js` 的 `SHELL` 加上檔名並把 `CACHE` 版本號加 1。
檔名開頭的數字就是步驟編號；在中間插入新步驟時，後面的檔案要跟著改名。

每次改版：把 `js/excel-io.js` 裡的 `VERSION` 和 `index.html` 裡每個 `?v=` 的數字一起加 1。
畫面右上角會顯示「版本 N」，可以用來確認瀏覽器跑的是不是最新的程式。

每一步做完會把結果放在 `IEApp.state`（第 1 步是 `mpsFiltered`，第 2 步是 `mpsTotals`，第 3 步是 `withP003`，第 4 步是 `withCT`，第 5 步是 `withDemand`，第 6 步是 `withSummary`，第 7 步是 `withRma`，第 8 步是 `withManpower`，第 9 步是 `withReport`，第 10 步是 `withRmaManpower`），
裡面的 `build()` 會產生一份新的活頁簿，下一步拿來接著加工。

## 登入與共用設定（Supabase）

部門的人共用同一份「記住的設定」：誰最後存，大家打開看到的就是誰的。

- 打開時要先登入（email + 密碼）。帳號由管理者在 Supabase 後台建立（Authentication → Users → Add user），不開放自行註冊。
- 設定存在 Supabase 專案的 `ie_manpower.settings` 表：一筆是 `(app, key) → value`。
  - `app`：專案代號（`js/config.js` 的 `PROJECT_ID`）。8 個人力需求 PWA 共用這張表，用 `app` 分開。
  - `key`：「步驟.輸入項目」，例如 `step9.otHours` 是第 9 步的每日加班小時。組名一律用步驟編號（`step1`～`step10`），第 2、5 步共用的 WIP／月份行是 `step2-5`（同一個軟件的兩步共用，不是各產線共用）。
  - `value`：那個輸入項目的值（JSON）。6 個月、Off line 各種人數這種一起看的值放在同一筆（陣列）。
  - 一個輸入項目一筆，所以兩個人同時改同一步的不同項目，不會互相蓋掉；每個項目都記得是誰改的。
- 誰改的、什麼時候改的由資料庫自動記錄（`updated_by`、`updated_at`），各步驟標題下面會顯示「共用設定：上次由 某某 於 10/19 16:05 修改」。
- 右上角顯示登入的人、同步狀態（已同步／同步中／離線）和「登出」。
- 讀設定一律讀這台電腦的本機備份（localStorage），所以離線也能用；改了設定先存本機，再在背景上傳（停止輸入 1 秒後）。離線時改的會等連上網路自動上傳。
- 打開時先把 Supabase 上最新的設定抓下來蓋掉本機備份，才載入各步驟（所以 `index.html` 裡沒有各步驟的 `<script>`，改由 `js/app.js` 的 `STEP_FILES` 載入）。
- 第一次登入時，這台電腦原本有、但 Supabase 上還沒有的設定會傳上去。
- 舊版的組名（`totals`、`columns`、`step3`…`step6`、`manpower`、`report`、`rma-manpower`）第一次打開新版時會自動改成步驟編號；版本 27 測試時存進 Supabase 的舊格式資料（key 沒有「.」的那幾筆）也會自動轉成新格式再傳一次。舊格式那幾筆網頁不能刪，可以在 Supabase 的 Table Editor 手動刪掉。

3F 取像存進資料庫的輸入項目：

| key | 對應的輸入框 |
| --- | --- |
| `step1.mode`、`step1.values`、`step1.fCol`、`step1.fOp`、`step1.fValues`、`step1.fSkipBlank` | 第 1 步：用哪一種選取方式（`a` 依 A 行／`col` 指定某一行）、A 行等於多少、指定的行、等於（`eq`）或不等於（`ne`）、指定行的值、「不等於」時是否排除那一行空白的列 |
| `step1.cols`、`step1.startRow`、`step1.selCol` | 第 1 步：保留的行、資料起始列、「依 A 行」方式用的行（預設 A） |
| `step2.startRow` | 第 2 步：資料起始列（報表版面不一樣時） |
| `step2-5.wipCol`、`step2-5.monthCols` | 第 2、5 步共用：WIP 在哪行、各月在哪幾行 |
| `step4.keyCol`、`step4.refKeyCol`、`step4.refValueCol`、`step4.refDivCol`、`step4.startRow` | 第 4 步：MPS 料號、P003 料號／Labor／Base Qty 的行、資料起始列 |
| `step5.startRow`、`step6.startRow` | 第 5、6 步：資料起始列 |
| `step7.items`、`step7.startRow` | 第 7 步：RMA 項目名稱和 CT.、資料起始列 |
| `step8.rows`、`step8.lastFirstCol` | 第 8 步：六個項目在人力報表第幾列、上次的當月行 |
| `step9.workMinutes`、`step9.effTop3`、`step9.effNonTop3`、`step9.attendance` | 第 9 步：工作時間、綜合效率（Top-3／Non-Top 3）、出勤率，各 6 個月 |
| `step9.otHours`、`step9.otRate`、`step9.offDirect`、`step9.offIndirect`、`step9.yoLimit` | 第 9 步：每日加班小時、加班出勤人數 %、Off line 直接／間接人數、YO 加班上限 |
| `step10.version` | 第 10 步：Cut-off 版次 |
- 不記住的欄位（平日工作天數、RMA 數量、Cut-off 日期）照舊每次輸入，不會上傳。
- 程式在 `js/cloud.js`，用 supabase-js 套件連線（`vendor/supabase.js`，版本 2.117.3，放在專案裡，離線也載得到）。雲端、本機自架、公司主機上的 Supabase 都用同一套程式，換資料庫只改 `js/config.js` 的 `SUPABASE_PUBLIC_URL`、`SUPABASE_PUBLISHABLE_KEY`。各步驟用 `IEApp.settings.read('step9')`／`IEApp.settings.write('step9', 物件)` 讀寫一整步的設定，`cloud.js` 負責拆成一個個輸入項目上傳、下載時再組回去。
- Supabase 的網址和公開鑰匙在 `js/config.js`。公開鑰匙本來就是公開用的；安全靠登入和資料表的權限規則（只有登入的人能讀寫，不能刪除）。`SUPABASE_SECRET_KEY`（舊式名稱：`SERVICE_ROLE_KEY`）不可以放進程式。
- 取名原則：`js/config.js` 的設定名稱照自架 Supabase 的 `.env` 取，兩邊找得到同一個名字。

  | `js/config.js` | 對應自架 `.env` | 說明 |
  | --- | --- | --- |
  | `SUPABASE_PUBLIC_URL` | `SUPABASE_PUBLIC_URL` | Supabase 網址 |
  | `SUPABASE_PUBLISHABLE_KEY` | `SUPABASE_PUBLISHABLE_KEY`（舊式名稱：`ANON_KEY`） | 公開鑰匙，`sb_publishable_` 開頭 |
  | `SUPABASE_SCHEMA` | `PGRST_DB_SCHEMAS` 清單裡的其中一個 | 這個程式用的 schema（`ie_manpower`） |

  版本 33 以前叫 `SUPABASE_URL`、`SUPABASE_KEY`，版本 35 起只認上面的新名稱。
- 登入資訊由 supabase-js 存在這台電腦的 `ie-manpower-3F-image-lens._auth`，會自己換新憑證。版本 31 以前存的 `._session` 第一次打開時會交給 supabase-js，已經登入的人不用重新登入。
- 有登入過的人打開時，如果 4 秒內連不上 Supabase，就先用這台電腦的設定，連上之後自動同步。

### 換一個 Supabase（本機自架、公司主機）

每個新的 Supabase 都要做一次：

1. 執行下面的 SQL（Studio 的 SQL Editor）。
2. 讓 API 看得到 `ie_manpower` schema：
   - 雲端：Settings → Data API → Exposed schemas 加上 `ie_manpower`。
   - Mac 上的 supabase CLI：`supabase/config.toml` 的 `[api]` 裡 `schemas` 加上 `"ie_manpower"`，再 `supabase stop` / `supabase start`。
   - Docker 自架：docker 資料夾 `.env` 的 `PGRST_DB_SCHEMAS` 加上 `,ie_manpower`，再 `docker compose up -d` 重啟。
3. 關閉自行註冊：CLI 是 `config.toml` 的 `[auth] enable_signup = false`；Docker 是 `.env` 的 `DISABLE_SIGNUP=true`。
4. 在 Studio 的 Authentication → Users 新增同事的帳號（email + 密碼，勾 Auto Confirm）。
5. 改 `js/config.js` 的 `SUPABASE_PUBLIC_URL`、`SUPABASE_PUBLISHABLE_KEY`（各環境要填什麼寫在檔案開頭）。

搬到公司主機時要注意：

- 別台電腦連的網址要寫主機的 IP 或名稱，不能寫 `127.0.0.1`／`localhost`（那是指使用者自己的電腦）。
- 用 https 網址打開的程式（例如 GitHub Pages）不能連 http 的 Supabase（瀏覽器會擋）。程式和 Supabase 要都是 http（例如程式也放在公司主機上），或都是 https。
- 離線快取（Service Worker）和「安裝成 App」只在 https 或 `localhost` 打開時才有；用 `http://主機IP` 打開時程式照常能用，但沒有離線功能。

建立資料庫的 SQL（雲端 Supabase 已經執行過；Exposed schemas 要加上 `ie_manpower`，見上面第 2 點）：

```sql
create schema if not exists ie_manpower;
grant usage on schema ie_manpower to authenticated;
create table if not exists ie_manpower.settings (
  app text not null, key text not null, value jsonb not null,
  updated_at timestamptz not null default now(), updated_by text,
  primary key (app, key)
);
alter table ie_manpower.settings enable row level security;
grant select, insert, update on ie_manpower.settings to authenticated;
create policy "settings_read"   on ie_manpower.settings for select to authenticated using (true);
create policy "settings_insert" on ie_manpower.settings for insert to authenticated with check (true);
create policy "settings_update" on ie_manpower.settings for update to authenticated using (true) with check (true);
create or replace function ie_manpower.stamp_settings() returns trigger language plpgsql security definer set search_path = '' as $$
begin new.updated_at := now(); new.updated_by := auth.jwt() ->> 'email'; return new; end $$;
create trigger settings_stamp before insert or update on ie_manpower.settings
for each row execute function ie_manpower.stamp_settings();
```

## 專案代號與複製成新專案

這個專案的代號是 `ie-manpower-3F-image-lens`，寫在 `js/config.js`（網頁和 `sw.js` 都讀這個檔案）。

- 離線快取的名稱：`ie-manpower-3F-image-lens-v26`（代號 + 版本號，版本號在 `sw.js` 的 `VERSION`）。新版啟用時，只刪「這個代號、但不是目前版本」的舊快取，別的專案的快取不碰。
- 記住的設定（localStorage）：`ie-manpower-3F-image-lens.step1` 這種名稱，用 `IEApp.storeKey('step1')` 產生。
- 舊版用的名稱（`ie-manpower.xxx` 設定、`ie-manpower-vNN` 快取）第一次打開新版時會自動搬過來或刪掉，記住的設定不會不見。

複製成新專案時要改的地方：

1. `js/config.js`：`PROJECT_ID` 改成新專案的代號；`LEGACY_STORE_PREFIX`、`LEGACY_CACHE_PATTERN` 改成 `null`。
2. `manifest.webmanifest`：`name`、`short_name`（最好也換 `theme_color` 或圖示）。
3. `啟動小程式.command`：`PORT` 改成別的號碼（例如 8766），本機測試時兩個專案才不會是同一個網址。
4. `sw.js` 的 `VERSION`、`js/excel-io.js` 的 `VERSION`、`index.html` 的 `?v=` 可以從 1 重新開始。

## 尚未開發

- RMA人力需求的假日工作天數、加班時間
