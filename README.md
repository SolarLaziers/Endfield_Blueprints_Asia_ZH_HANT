# Endfield_Blueprints_Asia_ZH_HANT

Original provenance note: the original README only contained the repo title.

## English

### Project overview

This repository manages the Valley Four workbook automation, verification, and bilingual usage guidance for the Asia Traditional Chinese workbook.

### Workbook overview

This workbook keeps the everyday user workflow separate from the source sheet and rebuild scripts.

### Which sheet to use

For normal workbook use, treat `四號谷地_視覺版` and `四號谷地_表格版` as the primary user workflow.

- `四號谷地_視覺版`: visual browsing view for the Valley Four data.
- `四號谷地_表格版`: main table view for daily review.
- `四號谷地_1`: source/backend worksheet that feeds the managed output.

### Filter and search

On `四號谷地_表格版`, use slicers and the built-in filters to narrow the generated table. Use `搜尋文字` for keyword search when you need to filter or search across mixed text values.

### Generated vs source

`四號谷地_表格版` and its related generated structures are managed output, while `四號谷地_1` is the source/backend sheet. Do not use the generated tables, slicers, or related structures as manual entry areas unless you plan to rebuild them again afterward.

### Automation

Rebuild the managed output table and slicers in write-enabled mode:

```powershell
pwsh -File .\scripts\apply_sihao_valley_table.ps1 -EnableWorkbookWrites
```

Optionally rebuild a different workbook with `-WorkbookPath`:

```powershell
pwsh -File .\scripts\apply_sihao_valley_table.ps1 -EnableWorkbookWrites -WorkbookPath ".\temp\Endfield Blueprints (Asia).xlsx"
```

Verify the workbook state afterward:

```powershell
python .\scripts\verify_sihao_valley_table.py
```

Optionally verify another workbook path, and optionally pass a baseline workbook as the second argument:

```powershell
python .\scripts\verify_sihao_valley_table.py ".\temp\Endfield Blueprints (Asia).xlsx" ".\baseline.xlsx"
```

Verification checks the source and output headers, row counts, and text integrity for workbook data.

### Provenance

Original provenance note: the original README only contained the repo title.

## 繁體中文

### 專案概覽

這個 repo 主要整理四號谷地活頁簿的自動化、驗證流程，以及提供給 Asia Traditional Chinese 活頁簿的雙語使用說明。

### 活頁簿概覽

這份活頁簿會把日常使用流程、來源工作表與重建腳本分開說明。

### 使用哪個工作表

一般使用時，請把 `四號谷地_視覺版` 與 `四號谷地_表格版` 視為主要使用流程。

- `四號谷地_視覺版`：用於瀏覽四號谷地資料的視覺版工作表。
- `四號谷地_表格版`：日常檢視資料時優先使用的表格版工作表。
- `四號谷地_1`：提供資料來源的來源/後端工作表。

### 篩選與搜尋

在 `四號谷地_表格版` 上，請使用 slicers 與內建篩選功能縮小資料範圍；需要關鍵字縮小結果時，請使用 `搜尋文字` 欄位來搜尋與篩選混合文字內容。

### 產生內容與來源

`四號谷地_表格版` 及其相關的產生結構屬於受控輸出，`四號谷地_1` 則是來源/後端工作表。除非你之後打算再次重建，否則不要把這些產生表格、slicers 或相關結構當成手動輸入區域。

### 自動化

以可寫入模式重建受控表格與 slicers：

```powershell
pwsh -File .\scripts\apply_sihao_valley_table.ps1 -EnableWorkbookWrites
```

若要改為重建其他活頁簿，可另外加上 `-WorkbookPath`：

```powershell
pwsh -File .\scripts\apply_sihao_valley_table.ps1 -EnableWorkbookWrites -WorkbookPath ".\temp\Endfield Blueprints (Asia).xlsx"
```

之後請驗證活頁簿狀態：

```powershell
python .\scripts\verify_sihao_valley_table.py
```

也可以驗證其他活頁簿路徑，並選擇性提供第二個 baseline 活頁簿參數：

```powershell
python .\scripts\verify_sihao_valley_table.py ".\temp\Endfield Blueprints (Asia).xlsx" ".\baseline.xlsx"
```

驗證會檢查來源與輸出標題、資料列數，以及文字內容完整性。

### 來源說明

補充來源資訊：最初的 README 只有 repo 標題。
