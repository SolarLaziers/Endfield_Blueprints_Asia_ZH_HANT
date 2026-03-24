# Sihao Valley Table Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Build a refreshable filtered Excel output sheet for `四號谷地_1` inside `Endfield Blueprints (Asia).xlsx` without changing source-sheet cells.

**Architecture:** Use PowerShell-based Excel desktop automation to add a dynamic named range, create a Power Query that loads a cleaned table to a managed output worksheet, and attach slicers to the loaded Excel Table. Reuse the managed query/table/sheet on reruns so refresh stays attached to one workbook surface, and validate both source-sheet immutability and workbook structure afterward with a lightweight Python checker.

**Tech Stack:** Windows PowerShell COM automation, Excel Power Query, Excel Tables and slicers, Python `openpyxl` for structural verification

---

## File Structure

- Create: `scripts/SihaoValleyTable.psm1` - pure helper functions for expected headers, deterministic naming, named-range formulas, and Power Query M generation
- Create: `scripts/apply_sihao_valley_table.ps1` - workbook orchestration script that opens Excel, validates source shape, creates the query/table/slicers, saves, and closes cleanly
- Create: `scripts/verify_sihao_valley_table.py` - post-run workbook verification for source-sheet immutability, sheet names, headers, row counts, and table structure
- Create: `tests/test_sihao_valley_helpers.ps1` - self-contained PowerShell tests for helper logic
- Modify: `README.md` - usage notes for running the automation and verifier
- Output artifact: `Endfield Blueprints (Asia).xlsx` - workbook updated in place with new query-backed sheet objects

### Task 1: Build the helper module and failing helper tests

**Files:**
- Create: `scripts/SihaoValleyTable.psm1`
- Create: `tests/test_sihao_valley_helpers.ps1`

- [ ] **Step 1: Write the failing helper tests**

```powershell
$modulePath = Join-Path $PSScriptRoot "..\scripts\SihaoValleyTable.psm1"
Import-Module $modulePath -Force

$expectedHeaders = @('時代','類別','藍圖名稱','藍圖代碼','提供者','備註')
if ((Get-SihaoExpectedHeaders) -join '|' -ne $expectedHeaders -join '|') {
    throw 'Expected header list mismatch'
}

$name = Get-DeterministicSheetName -BaseName '四號谷地_1_表格' -ExistingNames @('四號谷地_1_表格')
if ($name -ne '四號谷地_1_表格_2') {
    throw 'Deterministic suffix naming failed'
}

$m = New-SihaoQueryFormula -NamedRange '四號谷地_1_來源範圍'
if ($m -notmatch '搜尋文字') {
    throw 'Query formula must include 搜尋文字 column'
}
```

- [ ] **Step 2: Run the helper tests to verify they fail**

Run: `powershell -NoProfile -ExecutionPolicy Bypass -File "tests/test_sihao_valley_helpers.ps1"`
Expected: FAIL because the helper module and functions do not exist yet.

- [ ] **Step 3: Implement the minimal helper module**

```powershell
function Get-SihaoExpectedHeaders {
    @('時代','類別','藍圖名稱','藍圖代碼','提供者','備註')
}

function Get-DeterministicSheetName {
    param([string]$BaseName, [string[]]$ExistingNames)
    if ($ExistingNames -notcontains $BaseName) { return $BaseName }
    $index = 2
    while ($ExistingNames -contains "$BaseName`_$index") { $index++ }
    return "$BaseName`_$index"
}
```

- [ ] **Step 4: Expand the helper module with query and named-range builders**

```powershell
function New-SourceNamedRangeFormula {
    param([string]$SheetName)
    "=OFFSET('$SheetName'!$A$1,0,0,MAX(LOOKUP(2,1/('$SheetName'!$A:$A<>\"\"),ROW('$SheetName'!$A:$A)),LOOKUP(2,1/('$SheetName'!$B:$B<>\"\"),ROW('$SheetName'!$B:$B)),LOOKUP(2,1/('$SheetName'!$C:$C<>\"\"),ROW('$SheetName'!$C:$C)),LOOKUP(2,1/('$SheetName'!$D:$D<>\"\"),ROW('$SheetName'!$D:$D)),LOOKUP(2,1/('$SheetName'!$E:$E<>\"\"),ROW('$SheetName'!$E:$E)),LOOKUP(2,1/('$SheetName'!$F:$F<>\"\"),ROW('$SheetName'!$F:$F))),6)"
}

function New-SihaoQueryFormula {
    param([string]$NamedRange)
    @"
let
    Source = Excel.CurrentWorkbook(){[Name="$NamedRange"]}[Content],
    Promoted = Table.PromoteHeaders(Source, [PromoteAllScalars=true]),
    Filtered = Table.SelectRows(Promoted, each List.NonNullCount(Record.FieldValues(_)) > 0),
    Trimmed = Table.TransformColumns(Filtered, {{"時代", each if _ is text then Text.Trim(_) else _, type text}, {"類別", each if _ is text then Text.Trim(_) else _, type text}, {"藍圖名稱", each if _ is text then Text.Trim(_) else _, type text}, {"藍圖代碼", each if _ is text then Text.Trim(_) else _, type text}, {"提供者", each if _ is text then Text.Trim(_) else _, type text}, {"備註", each _, type text}}),
    AddedSearch = Table.AddColumn(Trimmed, "搜尋文字", each Text.Combine(List.Select({[藍圖名稱],[備註]}, each _ <> null and _ <> ""), " | "), type text)
in
    AddedSearch
"@
}
```

- [ ] **Step 5: Run the helper tests to verify they pass**

Run: `powershell -NoProfile -ExecutionPolicy Bypass -File "tests/test_sihao_valley_helpers.ps1"`
Expected: PASS with no thrown exceptions.

- [ ] **Step 6: Commit the helper groundwork**

```bash
git add scripts/SihaoValleyTable.psm1 tests/test_sihao_valley_helpers.ps1
git commit -m "feat: add sihao valley workbook helpers"
```

### Task 2: Add workbook validation and safe Excel session handling

**Files:**
- Modify: `scripts/SihaoValleyTable.psm1`
- Create: `scripts/apply_sihao_valley_table.ps1`
- Modify: `tests/test_sihao_valley_helpers.ps1`

- [ ] **Step 1: Extend the failing tests for source-sheet assumptions**

```powershell
$formula = New-SourceNamedRangeFormula -SheetName '四號谷地_1'
if ($formula -notmatch 'MAX\(LOOKUP') {
    throw 'Named range formula must stay dynamic across all six columns and tolerate blanks'
}

$queryName = Get-SihaoQueryName
if ($queryName -ne '四號谷地_1_表格查詢') {
    throw 'Unexpected query name'
}

$tableName = Get-SihaoTableName
if ($tableName -ne 'SihaoValleyTable') {
    throw 'Unexpected table name'
}
```

- [ ] **Step 2: Run the tests to verify they fail**

Run: `powershell -NoProfile -ExecutionPolicy Bypass -File "tests/test_sihao_valley_helpers.ps1"`
Expected: FAIL because the new helper functions are not implemented yet.

- [ ] **Step 3: Implement workbook constants and cleanup-safe helpers**

```powershell
function Get-SihaoQueryName { '四號谷地_1_表格查詢' }
function Get-SihaoNamedRangeName { '四號谷地_1_來源範圍' }
function Get-SihaoTableName { 'SihaoValleyTable' }
function Test-HeadersMatch {
    param([object[]]$ActualHeaders)
    ((Get-SihaoExpectedHeaders) -join '|') -eq ($ActualHeaders -join '|')
}
```

- [ ] **Step 4: Write the orchestration script shell with guarded open/save/close behavior**

```powershell
param(
    [string]$WorkbookPath = (Join-Path $PSScriptRoot '..\Endfield Blueprints (Asia).xlsx')
)

Import-Module (Join-Path $PSScriptRoot 'SihaoValleyTable.psm1') -Force

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

try {
    $workbook = $excel.Workbooks.Open((Resolve-Path $WorkbookPath))
    # validate source sheet and headers here
    $workbook.Save()
}
finally {
    if ($workbook) { $workbook.Close($true) }
    $excel.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($workbook) | Out-Null
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
}
```

- [ ] **Step 5: Add explicit source-sheet and row-1 header validation**

```powershell
$sourceSheet = $workbook.Worksheets.Item('四號谷地_1')
$actualHeaders = 1..6 | ForEach-Object { [string]$sourceSheet.Cells.Item(1, $_).Text }
if (-not (Test-HeadersMatch -ActualHeaders $actualHeaders)) {
    throw "Header mismatch on 四號谷地_1 row 1"
}
```

- [ ] **Step 6: Capture a pre-change snapshot of source-sheet values before any workbook mutation**

```powershell
$sourceSnapshot = 1..$sourceSheet.UsedRange.Rows.Count | ForEach-Object {
    $row = $_
    1..6 | ForEach-Object { [string]$sourceSheet.Cells.Item($row, $_).Text }
}
```

- [ ] **Step 7: Compare the source-sheet snapshot before saving and fail if any source cell changed**

```powershell
$postSnapshot = 1..$sourceSheet.UsedRange.Rows.Count | ForEach-Object {
    $row = $_
    1..6 | ForEach-Object { [string]$sourceSheet.Cells.Item($row, $_).Text }
}

if (ConvertTo-Json $sourceSnapshot -Depth 4 -Compress -ne (ConvertTo-Json $postSnapshot -Depth 4 -Compress)) {
    throw 'Source sheet values changed during automation'
}
```

- [ ] **Step 8: Run the helper tests again**

Run: `powershell -NoProfile -ExecutionPolicy Bypass -File "tests/test_sihao_valley_helpers.ps1"`
Expected: PASS.

- [ ] **Step 9: Smoke-run the script before query creation**

Run: `powershell -NoProfile -ExecutionPolicy Bypass -File "scripts/apply_sihao_valley_table.ps1"`
Expected: PASS through open/validate/save/close without creating the output objects yet.

- [ ] **Step 10: Commit the validated Excel session scaffolding**

```bash
git add scripts/SihaoValleyTable.psm1 scripts/apply_sihao_valley_table.ps1 tests/test_sihao_valley_helpers.ps1
git commit -m "feat: validate sihao valley workbook inputs"
```

### Task 3: Create the dynamic named range and Power Query-backed output table

**Files:**
- Modify: `scripts/SihaoValleyTable.psm1`
- Modify: `scripts/apply_sihao_valley_table.ps1`
- Modify: `tests/test_sihao_valley_helpers.ps1`

- [ ] **Step 1: Add a failing test for exact query-shape expectations**

```powershell
$m = New-SihaoQueryFormula -NamedRange '四號谷地_1_來源範圍'
foreach ($required in @('Table.PromoteHeaders','Table.SelectRows','搜尋文字','藍圖名稱','備註')) {
    if ($m -notmatch [regex]::Escape($required)) {
        throw "Missing query fragment: $required"
    }
}
```

- [ ] **Step 2: Run the tests and verify the new assertions fail if needed**

Run: `powershell -NoProfile -ExecutionPolicy Bypass -File "tests/test_sihao_valley_helpers.ps1"`
Expected: FAIL until the query builder output matches all required fragments.

- [ ] **Step 3: Implement named-range creation in the orchestration script**

```powershell
$namedRangeName = Get-SihaoNamedRangeName
$namedFormula = New-SourceNamedRangeFormula -SheetName '四號谷地_1'
try { $workbook.Names.Item($namedRangeName).Delete() } catch {}
$null = $workbook.Names.Add($namedRangeName, $namedFormula)
```

- [ ] **Step 4: Implement deterministic output sheet naming and query creation**

```powershell
$existingNames = @($workbook.Worksheets | ForEach-Object { $_.Name })
$queryName = Get-SihaoQueryName
$tableName = Get-SihaoTableName
$queryFormula = New-SihaoQueryFormula -NamedRange $namedRangeName

$managedSheet = $null
foreach ($sheet in $workbook.Worksheets) {
    foreach ($table in $sheet.ListObjects) {
        if ($table.Name -eq $tableName) { $managedSheet = $sheet }
    }
}

if ($managedSheet) {
    $outputSheet = $managedSheet
    foreach ($sheetName in @('時代_Slicer','類別_Slicer','提供者_Slicer')) {
        try { $outputSheet.Shapes.Item($sheetName).Delete() } catch {}
    }
} else {
    $outputSheetName = Get-DeterministicSheetName -BaseName '四號谷地_1_表格' -ExistingNames $existingNames
}

foreach ($cacheName in @('時代_Slicer','類別_Slicer','提供者_Slicer')) {
    try { $workbook.SlicerCaches.Item($cacheName).Delete() } catch {}
}

try { $workbook.Connections.Item('Query - 四號谷地_1_表格查詢').Delete() } catch {}
try { $workbook.Queries.Item($queryName).Delete() } catch {}

$null = $workbook.Queries.Add($queryName, $queryFormula)
```

- [ ] **Step 5: Reuse the managed output sheet when it exists, otherwise create it once**

```powershell
$tableName = Get-SihaoTableName
if (-not $outputSheet) {
    $outputSheet = $workbook.Worksheets.Add()
    $outputSheet.Name = $outputSheetName
}

try { $outputSheet.ListObjects.Item($tableName).Delete() } catch {}
$destination = $outputSheet.Range('A1')
$connection = @("OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=$queryName;Extended Properties=\"\"")
$listObject = $outputSheet.ListObjects.Add(0, $connection, $true, 1, $destination)
$listObject.Name = $tableName
$queryTable = $listObject.QueryTable
$queryTable.CommandType = 2
$queryTable.CommandText = @("SELECT * FROM [$queryName]")
$queryTable.Refresh($false)
```

- [ ] **Step 6: Run the helper tests again**

Run: `powershell -NoProfile -ExecutionPolicy Bypass -File "tests/test_sihao_valley_helpers.ps1"`
Expected: PASS.

- [ ] **Step 7: Run the workbook script to create or refresh the managed output sheet and table**

Run: `powershell -NoProfile -ExecutionPolicy Bypass -File "scripts/apply_sihao_valley_table.ps1"`
Expected: PASS with one managed output sheet such as `四號谷地_1_表格` and a loaded table that includes `搜尋文字`.

- [ ] **Step 8: Commit the query-backed table creation**

```bash
git add scripts/SihaoValleyTable.psm1 scripts/apply_sihao_valley_table.ps1 tests/test_sihao_valley_helpers.ps1
git commit -m "feat: create refreshable sihao valley table"
```

### Task 4: Add slicers and finish workbook formatting details

**Files:**
- Modify: `scripts/apply_sihao_valley_table.ps1`

- [ ] **Step 1: Add a failing postcondition check for slicer creation**

```powershell
$expectedSlicers = @('時代_Slicer', '類別_Slicer', '提供者_Slicer')
$actualSlicers = @($outputSheet.Shapes | Where-Object { $_.Name -like '*Slicer' } | ForEach-Object Name)
foreach ($expected in $expectedSlicers) {
    if ($actualSlicers -notcontains $expected) {
        throw "Missing slicer: $expected"
    }
}
```

- [ ] **Step 2: Add table styling and column autofit after the initial refresh**

```powershell
$listObject.TableStyle = 'TableStyleMedium2'
$outputSheet.Columns.AutoFit() | Out-Null
```

- [ ] **Step 3: Create slicer caches and slicers for the three requested fields**

```powershell
$slicerLeft = 20
foreach ($columnName in @('時代','類別','提供者')) {
    $cache = $workbook.SlicerCaches.Add2($listObject, $columnName)
    $null = $cache.Slicers.Add($outputSheet, , "$columnName`_Slicer", $columnName, $slicerLeft, 20, 144, 180)
    $slicerLeft += 150
}
```

- [ ] **Step 4: Freeze the header row and keep the output sheet readable**

```powershell
$excel.ActiveWindow.SplitRow = 1
$excel.ActiveWindow.FreezePanes = $true
```

- [ ] **Step 5: Re-run the workbook script and confirm the slicer postcondition passes**

Run: `powershell -NoProfile -ExecutionPolicy Bypass -File "scripts/apply_sihao_valley_table.ps1"`
Expected: PASS with slicers visible on the output sheet for `時代`, `類別`, and `提供者`.

- [ ] **Step 6: Commit the slicer and formatting pass**

```bash
git add scripts/apply_sihao_valley_table.ps1
git commit -m "feat: add sihao valley table slicers"
```

### Task 5: Add workbook verification and usage documentation

**Files:**
- Create: `scripts/verify_sihao_valley_table.py`
- Modify: `README.md`

- [ ] **Step 1: Write the failing verification script**

```python
from openpyxl import load_workbook
import sys

workbook_path = sys.argv[1] if len(sys.argv) > 1 else 'Endfield Blueprints (Asia).xlsx'
baseline_path = sys.argv[2] if len(sys.argv) > 2 else None
wb = load_workbook(workbook_path, read_only=False)
assert '四號谷地_1' in wb.sheetnames
assert any(name.startswith('四號谷地_1_表格') for name in wb.sheetnames)
```

- [ ] **Step 2: Run the verification script to verify it fails before the workbook is correct**

Run: `python "scripts/verify_sihao_valley_table.py"`
Expected: FAIL until the workbook contains the expected output sheet and headers.

- [ ] **Step 3: Implement full workbook verification**

```python
from openpyxl import load_workbook
import sys

SOURCE = '四號谷地_1'
EXPECTED = ['時代', '類別', '藍圖名稱', '藍圖代碼', '提供者', '備註', '搜尋文字']

workbook_path = sys.argv[1] if len(sys.argv) > 1 else 'Endfield Blueprints (Asia).xlsx'
baseline_path = sys.argv[2] if len(sys.argv) > 2 else None
wb = load_workbook(workbook_path, read_only=False)
source = wb[SOURCE]
target_name = next(name for name in wb.sheetnames if 'SihaoValleyTable' in wb[name].tables)
target = wb[target_name]

source_headers = [source.cell(1, idx).value for idx in range(1, 7)]
assert source_headers == EXPECTED[:6]
headers = [target.cell(1, idx).value for idx in range(1, 8)]
assert headers == EXPECTED

if baseline_path:
    baseline = load_workbook(baseline_path, read_only=False)[SOURCE]
    for row in range(1, source.max_row + 1):
        current = [source.cell(row, col).value for col in range(1, 7)]
        original = [baseline.cell(row, col).value for col in range(1, 7)]
        assert current == original
```

- [ ] **Step 4: Document how to run the automation and verification**

```markdown
## Workbook automation

- Run `powershell -NoProfile -ExecutionPolicy Bypass -File "scripts/apply_sihao_valley_table.ps1"`
- Verify with `python "scripts/verify_sihao_valley_table.py" [optional-workbook-path] [optional-baseline-workbook-path]`
```

- [ ] **Step 5: Run the automation and then run the verifier against a baseline copy**

Run: `copy "Endfield Blueprints (Asia).xlsx" "%TEMP%\sihao-baseline.xlsx" && powershell -NoProfile -ExecutionPolicy Bypass -File "scripts/apply_sihao_valley_table.ps1" && python "scripts/verify_sihao_valley_table.py" "Endfield Blueprints (Asia).xlsx" "%TEMP%\sihao-baseline.xlsx"`
Expected: PASS. The verifier should confirm source-sheet presence, full source-sheet immutability across the first six columns, output-sheet presence, expected headers, and nonzero data rows.

- [ ] **Step 6: Commit the verification tooling and docs**

```bash
git add scripts/verify_sihao_valley_table.py README.md
git commit -m "docs: add sihao valley workbook automation usage"
```

### Task 6: Final end-to-end validation for the workbook result

**Files:**
- Modify: `scripts/verify_sihao_valley_table.py`
- Modify: `README.md`

- [ ] **Step 1: Extend the verifier for row-count and text-integrity checks**

```python
def count_data_rows(ws, required_columns):
    count = 0
    for row in ws.iter_rows(min_row=2, max_col=required_columns, values_only=True):
        if any(value not in (None, '') for value in row):
            count += 1
    return count
```

- [ ] **Step 2: Run the verifier and confirm it catches row mismatches if present**

Run: `python "scripts/verify_sihao_valley_table.py"`
Expected: PASS only when the output row count matches the source nonblank row count and long text cells are preserved.

- [ ] **Step 3: Add explicit checks for mixed-language provider values and long notes**

```python
providers = [target.cell(row, 5).value for row in range(2, min(target.max_row, 20))]
notes = [target.cell(row, 6).value for row in range(2, min(target.max_row, 20))]
assert any(isinstance(value, str) and len(value) > 20 for value in notes if value)
assert any(isinstance(value, str) and value not in ('N/A', '') for value in providers if value is not None)
```

- [ ] **Step 4: Add a refresh-on-source-change test against a temporary workbook copy**

```powershell
$tempWorkbook = Join-Path $env:TEMP 'sihao-valley-refresh-test.xlsx'
Copy-Item 'Endfield Blueprints (Asia).xlsx' $tempWorkbook -Force

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$workbook = $excel.Workbooks.Open((Resolve-Path $tempWorkbook))
$sheet = $workbook.Worksheets.Item('四號谷地_1')
$nextRow = $sheet.UsedRange.Rows.Count + 1
$sheet.Cells.Item($nextRow, 1).Value2 = '測試時代'
$sheet.Cells.Item($nextRow, 2).Value2 = '測試類別'
$sheet.Cells.Item($nextRow, 3).Value2 = '刷新驗證藍圖'
$sheet.Cells.Item($nextRow, 4).Value2 = 'REFRESH-TEST-CODE'
$sheet.Cells.Item($nextRow, 5).Value2 = 'AutomationTest'
$sheet.Cells.Item($nextRow, 6).Value2 = 'refresh check row'
$workbook.Save()
$workbook.Close($true)
$excel.Quit()

powershell -NoProfile -ExecutionPolicy Bypass -File "scripts/apply_sihao_valley_table.ps1" -WorkbookPath $tempWorkbook
python "scripts/verify_sihao_valley_table.py" $tempWorkbook
Remove-Item $tempWorkbook -Force
```

- [ ] **Step 5: Re-run the full helper tests, automation, and verifier on the real workbook**

Run: `powershell -NoProfile -ExecutionPolicy Bypass -File "tests/test_sihao_valley_helpers.ps1" && powershell -NoProfile -ExecutionPolicy Bypass -File "scripts/apply_sihao_valley_table.ps1" && python "scripts/verify_sihao_valley_table.py"`
Expected: PASS end to end.

- [ ] **Step 6: Record the final manual validation checklist in README**

```markdown
- Open `Endfield Blueprints (Asia).xlsx`
- Check the managed `四號谷地_1_表格*` output sheet that contains `SihaoValleyTable`
- Confirm slicers exist for `時代`, `類別`, `提供者`
- Use the `搜尋文字` column filter search to narrow rows by partial keyword
```

- [ ] **Step 7: Commit the final validation pass**

```bash
git add scripts/verify_sihao_valley_table.py README.md
git commit -m "test: verify sihao valley workbook output"
```
