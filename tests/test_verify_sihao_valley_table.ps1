$ErrorActionPreference = 'Stop'

$repoRoot = Split-Path $PSScriptRoot -Parent
$applyScriptPath = Join-Path $repoRoot 'scripts\apply_sihao_valley_table.ps1'
$scriptPath = Join-Path $repoRoot 'scripts\verify_sihao_valley_table.py'
$workbookPath = Join-Path $repoRoot 'Endfield Blueprints (Asia).xlsx'
$baselinePath = $workbookPath

if (-not (Test-Path -LiteralPath $workbookPath)) {
    throw "Workbook not found: $workbookPath"
}

if (-not (Test-Path -LiteralPath $baselinePath)) {
    throw "Baseline workbook not found: $baselinePath"
}

function Remove-ComObject {
    param($ComObject)

    if ($null -ne $ComObject -and [System.Runtime.InteropServices.Marshal]::IsComObject($ComObject)) {
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($ComObject)
    }
}

function Copy-WorkbookToTemp {
    $tempDirectory = Join-Path ([System.IO.Path]::GetTempPath()) ([System.Guid]::NewGuid().ToString())
    [System.IO.Directory]::CreateDirectory($tempDirectory) | Out-Null
    $tempWorkbookPath = Join-Path $tempDirectory ([System.IO.Path]::GetFileName($workbookPath))
    Copy-Item -LiteralPath $workbookPath -Destination $tempWorkbookPath -Force
    return $tempWorkbookPath
}

function Add-UniqueSourceRow {
    param(
        [Parameter(Mandatory = $true)]
        [string]$TargetWorkbookPath,

        [Parameter(Mandatory = $true)]
        [string]$UniqueToken
    )

    $excel = $null
    $workbook = $null
    $worksheet = $null

    try {
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        $excel.DisplayAlerts = $false
        $excel.ScreenUpdating = $false

        $workbook = $excel.Workbooks.Open($TargetWorkbookPath)
        $worksheet = $workbook.Worksheets.Item('四號谷地_1')

        $lastRow = 1
        for ($columnIndex = 1; $columnIndex -le 6; $columnIndex++) {
            $columnLastRow = [int]$worksheet.Cells.Item($worksheet.Rows.Count, $columnIndex).End(-4162).Row
            if ($columnLastRow -gt $lastRow) {
                $lastRow = $columnLastRow
            }
        }

        $newRow = $lastRow + 1
        $worksheet.Cells.Item($newRow, 1).Value2 = '驗證區'
        $worksheet.Cells.Item($newRow, 2).Value2 = '測試'
        $worksheet.Cells.Item($newRow, 3).Value2 = "Task6 refresh $UniqueToken"
        $worksheet.Cells.Item($newRow, 4).Value2 = "TASK6$UniqueToken"
        $worksheet.Cells.Item($newRow, 5).Value2 = "混合Provider-$UniqueToken-김덕구"
        $worksheet.Cells.Item($newRow, 6).Value2 = "Task 6 refresh validation note $UniqueToken"

        $workbook.Save()
    }
    finally {
        if ($null -ne $workbook) {
            $workbook.Close($false)
        }

        Remove-ComObject $worksheet
        Remove-ComObject $workbook

        if ($null -ne $excel) {
            $excel.Quit()
        }

        Remove-ComObject $excel
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
    }
}

$output = & python $scriptPath $workbookPath $baselinePath 2>&1
$exitCode = $LASTEXITCODE

if ($exitCode -ne 0) {
    throw "verify_sihao_valley_table.py exited with code $exitCode.`nOutput:`n$output"
}

$tempWorkbookPath = Copy-WorkbookToTemp
$uniqueToken = [System.Guid]::NewGuid().ToString('N').Substring(0, 8)

try {
    Add-UniqueSourceRow -TargetWorkbookPath $tempWorkbookPath -UniqueToken $uniqueToken

    $applyOutput = & pwsh -File $applyScriptPath -EnableWorkbookWrites -WorkbookPath $tempWorkbookPath 2>&1
    $applyExitCode = $LASTEXITCODE
    if ($applyExitCode -ne 0) {
        throw "apply_sihao_valley_table.ps1 exited with code $applyExitCode.`nOutput:`n$applyOutput"
    }

    $refreshVerifyOutput = & python $scriptPath $tempWorkbookPath 2>&1
    $refreshVerifyExitCode = $LASTEXITCODE
    if ($refreshVerifyExitCode -ne 0) {
        throw "verify_sihao_valley_table.py failed for refreshed temp workbook with code $refreshVerifyExitCode.`nOutput:`n$refreshVerifyOutput"
    }

    $refreshCheckOutput = & python -X utf8 -c @"
from pathlib import Path
import warnings
from openpyxl import load_workbook
from openpyxl.utils import range_boundaries

workbook_path = Path(r'$tempWorkbookPath')
unique_token = '$uniqueToken'

with warnings.catch_warnings():
    warnings.filterwarnings('ignore', message='Unknown extension is not supported and will be removed', category=UserWarning)
    workbook = load_workbook(workbook_path, read_only=False, data_only=False)

try:
    output_sheet = None
    table = None
    for worksheet in workbook.worksheets:
        if 'SihaoValleyTable' in worksheet.tables:
            output_sheet = worksheet
            table = worksheet.tables['SihaoValleyTable']
            break

    if table is None:
        raise SystemExit('Managed output table not found in temp workbook.')

    min_col, min_row, max_col, max_row = range_boundaries(table.ref)
    matched = False
    for row in output_sheet.iter_rows(min_row=min_row + 1, max_row=max_row, min_col=min_col, max_col=max_col, values_only=True):
        if unique_token in str(row[2]) and unique_token in str(row[3]) and unique_token in str(row[4]):
            matched = True
            break

    if not matched:
        raise SystemExit(f'Unique source row {unique_token} was not found in managed output table.')

    print(f'Found refreshed row for {unique_token}.')
finally:
    workbook.close()
"@ 2>&1
    $refreshCheckExitCode = $LASTEXITCODE
    if ($refreshCheckExitCode -ne 0) {
        throw "Temp workbook refresh check failed with code $refreshCheckExitCode.`nOutput:`n$refreshCheckOutput"
    }
}
finally {
    if (Test-Path -LiteralPath $tempWorkbookPath) {
        Remove-Item -LiteralPath (Split-Path -Path $tempWorkbookPath -Parent) -Recurse -Force
    }
}

Write-Host 'Sihao valley verification integration test passed, including temp-workbook refresh validation.'
