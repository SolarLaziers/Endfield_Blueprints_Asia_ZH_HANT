param(
    [switch]$EnableWorkbookWrites,
    [string]$WorkbookPath = (Join-Path (Split-Path $PSScriptRoot -Parent) 'Endfield Blueprints (Asia).xlsx')
)

$ErrorActionPreference = 'Stop'

Set-StrictMode -Version Latest

$modulePath = Join-Path $PSScriptRoot 'SihaoValleyTable.psm1'

if (-not (Test-Path -LiteralPath $modulePath)) {
    throw "Module not found: $modulePath"
}

Import-Module $modulePath -Force

$workbookPath = (Resolve-Path -LiteralPath $WorkbookPath).Path
$sourceSheetName = [string]::Concat([char]0x56DB, [char]0x865F, [char]0x8C37, [char]0x5730, '_1')
$expectedHeaders = Get-SihaoExpectedHeaders
$sourceNamedRangeName = Get-SihaoSourceNamedRangeName
$queryName = Get-SihaoQueryName
$tableName = Get-SihaoTableName
$outputSheetBaseName = Get-SihaoOutputSheetBaseName
$openReadOnly = -not $EnableWorkbookWrites
$executionMode = if ($EnableWorkbookWrites) { 'write-enabled' } else { 'read-only' }

function Release-ComObject {
    param($ComObject)

    if ($null -ne $ComObject -and [System.Runtime.InteropServices.Marshal]::IsComObject($ComObject)) {
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($ComObject)
    }
}

function Get-ExcelRangeValueMatrix {
    param(
        [Parameter(Mandatory = $true)]
        $Range
    )

    $rowCount = [int]$Range.Rows.Count
    $columnCount = [int]$Range.Columns.Count
    $values = $Range.Value2
    $rows = New-Object 'System.Collections.Generic.List[object]'

    if ($rowCount -eq 1 -and $columnCount -eq 1) {
        $rows.Add(@($values))
        return ,$rows.ToArray()
    }

    if ($rowCount -eq 1) {
        $singleRow = for ($columnIndex = 1; $columnIndex -le $columnCount; $columnIndex++) {
            $values.GetValue(1, $columnIndex)
        }

        $rows.Add(@($singleRow))
        return ,$rows.ToArray()
    }

    if ($columnCount -eq 1) {
        for ($rowIndex = 1; $rowIndex -le $rowCount; $rowIndex++) {
            $rows.Add(@($values.GetValue($rowIndex, 1)))
        }

        return ,$rows.ToArray()
    }

    for ($rowIndex = 1; $rowIndex -le $rowCount; $rowIndex++) {
        $rowValues = for ($columnIndex = 1; $columnIndex -le $columnCount; $columnIndex++) {
            $values.GetValue($rowIndex, $columnIndex)
        }

        $rows.Add(@($rowValues))
    }

    return ,$rows.ToArray()
}

function Get-SourceSnapshot {
    param(
        [Parameter(Mandatory = $true)]
        $Worksheet
    )

    $lastRow = Get-SourceLastRow -Worksheet $Worksheet
    $startCell = $null
    $endCell = $null
    $range = $null

    try {
        $startCell = $Worksheet.Cells.Item(1, 1)
        $endCell = $Worksheet.Cells.Item($lastRow, 6)
        $range = $Worksheet.Range($startCell, $endCell)

        return Get-ExcelRangeValueMatrix -Range $range
    }
    finally {
        foreach ($comObject in @($range, $endCell, $startCell)) {
            Release-ComObject $comObject
        }
    }
}

function Get-SourceLastRow {
    param(
        [Parameter(Mandatory = $true)]
        $Worksheet
    )

    $startCell = $null
    $rowProbeCell = $null
    $rowProbeEndCell = $null

    try {
        $lastRow = 1

        for ($columnIndex = 1; $columnIndex -le 6; $columnIndex++) {
            $rowProbeCell = $Worksheet.Cells.Item($Worksheet.Rows.Count, $columnIndex)
            $rowProbeEndCell = $rowProbeCell.End(-4162)
            $columnLastRow = [int]$rowProbeEndCell.Row
            if ($columnLastRow -gt $lastRow) {
                $lastRow = $columnLastRow
            }

            Release-ComObject $rowProbeEndCell
            Release-ComObject $rowProbeCell
            $rowProbeEndCell = $null
            $rowProbeCell = $null
        }

        return $lastRow
    }
    finally {
        foreach ($comObject in @($startCell, $rowProbeEndCell, $rowProbeCell)) {
            Release-ComObject $comObject
        }
    }
}

function Assert-HeadersMatch {
    param(
        [Parameter(Mandatory = $true)]
        $Worksheet,

        [Parameter(Mandatory = $true)]
        [string[]]$ExpectedHeaders
    )

    $actualHeaders = @()
    $headerCell = $null

    try {
        for ($columnIndex = 1; $columnIndex -le $ExpectedHeaders.Count; $columnIndex++) {
            $headerCell = $Worksheet.Cells.Item(1, $columnIndex)
            $actualHeaders += [string]$headerCell.Text
            Release-ComObject $headerCell
            $headerCell = $null
        }
    }
    finally {
        Release-ComObject $headerCell
    }

    if ((ConvertTo-Json $actualHeaders -Compress) -ne (ConvertTo-Json $ExpectedHeaders -Compress)) {
        $expectedDisplay = $ExpectedHeaders -join ', '
        $actualDisplay = $actualHeaders -join ', '
        throw "Source sheet '$sourceSheetName' headers do not match exactly. Expected: [$expectedDisplay]. Actual: [$actualDisplay]."
    }
}

function Assert-SnapshotUnchanged {
    param(
        [Parameter(Mandatory = $true)]
        [object[]]$Before,

        [Parameter(Mandatory = $true)]
        [object[]]$After
    )

    if ((ConvertTo-Json $Before -Compress -Depth 4) -ne (ConvertTo-Json $After -Compress -Depth 4)) {
        throw "Source sheet '$sourceSheetName' changed in columns A:F during script execution."
    }
}

function Try-GetWorksheet {
    param(
        [Parameter(Mandatory = $true)]$Workbook,
        [Parameter(Mandatory = $true)][string]$WorksheetName
    )

    try {
        return $Workbook.Worksheets.Item($WorksheetName)
    }
    catch {
        return $null
    }
}

function Get-WorksheetNames {
    param(
        [Parameter(Mandatory = $true)]$Workbook
    )

    $names = New-Object 'System.Collections.Generic.List[string]'
    $worksheet = $null

    try {
        for ($worksheetIndex = 1; $worksheetIndex -le $Workbook.Worksheets.Count; $worksheetIndex++) {
            $worksheet = $Workbook.Worksheets.Item($worksheetIndex)
            $names.Add([string]$worksheet.Name)
            Release-ComObject $worksheet
            $worksheet = $null
        }
    }
    finally {
        Release-ComObject $worksheet
    }

    return ,$names.ToArray()
}

function Find-ManagedTableWorksheet {
    param(
        [Parameter(Mandatory = $true)]$Workbook,
        [Parameter(Mandatory = $true)][string]$ManagedTableName
    )

    $worksheet = $null
    $listObject = $null

    try {
        for ($worksheetIndex = 1; $worksheetIndex -le $Workbook.Worksheets.Count; $worksheetIndex++) {
            $worksheet = $Workbook.Worksheets.Item($worksheetIndex)

            for ($tableIndex = 1; $tableIndex -le $worksheet.ListObjects.Count; $tableIndex++) {
                $listObject = $worksheet.ListObjects.Item($tableIndex)
                if ([string]$listObject.Name -eq $ManagedTableName) {
                    Release-ComObject $listObject
                    $listObject = $null
                    return $worksheet
                }

                Release-ComObject $listObject
                $listObject = $null
            }

            Release-ComObject $worksheet
            $worksheet = $null
        }
    }
    finally {
        Release-ComObject $listObject
    }

    return $null
}

function Remove-WorkbookNameIfExists {
    param(
        [Parameter(Mandatory = $true)]$Workbook,
        [Parameter(Mandatory = $true)][string]$Name
    )

    $workbookName = $null

    try {
        $workbookName = $Workbook.Names.Item($Name)
        $workbookName.Delete()
    }
    catch {
    }
    finally {
        Release-ComObject $workbookName
    }
}

function Try-GetWorkbookName {
    param(
        [Parameter(Mandatory = $true)]$Workbook,
        [Parameter(Mandatory = $true)][string]$Name
    )

    try {
        return $Workbook.Names.Item($Name)
    }
    catch {
        return $null
    }
}

function Remove-WorkbookQueryIfExists {
    param(
        [Parameter(Mandatory = $true)]$Workbook,
        [Parameter(Mandatory = $true)][string]$QueryName
    )

    $query = $null

    try {
        $query = $Workbook.Queries.Item($QueryName)
        $query.Delete()
    }
    catch {
    }
    finally {
        Release-ComObject $query
    }
}

function Try-GetWorkbookQuery {
    param(
        [Parameter(Mandatory = $true)]$Workbook,
        [Parameter(Mandatory = $true)][string]$QueryName
    )

    try {
        return $Workbook.Queries.Item($QueryName)
    }
    catch {
        return $null
    }
}

function Remove-WorkbookConnectionIfExists {
    param(
        [Parameter(Mandatory = $true)]$Workbook,
        [Parameter(Mandatory = $true)][string[]]$ConnectionNames
    )

    foreach ($connectionName in $ConnectionNames) {
        $connection = $null

        try {
            $connection = $Workbook.Connections.Item($connectionName)
            [void]$connection.Delete()
        }
        catch {
        }
        finally {
            Release-ComObject $connection
        }
    }
}

function Try-GetWorkbookConnection {
    param(
        [Parameter(Mandatory = $true)]$Workbook,
        [Parameter(Mandatory = $true)][string]$ConnectionName
    )

    try {
        return $Workbook.Connections.Item($ConnectionName)
    }
    catch {
        return $null
    }
}

function Get-SlicerDefinitions {
    param(
        [Parameter(Mandatory = $true)][string]$ManagedTableName
    )

    return @(
        [pscustomobject]@{
            FieldName       = '時代'
            SlicerName      = "${ManagedTableName}_EraSlicer"
            Caption         = '時代'
            AnchorRow       = 2
            AnchorColumn    = 9
            Width           = 144
            Height          = 144
            NumberOfColumns = 1
        },
        [pscustomobject]@{
            FieldName       = '類別'
            SlicerName      = "${ManagedTableName}_CategorySlicer"
            Caption         = '類別'
            AnchorRow       = 14
            AnchorColumn    = 9
            Width           = 144
            Height          = 144
            NumberOfColumns = 1
        },
        [pscustomobject]@{
            FieldName       = '提供者'
            SlicerName      = "${ManagedTableName}_ProviderSlicer"
            Caption         = '提供者'
            AnchorRow       = 26
            AnchorColumn    = 9
            Width           = 144
            Height          = 144
            NumberOfColumns = 1
        }
    )
}

function Remove-OrphanedManagedSlicerCaches {
    param(
        [Parameter(Mandatory = $true)]$Workbook,
        [Parameter(Mandatory = $true)][string[]]$FieldNames
    )

    $slicerCache = $null

    try {
        for ($cacheIndex = $Workbook.SlicerCaches.Count; $cacheIndex -ge 1; $cacheIndex--) {
            $slicerCache = $Workbook.SlicerCaches.Item($cacheIndex)
            if (([string]$slicerCache.SourceName -in $FieldNames) -and ($slicerCache.Slicers.Count -eq 0)) {
                $slicerCache.Delete()
            }

            Release-ComObject $slicerCache
            $slicerCache = $null
        }
    }
    finally {
        Release-ComObject $slicerCache
    }
}

function Remove-SlicersByName {
    param(
        [Parameter(Mandatory = $true)]$Workbook,
        [Parameter(Mandatory = $true)][string[]]$SlicerNames
    )

    $slicerCache = $null
    $slicer = $null

    try {
        for ($cacheIndex = $Workbook.SlicerCaches.Count; $cacheIndex -ge 1; $cacheIndex--) {
            $slicerCache = $Workbook.SlicerCaches.Item($cacheIndex)

            for ($slicerIndex = $slicerCache.Slicers.Count; $slicerIndex -ge 1; $slicerIndex--) {
                $slicer = $slicerCache.Slicers.Item($slicerIndex)
                if ([string]$slicer.Name -in $SlicerNames) {
                    $slicer.Delete()
                }

                Release-ComObject $slicer
                $slicer = $null
            }

            Release-ComObject $slicerCache
            $slicerCache = $null
        }
    }
    finally {
        Release-ComObject $slicer
        Release-ComObject $slicerCache
    }
}

function Add-OrReplaceManagedSlicers {
    param(
        [Parameter(Mandatory = $true)]$Workbook,
        [Parameter(Mandatory = $true)]$Worksheet,
        [Parameter(Mandatory = $true)]$ListObject,
        [Parameter(Mandatory = $true)][string]$ManagedTableName
    )

    $definitions = Get-SlicerDefinitions -ManagedTableName $ManagedTableName
    Remove-SlicersByName -Workbook $Workbook -SlicerNames ($definitions | ForEach-Object { $_.SlicerName })
    Remove-OrphanedManagedSlicerCaches -Workbook $Workbook -FieldNames ($definitions | ForEach-Object { $_.FieldName })

    foreach ($definition in $definitions) {
        $anchorCell = $null
        $slicerCache = $null
        $slicer = $null

        try {
            $anchorCell = $Worksheet.Cells.Item($definition.AnchorRow, $definition.AnchorColumn)
            $slicerCache = $Workbook.SlicerCaches.Add2($ListObject, $definition.FieldName)
            $slicer = $slicerCache.Slicers.Add($Worksheet, [System.Type]::Missing, $definition.SlicerName, $definition.Caption, [double]$anchorCell.Left, [double]$anchorCell.Top, [double]$definition.Width, [double]$definition.Height)
            $slicer.NumberOfColumns = $definition.NumberOfColumns
        }
        finally {
            Release-ComObject $slicer
            Release-ComObject $slicerCache
            Release-ComObject $anchorCell
        }
    }
}

function Assert-ManagedSlicersPresent {
    param(
        [Parameter(Mandatory = $true)]$Workbook,
        [Parameter(Mandatory = $true)]$Worksheet,
        [Parameter(Mandatory = $true)][string]$ManagedTableName
    )

    $definitions = Get-SlicerDefinitions -ManagedTableName $ManagedTableName
    $missingSlicers = New-Object 'System.Collections.Generic.List[string]'

    foreach ($definition in $definitions) {
        $slicerCache = $null
        $slicer = $null
        $worksheetShape = $null
        $matchedCache = $false

        try {
            for ($cacheIndex = 1; $cacheIndex -le $Workbook.SlicerCaches.Count; $cacheIndex++) {
                $slicerCache = $Workbook.SlicerCaches.Item($cacheIndex)
                if ([string]$slicerCache.SourceName -eq $definition.FieldName) {
                    for ($slicerIndex = 1; $slicerIndex -le $slicerCache.Slicers.Count; $slicerIndex++) {
                        $slicer = $slicerCache.Slicers.Item($slicerIndex)
                        if ([string]$slicer.Name -eq $definition.SlicerName) {
                            $matchedCache = $true
                        }

                        Release-ComObject $slicer
                        $slicer = $null

                        if ($matchedCache) {
                            break
                        }
                    }
                }

                Release-ComObject $slicerCache
                $slicerCache = $null

                if ($matchedCache) {
                    break
                }
            }

            try {
                $worksheetShape = $Worksheet.Shapes.Item($definition.SlicerName)
            }
            catch {
                $worksheetShape = $null
            }

            if ((-not $matchedCache) -or ($null -eq $worksheetShape)) {
                $missingSlicers.Add($definition.FieldName)
            }
        }
        finally {
            Release-ComObject $slicer
            Release-ComObject $worksheetShape
            Release-ComObject $slicerCache
        }
    }

    if ($missingSlicers.Count -gt 0) {
        throw "Expected slicers were not present on worksheet '$($Worksheet.Name)': $($missingSlicers -join ', ')."
    }
}

function Format-ManagedOutputWorksheet {
    param(
        [Parameter(Mandatory = $true)]$Worksheet,
        [Parameter(Mandatory = $true)]$ListObject
    )

    $usedRange = $null

    try {
        $ListObject.TableStyle = 'TableStyleMedium2'
        $ListObject.ShowTableStyleRowStripes = $true
        $usedRange = $Worksheet.UsedRange
        [void]$usedRange.Columns.AutoFit()
    }
    finally {
        Release-ComObject $usedRange
    }
}

function Freeze-ManagedOutputHeaderRow {
    param(
        [Parameter(Mandatory = $true)]$Excel,
        [Parameter(Mandatory = $true)]$Worksheet
    )

    $freezeCell = $null
    $activeWindow = $null

    try {
        [void]$Worksheet.Activate()
        $freezeCell = $Worksheet.Range('A2')
        [void]$freezeCell.Select()
        $activeWindow = $Excel.ActiveWindow
        $activeWindow.FreezePanes = $false
        $activeWindow.SplitColumn = 0
        $activeWindow.SplitRow = 1
        $activeWindow.FreezePanes = $true
    }
    finally {
        Release-ComObject $activeWindow
        Release-ComObject $freezeCell
    }
}

function Add-OrUpdateWorkbookConnection {
    param(
        [Parameter(Mandatory = $true)]$Workbook,
        [Parameter(Mandatory = $true)][string]$QueryName
    )

    $connectionName = "Query - $QueryName"
    $commandText = @("SELECT * FROM [$QueryName]")

    Remove-WorkbookConnectionIfExists -Workbook $Workbook -ConnectionNames @($connectionName, $QueryName)

    return $Workbook.Connections.Add2(
        $connectionName,
        "Connection to the '$QueryName' query in the workbook.",
        "OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=`$Workbook$;Location=$QueryName;Extended Properties=`"`"",
        $commandText,
        2,
        $false,
        $false
    )
}

function Remove-ManagedTableFromWorksheet {
    param(
        [Parameter(Mandatory = $true)]$Worksheet,
        [Parameter(Mandatory = $true)][string]$ManagedTableName
    )

    $listObject = $null

    try {
        for ($tableIndex = 1; $tableIndex -le $Worksheet.ListObjects.Count; $tableIndex++) {
            $listObject = $Worksheet.ListObjects.Item($tableIndex)
            if ([string]$listObject.Name -eq $ManagedTableName) {
                [void]$listObject.Delete()
                break
            }

            Release-ComObject $listObject
            $listObject = $null
        }
    }
    finally {
        Release-ComObject $listObject
    }
}

function Remove-AllTablesFromWorksheet {
    param(
        [Parameter(Mandatory = $true)]$Worksheet
    )

    $listObject = $null

    try {
        for ($tableIndex = $Worksheet.ListObjects.Count; $tableIndex -ge 1; $tableIndex--) {
            $listObject = $Worksheet.ListObjects.Item($tableIndex)
            [void]$listObject.Delete()
            Release-ComObject $listObject
            $listObject = $null
        }
    }
    finally {
        Release-ComObject $listObject
    }
}

function Clear-WorksheetContents {
    param(
        [Parameter(Mandatory = $true)]$Worksheet
    )

    $cells = $null

    try {
        $cells = $Worksheet.Cells
        [void]$cells.Clear()
    }
    finally {
        Release-ComObject $cells
    }
}

function Get-OrCreateManagedOutputWorksheet {
    param(
        [Parameter(Mandatory = $true)]$Workbook,
        [Parameter(Mandatory = $true)][string]$ManagedTableName,
        [Parameter(Mandatory = $true)][string]$BaseSheetName
    )

    $managedWorksheet = Find-ManagedTableWorksheet -Workbook $Workbook -ManagedTableName $ManagedTableName
    if ($null -ne $managedWorksheet) {
        return $managedWorksheet
    }

    $existingNames = Get-WorksheetNames -Workbook $Workbook
    $preferredSheetName = Get-PreferredManagedSheetName -BaseName $BaseSheetName -ExistingNames $existingNames
    $existingWorksheet = Try-GetWorksheet -Workbook $Workbook -WorksheetName $preferredSheetName
    if ($null -ne $existingWorksheet) {
        return $existingWorksheet
    }

    $afterWorksheet = $null
    $newWorksheet = $null

    try {
        $afterWorksheet = $Workbook.Worksheets.Item($Workbook.Worksheets.Count)
        $newWorksheet = $Workbook.Worksheets.Add([System.Type]::Missing, $afterWorksheet, 1, [System.Type]::Missing)
        $newWorksheet.Name = $preferredSheetName
        return $newWorksheet
    }
    finally {
        Release-ComObject $afterWorksheet
    }
}

function Add-OrUpdateSourceNamedRange {
    param(
        [Parameter(Mandatory = $true)]$Workbook,
        [Parameter(Mandatory = $true)][string]$RangeName,
        [Parameter(Mandatory = $true)][string]$SheetName,
        [Parameter(Mandatory = $true)][int]$LastRow
    )

    $formula = New-SourceNamedRangeFormula -SheetName $SheetName -LastRow $LastRow
    Remove-WorkbookNameIfExists -Workbook $Workbook -Name $RangeName
    [void]$Workbook.Names.Add($RangeName, $formula)
}

function Add-OrUpdatePowerQuery {
    param(
        [Parameter(Mandatory = $true)]$Workbook,
        [Parameter(Mandatory = $true)][string]$QueryName,
        [Parameter(Mandatory = $true)][string]$NamedRangeName
    )

    $formula = New-SihaoQueryFormula -NamedRangeName $NamedRangeName
    Remove-WorkbookQueryIfExists -Workbook $Workbook -QueryName $QueryName
    [void]$Workbook.Queries.Add($QueryName, $formula)
}

function Get-ManagedTableFromWorksheet {
    param(
        [Parameter(Mandatory = $true)]$Worksheet,
        [Parameter(Mandatory = $true)][string]$ManagedTableName
    )

    $listObject = $null

    try {
        for ($tableIndex = 1; $tableIndex -le $Worksheet.ListObjects.Count; $tableIndex++) {
            $listObject = $Worksheet.ListObjects.Item($tableIndex)
            if ([string]$listObject.Name -eq $ManagedTableName) {
                return $listObject
            }

            Release-ComObject $listObject
            $listObject = $null
        }
    }
    finally {
        if ($null -ne $listObject -and [string]$listObject.Name -ne $ManagedTableName) {
            Release-ComObject $listObject
        }
    }

    return $null
}

function Test-PowerQueryMatchesDesiredFormula {
    param(
        [Parameter(Mandatory = $true)]$Workbook,
        [Parameter(Mandatory = $true)][string]$QueryName,
        [Parameter(Mandatory = $true)][string]$NamedRangeName
    )

    $query = $null

    try {
        $query = Try-GetWorkbookQuery -Workbook $Workbook -QueryName $QueryName
        if ($null -eq $query) {
            return $false
        }

        $expectedFormula = (New-SihaoQueryFormula -NamedRangeName $NamedRangeName) -replace "`r`n", "`n"
        $actualFormula = ([string]$query.Formula) -replace "`r`n", "`n"
        return $actualFormula -eq $expectedFormula
    }
    finally {
        Release-ComObject $query
    }
}

function Get-ListObjectHeaderTexts {
    param(
        [Parameter(Mandatory = $true)]$ListObject
    )

    $headers = New-Object 'System.Collections.Generic.List[string]'
    $listColumn = $null

    try {
        for ($columnIndex = 1; $columnIndex -le $ListObject.ListColumns.Count; $columnIndex++) {
            $listColumn = $ListObject.ListColumns.Item($columnIndex)
            $headers.Add(([string]$listColumn.Name).Trim())
            Release-ComObject $listColumn
            $listColumn = $null
        }
    }
    finally {
        Release-ComObject $listColumn
    }

    return ,$headers.ToArray()
}

function Add-OrRefreshManagedOutputTable {
    param(
        [Parameter(Mandatory = $true)]$Workbook,
        [Parameter(Mandatory = $true)]$Worksheet,
        [Parameter(Mandatory = $true)][string]$QueryName,
        [Parameter(Mandatory = $true)][string]$ManagedTableName
    )

    $destinationRange = $null
    $connection = $null
    $listObject = $null
    $queryTable = $null

    try {
        $commandText = @("SELECT * FROM [$QueryName]")
        $connection = Add-OrUpdateWorkbookConnection -Workbook $Workbook -QueryName $QueryName
        Remove-AllTablesFromWorksheet -Worksheet $Worksheet
        Clear-WorksheetContents -Worksheet $Worksheet
        $destinationRange = $Worksheet.Range('A1')
        $listObject = $Worksheet.ListObjects.Add(0, $connection, $true, 1, $destinationRange)
        $listObject.Name = $ManagedTableName

        $queryTable = $listObject.QueryTable
        $queryTable.CommandType = 2
        $queryTable.CommandText = $commandText
        $queryTable.BackgroundQuery = $false
        [void]$queryTable.Refresh($false)

        $headers = Get-ListObjectHeaderTexts -ListObject $listObject
        if ('搜尋文字' -notin $headers) {
            $headerDisplay = $headers -join ', '
            throw "Managed output table '$ManagedTableName' did not include the expected 搜尋文字 column. Headers: [$headerDisplay]."
        }

        return $listObject
    }
    finally {
        Release-ComObject $queryTable
        Release-ComObject $connection
        Release-ComObject $destinationRange
    }
}

function Refresh-ManagedOutputTable {
    param(
        [Parameter(Mandatory = $true)]$Worksheet,
        [Parameter(Mandatory = $true)][string]$ManagedTableName
    )

    $listObject = $null
    $queryTable = $null

    try {
        $listObject = Get-ManagedTableFromWorksheet -Worksheet $Worksheet -ManagedTableName $ManagedTableName
        if ($null -eq $listObject) {
            throw "Managed output table '$ManagedTableName' was not found on worksheet '$($Worksheet.Name)'."
        }

        $queryTable = $listObject.QueryTable
        $queryTable.BackgroundQuery = $false
        [void]$queryTable.Refresh($false)

        $headers = Get-ListObjectHeaderTexts -ListObject $listObject
        if ('搜尋文字' -notin $headers) {
            $headerDisplay = $headers -join ', '
            throw "Managed output table '$ManagedTableName' did not include the expected 搜尋文字 column. Headers: [$headerDisplay]."
        }

        return $listObject
    }
    finally {
        Release-ComObject $queryTable
    }
}

$excel = $null
$workbook = $null
$sourceWorksheet = $null
$outputWorksheet = $null
$outputTable = $null
$outputWorksheetName = $null
$workbookBackupPath = $null
$restoreWorkbookBackup = $false
$refreshExistingManagedTable = $false

try {
    if (-not (Test-Path -LiteralPath $workbookPath)) {
        throw "Workbook not found: $workbookPath"
    }

    if ($EnableWorkbookWrites) {
        $workbookBackupPath = [System.IO.Path]::Combine([System.IO.Path]::GetDirectoryName($workbookPath), ([System.IO.Path]::GetFileNameWithoutExtension($workbookPath) + '.sihao-backup' + [System.IO.Path]::GetExtension($workbookPath)))
        Copy-Item -LiteralPath $workbookPath -Destination $workbookBackupPath -Force
    }

    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.ScreenUpdating = $false
    $excel.EnableEvents = $false
    $excel.AskToUpdateLinks = $false

    try {
        $automationSecurity = [Microsoft.Office.Core.MsoAutomationSecurity]::msoAutomationSecurityForceDisable
        $excel.AutomationSecurity = $automationSecurity
    }
    catch {
    }

    $workbook = $excel.Workbooks.Open($workbookPath, 0, $openReadOnly)
    $sourceWorksheet = Try-GetWorksheet -Workbook $workbook -WorksheetName $sourceSheetName

    if ($null -eq $sourceWorksheet) {
        throw "Source worksheet '$sourceSheetName' was not found in '$workbookPath'."
    }

    $beforeSnapshot = Get-SourceSnapshot -Worksheet $sourceWorksheet
    Assert-HeadersMatch -Worksheet $sourceWorksheet -ExpectedHeaders $expectedHeaders
    $sourceLastRow = Get-SourceLastRow -Worksheet $sourceWorksheet
    if ($EnableWorkbookWrites) {
        $existingManagedWorksheet = Find-ManagedTableWorksheet -Workbook $workbook -ManagedTableName $tableName
        if (($null -ne $existingManagedWorksheet) -and (Test-PowerQueryMatchesDesiredFormula -Workbook $workbook -QueryName $queryName -NamedRangeName $sourceNamedRangeName)) {
            $outputWorksheet = $existingManagedWorksheet
            $existingManagedWorksheet = $null
            $refreshExistingManagedTable = $true
        }
        else {
            Release-ComObject $existingManagedWorksheet
            $outputWorksheet = Get-OrCreateManagedOutputWorksheet -Workbook $workbook -ManagedTableName $tableName -BaseSheetName $outputSheetBaseName
        }

        $outputWorksheetName = [string]$outputWorksheet.Name

        Add-OrUpdateSourceNamedRange -Workbook $workbook -RangeName $sourceNamedRangeName -SheetName $sourceSheetName -LastRow $sourceLastRow

        if (-not $refreshExistingManagedTable) {
            Remove-AllTablesFromWorksheet -Worksheet $outputWorksheet
            Clear-WorksheetContents -Worksheet $outputWorksheet
            Remove-WorkbookConnectionIfExists -Workbook $workbook -ConnectionNames @("Query - $queryName", $queryName)
            Remove-WorkbookQueryIfExists -Workbook $workbook -QueryName $queryName
            Add-OrUpdatePowerQuery -Workbook $workbook -QueryName $queryName -NamedRangeName $sourceNamedRangeName
        }

        $afterQuerySetupSnapshot = Get-SourceSnapshot -Worksheet $sourceWorksheet
        Assert-SnapshotUnchanged -Before $beforeSnapshot -After $afterQuerySetupSnapshot
        $workbook.Save()

        Release-ComObject $outputWorksheet
        $outputWorksheet = $null
        Release-ComObject $sourceWorksheet
        $sourceWorksheet = $null

        $workbook.Close($false)
        Release-ComObject $workbook
        $workbook = $null

        $workbook = $excel.Workbooks.Open($workbookPath, 0, $false)
        $sourceWorksheet = Try-GetWorksheet -Workbook $workbook -WorksheetName $sourceSheetName
        if ($null -eq $sourceWorksheet) {
            throw "Source worksheet '$sourceSheetName' was not found after reopening '$workbookPath'."
        }

        $outputWorksheet = Try-GetWorksheet -Workbook $workbook -WorksheetName $outputWorksheetName
        if ($null -eq $outputWorksheet) {
            throw "Managed output worksheet '$outputWorksheetName' was not found after reopening '$workbookPath'."
        }

        if ($refreshExistingManagedTable) {
            $outputTable = Refresh-ManagedOutputTable -Worksheet $outputWorksheet -ManagedTableName $tableName
        }
        else {
            $outputTable = Add-OrRefreshManagedOutputTable -Workbook $workbook -Worksheet $outputWorksheet -QueryName $queryName -ManagedTableName $tableName
        }
    }

    $afterSnapshot = Get-SourceSnapshot -Worksheet $sourceWorksheet
    Assert-SnapshotUnchanged -Before $beforeSnapshot -After $afterSnapshot

    if ($EnableWorkbookWrites) {
        $workbook.Save()
    }

    if ($EnableWorkbookWrites) {
        Format-ManagedOutputWorksheet -Worksheet $outputWorksheet -ListObject $outputTable
        Add-OrReplaceManagedSlicers -Workbook $workbook -Worksheet $outputWorksheet -ListObject $outputTable -ManagedTableName $tableName
        Assert-ManagedSlicersPresent -Workbook $workbook -Worksheet $outputWorksheet -ManagedTableName $tableName
        Freeze-ManagedOutputHeaderRow -Excel $excel -Worksheet $outputWorksheet
        $workbook.Save()
    }
}
catch {
    if ($EnableWorkbookWrites -and $null -ne $workbookBackupPath) {
        $restoreWorkbookBackup = $true
    }

    throw
}
finally {
    Release-ComObject $outputTable
    Release-ComObject $outputWorksheet
    Release-ComObject $sourceWorksheet

    if ($null -ne $workbook) {
        try {
            $workbook.Close($false)
        }
        catch {
        }

        Release-ComObject $workbook
    }

    if ($null -ne $excel) {
        try {
            $excel.Quit()
        }
        catch {
        }

        Release-ComObject $excel
    }

    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()

    if ($restoreWorkbookBackup -and $null -ne $workbookBackupPath -and (Test-Path -LiteralPath $workbookBackupPath)) {
        Copy-Item -LiteralPath $workbookBackupPath -Destination $workbookPath -Force
    }

    if (($null -ne $workbookBackupPath) -and (Test-Path -LiteralPath $workbookBackupPath)) {
        Remove-Item -LiteralPath $workbookBackupPath -Force
    }
}

if ($EnableWorkbookWrites) {
    Write-Host "Validated workbook '$workbookPath' in $executionMode mode, recreated named range '$sourceNamedRangeName', refreshed query '$queryName', loaded table '$tableName' on worksheet '$outputWorksheetName', added slicers for 時代/類別/提供者, froze the header row, and confirmed '$sourceSheetName' remained unchanged."
}
else {
    Write-Host "Validated workbook '$workbookPath' in $executionMode mode and confirmed '$sourceSheetName' remained unchanged."
}
