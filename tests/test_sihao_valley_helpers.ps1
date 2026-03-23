$ErrorActionPreference = 'Stop'

$modulePath = Join-Path $PSScriptRoot '..\scripts\SihaoValleyTable.psm1'

if (-not (Test-Path $modulePath)) {
    throw "Module not found: $modulePath"
}

Import-Module $modulePath -Force

function Assert-Equal {
    param(
        [Parameter(Mandatory = $true)]$Actual,
        [Parameter(Mandatory = $true)]$Expected,
        [Parameter(Mandatory = $true)][string]$Message
    )

    if ($null -eq $Actual -and $null -eq $Expected) {
        return
    }

    if ($Actual -is [System.Array] -or $Expected -is [System.Array]) {
        $actualJson = ConvertTo-Json $Actual -Compress
        $expectedJson = ConvertTo-Json $Expected -Compress
        if ($actualJson -ne $expectedJson) {
            throw "$Message`nExpected: $expectedJson`nActual:   $actualJson"
        }
        return
    }

    if ($Actual -ne $Expected) {
        throw "$Message`nExpected: $Expected`nActual:   $Actual"
    }
}

function Assert-Contains {
    param(
        [Parameter(Mandatory = $true)][string]$Actual,
        [Parameter(Mandatory = $true)][string]$ExpectedSubstring,
        [Parameter(Mandatory = $true)][string]$Message
    )

    if (-not $Actual.Contains($ExpectedSubstring)) {
        throw "$Message`nMissing: $ExpectedSubstring`nActual:  $Actual"
    }
}

function Assert-DoesNotContain {
    param(
        [Parameter(Mandatory = $true)][string]$Actual,
        [Parameter(Mandatory = $true)][string]$UnexpectedSubstring,
        [Parameter(Mandatory = $true)][string]$Message
    )

    if ($Actual.Contains($UnexpectedSubstring)) {
        throw "$Message`nUnexpected: $UnexpectedSubstring`nActual:    $Actual"
    }
}

$expectedHeaders = @('時代', '類別', '藍圖名稱', '藍圖代碼', '提供者', '備註')
Assert-Equal (Get-SihaoExpectedHeaders) $expectedHeaders 'Get-SihaoExpectedHeaders returns the expected source headers.'
Assert-Equal (Get-SihaoQueryName) '四號谷地_1_表格查詢' 'Get-SihaoQueryName returns the fixed query name.'
Assert-Equal (Get-SihaoTableName) 'SihaoValleyTable' 'Get-SihaoTableName returns the fixed table name.'
Assert-Equal (Get-SihaoSourceNamedRangeName) '四號谷地_1_來源範圍' 'Get-SihaoSourceNamedRangeName returns the fixed source named range name.'
Assert-Equal (Get-SihaoOutputSheetBaseName) '四號谷地_1_表格' 'Get-SihaoOutputSheetBaseName returns the fixed output sheet base name.'

Assert-Equal (Get-DeterministicSheetName -BaseName '四號谷地_1_表格' -ExistingNames @()) '四號谷地_1_表格' 'Get-DeterministicSheetName keeps the base name when unused.'
Assert-Equal (Get-DeterministicSheetName -BaseName '四號谷地_1_表格' -ExistingNames @('四號谷地_1_表格')) '四號谷地_1_表格_2' 'Get-DeterministicSheetName appends a deterministic suffix when the base name is already used.'
Assert-Equal (Get-DeterministicSheetName -BaseName '四號谷地_1_表格' -ExistingNames @('四號谷地_1_表格', '四號谷地_1_表格_2', '四號谷地_1_表格_3')) '四號谷地_1_表格_4' 'Get-DeterministicSheetName increments past existing suffixed names.'
Assert-Equal (Get-PreferredManagedSheetName -BaseName '四號谷地_1_表格' -ExistingNames @('資料', '四號谷地_1_表格_8', '其他')) '四號谷地_1_表格_8' 'Get-PreferredManagedSheetName reuses an existing deterministic managed output sheet when present.'
Assert-Equal (Get-PreferredManagedSheetName -BaseName '四號谷地_1_表格' -ExistingNames @('資料', '四號谷地_1_表格_10', '四號谷地_1_表格_2')) '四號谷地_1_表格_2' 'Get-PreferredManagedSheetName prefers the earliest deterministic suffix when multiple stale managed sheets exist.'
Assert-Equal (Get-PreferredManagedSheetName -BaseName '四號谷地_1_表格' -ExistingNames @('資料', '其他')) '四號谷地_1_表格' 'Get-PreferredManagedSheetName falls back to a new deterministic sheet name when no managed output sheet exists yet.'

$uniqueThirtyOneCharacterName = '1234567890123456789012345678901'
$uniqueTrimmedSheetName = Get-DeterministicSheetName -BaseName $uniqueThirtyOneCharacterName -ExistingNames @()
Assert-Equal $uniqueTrimmedSheetName '1234567890123456789012345678901' 'Get-DeterministicSheetName preserves a unique 31-character worksheet name.'
Assert-Equal $uniqueTrimmedSheetName.Length 31 'Get-DeterministicSheetName keeps a unique 31-character worksheet name within Excel''s limit.'

$uniqueOverLimitName = '12345678901234567890123456789012'
$uniqueCappedSheetName = Get-DeterministicSheetName -BaseName $uniqueOverLimitName -ExistingNames @()
Assert-Equal $uniqueCappedSheetName '1234567890123456789012345678901' 'Get-DeterministicSheetName trims an over-limit unique worksheet name to 31 characters.'
Assert-Equal $uniqueCappedSheetName.Length 31 'Get-DeterministicSheetName enforces Excel''s 31-character limit for unique worksheet names.'

$thirtyOneCharacterName = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ12345'
$trimmedSheetName = Get-DeterministicSheetName -BaseName $thirtyOneCharacterName -ExistingNames @($thirtyOneCharacterName)
Assert-Equal $trimmedSheetName 'ABCDEFGHIJKLMNOPQRSTUVWXYZ123_2' 'Get-DeterministicSheetName trims long names before adding a suffix.'
Assert-Equal $trimmedSheetName.Length 31 'Get-DeterministicSheetName keeps suffixed worksheet names within Excel''s 31-character limit.'

$normalNamedRangeFormula = New-SourceNamedRangeFormula -SheetName '四號谷地_1' -LastRow 27
Assert-DoesNotContain $normalNamedRangeFormula 'LET(' 'New-SourceNamedRangeFormula avoids LET for broader Excel compatibility.'
Assert-DoesNotContain $normalNamedRangeFormula 'COUNTA(' 'New-SourceNamedRangeFormula does not rely on COUNTA, which would truncate ranges after blanks.'
Assert-DoesNotContain $normalNamedRangeFormula 'LOOKUP(' 'New-SourceNamedRangeFormula leaves the last-row calculation outside the workbook formula.'
Assert-Contains $normalNamedRangeFormula '''四號谷地_1''!$A$1:$F$27' 'New-SourceNamedRangeFormula returns the explicit A:F range for the computed last row.'

$apostropheNamedRangeFormula = New-SourceNamedRangeFormula -SheetName "Sihao's Data" -LastRow 42
Assert-Contains $apostropheNamedRangeFormula '''Sihao''''s Data''!$A$1:$F$42' 'New-SourceNamedRangeFormula escapes apostrophes in sheet names for the returned source range.'

$queryFormula = New-SihaoQueryFormula -NamedRangeName '四號谷地_1_來源範圍'
Assert-Contains $queryFormula 'Excel.CurrentWorkbook(){[Name="四號谷地_1_來源範圍"]}[Content]' 'New-SihaoQueryFormula reads from the named range.'
Assert-Contains $queryFormula 'Table.PromoteHeaders' 'New-SihaoQueryFormula promotes the header row.'
Assert-Contains $queryFormula '{"時代", "類別", "藍圖名稱", "藍圖代碼", "提供者", "備註"}' 'New-SihaoQueryFormula enforces the expected source headers.'
Assert-DoesNotContain $queryFormula 'MissingField.UseNull' 'New-SihaoQueryFormula does not hide missing headers with MissingField.UseNull.'
Assert-Contains $queryFormula 'Table.TransformColumns' 'New-SihaoQueryFormula applies light normalization to the source columns.'
Assert-Contains $queryFormula 'each if _ is text then Text.Trim(_) else _' 'New-SihaoQueryFormula trims text values safely without altering non-text blanks.'
Assert-Contains $queryFormula '{"時代", each if _ is text then Text.Trim(_) else _, type nullable text}' 'New-SihaoQueryFormula normalizes 時代 with safe trimming.'
Assert-Contains $queryFormula '{"備註", each if _ is text then Text.Trim(_) else _, type nullable text}' 'New-SihaoQueryFormula normalizes 備註 with safe trimming.'
Assert-Contains $queryFormula '搜尋文字' 'New-SihaoQueryFormula adds the helper column.'
Assert-Contains $queryFormula 'try Text.From([藍圖名稱]) otherwise null' 'New-SihaoQueryFormula safely converts 藍圖名稱 when building 搜尋文字.'
Assert-Contains $queryFormula 'try Text.From([備註]) otherwise null' 'New-SihaoQueryFormula safely converts 備註 when building 搜尋文字.'

Write-Host 'All sihao valley helper tests passed.'
