function Get-SihaoExpectedHeaders {
    return @('時代', '類別', '藍圖名稱', '藍圖代碼', '提供者', '備註')
}

function Get-SihaoQueryName {
    return '四號谷地_1_表格查詢'
}

function Get-SihaoTableName {
    return 'SihaoValleyTable'
}

function Get-SihaoSourceNamedRangeName {
    return '四號谷地_1_來源範圍'
}

function Get-SihaoOutputSheetBaseName {
    return '四號谷地_1_表格'
}

function Get-DeterministicSheetName {
    param(
        [Parameter(Mandatory = $true)]
        [string]$BaseName,

        [AllowEmptyCollection()]
        [string[]]$ExistingNames
    )

    $normalizedBaseName = $BaseName.Substring(0, [Math]::Min($BaseName.Length, 31))

    if ($normalizedBaseName -notin $ExistingNames) {
        return $normalizedBaseName
    }

    $suffix = 2
    $candidateName = $null

    do {
        $suffixToken = "_{0}" -f $suffix
        $maxBaseLength = 31 - $suffixToken.Length
        $trimmedBaseName = $normalizedBaseName.Substring(0, [Math]::Min($normalizedBaseName.Length, $maxBaseLength))
        $candidateName = "{0}{1}" -f $trimmedBaseName, $suffixToken
        $suffix++
    } while ($candidateName -in $ExistingNames)

    return $candidateName
}

function Get-PreferredManagedSheetName {
    param(
        [Parameter(Mandatory = $true)]
        [string]$BaseName,

        [AllowEmptyCollection()]
        [string[]]$ExistingNames
    )

    if ($BaseName -in $ExistingNames) {
        return $BaseName
    }

    $matchingNames = @(
        $ExistingNames |
            Where-Object { $_ -match ('^{0}_(\d+)$' -f [regex]::Escape($BaseName)) } |
            Sort-Object {
                if ($_ -match '_(\d+)$') {
                    [int]$Matches[1]
                }
                else {
                    [int]::MaxValue
                }
            }
    )

    if ($matchingNames.Count -gt 0) {
        return $matchingNames[0]
    }

    return Get-DeterministicSheetName -BaseName $BaseName -ExistingNames $ExistingNames
}

function New-SourceNamedRangeFormula {
    param(
        [Parameter(Mandatory = $true)]
        [string]$SheetName,

        [Parameter(Mandatory = $true)]
        [ValidateRange(1, 1048576)]
        [int]$LastRow
    )

    $escapedSheetName = $SheetName -replace "'", "''"

    return ('=''{0}''!$A$1:$F${1}' -f $escapedSheetName, $LastRow)
}

function New-SihaoQueryFormula {
    param(
        [Parameter(Mandatory = $true)]
        [string]$NamedRangeName
    )

    $headersLiteral = (Get-SihaoExpectedHeaders | ForEach-Object { '"{0}"' -f $_ }) -join ', '

    return @"
let
    Source = Excel.CurrentWorkbook(){[Name="$NamedRangeName"]}[Content],
    PromotedHeaders = Table.PromoteHeaders(Source, [PromoteAllScalars=true]),
    SelectedColumns = Table.SelectColumns(PromotedHeaders, {$headersLiteral}),
    NormalizedColumns = Table.TransformColumns(SelectedColumns, {{"時代", each if _ is text then Text.Trim(_) else _, type nullable text}, {"類別", each if _ is text then Text.Trim(_) else _, type nullable text}, {"藍圖名稱", each if _ is text then Text.Trim(_) else _, type nullable text}, {"藍圖代碼", each if _ is text then Text.Trim(_) else _, type nullable text}, {"提供者", each if _ is text then Text.Trim(_) else _, type nullable text}, {"備註", each if _ is text then Text.Trim(_) else _, type nullable text}}),
    AddedSearchText = Table.AddColumn(NormalizedColumns, "搜尋文字", each Text.Combine(List.RemoveNulls({try Text.From([藍圖名稱]) otherwise null, try Text.From([備註]) otherwise null}), " "), type text)
in
    AddedSearchText
"@
}

Export-ModuleMember -Function Get-SihaoExpectedHeaders, Get-SihaoQueryName, Get-SihaoTableName, Get-SihaoSourceNamedRangeName, Get-SihaoOutputSheetBaseName, Get-DeterministicSheetName, Get-PreferredManagedSheetName, New-SourceNamedRangeFormula, New-SihaoQueryFormula
