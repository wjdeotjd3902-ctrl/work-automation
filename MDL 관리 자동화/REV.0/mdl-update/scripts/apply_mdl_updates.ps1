[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)][string]$PlanPath,
    [Parameter(Mandatory = $true)][string]$LogPath,
    [switch]$Apply
)

$ErrorActionPreference = 'Stop'

function Get-ColumnNumber([string]$Letters) {
    $number = 0
    foreach ($character in $Letters.ToUpper().ToCharArray()) {
        $number = ($number * 26) + ([int]$character - [int][char]'A' + 1)
    }
    return $number
}

function Get-MergeList($Worksheet) {
    $set = [System.Collections.Generic.HashSet[string]]::new()
    foreach ($cell in @($Worksheet.UsedRange.Cells)) {
        if ($cell.MergeCells) { [void]$set.Add([string]$cell.MergeArea.Address($false, $false)) }
    }
    return @($set | Sort-Object)
}

function Get-LayoutFingerprint($Worksheet) {
    $parts = [System.Collections.Generic.List[string]]::new()
    foreach ($merge in (Get-MergeList $Worksheet)) { $parts.Add("M:$merge") }
    for ($row = 1; $row -le [int]$Worksheet.UsedRange.Rows.Count; $row++) {
        $parts.Add("R:${row}:$($Worksheet.Rows.Item($row).RowHeight)")
    }
    for ($col = 1; $col -le [int]$Worksheet.UsedRange.Columns.Count; $col++) {
        $parts.Add("C:${col}:$($Worksheet.Columns.Item($col).ColumnWidth)")
    }
    foreach ($cell in @($Worksheet.UsedRange.Cells)) {
        $borders = foreach ($edge in 7, 8, 9, 10) {
            $border = $cell.Borders.Item($edge)
            "$edge,$($border.LineStyle),$($border.Weight),$($border.Color)"
        }
        $parts.Add("B:$($cell.Address($false,$false)):$($borders -join ';')")
    }
    $text = $parts -join "`n"
    $bytes = [Text.Encoding]::UTF8.GetBytes($text)
    $sha = [Security.Cryptography.SHA256]::Create()
    try { return ([BitConverter]::ToString($sha.ComputeHash($bytes))).Replace('-', '').ToLowerInvariant() }
    finally { $sha.Dispose() }
}

function Assert-WorkbookUnlocked([string]$Path) {
    try {
        $stream = [IO.File]::Open($Path, 'Open', 'ReadWrite', 'None')
        $stream.Dispose()
    } catch {
        throw "MDL 파일이 열려 있거나 잠겨 있습니다. Excel을 닫고 다시 실행하세요: $Path"
    }
}

function Resolve-TargetColumn($Profile, $Update) {
    $issue = @($Profile.issues | Where-Object { [int]$_.issue -eq [int]$Update.issue })
    if ($issue.Count -ne 1) { throw "Issue mapping not found or duplicated: $($Update.issue)" }
    switch ([string]$Update.field) {
        'revision' { return [string]$issue[0].revision }
        'submissionDate' { return [string]$issue[0].submissionDate }
        'reviewerReceiptDate' {
            $reviewer = $issue[0].reviewers.PSObject.Properties[[string]$Update.reviewer].Value
            return [string]$reviewer.receiptDate
        }
        'reviewerStatus' {
            $reviewer = $issue[0].reviewers.PSObject.Properties[[string]$Update.reviewer].Value
            return [string]$reviewer.status
        }
        default { throw "Unsupported field: $($Update.field)" }
    }
}

$plan = Get-Content -LiteralPath $PlanPath -Raw -Encoding utf8 | ConvertFrom-Json
$profile = Get-Content -LiteralPath ([string]$plan.profilePath) -Raw -Encoding utf8 | ConvertFrom-Json
$workbookPath = [IO.Path]::GetFullPath([string]$plan.workbookPath)
if (-not (Test-Path -LiteralPath $workbookPath -PathType Leaf)) { throw "Workbook not found: $workbookPath" }
$updates = @($plan.updates)
if ($updates.Count -eq 0) { throw 'Update plan contains no updates.' }
if ($Apply) { Assert-WorkbookUnlocked $workbookPath }

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$workbook = $null
$backupPath = $null
$resultItems = @()
$failures = @()

try {
    $workbook = $excel.Workbooks.Open($workbookPath, 0, (-not $Apply))
    $worksheet = $workbook.Worksheets.Item([string]$profile.sheetName)

    foreach ($required in @($profile.requiredHeaders)) {
        $actual = [string]$worksheet.Range([string]$required.cell).Value2
        if ($actual -ne [string]$required.value) {
            throw "MDL profile mismatch at $($required.cell): expected '$($required.value)', found '$actual'"
        }
    }
    $actualMerges = @(Get-MergeList $worksheet)
    $expectedMerges = @($profile.mergedRanges | Sort-Object)
    if (($actualMerges -join '|') -ne ($expectedMerges -join '|')) { throw 'MDL merged-cell layout differs from the project profile.' }

    $docColumn = Get-ColumnNumber ([string]$profile.documentNumberColumn)
    $rowMap = @{}
    for ($row = [int]$profile.firstDataRow; $row -le [int]$profile.lastDataRow; $row++) {
        $docNo = ([string]$worksheet.Cells.Item($row, $docColumn).Value2).Trim()
        if ($docNo) {
            if ($rowMap.ContainsKey($docNo)) { throw "Duplicate document number in MDL: $docNo" }
            $rowMap[$docNo] = $row
        }
    }

    $plannedCells = [System.Collections.Generic.HashSet[string]]::new()
    foreach ($update in $updates) {
        try {
            $docNo = ([string]$update.docNo).Trim()
            if (-not $rowMap.ContainsKey($docNo)) { throw "Document not found: $docNo" }
            if ([string]::IsNullOrWhiteSpace([string]$update.value)) { throw 'Blank values must be omitted from the plan.' }
            if ([string]::IsNullOrWhiteSpace([string]$update.evidence)) { throw 'Evidence is required.' }
            $columnLetters = Resolve-TargetColumn -Profile $profile -Update $update
            if (-not $columnLetters) { throw "Column mapping is blank for $($update.field)." }
            $column = Get-ColumnNumber $columnLetters
            $cell = $worksheet.Cells.Item([int]$rowMap[$docNo], $column)
            $address = [string]$cell.Address($false, $false)
            if (-not $plannedCells.Add($address)) { throw "Duplicate target cell in plan: $address" }

            $newValue = [string]$update.value
            if ($update.field -in @('submissionDate', 'reviewerReceiptDate')) {
                $parsed = [datetime]::ParseExact($newValue, 'yyyy-MM-dd', [Globalization.CultureInfo]::InvariantCulture)
                $writeValue = $parsed.ToOADate()
                $comparisonValue = $parsed.ToString('yyyy-MM-dd')
                $existingComparison = if ($null -eq $cell.Value2 -or [string]$cell.Value2 -eq '') { '' } else { ([datetime]::FromOADate([double]$cell.Value2)).ToString('yyyy-MM-dd') }
            } else {
                $writeValue = $newValue
                $comparisonValue = $newValue
                $existingComparison = ([string]$cell.Value2).Trim()
            }
            if ($update.field -eq 'reviewerStatus' -and @($profile.allowedStatuses).Count -gt 0 -and $newValue -notin @($profile.allowedStatuses)) {
                throw "Status '$newValue' is not allowed by the project profile."
            }
            if ($existingComparison -and $existingComparison -ne $comparisonValue -and -not [bool]$update.replaceExisting) {
                throw "Existing value conflict at ${address}: '$existingComparison' -> '$comparisonValue'"
            }

            $resultItems += [pscustomobject]@{
                docNo = $docNo
                issue = [int]$update.issue
                field = [string]$update.field
                reviewer = [string]$update.reviewer
                cell = $address
                oldValue = $existingComparison
                newValue = $comparisonValue
                evidence = [string]$update.evidence
                writeValue = $writeValue
            }
        } catch {
            $failures += [pscustomobject]@{ docNo = [string]$update.docNo; issue = $update.issue; field = [string]$update.field; error = $_.Exception.Message }
        }
    }

    if ($failures.Count -gt 0) { throw "Dry Run validation failed for $($failures.Count) update(s)." }
    $layoutBefore = Get-LayoutFingerprint $worksheet

    if ($Apply) {
        $backupFolder = Join-Path (Split-Path -Parent $workbookPath) '_MDL_Backup'
        if (-not (Test-Path -LiteralPath $backupFolder)) { New-Item -ItemType Directory -Path $backupFolder | Out-Null }
        $backupName = '{0}_{1}{2}' -f [IO.Path]::GetFileNameWithoutExtension($workbookPath), (Get-Date -Format 'yyyyMMdd_HHmmss'), [IO.Path]::GetExtension($workbookPath)
        $backupPath = Join-Path $backupFolder $backupName
        Copy-Item -LiteralPath $workbookPath -Destination $backupPath

        foreach ($item in $resultItems) { $worksheet.Range([string]$item.cell).Value2 = $item.writeValue }
        $workbook.Save()
        $workbook.Close($true)
        $workbook = $null

        $workbook = $excel.Workbooks.Open($workbookPath, 0, $true)
        $worksheet = $workbook.Worksheets.Item([string]$profile.sheetName)
        $layoutAfter = Get-LayoutFingerprint $worksheet
        if ($layoutAfter -ne $layoutBefore) { throw 'Layout verification failed after save.' }
        foreach ($item in $resultItems) {
            $cell = $worksheet.Range([string]$item.cell)
            if ($item.field -in @('submissionDate', 'reviewerReceiptDate')) {
                $actual = ([datetime]::FromOADate([double]$cell.Value2)).ToString('yyyy-MM-dd')
            } else { $actual = ([string]$cell.Value2).Trim() }
            if ($actual -ne [string]$item.newValue) { throw "Saved value verification failed at $($item.cell)." }
        }
    }
}
catch {
    $fatal = $_.Exception.Message
    if ($null -ne $workbook) { $workbook.Close($false); $workbook = $null }
    if ($Apply -and $backupPath -and (Test-Path -LiteralPath $backupPath)) {
        Copy-Item -LiteralPath $backupPath -Destination $workbookPath -Force
    }
    $failures += [pscustomobject]@{ docNo = $null; issue = $null; field = $null; error = $fatal }
}
finally {
    if ($null -ne $workbook) { $workbook.Close($false) }
    $excel.Quit()
    [Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel) | Out-Null
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}

$publicItems = @($resultItems | Select-Object docNo,issue,field,reviewer,cell,oldValue,newValue,evidence)
$result = [pscustomobject]@{
    mode = if ($Apply) { 'APPLY' } else { 'DRY_RUN' }
    project = [string]$profile.project
    workbookPath = $workbookPath
    backupPath = $backupPath
    requested = $updates.Count
    valid = $publicItems.Count
    failed = $failures.Count
    changes = $publicItems
    failures = $failures
}
$parent = Split-Path -Parent $LogPath
if ($parent -and -not (Test-Path -LiteralPath $parent)) { New-Item -ItemType Directory -Path $parent -Force | Out-Null }
$result | ConvertTo-Json -Depth 7 | Set-Content -LiteralPath $LogPath -Encoding utf8
$result | Select-Object mode,project,workbookPath,backupPath,requested,valid,failed | ConvertTo-Json -Compress
if ($failures.Count -gt 0) { exit 1 }
