[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)][string]$PlanPath,
    [Parameter(Mandatory = $true)][string]$LogPath,
    [switch]$Apply
)

$ErrorActionPreference = 'Stop'

function Get-OutlookFolder {
    param($Namespace, [string]$MailboxName, [string]$RelativePath)
    $folder = $Namespace.Folders.Item($MailboxName)
    foreach ($part in ($RelativePath -split '[\\/]' | Where-Object { $_ })) {
        $folder = $folder.Folders.Item($part)
    }
    return $folder
}

function Get-NewCategories {
    param([string]$Existing, [string]$Classification)
    $kept = @($Existing -split '\s*,\s*' | Where-Object {
        $_ -and $_ -notin @('접수', '제출')
    })
    return (@($kept) + $Classification) -join ', '
}

$plan = Get-Content -LiteralPath $PlanPath -Raw -Encoding utf8 | ConvertFrom-Json
$items = @($plan.items)
if (-not $plan.mailbox -or -not $plan.sourceFolder -or -not $plan.targetFolder) {
    throw 'Plan must contain mailbox, sourceFolder, and targetFolder.'
}
if ($items.Count -eq 0) { throw 'Plan contains no items.' }

$duplicateIds = @($items | Group-Object entryId | Where-Object Count -gt 1)
if ($duplicateIds.Count -gt 0) { throw 'Plan contains duplicate EntryID values.' }

$invalid = @($items | Where-Object { $_.classification -notin @('접수', '제출') -or -not $_.entryId })
if ($invalid.Count -gt 0) { throw 'Every item must contain an EntryID and classification 접수 or 제출.' }

$outlook = New-Object -ComObject Outlook.Application
$namespace = $outlook.GetNamespace('MAPI')
$source = Get-OutlookFolder -Namespace $namespace -MailboxName ([string]$plan.mailbox) -RelativePath ([string]$plan.sourceFolder)
$target = Get-OutlookFolder -Namespace $namespace -MailboxName ([string]$plan.mailbox) -RelativePath ([string]$plan.targetFolder)
$validated = @()
$failures = @()

foreach ($planned in $items) {
    try {
        $item = if ($planned.storeId) {
            $namespace.GetItemFromID([string]$planned.entryId, [string]$planned.storeId)
        } else {
            $namespace.GetItemFromID([string]$planned.entryId)
        }
        if ($null -eq $item -or [int]$item.Class -ne 43) { throw 'Mail item not found.' }
        $parentId = [string]$item.Parent.EntryID
        if ($parentId -ne [string]$source.EntryID -and $parentId -ne [string]$target.EntryID) {
            throw "Mail is outside the planned source/target folders: $($item.Parent.FolderPath)"
        }

        $record = [ordered]@{
            entryId = [string]$planned.entryId
            classification = [string]$planned.classification
            subject = [string]$item.Subject
            previousCategories = [string]$item.Categories
            newCategories = Get-NewCategories -Existing ([string]$item.Categories) -Classification ([string]$planned.classification)
            previousFolder = [string]$item.Parent.FolderPath
            result = 'validated'
            newEntryId = $null
        }

        if ($Apply) {
            $item.Categories = $record.newCategories
            $item.Save()
            if ([string]$item.Parent.EntryID -ne [string]$target.EntryID) {
                $moved = $item.Move($target)
                $record.newEntryId = [string]$moved.EntryID
            } else {
                $record.newEntryId = [string]$item.EntryID
            }
            $record.result = 'applied'
        }
        $validated += [pscustomobject]$record
    } catch {
        $failures += [pscustomobject]@{
            entryId = [string]$planned.entryId
            classification = [string]$planned.classification
            error = $_.Exception.Message
        }
    }
}

$result = [pscustomobject]@{
    mode = if ($Apply) { 'APPLY' } else { 'DRY_RUN' }
    mailbox = [string]$plan.mailbox
    sourceFolder = [string]$source.FolderPath
    targetFolder = [string]$target.FolderPath
    requested = $items.Count
    receiptCount = @($items | Where-Object classification -eq '접수').Count
    submissionCount = @($items | Where-Object classification -eq '제출').Count
    successCount = $validated.Count
    failureCount = $failures.Count
    items = $validated
    failures = $failures
}

$parent = Split-Path -Parent $LogPath
if ($parent -and -not (Test-Path -LiteralPath $parent)) {
    New-Item -ItemType Directory -Path $parent -Force | Out-Null
}
$result | ConvertTo-Json -Depth 7 | Set-Content -LiteralPath $LogPath -Encoding utf8
$result | Select-Object mode,mailbox,sourceFolder,targetFolder,requested,receiptCount,submissionCount,successCount,failureCount | ConvertTo-Json -Compress

