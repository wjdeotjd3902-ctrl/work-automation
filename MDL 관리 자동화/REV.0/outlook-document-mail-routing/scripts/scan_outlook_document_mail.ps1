[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)][string]$Mailbox,
    [Parameter(Mandatory = $true)][string]$SourceFolder,
    [Parameter(Mandatory = $true)][datetime]$StartDate,
    [Parameter(Mandatory = $true)][datetime]$EndDate,
    [Parameter(Mandatory = $true)][string]$OutputPath,
    [int]$MaxBodyChars = 6000
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

function Get-NewBody {
    param([string]$Body)
    if ([string]::IsNullOrWhiteSpace($Body)) { return '' }
    $cut = $Body.Length
    $markers = @(
        '(?im)^\s*-----Original Message-----\s*$',
        '(?im)^\s*From:\s*',
        '(?im)^\s*보낸 사람:\s*',
        '(?im)^\s*발신:\s*'
    )
    foreach ($marker in $markers) {
        $match = [regex]::Match($Body, $marker)
        if ($match.Success -and $match.Index -lt $cut) { $cut = $match.Index }
    }
    $newBody = $Body.Substring(0, $cut).Trim()
    if ($newBody.Length -gt $MaxBodyChars) { return $newBody.Substring(0, $MaxBodyChars) }
    return $newBody
}

if ($EndDate.Date -lt $StartDate.Date) { throw 'EndDate must be on or after StartDate.' }

$outlook = New-Object -ComObject Outlook.Application
$namespace = $outlook.GetNamespace('MAPI')
$folder = Get-OutlookFolder -Namespace $namespace -MailboxName $Mailbox -RelativePath $SourceFolder
$from = $StartDate.Date
$until = $EndDate.Date.AddDays(1)
$scanned = 0
$readFailures = @()
$mails = @()

foreach ($item in @($folder.Items)) {
    try {
        if ([int]$item.Class -ne 43) { continue }
        $received = [datetime]$item.ReceivedTime
        if ($received -lt $from -or $received -ge $until) { continue }
        $scanned++
        $attachments = @()
        for ($i = 1; $i -le [int]$item.Attachments.Count; $i++) {
            $attachment = $item.Attachments.Item($i)
            $attachments += [pscustomobject]@{
                index = $i
                name = [string]$attachment.FileName
                size = [int64]$attachment.Size
            }
        }
        $mails += [pscustomobject]@{
            entryId = [string]$item.EntryID
            storeId = [string]$folder.StoreID
            sourceFolder = [string]$folder.FolderPath
            received = $received.ToString('yyyy-MM-dd HH:mm:ss')
            sent = ([datetime]$item.SentOn).ToString('yyyy-MM-dd HH:mm:ss')
            senderName = [string]$item.SenderName
            senderEmail = [string]$item.SenderEmailAddress
            to = [string]$item.To
            cc = [string]$item.CC
            subject = [string]$item.Subject
            categories = [string]$item.Categories
            newBody = Get-NewBody -Body ([string]$item.Body)
            attachments = $attachments
        }
    } catch {
        $readFailures += [pscustomobject]@{
            subject = [string]$item.Subject
            error = $_.Exception.Message
        }
    }
}

$result = [pscustomobject]@{
    mailbox = $Mailbox
    sourceFolder = [string]$folder.FolderPath
    startDate = $from.ToString('yyyy-MM-dd')
    endDate = $EndDate.Date.ToString('yyyy-MM-dd')
    scanned = $scanned
    readFailureCount = $readFailures.Count
    readFailures = $readFailures
    mails = $mails
}

$parent = Split-Path -Parent $OutputPath
if ($parent -and -not (Test-Path -LiteralPath $parent)) {
    New-Item -ItemType Directory -Path $parent -Force | Out-Null
}
$result | ConvertTo-Json -Depth 7 | Set-Content -LiteralPath $OutputPath -Encoding utf8
$result | Select-Object mailbox,sourceFolder,startDate,endDate,scanned,readFailureCount | ConvertTo-Json -Compress

