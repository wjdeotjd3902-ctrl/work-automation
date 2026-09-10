[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)][string]$Mailbox,
    [Parameter(Mandatory = $true)][string]$FolderPath,
    [Parameter(Mandatory = $true)][datetime]$StartDate,
    [Parameter(Mandatory = $true)][datetime]$EndDate,
    [Parameter(Mandatory = $true)][string]$Destination,
    [string[]]$Categories = @(),
    [switch]$IncludeMsg,
    [switch]$IncludeInlineImages,
    [switch]$Apply
)

$ErrorActionPreference = 'Stop'

function Get-ShortHash([string]$Text, [int]$Length = 10) {
    $sha = [Security.Cryptography.SHA256]::Create()
    try {
        $bytes = [Text.Encoding]::UTF8.GetBytes($Text)
        $hex = -join ($sha.ComputeHash($bytes) | ForEach-Object { $_.ToString('x2') })
        return $hex.Substring(0, $Length)
    }
    finally { $sha.Dispose() }
}

function Get-OutlookFolder($Root, [string]$Path) {
    $current = $Root
    $segments = @($Path -split '[\\/]' | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
    foreach ($segment in $segments) { $current = $current.Folders.Item($segment) }
    return $current
}

function Get-Property($Attachment, [string]$Schema) {
    try { return $Attachment.PropertyAccessor.GetProperty($Schema) } catch { return $null }
}

function Test-Category([string]$MailCategories, [string[]]$Wanted) {
    if ($Wanted.Count -eq 0) { return $true }
    $actual = @($MailCategories -split '[,;]' | ForEach-Object { $_.Trim() } | Where-Object { $_ })
    foreach ($category in $Wanted) { if ($actual -contains $category) { return $true } }
    return $false
}

if ($EndDate.Date -lt $StartDate.Date) { throw 'EndDate must be on or after StartDate.' }
$resolvedDestination = [IO.Path]::GetFullPath($Destination)
if ($resolvedDestination.Length -gt 120) {
    throw "Destination is too long ($($resolvedDestination.Length) characters). Choose a shallow root such as C:\OA\Project."
}

$outlook = New-Object -ComObject Outlook.Application
$namespace = $outlook.GetNamespace('MAPI')
$store = $namespace.Folders.Item($Mailbox)
$folder = Get-OutlookFolder $store $FolderPath
$endExclusive = $EndDate.Date.AddDays(1)
$imageExtensions = @('.bmp', '.gif', '.jpeg', '.jpg', '.png', '.svg', '.tif', '.tiff', '.webp')

$manifest = [Collections.Generic.List[object]]::new()
$mailCount = 0
$planned = 0
$saved = 0
$skipped = 0
$excluded = 0
$failed = 0
$readFailures = 0

foreach ($item in $folder.Items) {
    if ($item.Class -ne 43) { continue }
    try {
        $received = [datetime]$item.ReceivedTime
        if ($received -lt $StartDate.Date -or $received -ge $endExclusive) { continue }
        if (-not (Test-Category ([string]$item.Categories) $Categories)) { continue }
        $mailCount++
        $mailKey = '{0}_{1}' -f $received.ToString('yyyyMMdd_HHmmss'), (Get-ShortHash ([string]$item.EntryID))
        $mailDirectory = Join-Path $resolvedDestination $mailKey

        for ($index = 1; $index -le $item.Attachments.Count; $index++) {
            $attachment = $item.Attachments.Item($index)
            $originalName = [string]$attachment.FileName
            $extension = [IO.Path]::GetExtension($originalName).ToLowerInvariant()
            $reason = $null

            if (-not $IncludeMsg -and ($extension -eq '.msg' -or $attachment.Type -eq 5)) {
                $reason = 'embedded-message'
            }
            elseif (-not $IncludeInlineImages -and $imageExtensions -contains $extension) {
                $hidden = Get-Property $attachment 'http://schemas.microsoft.com/mapi/proptag/0x7FFE000B'
                $contentId = Get-Property $attachment 'http://schemas.microsoft.com/mapi/proptag/0x3712001F'
                $flags = Get-Property $attachment 'http://schemas.microsoft.com/mapi/proptag/0x37140003'
                if ($hidden -eq $true -or -not [string]::IsNullOrWhiteSpace([string]$contentId) -or (($flags -as [int]) -band 4)) {
                    $reason = 'inline-image'
                }
            }

            if ($reason) {
                $excluded++
                $manifest.Add([pscustomobject]@{
                    received = $received.ToString('yyyy-MM-dd HH:mm:ss'); subject = [string]$item.Subject
                    entryId = [string]$item.EntryID; attachmentIndex = $index; originalName = $originalName
                    savedPath = $null; size = [int64]$attachment.Size; result = 'excluded'; reason = $reason
                })
                continue
            }

            $planned++
            $shortName = 'a{0:D3}_{1}{2}' -f $index, (Get-ShortHash $originalName 8), $extension
            $savePath = Join-Path $mailDirectory $shortName
            if ($savePath.Length -gt 240) { throw "Generated save path exceeds 240 characters: $savePath" }
            $result = 'planned'
            $failureReason = $null

            if ($Apply) {
                try {
                    New-Item -ItemType Directory -Path $mailDirectory -Force | Out-Null
                    if (Test-Path -LiteralPath $savePath) {
                        $existing = Get-Item -LiteralPath $savePath
                        if ($existing.Length -eq [int64]$attachment.Size) {
                            $result = 'skipped-existing'
                            $skipped++
                        }
                        else {
                            $version = 2
                            do {
                                $candidate = Join-Path $mailDirectory ('a{0:D3}_{1}_v{2}{3}' -f $index, (Get-ShortHash $originalName 8), $version, $extension)
                                $version++
                            } while (Test-Path -LiteralPath $candidate)
                            $savePath = $candidate
                            $attachment.SaveAsFile($savePath)
                            $result = 'saved-versioned'
                            $saved++
                        }
                    }
                    else {
                        $attachment.SaveAsFile($savePath)
                        $result = 'saved'
                        $saved++
                    }
                }
                catch {
                    $result = 'failed'
                    $failureReason = $_.Exception.Message
                    $failed++
                }
            }

            $manifest.Add([pscustomobject]@{
                received = $received.ToString('yyyy-MM-dd HH:mm:ss'); subject = [string]$item.Subject
                entryId = [string]$item.EntryID; attachmentIndex = $index; originalName = $originalName
                savedPath = $savePath; size = [int64]$attachment.Size; result = $result; reason = $failureReason
            })
        }
    }
    catch {
        $readFailures++
        $manifest.Add([pscustomobject]@{
            received = $null; subject = [string]$item.Subject; entryId = [string]$item.EntryID
            attachmentIndex = $null; originalName = $null; savedPath = $null; size = $null
            result = 'mail-read-failed'; reason = $_.Exception.Message
        })
    }
}

if ($Apply) {
    New-Item -ItemType Directory -Path $resolvedDestination -Force | Out-Null
    $manifest | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath (Join-Path $resolvedDestination 'manifest.json') -Encoding utf8
    $manifest | Export-Csv -LiteralPath (Join-Path $resolvedDestination 'manifest.csv') -NoTypeInformation -Encoding utf8
}

[pscustomobject]@{
    mode = $(if ($Apply) { 'apply' } else { 'dry-run' })
    folder = [string]$folder.FolderPath
    startDate = $StartDate.ToString('yyyy-MM-dd')
    endDate = $EndDate.ToString('yyyy-MM-dd')
    matchedMails = $mailCount
    plannedAttachments = $planned
    saved = $saved
    skipped = $skipped
    excluded = $excluded
    failed = $failed
    mailReadFailures = $readFailures
    destination = $resolvedDestination
} | ConvertTo-Json -Compress
