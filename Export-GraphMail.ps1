
#Requires -Version 5.1
#Requires -Modules Microsoft.Graph.Authentication, Microsoft.Graph.Mail

<#
.SYNOPSIS
Optimised Microsoft Graph email archiver.

.DESCRIPTION
Pass 1:
- Retrieves messages without the Filed category.
- Exports them as MIME .eml files.
- Adds the Filed category.

Pass 2:
- Retrieves only messages older than OlderThanDate.
- Checks their Filed category.
- Moves filed messages to Deleted Items.

Features:
- Year/month archive directories based on sent date.
- Nested progress bars.
- Server-side Graph filtering.
- Category filter fallback if Graph rejects the query.
- Local category verification to prevent incorrect processing.
- No additional GET request before category updates.
- Temporary downloads and no-overwrite protection.
- Multiple mailboxes and shared mailbox support.
- WhatIf and Verbose support.

Only root-level folders are matched.
Do not run concurrent instances against the same folders/destination.
#>

[CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'Medium')]
param(
    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string]$FolderNames,

    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string]$DestinationPath,

    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string[]]$MailboxIds,

    [datetime]$OlderThanDate = (Get-Date).AddDays(-365),

    [switch]$SharedMailbox
)

# ============================================================
# CONFIGURATION
# ============================================================

$ErrorActionPreference = 'Stop'

$startTime = Get-Date

$stats = [ordered]@{
    Examined       = 0
    Exported       = 0
    AlreadyFiled   = 0
    Moved          = 0
    WouldExport    = 0
    WouldMove      = 0
    Skipped        = 0
    Errors         = 0
}

# ============================================================
# DISPLAY FUNCTIONS
# ============================================================

function Write-Banner {
    param([string]$Message)

    Write-Host ''
    Write-Host '============================================================' -ForegroundColor DarkCyan
    Write-Host "  $Message" -ForegroundColor Cyan
    Write-Host '============================================================' -ForegroundColor DarkCyan
    Write-Host ''
}

function Write-Status {
    param(
        [string]$Message,
        [ConsoleColor]$Color = [ConsoleColor]::Gray
    )

    $time = Get-Date -Format 'HH:mm:ss'

    Write-Host "[$time] " -NoNewline -ForegroundColor DarkGray
    Write-Host $Message -ForegroundColor $Color
}

function Format-Elapsed {
    param([timespan]$Elapsed)

    if ($Elapsed.TotalDays -ge 1) {
        return $Elapsed.ToString('d\.hh\:mm\:ss')
    }

    return $Elapsed.ToString('hh\:mm\:ss')
}

function Update-ArchiveProgress {
    param(
        [int]$FolderIndex,
        [int]$FolderCount,
        [string]$FolderName,
        [int]$PassIndex,
        [int]$MessageIndex,
        [int]$MessageCount,
        [string]$Operation,
        [string]$Subject
    )

    $fraction = if ($MessageCount -gt 0) {
        $MessageIndex / $MessageCount
    }
    else {
        0
    }

    # Two passes per folder.
    $overallPercent = [math]::Min(
        100,
        [math]::Round(
            (
                (
                    ($FolderIndex - 1) * 2 +
                    ($PassIndex - 1) +
                    $fraction
                ) / ($FolderCount * 2)
            ) * 100
        )
    )

    $folderPercent = [math]::Round($fraction * 100)

    $passName = if ($PassIndex -eq 1) {
        'Export unfiled messages'
    }
    else {
        'Move old filed messages'
    }

    $overallStatus = "Folder $FolderIndex of $FolderCount | Examined: $($stats.Examined) | Exported: $($stats.Exported) | Moved: $($stats.Moved) | Errors: $($stats.Errors)"

    Write-Progress `
        -Id 1 `
        -Activity 'MICROSOFT GRAPH EMAIL ARCHIVER' `
        -Status $overallStatus `
        -PercentComplete $overallPercent `
        -CurrentOperation $passName

    $subjectDisplay = $Subject

    if ($subjectDisplay.Length -gt 75) {
        $subjectDisplay = $subjectDisplay.Substring(0, 72) + '...'
    }

    if ([string]::IsNullOrWhiteSpace($subjectDisplay)) {
        $subjectDisplay = '(No subject)'
    }

    Write-Progress `
        -Id 2 `
        -ParentId 1 `
        -Activity "$passName - $FolderName" `
        -Status "Message $MessageIndex of $MessageCount ($folderPercent%)" `
        -PercentComplete $folderPercent `
        -CurrentOperation "$Operation | $subjectDisplay"
}

# ============================================================
# GRAPH MESSAGE RETRIEVAL
# ============================================================

function Get-ArchiveMessages {
    param(
        [string]$MailboxId,
        [string]$FolderId,
        [string]$Filter,
        [string]$Description,
        [switch]$AllowCategoryFallback
    )

    Write-Status "Querying Graph: $Description" Cyan

    try {

        $result = @(
            Get-MgUserMailFolderMessage `
                -UserId $MailboxId `
                -MailFolderId $FolderId `
                -All `
                -PageSize 100 `
                -Filter $Filter `
                -Property 'id,subject,sentDateTime,categories' `
                -Headers $headers `
                -ErrorAction Stop
        )

        return $result
    }
    catch {

        if (-not $AllowCategoryFallback) {
            throw
        }

        Write-Warning "Graph rejected the category filter. Falling back to an unfiltered folder query."
        Write-Warning "Reason: $($_.Exception.Message)"

        $result = @(
            Get-MgUserMailFolderMessage `
                -UserId $MailboxId `
                -MailFolderId $FolderId `
                -All `
                -PageSize 100 `
                -Property 'id,subject,sentDateTime,categories' `
                -Headers $headers `
                -ErrorAction Stop
        )

        return $result
    }
}

# ============================================================
# VALIDATE INPUT
# ============================================================

$requestedNames = @(
    $FolderNames.Split(',') |
        ForEach-Object { $_.Trim() } |
        Where-Object { $_ } |
        Select-Object -Unique
)

if ($requestedNames.Count -eq 0) {
    throw 'Provide at least one mail folder name.'
}

$MailboxIds = @(
    $MailboxIds |
        ForEach-Object { $_.Trim() } |
        Where-Object { $_ } |
        Select-Object -Unique
)

if ($MailboxIds.Count -eq 0) {
    throw 'Provide at least one mailbox.'
}

$destination = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath(
    $DestinationPath
)

$cutoffUtc = $OlderThanDate.ToUniversalTime()

# ISO 8601 UTC value for Graph OData filtering.
$cutoffFilter = $cutoffUtc.ToString(
    "yyyy-MM-ddTHH:mm:ssZ",
    [Globalization.CultureInfo]::InvariantCulture
)

$headers = @{
    Prefer = 'IdType="ImmutableId"'
}

$scopes = @('Mail.ReadWrite')

if ($SharedMailbox) {
    $scopes += 'Mail.ReadWrite.Shared'
}

# ============================================================
# START
# ============================================================

Write-Banner 'MICROSOFT GRAPH EMAIL ARCHIVER - OPTIMISED'

Write-Status 'Initialising archive operation...' Cyan

Write-Status "Destination: $destination"
Write-Status "Archive cutoff: $OlderThanDate"
Write-Status "Mailboxes: $($MailboxIds -join ', ')"
Write-Status "Folders: $($requestedNames -join ', ')"

if ($WhatIfPreference) {
    Write-Status 'WHATIF MODE ENABLED - No changes will be made.' Yellow
}

if ($SharedMailbox) {
    Write-Status 'Shared mailbox access enabled.' Cyan
}

if (-not $WhatIfPreference) {
    [System.IO.Directory]::CreateDirectory($destination) | Out-Null
}

# ============================================================
# CONNECT TO GRAPH
# ============================================================

Write-Host ''
Write-Status 'Connecting to Microsoft Graph...' Cyan

Connect-MgGraph `
    -Scopes $scopes `
    -ContextScope Process `
    -NoWelcome `
    -ErrorAction Stop

Write-Status 'Microsoft Graph connection established.' Green

# ============================================================
# DISCOVER FOLDERS
# ============================================================

Write-Host ''
Write-Status 'Discovering mail folders...' Cyan

$folders = @(
    foreach ($mailbox in $MailboxIds) {

        Write-Status "Searching mailbox: $mailbox" DarkGray

        $candidates = @(
            Get-MgUserMailFolder `
                -UserId $mailbox `
                -All `
                -Property 'id,displayName' `
                -Headers $headers `
                -ErrorAction Stop
        )

        $deletedItems = Get-MgUserMailFolder `
            -UserId $mailbox `
            -MailFolderId 'deleteditems' `
            -Property 'id' `
            -Headers $headers `
            -ErrorAction Stop

        foreach ($candidate in $candidates) {

            if ($requestedNames -contains $candidate.DisplayName) {

                [pscustomobject]@{
                    MailboxId = $mailbox
                    Id = $candidate.Id
                    Name = $candidate.DisplayName
                    DisplayPath = "$mailbox/$($candidate.DisplayName)"
                    IsDeletedItems = ($candidate.Id -eq $deletedItems.Id)
                }
            }
        }
    }
)

foreach ($name in $requestedNames) {

    if ($name -notin @($folders | ForEach-Object { $_.Name })) {
        throw "Mail folder '$name' not found. No emails processed."
    }
}

Write-Status "Found $($folders.Count) matching folders." Green

foreach ($folder in $folders) {
    Write-Status "  $($folder.DisplayPath)" DarkGray
}

# ============================================================
# PROCESS FOLDERS
# ============================================================

$folderCount = $folders.Count
$folderIndex = 0

foreach ($folder in $folders) {

    $folderIndex++

    Write-Host ''
    Write-Host '------------------------------------------------------------' -ForegroundColor DarkCyan
    Write-Host " FOLDER $folderIndex OF $folderCount" -ForegroundColor Cyan
    Write-Host " $($folder.DisplayPath)" -ForegroundColor White
    Write-Host '------------------------------------------------------------' -ForegroundColor DarkCyan

    # ========================================================
    # PASS 1 - EXPORT UNFILED MESSAGES
    # ========================================================

    Write-Host ''
    Write-Status 'PASS 1: Searching for unfiled messages...' Cyan

    $unfiledFilter = "not(categories/any(c:c eq 'Filed'))"

    $messages = @(
        Get-ArchiveMessages `
            -MailboxId $folder.MailboxId `
            -FolderId $folder.Id `
            -Filter $unfiledFilter `
            -Description 'Messages without Filed category' `
            -AllowCategoryFallback
    )

    $messageCount = $messages.Count

    Write-Status "Pass 1 returned $messageCount message(s)." Green

    if ($messageCount -gt 0) {

        $messageIndex = 0

        foreach ($email in $messages) {

            $messageIndex++
            $stats.Examined++

            $filePath = $null
            $temporaryPath = $null

            $subjectDisplay = [string]$email.Subject

            if ([string]::IsNullOrWhiteSpace($subjectDisplay)) {
                $subjectDisplay = '(No subject)'
            }

            try {

                Update-ArchiveProgress `
                    -FolderIndex $folderIndex `
                    -FolderCount $folderCount `
                    -FolderName $folder.Name `
                    -PassIndex 1 `
                    -MessageIndex $messageIndex `
                    -MessageCount $messageCount `
                    -Operation 'Checking message' `
                    -Subject $subjectDisplay

                # Always verify the category locally.
                # This protects against a server-side filter being
                # ignored or behaving differently than expected.

                if (@($email.Categories) -contains 'Filed') {

                    $stats.AlreadyFiled++
                    Write-Verbose "Skipping already filed message: $subjectDisplay"
                    continue
                }

                if ($null -eq $email.SentDateTime) {
                    throw 'Message has no sent timestamp.'
                }

                $sent = [datetimeoffset]$email.SentDateTime
                $localSent = $sent.LocalDateTime

                # ------------------------------------------------
                # YEAR/MONTH ARCHIVE DIRECTORY
                # ------------------------------------------------

                $yearFolder = $localSent.ToString('yyyy')
                $monthFolder = $localSent.ToString('MM')

                $archiveFolder = Join-Path $destination $yearFolder
                $archiveFolder = Join-Path $archiveFolder $monthFolder

                # ------------------------------------------------
                # GENERATE FILENAME
                # ------------------------------------------------

                $timestamp = $localSent.ToString('yyyyMMdd HHmmss')

                $subject = [string]$email.Subject -replace '[\\/:*?"<>|\x00-\x1F]', ''

                $subject = $subject.Trim().TrimEnd('.') -replace ' {2,}', ' '

                if ([string]::IsNullOrWhiteSpace($subject)) {
                    $subject = 'No subject'
                }

                $prefix = "$timestamp "
                $extension = '.eml'

                $basePath = Join-Path $archiveFolder "$prefix$extension"

                $maxSubjectLength = 240 - $basePath.Length

                if ($maxSubjectLength -lt 1) {
                    throw "Destination path is too long: $archiveFolder"
                }

                if ($subject.Length -gt $maxSubjectLength) {
                    $subject = $subject.Substring(0, $maxSubjectLength).TrimEnd()
                }

                $fileName = "$prefix$subject$extension"
                $filePath = Join-Path $archiveFolder $fileName

                if (Test-Path -LiteralPath $filePath) {
                    throw "Refusing to overwrite an existing email file: $filePath"
                }

                # ------------------------------------------------
                # CONFIRM EXPORT
                # ------------------------------------------------

                if (-not $PSCmdlet.ShouldProcess(
                    $filePath,
                    'Export message and add Filed category'
                )) {

                    if ($WhatIfPreference) {

                        $stats.WouldExport++

                        # An old unfiled message would also be moved
                        # after its export and categorisation.

                        if (
                            $sent.UtcDateTime -lt $cutoffUtc -and
                            -not $folder.IsDeletedItems
                        ) {
                            $stats.WouldMove++
                        }
                    }
                    else {
                        $stats.Skipped++
                    }

                    continue
                }

                # ------------------------------------------------
                # CREATE YEAR/MONTH DIRECTORY
                # ------------------------------------------------

                [System.IO.Directory]::CreateDirectory(
                    $archiveFolder
                ) | Out-Null

                # ------------------------------------------------
                # DOWNLOAD MIME
                # ------------------------------------------------

                Update-ArchiveProgress `
                    -FolderIndex $folderIndex `
                    -FolderCount $folderCount `
                    -FolderName $folder.Name `
                    -PassIndex 1 `
                    -MessageIndex $messageIndex `
                    -MessageCount $messageCount `
                    -Operation 'Downloading MIME message' `
                    -Subject $subjectDisplay

                Write-Verbose "Downloading: $fileName"

                $temporaryPath = Join-Path $archiveFolder (
                    '.' + [guid]::NewGuid().ToString('N') + '.tmp'
                )

                Get-MgUserMessageContent `
                    -UserId $folder.MailboxId `
                    -MessageId $email.Id `
                    -OutFile $temporaryPath `
                    -Headers $headers `
                    -ErrorAction Stop | Out-Null

                if ((Get-Item -LiteralPath $temporaryPath).Length -eq 0) {
                    throw 'Graph returned an empty MIME file.'
                }

                # Publish only after successful download.
                # File.Move refuses to overwrite existing files.

                [System.IO.File]::Move(
                    $temporaryPath,
                    $filePath
                )

                $temporaryPath = $null

                # ------------------------------------------------
                # UPDATE CATEGORY
                # ------------------------------------------------

                Update-ArchiveProgress `
                    -FolderIndex $folderIndex `
                    -FolderCount $folderCount `
                    -FolderName $folder.Name `
                    -PassIndex 1 `
                    -MessageIndex $messageIndex `
                    -MessageCount $messageCount `
                    -Operation 'Adding Filed category' `
                    -Subject $subjectDisplay

                # Use categories from the initial message listing.
                # This avoids one additional GET per exported message.

                $categories = @(
                    @($email.Categories) + @('Filed') |
                        Where-Object { $_ } |
                        Select-Object -Unique
                )

                Update-MgUserMessage `
                    -UserId $folder.MailboxId `
                    -MessageId $email.Id `
                    -BodyParameter @{ categories = $categories } `
                    -Headers $headers `
                    -ErrorAction Stop

                $stats.Exported++

                Write-Status "EXPORTED: $yearFolder\$monthFolder\$fileName" Green

                Write-Verbose "Exported and categorized: $fileName"
            }
            catch {

                $failure = $_
                $stats.Errors++

                Write-Host ''
                Write-Host '!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!' -ForegroundColor Red
                Write-Host ' EXPORT PROCESSING ERROR' -ForegroundColor Red
                Write-Host '!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!' -ForegroundColor Red

                Write-Host "Mail folder:    $($folder.DisplayPath)" -ForegroundColor Yellow
                Write-Host "Message number: $messageIndex of $messageCount"
                Write-Host "Output path:    $filePath"
                Write-Host "Subject:        $($email.Subject)"
                Write-Host "Sent time:  $($email.SentDateTime)"
                Write-Host "Message ID:     $($email.Id)"
                Write-Host "Categories:     $($email.Categories -join ', ')"

                Write-Host ''
                Write-Host "Error ID: $($failure.FullyQualifiedErrorId)" -ForegroundColor Red
                Write-Host $failure.InvocationInfo.PositionMessage
                Write-Host $failure.ScriptStackTrace
                Write-Host ($failure.Exception.ToString())
                Write-Host ''
            }
            finally {

                if (
                    $temporaryPath -and
                    (Test-Path -LiteralPath $temporaryPath)
                ) {
                    Remove-Item `
                        -LiteralPath $temporaryPath `
                        -Force `
                        -ErrorAction SilentlyContinue
                }
            }
        }
    }
    else {
        Write-Status 'No unfiled messages found.' Green
    }

    # ========================================================
    # PASS 2 - MOVE OLD FILED MESSAGES
    # ========================================================

    Write-Host ''
    Write-Status "PASS 2: Searching for filed messages sent before $OlderThanDate..." Cyan

    if ($folder.IsDeletedItems) {

        Write-Status 'This is the Deleted Items folder. Skipping deletion pass.' Yellow
        $oldMessages = @()
    }
    else {

        # sentDateTime filtering is supported by the Graph
        # message listing endpoint.

        $oldFilter = "sentDateTime lt $cutoffFilter"

        $oldMessages = @(
            Get-ArchiveMessages `
                -MailboxId $folder.MailboxId `
                -FolderId $folder.Id `
                -Filter $oldFilter `
                -Description "Messages sent before $cutoffFilter"
        )
    }

    $oldMessageCount = $oldMessages.Count

    Write-Status "Pass 2 returned $oldMessageCount old message(s)." Green

    if ($oldMessageCount -gt 0) {

        $oldIndex = 0

        foreach ($email in $oldMessages) {

            $oldIndex++
            $stats.Examined++

            $subjectDisplay = [string]$email.Subject

            if ([string]::IsNullOrWhiteSpace($subjectDisplay)) {
                $subjectDisplay = '(No subject)'
            }

            try {

                Update-ArchiveProgress `
                    -FolderIndex $folderIndex `
                    -FolderCount $folderCount `
                    -FolderName $folder.Name `
                    -PassIndex 2 `
                    -MessageIndex $oldIndex `
                    -MessageCount $oldMessageCount `
                    -Operation 'Checking retention eligibility' `
                    -Subject $subjectDisplay

                if ($null -eq $email.SentDateTime) {
                    throw 'Message has no sent timestamp.'
                }

                $sent = [datetimeoffset]$email.SentDateTime

                # Verify cutoff locally.

                if ($sent.UtcDateTime -ge $cutoffUtc) {
                    continue
                }

                # Only messages already carrying Filed are eligible.
                # Newly exported messages from Pass 1 are also now Filed.

                if (-not (@($email.Categories) -contains 'Filed')) {
                    Write-Verbose "Skipping old but unfiled message: $subjectDisplay"
                    continue
                }

                # ------------------------------------------------
                # MOVE TO DELETED ITEMS
                # ------------------------------------------------

                if ($PSCmdlet.ShouldProcess(
                    "$($folder.DisplayPath): $subjectDisplay",
                    'Move old filed message to Deleted Items'
                )) {

                    Update-ArchiveProgress `
                        -FolderIndex $folderIndex `
                        -FolderCount $folderCount `
                        -FolderName $folder.Name `
                        -PassIndex 2 `
                        -MessageIndex $oldIndex `
                        -MessageCount $oldMessageCount `
                        -Operation 'Moving to Deleted Items' `
                        -Subject $subjectDisplay

                    Move-MgUserMessage `
                        -UserId $folder.MailboxId `
                        -MessageId $email.Id `
                        -BodyParameter @{ destinationId = 'deleteditems' } `
                        -Headers $headers `
                        -ErrorAction Stop | Out-Null

                    $stats.Moved++

                    Write-Status "MOVED TO DELETED ITEMS: $subjectDisplay" Yellow
                }
                elseif ($WhatIfPreference) {

                    $stats.WouldMove++
                }
                else {

                    $stats.Skipped++
                }
            }
            catch {

                $failure = $_
                $stats.Errors++

                Write-Host ''
                Write-Host '!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!' -ForegroundColor Red
                Write-Host ' MOVE PROCESSING ERROR' -ForegroundColor Red
                Write-Host '!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!' -ForegroundColor Red

                Write-Host "Mail folder:    $($folder.DisplayPath)" -ForegroundColor Yellow
                Write-Host "Message number: $oldIndex of $oldMessageCount"
                Write-Host "Subject:        $($email.Subject)"
                Write-Host "Sent time:  $($email.SentDateTime)"
                Write-Host "Message ID:     $($email.Id)"
                Write-Host "Categories:     $($email.Categories -join ', ')"

                Write-Host ''
                Write-Host "Error ID: $($failure.FullyQualifiedErrorId)" -ForegroundColor Red
                Write-Host $failure.InvocationInfo.PositionMessage
                Write-Host $failure.ScriptStackTrace
                Write-Host ($failure.Exception.ToString())
                Write-Host ''
            }
        }
    }
    else {
        Write-Status 'No old messages found.' Green
    }

    # --------------------------------------------------------
    # FOLDER COMPLETED
    # --------------------------------------------------------

    Write-Progress `
        -Id 2 `
        -Activity "Processing folder: $($folder.Name)" `
        -Completed

    Write-Status "Finished folder: $($folder.DisplayPath)" Green
}

# ============================================================
# COMPLETE PROGRESS BARS
# ============================================================

Write-Progress -Id 2 -Activity 'Processing messages' -Completed

Write-Progress -Id 1 -Activity 'MICROSOFT GRAPH EMAIL ARCHIVER' -Completed

# ============================================================
# FINAL SUMMARY
# ============================================================

$elapsed = (Get-Date) - $startTime

$rate = if ($elapsed.TotalSeconds -gt 0) {
    [math]::Round($stats.Examined / $elapsed.TotalSeconds, 2)
}
else {
    0
}

Write-Banner 'ARCHIVING COMPLETED'

Write-Host '  PROCESSING SUMMARY' -ForegroundColor Cyan
Write-Host '  ------------------' -ForegroundColor DarkGray

Write-Host ("  Messages examined:       {0,8}" -f $stats.Examined)

Write-Host ("  Successfully exported:   {0,8}" -f $stats.Exported) -ForegroundColor Green

Write-Host ("  Already filed:           {0,8}" -f $stats.AlreadyFiled)

Write-Host ("  Moved to Deleted Items:  {0,8}" -f $stats.Moved) -ForegroundColor Yellow

Write-Host ("  Skipped:                 {0,8}" -f $stats.Skipped)

Write-Host ("  Processing errors:       {0,8}" -f $stats.Errors) -ForegroundColor $(if ($stats.Errors -gt 0) { 'Red' } else { 'Green' })

if ($WhatIfPreference) {

    Write-Host ''
    Write-Host '  WHATIF PREVIEW' -ForegroundColor Yellow

    Write-Host ("  Would export:            {0,8}" -f $stats.WouldExport)

    Write-Host ("  Would move:              {0,8}" -f $stats.WouldMove)
}

Write-Host ''
Write-Host '  PERFORMANCE' -ForegroundColor Cyan
Write-Host '  -----------' -ForegroundColor DarkGray

Write-Host "  Elapsed time:            $(Format-Elapsed $elapsed)"

Write-Host "  Processing rate:         $rate message evaluations/sec"

Write-Host ''
Write-Host "  Destination: $destination" -ForegroundColor Gray
Write-Host "  Cutoff date: $OlderThanDate" -ForegroundColor Gray

Write-Host ''
Write-Host '============================================================' -ForegroundColor DarkCyan

if ($stats.Errors -gt 0) {
    Write-Host "  COMPLETED WITH $($stats.Errors) ERROR(S)" -ForegroundColor Red
}
else {
    Write-Host '  ALL OPERATIONS COMPLETED SUCCESSFULLY' -ForegroundColor Green
}

Write-Host '============================================================'
Write-Host ''

if ($stats.Errors -gt 0) {
    throw "Completed with $($stats.Errors) message processing error(s). Review the output above."
}
