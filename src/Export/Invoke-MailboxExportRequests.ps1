Set-StrictMode -Version Latest

function New-ImtExportContentFilter {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory=$true)][datetime]$StartAt,
    [Parameter(Mandatory=$true)][datetime]$EndAt,
    [string]$SubjectContains,
    [string[]]$SubjectKeywords,
    [string[]]$SenderFilters,
    [string[]]$RecipientFilters,
    [switch]$OutboundOnly,
    [switch]$RequireAttachment
  )

  $dtStart = $StartAt.ToString('MM/dd/yyyy HH:mm:ss')
  $dtEnd = $EndAt.ToString('MM/dd/yyyy HH:mm:ss')
  $dateClause = if ($OutboundOnly) {
    "(Sent -ge '$dtStart' -and Sent -le '$dtEnd')"
  } else {
    "((Received -ge '$dtStart' -and Received -le '$dtEnd') -or (Sent -ge '$dtStart' -and Sent -le '$dtEnd'))"
  }

  $subjectParts = @()
  if (-not [string]::IsNullOrWhiteSpace($SubjectContains)) {
    $value = $SubjectContains.Replace("'", "''")
    $subjectParts += "(Subject -like '*$value*')"
  }

  if ($SubjectKeywords -and $SubjectKeywords.Count -gt 0) {
    $subjectParts += @(
      $SubjectKeywords |
        Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
        ForEach-Object {
          $keyword = $_.Trim().Replace("'", "''")
          "(Subject -like '*$keyword*')"
        }
    )
  }

  $content = if ($subjectParts.Count -gt 0) {
    "$dateClause -and ($($subjectParts -join ' -or '))"
  } else {
    $dateClause
  }

  $senderParts = @()
  if ($SenderFilters -and $SenderFilters.Count -gt 0) {
    $senderParts += @(
      $SenderFilters |
        Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
        ForEach-Object {
          $sender = $_.Trim().Replace("'", "''")
          "(Sender -eq '$sender')"
        }
    )
  }

  if ($senderParts.Count -gt 0) {
    $content = "$content -and ($($senderParts -join ' -or '))"
  }

  $recipientParts = @()
  if ($RecipientFilters -and $RecipientFilters.Count -gt 0) {
    $recipientParts += @(
      $RecipientFilters |
        Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
        ForEach-Object {
          $recipient = $_.Trim().Replace("'", "''")
          "((To -like '*$recipient*') -or (Cc -like '*$recipient*'))"
        }
    )
  }

  if ($recipientParts.Count -gt 0) {
    $content = "$content -and ($($recipientParts -join ' -or '))"
  }

  if ($RequireAttachment) {
    $content = "$content -and (HasAttachment -eq `$true)"
  }

  $content
}

function Invoke-ImtMailboxExportRequests {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory=$true)]$RunContext,
    [Parameter(Mandatory=$true)][string[]]$TargetAddresses,
    [string[]]$EffectiveSenderFilters
  )

  $resolved = Resolve-ImtMailboxesByAddressSet -Addresses $TargetAddresses
  $mailboxes = @($resolved.Mailboxes)

  if ($mailboxes.Count -eq 0) {
    return New-ImtModuleResult -StepName 'MailboxExport' -Status 'WARN' -Summary 'No mailboxes could be resolved for export from target addresses.' -Data ([pscustomobject]@{
      ExportRows = @()
      ExportCsv = $null
      Preflight = $null
    }) -Metrics @{
      MailboxCount = 0
      ExportRows = 0
    } -Errors @()
  }

  $hostingServers = @()
  foreach ($mailbox in $mailboxes) {
    $hostingServer = Get-ImtActiveMailboxServer -Mailbox $mailbox
    if ($hostingServer) {
      $hostingServers += $hostingServer
    }
  }
  $hostingServers = @($hostingServers | Sort-Object -Unique)

  $preflight = Test-ImtMailboxExportPrerequisites -PstRoot $RunContext.Inputs.ExportPstRoot -MailboxServers $hostingServers -SkipRemotePathValidation:$RunContext.Inputs.SkipDagPathValidation
  if (-not $preflight.Ready) {
    foreach ($issue in $preflight.Issues) {
      Write-ImtLog -Level WARN -Step 'MailboxExport' -EventType Progress -Message $issue
    }
    foreach ($warning in $preflight.Warnings) {
      Write-ImtLog -Level WARN -Step 'MailboxExport' -EventType Progress -Message $warning
    }

    return New-ImtModuleResult -StepName 'MailboxExport' -Status 'FAIL' -Summary 'Mailbox export prerequisites failed. Export requests were not created.' -Data ([pscustomobject]@{
      ExportRows = @()
      ExportCsv = $null
      Preflight = $preflight
    }) -Metrics @{
      MailboxCount = $mailboxes.Count
      ExportRows = 0
    } -Errors @($preflight.Issues)
  }

  $contentFilter = New-ImtExportContentFilter `
    -StartAt $RunContext.Start `
    -EndAt $RunContext.End `
    -SubjectContains $RunContext.Inputs.SubjectLike `
    -SubjectKeywords $RunContext.Inputs.Keywords `
    -SenderFilters $EffectiveSenderFilters `
    -RecipientFilters $RunContext.Inputs.Recipients `
    -OutboundOnly:$RunContext.Inputs.OutboundOnly `
    -RequireAttachment:$RunContext.Inputs.HasAttachmentOnly

  Write-ImtLog -Level DEBUG -Step 'MailboxExport' -EventType Progress -Message ("ContentFilter: {0}" -f $contentFilter)
  $exportRows = New-Object System.Collections.Generic.List[object]

  foreach ($mailbox in $mailboxes) {
    $mailboxIdentity = $mailbox.Identity
    $mailboxSmtp = $mailbox.PrimarySmtpAddress.ToString()
    $safeMailbox = ($mailboxSmtp -replace '[^\w@.-]','_')
    $pstPath = Join-Path $RunContext.Inputs.ExportPstRoot ("{0}_{1}.pst" -f $safeMailbox, $RunContext.Timestamp)
    $requestName = ("IMT_{0}_{1}" -f $safeMailbox, $RunContext.Timestamp)

    try {
      $request = New-MailboxExportRequest -Mailbox $mailboxIdentity -Name $requestName -FilePath $pstPath -ContentFilter $contentFilter -ErrorAction Stop
      [void]$exportRows.Add([pscustomobject]@{
        Mailbox = $mailboxSmtp
        Archive = $false
        RequestName = $request.Name
        Status = $request.Status
        FilePath = $pstPath
      })
      Write-ImtLog -Level DEBUG -Step 'MailboxExport' -EventType Progress -Message ("Created export request: {0} -> {1}" -f $request.Name, $pstPath)
    } catch {
      [void]$exportRows.Add([pscustomobject]@{
        Mailbox = $mailboxSmtp
        Archive = $false
        RequestName = $requestName
        Status = 'Failed'
        FilePath = $pstPath
      })
      Write-ImtLog -Level WARN -Step 'MailboxExport' -EventType Progress -Message ("Export request failed for {0}: {1}" -f $mailboxSmtp, $_.Exception.Message)
    }

    if ($RunContext.Inputs.IncludeArchive) {
      $archivePstPath = Join-Path $RunContext.Inputs.ExportPstRoot ("{0}_Archive_{1}.pst" -f $safeMailbox, $RunContext.Timestamp)
      $archiveRequestName = ("IMT_{0}_Archive_{1}" -f $safeMailbox, $RunContext.Timestamp)

      try {
        $archiveRequest = New-MailboxExportRequest -Mailbox $mailboxIdentity -IsArchive -Name $archiveRequestName -FilePath $archivePstPath -ContentFilter $contentFilter -ErrorAction Stop
        [void]$exportRows.Add([pscustomobject]@{
          Mailbox = $mailboxSmtp
          Archive = $true
          RequestName = $archiveRequest.Name
          Status = $archiveRequest.Status
          FilePath = $archivePstPath
        })
        Write-ImtLog -Level DEBUG -Step 'MailboxExport' -EventType Progress -Message ("Created archive export request: {0} -> {1}" -f $archiveRequest.Name, $archivePstPath)
      } catch {
        [void]$exportRows.Add([pscustomobject]@{
          Mailbox = $mailboxSmtp
          Archive = $true
          RequestName = $archiveRequestName
          Status = 'Failed'
          FilePath = $archivePstPath
        })
        Write-ImtLog -Level WARN -Step 'MailboxExport' -EventType Progress -Message ("Archive export failed for {0}: {1}" -f $mailboxSmtp, $_.Exception.Message)
      }
    }
  }

  $exportCsv = Join-Path $RunContext.OutputDir ("MTL_ExportRequests_{0}.csv" -f $RunContext.Timestamp)
  $exportRows | Export-Csv -Path $exportCsv -NoTypeInformation -Encoding UTF8

  $failed = @($exportRows | Where-Object { $_.Status -eq 'Failed' }).Count
  $status = if ($failed -gt 0) { 'WARN' } else { 'OK' }

  New-ImtModuleResult -StepName 'MailboxExport' -Status $status -Summary ("Export requests: {0}; Failed: {1}; Csv: {2}" -f $exportRows.Count, $failed, $exportCsv) -Data ([pscustomobject]@{
    ExportRows = $exportRows.ToArray()
    ExportCsv = $exportCsv
    Preflight = $preflight
  }) -Metrics @{
    MailboxCount = $mailboxes.Count
    ExportRows = $exportRows.Count
    FailedRows = $failed
  } -Errors @()
}
