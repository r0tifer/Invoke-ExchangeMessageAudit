Set-StrictMode -Version Latest

function Export-ImtTrackingReports {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory=$true)]$RunContext,
    [Parameter(Mandatory=$true)][AllowEmptyCollection()][object[]]$Results,
    [Parameter(Mandatory=$true)][AllowEmptyCollection()][string[]]$BaseTargetAddresses,
    [AllowEmptyCollection()][object[]]$ClientAttributionRows,
    [AllowEmptyCollection()][object[]]$ClientAuditRows,
    [AllowEmptyCollection()][object[]]$ClientProtocolRows
  )

  $resultRowSet = @($Results)
  $clientAttributionRowSet = @($ClientAttributionRows)
  $clientAuditRowSet = @($ClientAuditRows)
  $clientProtocolRowSet = @($ClientProtocolRows)
  $inputKeywords = @($RunContext.Inputs.Keywords)

  if (@($resultRowSet).Count -eq 0) {
    return New-ImtModuleResult -StepName 'TrackingReport' -Status 'WARN' -Summary 'No results to report/export.' -Data ([pscustomobject]@{
      CsvMain = $null
      ClientAttributionCsv = $null
      ClientAuditCsv = $null
      ClientProtocolCsv = $null
      ClientAttributionRows = @()
      ClientAuditRows = @()
      ClientProtocolRows = @()
      TrackingKeywordRows = @()
      TrackingKeywordMailboxRows = @()
      DailyCounts = @()
    }) -Metrics @{
      ResultCount = 0
    } -Errors @()
  }

  $csvMain = Join-Path $RunContext.OutputDir ("MTL_{0}_from-{1}_{2}.csv" -f $RunContext.SafeRecipient, $RunContext.SafeSender, $RunContext.Timestamp)
  $resultRowSet |
    Sort-Object Timestamp |
    Select-Object Timestamp,EventId,Source,ServerHostname,ClientHostname,ConnectorId,
                  Sender,
                  @{n='Recipients';e={($_.Recipients -join ';')}},
                  MessageSubject,MessageId,InternalMessageId,RecipientStatus,TotalBytes,SourceContext |
    Export-Csv -Path $csvMain -NoTypeInformation -Encoding UTF8

  $clientAttributionCsv = $null
  if (@($clientAttributionRowSet).Count -gt 0) {
    $clientAttributionCsv = Join-Path $RunContext.OutputDir ("MTL_ClientAttribution_{0}.csv" -f $RunContext.Timestamp)
    $clientAttributionRowSet |
      Sort-Object Mailbox,SubmittedAt,Subject |
      Export-Csv -Path $clientAttributionCsv -NoTypeInformation -Encoding UTF8
  }

  $clientAuditCsv = $null
  if (@($clientAuditRowSet).Count -gt 0) {
    $clientAuditCsv = Join-Path $RunContext.OutputDir ("MTL_ClientAttribution_Audit_{0}.csv" -f $RunContext.Timestamp)
    $clientAuditRowSet |
      Select-Object LastAccessed,Operation,OperationResult,LogonType,MailboxOwnerUPN,LogonUserDisplayName,
                    ItemSubject,ClientInfoString,ClientIPAddress,ClientMachineName,ClientProcessName,
                    ClientVersion,FolderPathName |
      Export-Csv -Path $clientAuditCsv -NoTypeInformation -Encoding UTF8
  }

  $clientProtocolCsv = $null
  if (@($clientProtocolRowSet).Count -gt 0) {
    $clientProtocolCsv = Join-Path $RunContext.OutputDir ("MTL_ClientAttribution_Protocol_{0}.csv" -f $RunContext.Timestamp)
    $clientProtocolRowSet |
      Sort-Object Mailbox,Timestamp,EvidenceType |
      Export-Csv -Path $clientProtocolCsv -NoTypeInformation -Encoding UTF8
  }

  $dailyCounts = @(
    $resultRowSet |
      Group-Object { $_.Timestamp.Date } |
      Sort-Object Name |
      Select-Object @{n='Date';e={$_.Name.ToString('yyyy-MM-dd')}}, @{n='Count';e={$_.Count}}
  )

  $trackingKeywordRows = @()
  $trackingKeywordMailboxRows = @()

  if ($RunContext.Inputs.Keywords -and @($inputKeywords).Count -gt 0) {
    $normalizedKeywords = @(
      $inputKeywords |
        Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
        ForEach-Object { $_.Trim() } |
        Select-Object -Unique
    )

    $keywordRows = @()
    $keywordMailboxRows = @()

    foreach ($keyword in $normalizedKeywords) {
      $hits = @()
      foreach ($event in $resultRowSet) {
        $subject = ($event.MessageSubject -as [string])
        if ([string]::IsNullOrWhiteSpace($subject)) {
          continue
        }
        if ($subject.IndexOf($keyword, [System.StringComparison]::OrdinalIgnoreCase) -ge 0) {
          $hits += $event
        }
      }

      $eventCount = @($hits).Count
      $messageSet = @{}
      foreach ($hit in $hits) {
        $messageId = ($hit.MessageId -as [string])
        if (-not [string]::IsNullOrWhiteSpace($messageId)) {
          $messageSet[$messageId.ToLowerInvariant()] = $true
        }
      }

      $keywordRows += [pscustomobject]@{
        Keyword = $keyword
        EventHitCount = $eventCount
        DistinctMessageIdHitCount = @($messageSet.Keys).Count
      }
    }

    $resolvedMailboxAddresses = @(
      $BaseTargetAddresses |
        ForEach-Object { $_ -as [string] } |
        Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
        ForEach-Object { $_.ToLowerInvariant() } |
        Select-Object -Unique
    )

    foreach ($mailboxAddress in $resolvedMailboxAddresses) {
      foreach ($keyword in $normalizedKeywords) {
        $mailboxHits = @()
        foreach ($event in $resultRowSet) {
          $subject = ($event.MessageSubject -as [string])
          if ([string]::IsNullOrWhiteSpace($subject)) { continue }
          if ($subject.IndexOf($keyword, [System.StringComparison]::OrdinalIgnoreCase) -lt 0) { continue }

          $senderHit = $false
          $senderValue = ($event.Sender -as [string])
          if ($senderValue) {
            $senderHit = ($senderValue.ToLowerInvariant() -eq $mailboxAddress)
          }

          $recipientHit = $false
          if ($event.Recipients) {
            foreach ($recipient in $event.Recipients) {
              $recipientValue = ($recipient -as [string])
              if ($recipientValue -and ($recipientValue.ToLowerInvariant() -eq $mailboxAddress)) {
                $recipientHit = $true
                break
              }
            }
          }

          if ($senderHit -or $recipientHit) {
            $mailboxHits += $event
          }
        }

        $mailboxMessageSet = @{}
        foreach ($mailboxHit in $mailboxHits) {
          $messageId = ($mailboxHit.MessageId -as [string])
          if (-not [string]::IsNullOrWhiteSpace($messageId)) {
            $mailboxMessageSet[$messageId.ToLowerInvariant()] = $true
          }
        }

        $keywordMailboxRows += [pscustomobject]@{
          Mailbox = $mailboxAddress
          Keyword = $keyword
          EventHitCount = @($mailboxHits).Count
          DistinctMessageIdHitCount = @($mailboxMessageSet.Keys).Count
        }
      }
    }

    if (@($keywordRows).Count -gt 0) {
      $trackingKeywordRows = @($keywordRows)
      $keywordCsv = Join-Path $RunContext.OutputDir ("MTL_KeywordHits_{0}.csv" -f $RunContext.Timestamp)
      $trackingKeywordRows | Sort-Object Keyword | Export-Csv -Path $keywordCsv -NoTypeInformation -Encoding UTF8
      Write-ImtLog -Level DEBUG -Step 'TrackingReport' -EventType Progress -Message ("Keyword hit summary: {0}" -f $keywordCsv)
    }

    if (@($keywordMailboxRows).Count -gt 0) {
      $trackingKeywordMailboxRows = @($keywordMailboxRows)
      $keywordMailboxCsv = Join-Path $RunContext.OutputDir ("MTL_KeywordHits_ByMailbox_{0}.csv" -f $RunContext.Timestamp)
      $trackingKeywordMailboxRows | Sort-Object Mailbox,Keyword | Export-Csv -Path $keywordMailboxCsv -NoTypeInformation -Encoding UTF8
      Write-ImtLog -Level DEBUG -Step 'TrackingReport' -EventType Progress -Message ("Keyword-by-mailbox summary: {0}" -f $keywordMailboxCsv)
    }
  }

  New-ImtModuleResult -StepName 'TrackingReport' -Status 'OK' -Summary ("Tracking report exported: {0}" -f $csvMain) -Data ([pscustomobject]@{
    CsvMain = $csvMain
    ClientAttributionCsv = $clientAttributionCsv
    ClientAuditCsv = $clientAuditCsv
    ClientProtocolCsv = $clientProtocolCsv
    ClientAttributionRows = @($clientAttributionRowSet)
    ClientProtocolRows = @($clientProtocolRowSet)
    TrackingKeywordRows = @($trackingKeywordRows)
    TrackingKeywordMailboxRows = @($trackingKeywordMailboxRows)
    DailyCounts = @($dailyCounts)
  }) -Metrics @{
    ResultCount = @($resultRowSet).Count
    ClientAttributionRows = @($clientAttributionRowSet).Count
    ClientAuditRows = @($clientAuditRowSet).Count
    ClientProtocolRows = @($clientProtocolRowSet).Count
    TrackingKeywordRows = @($trackingKeywordRows).Count
    TrackingKeywordMailboxRows = @($trackingKeywordMailboxRows).Count
    DailyCounts = @($dailyCounts).Count
  } -Errors @()
}
