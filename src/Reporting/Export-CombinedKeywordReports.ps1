Set-StrictMode -Version Latest

function Export-ImtCombinedKeywordReports {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory=$true)]$RunContext,
    [Parameter(Mandatory=$true)][AllowEmptyCollection()][string[]]$BaseTargetAddresses,
    [object[]]$TrackingKeywordRows,
    [object[]]$TrackingKeywordMailboxRows,
    [object[]]$DirectKeywordRows
  )

  if (-not $RunContext.Inputs.Keywords -or $RunContext.Inputs.Keywords.Count -eq 0) {
    return New-ImtModuleResult -StepName 'KeywordCombined' -Status 'SKIP' -Summary 'Combined keyword reports skipped (no keywords provided).' -Data ([pscustomobject]@{
      CombinedRows = @()
      CombinedByMailboxRows = @()
      CombinedCsv = $null
      CombinedByMailboxCsv = $null
    }) -Metrics @{} -Errors @()
  }

  $normalizedKeywords = @(
    $RunContext.Inputs.Keywords |
      Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
      ForEach-Object { $_.Trim() } |
      Select-Object -Unique
  )

  $combinedRows = @()
  foreach ($keywordName in $normalizedKeywords) {
    $trackingRows = @($TrackingKeywordRows | Where-Object { $_.Keyword -eq $keywordName })
    $transportEventCount = if ($trackingRows.Count -gt 0) { [int]($trackingRows[0].EventHitCount -as [int]) } else { 0 }
    $transportMessageCount = if ($trackingRows.Count -gt 0) { [int]($trackingRows[0].DistinctMessageIdHitCount -as [int]) } else { 0 }

    $directRows = @($DirectKeywordRows | Where-Object { $_.Keyword -eq $keywordName -and $_.Status -eq 'OK' })
    $mailboxHits = 0
    $mailboxesWithHits = @{}

    foreach ($directRow in $directRows) {
      $hitCount = ($directRow.HitCount -as [int])
      if ($null -eq $hitCount) { $hitCount = 0 }
      $mailboxHits += [int]$hitCount

      if ($hitCount -gt 0) {
        $mailboxName = ($directRow.Mailbox -as [string])
        if (-not [string]::IsNullOrWhiteSpace($mailboxName)) {
          $mailboxesWithHits[$mailboxName.ToLowerInvariant()] = $true
        }
      }
    }

    $combinedRows += [pscustomobject]@{
      Keyword = $keywordName
      TransportEventHitCount = $transportEventCount
      TransportDistinctMessageIdHitCount = $transportMessageCount
      MailboxEstimatedItemHitCount = $mailboxHits
      MailboxesWithMailboxHits = @($mailboxesWithHits.Keys).Count
    }
  }

  $transportByMailboxKeyword = @{}
  foreach ($trackingMailboxRow in $TrackingKeywordMailboxRows) {
    $mailbox = ($trackingMailboxRow.Mailbox -as [string])
    $keyword = ($trackingMailboxRow.Keyword -as [string])
    if ([string]::IsNullOrWhiteSpace($mailbox) -or [string]::IsNullOrWhiteSpace($keyword)) {
      continue
    }

    $key = "{0}|{1}" -f $mailbox.Trim().ToLowerInvariant(), $keyword.Trim().ToLowerInvariant()
    $eventCount = ($trackingMailboxRow.EventHitCount -as [int])
    $messageCount = ($trackingMailboxRow.DistinctMessageIdHitCount -as [int])
    if ($null -eq $eventCount) { $eventCount = 0 }
    if ($null -eq $messageCount) { $messageCount = 0 }

    $transportByMailboxKeyword[$key] = [pscustomobject]@{
      EventHitCount = [int]$eventCount
      DistinctMessageIdHitCount = [int]$messageCount
    }
  }

  $mailboxByKeyword = @{}
  foreach ($directRow in ($DirectKeywordRows | Where-Object { $_.Status -eq 'OK' })) {
    $mailbox = ($directRow.Mailbox -as [string])
    $keyword = ($directRow.Keyword -as [string])
    if ([string]::IsNullOrWhiteSpace($mailbox) -or [string]::IsNullOrWhiteSpace($keyword)) {
      continue
    }

    $key = "{0}|{1}" -f $mailbox.Trim().ToLowerInvariant(), $keyword.Trim().ToLowerInvariant()
    if (-not $mailboxByKeyword.ContainsKey($key)) {
      $mailboxByKeyword[$key] = 0
    }

    $hitCount = ($directRow.HitCount -as [int])
    if ($null -eq $hitCount) { $hitCount = 0 }
    $mailboxByKeyword[$key] = ([int]$mailboxByKeyword[$key]) + ([int]$hitCount)
  }

  $mailboxSet = @{}
  foreach ($address in $BaseTargetAddresses) {
    $value = ($address -as [string])
    if (-not [string]::IsNullOrWhiteSpace($value)) {
      $mailboxSet[$value.Trim().ToLowerInvariant()] = $true
    }
  }

  foreach ($trackingMailboxRow in $TrackingKeywordMailboxRows) {
    $value = ($trackingMailboxRow.Mailbox -as [string])
    if (-not [string]::IsNullOrWhiteSpace($value)) {
      $mailboxSet[$value.Trim().ToLowerInvariant()] = $true
    }
  }

  foreach ($directRow in ($DirectKeywordRows | Where-Object { $_.Status -eq 'OK' })) {
    $value = ($directRow.Mailbox -as [string])
    if (-not [string]::IsNullOrWhiteSpace($value)) {
      $mailboxSet[$value.Trim().ToLowerInvariant()] = $true
    }
  }

  $combinedByMailboxRows = @()
  foreach ($mailbox in (@($mailboxSet.Keys) | Sort-Object)) {
    foreach ($keywordName in $normalizedKeywords) {
      $key = "{0}|{1}" -f $mailbox, $keywordName.ToLowerInvariant()

      $transportEventCount = 0
      $transportMessageCount = 0
      if ($transportByMailboxKeyword.ContainsKey($key)) {
        $transportEventCount = [int]$transportByMailboxKeyword[$key].EventHitCount
        $transportMessageCount = [int]$transportByMailboxKeyword[$key].DistinctMessageIdHitCount
      }

      $mailboxHitCount = 0
      if ($mailboxByKeyword.ContainsKey($key)) {
        $mailboxHitCount = [int]$mailboxByKeyword[$key]
      }

      $combinedByMailboxRows += [pscustomobject]@{
        Mailbox = $mailbox
        Keyword = $keywordName
        TransportEventHitCount = $transportEventCount
        TransportDistinctMessageIdHitCount = $transportMessageCount
        MailboxEstimatedItemHitCount = $mailboxHitCount
      }
    }
  }

  $combinedCsv = $null
  if ($combinedRows.Count -gt 0) {
    $combinedCsv = Join-Path $RunContext.OutputDir ("MTL_KeywordHits_Combined_{0}.csv" -f $RunContext.Timestamp)
    $combinedRows | Sort-Object Keyword | Export-Csv -Path $combinedCsv -NoTypeInformation -Encoding UTF8
  }

  $combinedByMailboxCsv = $null
  if ($combinedByMailboxRows.Count -gt 0) {
    $combinedByMailboxCsv = Join-Path $RunContext.OutputDir ("MTL_KeywordHits_Combined_ByMailbox_{0}.csv" -f $RunContext.Timestamp)
    $combinedByMailboxRows | Sort-Object Mailbox,Keyword | Export-Csv -Path $combinedByMailboxCsv -NoTypeInformation -Encoding UTF8
  }

  New-ImtModuleResult -StepName 'KeywordCombined' -Status 'OK' -Summary ("Combined keyword reports exported. OverallRows={0}; ByMailboxRows={1}" -f $combinedRows.Count, $combinedByMailboxRows.Count) -Data ([pscustomobject]@{
    CombinedRows = @($combinedRows)
    CombinedByMailboxRows = @($combinedByMailboxRows)
    CombinedCsv = $combinedCsv
    CombinedByMailboxCsv = $combinedByMailboxCsv
  }) -Metrics @{
    CombinedRows = $combinedRows.Count
    CombinedByMailboxRows = $combinedByMailboxRows.Count
  } -Errors @()
}
