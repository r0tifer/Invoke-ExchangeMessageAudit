Set-StrictMode -Version Latest

function Get-ImtLatestStepData {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory=$true)][object[]]$StepResults,
    [Parameter(Mandatory=$true)][string]$StepName
  )

  $stepResult = @(
    $StepResults |
      Where-Object {
        ($_.StepName -as [string]) -eq $StepName -and
        $_.Data
      }
  ) | Select-Object -Last 1

  if ($stepResult) {
    return $stepResult.Data
  }

  $null
}

function Write-ImtRunSummary {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory=$true)]$RunContext,
    [Parameter(Mandatory=$true)][object[]]$StepResults,
    [Parameter(Mandatory=$true)][datetime]$StartedAt,
    [Parameter(Mandatory=$true)][datetime]$EndedAt
  )

  $summaryObject = New-ImtRunSummary -StepResults $StepResults -StartedAt $StartedAt -EndedAt $EndedAt

  $stepOutcomes = @(
    foreach ($step in @($StepResults)) {
      [pscustomobject]@{
        Step = $step.StepName
        Status = $step.Status
        Summary = $step.Summary
      }
    }
  )

  $directSearchData = Get-ImtLatestStepData -StepResults $StepResults -StepName 'DirectMailboxSearch'
  $mailboxEvidenceData = Get-ImtLatestStepData -StepResults $StepResults -StepName 'MailboxEvidence'
  $combinedData = Get-ImtLatestStepData -StepResults $StepResults -StepName 'KeywordCombined'

  $finalKeywordByMailboxRows = @()
  $finalMailboxFindingRows = @()
  $finalEvidenceDetailRows = @()
  $finalRealHitRows = @()
  $finalNearMissRows = @()

  if ($combinedData -and $combinedData.PSObject.Properties.Name -contains 'CombinedByMailboxRows' -and $combinedData.CombinedByMailboxRows) {
    $finalKeywordByMailboxRows = @(
      foreach ($row in @($combinedData.CombinedByMailboxRows | Sort-Object Mailbox, Keyword)) {
        [pscustomobject]@{
          Mailbox = $row.Mailbox
          Keyword = $row.Keyword
          TransportEventHitCount = [int]($row.TransportEventHitCount -as [int])
          TransportDistinctMessageIdHitCount = [int]($row.TransportDistinctMessageIdHitCount -as [int])
          MailboxEstimatedItemHitCount = [int]($row.MailboxEstimatedItemHitCount -as [int])
        }
      }
    )
  }

  $evidenceSummaryByMailbox = @{}
  if ($mailboxEvidenceData -and $mailboxEvidenceData.PSObject.Properties.Name -contains 'SummaryRows') {
    foreach ($row in @($mailboxEvidenceData.SummaryRows)) {
      $mailbox = ($row.Mailbox -as [string])
      if ([string]::IsNullOrWhiteSpace($mailbox)) {
        continue
      }

      $evidenceSummaryByMailbox[$mailbox.Trim().ToLowerInvariant()] = [pscustomobject]@{
        Mailbox = $mailbox.Trim()
        EvidenceRowCount = [int]($row.EvidenceRowCount -as [int])
        TransportCorrelatedCount = [int]($row.TransportCorrelatedCount -as [int])
      }
    }
  }

  if ($mailboxEvidenceData -and $mailboxEvidenceData.PSObject.Properties.Name -contains 'EvidenceRows' -and $mailboxEvidenceData.EvidenceRows) {
    $finalEvidenceDetailRows = @(
      foreach ($row in @($mailboxEvidenceData.EvidenceRows | Sort-Object SourceMailbox, SentTime, Subject)) {
        [pscustomobject]@{
          Mailbox = $row.SourceMailbox
          MailboxLocation = $row.MailboxLocation
          SentTime = $row.SentTime
          Subject = $row.Subject
          To = $row.To
          Cc = $row.Cc
          AttachmentCount = if ($null -ne $row.AttachmentCount) { [int]($row.AttachmentCount -as [int]) } else { 0 }
          TransportCorrelated = [bool]$row.TransportCorrelated
          TrackingMessageId = $row.TrackingMessageId
          EvidenceFolder = $row.EvidenceFolder
        }
      }
    )
  }

  $mailboxFindingByMailbox = @{}
  if ($directSearchData -and $directSearchData.PSObject.Properties.Name -contains 'DirectRows' -and $directSearchData.DirectRows) {
    foreach ($row in @($directSearchData.DirectRows | Sort-Object Mailbox)) {
      $mailbox = ($row.Mailbox -as [string])
      if ([string]::IsNullOrWhiteSpace($mailbox)) {
        continue
      }

      $mailboxKey = $mailbox.Trim().ToLowerInvariant()
      $evidenceSummary = $null
      if ($evidenceSummaryByMailbox.ContainsKey($mailboxKey)) {
        $evidenceSummary = $evidenceSummaryByMailbox[$mailboxKey]
      }

      $mailboxFindingByMailbox[$mailboxKey] = [pscustomobject]@{
        Mailbox = $mailbox.Trim()
        ResultItemsCount = if ($null -ne $row.ResultItemsCount) { [int]($row.ResultItemsCount -as [int]) } else { $null }
        DateRangeItemsCount = if ($null -ne $row.DateRangeItemsCount) { [int]($row.DateRangeItemsCount -as [int]) } else { $null }
        PerKeywordHitTotal = if ($null -ne $row.PerKeywordHitTotal) { [int]($row.PerKeywordHitTotal -as [int]) } else { 0 }
        MatchedKeywordsInRange = $row.MatchedKeywordsInRange
        ResultItemsSize = $row.ResultItemsSize
        EvidenceRowCount = if ($evidenceSummary) { [int]$evidenceSummary.EvidenceRowCount } else { 0 }
        TransportCorrelatedCount = if ($evidenceSummary) { [int]$evidenceSummary.TransportCorrelatedCount } else { 0 }
        Status = $row.Status
        Error = $row.Error
      }
    }
  }

  foreach ($entry in $evidenceSummaryByMailbox.GetEnumerator()) {
    if ($mailboxFindingByMailbox.ContainsKey($entry.Key)) {
      continue
    }

    $row = $entry.Value
    $mailboxFindingByMailbox[$entry.Key] = [pscustomobject]@{
      Mailbox = $row.Mailbox
      ResultItemsCount = $null
      DateRangeItemsCount = $null
      PerKeywordHitTotal = 0
      MatchedKeywordsInRange = $null
      ResultItemsSize = $null
      EvidenceRowCount = [int]$row.EvidenceRowCount
      TransportCorrelatedCount = [int]$row.TransportCorrelatedCount
      Status = 'EvidenceOnly'
      Error = $null
    }
  }

  $finalMailboxFindingRows = @(
    $mailboxFindingByMailbox.GetEnumerator() |
      Sort-Object Name |
      ForEach-Object { $_.Value }
  )

  $finalRealHitRows = @(
    $finalMailboxFindingRows |
      Where-Object {
        $resultCount = ($_.ResultItemsCount -as [int])
        $null -ne $resultCount -and $resultCount -gt 0
      } |
      Sort-Object @{ Expression = { [int]($_.ResultItemsCount -as [int]) }; Descending = $true }, Mailbox |
      ForEach-Object {
        [pscustomobject]@{
          Mailbox = $_.Mailbox
          ResultItemsCount = [int]($_.ResultItemsCount -as [int])
          DateRangeItemsCount = if ($null -ne $_.DateRangeItemsCount) { [int]($_.DateRangeItemsCount -as [int]) } else { $null }
          EvidenceRowCount = [int]($_.EvidenceRowCount -as [int])
          TransportCorrelatedCount = [int]($_.TransportCorrelatedCount -as [int])
          MatchedKeywordsInRange = $_.MatchedKeywordsInRange
        }
      }
  )

  $finalNearMissRows = @(
    $finalMailboxFindingRows |
      Where-Object {
        $dateRangeCount = ($_.DateRangeItemsCount -as [int])
        $resultCount = ($_.ResultItemsCount -as [int])
        $null -ne $dateRangeCount -and
        $dateRangeCount -gt 0 -and
        (($null -eq $_.ResultItemsCount) -or $resultCount -eq 0)
      } |
      Sort-Object @{ Expression = { [int]($_.DateRangeItemsCount -as [int]) }; Descending = $true }, Mailbox |
      ForEach-Object {
        [pscustomobject]@{
          Mailbox = $_.Mailbox
          DateRangeItemsCount = [int]($_.DateRangeItemsCount -as [int])
          PerKeywordHitTotal = [int]($_.PerKeywordHitTotal -as [int])
          MatchedKeywordsInRange = $_.MatchedKeywordsInRange
          EvidenceRowCount = [int]($_.EvidenceRowCount -as [int])
        }
      }
  )

  $summaryObject | Add-Member -NotePropertyName StepOutcomes -NotePropertyValue @($stepOutcomes) -Force
  $summaryObject | Add-Member -NotePropertyName FinalMailboxFindingRows -NotePropertyValue @($finalMailboxFindingRows) -Force
  $summaryObject | Add-Member -NotePropertyName FinalEvidenceDetailRows -NotePropertyValue @($finalEvidenceDetailRows) -Force
  $summaryObject | Add-Member -NotePropertyName FinalRealHitRows -NotePropertyValue @($finalRealHitRows) -Force
  $summaryObject | Add-Member -NotePropertyName FinalNearMissRows -NotePropertyValue @($finalNearMissRows) -Force
  $summaryObject | Add-Member -NotePropertyName FinalKeywordByMailboxRows -NotePropertyValue @($finalKeywordByMailboxRows) -Force

  $line = "TotalSteps={0}; OK={1}; WARN={2}; FAIL={3}; SKIP={4}; DurationSeconds={5}" -f `
    $summaryObject.TotalSteps, `
    $summaryObject.Counts.OK, `
    $summaryObject.Counts.WARN, `
    $summaryObject.Counts.FAIL, `
    $summaryObject.Counts.SKIP, `
    ($summaryObject.DurationSeconds -as [string])

  $level = if ($summaryObject.Counts.FAIL -gt 0) { 'ERROR' } else { 'INFO' }
  Write-ImtLog -Level $level -Step 'RunSummary' -EventType Summary -Message $line

  New-ImtModuleResult -StepName 'RunSummary' -Status 'OK' -Summary $line -Data $summaryObject -Metrics @{
    TotalSteps = $summaryObject.TotalSteps
    OK = $summaryObject.Counts.OK
    WARN = $summaryObject.Counts.WARN
    FAIL = $summaryObject.Counts.FAIL
    SKIP = $summaryObject.Counts.SKIP
  } -Errors @()
}
