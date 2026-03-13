Set-StrictMode -Version Latest

function ConvertTo-ImtAddressSet {
  [CmdletBinding()]
  param(
    [string[]]$Values
  )

  $set = @{}
  foreach ($value in @($Values)) {
    $raw = ($value -as [string])
    if ([string]::IsNullOrWhiteSpace($raw)) {
      continue
    }

    foreach ($part in ($raw -split ';')) {
      $clean = ($part -as [string])
      if ([string]::IsNullOrWhiteSpace($clean)) {
        continue
      }

      $set[$clean.Trim().ToLowerInvariant()] = $true
    }
  }

  $set
}

function Find-ImtTrackingCorrelation {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory=$true)]$EvidenceRow,
    [object[]]$TrackingResults,
    [hashtable]$TrackingByMessageId
  )

  $messageId = ($EvidenceRow.InternetMessageId -as [string])
  if (-not [string]::IsNullOrWhiteSpace($messageId) -and $TrackingByMessageId -and $TrackingByMessageId.ContainsKey($messageId.Trim().ToLowerInvariant())) {
    $match = @($TrackingByMessageId[$messageId.Trim().ToLowerInvariant()] | Select-Object -First 1)
    if ($match.Count -gt 0) {
      return [pscustomobject]@{
        TransportCorrelated = $true
        TrackingMessageId = $match[0].MessageId
      }
    }
  }

  $sourceMailbox = ($EvidenceRow.SourceMailbox -as [string])
  $sentTime = $EvidenceRow.SentTime
  $targetRecipients = ConvertTo-ImtAddressSet -Values @($EvidenceRow.To, $EvidenceRow.Cc)

  if ([string]::IsNullOrWhiteSpace($sourceMailbox) -or $targetRecipients.Count -eq 0 -or -not $sentTime) {
    return [pscustomobject]@{
      TransportCorrelated = $false
      TrackingMessageId = $null
    }
  }

  $senderLower = $sourceMailbox.Trim().ToLowerInvariant()
  $match = @(
    $TrackingResults |
      Where-Object {
        $trackingSender = ($_.Sender -as [string])
        if ([string]::IsNullOrWhiteSpace($trackingSender) -or $trackingSender.Trim().ToLowerInvariant() -ne $senderLower) {
          return $false
        }

        $recipientOverlap = $false
        foreach ($recipient in @($_.Recipients)) {
          $recipientValue = ($recipient -as [string])
          if (-not [string]::IsNullOrWhiteSpace($recipientValue) -and $targetRecipients.ContainsKey($recipientValue.Trim().ToLowerInvariant())) {
            $recipientOverlap = $true
            break
          }
        }

        if (-not $recipientOverlap) {
          return $false
        }

        $timestamp = $_.Timestamp
        if (-not $timestamp) {
          return $false
        }

        [math]::Abs((New-TimeSpan -Start $timestamp -End $sentTime).TotalMinutes) -le 120
      } |
      Sort-Object Timestamp
  )

  if ($match.Count -gt 0) {
    return [pscustomobject]@{
      TransportCorrelated = $true
      TrackingMessageId = $match[0].MessageId
    }
  }

  [pscustomobject]@{
    TransportCorrelated = $false
    TrackingMessageId = $null
  }
}

function Export-ImtMailboxEvidenceReports {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory=$true)]$RunContext,
    [object[]]$EvidenceRows,
    [object[]]$TrackingResults
  )

  $rows = @($EvidenceRows)
  if ($rows.Count -eq 0) {
    $summary = if ($RunContext.Inputs.DetailedMailboxEvidence) {
      'Mailbox evidence requested but no matching evidence rows were collected.'
    } else {
      'Mailbox evidence reporting skipped (no evidence rows available).'
    }

    return New-ImtModuleResult -StepName 'MailboxEvidence' -Status 'SKIP' -Summary $summary -Data ([pscustomobject]@{
      EvidenceRows = @()
      EvidenceCsv = $null
      SummaryRows = @()
      TrackingAvailable = $false
    }) -Metrics @{
      EvidenceRows = 0
      SummaryRows = 0
    } -Errors @()
  }

  $trackingByMessageId = @{}
  foreach ($result in @($TrackingResults)) {
    $trackingMessageId = ($result.MessageId -as [string])
    if ([string]::IsNullOrWhiteSpace($trackingMessageId)) {
      continue
    }

    $key = $trackingMessageId.Trim().ToLowerInvariant()
    if (-not $trackingByMessageId.ContainsKey($key)) {
      $trackingByMessageId[$key] = New-Object System.Collections.Generic.List[object]
    }

    [void]$trackingByMessageId[$key].Add($result)
  }

  $trackingAvailable = @($TrackingResults).Count -gt 0
  $exportRows = New-Object System.Collections.Generic.List[object]

  foreach ($row in $rows) {
    $correlation = Find-ImtTrackingCorrelation -EvidenceRow $row -TrackingResults @($TrackingResults) -TrackingByMessageId $trackingByMessageId
    [void]$exportRows.Add([pscustomobject]@{
      SourceMailbox = $row.SourceMailbox
      MailboxLocation = $row.MailboxLocation
      SentTime = $row.SentTime
      From = $row.From
      To = $row.To
      Cc = $row.Cc
      Subject = $row.Subject
      InternetMessageId = $row.InternetMessageId
      HasAttachments = [bool]$row.HasAttachments
      AttachmentCount = if ($null -ne $row.AttachmentCount) { [int]$row.AttachmentCount } else { 0 }
      ItemSize = $row.ItemSize
      EvidenceFolder = $row.EvidenceFolder
      TransportCorrelated = [bool]$correlation.TransportCorrelated
      TrackingMessageId = $correlation.TrackingMessageId
    })
  }

  $evidenceCsv = Join-Path $RunContext.OutputDir ("MTL_MailboxEvidence_{0}.csv" -f $RunContext.Timestamp)
  $exportRows | Sort-Object SourceMailbox,SentTime,Subject | Export-Csv -Path $evidenceCsv -NoTypeInformation -Encoding UTF8

  $summaryRows = @(
    $exportRows |
      Group-Object SourceMailbox |
      Sort-Object Name |
      ForEach-Object {
        [pscustomobject]@{
          Mailbox = $_.Name
          EvidenceRowCount = $_.Count
          TransportCorrelatedCount = @($_.Group | Where-Object { $_.TransportCorrelated }).Count
        }
      }
  )

  $status = 'OK'
  $errors = @()
  if (-not $trackingAvailable) {
    $status = 'WARN'
    $errors += 'Transport correlation skipped because no message tracking results were available for the selected window.'
  }

  New-ImtModuleResult -StepName 'MailboxEvidence' -Status $status -Summary ("Mailbox evidence rows={0}; csv={1}" -f $exportRows.Count, $evidenceCsv) -Data ([pscustomobject]@{
    EvidenceRows = $exportRows.ToArray()
    EvidenceCsv = $evidenceCsv
    SummaryRows = @($summaryRows)
    TrackingAvailable = $trackingAvailable
  }) -Metrics @{
    EvidenceRows = $exportRows.Count
    SummaryRows = $summaryRows.Count
    TrackingAvailable = $trackingAvailable
  } -Errors @($errors)
}
