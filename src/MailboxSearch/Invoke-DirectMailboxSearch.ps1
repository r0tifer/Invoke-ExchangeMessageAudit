Set-StrictMode -Version Latest

function New-ImtMailboxSearchQuery {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory=$true)][datetime]$StartAt,
    [Parameter(Mandatory=$true)][datetime]$EndAt,
    [string]$SubjectContains,
    [string[]]$SubjectKeywords,
    [string[]]$SenderFilters,
    [switch]$RequireAttachment
  )

  $startDateOnly = $StartAt.ToString('MM/dd/yyyy')
  $endDateOnly = $EndAt.ToString('MM/dd/yyyy')
  $dateClause = "((received:$startDateOnly..$endDateOnly) OR (sent:$startDateOnly..$endDateOnly))"

  $subjectParts = New-Object System.Collections.Generic.List[string]
  if (-not [string]::IsNullOrWhiteSpace($SubjectContains)) {
    [void]$subjectParts.Add("(`"$SubjectContains`")")
  }

  if ($SubjectKeywords -and $SubjectKeywords.Count -gt 0) {
    foreach ($keyword in ($SubjectKeywords | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })) {
      $clean = $keyword.Trim()
      [void]$subjectParts.Add("(`"$clean`")")
      if ($clean -like '*/*') {
        $spaced = $clean -replace '/', ' '
        if ($spaced -ne $clean) { [void]$subjectParts.Add("(`"$spaced`")") }

        $noSlash = $clean -replace '/', ''
        if ($noSlash -ne $clean) { [void]$subjectParts.Add("(`"$noSlash`")") }
      }
    }
  }

  $query = if ($subjectParts.Count -gt 0) {
    "($dateClause AND ($($subjectParts -join ' OR ')))"
  } else {
    $dateClause
  }

  $senderParts = New-Object System.Collections.Generic.List[string]
  if ($SenderFilters -and $SenderFilters.Count -gt 0) {
    foreach ($sender in ($SenderFilters | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })) {
      [void]$senderParts.Add("(from:`"$($sender.Trim())`")")
    }
  }

  if ($senderParts.Count -gt 0) {
    $query = "($query AND ($($senderParts -join ' OR ')))"
  }

  if ($RequireAttachment) {
    $query = "($query AND (hasattachment:true))"
  }

  $query
}

function Get-ImtSearchMailboxEstimate {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory=$true)]$MailboxIdentity,
    [Parameter(Mandatory=$true)][string]$SearchQuery,
    [switch]$SearchDumpster
  )

  try {
    if ($SearchDumpster) {
      $result = Search-Mailbox -Identity $MailboxIdentity -SearchQuery $SearchQuery -EstimateResultOnly -SearchDumpster -ErrorAction Stop -WarningAction SilentlyContinue
    } else {
      $result = Search-Mailbox -Identity $MailboxIdentity -SearchQuery $SearchQuery -EstimateResultOnly -ErrorAction Stop -WarningAction SilentlyContinue
    }

    $count = 0
    $size = 'n/a'
    if ($result -and $null -ne $result.ResultItemsCount) {
      $count = [int]$result.ResultItemsCount
    }
    if ($result -and $null -ne $result.ResultItemsSize) {
      $size = $result.ResultItemsSize.ToString()
    }

    [pscustomobject]@{
      Success = $true
      Count = $count
      Size = $size
      Error = $null
    }
  } catch {
    [pscustomobject]@{
      Success = $false
      Count = $null
      Size = $null
      Error = $_.Exception.Message
    }
  }
}

function Invoke-ImtDirectMailboxSearch {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory=$true)]$RunContext,
    [Parameter(Mandatory=$true)][string[]]$BaseTargetAddresses,
    [string[]]$EffectiveSenderFilters
  )

  if (-not (Get-Command -Name Search-Mailbox -ErrorAction SilentlyContinue)) {
    return New-ImtModuleResult -StepName 'DirectMailboxSearch' -Status 'FAIL' -Summary 'Search-Mailbox cmdlet is unavailable in this session.' -Data ([pscustomobject]@{
      DirectRows = @()
      DirectKeywordRows = @()
      NonSearchable = @()
      DirectCsv = $null
      DirectKeywordCsv = $null
      NonSearchableCsv = $null
    }) -Metrics @{} -Errors @('Search-Mailbox cmdlet unavailable')
  }

  $resolved = Resolve-ImtMailboxesByAddressSet -Addresses $BaseTargetAddresses -CaptureUnresolved
  $targetMailboxes = @($resolved.Mailboxes)
  $nonSearchable = @($resolved.Unresolved)

  if ($targetMailboxes.Count -eq 0) {
    return New-ImtModuleResult -StepName 'DirectMailboxSearch' -Status 'WARN' -Summary 'No target mailboxes could be resolved for direct search.' -Data ([pscustomobject]@{
      DirectRows = @()
      DirectKeywordRows = @()
      NonSearchable = $nonSearchable
      DirectCsv = $null
      DirectKeywordCsv = $null
      NonSearchableCsv = $null
    }) -Metrics @{
      TargetMailboxes = 0
      NonSearchable = $nonSearchable.Count
    } -Errors @()
  }

  $searchQuery = New-ImtMailboxSearchQuery -StartAt $RunContext.Start -EndAt $RunContext.End -SubjectContains $RunContext.Inputs.SubjectLike -SubjectKeywords $RunContext.Inputs.Keywords -SenderFilters $EffectiveSenderFilters -RequireAttachment:$RunContext.Inputs.HasAttachmentOnly
  $dateOnlyQuery = New-ImtMailboxSearchQuery -StartAt $RunContext.Start -EndAt $RunContext.End -SubjectContains $null -SubjectKeywords @() -SenderFilters $EffectiveSenderFilters -RequireAttachment:$RunContext.Inputs.HasAttachmentOnly

  Write-ImtLog -Level DEBUG -Step 'DirectMailboxSearch' -EventType Progress -Message ("Mailboxes={0}; Query={1}" -f $targetMailboxes.Count, $searchQuery)

  $directRows = New-Object System.Collections.Generic.List[object]
  $directKeywordRows = New-Object System.Collections.Generic.List[object]

  foreach ($mailbox in $targetMailboxes) {
    $mailboxIdentity = $mailbox.Identity
    $mailboxSmtp = $mailbox.PrimarySmtpAddress.ToString()

    try {
      $combined = Get-ImtSearchMailboxEstimate -MailboxIdentity $mailboxIdentity -SearchQuery $searchQuery -SearchDumpster:$RunContext.Inputs.SearchDumpsterDirectly
      if (-not $combined.Success) {
        throw $combined.Error
      }

      $dateOnly = Get-ImtSearchMailboxEstimate -MailboxIdentity $mailboxIdentity -SearchQuery $dateOnlyQuery -SearchDumpster:$RunContext.Inputs.SearchDumpsterDirectly
      $dateOnlyCount = if ($dateOnly.Success -and $null -ne $dateOnly.Count) { [int]$dateOnly.Count } else { $null }

      $matchedKeywords = @()
      $keywordSummaryParts = @()
      $keywordTotalHits = 0

      if ($RunContext.Inputs.Keywords -and $RunContext.Inputs.Keywords.Count -gt 0) {
        foreach ($keyword in ($RunContext.Inputs.Keywords | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })) {
          $keywordQuery = New-ImtMailboxSearchQuery -StartAt $RunContext.Start -EndAt $RunContext.End -SubjectContains $null -SubjectKeywords @($keyword) -SenderFilters $EffectiveSenderFilters -RequireAttachment:$RunContext.Inputs.HasAttachmentOnly
          $keywordResult = Get-ImtSearchMailboxEstimate -MailboxIdentity $mailboxIdentity -SearchQuery $keywordQuery -SearchDumpster:$RunContext.Inputs.SearchDumpsterDirectly

          if ($keywordResult.Success -and $null -ne $keywordResult.Count -and [int]$keywordResult.Count -gt 0) {
            $matchedKeywords += $keyword.Trim()
            $keywordTotalHits += [int]$keywordResult.Count
          }

          if ($keywordResult.Success -and $null -ne $keywordResult.Count) {
            $keywordSummaryParts += ("{0}={1}" -f $keyword.Trim(), [int]$keywordResult.Count)
            [void]$directKeywordRows.Add([pscustomobject]@{
              Mailbox = $mailboxSmtp
              Keyword = $keyword.Trim()
              HitCount = [int]$keywordResult.Count
              Status = 'OK'
              Error = $null
            })
          } else {
            $keywordSummaryParts += ("{0}=ERR" -f $keyword.Trim())
            [void]$directKeywordRows.Add([pscustomobject]@{
              Mailbox = $mailboxSmtp
              Keyword = $keyword.Trim()
              HitCount = $null
              Status = 'Failed'
              Error = $keywordResult.Error
            })
          }
        }
      }

      $items = [int]$combined.Count
      $size = $combined.Size

      [void]$directRows.Add([pscustomobject]@{
        Mailbox = $mailboxSmtp
        ResultItemsCount = $items
        ResultItemsSize = $size
        DateRangeItemsCount = $dateOnlyCount
        MatchedKeywordsInRange = ($matchedKeywords -join ';')
        PerKeywordHitSummary = ($keywordSummaryParts -join ';')
        PerKeywordHitTotal = $keywordTotalHits
        Status = 'OK'
        Error = $null
      })

      Write-ImtLog -Level DEBUG -Step 'DirectMailboxSearch' -EventType Progress -Message ("{0}: matched={1}; date-range total={2}" -f $mailboxSmtp, $items, ($dateOnlyCount -as [string]))
    } catch {
      [void]$directRows.Add([pscustomobject]@{
        Mailbox = $mailboxSmtp
        ResultItemsCount = $null
        ResultItemsSize = $null
        DateRangeItemsCount = $null
        MatchedKeywordsInRange = $null
        PerKeywordHitSummary = $null
        PerKeywordHitTotal = $null
        Status = 'Failed'
        Error = $_.Exception.Message
      })

      Write-ImtLog -Level WARN -Step 'DirectMailboxSearch' -EventType Progress -Message ("Direct mailbox search failed for {0}: {1}" -f $mailboxSmtp, $_.Exception.Message)
    }
  }

  $directCsv = Join-Path $RunContext.OutputDir ("MTL_DirectMailboxSearch_{0}.csv" -f $RunContext.Timestamp)
  $directRows | Export-Csv -Path $directCsv -NoTypeInformation -Encoding UTF8

  $directKeywordRowsAll = @($directKeywordRows | ForEach-Object { $_ })
  $directKeywordCsv = $null
  if ($directKeywordRowsAll.Count -gt 0) {
    $directKeywordCsv = Join-Path $RunContext.OutputDir ("MTL_DirectMailboxSearch_KeywordHits_{0}.csv" -f $RunContext.Timestamp)
    $directKeywordRowsAll | Export-Csv -Path $directKeywordCsv -NoTypeInformation -Encoding UTF8
  }

  $nonSearchableCsv = $null
  if ($nonSearchable.Count -gt 0) {
    $nonSearchableCsv = Join-Path $RunContext.OutputDir ("MTL_DirectMailboxSearch_Skipped_{0}.csv" -f $RunContext.Timestamp)
    $nonSearchable | Export-Csv -Path $nonSearchableCsv -NoTypeInformation -Encoding UTF8
  }

  $failedRows = @($directRows | Where-Object { $_.Status -eq 'Failed' }).Count
  $status = if ($failedRows -gt 0) { 'WARN' } else { 'OK' }

  New-ImtModuleResult -StepName 'DirectMailboxSearch' -Status $status -Summary ("Direct mailbox search rows={0}; failed={1}; csv={2}" -f $directRows.Count, $failedRows, $directCsv) -Data ([pscustomobject]@{
    DirectRows = $directRows.ToArray()
    DirectKeywordRows = @($directKeywordRowsAll)
    NonSearchable = @($nonSearchable)
    DirectCsv = $directCsv
    DirectKeywordCsv = $directKeywordCsv
    NonSearchableCsv = $nonSearchableCsv
  }) -Metrics @{
    TargetMailboxes = $targetMailboxes.Count
    DirectRows = $directRows.Count
    DirectKeywordRows = $directKeywordRowsAll.Count
    NonSearchable = $nonSearchable.Count
    FailedRows = $failedRows
  } -Errors @()
}
