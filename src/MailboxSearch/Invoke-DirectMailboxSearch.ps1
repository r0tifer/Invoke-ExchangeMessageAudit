Set-StrictMode -Version Latest

function New-ImtMailboxSearchQuery {
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

  $startDateOnly = $StartAt.ToString('MM/dd/yyyy')
  $endDateOnly = $EndAt.ToString('MM/dd/yyyy')
  $dateClause = if ($OutboundOnly) {
    "(sent:$startDateOnly..$endDateOnly)"
  } else {
    "((received:$startDateOnly..$endDateOnly) OR (sent:$startDateOnly..$endDateOnly))"
  }

  $queryParts = New-Object System.Collections.Generic.List[string]
  [void]$queryParts.Add($dateClause)

  $subjectParts = New-Object System.Collections.Generic.List[string]
  if (-not [string]::IsNullOrWhiteSpace($SubjectContains)) {
    [void]$subjectParts.Add("(`"$($SubjectContains.Trim())`")")
  }

  if ($SubjectKeywords -and $SubjectKeywords.Count -gt 0) {
    foreach ($keyword in ($SubjectKeywords | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })) {
      $clean = $keyword.Trim()
      [void]$subjectParts.Add("(`"$clean`")")
      if ($clean -like '*/*') {
        $spaced = $clean -replace '/', ' '
        if ($spaced -ne $clean) {
          [void]$subjectParts.Add("(`"$spaced`")")
        }

        $noSlash = $clean -replace '/', ''
        if ($noSlash -ne $clean) {
          [void]$subjectParts.Add("(`"$noSlash`")")
        }
      }
    }
  }

  if ($subjectParts.Count -gt 0) {
    [void]$queryParts.Add("($($subjectParts -join ' OR '))")
  }

  $senderParts = New-Object System.Collections.Generic.List[string]
  if ($SenderFilters -and $SenderFilters.Count -gt 0) {
    foreach ($sender in ($SenderFilters | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })) {
      [void]$senderParts.Add("(from:`"$($sender.Trim())`")")
    }
  }

  if ($senderParts.Count -gt 0) {
    [void]$queryParts.Add("($($senderParts -join ' OR '))")
  }

  $recipientParts = New-Object System.Collections.Generic.List[string]
  if ($RecipientFilters -and $RecipientFilters.Count -gt 0) {
    foreach ($recipient in ($RecipientFilters | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })) {
      [void]$recipientParts.Add("(to:`"$($recipient.Trim())`")")
    }
  }

  if ($recipientParts.Count -gt 0) {
    [void]$queryParts.Add("($($recipientParts -join ' OR '))")
  }

  if ($RequireAttachment) {
    [void]$queryParts.Add('(hasattachment:true)')
  }

  "($($queryParts -join ' AND '))"
}

function Get-ImtSearchMailboxEstimate {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory=$true)]$MailboxIdentity,
    [Parameter(Mandatory=$true)][string]$SearchQuery,
    [switch]$SearchDumpster,
    [switch]$IncludeArchive
  )

  try {
    $params = @{
      Identity = $MailboxIdentity
      SearchQuery = $SearchQuery
      EstimateResultOnly = $true
      ErrorAction = 'Stop'
      WarningAction = 'SilentlyContinue'
    }

    if ($SearchDumpster) {
      $params.SearchDumpster = $true
    }

    if (-not $IncludeArchive) {
      $params.DoNotIncludeArchive = $true
    }

    $result = Search-Mailbox @params

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

function Invoke-ImtSearchMailboxCopy {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory=$true)]$MailboxIdentity,
    [Parameter(Mandatory=$true)][string]$SearchQuery,
    [Parameter(Mandatory=$true)][string]$TargetMailbox,
    [Parameter(Mandatory=$true)][string]$TargetFolder,
    [switch]$SearchDumpster,
    [switch]$IncludeArchive
  )

  try {
    $params = @{
      Identity = $MailboxIdentity
      SearchQuery = $SearchQuery
      TargetMailbox = $TargetMailbox
      TargetFolder = $TargetFolder
      LogLevel = 'Full'
      ErrorAction = 'Stop'
      WarningAction = 'SilentlyContinue'
    }

    if ($SearchDumpster) {
      $params.SearchDumpster = $true
    }

    if (-not $IncludeArchive) {
      $params.DoNotIncludeArchive = $true
    }

    $result = Search-Mailbox @params

    [pscustomobject]@{
      Success = $true
      Count = if ($result -and $null -ne $result.ResultItemsCount) { [int]$result.ResultItemsCount } else { 0 }
      Size = if ($result -and $null -ne $result.ResultItemsSize) { $result.ResultItemsSize.ToString() } else { $null }
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

function Test-ImtIsOrgWideSearchableMailbox {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory=$true)]$Mailbox
  )

  $details = ''
  if ($Mailbox.PSObject.Properties.Name -contains 'RecipientTypeDetails' -and $Mailbox.RecipientTypeDetails) {
    $details = $Mailbox.RecipientTypeDetails.ToString()
  }

  $details -in @('UserMailbox', 'SharedMailbox')
}

function Resolve-ImtMailboxAuditScope {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory=$true)]$RunContext,
    [string[]]$BaseTargetAddresses
  )

  $scopeRows = New-Object System.Collections.Generic.List[object]
  $mailboxes = New-Object System.Collections.Generic.List[object]
  $unresolved = New-Object System.Collections.Generic.List[object]

  if ($RunContext.Inputs.SourceMailboxes -and $RunContext.Inputs.SourceMailboxes.Count -gt 0) {
    foreach ($identity in @($RunContext.Inputs.SourceMailboxes)) {
      try {
        $mailbox = Get-Mailbox -Identity $identity -ErrorAction Stop
        [void]$mailboxes.Add($mailbox)
        [void]$scopeRows.Add([pscustomobject]@{
          ScopeMode = 'Explicit'
          Mailbox = $mailbox.PrimarySmtpAddress.ToString()
          RecipientTypeDetails = if ($mailbox.RecipientTypeDetails) { $mailbox.RecipientTypeDetails.ToString() } else { $null }
          Included = $true
          Reason = 'Explicit source mailbox'
        })
      } catch {
        [void]$unresolved.Add([pscustomobject]@{
          Address = $identity
          Reason = $_.Exception.Message
          RecipientType = 'Unresolved'
        })
      }
    }
  } elseif ($RunContext.Inputs.SearchAllMailboxes) {
    foreach ($mailbox in @(Get-Mailbox -ResultSize Unlimited -ErrorAction Stop)) {
      $mailboxSmtp = $mailbox.PrimarySmtpAddress.ToString()
      if (Test-ImtIsOrgWideSearchableMailbox -Mailbox $mailbox) {
        [void]$mailboxes.Add($mailbox)
        [void]$scopeRows.Add([pscustomobject]@{
          ScopeMode = 'OrgWide'
          Mailbox = $mailboxSmtp
          RecipientTypeDetails = if ($mailbox.RecipientTypeDetails) { $mailbox.RecipientTypeDetails.ToString() } else { $null }
          Included = $true
          Reason = 'Org-wide searchable mailbox'
        })
      } else {
        [void]$scopeRows.Add([pscustomobject]@{
          ScopeMode = 'OrgWide'
          Mailbox = $mailboxSmtp
          RecipientTypeDetails = if ($mailbox.RecipientTypeDetails) { $mailbox.RecipientTypeDetails.ToString() } else { $null }
          Included = $false
          Reason = 'Skipped non-user/shared mailbox'
        })
      }
    }
  } else {
    $resolved = Resolve-ImtMailboxesByAddressSet -Addresses $BaseTargetAddresses -CaptureUnresolved
    foreach ($mailbox in @($resolved.Mailboxes)) {
      [void]$mailboxes.Add($mailbox)
      [void]$scopeRows.Add([pscustomobject]@{
        ScopeMode = 'TargetAddress'
        Mailbox = $mailbox.PrimarySmtpAddress.ToString()
        RecipientTypeDetails = if ($mailbox.RecipientTypeDetails) { $mailbox.RecipientTypeDetails.ToString() } else { $null }
        Included = $true
        Reason = 'Resolved from existing target address set'
      })
    }

    foreach ($row in @($resolved.Unresolved)) {
      [void]$unresolved.Add($row)
    }
  }

  [pscustomobject]@{
    Mailboxes = @($mailboxes.ToArray() | Sort-Object DistinguishedName -Unique)
    ScopeRows = $scopeRows.ToArray()
    Unresolved = $unresolved.ToArray()
  }
}

function Get-ImtEwsAssemblyCandidatePaths {
  [CmdletBinding()]
  param()

  @(
    'C:\Program Files\Microsoft\Exchange\Web Services\2.2\Microsoft.Exchange.WebServices.dll'
    'C:\Program Files\Microsoft\Exchange\Web Services\2.1\Microsoft.Exchange.WebServices.dll'
    'C:\Program Files\Microsoft\Exchange\Web Services\2.0\Microsoft.Exchange.WebServices.dll'
  )
}

function Import-ImtEwsManagedApi {
  [CmdletBinding()]
  param()

  if ('Microsoft.Exchange.WebServices.Data.ExchangeService' -as [type]) {
    return
  }

  $loadedAssembly = [AppDomain]::CurrentDomain.GetAssemblies() | Where-Object { $_.GetName().Name -eq 'Microsoft.Exchange.WebServices' } | Select-Object -First 1
  if ($loadedAssembly) {
    return
  }

  foreach ($path in @(Get-ImtEwsAssemblyCandidatePaths)) {
    if (-not (Test-Path -LiteralPath $path)) {
      continue
    }

    Add-Type -Path $path
    if ('Microsoft.Exchange.WebServices.Data.ExchangeService' -as [type]) {
      return
    }
  }

  throw 'Microsoft.Exchange.WebServices.dll is required for -DetailedMailboxEvidence but could not be loaded.'
}

function Get-ImtEwsServiceForMailbox {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory=$true)][string]$MailboxSmtp
  )

  Import-ImtEwsManagedApi

  $preferredVersion = [enum]::GetValues([Microsoft.Exchange.WebServices.Data.ExchangeVersion]) | Where-Object { $_.ToString() -eq 'Exchange2013_SP1' } | Select-Object -First 1
  if (-not $preferredVersion) {
    $preferredVersion = [enum]::GetValues([Microsoft.Exchange.WebServices.Data.ExchangeVersion]) | Select-Object -First 1
  }

  $service = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService($preferredVersion)
  $service.UseDefaultCredentials = $true
  $service.AutodiscoverUrl($MailboxSmtp, { param($url) $url.Scheme -eq 'https' })
  $service
}

function Join-ImtEwsAddressCollection {
  [CmdletBinding()]
  param(
    [object[]]$Recipients
  )

  if (-not $Recipients) {
    return $null
  }

  @(
    $Recipients |
      ForEach-Object {
        if ($_.Address) {
          $_.Address
        } elseif ($_.Name) {
          $_.Name
        }
      } |
      Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
  ) -join ';'
}

function Get-ImtEwsFolderByPath {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory=$true)]$Service,
    [Parameter(Mandatory=$true)][string]$MailboxSmtp,
    [Parameter(Mandatory=$true)][string]$FolderPath
  )

  $segments = @(
    $FolderPath -split '[\\/]+' |
      Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
  )

  if ($segments.Count -eq 0) {
    return $null
  }

  $folderId = New-Object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::MsgFolderRoot, $MailboxSmtp)
  $currentFolder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($Service, $folderId)

  foreach ($segment in $segments) {
    $folderView = New-Object Microsoft.Exchange.WebServices.Data.FolderView(100)
    $folderView.Traversal = [Microsoft.Exchange.WebServices.Data.FolderTraversal]::Shallow
    $filter = New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+IsEqualTo([Microsoft.Exchange.WebServices.Data.FolderSchema]::DisplayName, $segment)
    $result = $Service.FindFolders($currentFolder.Id, $filter, $folderView)

    if ($result.TotalCount -lt 1) {
      return $null
    }

    $currentFolder = $result.Folders[0]
  }

  $currentFolder
}

function Get-ImtEvidenceItemsFromTargetMailbox {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory=$true)][string]$EvidenceMailbox,
    [Parameter(Mandatory=$true)][string]$FolderPath,
    [Parameter(Mandatory=$true)][string]$SourceMailbox,
    [switch]$IncludeArchive
  )

  $service = Get-ImtEwsServiceForMailbox -MailboxSmtp $EvidenceMailbox
  $folder = Get-ImtEwsFolderByPath -Service $service -MailboxSmtp $EvidenceMailbox -FolderPath $FolderPath
  if (-not $folder) {
    throw ("Evidence folder '{0}' was not found in mailbox '{1}'." -f $FolderPath, $EvidenceMailbox)
  }

  $propertySet = New-Object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::FirstClassProperties)
  [void]$propertySet.Add([Microsoft.Exchange.WebServices.Data.EmailMessageSchema]::InternetMessageId)
  [void]$propertySet.Add([Microsoft.Exchange.WebServices.Data.ItemSchema]::Attachments)

  $rows = New-Object System.Collections.Generic.List[object]
  $offset = 0
  $pageSize = 250

  do {
    $view = New-Object Microsoft.Exchange.WebServices.Data.ItemView($pageSize, $offset)
    $view.PropertySet = $propertySet

    $findResults = $service.FindItems($folder.Id, $view)
    if ($findResults.Items.Count -gt 0) {
      $null = $service.LoadPropertiesForItems($findResults.Items, $propertySet)
    }

    foreach ($item in @($findResults.Items)) {
      $attachmentNames = @(
        $item.Attachments |
          ForEach-Object { $_.Name } |
          Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
      )

      if ($attachmentNames -contains 'Search Results.csv') {
        continue
      }

      [void]$rows.Add([pscustomobject]@{
        SourceMailbox = $SourceMailbox
        MailboxLocation = if ($IncludeArchive) { 'PrimaryOrArchive' } else { 'Primary' }
        SentTime = $item.DateTimeSent
        From = if ($item.From) { $item.From.Address } else { $null }
        To = Join-ImtEwsAddressCollection -Recipients $item.ToRecipients
        Cc = Join-ImtEwsAddressCollection -Recipients $item.CcRecipients
        Subject = $item.Subject
        InternetMessageId = $item.InternetMessageId
        HasAttachments = [bool]$item.HasAttachments
        AttachmentCount = @($item.Attachments).Count
        ItemSize = $item.Size
        EvidenceFolder = $FolderPath
        TransportCorrelated = $false
        TrackingMessageId = $null
      })
    }

    $offset += @($findResults.Items).Count
  } while ($findResults.MoreAvailable)

  @($rows)
}

function Invoke-ImtDirectMailboxSearch {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory=$true)]$RunContext,
    [string[]]$BaseTargetAddresses,
    [string[]]$EffectiveSenderFilters
  )

  if (-not (Get-Command -Name Search-Mailbox -ErrorAction SilentlyContinue)) {
    return New-ImtModuleResult -StepName 'DirectMailboxSearch' -Status 'FAIL' -Summary 'Search-Mailbox cmdlet is unavailable in this session.' -Data ([pscustomobject]@{
      DirectRows = @()
      DirectKeywordRows = @()
      NonSearchable = @()
      ScopeRows = @()
      MatchedSourceMailboxAddresses = @()
      EvidenceRows = @()
      EvidenceErrors = @()
      DirectCsv = $null
      DirectKeywordCsv = $null
      NonSearchableCsv = $null
      ScopeCsv = $null
    }) -Metrics @{} -Errors @('Search-Mailbox cmdlet unavailable')
  }

  $scope = Resolve-ImtMailboxAuditScope -RunContext $RunContext -BaseTargetAddresses $BaseTargetAddresses
  $targetMailboxes = @($scope.Mailboxes)
  $nonSearchable = @($scope.Unresolved)
  $scopeRows = @($scope.ScopeRows)

  if ($targetMailboxes.Count -eq 0) {
    return New-ImtModuleResult -StepName 'DirectMailboxSearch' -Status 'WARN' -Summary 'No target mailboxes could be resolved for direct search.' -Data ([pscustomobject]@{
      DirectRows = @()
      DirectKeywordRows = @()
      NonSearchable = $nonSearchable
      ScopeRows = $scopeRows
      MatchedSourceMailboxAddresses = @()
      EvidenceRows = @()
      EvidenceErrors = @()
      DirectCsv = $null
      DirectKeywordCsv = $null
      NonSearchableCsv = $null
      ScopeCsv = $null
    }) -Metrics @{
      TargetMailboxes = 0
      NonSearchable = $nonSearchable.Count
    } -Errors @()
  }

  $searchQuery = New-ImtMailboxSearchQuery `
    -StartAt $RunContext.Start `
    -EndAt $RunContext.End `
    -SubjectContains $RunContext.Inputs.SubjectLike `
    -SubjectKeywords $RunContext.Inputs.Keywords `
    -SenderFilters $EffectiveSenderFilters `
    -RecipientFilters $RunContext.Inputs.Recipients `
    -OutboundOnly:$RunContext.Inputs.OutboundOnly `
    -RequireAttachment:$RunContext.Inputs.HasAttachmentOnly

  $dateOnlyQuery = New-ImtMailboxSearchQuery `
    -StartAt $RunContext.Start `
    -EndAt $RunContext.End `
    -SubjectContains $null `
    -SubjectKeywords @() `
    -SenderFilters $EffectiveSenderFilters `
    -RecipientFilters @() `
    -OutboundOnly:$RunContext.Inputs.OutboundOnly `
    -RequireAttachment:$RunContext.Inputs.HasAttachmentOnly

  Write-ImtLog -Level DEBUG -Step 'DirectMailboxSearch' -EventType Progress -Message ("Mailboxes={0}; Query={1}" -f $targetMailboxes.Count, $searchQuery)

  $directRows = New-Object System.Collections.Generic.List[object]
  $directKeywordRows = New-Object System.Collections.Generic.List[object]
  $evidenceRows = New-Object System.Collections.Generic.List[object]
  $evidenceErrors = New-Object System.Collections.Generic.List[string]
  $matchedMailboxSet = @{}

  $evidenceMailboxSmtp = $null
  if ($RunContext.Inputs.DetailedMailboxEvidence) {
    $evidenceMailboxSmtp = Resolve-ImtParticipantSmtp -Identity $RunContext.Inputs.EvidenceMailbox
    $evidenceMailbox = Resolve-ImtMailboxByAddress -Address $evidenceMailboxSmtp
    if (-not $evidenceMailbox) {
      return New-ImtModuleResult -StepName 'DirectMailboxSearch' -Status 'FAIL' -Summary ("Evidence mailbox '{0}' could not be resolved." -f $RunContext.Inputs.EvidenceMailbox) -Data ([pscustomobject]@{
        DirectRows = @()
        DirectKeywordRows = @()
        NonSearchable = $nonSearchable
        ScopeRows = $scopeRows
        MatchedSourceMailboxAddresses = @()
        EvidenceRows = @()
        EvidenceErrors = @()
        DirectCsv = $null
        DirectKeywordCsv = $null
        NonSearchableCsv = $null
        ScopeCsv = $null
      }) -Metrics @{} -Errors @('Evidence mailbox could not be resolved')
    }
  }

  foreach ($mailbox in $targetMailboxes) {
    $mailboxIdentity = $mailbox.Identity
    $mailboxSmtp = $mailbox.PrimarySmtpAddress.ToString()

    try {
      $combined = Get-ImtSearchMailboxEstimate -MailboxIdentity $mailboxIdentity -SearchQuery $searchQuery -SearchDumpster:$RunContext.Inputs.SearchDumpsterDirectly -IncludeArchive:$RunContext.Inputs.IncludeArchive
      if (-not $combined.Success) {
        throw $combined.Error
      }

      $dateOnly = Get-ImtSearchMailboxEstimate -MailboxIdentity $mailboxIdentity -SearchQuery $dateOnlyQuery -SearchDumpster:$RunContext.Inputs.SearchDumpsterDirectly -IncludeArchive:$RunContext.Inputs.IncludeArchive
      $dateOnlyCount = if ($dateOnly.Success -and $null -ne $dateOnly.Count) { [int]$dateOnly.Count } else { $null }

      $matchedKeywords = @()
      $keywordSummaryParts = @()
      $keywordTotalHits = 0

      if ($RunContext.Inputs.Keywords -and $RunContext.Inputs.Keywords.Count -gt 0) {
        foreach ($keyword in ($RunContext.Inputs.Keywords | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })) {
          $keywordQuery = New-ImtMailboxSearchQuery `
            -StartAt $RunContext.Start `
            -EndAt $RunContext.End `
            -SubjectContains $null `
            -SubjectKeywords @($keyword) `
            -SenderFilters $EffectiveSenderFilters `
            -RecipientFilters $RunContext.Inputs.Recipients `
            -OutboundOnly:$RunContext.Inputs.OutboundOnly `
            -RequireAttachment:$RunContext.Inputs.HasAttachmentOnly

          $keywordResult = Get-ImtSearchMailboxEstimate -MailboxIdentity $mailboxIdentity -SearchQuery $keywordQuery -SearchDumpster:$RunContext.Inputs.SearchDumpsterDirectly -IncludeArchive:$RunContext.Inputs.IncludeArchive

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

      if ($items -gt 0) {
        $matchedMailboxSet[$mailboxSmtp.ToLowerInvariant()] = $mailboxSmtp
      }

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

      if ($RunContext.Inputs.DetailedMailboxEvidence -and $items -gt 0) {
        $safeMailbox = ($mailboxSmtp -replace '[^\w@.-]', '_')
        $targetFolder = ("IMT_Evidence_{0}\{1}" -f $RunContext.Timestamp, $safeMailbox)
        $copyResult = Invoke-ImtSearchMailboxCopy -MailboxIdentity $mailboxIdentity -SearchQuery $searchQuery -TargetMailbox $evidenceMailboxSmtp -TargetFolder $targetFolder -SearchDumpster:$RunContext.Inputs.SearchDumpsterDirectly -IncludeArchive:$RunContext.Inputs.IncludeArchive

        if (-not $copyResult.Success) {
          [void]$evidenceErrors.Add(("Evidence copy failed for {0}: {1}" -f $mailboxSmtp, $copyResult.Error))
        } else {
          try {
            foreach ($evidenceRow in @(Get-ImtEvidenceItemsFromTargetMailbox -EvidenceMailbox $evidenceMailboxSmtp -FolderPath $targetFolder -SourceMailbox $mailboxSmtp -IncludeArchive:$RunContext.Inputs.IncludeArchive)) {
              [void]$evidenceRows.Add($evidenceRow)
            }
          } catch {
            [void]$evidenceErrors.Add(("Evidence enumeration failed for {0}: {1}" -f $mailboxSmtp, $_.Exception.Message))
          }
        }
      }

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

  $scopeCsv = $null
  if ($scopeRows.Count -gt 0) {
    $scopeCsv = Join-Path $RunContext.OutputDir ("MTL_DirectMailboxSearch_Scope_{0}.csv" -f $RunContext.Timestamp)
    $scopeRows | Export-Csv -Path $scopeCsv -NoTypeInformation -Encoding UTF8
  }

  $failedRows = @($directRows | Where-Object { $_.Status -eq 'Failed' }).Count
  $status = if ($failedRows -gt 0 -or $evidenceErrors.Count -gt 0) { 'WARN' } else { 'OK' }

  New-ImtModuleResult -StepName 'DirectMailboxSearch' -Status $status -Summary ("Direct mailbox search rows={0}; failed={1}; evidenceRows={2}; csv={3}" -f $directRows.Count, $failedRows, $evidenceRows.Count, $directCsv) -Data ([pscustomobject]@{
    DirectRows = $directRows.ToArray()
    DirectKeywordRows = @($directKeywordRowsAll)
    NonSearchable = @($nonSearchable)
    ScopeRows = @($scopeRows)
    MatchedSourceMailboxAddresses = @($matchedMailboxSet.Values | Sort-Object)
    EvidenceRows = $evidenceRows.ToArray()
    EvidenceErrors = $evidenceErrors.ToArray()
    DirectCsv = $directCsv
    DirectKeywordCsv = $directKeywordCsv
    NonSearchableCsv = $nonSearchableCsv
    ScopeCsv = $scopeCsv
  }) -Metrics @{
    TargetMailboxes = $targetMailboxes.Count
    DirectRows = $directRows.Count
    DirectKeywordRows = $directKeywordRowsAll.Count
    NonSearchable = $nonSearchable.Count
    ScopeRows = $scopeRows.Count
    MatchedMailboxes = @($matchedMailboxSet.Keys).Count
    EvidenceRows = $evidenceRows.Count
    EvidenceErrors = $evidenceErrors.Count
    FailedRows = $failedRows
  } -Errors @($evidenceErrors)
}
