Set-StrictMode -Version Latest

function Get-ImtTrackingPropertyValue {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory = $true)]
    [object]$InputObject,

    [Parameter(Mandatory = $true)]
    [string]$PropertyName
  )

  if ($null -eq $InputObject) {
    return $null
  }

  $property = $InputObject.PSObject.Properties[$PropertyName]
  if ($null -eq $property) {
    return $null
  }

  $property.Value
}

function Get-ImtTrackingFirstAvailablePropertyValue {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory = $true)]
    [object]$InputObject,

    [Parameter(Mandatory = $true)]
    [string[]]$PropertyNames
  )

  foreach ($propertyName in $PropertyNames) {
    $value = Get-ImtTrackingPropertyValue -InputObject $InputObject -PropertyName $propertyName
    if ($null -ne $value -and -not [string]::IsNullOrWhiteSpace(($value -as [string]))) {
      return $value
    }
  }

  $null
}

function Join-ImtTrackingDistinctValues {
  [CmdletBinding()]
  param(
    [AllowEmptyCollection()]
    [object[]]$Values,

    [string]$Separator = '; '
  )

  @(
    $Values |
      ForEach-Object { $_ -as [string] } |
      Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
      ForEach-Object { $_.Trim() } |
      Select-Object -Unique
  ) -join $Separator
}

function Get-ImtTrackingMessageKey {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory = $true)]
    [object]$Row
  )

  $messageId = Get-ImtTrackingPropertyValue -InputObject $Row -PropertyName 'MessageId'
  if (-not [string]::IsNullOrWhiteSpace(($messageId -as [string]))) {
    return ('mid:{0}' -f $messageId.ToString().Trim().ToLowerInvariant())
  }

  $internalMessageId = Get-ImtTrackingPropertyValue -InputObject $Row -PropertyName 'InternalMessageId'
  if (-not [string]::IsNullOrWhiteSpace(($internalMessageId -as [string]))) {
    return ('imid:{0}' -f $internalMessageId.ToString().Trim())
  }

  $timestamp = Get-ImtTrackingPropertyValue -InputObject $Row -PropertyName 'Timestamp'
  $sender = Get-ImtTrackingPropertyValue -InputObject $Row -PropertyName 'Sender'
  $subject = Get-ImtTrackingPropertyValue -InputObject $Row -PropertyName 'MessageSubject'

  'fallback:{0}|{1}|{2}' -f (($timestamp -as [datetime]).ToString('o')), ($sender -as [string]), ($subject -as [string])
}

function Select-ImtPrimaryTrackingRow {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory = $true)]
    [AllowEmptyCollection()]
    [object[]]$Rows
  )

  $precedence = @{
    SUBMIT = 1
    SEND = 2
    RECEIVE = 3
    DELIVER = 4
    RESOLVE = 5
    EXPAND = 6
    FAIL = 7
    DEFER = 8
  }

  @($Rows) |
    Sort-Object `
      @{ Expression = {
          $eventId = (Get-ImtTrackingPropertyValue -InputObject $_ -PropertyName 'EventId') -as [string]
          $normalized = if ($eventId) { $eventId.Trim().ToUpperInvariant() } else { '' }
          if ($precedence.ContainsKey($normalized)) { $precedence[$normalized] } else { 100 }
        }
      }, `
      @{ Expression = {
          $timestamp = Get-ImtTrackingPropertyValue -InputObject $_ -PropertyName 'Timestamp'
          if ($timestamp) { [datetime]$timestamp } else { [datetime]::MaxValue }
        }
      } |
    Select-Object -First 1
}

function Get-ImtTrackingRecipientSummary {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory = $true)]
    [AllowEmptyCollection()]
    [object[]]$Rows
  )

  $recipientSet = New-Object System.Collections.Generic.List[string]

  foreach ($row in @($Rows)) {
    foreach ($recipient in @(Get-ImtTrackingPropertyValue -InputObject $row -PropertyName 'Recipients')) {
      if (-not [string]::IsNullOrWhiteSpace(($recipient -as [string]))) {
        [void]$recipientSet.Add($recipient.ToString().Trim())
      }
    }
  }

  Join-ImtTrackingDistinctValues -Values $recipientSet.ToArray()
}

function Get-ImtTrackingTrailHints {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory = $true)]
    [AllowEmptyCollection()]
    [object[]]$Rows
  )

  $clientHostnames = New-Object System.Collections.Generic.List[string]
  $clientIps = New-Object System.Collections.Generic.List[string]
  $connectorIds = New-Object System.Collections.Generic.List[string]
  $sources = New-Object System.Collections.Generic.List[string]
  $sourceContexts = New-Object System.Collections.Generic.List[string]
  $serverNames = New-Object System.Collections.Generic.List[string]

  foreach ($row in @($Rows)) {
    $clientHostname = Get-ImtTrackingPropertyValue -InputObject $row -PropertyName 'ClientHostname'
    if (-not [string]::IsNullOrWhiteSpace(($clientHostname -as [string]))) {
      [void]$clientHostnames.Add($clientHostname.ToString().Trim())
    }

    $clientIp = Get-ImtTrackingFirstAvailablePropertyValue -InputObject $row -PropertyNames @('ClientIP', 'ClientIp', 'OriginalClientIP', 'OriginalClientIp')
    if (-not [string]::IsNullOrWhiteSpace(($clientIp -as [string]))) {
      [void]$clientIps.Add($clientIp.ToString().Trim())
    }

    $connectorId = Get-ImtTrackingPropertyValue -InputObject $row -PropertyName 'ConnectorId'
    if (-not [string]::IsNullOrWhiteSpace(($connectorId -as [string]))) {
      [void]$connectorIds.Add($connectorId.ToString().Trim())
    }

    $source = Get-ImtTrackingPropertyValue -InputObject $row -PropertyName 'Source'
    if (-not [string]::IsNullOrWhiteSpace(($source -as [string]))) {
      [void]$sources.Add($source.ToString().Trim())
    }

    $sourceContext = Get-ImtTrackingPropertyValue -InputObject $row -PropertyName 'SourceContext'
    if (-not [string]::IsNullOrWhiteSpace(($sourceContext -as [string]))) {
      [void]$sourceContexts.Add($sourceContext.ToString().Trim())
    }

    $serverName = Get-ImtTrackingPropertyValue -InputObject $row -PropertyName 'ServerHostname'
    if (-not [string]::IsNullOrWhiteSpace(($serverName -as [string]))) {
      [void]$serverNames.Add($serverName.ToString().Trim())
    }
  }

  [pscustomobject]@{
    ClientHostname = Join-ImtTrackingDistinctValues -Values $clientHostnames.ToArray()
    ClientIPAddress = Join-ImtTrackingDistinctValues -Values $clientIps.ToArray()
    ConnectorIds = Join-ImtTrackingDistinctValues -Values $connectorIds.ToArray()
    Sources = Join-ImtTrackingDistinctValues -Values $sources.ToArray()
    SourceContextSample = Join-ImtTrackingDistinctValues -Values (@($sourceContexts.ToArray()) | Select-Object -First 3)
    ServerHostnames = Join-ImtTrackingDistinctValues -Values $serverNames.ToArray()
  }
}

function Find-ImtTrackingBestMailboxAuditMatch {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory = $true)]
    [pscustomobject]$Message,

    [AllowEmptyCollection()]
    [object[]]$AuditRows,

    [ValidateRange(1, 120)]
    [int]$WindowMinutes = 10
  )

  if (-not $AuditRows -or @($AuditRows).Count -eq 0) {
    return $null
  }

  $messageSubject = ($Message.Subject -as [string])
  $normalizedSubject = if ($messageSubject) { $messageSubject.Trim() } else { '' }
  $messageTimestamp = [datetime]$Message.SubmittedAt

  $scoredRows = New-Object System.Collections.Generic.List[object]

  foreach ($row in @($AuditRows)) {
    $lastAccessed = Get-ImtTrackingPropertyValue -InputObject $row -PropertyName 'LastAccessed'
    if (-not $lastAccessed) {
      continue
    }

    $lastAccessedDate = [datetime]$lastAccessed
    $deltaMinutes = [math]::Abs((New-TimeSpan -Start $messageTimestamp -End $lastAccessedDate).TotalMinutes)
    if ($deltaMinutes -gt $WindowMinutes) {
      continue
    }

    $itemSubject = (Get-ImtTrackingPropertyValue -InputObject $row -PropertyName 'ItemSubject') -as [string]
    $normalizedItemSubject = if ($itemSubject) { $itemSubject.Trim() } else { '' }

    $exactSubjectMatch = $false
    $containsSubjectMatch = $false

    if (-not [string]::IsNullOrWhiteSpace($normalizedSubject) -and -not [string]::IsNullOrWhiteSpace($normalizedItemSubject)) {
      $exactSubjectMatch = $normalizedItemSubject.Equals($normalizedSubject, [System.StringComparison]::OrdinalIgnoreCase)
      if (-not $exactSubjectMatch) {
        $containsSubjectMatch = $normalizedItemSubject.IndexOf($normalizedSubject, [System.StringComparison]::OrdinalIgnoreCase) -ge 0
      }
    }

    $operation = (Get-ImtTrackingPropertyValue -InputObject $row -PropertyName 'Operation') -as [string]
    $normalizedOperation = if ($operation) { $operation.Trim() } else { '' }

    $hasClientData = @(
      (Get-ImtTrackingPropertyValue -InputObject $row -PropertyName 'ClientInfoString'),
      (Get-ImtTrackingPropertyValue -InputObject $row -PropertyName 'ClientIPAddress'),
      (Get-ImtTrackingPropertyValue -InputObject $row -PropertyName 'ClientMachineName'),
      (Get-ImtTrackingPropertyValue -InputObject $row -PropertyName 'ClientProcessName'),
      (Get-ImtTrackingPropertyValue -InputObject $row -PropertyName 'ClientVersion')
    ) | Where-Object { -not [string]::IsNullOrWhiteSpace(($_ -as [string])) }

    if (-not $exactSubjectMatch -and -not $containsSubjectMatch -and $normalizedOperation -notin @('SendAs', 'SendOnBehalf')) {
      continue
    }

    $score = 0

    if ($exactSubjectMatch) {
      $score += 100
    } elseif ($containsSubjectMatch) {
      $score += 60
    }

    switch ($normalizedOperation) {
      'SendAs' { $score += 50 }
      'SendOnBehalf' { $score += 50 }
      'Update' { $score += 20 }
      'Create' { $score += 10 }
      'Move' { $score += 10 }
      'MoveToDeletedItems' { $score += 5 }
    }

    if ($hasClientData.Count -gt 0) {
      $score += 20
    }

    $score += [math]::Max(0, (20 - [int][math]::Floor($deltaMinutes)))

    [void]$scoredRows.Add([pscustomobject]@{
      Row = $row
      Score = $score
      DeltaMinutes = [math]::Round($deltaMinutes, 2)
      ExactSubjectMatch = $exactSubjectMatch
      HasClientData = ($hasClientData.Count -gt 0)
      Operation = $normalizedOperation
    })
  }

  $bestMatch = $scoredRows |
    Sort-Object `
      @{ Expression = 'Score'; Descending = $true }, `
      @{ Expression = 'DeltaMinutes'; Descending = $false } |
    Select-Object -First 1

  if (-not $bestMatch) {
    return $null
  }

  $confidence = 'Low'
  if (($bestMatch.Operation -in @('SendAs', 'SendOnBehalf')) -or ($bestMatch.ExactSubjectMatch -and $bestMatch.HasClientData -and $bestMatch.DeltaMinutes -le 2)) {
    $confidence = 'High'
  } elseif ($bestMatch.ExactSubjectMatch -and $bestMatch.HasClientData) {
    $confidence = 'Medium'
  }

  [pscustomobject]@{
    Row = $bestMatch.Row
    Score = $bestMatch.Score
    DeltaMinutes = $bestMatch.DeltaMinutes
    Confidence = $confidence
  }
}

function Resolve-ImtTrackingLikelyClient {
  [CmdletBinding()]
  param(
    [string]$ClientInfoString,
    [string]$ClientProcessName,
    [string]$ClientHostname
  )

  $clientInfo = if ($ClientInfoString) { $ClientInfoString.Trim() } else { '' }
  $processName = if ($ClientProcessName) { $ClientProcessName.Trim() } else { '' }

  if ($processName -match '^OUTLOOK\.EXE$') {
    return 'Outlook desktop'
  }

  if ($clientInfo -match '(?i)activesync|airsync') {
    return 'Mobile client via ActiveSync'
  }

  if ($clientInfo -match '(?i)owa') {
    return 'Outlook on the web'
  }

  if ($clientInfo -match '(?i)mapi|rpc|msoutlook') {
    return 'Outlook desktop'
  }

  if (-not [string]::IsNullOrWhiteSpace($ClientHostname)) {
    return ('SMTP or relay host: {0}' -f $ClientHostname)
  }

  $null
}

function Resolve-ImtTrackingDeviceAssessment {
  [CmdletBinding()]
  param(
    [pscustomobject]$AuditMatch,
    [Parameter(Mandatory = $true)]
    [pscustomobject]$TrailHints
  )

  if ($AuditMatch) {
    $auditRow = $AuditMatch.Row
    $clientInfoString = (Get-ImtTrackingPropertyValue -InputObject $auditRow -PropertyName 'ClientInfoString') -as [string]
    $clientProcessName = (Get-ImtTrackingPropertyValue -InputObject $auditRow -PropertyName 'ClientProcessName') -as [string]
    $clientMachineName = (Get-ImtTrackingPropertyValue -InputObject $auditRow -PropertyName 'ClientMachineName') -as [string]
    $clientVersion = (Get-ImtTrackingPropertyValue -InputObject $auditRow -PropertyName 'ClientVersion') -as [string]
    $clientIpAddress = (Get-ImtTrackingPropertyValue -InputObject $auditRow -PropertyName 'ClientIPAddress') -as [string]

    return [pscustomobject]@{
      AttributionSource = 'MailboxAudit'
      AttributionConfidence = $AuditMatch.Confidence
      LikelyClient = Resolve-ImtTrackingLikelyClient -ClientInfoString $clientInfoString -ClientProcessName $clientProcessName -ClientHostname $TrailHints.ClientHostname
      ClientInfoString = $clientInfoString
      ClientProcessName = $clientProcessName
      ClientMachineName = $clientMachineName
      ClientVersion = $clientVersion
      ClientIPAddress = $clientIpAddress
      TransportClientHostname = $TrailHints.ClientHostname
      TransportClientIPAddress = $TrailHints.ClientIPAddress
      EvidenceNote = ('Mailbox audit match within {0} minute(s).' -f $AuditMatch.DeltaMinutes)
    }
  }

  if (-not [string]::IsNullOrWhiteSpace($TrailHints.ClientHostname) -or -not [string]::IsNullOrWhiteSpace($TrailHints.ClientIPAddress)) {
    return [pscustomobject]@{
      AttributionSource = 'Transport'
      AttributionConfidence = 'Low'
      LikelyClient = Resolve-ImtTrackingLikelyClient -ClientInfoString $null -ClientProcessName $null -ClientHostname $TrailHints.ClientHostname
      ClientInfoString = $null
      ClientProcessName = $null
      ClientMachineName = $null
      ClientVersion = $null
      ClientIPAddress = $TrailHints.ClientIPAddress
      TransportClientHostname = $TrailHints.ClientHostname
      TransportClientIPAddress = $TrailHints.ClientIPAddress
      EvidenceNote = 'Derived from message tracking only. Confirm with IIS, ActiveSync, MAPI/HTTP, IMAP, or SMTP protocol logs if needed.'
    }
  }

  [pscustomobject]@{
    AttributionSource = 'Undetermined'
    AttributionConfidence = 'None'
    LikelyClient = $null
    ClientInfoString = $null
    ClientProcessName = $null
    ClientMachineName = $null
    ClientVersion = $null
    ClientIPAddress = $null
    TransportClientHostname = $TrailHints.ClientHostname
    TransportClientIPAddress = $TrailHints.ClientIPAddress
    EvidenceNote = 'Exchange tracking did not expose a client host, and no correlated mailbox audit row was found.'
  }
}

function Get-ImtMailboxAuditQueryParameters {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory = $true)]
    [string]$Identity,

    [Parameter(Mandatory = $true)]
    [datetime]$StartDate,

    [Parameter(Mandatory = $true)]
    [datetime]$EndDate
  )

  $params = @{
    Identity = $Identity
    LogonTypes = @('Owner', 'Delegate', 'Admin')
    ShowDetails = $true
    StartDate = $StartDate
    EndDate = $EndDate
    ErrorAction = 'Stop'
  }

  $command = Get-Command -Name Search-MailboxAuditLog -ErrorAction SilentlyContinue
  if (-not $command) {
    return $params
  }

  $resultSizeParameter = $command.Parameters['ResultSize']
  if (-not $resultSizeParameter) {
    return $params
  }

  $parameterType = $resultSizeParameter.ParameterType
  if ($parameterType -eq [int] -or $parameterType -eq [int32]) {
    $params.ResultSize = 250000
    return $params
  }

  if (
    $parameterType -eq [string] -or
    $parameterType -eq [object] -or
    (($parameterType.FullName -as [string]) -match 'Unlimited')
  ) {
    $params.ResultSize = 'Unlimited'
  }

  $params
}

function Invoke-ImtMessageClientAccessAudit {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory = $true)]
    $RunContext,

    [Parameter(Mandatory = $true)]
    [AllowEmptyCollection()]
    [object[]]$Results,

    [AllowEmptyCollection()]
    [string[]]$CandidateMailboxAddresses
  )

  if (@($Results).Count -eq 0) {
    return New-ImtModuleResult -StepName 'MessageClientAccess' -Status 'SKIP' -Summary 'Client access correlation skipped because there were no tracking results.' -Data ([pscustomobject]@{
      Rows = @()
      AuditRows = @()
      AuditAvailable = $false
      AuditFailures = @()
    }) -Metrics @{
      MessageRows = 0
      MailboxAuditRows = 0
      AuditFailures = 0
    } -Errors @()
  }

  $correlationWindowMinutes = 10
  $candidateMailboxSet = @{}
  foreach ($address in @($CandidateMailboxAddresses)) {
    $normalizedAddress = ($address -as [string])
    if ($normalizedAddress) {
      $candidateMailboxSet[$normalizedAddress.Trim().ToLowerInvariant()] = $true
    }
  }

  $senderSet = @{}
  foreach ($result in @($Results)) {
    $senderValue = (Get-ImtTrackingPropertyValue -InputObject $result -PropertyName 'Sender') -as [string]
    if (-not [string]::IsNullOrWhiteSpace($senderValue)) {
      $senderSet[$senderValue.Trim().ToLowerInvariant()] = $true
    }
  }

  $auditAvailable = $null -ne (Get-Command -Name Search-MailboxAuditLog -ErrorAction SilentlyContinue)
  $auditFailures = New-Object System.Collections.Generic.List[string]
  $auditRows = New-Object System.Collections.Generic.List[object]
  $auditRowsBySender = @{}

  foreach ($sender in @($senderSet.Keys | Sort-Object)) {
    $shouldAttemptAudit = ($candidateMailboxSet.Count -eq 0) -or $candidateMailboxSet.ContainsKey($sender)
    if (-not $shouldAttemptAudit) {
      continue
    }

    if (-not $auditAvailable) {
      continue
    }

    $resolvedMailbox = Resolve-ImtMailboxByAddress -Address $sender
    if (-not $resolvedMailbox) {
      Write-ImtLog -Level DEBUG -Step 'MessageClientAccess' -EventType Progress -Message ("Skipping mailbox audit correlation for sender '{0}' because it could not be resolved to a local mailbox." -f $sender)
      continue
    }

    try {
      Write-ImtLog -Level DEBUG -Step 'MessageClientAccess' -EventType Progress -Message ("Querying mailbox audit log for sender '{0}'." -f $sender)
      $mailboxAuditParams = Get-ImtMailboxAuditQueryParameters `
        -Identity $sender `
        -StartDate $RunContext.Start.AddMinutes(-$correlationWindowMinutes) `
        -EndDate $RunContext.End.AddMinutes($correlationWindowMinutes)
      $senderAuditRows = @(
        Search-MailboxAuditLog @mailboxAuditParams
      )

      $auditRowsBySender[$sender] = @($senderAuditRows)
      foreach ($auditRow in @($senderAuditRows)) {
        [void]$auditRows.Add($auditRow)
      }
    } catch {
      [void]$auditFailures.Add(('{0}: {1}' -f $sender, $_.Exception.Message))
      Write-ImtLog -Level WARN -Step 'MessageClientAccess' -EventType Progress -Message ("Mailbox audit correlation failed for sender '{0}': {1}" -f $sender, $_.Exception.Message)
    }
  }

  if (-not $auditAvailable) {
    Write-ImtLog -Level WARN -Step 'MessageClientAccess' -EventType Progress -Message 'Search-MailboxAuditLog is unavailable in this session. Falling back to tracking-only device hints.'
  }

  $messageRows = New-Object System.Collections.Generic.List[object]

  foreach ($group in @($Results | Group-Object { Get-ImtTrackingMessageKey -Row $_ })) {
    $groupRows = @($group.Group | Sort-Object Timestamp)
    $primaryRow = Select-ImtPrimaryTrackingRow -Rows $groupRows
    if (-not $primaryRow) {
      continue
    }

    $submittedAt = [datetime](Get-ImtTrackingPropertyValue -InputObject $primaryRow -PropertyName 'Timestamp')
    $sender = (Get-ImtTrackingPropertyValue -InputObject $primaryRow -PropertyName 'Sender') -as [string]
    $normalizedSender = if ($sender) { $sender.Trim().ToLowerInvariant() } else { $null }
    $trailHints = Get-ImtTrackingTrailHints -Rows $groupRows
    $eventIdsSeen = Join-ImtTrackingDistinctValues -Values @(
      $groupRows |
        ForEach-Object { Get-ImtTrackingPropertyValue -InputObject $_ -PropertyName 'EventId' }
    )

    $senderAuditRows = if ($normalizedSender -and $auditRowsBySender.ContainsKey($normalizedSender)) {
      @($auditRowsBySender[$normalizedSender])
    } else {
      @()
    }

    $auditMatch = Find-ImtTrackingBestMailboxAuditMatch -Message ([pscustomobject]@{
        SubmittedAt = $submittedAt
        Subject = (Get-ImtTrackingPropertyValue -InputObject $primaryRow -PropertyName 'MessageSubject') -as [string]
      }) `
      -AuditRows $senderAuditRows `
      -WindowMinutes $correlationWindowMinutes

    $deviceAssessment = Resolve-ImtTrackingDeviceAssessment -AuditMatch $auditMatch -TrailHints $trailHints

    [void]$messageRows.Add([pscustomobject]@{
      Mailbox = $sender
      SubmittedAt = $submittedAt
      PrimaryEventId = (Get-ImtTrackingPropertyValue -InputObject $primaryRow -PropertyName 'EventId') -as [string]
      EventIdsSeen = $eventIdsSeen
      Sender = $sender
      Recipients = Get-ImtTrackingRecipientSummary -Rows $groupRows
      Subject = (Get-ImtTrackingPropertyValue -InputObject $primaryRow -PropertyName 'MessageSubject') -as [string]
      MessageId = (Get-ImtTrackingPropertyValue -InputObject $primaryRow -PropertyName 'MessageId') -as [string]
      InternalMessageId = (Get-ImtTrackingPropertyValue -InputObject $primaryRow -PropertyName 'InternalMessageId') -as [string]
      ServerHostnames = $trailHints.ServerHostnames
      ConnectorIds = $trailHints.ConnectorIds
      SourceTypes = $trailHints.Sources
      SourceContextSample = $trailHints.SourceContextSample
      AttributionSource = $deviceAssessment.AttributionSource
      AttributionConfidence = $deviceAssessment.AttributionConfidence
      LikelyClient = $deviceAssessment.LikelyClient
      ClientInfoString = $deviceAssessment.ClientInfoString
      ClientProcessName = $deviceAssessment.ClientProcessName
      ClientMachineName = $deviceAssessment.ClientMachineName
      ClientVersion = $deviceAssessment.ClientVersion
      ClientIPAddress = $deviceAssessment.ClientIPAddress
      TransportClientHostname = $deviceAssessment.TransportClientHostname
      TransportClientIPAddress = $deviceAssessment.TransportClientIPAddress
      EvidenceNote = $deviceAssessment.EvidenceNote
      MailboxAuditOperation = if ($auditMatch) { (Get-ImtTrackingPropertyValue -InputObject $auditMatch.Row -PropertyName 'Operation') -as [string] } else { $null }
      MailboxAuditLogonType = if ($auditMatch) { (Get-ImtTrackingPropertyValue -InputObject $auditMatch.Row -PropertyName 'LogonType') -as [string] } else { $null }
      MailboxAuditLastAccessed = if ($auditMatch) { [datetime](Get-ImtTrackingPropertyValue -InputObject $auditMatch.Row -PropertyName 'LastAccessed') } else { $null }
      MailboxAuditDeltaMinutes = if ($auditMatch) { $auditMatch.DeltaMinutes } else { $null }
    })
  }

  $rowArray = @($messageRows.ToArray() | Sort-Object Mailbox, SubmittedAt, Subject)
  $confidenceCounts = @{
    High = @($rowArray | Where-Object { $_.AttributionConfidence -eq 'High' }).Count
    Medium = @($rowArray | Where-Object { $_.AttributionConfidence -eq 'Medium' }).Count
    Low = @($rowArray | Where-Object { $_.AttributionConfidence -eq 'Low' }).Count
    None = @($rowArray | Where-Object { $_.AttributionConfidence -eq 'None' }).Count
  }

  $status = if (-not $auditAvailable -or $auditFailures.Count -gt 0 -or $confidenceCounts.None -gt 0) { 'WARN' } else { 'OK' }
  $summary = "Client attribution rows={0}; High={1}; Medium={2}; Low={3}; None={4}; MailboxAuditRows={5}; AuditFailures={6}" -f `
    $rowArray.Count, `
    $confidenceCounts.High, `
    $confidenceCounts.Medium, `
    $confidenceCounts.Low, `
    $confidenceCounts.None, `
    $auditRows.Count, `
    $auditFailures.Count

  New-ImtModuleResult -StepName 'MessageClientAccess' -Status $status -Summary $summary -Data ([pscustomobject]@{
    Rows = $rowArray
    AuditRows = @($auditRows.ToArray())
    AuditAvailable = [bool]$auditAvailable
    AuditFailures = @($auditFailures.ToArray())
  }) -Metrics @{
    MessageRows = $rowArray.Count
    MailboxAuditRows = $auditRows.Count
    HighConfidence = $confidenceCounts.High
    MediumConfidence = $confidenceCounts.Medium
    LowConfidence = $confidenceCounts.Low
    NoConfidence = $confidenceCounts.None
    AuditFailures = $auditFailures.Count
  } -Errors @($auditFailures)
}
