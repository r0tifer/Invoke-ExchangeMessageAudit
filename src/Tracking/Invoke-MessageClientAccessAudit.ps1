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
  $clientTypes = New-Object System.Collections.Generic.List[string]
  $submissionAssistants = New-Object System.Collections.Generic.List[string]

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
      $normalizedSourceContext = $sourceContext.ToString().Trim()
      [void]$sourceContexts.Add($normalizedSourceContext)

      foreach ($match in @([regex]::Matches($normalizedSourceContext, '(?i)(?:^|[,;]\s*)ClientType:(?<Value>[^,;]+)'))) {
        $value = ($match.Groups['Value'].Value -as [string])
        if (-not [string]::IsNullOrWhiteSpace($value)) {
          [void]$clientTypes.Add($value.Trim())
        }
      }

      foreach ($match in @([regex]::Matches($normalizedSourceContext, '(?i)(?:^|[,;]\s*)SubmissionAssistant:(?<Value>[^,;]+)'))) {
        $value = ($match.Groups['Value'].Value -as [string])
        if (-not [string]::IsNullOrWhiteSpace($value)) {
          [void]$submissionAssistants.Add($value.Trim())
        }
      }
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
    TrackingClientType = Join-ImtTrackingDistinctValues -Values $clientTypes.ToArray()
    SubmissionAssistant = Join-ImtTrackingDistinctValues -Values $submissionAssistants.ToArray()
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
    [string]$TrackingClientType,
    [string]$ClientHostname
  )

  $clientInfo = if ($ClientInfoString) { $ClientInfoString.Trim() } else { '' }
  $processName = if ($ClientProcessName) { $ClientProcessName.Trim() } else { '' }
  $trackingClientTypeValue = if ($TrackingClientType) { $TrackingClientType.Trim() } else { '' }

  if ($processName -match '^OUTLOOK\.EXE$') {
    return 'Outlook desktop'
  }

  if ($trackingClientTypeValue -match '(?i)activesync|airsync') {
    return 'Mobile client via ActiveSync'
  }

  if ($trackingClientTypeValue -match '(?i)owa') {
    return 'Outlook on the web'
  }

  if ($trackingClientTypeValue -match '(?i)mapi|rpc|outlook') {
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
    [pscustomobject]$ProtocolMatch,
    [pscustomobject]$ActiveSyncMatch,
    [Parameter(Mandatory = $true)]
    [pscustomobject]$TrailHints
  )

  $trailTrackingClientType = (Get-ImtTrackingPropertyValue -InputObject $TrailHints -PropertyName 'TrackingClientType') -as [string]
  $trailSubmissionAssistant = (Get-ImtTrackingPropertyValue -InputObject $TrailHints -PropertyName 'SubmissionAssistant') -as [string]

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
      LikelyClient = Resolve-ImtTrackingLikelyClient -ClientInfoString $clientInfoString -ClientProcessName $clientProcessName -TrackingClientType $trailTrackingClientType -ClientHostname $TrailHints.ClientHostname
      ClientInfoString = $clientInfoString
      ClientProcessName = $clientProcessName
      ClientMachineName = $clientMachineName
      ClientVersion = $clientVersion
      ClientIPAddress = $clientIpAddress
      TransportClientHostname = $TrailHints.ClientHostname
      TransportClientIPAddress = $TrailHints.ClientIPAddress
      ProtocolEvidenceType = $null
      ProtocolLogServer = $null
      ProtocolLogPath = $null
      ProtocolUserAgent = $null
      ProtocolRemoteEndpoint = $null
      ProtocolTimestamp = $null
      ProtocolDeltaMinutes = $null
      EvidenceNote = ('Mailbox audit match within {0} minute(s).' -f $AuditMatch.DeltaMinutes)
    }
  }

  if ($ProtocolMatch) {
    $protocolRow = $ProtocolMatch.Row
    $evidenceType = (Get-ImtTrackingPropertyValue -InputObject $protocolRow -PropertyName 'EvidenceType') -as [string]
    $userAgent = (Get-ImtTrackingPropertyValue -InputObject $protocolRow -PropertyName 'UserAgent') -as [string]
    $clientIpAddress = (Get-ImtTrackingPropertyValue -InputObject $protocolRow -PropertyName 'ClientIPAddress') -as [string]
    $remoteEndpoint = (Get-ImtTrackingPropertyValue -InputObject $protocolRow -PropertyName 'RemoteEndpoint') -as [string]
    $likelyClient = switch -Regex ($evidenceType) {
      '^HttpProxyMapi$' { 'Outlook desktop'; break }
      '^HttpProxyOwa$' { 'Outlook on the web'; break }
      '^HttpProxyEas$' { 'Mobile client via ActiveSync'; break }
      '^HttpProxyEws$' { 'Exchange Web Services client'; break }
      '^SmtpReceive' {
        if (-not [string]::IsNullOrWhiteSpace($remoteEndpoint)) {
          "SMTP submission from $remoteEndpoint"
        } elseif (-not [string]::IsNullOrWhiteSpace($clientIpAddress)) {
          "SMTP submission from $clientIpAddress"
        } else {
          'SMTP submission'
        }
        break
      }
      default {
        Resolve-ImtTrackingLikelyClient -ClientInfoString $userAgent -ClientProcessName $null -TrackingClientType $trailTrackingClientType -ClientHostname $TrailHints.ClientHostname
      }
    }

    return [pscustomobject]@{
      AttributionSource = 'ProtocolLog'
      AttributionConfidence = $ProtocolMatch.Confidence
      LikelyClient = $likelyClient
      ClientInfoString = $userAgent
      ClientProcessName = $null
      ClientMachineName = $null
      ClientVersion = $null
      ClientIPAddress = if (-not [string]::IsNullOrWhiteSpace($clientIpAddress)) { $clientIpAddress } else { $remoteEndpoint }
      TransportClientHostname = $TrailHints.ClientHostname
      TransportClientIPAddress = $TrailHints.ClientIPAddress
      ProtocolEvidenceType = $evidenceType
      ProtocolLogServer = (Get-ImtTrackingPropertyValue -InputObject $protocolRow -PropertyName 'Server') -as [string]
      ProtocolLogPath = (Get-ImtTrackingPropertyValue -InputObject $protocolRow -PropertyName 'LogPath') -as [string]
      ProtocolUserAgent = $userAgent
      ProtocolRemoteEndpoint = $remoteEndpoint
      ProtocolTimestamp = Get-ImtTrackingPropertyValue -InputObject $protocolRow -PropertyName 'Timestamp'
      ProtocolDeltaMinutes = $ProtocolMatch.DeltaMinutes
      EvidenceNote = ('Protocol log match ({0}) within {1} minute(s).' -f $evidenceType, $ProtocolMatch.DeltaMinutes)
    }
  }

  if ($ActiveSyncMatch) {
    $activeSyncRow = $ActiveSyncMatch.Row
    $deviceId = (Get-ImtTrackingPropertyValue -InputObject $activeSyncRow -PropertyName 'DeviceId') -as [string]
    $deviceType = (Get-ImtTrackingPropertyValue -InputObject $activeSyncRow -PropertyName 'DeviceType') -as [string]
    $deviceModel = (Get-ImtTrackingPropertyValue -InputObject $activeSyncRow -PropertyName 'DeviceModel') -as [string]
    $deviceOs = (Get-ImtTrackingPropertyValue -InputObject $activeSyncRow -PropertyName 'DeviceOS') -as [string]
    $deviceFriendlyName = (Get-ImtTrackingPropertyValue -InputObject $activeSyncRow -PropertyName 'DeviceFriendlyName') -as [string]
    $deviceUserAgent = (Get-ImtTrackingPropertyValue -InputObject $activeSyncRow -PropertyName 'DeviceUserAgent') -as [string]
    $clientType = (Get-ImtTrackingPropertyValue -InputObject $activeSyncRow -PropertyName 'ClientType') -as [string]
    $lastSuccessSync = Get-ImtTrackingPropertyValue -InputObject $activeSyncRow -PropertyName 'LastSuccessSync'
    $lastSyncAttemptTime = Get-ImtTrackingPropertyValue -InputObject $activeSyncRow -PropertyName 'LastSyncAttemptTime'

    $clientInfoParts = New-Object System.Collections.Generic.List[string]
    foreach ($entry in @(
        if (-not [string]::IsNullOrWhiteSpace($clientType)) { 'ClientType={0}' -f $clientType.Trim() } else { $null }
        if (-not [string]::IsNullOrWhiteSpace($deviceUserAgent)) { 'UserAgent={0}' -f $deviceUserAgent.Trim() } else { $null }
        if (-not [string]::IsNullOrWhiteSpace($deviceId)) { 'DeviceId={0}' -f $deviceId.Trim() } else { $null }
      )) {
      if (-not [string]::IsNullOrWhiteSpace(($entry -as [string]))) {
        [void]$clientInfoParts.Add($entry)
      }
    }

    $evidenceNote = 'Matched ActiveSync device partnership details.'
    if ($ActiveSyncMatch.TimeProperty -and $ActiveSyncMatch.DeltaMinutes -ne $null) {
      $evidenceNote = ('Matched ActiveSync device evidence using {0} within {1} minute(s).' -f $ActiveSyncMatch.TimeProperty, $ActiveSyncMatch.DeltaMinutes)
    } elseif ($ActiveSyncMatch.DeltaMinutes -ne $null) {
      $evidenceNote = ('Matched ActiveSync device evidence within {0} minute(s).' -f $ActiveSyncMatch.DeltaMinutes)
    }

    $clientMachineName = $null
    foreach ($value in @($deviceFriendlyName, $deviceModel, $deviceType)) {
      if (-not [string]::IsNullOrWhiteSpace(($value -as [string]))) {
        $clientMachineName = $value.Trim()
        break
      }
    }

    return [pscustomobject]@{
      AttributionSource = 'ActiveSyncDevice'
      AttributionConfidence = $ActiveSyncMatch.Confidence
      LikelyClient = Format-ImtActiveSyncLikelyClient -DeviceRow $activeSyncRow
      ClientInfoString = Join-ImtTrackingDistinctValues -Values $clientInfoParts.ToArray()
      ClientProcessName = $null
      ClientMachineName = $clientMachineName
      ClientVersion = $deviceOs
      ClientIPAddress = $TrailHints.ClientIPAddress
      TransportClientHostname = $TrailHints.ClientHostname
      TransportClientIPAddress = $TrailHints.ClientIPAddress
      ProtocolEvidenceType = $null
      ProtocolLogServer = $null
      ProtocolLogPath = $null
      ProtocolUserAgent = $null
      ProtocolRemoteEndpoint = $null
      ProtocolTimestamp = $null
      ProtocolDeltaMinutes = $null
      ActiveSyncDeviceId = $deviceId
      ActiveSyncDeviceType = $deviceType
      ActiveSyncDeviceModel = $deviceModel
      ActiveSyncDeviceOS = $deviceOs
      ActiveSyncDeviceFriendlyName = $deviceFriendlyName
      ActiveSyncDeviceUserAgent = $deviceUserAgent
      ActiveSyncClientType = $clientType
      ActiveSyncLastSuccessSync = $lastSuccessSync
      ActiveSyncLastSyncAttemptTime = $lastSyncAttemptTime
      ActiveSyncDeltaMinutes = $ActiveSyncMatch.DeltaMinutes
      EvidenceNote = $evidenceNote
    }
  }

  if (-not [string]::IsNullOrWhiteSpace($trailTrackingClientType)) {
    $clientType = $trailTrackingClientType
    $submissionAssistant = $trailSubmissionAssistant
    $evidenceNote = 'Derived from message tracking SourceContext.'
    if (-not [string]::IsNullOrWhiteSpace($submissionAssistant)) {
      $evidenceNote = ('Derived from message tracking SourceContext (ClientType={0}; SubmissionAssistant={1}).' -f $clientType, $submissionAssistant)
    } else {
      $evidenceNote = ('Derived from message tracking SourceContext (ClientType={0}).' -f $clientType)
    }

    return [pscustomobject]@{
      AttributionSource = 'TrackingSourceContext'
      AttributionConfidence = 'Medium'
      LikelyClient = Resolve-ImtTrackingLikelyClient -ClientInfoString $null -ClientProcessName $null -TrackingClientType $clientType -ClientHostname $TrailHints.ClientHostname
      ClientInfoString = if (-not [string]::IsNullOrWhiteSpace($submissionAssistant)) { ('ClientType={0}; SubmissionAssistant={1}' -f $clientType, $submissionAssistant) } else { ('ClientType={0}' -f $clientType) }
      ClientProcessName = $null
      ClientMachineName = $null
      ClientVersion = $null
      ClientIPAddress = $TrailHints.ClientIPAddress
      TransportClientHostname = $TrailHints.ClientHostname
      TransportClientIPAddress = $TrailHints.ClientIPAddress
      ProtocolEvidenceType = $null
      ProtocolLogServer = $null
      ProtocolLogPath = $null
      ProtocolUserAgent = $null
      ProtocolRemoteEndpoint = $null
      ProtocolTimestamp = $null
      ProtocolDeltaMinutes = $null
      EvidenceNote = $evidenceNote
    }
  }

  if (-not [string]::IsNullOrWhiteSpace($TrailHints.ClientHostname) -or -not [string]::IsNullOrWhiteSpace($TrailHints.ClientIPAddress)) {
    return [pscustomobject]@{
      AttributionSource = 'Transport'
      AttributionConfidence = 'Low'
      LikelyClient = Resolve-ImtTrackingLikelyClient -ClientInfoString $null -ClientProcessName $null -TrackingClientType $trailTrackingClientType -ClientHostname $TrailHints.ClientHostname
      ClientInfoString = $null
      ClientProcessName = $null
      ClientMachineName = $null
      ClientVersion = $null
      ClientIPAddress = $TrailHints.ClientIPAddress
      TransportClientHostname = $TrailHints.ClientHostname
      TransportClientIPAddress = $TrailHints.ClientIPAddress
      ProtocolEvidenceType = $null
      ProtocolLogServer = $null
      ProtocolLogPath = $null
      ProtocolUserAgent = $null
      ProtocolRemoteEndpoint = $null
      ProtocolTimestamp = $null
      ProtocolDeltaMinutes = $null
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
    ProtocolEvidenceType = $null
    ProtocolLogServer = $null
    ProtocolLogPath = $null
    ProtocolUserAgent = $null
    ProtocolRemoteEndpoint = $null
    ProtocolTimestamp = $null
    ProtocolDeltaMinutes = $null
    EvidenceNote = 'Exchange tracking did not expose a client host, and no correlated mailbox audit, ActiveSync device, or protocol-log row was found.'
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

  # Exchange on-prem frequently exposes Search-MailboxAuditLog through a proxy
  # command whose local parameter metadata does not reflect the remote Int32 cap.
  # Use the server-accepted ceiling consistently instead of relying on proxy types.
  $params.ResultSize = 250000

  $params
}

function ConvertTo-ImtProtocolComparableTimestamp {
  [CmdletBinding()]
  param(
    [object]$Value
  )

  if ($null -eq $Value) {
    return $null
  }

  try {
    $timestamp = [datetime]$Value
    if ($timestamp.Kind -eq [System.DateTimeKind]::Utc) {
      return $timestamp.ToLocalTime()
    }

    $timestamp
  } catch {
    $null
  }
}

function Get-ImtExchangeInstallRoot {
  [CmdletBinding()]
  param()

  if (-not [string]::IsNullOrWhiteSpace($env:ExchangeInstallPath)) {
    return $env:ExchangeInstallPath.TrimEnd('\')
  }

  'C:\Program Files\Microsoft\Exchange Server\V15'
}

function Join-ImtWindowsPath {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory = $true)]
    [string]$BasePath,

    [Parameter(Mandatory = $true)]
    [string]$ChildPath
  )

  ('{0}\{1}' -f $BasePath.TrimEnd('\'), $ChildPath.TrimStart('\'))
}

function Test-ImtLiteralPathExists {
  [CmdletBinding()]
  param(
    [string]$Path
  )

  if ([string]::IsNullOrWhiteSpace($Path)) {
    return $false
  }

  try {
    [bool](Test-Path -LiteralPath $Path -ErrorAction Stop)
  } catch {
    $false
  }
}

function Get-ImtProtocolLogDirectoryPath {
  [CmdletBinding()]
  param(
    [string]$Server,
    [Parameter(Mandatory = $true)]
    [string]$RelativePath
  )

  $exchangeRoot = Get-ImtExchangeInstallRoot
  $qualifier = Split-Path -Path $exchangeRoot -Qualifier
  $relativeExchangeRoot = $exchangeRoot.Substring($qualifier.Length).TrimStart('\')
  $driveShare = '{0}$' -f $qualifier.TrimEnd(':')

  if ([string]::IsNullOrWhiteSpace($Server)) {
    return Join-ImtWindowsPath -BasePath $exchangeRoot -ChildPath $RelativePath
  }

  $serverValue = $Server.Trim()
  $localComputer = $env:COMPUTERNAME
  $isLocal = $false
  if (-not [string]::IsNullOrWhiteSpace($localComputer)) {
    $isLocal = $serverValue.Equals($localComputer, [System.StringComparison]::OrdinalIgnoreCase) -or
      $serverValue.Split('.')[0].Equals($localComputer, [System.StringComparison]::OrdinalIgnoreCase)
  }

  if ($isLocal) {
    return Join-ImtWindowsPath -BasePath $exchangeRoot -ChildPath $RelativePath
  }

  $uncRoot = '\\{0}\{1}' -f $serverValue, $driveShare
  $remoteExchangeRoot = if ([string]::IsNullOrWhiteSpace($relativeExchangeRoot)) {
    $uncRoot
  } else {
    Join-ImtWindowsPath -BasePath $uncRoot -ChildPath $relativeExchangeRoot
  }

  Join-ImtWindowsPath -BasePath $remoteExchangeRoot -ChildPath $RelativePath
}

function Get-ImtTimeWindowTokens {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory = $true)]
    [datetime]$StartDate,

    [Parameter(Mandatory = $true)]
    [datetime]$EndDate,

    [Parameter(Mandatory = $true)]
    [ValidateSet('Hour', 'Day')]
    [string]$Granularity
  )

  $tokens = New-Object System.Collections.Generic.List[string]

  if ($Granularity -eq 'Hour') {
    $cursor = [datetime]::new($StartDate.Year, $StartDate.Month, $StartDate.Day, $StartDate.Hour, 0, 0)
    while ($cursor -le $EndDate) {
      [void]$tokens.Add($cursor.ToString('yyyyMMddHH'))
      $cursor = $cursor.AddHours(1)
    }
  } else {
    $cursor = [datetime]::new($StartDate.Year, $StartDate.Month, $StartDate.Day, 0, 0, 0)
    while ($cursor -le $EndDate) {
      [void]$tokens.Add($cursor.ToString('yyyyMMdd'))
      $cursor = $cursor.AddDays(1)
    }
  }

  @($tokens.ToArray() | Select-Object -Unique)
}

function Get-ImtProtocolWindowTokens {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory = $true)]
    [datetime]$StartDate,

    [Parameter(Mandatory = $true)]
    [datetime]$EndDate,

    [Parameter(Mandatory = $true)]
    [ValidateSet('Hour', 'Day')]
    [string]$Granularity
  )

  $tokenSets = New-Object System.Collections.Generic.List[string]

  foreach ($token in @(Get-ImtTimeWindowTokens -StartDate $StartDate -EndDate $EndDate -Granularity $Granularity)) {
    [void]$tokenSets.Add($token)
  }

  $utcStart = $StartDate.ToUniversalTime()
  $utcEnd = $EndDate.ToUniversalTime()
  foreach ($token in @(Get-ImtTimeWindowTokens -StartDate $utcStart -EndDate $utcEnd -Granularity $Granularity)) {
    [void]$tokenSets.Add($token)
  }

  @($tokenSets.ToArray() | Select-Object -Unique)
}

function Get-ImtExchangeCsvHeaders {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory = $true)]
    [string]$Path
  )

  $headerLine = @(
    Select-String -Path $Path -Pattern '^#Fields:' -ErrorAction Stop
  ) | Select-Object -Last 1

  if (-not $headerLine) {
    return @()
  }

  $headerText = $headerLine.Line.Substring(8).Trim()
  if ([string]::IsNullOrWhiteSpace($headerText)) {
    return @()
  }

  if ($headerText -notmatch ',') {
    return @()
  }

  @(
    $headerText -split '\s*,\s*' |
      Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
  )
}

function ConvertFrom-ImtExchangeCsvMatch {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory = $true)]
    [string]$Line,

    [Parameter(Mandatory = $true)]
    [string[]]$Headers
  )

  if ([string]::IsNullOrWhiteSpace($Line) -or $Line -match '^#' -or $Headers.Count -eq 0) {
    return $null
  }

  try {
    $Line | ConvertFrom-Csv -Header $Headers
  } catch {
    $null
  }
}

function Get-ImtProtocolSearchPatterns {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory = $true)]
    [string]$Sender,

    $ResolvedMailbox
  )

  $patterns = New-Object System.Collections.Generic.List[string]
  [void]$patterns.Add($Sender.Trim())

  if ($ResolvedMailbox) {
    $exchangeGuid = Get-ImtTrackingFirstAvailablePropertyValue -InputObject $ResolvedMailbox -PropertyNames @('ExchangeGuid', 'Guid')
    if ($exchangeGuid) {
      $guidText = ($exchangeGuid -as [string])
      if (-not [string]::IsNullOrWhiteSpace($guidText)) {
        $normalizedGuid = $guidText.Trim()
        [void]$patterns.Add(('MailboxId={0}@' -f $normalizedGuid))
        [void]$patterns.Add($normalizedGuid)
      }
    }
  }

  @($patterns.ToArray() | Select-Object -Unique)
}

function Get-ImtProtocolCandidateFiles {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory = $true)]
    [string]$DirectoryPath,

    [Parameter(Mandatory = $true)]
    [string[]]$Tokens
  )

  if (-not (Test-ImtLiteralPathExists -Path $DirectoryPath)) {
    return @()
  }

  $files = New-Object System.Collections.Generic.List[string]
  $normalizedTokens = @(
    $Tokens |
      ForEach-Object { $_ -as [string] } |
      Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
      Select-Object -Unique
  )

  foreach ($token in $normalizedTokens) {
    foreach ($path in @(
        Get-ChildItem -LiteralPath $DirectoryPath -Filter ("*{0}*" -f $token) -ErrorAction Stop |
          Where-Object {
            ($_.PSObject.Properties.Match('PSIsContainer').Count -eq 0) -or (-not $_.PSIsContainer)
          } |
          Select-Object -ExpandProperty FullName
      )) {
      [void]$files.Add($path)
    }
  }

  @($files.ToArray() | Select-Object -Unique)
}

function Get-ImtProtocolRowsFromFiles {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory = $true)]
    [string[]]$Paths,

    [Parameter(Mandatory = $true)]
    [string[]]$Patterns,

    [Parameter(Mandatory = $true)]
    [string]$Mailbox,

    [Parameter(Mandatory = $true)]
    [string]$EvidenceType,

    [string]$Server,

    [Parameter(Mandatory = $true)]
    [datetime]$StartDate,

    [Parameter(Mandatory = $true)]
    [datetime]$EndDate
  )

  $rows = New-Object System.Collections.Generic.List[object]
  $pathSet = @(
    $Paths |
      ForEach-Object { $_ -as [string] } |
      Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
      Select-Object -Unique
  )
  if ($pathSet.Count -eq 0) {
    return @()
  }

  $matchesByPath = @(
    Select-String -Path $pathSet -Pattern $Patterns -SimpleMatch -ErrorAction Stop |
      Group-Object Path
  )

  foreach ($matchGroup in $matchesByPath) {
    $path = $matchGroup.Name
    $headers = @(Get-ImtExchangeCsvHeaders -Path $path)
    if ($headers.Count -eq 0) {
      continue
    }

    foreach ($match in @($matchGroup.Group)) {
      $parsedRow = ConvertFrom-ImtExchangeCsvMatch -Line $match.Line -Headers $headers
      if (-not $parsedRow) {
        continue
      }

      $timestamp = ConvertTo-ImtProtocolComparableTimestamp -Value (Get-ImtTrackingFirstAvailablePropertyValue -InputObject $parsedRow -PropertyNames @('date-time', 'DateTime', 'datetime'))
      if (-not $timestamp) {
        continue
      }

      if ($timestamp -lt $StartDate -or $timestamp -gt $EndDate) {
        continue
      }

      [void]$rows.Add([pscustomobject]@{
        Mailbox = $Mailbox
        EvidenceType = $EvidenceType
        Server = if (-not [string]::IsNullOrWhiteSpace($Server)) { $Server } else { (Get-ImtTrackingFirstAvailablePropertyValue -InputObject $parsedRow -PropertyNames @('ServerHostName', 'server-host-name', 's-computername')) }
        LogPath = $path
        Timestamp = $timestamp
        Protocol = (Get-ImtTrackingFirstAvailablePropertyValue -InputObject $parsedRow -PropertyNames @('Protocol', 'protocol'))
        UrlStem = (Get-ImtTrackingFirstAvailablePropertyValue -InputObject $parsedRow -PropertyNames @('UrlStem', 'url-stem', 'cs-uri-stem'))
        UserAgent = (Get-ImtTrackingFirstAvailablePropertyValue -InputObject $parsedRow -PropertyNames @('UserAgent', 'user-agent', 'cs(User-Agent)'))
        ClientIPAddress = (Get-ImtTrackingFirstAvailablePropertyValue -InputObject $parsedRow -PropertyNames @('ClientIpAddress', 'ClientIPAddress', 'client-ip-address', 'c-ip'))
        AuthenticatedUser = (Get-ImtTrackingFirstAvailablePropertyValue -InputObject $parsedRow -PropertyNames @('AuthenticatedUser', 'authenticated-user', 'UserEmail', 'user-email', 'cs-username'))
        AnchorMailbox = (Get-ImtTrackingFirstAvailablePropertyValue -InputObject $parsedRow -PropertyNames @('AnchorMailbox', 'anchor-mailbox'))
        RemoteEndpoint = (Get-ImtTrackingFirstAvailablePropertyValue -InputObject $parsedRow -PropertyNames @('remote-endpoint', 'RemoteEndpoint'))
        SessionId = (Get-ImtTrackingFirstAvailablePropertyValue -InputObject $parsedRow -PropertyNames @('session-id', 'SessionId'))
        Event = (Get-ImtTrackingFirstAvailablePropertyValue -InputObject $parsedRow -PropertyNames @('event', 'Event'))
        Data = (Get-ImtTrackingFirstAvailablePropertyValue -InputObject $parsedRow -PropertyNames @('data', 'Data', 'GenericInfo', 'generic-info'))
      })
    }
  }

  @($rows.ToArray())
}

function Get-ImtProtocolEvidenceRowsForSenders {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory = $true)]
    $RunContext,

    [Parameter(Mandatory = $true)]
    [string[]]$Senders,

    [AllowEmptyCollection()]
    [string[]]$Servers,

    [hashtable]$ResolvedMailboxBySender
  )

  $startDate = $RunContext.Start.AddMinutes(-15)
  $endDate = $RunContext.End.AddMinutes(15)
  $hourTokens = Get-ImtProtocolWindowTokens -StartDate $startDate -EndDate $endDate -Granularity Hour
  $dayTokens = Get-ImtProtocolWindowTokens -StartDate $startDate -EndDate $endDate -Granularity Day

  $serverList = @(
    $Servers |
      ForEach-Object { $_ -as [string] } |
      Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
      Select-Object -Unique
  )

  if ($serverList.Count -eq 0) {
    $serverList = @($null)
  }

  $logSpecs = @(
    @{ EvidenceType = 'HttpProxyMapi'; RelativePath = 'Logging\HttpProxy\Mapi'; Tokens = $hourTokens }
    @{ EvidenceType = 'HttpProxyOwa'; RelativePath = 'Logging\HttpProxy\Owa'; Tokens = $hourTokens }
    @{ EvidenceType = 'HttpProxyEas'; RelativePath = 'Logging\HttpProxy\Eas'; Tokens = $hourTokens }
    @{ EvidenceType = 'HttpProxyEws'; RelativePath = 'Logging\HttpProxy\Ews'; Tokens = $hourTokens }
    @{ EvidenceType = 'SmtpReceiveFrontEnd'; RelativePath = 'TransportRoles\Logs\FrontEnd\ProtocolLog\SmtpReceive'; Tokens = $dayTokens }
    @{ EvidenceType = 'SmtpReceiveHub'; RelativePath = 'TransportRoles\Logs\Hub\ProtocolLog\SmtpReceive'; Tokens = $dayTokens }
  )

  $rows = New-Object System.Collections.Generic.List[object]
  $failures = New-Object System.Collections.Generic.List[string]
  $candidateFileCache = @{}

  foreach ($sender in @($Senders | Sort-Object -Unique)) {
    $resolvedMailbox = $null
    if ($ResolvedMailboxBySender -and $ResolvedMailboxBySender.ContainsKey($sender)) {
      $resolvedMailbox = $ResolvedMailboxBySender[$sender]
    }

    $patterns = Get-ImtProtocolSearchPatterns -Sender $sender -ResolvedMailbox $resolvedMailbox

    foreach ($server in $serverList) {
      foreach ($spec in $logSpecs) {
        try {
          $cacheKey = '{0}|{1}|{2}' -f ($server -as [string]), $spec.EvidenceType, ($spec.Tokens -join ',')
          if (-not $candidateFileCache.ContainsKey($cacheKey)) {
            $directoryPath = Get-ImtProtocolLogDirectoryPath -Server $server -RelativePath $spec.RelativePath
            $candidateFileCache[$cacheKey] = @(Get-ImtProtocolCandidateFiles -DirectoryPath $directoryPath -Tokens $spec.Tokens)
          }

          $candidateFiles = @($candidateFileCache[$cacheKey])
          if ($candidateFiles.Count -eq 0) {
            continue
          }

          foreach ($protocolRow in @(Get-ImtProtocolRowsFromFiles -Paths $candidateFiles -Patterns $patterns -Mailbox $sender -EvidenceType $spec.EvidenceType -Server $server -StartDate $startDate -EndDate $endDate)) {
            [void]$rows.Add($protocolRow)
          }
        } catch {
          [void]$failures.Add(('{0}|{1}: {2}' -f ($server -as [string]), $spec.EvidenceType, $_.Exception.Message))
        }
      }
    }
  }

  $dedupedRows = @(
    $rows.ToArray() |
      Group-Object {
        '{0}|{1}|{2}|{3}|{4}' -f `
          (Get-ImtTrackingPropertyValue -InputObject $_ -PropertyName 'Mailbox'), `
          (Get-ImtTrackingPropertyValue -InputObject $_ -PropertyName 'EvidenceType'), `
          ((Get-ImtTrackingPropertyValue -InputObject $_ -PropertyName 'Timestamp') -as [datetime]), `
          (Get-ImtTrackingPropertyValue -InputObject $_ -PropertyName 'LogPath'), `
          (Get-ImtTrackingPropertyValue -InputObject $_ -PropertyName 'SessionId')
      } |
      ForEach-Object { $_.Group[0] }
  )

  [pscustomobject]@{
    Rows = @($dedupedRows)
    Failures = @($failures.ToArray() | Select-Object -Unique)
  }
}

function Find-ImtTrackingBestProtocolLogMatch {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory = $true)]
    [pscustomobject]$Message,

    [AllowEmptyCollection()]
    [object[]]$ProtocolRows,

    [ValidateRange(1, 120)]
    [int]$WindowMinutes = 15
  )

  if (-not $ProtocolRows -or @($ProtocolRows).Count -eq 0) {
    return $null
  }

  $messageTimestamp = [datetime]$Message.SubmittedAt
  $messageSender = ($Message.Sender -as [string])
  $messageRecipients = @(
    $Message.Recipients |
      ForEach-Object { $_ -as [string] } |
      Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
  )

  $scoredRows = New-Object System.Collections.Generic.List[object]

  foreach ($row in @($ProtocolRows)) {
    $timestamp = Get-ImtTrackingPropertyValue -InputObject $row -PropertyName 'Timestamp'
    if (-not $timestamp) {
      continue
    }

    $rowTimestamp = [datetime]$timestamp
    $deltaMinutes = [math]::Abs((New-TimeSpan -Start $messageTimestamp -End $rowTimestamp).TotalMinutes)
    if ($deltaMinutes -gt $WindowMinutes) {
      continue
    }

    $evidenceType = (Get-ImtTrackingPropertyValue -InputObject $row -PropertyName 'EvidenceType') -as [string]
    $userAgent = (Get-ImtTrackingPropertyValue -InputObject $row -PropertyName 'UserAgent') -as [string]
    $urlStem = (Get-ImtTrackingPropertyValue -InputObject $row -PropertyName 'UrlStem') -as [string]
    $evidenceBlob = Join-ImtTrackingDistinctValues -Values @(
      (Get-ImtTrackingPropertyValue -InputObject $row -PropertyName 'AuthenticatedUser'),
      (Get-ImtTrackingPropertyValue -InputObject $row -PropertyName 'AnchorMailbox'),
      (Get-ImtTrackingPropertyValue -InputObject $row -PropertyName 'Data'),
      $urlStem,
      $userAgent
    )

    $recipientHits = 0
    foreach ($recipient in $messageRecipients) {
      if ($evidenceBlob.IndexOf($recipient, [System.StringComparison]::OrdinalIgnoreCase) -ge 0) {
        $recipientHits++
      }
    }

    $score = switch -Regex ($evidenceType) {
      '^SmtpReceive' { 90; break }
      '^HttpProxyMapi$' { 70; break }
      '^HttpProxyOwa$' { 65; break }
      '^HttpProxyEas$' { 60; break }
      '^HttpProxyEws$' { 55; break }
      default { 40 }
    }

    if (-not [string]::IsNullOrWhiteSpace($messageSender) -and $evidenceBlob.IndexOf($messageSender, [System.StringComparison]::OrdinalIgnoreCase) -ge 0) {
      $score += 20
    }

    if ($recipientHits -gt 0) {
      $score += (15 * $recipientHits)
    }

    if ($urlStem -match '(?i)/mapi/emsmdb/') {
      $score += 10
    }

    if ($userAgent -match '(?i)Outlook|MAPI') {
      $score += 10
    }

    $score += [math]::Max(0, (30 - [int][math]::Floor($deltaMinutes * 2)))

    [void]$scoredRows.Add([pscustomobject]@{
      Row = $row
      Score = $score
      DeltaMinutes = [math]::Round($deltaMinutes, 2)
      EvidenceType = $evidenceType
      RecipientHits = $recipientHits
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
  if ($bestMatch.EvidenceType -match '^SmtpReceive' -and $bestMatch.DeltaMinutes -le 2 -and $bestMatch.RecipientHits -gt 0) {
    $confidence = 'High'
  } elseif ($bestMatch.DeltaMinutes -le 5) {
    $confidence = 'Medium'
  }

  [pscustomobject]@{
    Row = $bestMatch.Row
    Score = $bestMatch.Score
    DeltaMinutes = $bestMatch.DeltaMinutes
    Confidence = $confidence
  }
}

function Format-ImtActiveSyncLikelyClient {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory = $true)]
    [object]$DeviceRow
  )

  $friendlyName = (Get-ImtTrackingPropertyValue -InputObject $DeviceRow -PropertyName 'DeviceFriendlyName') -as [string]
  $deviceModel = (Get-ImtTrackingPropertyValue -InputObject $DeviceRow -PropertyName 'DeviceModel') -as [string]
  $deviceType = (Get-ImtTrackingPropertyValue -InputObject $DeviceRow -PropertyName 'DeviceType') -as [string]
  $deviceOs = (Get-ImtTrackingPropertyValue -InputObject $DeviceRow -PropertyName 'DeviceOS') -as [string]

  $parts = New-Object System.Collections.Generic.List[string]
  foreach ($value in @($friendlyName, $deviceModel, $deviceType)) {
    $text = ($value -as [string])
    if (-not [string]::IsNullOrWhiteSpace($text)) {
      $normalized = $text.Trim()
      if (-not ($parts.Contains($normalized))) {
        [void]$parts.Add($normalized)
      }
    }
  }

  $label = if ($parts.Count -gt 0) {
    $parts -join ' / '
  } else {
    'Mobile client'
  }

  if (-not [string]::IsNullOrWhiteSpace($deviceOs)) {
    return ('{0} ({1}) via ActiveSync' -f $label, $deviceOs.Trim())
  }

  ('{0} via ActiveSync' -f $label)
}

function Find-ImtTrackingBestActiveSyncDeviceMatch {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory = $true)]
    [pscustomobject]$Message,

    [Parameter(Mandatory = $true)]
    [pscustomobject]$TrailHints,

    [AllowEmptyCollection()]
    [object[]]$DeviceRows
  )

  if (-not $DeviceRows -or @($DeviceRows).Count -eq 0) {
    return $null
  }

  $messageTimestamp = [datetime]$Message.SubmittedAt
  $trackingClientType = (Get-ImtTrackingPropertyValue -InputObject $TrailHints -PropertyName 'TrackingClientType') -as [string]
  $requiresActiveSync = -not [string]::IsNullOrWhiteSpace($trackingClientType) -and $trackingClientType -match '(?i)activesync|airsync'

  $scoredRows = New-Object System.Collections.Generic.List[object]
  $deviceRowSet = @($DeviceRows)

  foreach ($row in $deviceRowSet) {
    $candidateTimestamps = @()
    foreach ($propertyName in @('LastSyncAttemptTime', 'LastSuccessSync', 'LastPolicyUpdateTime', 'FirstSyncTime')) {
      $timestamp = ConvertTo-ImtProtocolComparableTimestamp -Value (Get-ImtTrackingPropertyValue -InputObject $row -PropertyName $propertyName)
      if ($timestamp) {
        $candidateTimestamps += [pscustomobject]@{
          PropertyName = $propertyName
          Timestamp = $timestamp
          DeltaMinutes = [math]::Abs((New-TimeSpan -Start $messageTimestamp -End $timestamp).TotalMinutes)
        }
      }
    }

    $nearestTimestamp = $candidateTimestamps |
      Sort-Object `
        @{ Expression = 'DeltaMinutes'; Descending = $false }, `
        @{ Expression = 'PropertyName'; Descending = $false } |
      Select-Object -First 1

    $deviceType = (Get-ImtTrackingPropertyValue -InputObject $row -PropertyName 'DeviceType') -as [string]
    $deviceModel = (Get-ImtTrackingPropertyValue -InputObject $row -PropertyName 'DeviceModel') -as [string]
    $deviceOs = (Get-ImtTrackingPropertyValue -InputObject $row -PropertyName 'DeviceOS') -as [string]
    $friendlyName = (Get-ImtTrackingPropertyValue -InputObject $row -PropertyName 'DeviceFriendlyName') -as [string]
    $deviceUserAgent = (Get-ImtTrackingPropertyValue -InputObject $row -PropertyName 'DeviceUserAgent') -as [string]
    $clientType = (Get-ImtTrackingPropertyValue -InputObject $row -PropertyName 'ClientType') -as [string]
    $deviceAccessState = (Get-ImtTrackingPropertyValue -InputObject $row -PropertyName 'DeviceAccessState') -as [string]

    $score = 0
    if ($requiresActiveSync) {
      $score += 50
    }

    if ($clientType -match '(?i)activesync|airsync|eas') {
      $score += 30
    }

    if (-not [string]::IsNullOrWhiteSpace($deviceType) -or -not [string]::IsNullOrWhiteSpace($deviceModel) -or -not [string]::IsNullOrWhiteSpace($friendlyName)) {
      $score += 15
    }

    if (-not [string]::IsNullOrWhiteSpace($deviceOs) -or -not [string]::IsNullOrWhiteSpace($deviceUserAgent)) {
      $score += 10
    }

    if ($deviceAccessState -match '(?i)^allowed$') {
      $score += 10
    }

    if ($nearestTimestamp) {
      if ($nearestTimestamp.DeltaMinutes -le 15) {
        $score += 40
      } elseif ($nearestTimestamp.DeltaMinutes -le 60) {
        $score += 30
      } elseif ($nearestTimestamp.DeltaMinutes -le 720) {
        $score += 20
      } elseif ($nearestTimestamp.DeltaMinutes -le 1440) {
        $score += 10
      }
    } elseif ($deviceRowSet.Count -eq 1) {
      $score += 10
    }

    [void]$scoredRows.Add([pscustomobject]@{
      Row = $row
      Score = $score
      DeltaMinutes = if ($nearestTimestamp) { [math]::Round($nearestTimestamp.DeltaMinutes, 2) } else { $null }
      TimeProperty = if ($nearestTimestamp) { $nearestTimestamp.PropertyName } else { $null }
      HasRichDeviceData = (-not [string]::IsNullOrWhiteSpace($deviceType)) -or (-not [string]::IsNullOrWhiteSpace($deviceModel)) -or (-not [string]::IsNullOrWhiteSpace($friendlyName))
      DeviceRowCount = $deviceRowSet.Count
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
  if ($bestMatch.DeltaMinutes -ne $null -and $bestMatch.DeltaMinutes -le 15 -and $bestMatch.HasRichDeviceData) {
    $confidence = 'High'
  } elseif (($bestMatch.DeltaMinutes -ne $null -and $bestMatch.DeltaMinutes -le 1440) -or ($bestMatch.DeviceRowCount -eq 1 -and $bestMatch.HasRichDeviceData)) {
    $confidence = 'Medium'
  }

  [pscustomobject]@{
    Row = $bestMatch.Row
    Score = $bestMatch.Score
    DeltaMinutes = $bestMatch.DeltaMinutes
    TimeProperty = $bestMatch.TimeProperty
    Confidence = $confidence
  }
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
    [string[]]$CandidateMailboxAddresses,

    [AllowEmptyCollection()]
    [string[]]$Servers
  )

  if (@($Results).Count -eq 0) {
    return New-ImtModuleResult -StepName 'MessageClientAccess' -Status 'SKIP' -Summary 'Client access correlation skipped because there were no tracking results.' -Data ([pscustomobject]@{
      Rows = @()
      AuditRows = @()
      ProtocolRows = @()
      ActiveSyncRows = @()
      AuditAvailable = $false
      AuditFailures = @()
      ProtocolFailures = @()
      ActiveSyncFailures = @()
    }) -Metrics @{
      MessageRows = 0
      MailboxAuditRows = 0
      ProtocolRows = 0
      ActiveSyncRows = 0
      AuditFailures = 0
      ProtocolFailures = 0
      ActiveSyncFailures = 0
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

  $senderTrailHintsBySender = @{}
  foreach ($senderKey in @($senderSet.Keys | Sort-Object)) {
    $senderRows = @(
      $Results |
        Where-Object {
          $rowSender = (Get-ImtTrackingPropertyValue -InputObject $_ -PropertyName 'Sender') -as [string]
          $rowSender -and $rowSender.Trim().ToLowerInvariant() -eq $senderKey
        }
    )

    if ($senderRows.Count -gt 0) {
      $senderTrailHintsBySender[$senderKey] = Get-ImtTrackingTrailHints -Rows $senderRows
    }
  }

  $resolvedMailboxBySender = @{}
  foreach ($sender in @($senderSet.Keys | Sort-Object)) {
    $shouldAttemptMailboxLookup = ($candidateMailboxSet.Count -eq 0) -or $candidateMailboxSet.ContainsKey($sender)
    if (-not $shouldAttemptMailboxLookup) {
      continue
    }

    $resolvedMailbox = Resolve-ImtMailboxByAddress -Address $sender
    if ($resolvedMailbox) {
      $resolvedMailboxBySender[$sender] = $resolvedMailbox
    } else {
      Write-ImtLog -Level DEBUG -Step 'MessageClientAccess' -EventType Progress -Message ("Unable to resolve sender '{0}' to a mailbox for mailbox/device correlation." -f $sender)
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

    $resolvedMailbox = if ($resolvedMailboxBySender.ContainsKey($sender)) { $resolvedMailboxBySender[$sender] } else { $null }
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

  $protocolEvidenceResult = Get-ImtProtocolEvidenceRowsForSenders -RunContext $RunContext -Senders @($senderSet.Keys) -Servers $Servers -ResolvedMailboxBySender $resolvedMailboxBySender
  $protocolRows = @($protocolEvidenceResult.Rows)
  $protocolFailures = @($protocolEvidenceResult.Failures)

  foreach ($failure in @($protocolFailures)) {
    Write-ImtLog -Level WARN -Step 'MessageClientAccess' -EventType Progress -Message ("Protocol log evidence query warning: {0}" -f $failure)
  }

  $activeSyncRows = @()
  $activeSyncFailures = New-Object System.Collections.Generic.List[string]
  $activeSyncRowsBySender = @{}

  foreach ($sender in @($senderSet.Keys | Sort-Object)) {
    $trailHints = if ($senderTrailHintsBySender.ContainsKey($sender)) { $senderTrailHintsBySender[$sender] } else { $null }
    $trackingClientType = if ($trailHints) { (Get-ImtTrackingPropertyValue -InputObject $trailHints -PropertyName 'TrackingClientType') -as [string] } else { $null }
    $isActiveSyncCandidate = -not [string]::IsNullOrWhiteSpace($trackingClientType) -and $trackingClientType -match '(?i)activesync|airsync'
    if (-not $isActiveSyncCandidate) {
      continue
    }

    $mailboxLookupIdentity = $sender
    if ($resolvedMailboxBySender.ContainsKey($sender)) {
      $resolvedMailbox = $resolvedMailboxBySender[$sender]
      $mailboxLookupIdentity = (Get-ImtTrackingFirstAvailablePropertyValue -InputObject $resolvedMailbox -PropertyNames @('PrimarySmtpAddress', 'Identity', 'Alias')) -as [string]
      if ([string]::IsNullOrWhiteSpace($mailboxLookupIdentity)) {
        $mailboxLookupIdentity = $sender
      }
    }

    try {
      $activeSyncEvidenceResult = Get-ImtActiveSyncDeviceEvidence -MailboxIdentity $mailboxLookupIdentity
      $senderActiveSyncRows = @($activeSyncEvidenceResult.Rows)
      $activeSyncRowsBySender[$sender] = $senderActiveSyncRows
      foreach ($row in $senderActiveSyncRows) {
        $activeSyncRows += $row
      }
      foreach ($failure in @($activeSyncEvidenceResult.Failures)) {
        [void]$activeSyncFailures.Add(('{0}: {1}' -f $sender, $failure))
      }
    } catch {
      [void]$activeSyncFailures.Add(('{0}: {1}' -f $sender, $_.Exception.Message))
    }
  }

  foreach ($failure in @($activeSyncFailures)) {
    Write-ImtLog -Level WARN -Step 'MessageClientAccess' -EventType Progress -Message ("ActiveSync device correlation warning: {0}" -f $failure)
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

    $senderProtocolRows = @(
      $protocolRows |
        Where-Object {
          $rowMailbox = (Get-ImtTrackingPropertyValue -InputObject $_ -PropertyName 'Mailbox') -as [string]
          $rowMailbox -and $normalizedSender -and $rowMailbox.Trim().ToLowerInvariant() -eq $normalizedSender
        }
    )

    $protocolMatch = Find-ImtTrackingBestProtocolLogMatch -Message ([pscustomobject]@{
        SubmittedAt = $submittedAt
        Sender = $sender
        Recipients = @(
          $groupRows |
            ForEach-Object { @(Get-ImtTrackingPropertyValue -InputObject $_ -PropertyName 'Recipients') } |
            ForEach-Object { $_ -as [string] } |
            Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
            Select-Object -Unique
        )
      }) `
      -ProtocolRows $senderProtocolRows

    $senderActiveSyncRows = if ($normalizedSender -and $activeSyncRowsBySender.ContainsKey($normalizedSender)) {
      @($activeSyncRowsBySender[$normalizedSender])
    } else {
      @()
    }

    $activeSyncMatch = Find-ImtTrackingBestActiveSyncDeviceMatch -Message ([pscustomobject]@{
        SubmittedAt = $submittedAt
      }) `
      -TrailHints $trailHints `
      -DeviceRows $senderActiveSyncRows

    $deviceAssessment = Resolve-ImtTrackingDeviceAssessment -AuditMatch $auditMatch -ProtocolMatch $protocolMatch -ActiveSyncMatch $activeSyncMatch -TrailHints $trailHints

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
      ProtocolEvidenceType = $deviceAssessment.ProtocolEvidenceType
      ProtocolLogServer = $deviceAssessment.ProtocolLogServer
      ProtocolLogPath = $deviceAssessment.ProtocolLogPath
      ProtocolUserAgent = $deviceAssessment.ProtocolUserAgent
      ProtocolRemoteEndpoint = $deviceAssessment.ProtocolRemoteEndpoint
      ProtocolTimestamp = $deviceAssessment.ProtocolTimestamp
      ProtocolDeltaMinutes = $deviceAssessment.ProtocolDeltaMinutes
      ActiveSyncDeviceId = Get-ImtTrackingPropertyValue -InputObject $deviceAssessment -PropertyName 'ActiveSyncDeviceId'
      ActiveSyncDeviceType = Get-ImtTrackingPropertyValue -InputObject $deviceAssessment -PropertyName 'ActiveSyncDeviceType'
      ActiveSyncDeviceModel = Get-ImtTrackingPropertyValue -InputObject $deviceAssessment -PropertyName 'ActiveSyncDeviceModel'
      ActiveSyncDeviceOS = Get-ImtTrackingPropertyValue -InputObject $deviceAssessment -PropertyName 'ActiveSyncDeviceOS'
      ActiveSyncDeviceFriendlyName = Get-ImtTrackingPropertyValue -InputObject $deviceAssessment -PropertyName 'ActiveSyncDeviceFriendlyName'
      ActiveSyncDeviceUserAgent = Get-ImtTrackingPropertyValue -InputObject $deviceAssessment -PropertyName 'ActiveSyncDeviceUserAgent'
      ActiveSyncClientType = Get-ImtTrackingPropertyValue -InputObject $deviceAssessment -PropertyName 'ActiveSyncClientType'
      ActiveSyncLastSuccessSync = Get-ImtTrackingPropertyValue -InputObject $deviceAssessment -PropertyName 'ActiveSyncLastSuccessSync'
      ActiveSyncLastSyncAttemptTime = Get-ImtTrackingPropertyValue -InputObject $deviceAssessment -PropertyName 'ActiveSyncLastSyncAttemptTime'
      ActiveSyncDeltaMinutes = Get-ImtTrackingPropertyValue -InputObject $deviceAssessment -PropertyName 'ActiveSyncDeltaMinutes'
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

  $status = if (-not $auditAvailable -or $auditFailures.Count -gt 0 -or $protocolFailures.Count -gt 0 -or $activeSyncFailures.Count -gt 0 -or $confidenceCounts.None -gt 0) { 'WARN' } else { 'OK' }
  $summary = "Client attribution rows={0}; High={1}; Medium={2}; Low={3}; None={4}; MailboxAuditRows={5}; ProtocolRows={6}; ActiveSyncRows={7}; AuditFailures={8}; ProtocolFailures={9}; ActiveSyncFailures={10}" -f `
    $rowArray.Count, `
    $confidenceCounts.High, `
    $confidenceCounts.Medium, `
    $confidenceCounts.Low, `
    $confidenceCounts.None, `
    $auditRows.Count, `
    $protocolRows.Count, `
    @($activeSyncRows).Count, `
    $auditFailures.Count, `
    $protocolFailures.Count, `
    $activeSyncFailures.Count

  New-ImtModuleResult -StepName 'MessageClientAccess' -Status $status -Summary $summary -Data ([pscustomobject]@{
    Rows = $rowArray
    AuditRows = @($auditRows.ToArray())
    ProtocolRows = @($protocolRows)
    ActiveSyncRows = @($activeSyncRows)
    AuditAvailable = [bool]$auditAvailable
    AuditFailures = @($auditFailures.ToArray())
    ProtocolFailures = @($protocolFailures)
    ActiveSyncFailures = @($activeSyncFailures)
  }) -Metrics @{
    MessageRows = $rowArray.Count
    MailboxAuditRows = $auditRows.Count
    ProtocolRows = $protocolRows.Count
    ActiveSyncRows = @($activeSyncRows).Count
    HighConfidence = $confidenceCounts.High
    MediumConfidence = $confidenceCounts.Medium
    LowConfidence = $confidenceCounts.Low
    NoConfidence = $confidenceCounts.None
    AuditFailures = $auditFailures.Count
    ProtocolFailures = $protocolFailures.Count
    ActiveSyncFailures = $activeSyncFailures.Count
  } -Errors @(@($auditFailures) + @($protocolFailures) + @($activeSyncFailures))
}
