Set-StrictMode -Version Latest

function Get-ImtActiveSyncPropertyValue {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory = $true)]
    [object]$InputObject,

    [Parameter(Mandatory = $true)]
    [string[]]$PropertyNames
  )

  if ($null -eq $InputObject) {
    return $null
  }

  foreach ($propertyName in $PropertyNames) {
    $property = $InputObject.PSObject.Properties[$propertyName]
    if ($null -ne $property -and $null -ne $property.Value) {
      return $property.Value
    }
  }

  $null
}

function Get-ImtActiveSyncEvidenceKeys {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory = $true)]
    [object]$InputObject
  )

  $keys = New-Object System.Collections.Generic.List[string]

  foreach ($value in @(
      Get-ImtActiveSyncPropertyValue -InputObject $InputObject -PropertyNames @('Identity')
      Get-ImtActiveSyncPropertyValue -InputObject $InputObject -PropertyNames @('DeviceID', 'DeviceId')
      Get-ImtActiveSyncPropertyValue -InputObject $InputObject -PropertyNames @('Guid')
    )) {
    $text = ($value -as [string])
    if (-not [string]::IsNullOrWhiteSpace($text)) {
      [void]$keys.Add($text.Trim().ToLowerInvariant())
    }
  }

  @($keys.ToArray() | Select-Object -Unique)
}

function New-ImtActiveSyncEvidenceRow {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory = $true)]
    [string]$Mailbox,

    $StatsRow,

    $DeviceRow
  )

  $primaryStatsOrDevice = if ($null -ne $StatsRow) { $StatsRow } else { $DeviceRow }
  $primaryDeviceOrStats = if ($null -ne $DeviceRow) { $DeviceRow } else { $StatsRow }

  [pscustomobject]@{
    Mailbox = $Mailbox
    Identity = (Get-ImtActiveSyncPropertyValue -InputObject $primaryStatsOrDevice -PropertyNames @('Identity')) -as [string]
    DeviceId = (Get-ImtActiveSyncPropertyValue -InputObject $primaryStatsOrDevice -PropertyNames @('DeviceID', 'DeviceId')) -as [string]
    DeviceType = (Get-ImtActiveSyncPropertyValue -InputObject $primaryStatsOrDevice -PropertyNames @('DeviceType')) -as [string]
    DeviceModel = (Get-ImtActiveSyncPropertyValue -InputObject $primaryStatsOrDevice -PropertyNames @('DeviceModel', 'Model')) -as [string]
    DeviceOS = (Get-ImtActiveSyncPropertyValue -InputObject $primaryStatsOrDevice -PropertyNames @('DeviceOS', 'OS')) -as [string]
    DeviceFriendlyName = (Get-ImtActiveSyncPropertyValue -InputObject $primaryDeviceOrStats -PropertyNames @('FriendlyName', 'DeviceFriendlyName')) -as [string]
    DeviceUserAgent = (Get-ImtActiveSyncPropertyValue -InputObject $primaryDeviceOrStats -PropertyNames @('DeviceUserAgent', 'UserAgent')) -as [string]
    ClientType = (Get-ImtActiveSyncPropertyValue -InputObject $primaryDeviceOrStats -PropertyNames @('ClientType')) -as [string]
    DeviceAccessState = (Get-ImtActiveSyncPropertyValue -InputObject $primaryDeviceOrStats -PropertyNames @('DeviceAccessState')) -as [string]
    DeviceAccessStateReason = (Get-ImtActiveSyncPropertyValue -InputObject $primaryDeviceOrStats -PropertyNames @('DeviceAccessStateReason')) -as [string]
    Status = (Get-ImtActiveSyncPropertyValue -InputObject $primaryStatsOrDevice -PropertyNames @('Status')) -as [string]
    Guid = (Get-ImtActiveSyncPropertyValue -InputObject $primaryStatsOrDevice -PropertyNames @('Guid')) -as [string]
    FirstSyncTime = Get-ImtActiveSyncPropertyValue -InputObject $primaryStatsOrDevice -PropertyNames @('FirstSyncTime')
    LastSuccessSync = Get-ImtActiveSyncPropertyValue -InputObject $primaryStatsOrDevice -PropertyNames @('LastSuccessSync')
    LastSyncAttemptTime = Get-ImtActiveSyncPropertyValue -InputObject $primaryStatsOrDevice -PropertyNames @('LastSyncAttemptTime')
    LastPolicyUpdateTime = Get-ImtActiveSyncPropertyValue -InputObject $primaryStatsOrDevice -PropertyNames @('LastPolicyUpdateTime')
    LastDeviceWipeRequestTime = Get-ImtActiveSyncPropertyValue -InputObject $primaryStatsOrDevice -PropertyNames @('LastDeviceWipeRequestTime')
  }
}

function Get-ImtActiveSyncDeviceEvidence {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory = $true)]
    [string]$MailboxIdentity
  )

  $mobileDeviceStatsCommand = Get-Command -Name Get-MobileDeviceStatistics -ErrorAction SilentlyContinue
  $mobileDeviceCommand = Get-Command -Name Get-MobileDevice -ErrorAction SilentlyContinue

  if (-not $mobileDeviceStatsCommand -and -not $mobileDeviceCommand) {
    return [pscustomobject]@{
      Available = $false
      Rows = @()
      Failures = @()
    }
  }

  $failures = New-Object System.Collections.Generic.List[string]
  $statsRows = @()
  $deviceRows = @()

  if ($mobileDeviceStatsCommand) {
    try {
      $statsRows = @(Get-MobileDeviceStatistics -Mailbox $MailboxIdentity -ErrorAction Stop)
    } catch {
      [void]$failures.Add(('Get-MobileDeviceStatistics: {0}' -f $_.Exception.Message))
    }
  }

  if ($mobileDeviceCommand) {
    try {
      $deviceRows = @(Get-MobileDevice -Mailbox $MailboxIdentity -ErrorAction Stop)
    } catch {
      [void]$failures.Add(('Get-MobileDevice: {0}' -f $_.Exception.Message))
    }
  }

  $deviceByKey = @{}
  foreach ($deviceRow in @($deviceRows)) {
    foreach ($key in @(Get-ImtActiveSyncEvidenceKeys -InputObject $deviceRow)) {
      if (-not $deviceByKey.ContainsKey($key)) {
        $deviceByKey[$key] = $deviceRow
      }
    }
  }

  $rows = New-Object System.Collections.Generic.List[object]
  $matchedKeys = New-Object System.Collections.Generic.HashSet[string]

  foreach ($statsRow in @($statsRows)) {
    $matchedDeviceRow = $null
    foreach ($key in @(Get-ImtActiveSyncEvidenceKeys -InputObject $statsRow)) {
      if ($deviceByKey.ContainsKey($key)) {
        $matchedDeviceRow = $deviceByKey[$key]
        [void]$matchedKeys.Add($key)
        break
      }
    }

    [void]$rows.Add((New-ImtActiveSyncEvidenceRow -Mailbox $MailboxIdentity -StatsRow $statsRow -DeviceRow $matchedDeviceRow))
  }

  foreach ($deviceRow in @($deviceRows)) {
    $deviceKeys = @(Get-ImtActiveSyncEvidenceKeys -InputObject $deviceRow)
    $alreadyRepresented = $false
    foreach ($key in $deviceKeys) {
      if ($matchedKeys.Contains($key)) {
        $alreadyRepresented = $true
        break
      }
    }

    if (-not $alreadyRepresented) {
      [void]$rows.Add((New-ImtActiveSyncEvidenceRow -Mailbox $MailboxIdentity -StatsRow $null -DeviceRow $deviceRow))
    }
  }

  [pscustomobject]@{
    Available = $true
    Rows = @($rows.ToArray())
    Failures = @($failures.ToArray() | Select-Object -Unique)
  }
}
