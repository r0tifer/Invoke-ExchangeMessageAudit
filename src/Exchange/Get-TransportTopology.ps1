Set-StrictMode -Version Latest

function Get-ImtTransportTopology {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory=$true)]$RunContext
  )

  $allTransport = @(Get-TransportService)
  if ($allTransport.Count -eq 0) {
    throw "No TransportService servers found. Run in the Exchange Management Shell."
  }

  $hasIsFrontendProp = $null -ne ($allTransport | Get-Member -Name IsFrontendTransportServer -MemberType NoteProperty)
  $transportTargets = if ($hasIsFrontendProp) {
    $allTransport | Where-Object { -not $_.IsFrontendTransportServer }
  } else {
    $allTransport
  }

  $servers = @($transportTargets | Select-Object -ExpandProperty Name)
  $versionInfo = @{}
  foreach ($server in $servers) {
    try {
      $exchangeServer = Get-ExchangeServer -Identity $server -ErrorAction Stop
      $major = [int]$exchangeServer.AdminDisplayVersion.Major
      $minor = [int]$exchangeServer.AdminDisplayVersion.Minor
      $versionLabel = switch ("$major.$minor") {
        '15.0' { 'Exchange 2013'; break }
        '15.1' { 'Exchange 2016'; break }
        '15.2' { 'Exchange 2019'; break }
        default { "Exchange $($exchangeServer.AdminDisplayVersion)" }
      }
      $versionInfo[$server] = $versionLabel
    } catch {
      $versionInfo[$server] = 'Unknown (RBAC/permissions?)'
    }
  }

  $detected = ($versionInfo.GetEnumerator() | Sort-Object Key | ForEach-Object { "{0}=>{1}" -f $_.Key, $_.Value }) -join '; '

  New-ImtModuleResult -StepName 'DiscoverTransport' -Status 'OK' -Summary ("Servers={0}; Versions={1}" -f $servers.Count, $detected) -Data ([pscustomobject]@{
    AllTransport = @($allTransport)
    TransportTargets = @($transportTargets)
    Servers = @($servers)
    VersionInfo = $versionInfo
  }) -Metrics @{
    TransportServers = $allTransport.Count
    TraceServers = $servers.Count
  } -Errors @()
}
