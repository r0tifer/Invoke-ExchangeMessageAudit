Set-StrictMode -Version Latest

function Get-ImtRetentionSnapshot {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory=$true)]$RunContext,
    [Parameter(Mandatory=$true)][object[]]$TransportTargets
  )

  $retentionRows = @()
  $oldestByServer = @{}

  foreach ($service in $TransportTargets) {
    $row = '' | Select-Object Name,MessageTrackingLogPath,MessageTrackingLogMaxAge,MessageTrackingLogMaxDirectorySize,MessageTrackingLogMaxFileSize,OldestLog,NewestLog
    $row.Name = $service.Name
    $row.MessageTrackingLogPath = $service.MessageTrackingLogPath
    $row.MessageTrackingLogMaxAge = $service.MessageTrackingLogMaxAge
    $row.MessageTrackingLogMaxDirectorySize = $service.MessageTrackingLogMaxDirectorySize
    $row.MessageTrackingLogMaxFileSize = $service.MessageTrackingLogMaxFileSize

    try {
      if ($service.MessageTrackingLogPath -and (Test-Path -LiteralPath $service.MessageTrackingLogPath)) {
        $files = Get-ChildItem -LiteralPath $service.MessageTrackingLogPath -Filter *.log -ErrorAction Stop
        if ($files) {
          $oldest = ($files | Sort-Object LastWriteTime | Select-Object -First 1).LastWriteTimeUtc
          $newest = ($files | Sort-Object LastWriteTime -Descending | Select-Object -First 1).LastWriteTimeUtc
          $row.OldestLog = $oldest
          $row.NewestLog = $newest
          $oldestByServer[$service.Name] = $oldest
        }
      }
    } catch {
      Write-ImtLog -Level DEBUG -Step 'RetentionSnapshot' -EventType Progress -Message ("Unable to inspect log path on {0}: {1}" -f $service.Name, $_.Exception.Message)
    }

    $retentionRows += $row
  }

  foreach ($server in $oldestByServer.Keys) {
    $oldestValue = $oldestByServer[$server]
    if ($oldestValue -and $RunContext.Start.ToUniversalTime() -lt $oldestValue) {
      Write-ImtLog -Level WARN -Step 'RetentionSnapshot' -EventType Progress -Message ("{0}: Oldest tracking log is {1:u}. Requested start {2:u} is earlier; older data is unavailable." -f $server, $oldestValue, $RunContext.Start.ToUniversalTime())
    }
  }

  New-ImtModuleResult -StepName 'RetentionSnapshot' -Status 'OK' -Summary ("Retention rows collected: {0}" -f $retentionRows.Count) -Data ([pscustomobject]@{
    RetentionRows = @($retentionRows)
    OldestByServer = $oldestByServer
  }) -Metrics @{
    RetentionRows = $retentionRows.Count
    ServersWithOldestLog = $oldestByServer.Count
  } -Errors @()
}

function Export-ImtRetentionSnapshot {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory=$true)]$RunContext,
    [Parameter(Mandatory=$true)][object[]]$RetentionRows
  )

  $retentionCsv = Join-Path $RunContext.OutputDir ("MTL_Retention_{0}.csv" -f $RunContext.Timestamp)
  $RetentionRows | Export-Csv -Path $retentionCsv -NoTypeInformation -Encoding UTF8

  New-ImtModuleResult -StepName 'RetentionExport' -Status 'OK' -Summary ("Retention snapshot exported: {0}" -f $retentionCsv) -Data ([pscustomobject]@{
    RetentionCsv = $retentionCsv
  }) -Metrics @{
    Rows = $RetentionRows.Count
  } -Errors @()
}
