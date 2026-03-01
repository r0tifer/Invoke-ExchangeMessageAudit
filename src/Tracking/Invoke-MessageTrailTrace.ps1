Set-StrictMode -Version Latest

function Invoke-ImtMessageTrailTrace {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory=$true)]$RunContext,
    [Parameter(Mandatory=$true)][object[]]$Results,
    [Parameter(Mandatory=$true)][string[]]$Servers,
    [Parameter(Mandatory=$true)][hashtable]$VersionInfo
  )

  $traceMessageId = $RunContext.Inputs.TraceMessageId
  if ($RunContext.Inputs.TraceLatest -and -not $traceMessageId) {
    $traceMessageId = ($Results | Sort-Object Timestamp | Select-Object -Last 1 -ExpandProperty MessageId)
    if (-not $traceMessageId) {
      return New-ImtModuleResult -StepName 'MessageTrailTrace' -Status 'WARN' -Summary 'TraceLatest requested but no results were available to select a MessageId.' -Data ([pscustomobject]@{
        TraceMessageId = $null
        TrailRows = @()
        TrailCsv = $null
      }) -Metrics @{} -Errors @()
    }
  }

  if (-not $traceMessageId) {
    return New-ImtModuleResult -StepName 'MessageTrailTrace' -Status 'SKIP' -Summary 'Message trail trace skipped (no TraceMessageId or TraceLatest).' -Data ([pscustomobject]@{
      TraceMessageId = $null
      TrailRows = @()
      TrailCsv = $null
    }) -Metrics @{} -Errors @()
  }

  $trail = New-Object System.Collections.Generic.List[object]
  $failures = 0

  foreach ($server in $Servers) {
    $version = $VersionInfo[$server]
    Write-ImtLog -Level DEBUG -Step 'MessageTrailTrace' -EventType Progress -Message ("Tracing on {0} ({1})" -f $server, $version)

    try {
      $chunk = Get-MessageTrackingLog -Server $server -Start $RunContext.Start -End $RunContext.End -MessageId $traceMessageId -ResultSize Unlimited
      if ($chunk) {
        foreach ($item in $chunk) {
          [void]$trail.Add($item)
        }
      }
    } catch {
      $failures++
      Write-ImtLog -Level WARN -Step 'MessageTrailTrace' -EventType Progress -Message ("[{0}] trace failed: {1}" -f $server, $_.Exception.Message)
    }
  }

  if ($trail.Count -eq 0) {
    $status = if ($failures -gt 0) { 'FAIL' } else { 'WARN' }
    $summary = if ($failures -gt 0) {
      "No trail events found for {0}; {1} server trace failure(s)." -f $traceMessageId, $failures
    } else {
      "No trail events found for {0} in the selected window." -f $traceMessageId
    }

    return New-ImtModuleResult -StepName 'MessageTrailTrace' -Status $status -Summary $summary -Data ([pscustomobject]@{
      TraceMessageId = $traceMessageId
      TrailRows = @()
      TrailCsv = $null
    }) -Metrics @{
      TrailRows = 0
      Failures = $failures
    } -Errors @()
  }

  $safeMid = ($traceMessageId -replace '[^\w@.<>\-@]','_')
  $trailCsv = Join-Path $RunContext.OutputDir ("MTL_Trail_{0}_{1}.csv" -f $safeMid, $RunContext.Timestamp)

  $trail |
    Sort-Object Timestamp |
    Select-Object Timestamp,EventId,Source,ServerHostname,ClientHostname,ConnectorId,
                  Sender,
                  @{n='Recipients';e={($_.Recipients -join ';')}},
                  MessageSubject,MessageId,InternalMessageId,RecipientStatus,TotalBytes,SourceContext |
    Export-Csv -Path $trailCsv -NoTypeInformation -Encoding UTF8

  New-ImtModuleResult -StepName 'MessageTrailTrace' -Status 'OK' -Summary ("Trail exported: {0}; Rows={1}" -f $trailCsv, $trail.Count) -Data ([pscustomobject]@{
    TraceMessageId = $traceMessageId
    TrailRows = $trail.ToArray()
    TrailCsv = $trailCsv
  }) -Metrics @{
    TrailRows = $trail.Count
    Failures = $failures
  } -Errors @()
}
