Set-StrictMode -Version Latest

function New-ImtModuleResult {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory=$true)][string]$StepName,
    [Parameter(Mandatory=$true)][ValidateSet('OK','WARN','FAIL','SKIP')][string]$Status,
    [Parameter(Mandatory=$true)][string]$Summary,
    [object]$Data,
    [hashtable]$Metrics,
    [string[]]$Errors
  )

  if (-not $Metrics) {
    $Metrics = @{}
  }
  if (-not $Errors) {
    $Errors = @()
  }

  [pscustomobject]@{
    StepName = $StepName
    Status = $Status
    Summary = $Summary
    Data = $Data
    Metrics = $Metrics
    Errors = @($Errors)
  }
}

function New-ImtRunSummary {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory=$true)][object[]]$StepResults,
    [datetime]$StartedAt,
    [datetime]$EndedAt
  )

  $counts = @{
    OK = 0
    WARN = 0
    FAIL = 0
    SKIP = 0
  }

  foreach ($step in $StepResults) {
    $status = ($step.Status -as [string])
    if ($counts.ContainsKey($status)) {
      $counts[$status]++
    }
  }

  $duration = $null
  if ($StartedAt -and $EndedAt) {
    $duration = [math]::Round((New-TimeSpan -Start $StartedAt -End $EndedAt).TotalSeconds, 2)
  }

  [pscustomobject]@{
    TotalSteps = $StepResults.Count
    Counts = [pscustomobject]$counts
    StartedAt = $StartedAt
    EndedAt = $EndedAt
    DurationSeconds = $duration
  }
}
