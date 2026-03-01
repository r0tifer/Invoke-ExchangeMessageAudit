Set-StrictMode -Version Latest

function Write-ImtRunSummary {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory=$true)]$RunContext,
    [Parameter(Mandatory=$true)][object[]]$StepResults,
    [Parameter(Mandatory=$true)][datetime]$StartedAt,
    [Parameter(Mandatory=$true)][datetime]$EndedAt
  )

  $summaryObject = New-ImtRunSummary -StepResults $StepResults -StartedAt $StartedAt -EndedAt $EndedAt

  $line = "TotalSteps={0}; OK={1}; WARN={2}; FAIL={3}; SKIP={4}; DurationSeconds={5}" -f `
    $summaryObject.TotalSteps, `
    $summaryObject.Counts.OK, `
    $summaryObject.Counts.WARN, `
    $summaryObject.Counts.FAIL, `
    $summaryObject.Counts.SKIP, `
    ($summaryObject.DurationSeconds -as [string])

  $level = if ($summaryObject.Counts.FAIL -gt 0) { 'ERROR' } else { 'INFO' }
  Write-ImtLog -Level $level -Step 'RunSummary' -EventType Summary -Message $line

  New-ImtModuleResult -StepName 'RunSummary' -Status 'OK' -Summary $line -Data $summaryObject -Metrics @{
    TotalSteps = $summaryObject.TotalSteps
    OK = $summaryObject.Counts.OK
    WARN = $summaryObject.Counts.WARN
    FAIL = $summaryObject.Counts.FAIL
    SKIP = $summaryObject.Counts.SKIP
  } -Errors @()
}
