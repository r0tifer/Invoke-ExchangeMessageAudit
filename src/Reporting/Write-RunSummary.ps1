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

  $stepOutcomes = @(
    foreach ($step in @($StepResults)) {
      [pscustomobject]@{
        Step = $step.StepName
        Status = $step.Status
        Summary = $step.Summary
      }
    }
  )

  $finalKeywordByMailboxRows = @()
  $combinedStepResult = @(
    $StepResults |
      Where-Object {
        ($_.StepName -as [string]) -eq 'KeywordCombined' -and
        $_.Data -and
        $_.Data.PSObject.Properties.Name -contains 'CombinedByMailboxRows'
      }
  ) | Select-Object -Last 1

  if ($combinedStepResult -and $combinedStepResult.Data -and $combinedStepResult.Data.CombinedByMailboxRows) {
    $finalKeywordByMailboxRows = @(
      foreach ($row in @($combinedStepResult.Data.CombinedByMailboxRows | Sort-Object Mailbox,Keyword)) {
        [pscustomobject]@{
          Mailbox = $row.Mailbox
          Keyword = $row.Keyword
          TransportEventHitCount = [int]($row.TransportEventHitCount -as [int])
          TransportDistinctMessageIdHitCount = [int]($row.TransportDistinctMessageIdHitCount -as [int])
          MailboxEstimatedItemHitCount = [int]($row.MailboxEstimatedItemHitCount -as [int])
        }
      }
    )
  }

  $summaryObject | Add-Member -NotePropertyName StepOutcomes -NotePropertyValue @($stepOutcomes) -Force
  $summaryObject | Add-Member -NotePropertyName FinalKeywordByMailboxRows -NotePropertyValue @($finalKeywordByMailboxRows) -Force

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
