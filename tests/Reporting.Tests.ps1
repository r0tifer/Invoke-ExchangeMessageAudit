Set-StrictMode -Version Latest

$repoRoot = Split-Path -Path $PSScriptRoot -Parent
. (Join-Path $repoRoot 'src\\Models\\New-ResultObjects.ps1')
. (Join-Path $repoRoot 'src\\Logging\\Write-ImtLog.ps1')
. (Join-Path $repoRoot 'src\\Reporting\\Write-StepTables.ps1')
. (Join-Path $repoRoot 'src\\Reporting\\Write-RunSummary.ps1')

Describe 'Write-ImtStepDataTables' {
  BeforeEach {
    $tempDir = Join-Path $env:TEMP ("imt-reporting-tests-{0}" -f ([guid]::NewGuid().ToString('N')))
    New-Item -ItemType Directory -Path $tempDir | Out-Null

    $script:StepLogPath = Join-Path $tempDir 'steps.log'
    $script:TranscriptPath = Join-Path $tempDir 'run.log'
    Initialize-ImtLogger -OutputLevel INFO -DisableTranscriptLog -StepLogPath $script:StepLogPath -TranscriptPath $script:TranscriptPath
  }

  AfterEach {
    Complete-ImtLogger
    if ($script:StepLogPath) {
      $dir = Split-Path -Path $script:StepLogPath -Parent
      Remove-Item -LiteralPath $dir -Recurse -Force -ErrorAction SilentlyContinue
    }
  }

  It 'does not throw for step results with empty metrics and no data' {
    $result = New-ImtModuleResult -StepName 'ValidateInputs' -Status OK -Summary 'Input validation passed.' -Data $null -Metrics @{} -Errors @()

    { Write-ImtStepDataTables -StepResult $result } | Should Not Throw
  }

  It 'does not throw for step results with exactly one metric' {
    $result = New-ImtModuleResult -StepName 'ValidateInputs' -Status OK -Summary 'Input validation passed.' -Data $null -Metrics @{ Checked = 1 } -Errors @()

    { Write-ImtStepDataTables -StepResult $result } | Should Not Throw
  }

  It 'prints mailbox separators for tracking mailbox keyword rows' {
    Mock Write-Host {}

    $result = New-ImtModuleResult -StepName 'TrackingReport' -Status OK -Summary 'Tracking report exported.' -Data ([pscustomobject]@{
      TrackingKeywordRows = @()
      TrackingKeywordMailboxRows = @(
        [pscustomobject]@{
          Mailbox = 'alex.rivera@example.org'
          Keyword = 'audit'
          EventHitCount = 3
          DistinctMessageIdHitCount = 2
        },
        [pscustomobject]@{
          Mailbox = 'jamie.chen@example.org'
          Keyword = 'audit'
          EventHitCount = 1
          DistinctMessageIdHitCount = 1
        }
      )
      DailyCounts = @()
    }) -Metrics @{} -Errors @()

    { Write-ImtStepDataTables -StepResult $result } | Should Not Throw

    Assert-MockCalled Write-Host -Times 2 -Exactly -ParameterFilter {
      ($Object -as [string]) -match '\[TrackingReport\] \[Mailbox\] Mailbox: '
    }
  }

  It 'prints mailbox separators for run summary final keyword findings' {
    Mock Write-Host {}

    $summaryData = [pscustomobject]@{
      TotalSteps = 3
      Counts = [pscustomobject]@{
        OK = 3
        WARN = 0
        FAIL = 0
        SKIP = 0
      }
      DurationSeconds = 12.34
      StepOutcomes = @(
        [pscustomobject]@{
          Step = 'ValidateInputs'
          Status = 'OK'
          Summary = 'Input validation passed.'
        }
      )
      FinalKeywordByMailboxRows = @(
        [pscustomobject]@{
          Mailbox = 'alex.rivera@example.org'
          Keyword = 'audit'
          TransportEventHitCount = 3
          TransportDistinctMessageIdHitCount = 2
          MailboxEstimatedItemHitCount = 5
        },
        [pscustomobject]@{
          Mailbox = 'jamie.chen@example.org'
          Keyword = 'audit'
          TransportEventHitCount = 1
          TransportDistinctMessageIdHitCount = 1
          MailboxEstimatedItemHitCount = 2
        }
      )
    }

    $result = New-ImtModuleResult -StepName 'RunSummary' -Status OK -Summary 'run summary' -Data $summaryData -Metrics @{} -Errors @()

    { Write-ImtStepDataTables -StepResult $result } | Should Not Throw

    Assert-MockCalled Write-Host -Times 2 -Exactly -ParameterFilter {
      ($Object -as [string]) -match '\[RunSummary\] \[Mailbox\] Mailbox: '
    }
  }
}

Describe 'Write-ImtRunSummary' {
  It 'includes final keyword rows by mailbox when KeywordCombined output is present' {
    $startedAt = [datetime]'2026-02-01T00:00:00Z'
    $endedAt = [datetime]'2026-02-01T00:01:00Z'

    $stepResults = @(
      (New-ImtModuleResult -StepName 'ValidateInputs' -Status OK -Summary 'Input validation passed.' -Data $null -Metrics @{} -Errors @())
      (New-ImtModuleResult -StepName 'KeywordCombined' -Status OK -Summary 'Combined keyword reports exported.' -Data ([pscustomobject]@{
        CombinedByMailboxRows = @(
          [pscustomobject]@{
            Mailbox = 'alex.rivera@example.org'
            Keyword = 'audit'
            TransportEventHitCount = 3
            TransportDistinctMessageIdHitCount = 2
            MailboxEstimatedItemHitCount = 5
          }
        )
      }) -Metrics @{} -Errors @())
    )

    $result = Write-ImtRunSummary -RunContext ([pscustomobject]@{}) -StepResults $stepResults -StartedAt $startedAt -EndedAt $endedAt

    @($result.Data.FinalKeywordByMailboxRows).Count | Should Be 1
    @($result.Data.StepOutcomes).Count | Should Be 2
  }
}
