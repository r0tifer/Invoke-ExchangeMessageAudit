Set-StrictMode -Version Latest

$repoRoot = Split-Path -Path $PSScriptRoot -Parent
. (Join-Path $repoRoot 'src\\Models\\New-ResultObjects.ps1')
. (Join-Path $repoRoot 'src\\Logging\\Write-ImtLog.ps1')
. (Join-Path $repoRoot 'src\\Reporting\\Write-StepTables.ps1')

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
}
