Set-StrictMode -Version Latest

$repoRoot = Split-Path -Path $PSScriptRoot -Parent
. (Join-Path $repoRoot 'src\\Models\\New-ResultObjects.ps1')
. (Join-Path $repoRoot 'src\\Logging\\Write-ImtLog.ps1')

Describe 'Logging level filter behavior' {
  It 'INFO shows Start/Result and suppresses INFO Progress' {
    Test-ImtShouldEmitConsole -CurrentOutputLevel INFO -MessageLevel INFO -EventType Start | Should Be $true
    Test-ImtShouldEmitConsole -CurrentOutputLevel INFO -MessageLevel INFO -EventType Result | Should Be $true
    Test-ImtShouldEmitConsole -CurrentOutputLevel INFO -MessageLevel INFO -EventType Progress | Should Be $false
  }

  It 'INFO still shows warnings' {
    Test-ImtShouldEmitConsole -CurrentOutputLevel INFO -MessageLevel WARN -EventType Progress | Should Be $true
  }

  It 'WARN suppresses INFO and keeps WARN/ERROR plus Summary' {
    Test-ImtShouldEmitConsole -CurrentOutputLevel WARN -MessageLevel INFO -EventType Result | Should Be $false
    Test-ImtShouldEmitConsole -CurrentOutputLevel WARN -MessageLevel WARN -EventType Progress | Should Be $true
    Test-ImtShouldEmitConsole -CurrentOutputLevel WARN -MessageLevel ERROR -EventType Result | Should Be $true
    Test-ImtShouldEmitConsole -CurrentOutputLevel WARN -MessageLevel INFO -EventType Summary | Should Be $true
  }

  It 'ERROR and CRITICAL thresholds are enforced' {
    Test-ImtShouldEmitConsole -CurrentOutputLevel ERROR -MessageLevel WARN -EventType Result | Should Be $false
    Test-ImtShouldEmitConsole -CurrentOutputLevel ERROR -MessageLevel ERROR -EventType Result | Should Be $true
    Test-ImtShouldEmitConsole -CurrentOutputLevel CRITICAL -MessageLevel ERROR -EventType Result | Should Be $false
    Test-ImtShouldEmitConsole -CurrentOutputLevel CRITICAL -MessageLevel CRITICAL -EventType Result | Should Be $true
  }

  It 'DEBUG emits everything' {
    Test-ImtShouldEmitConsole -CurrentOutputLevel DEBUG -MessageLevel INFO -EventType Progress | Should Be $true
    Test-ImtShouldEmitConsole -CurrentOutputLevel DEBUG -MessageLevel DEBUG -EventType Progress | Should Be $true
  }
}

Describe 'Logger step file behavior' {
  It 'writes all events to step log when transcript logging is enabled' {
    $tempDir = Join-Path $env:TEMP ("imt-tests-{0}" -f ([guid]::NewGuid().ToString('N')))
    New-Item -ItemType Directory -Path $tempDir | Out-Null
    $stepLog = Join-Path $tempDir 'steps.log'
    $transcript = Join-Path $tempDir 'run.log'

    try {
      Initialize-ImtLogger -OutputLevel INFO -StepLogPath $stepLog -TranscriptPath $transcript
      Write-ImtLog -Level INFO -Step 'UnitTest' -EventType Start -Message 'start'
      Write-ImtLog -Level DEBUG -Step 'UnitTest' -EventType Progress -Message 'progress'
      Write-ImtLog -Level INFO -Step 'UnitTest' -EventType Result -Message 'result'
      Complete-ImtLogger

      Test-Path -LiteralPath $stepLog | Should Be $true
      $lines = Get-Content -Path $stepLog
      $lines.Count | Should BeGreaterThan 3
    } finally {
      Complete-ImtLogger
      Remove-Item -LiteralPath $tempDir -Recurse -Force -ErrorAction SilentlyContinue
    }
  }
}
