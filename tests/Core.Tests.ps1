Set-StrictMode -Version Latest

$repoRoot = Split-Path -Path $PSScriptRoot -Parent
. (Join-Path $repoRoot 'src\\Models\\New-ResultObjects.ps1')
. (Join-Path $repoRoot 'src\\Core\\Initialize-RunContext.ps1')
. (Join-Path $repoRoot 'src\\Core\\Test-RunInputs.ps1')

Describe 'New-ImtModuleResult' {
  It 'builds a standard module result envelope' {
    $result = New-ImtModuleResult -StepName 'X' -Status OK -Summary 'done' -Data @{A=1} -Metrics @{Count=1} -Errors @()
    $result.StepName | Should Be 'X'
    $result.Status | Should Be 'OK'
    $result.Summary | Should Be 'done'
    $result.Metrics.Count | Should Be 1
  }
}

Describe 'Test-ImtRunInputs' {
  It 'throws when no participant/sender/recipient target is provided' {
    $ctx = Initialize-ImtRunContext -DaysBack 7 -OutputDir $env:TEMP -OutputLevel INFO
    { Test-ImtRunInputs -RunContext $ctx } | Should Throw
  }

  It 'throws when export is requested without ExportPstRoot' {
    $ctx = Initialize-ImtRunContext -Participants 'user@example.org' -DaysBack 7 -OutputDir $env:TEMP -ExportLocatedEmails -OutputLevel INFO
    { Test-ImtRunInputs -RunContext $ctx } | Should Throw
  }

  It 'throws when ExportPstRoot is not UNC' {
    $ctx = Initialize-ImtRunContext -Participants 'user@example.org' -DaysBack 7 -OutputDir $env:TEMP -ExportPstRoot 'C:\Exports' -OutputLevel INFO
    { Test-ImtRunInputs -RunContext $ctx } | Should Throw
  }

  It 'returns OK for valid minimal participant run' {
    $ctx = Initialize-ImtRunContext -Participants 'user@example.org' -DaysBack 7 -OutputDir $env:TEMP -OutputLevel INFO
    $result = Test-ImtRunInputs -RunContext $ctx
    $result.Status | Should Be 'OK'
  }
}

Describe 'Orchestrator usage path' {
  It 'prints usage when invoked without parameters' {
    $scriptPath = Join-Path $repoRoot 'Invoke-ExchangeMessageAudit.ps1'
    $output = & powershell -NoProfile -ExecutionPolicy Bypass -File $scriptPath 2>&1 | Out-String
    $output | Should Match 'Invoke-ExchangeMessageAudit.ps1'
    $output | Should Match 'Modular Exchange mail trace orchestrator'
  }
}

