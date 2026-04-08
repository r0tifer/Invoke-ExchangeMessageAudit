Set-StrictMode -Version Latest

$repoRoot = Split-Path -Path $PSScriptRoot -Parent
$testTempRoot = [System.IO.Path]::GetTempPath()
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
    $ctx = Initialize-ImtRunContext -DaysBack 7 -OutputDir $testTempRoot -OutputLevel INFO
    { Test-ImtRunInputs -RunContext $ctx } | Should Throw
  }

  It 'throws when export is requested without ExportPstRoot' {
    $ctx = Initialize-ImtRunContext -Participants 'user@example.org' -DaysBack 7 -OutputDir $testTempRoot -ExportLocatedEmails -OutputLevel INFO
    { Test-ImtRunInputs -RunContext $ctx } | Should Throw
  }

  It 'throws when ExportPstRoot is not UNC' {
    $ctx = Initialize-ImtRunContext -Participants 'user@example.org' -DaysBack 7 -OutputDir $testTempRoot -ExportPstRoot 'C:\Exports' -OutputLevel INFO
    { Test-ImtRunInputs -RunContext $ctx } | Should Throw
  }

  It 'returns OK for valid minimal participant run' {
    $ctx = Initialize-ImtRunContext -Participants 'user@example.org' -DaysBack 7 -OutputDir $testTempRoot -OutputLevel INFO
    $result = Test-ImtRunInputs -RunContext $ctx
    $result.Status | Should Be 'OK'
  }

  It 'normalizes Recipient and Recipients into one recipient list' {
    $ctx = Initialize-ImtRunContext -Recipient 'one@example.org' -Recipients @('two@example.org', 'one@example.org') -DaysBack 7 -OutputDir $testTempRoot -OutputLevel INFO
    @($ctx.Inputs.Recipients).Count | Should Be 2
    $ctx.Inputs.Recipient | Should Be 'one@example.org'
    @($ctx.Inputs.Recipients) -contains 'one@example.org' | Should Be $true
    @($ctx.Inputs.Recipients) -contains 'two@example.org' | Should Be $true
  }

  It 'throws when detailed mailbox evidence is requested without EvidenceMailbox' {
    $ctx = Initialize-ImtRunContext -Recipients 'target@example.org' -DaysBack 7 -OutputDir $testTempRoot -DetailedMailboxEvidence -OutputLevel INFO
    { Test-ImtRunInputs -RunContext $ctx } | Should Throw
  }

  It 'throws when SearchAllMailboxes and SourceMailboxes are combined' {
    $ctx = Initialize-ImtRunContext -Recipients 'target@example.org' -DaysBack 7 -OutputDir $testTempRoot -SearchAllMailboxes -SourceMailboxes 'user@example.org' -OutputLevel INFO
    { Test-ImtRunInputs -RunContext $ctx } | Should Throw
  }

  It 'stores CorrelateClientAccess in the run context inputs' {
    $ctx = Initialize-ImtRunContext -SenderAddress 'user@example.org' -DaysBack 7 -OutputDir $testTempRoot -CorrelateClientAccess -OutputLevel INFO
    $ctx.Inputs.CorrelateClientAccess | Should Be $true
  }
}

Describe 'Initialize-ImtRunContext logging paths' {
  It 'defaults logs to OutputDir when LogDir is not provided' {
    $outputDir = Join-Path $testTempRoot ("imt-core-tests-output-{0}" -f ([guid]::NewGuid().ToString('N')))
    New-Item -ItemType Directory -Path $outputDir | Out-Null

    try {
      $ctx = Initialize-ImtRunContext -Participants 'user@example.org' -DaysBack 7 -OutputDir $outputDir -OutputLevel INFO
      (Split-Path -Path $ctx.StepLogPath -Parent) | Should Be $outputDir
      (Split-Path -Path $ctx.RunTranscriptPath -Parent) | Should Be $outputDir
    } finally {
      Remove-Item -LiteralPath $outputDir -Recurse -Force -ErrorAction SilentlyContinue
    }
  }

  It 'writes logs to LogDir when provided' {
    $outputDir = Join-Path $testTempRoot ("imt-core-tests-output-{0}" -f ([guid]::NewGuid().ToString('N')))
    $logDir = Join-Path $testTempRoot ("imt-core-tests-logs-{0}" -f ([guid]::NewGuid().ToString('N')))
    New-Item -ItemType Directory -Path $outputDir | Out-Null

    try {
      $ctx = Initialize-ImtRunContext -Participants 'user@example.org' -DaysBack 7 -OutputDir $outputDir -LogDir $logDir -OutputLevel INFO
      (Split-Path -Path $ctx.StepLogPath -Parent) | Should Be $logDir
      (Split-Path -Path $ctx.RunTranscriptPath -Parent) | Should Be $logDir
      $ctx.LogDir | Should Be $logDir
    } finally {
      Remove-Item -LiteralPath $outputDir -Recurse -Force -ErrorAction SilentlyContinue
      Remove-Item -LiteralPath $logDir -Recurse -Force -ErrorAction SilentlyContinue
    }
  }
}

Describe 'Orchestrator usage path' {
  It 'prints usage when invoked without parameters' {
    $scriptPath = Join-Path $repoRoot 'Invoke-ExchangeMessageAudit.ps1'
    $output = & powershell -NoProfile -ExecutionPolicy Bypass -File $scriptPath 2>&1 | Out-String
    $output | Should Match 'Invoke-ExchangeMessageAudit'
    $output | Should Match 'Modular Exchange mail trace orchestrator'
  }
}

Describe 'Module packaging' {
  It 'imports the module manifest and exposes Invoke-ExchangeMessageAudit' {
    $manifestPath = Join-Path $repoRoot 'Invoke-ExchangeMessageAudit.psd1'
    Import-Module -Name $manifestPath -Force
    try {
      $command = Get-Command -Name Invoke-ExchangeMessageAudit -CommandType Function -ErrorAction Stop
      $command.Name | Should Be 'Invoke-ExchangeMessageAudit'
    } finally {
      Remove-Module -Name Invoke-ExchangeMessageAudit -Force -ErrorAction SilentlyContinue
    }
  }
}

