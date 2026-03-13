Set-StrictMode -Version Latest

$repoRoot = Split-Path -Path $PSScriptRoot -Parent
. (Join-Path $repoRoot 'src\Models\New-ResultObjects.ps1')
. (Join-Path $repoRoot 'src\Logging\Write-ImtLog.ps1')
. (Join-Path $repoRoot 'src\Core\Initialize-RunContext.ps1')
. (Join-Path $repoRoot 'src\Orchestration\Invoke-ExchangeMessageAudit.ps1')

function Test-ImtRunInputs { }
function Resolve-ImtParticipantsAndSenders { }
function Get-ImtTransportTopology { }
function Invoke-ImtMessageTrackingAudit { }
function Export-ImtTrackingReports { }
function Invoke-ImtDirectMailboxSearch { }
function Export-ImtMailboxEvidenceReports { }
function Invoke-ImtMailboxExportRequests { }
function Export-ImtCombinedKeywordReports { }
function Invoke-ImtMessageTrailTrace { }
function Write-ImtRunSummary { }
function Write-ImtStepDataTables { }

Describe 'Invoke-ExchangeMessageAudit export targeting' {
  BeforeEach {
    Mock Initialize-ImtLogger {}
    Mock Complete-ImtLogger {}
    Mock Write-ImtLog {}
    Mock Write-ImtStepDataTables {}
    Mock Test-ImtRunInputs {
      New-ImtModuleResult -StepName 'ValidateInputs' -Status 'OK' -Summary 'ok' -Data $null -Metrics @{} -Errors @()
    }
    Mock Resolve-ImtParticipantsAndSenders {
      New-ImtModuleResult -StepName 'ResolveIdentities' -Status 'OK' -Summary 'ok' -Data ([pscustomobject]@{
        ResolvedParticipants = @()
        TraceParticipants = @()
        EffectiveSenderFilters = @()
        BaseTargetAddresses = @()
      }) -Metrics @{} -Errors @()
    }
    Mock Get-ImtTransportTopology {
      New-ImtModuleResult -StepName 'DiscoverTransport' -Status 'OK' -Summary 'ok' -Data ([pscustomobject]@{
        Servers = @('EXCH-01')
        VersionInfo = @{ 'EXCH-01' = 'Exchange 2019' }
        TransportTargets = @()
      }) -Metrics @{} -Errors @()
    }
    Mock Invoke-ImtMessageTrackingAudit {
      New-ImtModuleResult -StepName 'MessageTrackingQuery' -Status 'OK' -Summary 'ok' -Data ([pscustomobject]@{
        Results = @()
      }) -Metrics @{} -Errors @()
    }
    Mock Export-ImtTrackingReports {
      New-ImtModuleResult -StepName 'TrackingReport' -Status 'OK' -Summary 'ok' -Data ([pscustomobject]@{
        TrackingKeywordRows = @()
        TrackingKeywordMailboxRows = @()
        DailyCounts = @()
      }) -Metrics @{} -Errors @()
    }
    Mock Invoke-ImtDirectMailboxSearch {
      New-ImtModuleResult -StepName 'DirectMailboxSearch' -Status 'OK' -Summary 'ok' -Data ([pscustomobject]@{
        DirectKeywordRows = @()
        MatchedSourceMailboxAddresses = @('rachel.aumavae@example.org')
        EvidenceRows = @()
      }) -Metrics @{} -Errors @()
    }
    Mock Export-ImtCombinedKeywordReports {
      New-ImtModuleResult -StepName 'KeywordCombined' -Status 'SKIP' -Summary 'skip' -Data ([pscustomobject]@{
        CombinedByMailboxRows = @()
      }) -Metrics @{} -Errors @()
    }
    Mock Invoke-ImtMessageTrailTrace {
      New-ImtModuleResult -StepName 'MessageTrailTrace' -Status 'SKIP' -Summary 'skip' -Data ([pscustomobject]@{}) -Metrics @{} -Errors @()
    }
    $script:ExportTargetAddresses = $null
    function Invoke-ImtMailboxExportRequests {
      param($RunContext, [string[]]$TargetAddresses, [string[]]$EffectiveSenderFilters)
      $script:ExportTargetAddresses = @($TargetAddresses)
      New-ImtModuleResult -StepName 'MailboxExport' -Status 'OK' -Summary ($TargetAddresses -join ',') -Data ([pscustomobject]@{}) -Metrics @{} -Errors @()
    }
    Mock Write-ImtRunSummary {
      New-ImtModuleResult -StepName 'RunSummary' -Status 'OK' -Summary 'summary' -Data ([pscustomobject]@{
        TotalSteps = 1
        Counts = [pscustomobject]@{ OK = 1; WARN = 0; FAIL = 0; SKIP = 0 }
        DurationSeconds = 1
        StepOutcomes = @()
        FinalKeywordByMailboxRows = @()
      }) -Metrics @{} -Errors @()
    }
  }

  It 'exports only matched source mailboxes when direct mailbox search identifies hits' {
    $tempDir = Join-Path $env:TEMP ("imt-orch-tests-{0}" -f ([guid]::NewGuid().ToString('N')))
    New-Item -ItemType Directory -Path $tempDir | Out-Null

    try {
      $null = Invoke-ExchangeMessageAudit `
        -Recipients 'riveracarolyn929@gmail.com' `
        -SearchAllMailboxes `
        -StartDate '2024-10-01T00:00:00' `
        -EndDate '2025-09-30T23:59:59' `
        -HasAttachmentOnly `
        -ExportLocatedEmails `
        -ExportPstRoot '\\fileserver\PST' `
        -SkipRetentionCheck `
        -DisableTranscriptLog `
        -OutputDir $tempDir `
        -OutputLevel INFO

      @($script:ExportTargetAddresses).Count | Should Be 1
      $script:ExportTargetAddresses[0] | Should Be 'rachel.aumavae@example.org'
    } finally {
      Remove-Item -LiteralPath $tempDir -Recurse -Force -ErrorAction SilentlyContinue
    }
  }
}
