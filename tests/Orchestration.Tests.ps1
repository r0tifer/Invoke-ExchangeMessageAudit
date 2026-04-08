Set-StrictMode -Version Latest

$repoRoot = Split-Path -Path $PSScriptRoot -Parent
$testTempRoot = [System.IO.Path]::GetTempPath()
. (Join-Path $repoRoot 'src\Models\New-ResultObjects.ps1')
. (Join-Path $repoRoot 'src\Logging\Write-ImtLog.ps1')
. (Join-Path $repoRoot 'src\Core\Initialize-RunContext.ps1')
. (Join-Path $repoRoot 'src\Orchestration\Invoke-ExchangeMessageAudit.ps1')

function Test-ImtRunInputs { }
function Resolve-ImtParticipantsAndSenders { }
function Get-ImtTransportTopology { }
function Invoke-ImtMessageTrackingAudit { }
function Invoke-ImtMessageClientAccessAudit { }
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
    $tempDir = Join-Path $testTempRoot ("imt-orch-tests-{0}" -f ([guid]::NewGuid().ToString('N')))
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

Describe 'Invoke-ExchangeMessageAudit client access correlation' {
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
        EffectiveSenderFilters = @('jproger@arcticslope.org')
        BaseTargetAddresses = @('jproger@arcticslope.org')
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
        Results = @(
          [pscustomobject]@{
            Sender = 'jproger@arcticslope.org'
            Recipients = @('target@example.org')
            MessageSubject = 'Quarterly update'
            MessageId = '<msg-01@example.org>'
            InternalMessageId = '101'
            EventId = 'SUBMIT'
            Timestamp = [datetime]'2026-04-06T15:31:00'
            ServerHostname = 'EXCH-01'
          }
        )
      }) -Metrics @{} -Errors @()
    }
    Mock Invoke-ImtMessageClientAccessAudit {
      New-ImtModuleResult -StepName 'MessageClientAccess' -Status 'OK' -Summary 'ok' -Data ([pscustomobject]@{
        Rows = @(
          [pscustomobject]@{
            Mailbox = 'jproger@arcticslope.org'
            SubmittedAt = [datetime]'2026-04-06T15:31:00'
            Subject = 'Quarterly update'
            Recipients = 'target@example.org'
            AttributionSource = 'MailboxAudit'
            AttributionConfidence = 'High'
            LikelyClient = 'Outlook desktop'
            ClientMachineName = 'JPROGER-LT'
            TransportClientHostname = $null
          }
        )
        AuditRows = @()
        ProtocolRows = @()
        ActiveSyncRows = @(
          [pscustomobject]@{
            Mailbox = 'jproger@arcticslope.org'
            DeviceId = 'ApplABC123'
            DeviceFriendlyName = 'Jeff iPhone'
          }
        )
        AuditFailures = @()
        ProtocolFailures = @()
        ActiveSyncFailures = @()
      }) -Metrics @{} -Errors @()
    }
    $script:ReportedClientRows = @()
    $script:ReportedActiveSyncRows = @()
    Mock Export-ImtTrackingReports {
      param($RunContext, [object[]]$Results, [string[]]$BaseTargetAddresses, [object[]]$ClientAttributionRows, [object[]]$ClientAuditRows, [object[]]$ClientProtocolRows, [object[]]$ClientActiveSyncRows)
      $script:ReportedClientRows = @($ClientAttributionRows)
      $script:ReportedActiveSyncRows = @($ClientActiveSyncRows)
      New-ImtModuleResult -StepName 'TrackingReport' -Status 'OK' -Summary 'ok' -Data ([pscustomobject]@{
        CsvMain = $null
        ClientAttributionCsv = $null
        ClientAuditCsv = $null
        ClientProtocolCsv = $null
        ClientActiveSyncCsv = $null
        ClientAttributionRows = @($ClientAttributionRows)
        ClientAuditRows = @($ClientAuditRows)
        ClientProtocolRows = @($ClientProtocolRows)
        ClientActiveSyncRows = @($ClientActiveSyncRows)
        TrackingKeywordRows = @()
        TrackingKeywordMailboxRows = @()
        DailyCounts = @()
      }) -Metrics @{} -Errors @()
    }
    Mock Invoke-ImtDirectMailboxSearch {
      New-ImtModuleResult -StepName 'DirectMailboxSearch' -Status 'OK' -Summary 'ok' -Data ([pscustomobject]@{
        DirectKeywordRows = @()
        MatchedSourceMailboxAddresses = @()
        EvidenceRows = @()
      }) -Metrics @{} -Errors @()
    }
    Mock Export-ImtMailboxEvidenceReports {
      New-ImtModuleResult -StepName 'MailboxEvidence' -Status 'SKIP' -Summary 'no evidence' -Data ([pscustomobject]@{
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

  It 'passes client attribution rows into tracking reporting when requested' {
    $tempDir = Join-Path $testTempRoot ("imt-orch-client-access-{0}" -f ([guid]::NewGuid().ToString('N')))
    New-Item -ItemType Directory -Path $tempDir | Out-Null

    try {
      $null = Invoke-ExchangeMessageAudit `
        -Sender 'jproger@arcticslope.org' `
        -StartDate '2026-04-06T15:00:00' `
        -EndDate '2026-04-06T22:00:00' `
        -CorrelateClientAccess `
        -SkipRetentionCheck `
        -DisableTranscriptLog `
        -OutputDir $tempDir `
        -OutputLevel INFO

      @($script:ReportedClientRows).Count | Should Be 1
      $script:ReportedClientRows[0].ClientMachineName | Should Be 'JPROGER-LT'
      @($script:ReportedActiveSyncRows).Count | Should Be 1
      $script:ReportedActiveSyncRows[0].DeviceId | Should Be 'ApplABC123'
    } finally {
      Remove-Item -LiteralPath $tempDir -Recurse -Force -ErrorAction SilentlyContinue
    }
  }
}

Describe 'Invoke-ExchangeMessageAudit empty tracking path' {
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
      New-ImtModuleResult -StepName 'MessageTrackingQuery' -Status 'WARN' -Summary 'no tracking rows' -Data ([pscustomobject]@{
        Results = @()
      }) -Metrics @{} -Errors @()
    }
    Mock Export-ImtTrackingReports {
      param($RunContext, [object[]]$Results, [string[]]$BaseTargetAddresses)
      New-ImtModuleResult -StepName 'TrackingReport' -Status 'WARN' -Summary 'No results to report/export.' -Data ([pscustomobject]@{
        CsvMain = $null
        TrackingKeywordRows = @()
        TrackingKeywordMailboxRows = @()
        DailyCounts = @()
      }) -Metrics @{ ResultCount = @($Results).Count } -Errors @()
    }
    Mock Invoke-ImtDirectMailboxSearch {
      New-ImtModuleResult -StepName 'DirectMailboxSearch' -Status 'OK' -Summary 'ok' -Data ([pscustomobject]@{
        DirectKeywordRows = @()
        MatchedSourceMailboxAddresses = @()
        EvidenceRows = @()
      }) -Metrics @{} -Errors @()
    }
    Mock Export-ImtMailboxEvidenceReports {
      New-ImtModuleResult -StepName 'MailboxEvidence' -Status 'SKIP' -Summary 'no evidence' -Data ([pscustomobject]@{
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
    Mock Write-ImtRunSummary {
      param($RunContext, $StepResults, $StartedAt, $EndedAt)
      New-ImtModuleResult -StepName 'RunSummary' -Status 'OK' -Summary 'summary' -Data ([pscustomobject]@{
        TotalSteps = @($StepResults).Count
        Counts = [pscustomobject]@{ OK = 1; WARN = 1; FAIL = 0; SKIP = 1 }
        DurationSeconds = 1
        StepOutcomes = @()
        FinalKeywordByMailboxRows = @()
      }) -Metrics @{} -Errors @()
    }
  }

  It 'does not fail when tracking returns zero rows' {
    $tempDir = Join-Path $testTempRoot ("imt-orch-empty-tracking-{0}" -f ([guid]::NewGuid().ToString('N')))
    New-Item -ItemType Directory -Path $tempDir | Out-Null

    try {
      { Invoke-ExchangeMessageAudit `
          -Recipients 'riveracarolyn929@gmail.com' `
          -SourceMailboxes 'Rachel Aumavae' `
          -StartDate '2024-10-01T00:00:00' `
          -EndDate '2025-09-30T23:59:59' `
          -HasAttachmentOnly `
          -OutboundOnly `
          -SkipRetentionCheck `
          -DisableTranscriptLog `
          -OutputDir $tempDir `
          -OutputLevel INFO } | Should Not Throw
    } finally {
      Remove-Item -LiteralPath $tempDir -Recurse -Force -ErrorAction SilentlyContinue
    }
  }
}
