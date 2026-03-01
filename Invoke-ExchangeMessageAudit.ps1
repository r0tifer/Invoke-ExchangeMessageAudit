<#
.SYNOPSIS
  Exchange mail trace orchestrator with modular task execution.

.DESCRIPTION
  Orchestrates modular scripts under src/ for tracing, reporting, preflight, export, direct mailbox search,
  retention snapshot, and optional trail tracing.
#>

[CmdletBinding()]
param(
  [string]$Recipient,
  [Alias('Sender')][string]$SenderAddress,
  [Alias('SenderList')][string[]]$Senders,
  [string[]]$Participants,
  [int]$DaysBack = 90,
  [datetime]$StartDate,
  [datetime]$EndDate,
  [string]$OutputDir = 'C:\Temp',
  [string]$SubjectLike,
  [string[]]$Keywords,
  [switch]$HasAttachmentOnly,
  [switch]$OnlyProblems,
  [string]$TraceMessageId,
  [switch]$TraceLatest,
  [switch]$SkipRetentionCheck,
  [switch]$PromptForMailboxExport,
  [switch]$ExportLocatedEmails,
  [string]$ExportPstRoot,
  [switch]$IncludeArchive,
  [switch]$SkipDagPathValidation,
  [switch]$PreflightOnly,
  [switch]$SearchMailboxesDirectly,
  [switch]$DisableTranscriptLog,
  [switch]$SearchDumpsterDirectly,
  [switch]$ExpandExportScopeFromMatchedTraffic,
  [ValidateSet('DEBUG','INFO','WARN','ERROR','CRITICAL')][string]$OutputLevel = 'INFO'
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Show-InvokeExchangeMessageAuditUsage {
  $usage = @"
Invoke-ExchangeMessageAudit.ps1
--------------------
Modular Exchange mail trace orchestrator.

Core Parameters
---------------
-Participants <string[]>
-Recipient <string>
-Sender / -Senders <string|string[]>
-StartDate <datetime> -EndDate <datetime>
-Keywords <string[]>
-OutputDir <path>
-OutputLevel <DEBUG|INFO|WARN|ERROR|CRITICAL> (default: INFO)

Export Parameters
-----------------
-ExportLocatedEmails
-PromptForMailboxExport
-ExportPstRoot <UNC path>
-IncludeArchive
-PreflightOnly

Examples
--------
.\Invoke-ExchangeMessageAudit.ps1 -Participants "user1@contoso.org","user2@contoso.org" -StartDate "2025-01-01" -EndDate "2025-01-31" -Keywords "audit" -OutputLevel INFO
.\Invoke-ExchangeMessageAudit.ps1 -Participants "user1@contoso.org" -ExportPstRoot "\\fileserver\PSTExports" -PreflightOnly -OutputLevel DEBUG
"@
  Write-Host $usage
}

if ($PSBoundParameters.Count -eq 0) {
  Show-InvokeExchangeMessageAuditUsage
  return
}

$moduleFiles = @(
  'src\\Models\\New-ResultObjects.ps1',
  'src\\Logging\\Write-ImtLog.ps1',
  'src\\Core\\Initialize-RunContext.ps1',
  'src\\Core\\Test-RunInputs.ps1',
  'src\\Identity\\Resolve-Participants.ps1',
  'src\\Exchange\\Get-TransportTopology.ps1',
  'src\\Exchange\\Get-RetentionSnapshot.ps1',
  'src\\Tracking\\Invoke-MessageTrackingAudit.ps1',
  'src\\Reporting\\Export-TrackingReports.ps1',
  'src\\Export\\Test-ExportPrerequisites.ps1',
  'src\\Export\\Invoke-MailboxExportRequests.ps1',
  'src\\MailboxSearch\\Invoke-DirectMailboxSearch.ps1',
  'src\\Reporting\\Export-CombinedKeywordReports.ps1',
  'src\\Tracking\\Invoke-MessageTrailTrace.ps1',
  'src\\Reporting\\Write-RunSummary.ps1'
)

foreach ($moduleFile in $moduleFiles) {
  $fullPath = Join-Path -Path $PSScriptRoot -ChildPath $moduleFile
  if (-not (Test-Path -LiteralPath $fullPath)) {
    throw "Required module file is missing: $fullPath"
  }
  . $fullPath
}

$runContext = Initialize-ImtRunContext `
  -Recipient $Recipient `
  -SenderAddress $SenderAddress `
  -Senders $Senders `
  -Participants $Participants `
  -DaysBack $DaysBack `
  -StartDate $StartDate `
  -EndDate $EndDate `
  -OutputDir $OutputDir `
  -SubjectLike $SubjectLike `
  -Keywords $Keywords `
  -HasAttachmentOnly:$HasAttachmentOnly `
  -OnlyProblems:$OnlyProblems `
  -TraceMessageId $TraceMessageId `
  -TraceLatest:$TraceLatest `
  -SkipRetentionCheck:$SkipRetentionCheck `
  -PromptForMailboxExport:$PromptForMailboxExport `
  -ExportLocatedEmails:$ExportLocatedEmails `
  -ExportPstRoot $ExportPstRoot `
  -IncludeArchive:$IncludeArchive `
  -SkipDagPathValidation:$SkipDagPathValidation `
  -PreflightOnly:$PreflightOnly `
  -SearchMailboxesDirectly:$SearchMailboxesDirectly `
  -DisableTranscriptLog:$DisableTranscriptLog `
  -SearchDumpsterDirectly:$SearchDumpsterDirectly `
  -ExpandExportScopeFromMatchedTraffic:$ExpandExportScopeFromMatchedTraffic `
  -OutputLevel $OutputLevel

Initialize-ImtLogger -OutputLevel $runContext.OutputLevel -DisableTranscriptLog:$runContext.Inputs.DisableTranscriptLog -StepLogPath $runContext.StepLogPath -TranscriptPath $runContext.RunTranscriptPath

$startedAt = Get-Date
$stepResults = New-Object System.Collections.Generic.List[object]

$identityData = $null
$topologyData = $null
$retentionData = $null
$trackingData = $null
$trackingReportData = $null
$directSearchData = $null

function Invoke-ImtStep {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory=$true)][string]$StepName,
    [Parameter(Mandatory=$true)][scriptblock]$Action
  )

  Write-ImtLog -Level INFO -Step $StepName -EventType Start -Message ("Starting {0}." -f $StepName)

  try {
    $result = & $Action
    if (-not $result) {
      $result = New-ImtModuleResult -StepName $StepName -Status 'FAIL' -Summary 'Step returned no result object.' -Data $null -Metrics @{} -Errors @('Missing module result')
    }
  } catch {
    $result = New-ImtModuleResult -StepName $StepName -Status 'FAIL' -Summary ("{0} failed: {1}" -f $StepName, $_.Exception.Message) -Data $null -Metrics @{} -Errors @($_.Exception.Message)
  }

  [void]$stepResults.Add($result)

  $level = switch ($result.Status) {
    'OK' { 'INFO' }
    'SKIP' { 'INFO' }
    'WARN' { 'WARN' }
    'FAIL' { 'ERROR' }
    default { 'ERROR' }
  }

  Write-ImtLog -Level $level -Step $StepName -EventType Result -Message $result.Summary
  if ($result.Status -eq 'FAIL') {
    throw $result.Summary
  }

  $result
}

function Add-ImtSkippedStep {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory=$true)][string]$StepName,
    [Parameter(Mandatory=$true)][string]$Summary
  )

  $result = New-ImtModuleResult -StepName $StepName -Status 'SKIP' -Summary $Summary -Data $null -Metrics @{} -Errors @()
  [void]$stepResults.Add($result)
  Write-ImtLog -Level INFO -Step $StepName -EventType Result -Message $Summary
  $result
}

$fatalError = $null
try {
  Invoke-ImtStep -StepName 'ValidateInputs' -Action {
    Test-ImtRunInputs -RunContext $runContext
  } | Out-Null

  $identityResult = Invoke-ImtStep -StepName 'ResolveIdentities' -Action {
    Resolve-ImtParticipantsAndSenders -RunContext $runContext
  }
  $identityData = $identityResult.Data

  $topologyResult = Invoke-ImtStep -StepName 'DiscoverTransport' -Action {
    Get-ImtTransportTopology -RunContext $runContext
  }
  $topologyData = $topologyResult.Data

  if ($runContext.Inputs.PreflightOnly) {
    $preflightAddresses = Get-ImtTargetAddressSet -RunContext $runContext -ResolvedParticipants $identityData.ResolvedParticipants -TraceParticipants $identityData.TraceParticipants -EffectiveSenderFilters $identityData.EffectiveSenderFilters -UseResolvedParticipants

    Invoke-ImtStep -StepName 'Preflight' -Action {
      Invoke-ImtExportPreflight -RunContext $runContext -TargetAddresses $preflightAddresses
    } | Out-Null

    Add-ImtSkippedStep -StepName 'MessageTrackingQuery' -Summary 'Skipped because PreflightOnly mode is enabled.' | Out-Null
    Add-ImtSkippedStep -StepName 'TrackingReport' -Summary 'Skipped because PreflightOnly mode is enabled.' | Out-Null
    Add-ImtSkippedStep -StepName 'MailboxExport' -Summary 'Skipped because PreflightOnly mode is enabled.' | Out-Null
    Add-ImtSkippedStep -StepName 'DirectMailboxSearch' -Summary 'Skipped because PreflightOnly mode is enabled.' | Out-Null
    Add-ImtSkippedStep -StepName 'KeywordCombined' -Summary 'Skipped because PreflightOnly mode is enabled.' | Out-Null
    Add-ImtSkippedStep -StepName 'MessageTrailTrace' -Summary 'Skipped because PreflightOnly mode is enabled.' | Out-Null
    Add-ImtSkippedStep -StepName 'RetentionExport' -Summary 'Skipped because PreflightOnly mode is enabled.' | Out-Null
  } else {
    if (-not $runContext.Inputs.SkipRetentionCheck) {
      $retentionResult = Invoke-ImtStep -StepName 'RetentionSnapshot' -Action {
        Get-ImtRetentionSnapshot -RunContext $runContext -TransportTargets $topologyData.TransportTargets
      }
      $retentionData = $retentionResult.Data
    } else {
      Add-ImtSkippedStep -StepName 'RetentionSnapshot' -Summary 'Skipped retention snapshot collection due to -SkipRetentionCheck.' | Out-Null
    }

    $trackingResult = Invoke-ImtStep -StepName 'MessageTrackingQuery' -Action {
      Invoke-ImtMessageTrackingAudit -RunContext $runContext -Servers $topologyData.Servers -VersionInfo $topologyData.VersionInfo -TraceParticipants $identityData.TraceParticipants -EffectiveSenderFilters $identityData.EffectiveSenderFilters
    }
    $trackingData = $trackingResult.Data

    $trackingReportResult = Invoke-ImtStep -StepName 'TrackingReport' -Action {
      Export-ImtTrackingReports -RunContext $runContext -Results $trackingData.Results -BaseTargetAddresses $identityData.BaseTargetAddresses
    }
    $trackingReportData = $trackingReportResult.Data

    $runExport = $false
    if ($runContext.Inputs.ExportLocatedEmails) {
      $runExport = $true
    } elseif ($runContext.Inputs.PromptForMailboxExport) {
      $prompt = Read-Host 'Create mailbox export requests for identified mailboxes now? (Y/N)'
      if ($prompt -match '^(?i)y(es)?$') {
        $runExport = $true
      }
    }

    if ($runExport) {
      $exportAddresses = Get-ImtTargetAddressSet -RunContext $runContext -ResolvedParticipants $identityData.ResolvedParticipants -TraceParticipants $identityData.TraceParticipants -EffectiveSenderFilters $identityData.EffectiveSenderFilters -UseResolvedParticipants -IncludeMatchedTraffic:$runContext.Inputs.ExpandExportScopeFromMatchedTraffic -MatchedResults $trackingData.Results

      Invoke-ImtStep -StepName 'MailboxExport' -Action {
        Invoke-ImtMailboxExportRequests -RunContext $runContext -TargetAddresses $exportAddresses -EffectiveSenderFilters $identityData.EffectiveSenderFilters
      } | Out-Null
    } else {
      Add-ImtSkippedStep -StepName 'MailboxExport' -Summary 'Mailbox export was not requested.' | Out-Null
    }

    if ($runContext.DoDirectMailboxSearch) {
      $directSearchResult = Invoke-ImtStep -StepName 'DirectMailboxSearch' -Action {
        Invoke-ImtDirectMailboxSearch -RunContext $runContext -BaseTargetAddresses $identityData.BaseTargetAddresses -EffectiveSenderFilters $identityData.EffectiveSenderFilters
      }
      $directSearchData = $directSearchResult.Data
    } else {
      Add-ImtSkippedStep -StepName 'DirectMailboxSearch' -Summary 'Direct mailbox search not requested for this run.' | Out-Null
      $directSearchData = [pscustomobject]@{
        DirectKeywordRows = @()
      }
    }

    Invoke-ImtStep -StepName 'KeywordCombined' -Action {
      Export-ImtCombinedKeywordReports -RunContext $runContext -BaseTargetAddresses $identityData.BaseTargetAddresses -TrackingKeywordRows $trackingReportData.TrackingKeywordRows -TrackingKeywordMailboxRows $trackingReportData.TrackingKeywordMailboxRows -DirectKeywordRows $directSearchData.DirectKeywordRows
    } | Out-Null

    Invoke-ImtStep -StepName 'MessageTrailTrace' -Action {
      Invoke-ImtMessageTrailTrace -RunContext $runContext -Results $trackingData.Results -Servers $topologyData.Servers -VersionInfo $topologyData.VersionInfo
    } | Out-Null

    if (-not $runContext.Inputs.SkipRetentionCheck -and $retentionData -and $retentionData.RetentionRows) {
      Invoke-ImtStep -StepName 'RetentionExport' -Action {
        Export-ImtRetentionSnapshot -RunContext $runContext -RetentionRows $retentionData.RetentionRows
      } | Out-Null
    } else {
      Add-ImtSkippedStep -StepName 'RetentionExport' -Summary 'Retention export skipped (retention data unavailable or check disabled).' | Out-Null
    }
  }
} catch {
  $fatalError = $_
  $message = if ($_.Exception) { $_.Exception.Message } else { ($_ | Out-String) }
  Write-ImtLog -Level CRITICAL -Step 'RunFailure' -EventType Result -Message $message

  if (-not @($stepResults | Where-Object { $_.StepName -eq 'RunFailure' }).Count) {
    [void]$stepResults.Add((New-ImtModuleResult -StepName 'RunFailure' -Status 'FAIL' -Summary $message -Data $null -Metrics @{} -Errors @($message)))
  }
} finally {
  $endedAt = Get-Date
  Write-ImtRunSummary -RunContext $runContext -StepResults @($stepResults) -StartedAt $startedAt -EndedAt $endedAt | Out-Null
  Complete-ImtLogger

  if ($fatalError) {
    throw $fatalError
  }
}

