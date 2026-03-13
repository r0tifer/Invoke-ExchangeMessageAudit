<#
.SYNOPSIS
  Compatibility launcher for Invoke-ExchangeMessageAudit module command.
#>

[CmdletBinding()]
param(
  [string]$Recipient,
  [string[]]$Recipients,
  [Alias('Sender')][string]$SenderAddress,
  [Alias('SenderList')][string[]]$Senders,
  [string[]]$Participants,
  [string[]]$SourceMailboxes,
  [int]$DaysBack = 90,
  [datetime]$StartDate,
  [datetime]$EndDate,
  [string]$OutputDir = 'C:\Temp',
  [string]$LogDir,
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
  [switch]$SearchAllMailboxes,
  [switch]$SearchMailboxesDirectly,
  [switch]$OutboundOnly,
  [switch]$DetailedMailboxEvidence,
  [string]$EvidenceMailbox,
  [switch]$DisableTranscriptLog,
  [switch]$SearchDumpsterDirectly,
  [switch]$ExpandExportScopeFromMatchedTraffic,
  [ValidateSet('DEBUG','INFO','WARN','ERROR','CRITICAL')][string]$OutputLevel = 'INFO'
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$moduleManifestPath = Join-Path -Path $PSScriptRoot -ChildPath 'Invoke-ExchangeMessageAudit.psd1'
Import-Module -Name $moduleManifestPath -Force

Invoke-ExchangeMessageAudit @PSBoundParameters

