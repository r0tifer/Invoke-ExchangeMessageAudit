Set-StrictMode -Version Latest

$sourceFiles = @(
  'src\Models\New-ResultObjects.ps1'
  'src\Logging\Write-ImtLog.ps1'
  'src\Core\Initialize-RunContext.ps1'
  'src\Core\Test-RunInputs.ps1'
  'src\Identity\Resolve-Participants.ps1'
  'src\Exchange\Get-TransportTopology.ps1'
  'src\Exchange\Get-RetentionSnapshot.ps1'
  'src\Tracking\Invoke-MessageTrackingAudit.ps1'
  'src\Tracking\Invoke-MessageClientAccessAudit.ps1'
  'src\Reporting\Export-TrackingReports.ps1'
  'src\Export\Test-ExportPrerequisites.ps1'
  'src\Export\Invoke-MailboxExportRequests.ps1'
  'src\MailboxSearch\Invoke-DirectMailboxSearch.ps1'
  'src\Reporting\Export-MailboxEvidenceReports.ps1'
  'src\Reporting\Export-CombinedKeywordReports.ps1'
  'src\Reporting\Write-StepTables.ps1'
  'src\Tracking\Invoke-MessageTrailTrace.ps1'
  'src\Reporting\Write-RunSummary.ps1'
  'src\Orchestration\Invoke-ExchangeMessageAudit.ps1'
)

foreach ($sourceFile in $sourceFiles) {
  $path = Join-Path -Path $PSScriptRoot -ChildPath $sourceFile
  if (-not (Test-Path -LiteralPath $path)) {
    throw "Required module file is missing: $path"
  }

  . $path
}

Export-ModuleMember -Function @('Invoke-ExchangeMessageAudit')
