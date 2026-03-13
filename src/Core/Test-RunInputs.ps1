Set-StrictMode -Version Latest

function Test-ImtRunInputs {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory=$true)]$RunContext
  )

  $inputs = $RunContext.Inputs

  if (-not $RunContext.RecipientMode -and -not $RunContext.HasSenderInput -and -not $RunContext.ParticipantMode -and -not $RunContext.HasMailboxScopeInput) {
    throw [System.ArgumentException]::new("Provide at least one of: -Participants, -Recipient/-Recipients, -Sender/-Senders, -SearchAllMailboxes, or -SourceMailboxes.")
  }

  if (($inputs.StartDate -and -not $inputs.EndDate) -or ($inputs.EndDate -and -not $inputs.StartDate)) {
    throw [System.ArgumentException]::new("Provide both -StartDate and -EndDate, or neither.")
  }

  if ($RunContext.End -lt $RunContext.Start) {
    throw [System.ArgumentException]::new("-EndDate must be greater than or equal to -StartDate.")
  }

  if (($inputs.PromptForMailboxExport -or $inputs.ExportLocatedEmails) -and [string]::IsNullOrWhiteSpace($inputs.ExportPstRoot)) {
    throw [System.ArgumentException]::new("Export requires -ExportPstRoot (UNC path).")
  }

  if ($inputs.DetailedMailboxEvidence -and [string]::IsNullOrWhiteSpace($inputs.EvidenceMailbox)) {
    throw [System.ArgumentException]::new("Detailed mailbox evidence requires -EvidenceMailbox.")
  }

  if ($inputs.SearchAllMailboxes -and $inputs.SourceMailboxes -and $inputs.SourceMailboxes.Count -gt 0) {
    throw [System.ArgumentException]::new("Use either -SearchAllMailboxes or -SourceMailboxes, not both.")
  }

  if ($inputs.PreflightOnly -and [string]::IsNullOrWhiteSpace($inputs.ExportPstRoot)) {
    throw [System.ArgumentException]::new("Preflight-only mode requires -ExportPstRoot (UNC path).")
  }

  if ($inputs.ExportPstRoot -and -not $inputs.ExportPstRoot.StartsWith('\\')) {
    throw [System.ArgumentException]::new("-ExportPstRoot must be a UNC path, for example \\fileserver\\PSTExports.")
  }

  New-ImtModuleResult -StepName 'ValidateInputs' -Status 'OK' -Summary 'Input validation passed.' -Data $null -Metrics @{} -Errors @()
}
