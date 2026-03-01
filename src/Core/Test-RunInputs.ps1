Set-StrictMode -Version Latest

function Test-ImtRunInputs {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory=$true)]$RunContext
  )

  $inputs = $RunContext.Inputs

  if (-not $RunContext.RecipientMode -and -not $RunContext.HasSenderInput -and -not $RunContext.ParticipantMode) {
    throw [System.ArgumentException]::new("Provide at least one of: -Participants, -Recipient, -Sender, or -Senders.")
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

  if ($inputs.PreflightOnly -and [string]::IsNullOrWhiteSpace($inputs.ExportPstRoot)) {
    throw [System.ArgumentException]::new("Preflight-only mode requires -ExportPstRoot (UNC path).")
  }

  if ($inputs.ExportPstRoot -and -not $inputs.ExportPstRoot.StartsWith('\\')) {
    throw [System.ArgumentException]::new("-ExportPstRoot must be a UNC path, for example \\fileserver\\PSTExports.")
  }

  New-ImtModuleResult -StepName 'ValidateInputs' -Status 'OK' -Summary 'Input validation passed.' -Data $null -Metrics @{} -Errors @()
}
