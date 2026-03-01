Set-StrictMode -Version Latest

function Initialize-ImtRunContext {
  [CmdletBinding()]
  param(
    [string]$Recipient,
    [string]$SenderAddress,
    [string[]]$Senders,
    [string[]]$Participants,
    [int]$DaysBack,
    [datetime]$StartDate,
    [datetime]$EndDate,
    [string]$OutputDir,
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

  if (-not (Test-Path -LiteralPath $OutputDir)) {
    New-Item -ItemType Directory -Path $OutputDir | Out-Null
  }

  $hasSenderInput = (-not [string]::IsNullOrWhiteSpace($SenderAddress)) -or ($Senders -and $Senders.Count -gt 0)
  $recipientMode = -not [string]::IsNullOrWhiteSpace($Recipient)
  $pairMode = $recipientMode -and $hasSenderInput
  $participantMode = $Participants -and $Participants.Count -gt 0

  $hasExplicitDateRange = $StartDate -or $EndDate
  if ($StartDate -and $EndDate) {
    $start = $StartDate
    $end = $EndDate
  } else {
    $start = (Get-Date).Date.AddDays(-[math]::Abs($DaysBack))
    $end = Get-Date
  }

  $ts = Get-Date -Format 'yyyyMMdd_HHmmss'
  $safeRecipient = if ($Recipient) { ($Recipient -replace '[^\w@.-]','_') } else { 'any-recipient' }

  $rawSenderList = @()
  if (-not [string]::IsNullOrWhiteSpace($SenderAddress)) {
    $rawSenderList += $SenderAddress.Trim()
  }
  if ($Senders -and $Senders.Count -gt 0) {
    $rawSenderList += @(
      $Senders |
        Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
        ForEach-Object { $_.Trim() }
    )
  }

  $safeSender = if ($rawSenderList.Count -gt 0) { (($rawSenderList -join '_') -replace '[^\w@.-]','_') } else { 'any-sender' }

  [pscustomobject]@{
    Timestamp = $ts
    Start = $start
    End = $end
    HasExplicitDateRange = [bool]$hasExplicitDateRange
    RawSenderList = @($rawSenderList)
    SafeRecipient = $safeRecipient
    SafeSender = $safeSender
    StepLogPath = Join-Path $OutputDir ("MTL_Steps_{0}.log" -f $ts)
    RunTranscriptPath = Join-Path $OutputDir ("MTL_RunTranscript_{0}.log" -f $ts)
    OutputLevel = $OutputLevel.ToUpperInvariant()
    OutputDir = $OutputDir
    HasSenderInput = $hasSenderInput
    RecipientMode = $recipientMode
    PairMode = $pairMode
    ParticipantMode = $participantMode
    DoDirectMailboxSearch = ($SearchMailboxesDirectly -or ($Keywords -and $Keywords.Count -gt 0) -or $HasAttachmentOnly)
    Inputs = [pscustomobject]@{
      Recipient = $Recipient
      Participants = @($Participants)
      DaysBack = $DaysBack
      StartDate = $StartDate
      EndDate = $EndDate
      SubjectLike = $SubjectLike
      Keywords = @($Keywords)
      HasAttachmentOnly = [bool]$HasAttachmentOnly
      OnlyProblems = [bool]$OnlyProblems
      TraceMessageId = $TraceMessageId
      TraceLatest = [bool]$TraceLatest
      SkipRetentionCheck = [bool]$SkipRetentionCheck
      PromptForMailboxExport = [bool]$PromptForMailboxExport
      ExportLocatedEmails = [bool]$ExportLocatedEmails
      ExportPstRoot = $ExportPstRoot
      IncludeArchive = [bool]$IncludeArchive
      SkipDagPathValidation = [bool]$SkipDagPathValidation
      PreflightOnly = [bool]$PreflightOnly
      SearchMailboxesDirectly = [bool]$SearchMailboxesDirectly
      DisableTranscriptLog = [bool]$DisableTranscriptLog
      SearchDumpsterDirectly = [bool]$SearchDumpsterDirectly
      ExpandExportScopeFromMatchedTraffic = [bool]$ExpandExportScopeFromMatchedTraffic
    }
  }
}
