Set-StrictMode -Version Latest

function Initialize-ImtRunContext {
  [CmdletBinding()]
  param(
    [string]$Recipient,
    [string[]]$Recipients,
    [string]$SenderAddress,
    [string[]]$Senders,
    [string[]]$Participants,
    [string[]]$SourceMailboxes,
    [int]$DaysBack,
    [datetime]$StartDate,
    [datetime]$EndDate,
    [string]$OutputDir,
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
    [switch]$CorrelateClientAccess,
    [switch]$DisableTranscriptLog,
    [switch]$SearchDumpsterDirectly,
    [switch]$ExpandExportScopeFromMatchedTraffic,
    [ValidateSet('DEBUG','INFO','WARN','ERROR','CRITICAL')][string]$OutputLevel = 'INFO'
  )

  if ([string]::IsNullOrWhiteSpace($OutputDir)) {
    $OutputDir = [System.IO.Path]::GetTempPath()
  }

  if (-not (Test-Path -LiteralPath $OutputDir)) {
    New-Item -ItemType Directory -Path $OutputDir | Out-Null
  }
  $effectiveLogDir = if ([string]::IsNullOrWhiteSpace($LogDir)) { $OutputDir } else { $LogDir }
  if (-not (Test-Path -LiteralPath $effectiveLogDir)) {
    New-Item -ItemType Directory -Path $effectiveLogDir | Out-Null
  }

  $normalizedRecipients = @()
  if (-not [string]::IsNullOrWhiteSpace($Recipient)) {
    $normalizedRecipients += $Recipient.Trim()
  }
  if ($Recipients -and $Recipients.Count -gt 0) {
    $normalizedRecipients += @(
      $Recipients |
        Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
        ForEach-Object { $_.Trim() }
    )
  }
  $normalizedRecipients = @($normalizedRecipients | Select-Object -Unique)

  $normalizedSourceMailboxes = @()
  if ($SourceMailboxes -and $SourceMailboxes.Count -gt 0) {
    $normalizedSourceMailboxes = @(
      $SourceMailboxes |
        Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
        ForEach-Object { $_.Trim() } |
        Select-Object -Unique
    )
  }

  $hasSenderInput = (-not [string]::IsNullOrWhiteSpace($SenderAddress)) -or ($Senders -and $Senders.Count -gt 0)
  $recipientMode = $normalizedRecipients.Count -gt 0
  $pairMode = $recipientMode -and $hasSenderInput
  $participantMode = $Participants -and $Participants.Count -gt 0
  $hasMailboxScopeInput = [bool]$SearchAllMailboxes -or ($normalizedSourceMailboxes.Count -gt 0)

  $hasExplicitDateRange = $StartDate -or $EndDate
  if ($StartDate -and $EndDate) {
    $start = $StartDate
    $end = $EndDate
  } else {
    $start = (Get-Date).Date.AddDays(-[math]::Abs($DaysBack))
    $end = Get-Date
  }

  $ts = Get-Date -Format 'yyyyMMdd_HHmmss'
  $safeRecipient = if ($normalizedRecipients.Count -gt 0) { (($normalizedRecipients -join '_') -replace '[^\w@.-]','_') } else { 'any-recipient' }

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
    StepLogPath = Join-Path $effectiveLogDir ("MTL_Steps_{0}.log" -f $ts)
    RunTranscriptPath = Join-Path $effectiveLogDir ("MTL_RunTranscript_{0}.log" -f $ts)
    OutputLevel = $OutputLevel.ToUpperInvariant()
    OutputDir = $OutputDir
    LogDir = $effectiveLogDir
    HasSenderInput = $hasSenderInput
    RecipientMode = $recipientMode
    PairMode = $pairMode
    ParticipantMode = $participantMode
    HasMailboxScopeInput = $hasMailboxScopeInput
    DoDirectMailboxSearch = ($SearchMailboxesDirectly -or ($Keywords -and $Keywords.Count -gt 0) -or $HasAttachmentOnly -or $SearchAllMailboxes -or ($normalizedSourceMailboxes.Count -gt 0) -or $DetailedMailboxEvidence)
    Inputs = [pscustomobject]@{
      Recipient = if ($normalizedRecipients.Count -gt 0) { $normalizedRecipients[0] } else { $null }
      Recipients = @($normalizedRecipients)
      Participants = @($Participants)
      SourceMailboxes = @($normalizedSourceMailboxes)
      DaysBack = $DaysBack
      StartDate = $StartDate
      EndDate = $EndDate
      LogDir = $effectiveLogDir
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
      SearchAllMailboxes = [bool]$SearchAllMailboxes
      SearchMailboxesDirectly = [bool]$SearchMailboxesDirectly
      OutboundOnly = [bool]$OutboundOnly
      DetailedMailboxEvidence = [bool]$DetailedMailboxEvidence
      EvidenceMailbox = $EvidenceMailbox
      CorrelateClientAccess = [bool]$CorrelateClientAccess
      DisableTranscriptLog = [bool]$DisableTranscriptLog
      SearchDumpsterDirectly = [bool]$SearchDumpsterDirectly
      ExpandExportScopeFromMatchedTraffic = [bool]$ExpandExportScopeFromMatchedTraffic
    }
  }
}
