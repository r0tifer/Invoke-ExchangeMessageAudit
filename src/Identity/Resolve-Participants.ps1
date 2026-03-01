Set-StrictMode -Version Latest

function Test-ImtIsValidSmtpAddress {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory=$true)][string]$Address
  )

  try {
    $mail = [System.Net.Mail.MailAddress]::new($Address)
    return ($mail.Address -eq $Address -and $Address.Contains('@'))
  } catch {
    return $false
  }
}

function Resolve-ImtParticipantSmtp {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory=$true)][string]$Identity
  )

  if ($Identity -match '@') {
    return $Identity.ToLowerInvariant()
  }

  try {
    $recipient = Get-Recipient -Identity $Identity -ErrorAction Stop
    if ($recipient.PrimarySmtpAddress) {
      return $recipient.PrimarySmtpAddress.ToString().ToLowerInvariant()
    }
  } catch {
    # fallback probes below
  }

  try {
    $escaped = $Identity.Replace("'", "''")
    $recipientMatches = @(Get-Recipient -ResultSize Unlimited -Filter "DisplayName -eq '$escaped'" -ErrorAction Stop)
    if ($recipientMatches.Count -gt 0 -and $recipientMatches[0].PrimarySmtpAddress) {
      return $recipientMatches[0].PrimarySmtpAddress.ToString().ToLowerInvariant()
    }
  } catch {
    # fallback probes below
  }

  try {
    $anrMatches = @(Get-Recipient -Anr $Identity -ResultSize 10 -ErrorAction Stop)
    if ($anrMatches.Count -gt 0) {
      $exactDisplay = @($anrMatches | Where-Object { ($_.DisplayName -as [string]) -eq $Identity })
      $picked = if ($exactDisplay.Count -gt 0) { $exactDisplay[0] } else { $anrMatches[0] }
      if ($picked.PrimarySmtpAddress) {
        if ($anrMatches.Count -gt 1) {
          Write-ImtLog -Level WARN -Step 'ResolveIdentities' -EventType Progress -Message ("Participant '{0}' matched multiple recipients via ANR. Using '{1} <{2}>'." -f $Identity, $picked.DisplayName, $picked.PrimarySmtpAddress)
        }
        return $picked.PrimarySmtpAddress.ToString().ToLowerInvariant()
      }
    }
  } catch {
    # final fallback below
  }

  Write-ImtLog -Level WARN -Step 'ResolveIdentities' -EventType Progress -Message ("Unable to resolve participant '{0}' to SMTP address. Using input as-is." -f $Identity)
  return $Identity.ToLowerInvariant()
}

function Resolve-ImtMailboxByAddress {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory=$true)][string]$Address
  )

  try {
    return Get-Mailbox -Identity $Address -ErrorAction Stop
  } catch {
    return $null
  }
}

function Get-ImtTargetAddressSet {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory=$true)]$RunContext,
    [string[]]$ResolvedParticipants,
    [string[]]$TraceParticipants,
    [string[]]$EffectiveSenderFilters,
    [switch]$UseResolvedParticipants,
    [switch]$IncludeMatchedTraffic,
    [object[]]$MatchedResults
  )

  $set = New-Object System.Collections.ArrayList

  if ($RunContext.ParticipantMode) {
    $sourceParticipants = if ($UseResolvedParticipants) { @($ResolvedParticipants) } else { @($TraceParticipants) }
    foreach ($p in $sourceParticipants) {
      $value = ($p -as [string])
      if ($value) {
        $normalized = $value.ToLowerInvariant()
        if (-not $set.Contains($normalized)) {
          [void]$set.Add($normalized)
        }
      }
    }
  }

  if ($RunContext.HasSenderInput) {
    foreach ($sender in @($EffectiveSenderFilters)) {
      $value = ($sender -as [string])
      if ($value) {
        $normalized = $value.ToLowerInvariant()
        if (-not $set.Contains($normalized)) {
          [void]$set.Add($normalized)
        }
      }
    }
  }

  if ($RunContext.RecipientMode) {
    $value = ($RunContext.Inputs.Recipient -as [string])
    if ($value) {
      $normalized = $value.ToLowerInvariant()
      if (-not $set.Contains($normalized)) {
        [void]$set.Add($normalized)
      }
    }
  }

  if ($IncludeMatchedTraffic -and $MatchedResults) {
    foreach ($result in $MatchedResults) {
      $senderValue = ($result.Sender -as [string])
      if ($senderValue) {
        $normalizedSender = $senderValue.ToLowerInvariant()
        if (-not $set.Contains($normalizedSender)) {
          [void]$set.Add($normalizedSender)
        }
      }

      if ($result.Recipients) {
        foreach ($recipient in $result.Recipients) {
          $recipientValue = ($recipient -as [string])
          if ($recipientValue) {
            $normalizedRecipient = $recipientValue.ToLowerInvariant()
            if (-not $set.Contains($normalizedRecipient)) {
              [void]$set.Add($normalizedRecipient)
            }
          }
        }
      }
    }
  }

  @($set)
}

function Resolve-ImtMailboxesByAddressSet {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory=$true)][string[]]$Addresses,
    [switch]$CaptureUnresolved
  )

  $mailboxes = @()
  $unresolved = New-Object System.Collections.Generic.List[object]

  foreach ($address in $Addresses) {
    $mailbox = Resolve-ImtMailboxByAddress -Address $address
    if ($mailbox) {
      $mailboxes += $mailbox
    } elseif ($CaptureUnresolved) {
      $recipientType = 'Unresolved'
      try {
        $recipient = Get-Recipient -Identity $address -ErrorAction SilentlyContinue
        if ($recipient.RecipientType) {
          $recipientType = $recipient.RecipientType.ToString()
        }
      } catch {
        # ignore
      }

      [void]$unresolved.Add([pscustomobject]@{
        Address = $address
        Reason = 'Not a local searchable mailbox in this org/session'
        RecipientType = $recipientType
      })
    }
  }

  [pscustomobject]@{
    Mailboxes = @($mailboxes | Sort-Object DistinguishedName -Unique)
    Unresolved = @($unresolved)
  }
}

function Resolve-ImtParticipantsAndSenders {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory=$true)]$RunContext
  )

  $resolvedParticipants = @()
  $traceParticipants = @()
  $effectiveSenderFilters = @()

  if ($RunContext.ParticipantMode) {
    $resolvedParticipants = @(
      $RunContext.Inputs.Participants |
        Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
        ForEach-Object { Resolve-ImtParticipantSmtp -Identity $_.Trim() } |
        Select-Object -Unique
    )

    if ($resolvedParticipants.Count -eq 0) {
      throw "No valid participant identities resolved."
    }

    $traceParticipants = @(
      $resolvedParticipants |
        Where-Object { Test-ImtIsValidSmtpAddress -Address $_ } |
        Select-Object -Unique
    )

    $invalidForTrace = @($resolvedParticipants | Where-Object { -not (Test-ImtIsValidSmtpAddress -Address $_) })
    foreach ($bad in $invalidForTrace) {
      Write-ImtLog -Level WARN -Step 'ResolveIdentities' -EventType Progress -Message ("Skipping participant '{0}' for message tracking query because it is not a valid SMTP address." -f $bad)
    }

    if ($traceParticipants.Count -eq 0) {
      throw "No valid SMTP participants remain for tracking query. Provide SMTP addresses or resolvable identities."
    }
  }

  if ($RunContext.RawSenderList.Count -gt 0) {
    $resolvedSenders = @($RunContext.RawSenderList | ForEach-Object { Resolve-ImtParticipantSmtp -Identity $_ } | Select-Object -Unique)
    $effectiveSenderFilters = @($resolvedSenders | Where-Object { Test-ImtIsValidSmtpAddress -Address $_ } | Select-Object -Unique)

    $invalidSenders = @($resolvedSenders | Where-Object { -not (Test-ImtIsValidSmtpAddress -Address $_) } | Select-Object -Unique)
    foreach ($badSender in $invalidSenders) {
      Write-ImtLog -Level WARN -Step 'ResolveIdentities' -EventType Progress -Message ("Skipping sender '{0}' because it is not a valid SMTP address." -f $badSender)
    }

    if ($effectiveSenderFilters.Count -eq 0 -and -not $RunContext.ParticipantMode -and -not $RunContext.RecipientMode) {
      throw "No valid sender filters remain. Provide valid SMTP senders or resolvable sender identities."
    }
  }

  $baseTargetAddresses = Get-ImtTargetAddressSet -RunContext $RunContext -ResolvedParticipants $resolvedParticipants -TraceParticipants $traceParticipants -EffectiveSenderFilters $effectiveSenderFilters

  $summary = "Participants={0}; TraceParticipants={1}; SenderFilters={2}; BaseTargets={3}" -f $resolvedParticipants.Count, $traceParticipants.Count, $effectiveSenderFilters.Count, $baseTargetAddresses.Count
  New-ImtModuleResult -StepName 'ResolveIdentities' -Status 'OK' -Summary $summary -Data ([pscustomobject]@{
    ResolvedParticipants = @($resolvedParticipants)
    TraceParticipants = @($traceParticipants)
    EffectiveSenderFilters = @($effectiveSenderFilters)
    BaseTargetAddresses = @($baseTargetAddresses)
  }) -Metrics @{
    ResolvedParticipants = $resolvedParticipants.Count
    TraceParticipants = $traceParticipants.Count
    EffectiveSenderFilters = $effectiveSenderFilters.Count
    BaseTargetAddresses = $baseTargetAddresses.Count
  } -Errors @()
}
