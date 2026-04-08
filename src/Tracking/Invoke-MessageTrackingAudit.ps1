Set-StrictMode -Version Latest

function Invoke-ImtMessageTrackingAudit {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory=$true)]$RunContext,
    [Parameter(Mandatory=$true)][string[]]$Servers,
    [Parameter(Mandatory=$true)][hashtable]$VersionInfo,
    [string[]]$TraceParticipants,
    [string[]]$EffectiveSenderFilters
  )

  $results = New-Object System.Collections.Generic.List[object]
  $serverFailures = 0
  $rawRecipientInputs = @()
  if ($RunContext.Inputs -and ($RunContext.Inputs.PSObject.Properties.Name -contains 'Recipients')) {
    $rawRecipientInputs = @($RunContext.Inputs.Recipients)
  } elseif ($RunContext.Inputs -and ($RunContext.Inputs.PSObject.Properties.Name -contains 'Recipient') -and $RunContext.Inputs.Recipient) {
    $rawRecipientInputs = @($RunContext.Inputs.Recipient)
  }
  $recipientFilters = @(
    $rawRecipientInputs |
      ForEach-Object { $_ -as [string] } |
      Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
      ForEach-Object { $_.Trim().ToLowerInvariant() } |
      Select-Object -Unique
  )

  foreach ($server in $Servers) {
    $version = $VersionInfo[$server]
    Write-ImtLog -Level DEBUG -Step 'MessageTrackingQuery' -EventType Progress -Message ("Querying {0} ({1})" -f $server, $version)

    try {
      $chunk = @()

      if ($RunContext.ParticipantMode) {
        $fromHits = @(
          foreach ($participant in $TraceParticipants) {
            Get-MessageTrackingLog -Server $server -Start $RunContext.Start -End $RunContext.End -ResultSize Unlimited -Sender $participant
          }
        )
        $toHits = @(
          foreach ($participant in $TraceParticipants) {
            Get-MessageTrackingLog -Server $server -Start $RunContext.Start -End $RunContext.End -ResultSize Unlimited -Recipients $participant
          }
        )
        $chunk = @($fromHits; $toHits)
      } elseif ($RunContext.RecipientMode -and $EffectiveSenderFilters.Count -gt 0) {
        $comboHits = @(
          foreach ($sender in $EffectiveSenderFilters) {
            foreach ($recipient in $recipientFilters) {
              Get-MessageTrackingLog -Server $server -Start $RunContext.Start -End $RunContext.End -ResultSize Unlimited -Sender $sender -Recipients $recipient
            }
          }
        )
        $chunk = @($comboHits)
      } elseif ($RunContext.RecipientMode) {
        $recipientHits = @(
          foreach ($recipient in $recipientFilters) {
            Get-MessageTrackingLog -Server $server -Start $RunContext.Start -End $RunContext.End -ResultSize Unlimited -Recipients $recipient
          }
        )
        $chunk = @($recipientHits)
      } elseif ($EffectiveSenderFilters.Count -gt 0) {
        $fromSenderHits = @(
          foreach ($sender in $EffectiveSenderFilters) {
            Get-MessageTrackingLog -Server $server -Start $RunContext.Start -End $RunContext.End -ResultSize Unlimited -Sender $sender
          }
        )
        $chunk = @($fromSenderHits)
      }

      # Force single-object query results into an array to keep pipeline behavior consistent.
      $chunk = @($chunk)

      if ($chunk) {
        $preFilterCount = @($chunk).Count

        if ($RunContext.Inputs.SubjectLike) {
          $chunk = $chunk | Where-Object { $_.MessageSubject -match [regex]::Escape($RunContext.Inputs.SubjectLike) }
        }

        if ($RunContext.Inputs.Keywords -and $RunContext.Inputs.Keywords.Count -gt 0) {
          $keywordPattern = (
            $RunContext.Inputs.Keywords |
              Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
              ForEach-Object { [regex]::Escape($_.Trim()) }
          ) -join '|'

          if ($keywordPattern) {
            $chunk = $chunk | Where-Object { ($_.MessageSubject -as [string]) -match $keywordPattern }
          }
        }

        if ($RunContext.ParticipantMode) {
          $chunk = $chunk | Where-Object {
            $senderMatch = $false
            $senderRaw = ($_.Sender -as [string])
            if ($senderRaw) {
              $senderMatch = $TraceParticipants -contains $senderRaw.ToLowerInvariant()
            }

            $recipientMatch = $false
            if ($_.Recipients) {
              $recipientMatch = @(
                $_.Recipients |
                  ForEach-Object {
                    $recipientValue = ($_ -as [string])
                    if ($recipientValue) {
                      $recipientValue.ToLowerInvariant()
                    }
                  } |
                  Where-Object { $_ -and ($TraceParticipants -contains $_) }
              ).Count -gt 0
            }

            $senderMatch -or $recipientMatch
          }
        }

        if ($recipientFilters.Count -gt 0) {
          $chunk = $chunk | Where-Object {
            if (-not $_.Recipients) {
              return $false
            }

            @(
              $_.Recipients |
                ForEach-Object {
                  $recipientValue = ($_ -as [string])
                  if ($recipientValue) {
                    $recipientValue.ToLowerInvariant()
                  }
                } |
                Where-Object { $_ -and ($recipientFilters -contains $_) }
            ).Count -gt 0
          }
        }

        if ($EffectiveSenderFilters.Count -gt 0) {
          $chunk = $chunk | Where-Object {
            $senderValue = ($_.Sender -as [string])
            if (-not $senderValue) {
              return $false
            }
            $EffectiveSenderFilters -contains $senderValue.ToLowerInvariant()
          }
        }

        if ($RunContext.Inputs.OnlyProblems) {
          $chunk = $chunk | Where-Object {
            $_.EventId -in 'FAIL','DEFER' -or
            (($_.RecipientStatus -join ' ') -match '(?i)(fail|defer|retry|smtp;4\.|smtp;5\.)')
          }
        }

        $chunk | Group-Object {
          "{0}|{1}|{2}|{3}|{4}" -f $_.ServerHostname, $_.InternalMessageId, $_.EventId, $_.Timestamp.ToUniversalTime().Ticks, $_.MessageId
        } | ForEach-Object {
          [void]$results.Add($_.Group[0])
        }

        $postFilterCount = @($chunk).Count
        Write-ImtLog -Level DEBUG -Step 'MessageTrackingQuery' -EventType Progress -Message ("Server={0}; Raw={1}; AfterFilters={2}" -f $server, $preFilterCount, $postFilterCount)
      } else {
        Write-ImtLog -Level DEBUG -Step 'MessageTrackingQuery' -EventType Progress -Message ("Server={0}; Raw=0" -f $server)
      }
    } catch {
      $serverFailures++
      Write-ImtLog -Level WARN -Step 'MessageTrackingQuery' -EventType Progress -Message ("[{0}] query failed: {1}" -f $server, $_.Exception.Message)
    }
  }

  if ($results.Count -eq 0) {
    $status = if ($serverFailures -gt 0) { 'FAIL' } else { 'WARN' }
    $summary = if ($serverFailures -gt 0) {
      "No tracking results returned; {0} server query failure(s)." -f $serverFailures
    } else {
      'No tracking results returned for selected filters/window.'
    }
    return New-ImtModuleResult -StepName 'MessageTrackingQuery' -Status $status -Summary $summary -Data ([pscustomobject]@{
      Results = @()
    }) -Metrics @{
      ResultCount = 0
      ServerFailures = $serverFailures
      ServerCount = $Servers.Count
    } -Errors @()
  }

  New-ImtModuleResult -StepName 'MessageTrackingQuery' -Status 'OK' -Summary ("Tracking results collected: {0}" -f $results.Count) -Data ([pscustomobject]@{
    Results = $results.ToArray()
  }) -Metrics @{
    ResultCount = $results.Count
    ServerFailures = $serverFailures
    ServerCount = $Servers.Count
  } -Errors @()
}
