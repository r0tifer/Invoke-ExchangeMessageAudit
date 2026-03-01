Set-StrictMode -Version Latest

function ConvertTo-ImtMetricTableRows {
  [CmdletBinding()]
  param(
    [hashtable]$Metrics
  )

  if (-not $Metrics -or $Metrics.Count -eq 0) {
    return @()
  }

  @(
    $Metrics.GetEnumerator() |
      Sort-Object Key |
      ForEach-Object {
        [pscustomobject]@{
          Metric = $_.Key
          Value = $_.Value
        }
      }
  )
}

function Write-ImtFormattedTable {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory=$true)][string]$StepName,
    [Parameter(Mandatory=$true)][string]$Title,
    [object[]]$Rows,
    [string[]]$Columns,
    [int]$MaxRows = 25
  )

  $rowArray = @($Rows)
  if ($rowArray.Count -eq 0) {
    return $false
  }

  $displayRows = if ($MaxRows -gt 0) {
    @($rowArray | Select-Object -First $MaxRows)
  } else {
    $rowArray
  }

  Write-Host ("[INFO] [{0}] [Table] {1}" -f $StepName, $Title) -ForegroundColor Gray

  $tableInput = if ($Columns -and $Columns.Count -gt 0) {
    $displayRows | Select-Object -Property $Columns
  } else {
    $displayRows
  }

  $tableText = $tableInput | Format-Table -AutoSize | Out-String -Width 4096
  Write-Host $tableText -ForegroundColor Gray

  if ($displayRows.Count -lt $rowArray.Count) {
    Write-Host ("[INFO] [{0}] [Table] Showing first {1} of {2} rows." -f $StepName, $displayRows.Count, $rowArray.Count) -ForegroundColor DarkGray
  }

  return $true
}

function Write-ImtStepDataTables {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory=$true)]$StepResult,
    [int]$MaxRows = 25
  )

  if (-not $StepResult) {
    return
  }

  $loggerState = Get-ImtLoggerState
  $outputLevel = if ($loggerState -and $loggerState.OutputLevel) {
    ($loggerState.OutputLevel -as [string]).ToUpperInvariant()
  } else {
    'INFO'
  }

  if (-not (Test-ImtShouldEmitConsole -CurrentOutputLevel $outputLevel -MessageLevel INFO -EventType Result)) {
    return
  }

  $stepName = ($StepResult.StepName -as [string])
  $data = $StepResult.Data
  $hasDetails = $false

  switch ($stepName) {
    'ResolveIdentities' {
      if ($data) {
        $rows = @(
          foreach ($participant in @($data.TraceParticipants)) {
            [pscustomobject]@{
              Category = 'TraceParticipant'
              Value = $participant
            }
          }

          foreach ($sender in @($data.EffectiveSenderFilters)) {
            [pscustomobject]@{
              Category = 'SenderFilter'
              Value = $sender
            }
          }

          foreach ($target in @($data.BaseTargetAddresses)) {
            [pscustomobject]@{
              Category = 'BaseTarget'
              Value = $target
            }
          }
        )

        if (Write-ImtFormattedTable -StepName $stepName -Title 'Resolved identity targets' -Rows $rows -Columns @('Category','Value') -MaxRows $MaxRows) {
          $hasDetails = $true
        }
      }
    }

    'DiscoverTransport' {
      if ($data -and $data.Servers) {
        $rows = @(
          foreach ($server in @($data.Servers)) {
            [pscustomobject]@{
              Server = $server
              Version = if ($data.VersionInfo) { $data.VersionInfo[$server] } else { $null }
            }
          }
        )

        if (Write-ImtFormattedTable -StepName $stepName -Title 'Transport servers' -Rows $rows -Columns @('Server','Version') -MaxRows $MaxRows) {
          $hasDetails = $true
        }
      }
    }

    'Preflight' {
      if ($data -and $data.Preflight) {
        $overview = @([pscustomobject]@{
          Ready = [bool]$data.Preflight.Ready
          MailboxCount = @($data.Mailboxes).Count
          HostingServerCount = @($data.HostingServers).Count
          IssueCount = @($data.Preflight.Issues).Count
          WarningCount = @($data.Preflight.Warnings).Count
        })

        if (Write-ImtFormattedTable -StepName $stepName -Title 'Preflight overview' -Rows $overview -Columns @('Ready','MailboxCount','HostingServerCount','IssueCount','WarningCount') -MaxRows $MaxRows) {
          $hasDetails = $true
        }

        $issueRows = @(
          foreach ($issue in @($data.Preflight.Issues)) {
            [pscustomobject]@{
              Severity = 'Issue'
              Message = $issue
            }
          }
          foreach ($warning in @($data.Preflight.Warnings)) {
            [pscustomobject]@{
              Severity = 'Warning'
              Message = $warning
            }
          }
        )
        if (Write-ImtFormattedTable -StepName $stepName -Title 'Preflight issues and warnings' -Rows $issueRows -Columns @('Severity','Message') -MaxRows $MaxRows) {
          $hasDetails = $true
        }
      }
    }

    'RetentionSnapshot' {
      if ($data -and $data.RetentionRows) {
        if (Write-ImtFormattedTable -StepName $stepName -Title 'Transport retention snapshot' -Rows @($data.RetentionRows) -Columns @('Name','MessageTrackingLogMaxAge','OldestLog','NewestLog') -MaxRows $MaxRows) {
          $hasDetails = $true
        }
      }
    }

    'MessageTrackingQuery' {
      if ($data -and $data.Results) {
        $serverRows = @(
          @($data.Results) |
            Group-Object ServerHostname |
            Sort-Object Name |
            ForEach-Object {
              $messageSet = @{}
              foreach ($row in $_.Group) {
                $messageId = ($row.MessageId -as [string])
                if (-not [string]::IsNullOrWhiteSpace($messageId)) {
                  $messageSet[$messageId.ToLowerInvariant()] = $true
                }
              }

              [pscustomobject]@{
                Server = if ([string]::IsNullOrWhiteSpace($_.Name)) { '(unknown)' } else { $_.Name }
                TransportEventHitCount = $_.Count
                TransportDistinctMessageIdHitCount = @($messageSet.Keys).Count
              }
            }
        )

        if (Write-ImtFormattedTable -StepName $stepName -Title 'Tracking hits by server' -Rows $serverRows -Columns @('Server','TransportEventHitCount','TransportDistinctMessageIdHitCount') -MaxRows $MaxRows) {
          $hasDetails = $true
        }
      }
    }

    'TrackingReport' {
      if ($data -and $data.TrackingKeywordRows) {
        $rows = @(
          foreach ($row in @($data.TrackingKeywordRows | Sort-Object Keyword)) {
            [pscustomobject]@{
              Keyword = $row.Keyword
              TransportEventHitCount = [int]($row.EventHitCount -as [int])
              TransportDistinctMessageIdHitCount = [int]($row.DistinctMessageIdHitCount -as [int])
            }
          }
        )

        if (Write-ImtFormattedTable -StepName $stepName -Title 'Tracking keyword summary' -Rows $rows -Columns @('Keyword','TransportEventHitCount','TransportDistinctMessageIdHitCount') -MaxRows $MaxRows) {
          $hasDetails = $true
        }
      }

      if ($data -and $data.DailyCounts) {
        if (Write-ImtFormattedTable -StepName $stepName -Title 'Tracking daily counts' -Rows @($data.DailyCounts) -Columns @('Date','Count') -MaxRows $MaxRows) {
          $hasDetails = $true
        }
      }
    }

    'MailboxExport' {
      if ($data -and $data.ExportRows) {
        if (Write-ImtFormattedTable -StepName $stepName -Title 'Mailbox export requests' -Rows @($data.ExportRows) -Columns @('Mailbox','Archive','RequestName','Status','FilePath') -MaxRows $MaxRows) {
          $hasDetails = $true
        }
      }
    }

    'DirectMailboxSearch' {
      if ($data -and $data.DirectKeywordRows) {
        $keywordTotals = @{}
        foreach ($row in @($data.DirectKeywordRows | Where-Object { $_.Status -eq 'OK' })) {
          $keyword = ($row.Keyword -as [string])
          if ([string]::IsNullOrWhiteSpace($keyword)) {
            continue
          }

          $keywordKey = $keyword.Trim().ToLowerInvariant()
          if (-not $keywordTotals.ContainsKey($keywordKey)) {
            $keywordTotals[$keywordKey] = [pscustomobject]@{
              Keyword = $keyword.Trim()
              MailboxEstimatedItemHitCount = 0
            }
          }

          $hitCount = ($row.HitCount -as [int])
          if ($null -eq $hitCount) { $hitCount = 0 }
          $keywordTotals[$keywordKey].MailboxEstimatedItemHitCount = [int]$keywordTotals[$keywordKey].MailboxEstimatedItemHitCount + [int]$hitCount
        }

        $rows = @($keywordTotals.GetEnumerator() | Sort-Object Name | ForEach-Object { $_.Value })
        if (Write-ImtFormattedTable -StepName $stepName -Title 'Direct mailbox keyword summary' -Rows $rows -Columns @('Keyword','MailboxEstimatedItemHitCount') -MaxRows $MaxRows) {
          $hasDetails = $true
        }
      }
    }

    'KeywordCombined' {
      if ($data -and $data.CombinedRows) {
        if (Write-ImtFormattedTable -StepName $stepName -Title 'Combined keyword summary' -Rows @($data.CombinedRows | Sort-Object Keyword) -Columns @('Keyword','TransportEventHitCount','TransportDistinctMessageIdHitCount','MailboxEstimatedItemHitCount') -MaxRows $MaxRows) {
          $hasDetails = $true
        }
      }
    }

    'MessageTrailTrace' {
      if ($data -and $data.TrailRows) {
        $rows = @(
          @($data.TrailRows) |
            Group-Object EventId |
            Sort-Object Name |
            ForEach-Object {
              [pscustomobject]@{
                EventId = if ([string]::IsNullOrWhiteSpace($_.Name)) { '(unknown)' } else { $_.Name }
                Count = $_.Count
              }
            }
        )

        if (Write-ImtFormattedTable -StepName $stepName -Title 'Trail events by EventId' -Rows $rows -Columns @('EventId','Count') -MaxRows $MaxRows) {
          $hasDetails = $true
        }
      }
    }

    'RetentionExport' {
      if ($data) {
        $rows = @([pscustomobject]@{
          RetentionCsv = $data.RetentionCsv
        })
        if (Write-ImtFormattedTable -StepName $stepName -Title 'Retention export artifact' -Rows $rows -Columns @('RetentionCsv') -MaxRows $MaxRows) {
          $hasDetails = $true
        }
      }
    }
  }

  if (-not $hasDetails) {
    $metricRows = @()
    if ($StepResult.PSObject.Properties.Name -contains 'Metrics') {
      $metricRows = ConvertTo-ImtMetricTableRows -Metrics $StepResult.Metrics
    }

    if ($metricRows.Count -gt 0) {
      [void](Write-ImtFormattedTable -StepName $stepName -Title 'Step metrics' -Rows $metricRows -Columns @('Metric','Value') -MaxRows $MaxRows)
    } else {
      $outcome = @([pscustomobject]@{
        Step = $stepName
        Status = $StepResult.Status
        Summary = $StepResult.Summary
      })
      [void](Write-ImtFormattedTable -StepName $stepName -Title 'Step outcome' -Rows $outcome -Columns @('Step','Status','Summary') -MaxRows $MaxRows)
    }
  }
}
