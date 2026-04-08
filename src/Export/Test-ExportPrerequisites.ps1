Set-StrictMode -Version Latest

function Get-ImtActiveMailboxServer {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory=$true)]$Mailbox
  )

  try {
    if ($Mailbox.Database) {
      $copies = @(Get-MailboxDatabaseCopyStatus -Identity $Mailbox.Database -ErrorAction Stop)
      $mounted = $copies | Where-Object { $_.Status -eq 'Mounted' } | Select-Object -First 1
      if ($mounted -and $mounted.Name) {
        return ($mounted.Name -split '\\')[0]
      }
    }
  } catch {
    # fallback below
  }

  if ($Mailbox.ServerName) {
    return $Mailbox.ServerName
  }

  return $null
}

function Test-ImtMailboxExportPrerequisites {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory=$true)][string]$PstRoot,
    [string[]]$MailboxServers,
    [switch]$SkipRemotePathValidation
  )

  $issues = New-Object System.Collections.Generic.List[string]
  $guidance = New-Object System.Collections.Generic.List[string]
  $warnings = New-Object System.Collections.Generic.List[string]

  if (-not (Get-Command -Name New-MailboxExportRequest -ErrorAction SilentlyContinue)) {
    [void]$issues.Add('Cmdlet New-MailboxExportRequest is not available in this session.')
    [void]$guidance.Add('Run from Exchange Management Shell on an Exchange mailbox server with mailbox tools installed.')
  }

  $currentIdentity = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name
  $hasExportRole = $false
  $roleCheckPerformed = $false
  $shortName = ($currentIdentity -split '\\')[-1]
  $roleCandidates = @($currentIdentity, $shortName) | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Select-Object -Unique

  foreach ($candidate in $roleCandidates) {
    try {
      $directAssignments = @(Get-ManagementRoleAssignment -Role 'Mailbox Import Export' -RoleAssignee $candidate -ErrorAction Stop)
      $roleCheckPerformed = $true
      if ($directAssignments.Count -gt 0) {
        $hasExportRole = $true
        break
      }
    } catch {
      # continue probing
    }
  }

  if (-not $hasExportRole) {
    try {
      $assignments = @(Get-ManagementRoleAssignment -Role 'Mailbox Import Export' -GetEffectiveUsers -ErrorAction Stop)
      $roleCheckPerformed = $true
      foreach ($assignment in $assignments) {
        $effective = ($assignment.EffectiveUserName -as [string])
        if (-not $effective) { continue }
        $effectiveLower = $effective.ToLowerInvariant()
        if ($effectiveLower -eq $currentIdentity.ToLowerInvariant()) { $hasExportRole = $true; break }
        if (($effectiveLower -split '\\')[-1] -eq $shortName.ToLowerInvariant()) { $hasExportRole = $true; break }
        if ($effectiveLower -like ("{0}@*" -f $shortName.ToLowerInvariant())) { $hasExportRole = $true; break }
      }
    } catch {
      [void]$warnings.Add('Unable to fully verify RBAC role assignments for Mailbox Import Export in this session.')
      [void]$guidance.Add("Verify manually with: Get-ManagementRoleAssignment -Role 'Mailbox Import Export' -GetEffectiveUsers")
    }
  }

  if (-not $hasExportRole -and $roleCheckPerformed) {
    [void]$issues.Add(("Current user '{0}' does not appear to have the Mailbox Import Export role." -f $currentIdentity))
    [void]$guidance.Add("Assign role: New-ManagementRoleAssignment -Role 'Mailbox Import Export' -User '<adminUser>'")
    [void]$guidance.Add('Restart Exchange Management Shell and re-run the script after RBAC replication.')
  }

  if (-not $PstRoot.StartsWith('\\')) {
    [void]$issues.Add(("Export path '{0}' is not a UNC path." -f $PstRoot))
    [void]$guidance.Add('Use a UNC path like \\fileserver\PSTExports.')
  } else {
    try {
      if (-not (Test-Path -LiteralPath $PstRoot)) {
        [void]$issues.Add(("UNC path '{0}' is not reachable from this session." -f $PstRoot))
        [void]$guidance.Add('Create/validate the share and ensure Exchange mailbox servers can write to it.')
      }
    } catch {
      [void]$issues.Add(("UNC path '{0}' could not be validated." -f $PstRoot))
      [void]$guidance.Add("Check share and NTFS permissions for 'Exchange Trusted Subsystem' (Modify/Write).")
    }
  }

  if ($MailboxServers -and $MailboxServers.Count -gt 0) {
    foreach ($server in ($MailboxServers | Sort-Object -Unique)) {
      try {
        $mrs = Get-Service -ComputerName $server -Name MSExchangeMailboxReplication -ErrorAction Stop
        if ($mrs.Status -ne 'Running') {
          [void]$issues.Add(("MSExchangeMailboxReplication is not running on {0}." -f $server))
          [void]$guidance.Add(("Start service on {0}: Start-Service MSExchangeMailboxReplication" -f $server))
        }
      } catch {
        [void]$issues.Add(("Unable to verify MSExchangeMailboxReplication service on {0}." -f $server))
        [void]$guidance.Add(("Verify service manually on {0} and ensure MRS is running." -f $server))
      }

      if (-not $SkipRemotePathValidation) {
        try {
          $remotePathOk = Invoke-Command -ComputerName $server -ScriptBlock { param($pathToCheck) Test-Path -LiteralPath $pathToCheck } -ArgumentList $PstRoot -ErrorAction Stop
          if (-not $remotePathOk) {
            [void]$issues.Add(("UNC path '{0}' is not reachable from mailbox server {1}." -f $PstRoot, $server))
            [void]$guidance.Add(("Grant share and NTFS write access for Exchange server computer accounts / Exchange Trusted Subsystem; verify from {0}." -f $server))
          }
        } catch {
          try {
            Enter-PSSession -ComputerName $server -ErrorAction Stop
            try {
              $remotePathFallback = Test-Path -LiteralPath $PstRoot
            } finally {
              Exit-PSSession
            }

            if (-not $remotePathFallback) {
              [void]$issues.Add(("UNC path '{0}' is not reachable from mailbox server {1} (Enter-PSSession fallback)." -f $PstRoot, $server))
              [void]$guidance.Add(("Grant share and NTFS write access for Exchange server computer accounts / Exchange Trusted Subsystem; verify from {0}." -f $server))
            }
          } catch {
            [void]$warnings.Add(("Unable to validate UNC path '{0}' from mailbox server {1} (Invoke-Command and Enter-PSSession unavailable/denied)." -f $PstRoot, $server))
            [void]$guidance.Add(("Enable/allow PowerShell remoting to {0} or run a manual Test-Path on that server for '{1}'." -f $server, $PstRoot))
          }
        }
      }
    }
  }

  [pscustomobject]@{
    Ready = ($issues.Count -eq 0)
    Issues = @($issues)
    Warnings = @($warnings)
    Guidance = @($guidance)
  }
}

function Invoke-ImtExportPreflight {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory=$true)]$RunContext,
    [Parameter(Mandatory=$true)][string[]]$TargetAddresses
  )

  $resolved = Resolve-ImtMailboxesByAddressSet -Addresses $TargetAddresses
  $mailboxes = @($resolved.Mailboxes)

  if ($mailboxes.Count -eq 0) {
    Write-ImtLog -Level WARN -Step 'Preflight' -EventType Progress -Message 'No mailboxes could be resolved from provided identities.'
  }

  $hostingServers = @()
  foreach ($mailbox in $mailboxes) {
    $hostingServer = Get-ImtActiveMailboxServer -Mailbox $mailbox
    if ($hostingServer) {
      $hostingServers += $hostingServer
    }
  }

  $hostingServers = @($hostingServers | Sort-Object -Unique)
  if ($hostingServers.Count -eq 0) {
    Write-ImtLog -Level WARN -Step 'Preflight' -EventType Progress -Message 'No hosting servers could be determined from resolved mailboxes.'
  }

  $preflight = Test-ImtMailboxExportPrerequisites -PstRoot $RunContext.Inputs.ExportPstRoot -MailboxServers $hostingServers -SkipRemotePathValidation:$RunContext.Inputs.SkipDagPathValidation
  $status = if ($preflight.Ready) { 'OK' } else { 'FAIL' }

  if (-not $preflight.Ready) {
    foreach ($issue in $preflight.Issues) {
      Write-ImtLog -Level WARN -Step 'Preflight' -EventType Progress -Message $issue
    }
  }

  if ($preflight.Warnings.Count -gt 0) {
    foreach ($warning in $preflight.Warnings) {
      Write-ImtLog -Level WARN -Step 'Preflight' -EventType Progress -Message $warning
    }
  }

  New-ImtModuleResult -StepName 'Preflight' -Status $status -Summary ("MailboxCount={0}; HostingServers={1}; Issues={2}; Warnings={3}" -f $mailboxes.Count, ($hostingServers -join ','), $preflight.Issues.Count, $preflight.Warnings.Count) -Data ([pscustomobject]@{
    Mailboxes = @($mailboxes)
    HostingServers = @($hostingServers)
    Preflight = $preflight
  }) -Metrics @{
    MailboxCount = $mailboxes.Count
    HostingServers = $hostingServers.Count
    Issues = $preflight.Issues.Count
    Warnings = $preflight.Warnings.Count
  } -Errors @($preflight.Issues)
}
