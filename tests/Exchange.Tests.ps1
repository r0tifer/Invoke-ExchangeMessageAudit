Set-StrictMode -Version Latest

$repoRoot = Split-Path -Path $PSScriptRoot -Parent
. (Join-Path $repoRoot 'src\\Exchange\\Get-ActiveSyncDeviceEvidence.ps1')

function Get-MobileDevice {
  [CmdletBinding()]
  param(
    [string]$Mailbox
  )
}

function Get-MobileDeviceStatistics {
  [CmdletBinding()]
  param(
    [string]$Mailbox
  )
}

Describe 'Get-ImtActiveSyncDeviceEvidence' {
  It 'merges mobile device statistics and device rows by device identity' {
    Mock Get-MobileDeviceStatistics {
      [pscustomobject]@{
        Identity = 'jeffrey.roger@arcticslope.org\\ExchangeActiveSyncDevices\\ApplABC123'
        DeviceID = 'ApplABC123'
        DeviceType = 'iPhone'
        DeviceModel = 'iPhone 15'
        DeviceOS = 'iOS 17.4'
        Status = 'OK'
        LastSuccessSync = [datetime]'2026-04-06T19:01:10'
        LastSyncAttemptTime = [datetime]'2026-04-06T19:01:00'
      }
    }

    Mock Get-MobileDevice {
      [pscustomobject]@{
        Identity = 'jeffrey.roger@arcticslope.org\\ExchangeActiveSyncDevices\\ApplABC123'
        DeviceID = 'ApplABC123'
        FriendlyName = 'Jeff iPhone'
        DeviceUserAgent = 'Apple-iPhone/1704'
        ClientType = 'AirSync'
        DeviceAccessState = 'Allowed'
      }
    }

    $result = Get-ImtActiveSyncDeviceEvidence -MailboxIdentity 'jeffrey.roger@arcticslope.org'

    $result.Available | Should Be $true
    @($result.Rows).Count | Should Be 1
    @($result.Failures).Count | Should Be 0
    @($result.Rows)[0].Mailbox | Should Be 'jeffrey.roger@arcticslope.org'
    @($result.Rows)[0].DeviceId | Should Be 'ApplABC123'
    @($result.Rows)[0].DeviceFriendlyName | Should Be 'Jeff iPhone'
    @($result.Rows)[0].ClientType | Should Be 'AirSync'
    @($result.Rows)[0].LastSuccessSync | Should Be ([datetime]'2026-04-06T19:01:10')
  }
}
