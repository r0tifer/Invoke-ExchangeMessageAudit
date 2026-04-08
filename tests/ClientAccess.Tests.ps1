Set-StrictMode -Version Latest

$repoRoot = Split-Path -Path $PSScriptRoot -Parent
. (Join-Path $repoRoot 'src\\Models\\New-ResultObjects.ps1')
. (Join-Path $repoRoot 'src\\Logging\\Write-ImtLog.ps1')
. (Join-Path $repoRoot 'src\\Identity\\Resolve-Participants.ps1')
. (Join-Path $repoRoot 'src\\Tracking\\Invoke-MessageClientAccessAudit.ps1')

function Get-Recipient { }
function Get-Mailbox { }
function Search-MailboxAuditLog { }

Describe 'Invoke-ImtMessageClientAccessAudit' {
  BeforeEach {
    Mock Write-ImtLog {}
  }

  It 'correlates mailbox audit client details to grouped tracking results' {
    $runContext = [pscustomobject]@{
      Start = [datetime]'2026-04-06T15:00:00'
      End = [datetime]'2026-04-06T22:00:00'
    }

    Mock Resolve-ImtMailboxByAddress {
      param([string]$Address)

      if ($Address -eq 'jproger@arcticslope.org') {
        return [pscustomobject]@{
          Identity = 'jproger'
          PrimarySmtpAddress = 'jproger@arcticslope.org'
        }
      }
    }

    Mock Search-MailboxAuditLog {
      [pscustomobject]@{
        LastAccessed = [datetime]'2026-04-06T15:31:00'
        Operation = 'Update'
        OperationResult = 'Succeeded'
        LogonType = 'Owner'
        ItemSubject = 'Quarterly update'
        ClientInfoString = 'Client=MSExchangeRPC'
        ClientIPAddress = '10.0.0.24'
        ClientMachineName = 'JPROGER-LT'
        ClientProcessName = 'OUTLOOK.EXE'
        ClientVersion = '16.0.18730.20122'
        FolderPathName = '\\Sent Items'
      }
    }

    $results = @(
      [pscustomobject]@{
        Sender = 'jproger@arcticslope.org'
        Recipients = @('target@example.org')
        MessageSubject = 'Quarterly update'
        MessageId = '<msg-01@example.org>'
        InternalMessageId = '101'
        EventId = 'SEND'
        Timestamp = [datetime]'2026-04-06T15:31:30'
        ServerHostname = 'EXCH-01'
        Source = 'STOREDRIVER'
        ClientHostname = $null
        ConnectorId = $null
        SourceContext = 'MDB:01'
      }
      [pscustomobject]@{
        Sender = 'jproger@arcticslope.org'
        Recipients = @('target@example.org')
        MessageSubject = 'Quarterly update'
        MessageId = '<msg-01@example.org>'
        InternalMessageId = '101'
        EventId = 'SUBMIT'
        Timestamp = [datetime]'2026-04-06T15:30:45'
        ServerHostname = 'EXCH-01'
        Source = 'STOREDRIVER'
        ClientHostname = $null
        ConnectorId = $null
        SourceContext = 'MDB:01'
      }
    )

    $result = Invoke-ImtMessageClientAccessAudit -RunContext $runContext -Results $results -CandidateMailboxAddresses @('jproger@arcticslope.org')

    $result.Status | Should Be 'OK'
    @($result.Data.Rows).Count | Should Be 1
    @($result.Data.Rows)[0].AttributionSource | Should Be 'MailboxAudit'
    @($result.Data.Rows)[0].LikelyClient | Should Be 'Outlook desktop'
    @($result.Data.Rows)[0].ClientMachineName | Should Be 'JPROGER-LT'
  }
}

Describe 'Get-ImtMailboxAuditQueryParameters' {
  It 'uses an integer result size when the command metadata requires Int32' {
    function Search-MailboxAuditLog {
      param(
        [string]$Identity,
        [string[]]$LogonTypes,
        [switch]$ShowDetails,
        [datetime]$StartDate,
        [datetime]$EndDate,
        [int]$ResultSize
      )
    }

    $params = Get-ImtMailboxAuditQueryParameters `
      -Identity 'jproger@arcticslope.org' `
      -StartDate ([datetime]'2026-04-06T15:00:00') `
      -EndDate ([datetime]'2026-04-06T22:00:00')

    $params.ResultSize | Should Be 250000
  }
}

Describe 'Resolve-ImtTrackingDeviceAssessment' {
  It 'falls back to transport client host when mailbox audit data is unavailable' {
    $trailHints = [pscustomobject]@{
      ClientHostname = 'JPROGER-LT'
      ClientIPAddress = '10.0.0.24'
      ConnectorIds = ''
      Sources = 'SMTP'
      SourceContextSample = ''
      ServerHostnames = 'EXCH-SE-01'
    }

    $result = Resolve-ImtTrackingDeviceAssessment -AuditMatch $null -TrailHints $trailHints

    $result.AttributionSource | Should Be 'Transport'
    $result.AttributionConfidence | Should Be 'Low'
    $result.TransportClientHostname | Should Be 'JPROGER-LT'
  }
}
