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

  It 'falls back to protocol logs when mailbox audit rows are unavailable' {
    $runContext = [pscustomobject]@{
      Start = [datetime]'2026-04-06T15:00:00'
      End = [datetime]'2026-04-06T22:00:00'
    }

    Mock Resolve-ImtMailboxByAddress {
      [pscustomobject]@{
        Identity = 'jroger'
        PrimarySmtpAddress = 'jeffrey.roger@arcticslope.org'
        ExchangeGuid = '7f000a6b-7628-4c1c-8f74-daa11b39b3ab'
      }
    }

    Mock Search-MailboxAuditLog { @() }

    Mock Get-ImtProtocolEvidenceRowsForSenders {
      [pscustomobject]@{
        Rows = @(
          [pscustomobject]@{
            Mailbox = 'jeffrey.roger@arcticslope.org'
            EvidenceType = 'HttpProxyMapi'
            Server = 'EXCH-SE-01'
            LogPath = 'C:\Logs\HttpProxy_2026040619-1.LOG'
            Timestamp = [datetime]'2026-04-06T19:00:10'
            Protocol = 'Mapi'
            UrlStem = '/mapi/emsmdb/'
            UserAgent = 'Microsoft Office/16.0 (Windows NT 10.0; Microsoft Outlook 16.0.10417; Pro)'
            ClientIPAddress = '172.16.111.56'
            AuthenticatedUser = 'jeffrey.roger@arcticslope.org'
            AnchorMailbox = 'MailboxId=7f000a6b-7628-4c1c-8f74-daa11b39b3ab@arcticslope.org'
            RemoteEndpoint = $null
            SessionId = 'abc'
            Event = $null
            Data = 'Mapi'
          }
        )
        Failures = @()
      }
    }

    $results = @(
      [pscustomobject]@{
        Sender = 'Jeffrey.Roger@arcticslope.org'
        Recipients = @('Susan.Miklavcic@arcticslope.org')
        MessageSubject = 'Resignation from medical Staff'
        MessageId = '<msg-01@example.org>'
        InternalMessageId = '101'
        EventId = 'SUBMIT'
        Timestamp = [datetime]'2026-04-06T19:00:15'
        ServerHostname = 'EXCH-SE-02'
        Source = 'SMTP'
        ClientHostname = 'EXCH-SE-02'
        ConnectorId = 'InboundProxy'
        SourceContext = 'MDB:01'
      }
    )

    $result = Invoke-ImtMessageClientAccessAudit `
      -RunContext $runContext `
      -Results $results `
      -CandidateMailboxAddresses @('jeffrey.roger@arcticslope.org') `
      -Servers @('EXCH-SE-01', 'EXCH-SE-02')

    $result.Status | Should Be 'OK'
    @($result.Data.Rows).Count | Should Be 1
    @($result.Data.Rows)[0].AttributionSource | Should Be 'ProtocolLog'
    @($result.Data.Rows)[0].LikelyClient | Should Be 'Outlook desktop'
    @($result.Data.Rows)[0].ClientIPAddress | Should Be '172.16.111.56'
    @($result.Data.Rows)[0].ProtocolEvidenceType | Should Be 'HttpProxyMapi'
  }

  It 'uses tracking source context when mailbox audit and protocol logs are unavailable' {
    $runContext = [pscustomobject]@{
      Start = [datetime]'2026-04-06T15:00:00'
      End = [datetime]'2026-04-06T22:00:00'
    }

    Mock Resolve-ImtMailboxByAddress {
      [pscustomobject]@{
        Identity = 'jroger'
        PrimarySmtpAddress = 'jeffrey.roger@arcticslope.org'
        ExchangeGuid = '1320642c-99cb-4a3f-bd0c-240948aebd03'
      }
    }

    Mock Search-MailboxAuditLog { @() }

    Mock Get-ImtProtocolEvidenceRowsForSenders {
      [pscustomobject]@{
        Rows = @()
        Failures = @()
      }
    }

    $results = @(
      [pscustomobject]@{
        Sender = 'Jeffrey.Roger@arcticslope.org'
        Recipients = @('Susan.Miklavcic@arcticslope.org')
        MessageSubject = 'Resignation from medical Staff'
        MessageId = '<msg-01@example.org>'
        InternalMessageId = '101'
        EventId = 'SUBMIT'
        Timestamp = [datetime]'2026-04-06T19:00:15'
        ServerHostname = 'EXCH-SE-03.asna.alaska.ihs.gov'
        Source = 'STOREDRIVER'
        ClientHostname = 'EXCH-SE-02'
        ClientIp = '172.16.2.16'
        ConnectorId = $null
        SourceContext = 'MDB:f4f3423e-3ce8-4dd2-a4de-5a5b79f23b63, Mailbox:1320642c-99cb-4a3f-bd0c-240948aebd03, Event:28662855, MessageClass:IPM.Note, CreationTime:2026-04-07T03:00:15.485Z, ClientType:AirSync, SubmissionAssistant:MailboxTransportSubmissionEmailAssistant'
      }
    )

    $result = Invoke-ImtMessageClientAccessAudit `
      -RunContext $runContext `
      -Results $results `
      -CandidateMailboxAddresses @('jeffrey.roger@arcticslope.org') `
      -Servers @('EXCH-SE-01', 'EXCH-SE-02', 'EXCH-SE-03')

    $result.Status | Should Be 'OK'
    @($result.Data.Rows).Count | Should Be 1
    @($result.Data.Rows)[0].AttributionSource | Should Be 'TrackingSourceContext'
    @($result.Data.Rows)[0].AttributionConfidence | Should Be 'Medium'
    @($result.Data.Rows)[0].LikelyClient | Should Be 'Mobile client via ActiveSync'
    @($result.Data.Rows)[0].ClientInfoString | Should Match 'ClientType=AirSync'
    @($result.Data.Rows)[0].ClientIPAddress | Should Be '172.16.2.16'
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

  It 'uses the same integer result size when proxy metadata reports object' {
    function Search-MailboxAuditLog {
      param(
        [string]$Identity,
        [string[]]$LogonTypes,
        [switch]$ShowDetails,
        [datetime]$StartDate,
        [datetime]$EndDate,
        [object]$ResultSize
      )
    }

    $params = Get-ImtMailboxAuditQueryParameters `
      -Identity 'jproger@arcticslope.org' `
      -StartDate ([datetime]'2026-04-06T15:00:00') `
      -EndDate ([datetime]'2026-04-06T22:00:00')

    $params.ResultSize | Should Be 250000
  }
}

Describe 'Get-ImtProtocolCandidateFiles' {
  BeforeEach {
    Mock Test-ImtLiteralPathExists { $true }
  }

  It 'queries token-filtered file sets and de-duplicates overlapping results' {
    Mock Get-ChildItem {
      param($LiteralPath, $File, $Filter)

      switch ($Filter) {
        '*2026040619*' {
          @([pscustomobject]@{ FullName = 'C:\Logs\HttpProxy_2026040619-1.LOG' })
          break
        }
        '*2026040620*' {
          @(
            [pscustomobject]@{ FullName = 'C:\Logs\HttpProxy_2026040619-1.LOG' }
            [pscustomobject]@{ FullName = 'C:\Logs\HttpProxy_2026040620-1.LOG' }
          )
          break
        }
        default { @() }
      }
    }

    $files = @(Get-ImtProtocolCandidateFiles -DirectoryPath 'C:\Logs' -Tokens @('2026040619', '2026040620'))

    $files.Count | Should Be 2
    $files[0] | Should Be 'C:\Logs\HttpProxy_2026040619-1.LOG'
    $files[1] | Should Be 'C:\Logs\HttpProxy_2026040620-1.LOG'
    Assert-MockCalled Get-ChildItem -Times 1 -ParameterFilter { $Filter -eq '*2026040619*' }
    Assert-MockCalled Get-ChildItem -Times 1 -ParameterFilter { $Filter -eq '*2026040620*' }
  }
}

Describe 'Get-ImtProtocolWindowTokens' {
  It 'includes both local and UTC hour tokens for evening windows that cross UTC midnight' {
    $start = [datetime]'2026-04-06T15:00:00'
    $end = [datetime]'2026-04-06T22:00:00'

    $tokens = @(Get-ImtProtocolWindowTokens -StartDate $start -EndDate $end -Granularity Hour)

    @($tokens | Where-Object { $_ -eq '2026040615' }).Count | Should BeGreaterThan 0
    @($tokens | Where-Object { $_ -eq '2026040623' }).Count | Should BeGreaterThan 0
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
