Set-StrictMode -Version Latest

$repoRoot = Split-Path -Path $PSScriptRoot -Parent
. (Join-Path $repoRoot 'src\\Models\\New-ResultObjects.ps1')
. (Join-Path $repoRoot 'src\\Logging\\Write-ImtLog.ps1')
. (Join-Path $repoRoot 'src\\Tracking\\Invoke-MessageTrackingAudit.ps1')

# Define Exchange cmdlet placeholder so Pester 3 can mock it.
function Get-MessageTrackingLog { }

Describe 'Invoke-ImtMessageTrackingAudit' {
  BeforeEach {
    Mock Write-ImtLog {}
  }

  It 'collects participant query results when mixed query calls return single objects and nulls' {
    $runContext = [pscustomobject]@{
      ParticipantMode = $true
      RecipientMode = $false
      Start = [datetime]'2026-02-01T00:00:00Z'
      End = [datetime]'2026-02-02T00:00:00Z'
      Inputs = [pscustomobject]@{
        SubjectLike = $null
        Keywords = @()
        OnlyProblems = $false
        Recipient = $null
      }
    }

    Mock Get-MessageTrackingLog {
      param(
        [string]$Server,
        [datetime]$Start,
        [datetime]$End,
        $ResultSize,
        [string]$Sender,
        [string]$Recipients
      )

      if ($Sender -eq 'dakota.miller@arcticslope.org') {
        return [pscustomobject]@{
          ServerHostname = $Server
          InternalMessageId = '101'
          EventId = 'SEND'
          Timestamp = [datetime]'2026-02-01T12:00:00Z'
          MessageId = '<id-101@example.org>'
          Sender = 'dakota.miller@arcticslope.org'
          Recipients = @('jalaya.duarte@arcticslope.org')
          MessageSubject = 'eligibility grant review'
          RecipientStatus = @()
        }
      }

      if ($Recipients -eq 'jalaya.duarte@arcticslope.org') {
        return [pscustomobject]@{
          ServerHostname = $Server
          InternalMessageId = '102'
          EventId = 'RECEIVE'
          Timestamp = [datetime]'2026-02-01T12:01:00Z'
          MessageId = '<id-102@example.org>'
          Sender = 'external.sender@example.org'
          Recipients = @('jalaya.duarte@arcticslope.org')
          MessageSubject = 'provider audit follow-up'
          RecipientStatus = @()
        }
      }

      # No output when no rows match (mirrors Get-MessageTrackingLog behavior).
    }

    $result = Invoke-ImtMessageTrackingAudit `
      -RunContext $runContext `
      -Servers @('EXCH-SE-01') `
      -VersionInfo @{ 'EXCH-SE-01' = 'Exchange 2019' } `
      -TraceParticipants @('dakota.miller@arcticslope.org', 'jalaya.duarte@arcticslope.org') `
      -EffectiveSenderFilters @()

    $result.Status | Should Be 'OK'
    @($result.Data.Results).Count | Should Be 2
  }
}
