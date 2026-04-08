Set-StrictMode -Version Latest

$repoRoot = Split-Path -Path $PSScriptRoot -Parent
. (Join-Path $repoRoot 'src\\Models\\New-ResultObjects.ps1')
. (Join-Path $repoRoot 'src\\Logging\\Write-ImtLog.ps1')
. (Join-Path $repoRoot 'src\\Core\\Initialize-RunContext.ps1')
. (Join-Path $repoRoot 'src\\Identity\\Resolve-Participants.ps1')

function Get-Recipient { }
function Get-Mailbox { }

Describe 'Resolve-ImtParticipantsAndSenders' {
  BeforeEach {
    Mock Write-ImtLog {}
  }

  It 'handles a single explicit sender without scalar Count failures' {
    $runContext = Initialize-ImtRunContext `
      -SenderAddress 'jproger@arcticslope.org' `
      -StartDate '2026-04-06T15:00:00' `
      -EndDate '2026-04-06T22:00:00' `
      -OutputDir ([System.IO.Path]::GetTempPath()) `
      -OutputLevel INFO

    Mock Get-Recipient {
      param(
        [string]$Identity,
        [string]$Anr,
        [string]$Filter,
        [int]$ResultSize
      )

      if ($Identity) {
        return [pscustomobject]@{
          PrimarySmtpAddress = $Identity
          DisplayName = $Identity
        }
      }
    }

    $result = Resolve-ImtParticipantsAndSenders -RunContext $runContext

    $result.Status | Should Be 'OK'
    @($result.Data.EffectiveSenderFilters).Count | Should Be 1
    $result.Data.EffectiveSenderFilters[0] | Should Be 'jproger@arcticslope.org'
    @($result.Data.BaseTargetAddresses).Count | Should Be 1
    $result.Data.BaseTargetAddresses[0] | Should Be 'jproger@arcticslope.org'
  }
}
