Set-StrictMode -Version Latest

$repoRoot = Split-Path -Path $PSScriptRoot -Parent
. (Join-Path $repoRoot 'src\Models\New-ResultObjects.ps1')
. (Join-Path $repoRoot 'src\Logging\Write-ImtLog.ps1')
. (Join-Path $repoRoot 'src\Identity\Resolve-Participants.ps1')
. (Join-Path $repoRoot 'src\MailboxSearch\Invoke-DirectMailboxSearch.ps1')

function Get-Mailbox { }
function Get-Recipient { }
function Search-Mailbox { }

Describe 'New-ImtMailboxSearchQuery' {
  It 'builds outbound recipient-aware attachment query' {
    $query = New-ImtMailboxSearchQuery `
      -StartAt ([datetime]'2024-10-01T00:00:00') `
      -EndAt ([datetime]'2025-09-30T23:59:59') `
      -SubjectContains 'eligibility letter' `
      -SubjectKeywords @('provider') `
      -SenderFilters @('rachel.aumavae@example.org') `
      -RecipientFilters @('riveracarolyn929@gmail.com', 'laserino77@gmail.com') `
      -OutboundOnly `
      -RequireAttachment

    $query | Should Match 'sent:10/01/2024\.\.09/30/2025'
    $query | Should Match 'from:'
    $query | Should Match 'to:'
    $query | Should Match 'cc:'
    $query | Should Match 'hasattachment:true'
  }
}

Describe 'New-ImtExportContentFilter' {
  BeforeAll {
    . (Join-Path $repoRoot 'src\Export\Invoke-MailboxExportRequests.ps1')
  }

  It 'builds recipient filters against both To and Cc' {
    $filter = New-ImtExportContentFilter `
      -StartAt ([datetime]'2024-10-01T00:00:00') `
      -EndAt ([datetime]'2025-09-30T23:59:59') `
      -SenderFilters @('rachel.aumavae@example.org') `
      -RecipientFilters @('riveracarolyn929@gmail.com') `
      -RequireAttachment

    $filter | Should Match "From -eq 'rachel\.aumavae@example\.org'"
    $filter | Should Match "To -like '\*riveracarolyn929@gmail\.com\*'"
    $filter | Should Match "Cc -like '\*riveracarolyn929@gmail\.com\*'"
    $filter | Should Match 'HasAttachment -eq \$true'
  }
}

Describe 'Resolve-ImtMailboxAuditScope' {
  It 'filters org-wide scope down to user and shared mailboxes' {
    $runContext = [pscustomobject]@{
      Inputs = [pscustomobject]@{
        SearchAllMailboxes = $true
        SourceMailboxes = @()
      }
    }

    Mock Get-Mailbox {
      @(
        [pscustomobject]@{
          Identity = 'user01'
          DistinguishedName = 'dn-user01'
          PrimarySmtpAddress = 'user01@example.org'
          RecipientTypeDetails = 'UserMailbox'
        }
        [pscustomobject]@{
          Identity = 'shared01'
          DistinguishedName = 'dn-shared01'
          PrimarySmtpAddress = 'shared01@example.org'
          RecipientTypeDetails = 'SharedMailbox'
        }
        [pscustomobject]@{
          Identity = 'audit01'
          DistinguishedName = 'dn-audit01'
          PrimarySmtpAddress = 'audit01@example.org'
          RecipientTypeDetails = 'AuditLogMailbox'
        }
      )
    }

    $result = Resolve-ImtMailboxAuditScope -RunContext $runContext -BaseTargetAddresses @()

    @($result.Mailboxes).Count | Should Be 2
    @($result.ScopeRows | Where-Object { $_.Included }).Count | Should Be 2
    @($result.ScopeRows | Where-Object { -not $_.Included }).Count | Should Be 1
  }

  It 'resolves explicit source mailboxes from display names' {
    $runContext = [pscustomobject]@{
      Inputs = [pscustomobject]@{
        SearchAllMailboxes = $false
        SourceMailboxes = @('Rachel Aumavae', 'Joshua Stein')
      }
    }

    Mock Get-Recipient {
      param(
        [string]$Identity,
        [string]$Anr,
        [string]$Filter,
        [int]$ResultSize
      )

      if ($Identity -eq 'Rachel Aumavae') {
        return [pscustomobject]@{
          PrimarySmtpAddress = 'rachel.aumavae@example.org'
          DisplayName = 'Rachel Aumavae'
        }
      }

      if ($Identity -eq 'Joshua Stein') {
        return [pscustomobject]@{
          PrimarySmtpAddress = 'joshua.stein@example.org'
          DisplayName = 'Joshua Stein'
        }
      }
    }

    Mock Get-Mailbox {
      param([string]$Identity)

      switch ($Identity) {
        'Rachel Aumavae' { throw 'not found by display name' }
        'Joshua Stein' { throw 'not found by display name' }
        'rachel.aumavae@example.org' {
          return [pscustomobject]@{
            Identity = 'rachel01'
            DistinguishedName = 'dn-rachel01'
            PrimarySmtpAddress = 'rachel.aumavae@example.org'
            RecipientTypeDetails = 'UserMailbox'
          }
        }
        'joshua.stein@example.org' {
          return [pscustomobject]@{
            Identity = 'joshua01'
            DistinguishedName = 'dn-joshua01'
            PrimarySmtpAddress = 'joshua.stein@example.org'
            RecipientTypeDetails = 'UserMailbox'
          }
        }
      }
    }

    $result = Resolve-ImtMailboxAuditScope -RunContext $runContext -BaseTargetAddresses @()

    @($result.Mailboxes).Count | Should Be 2
    @($result.ScopeRows | Where-Object { $_.Included }).Count | Should Be 2
    @($result.Unresolved).Count | Should Be 0
  }
}

Describe 'Invoke-ImtDirectMailboxSearch' {
  BeforeEach {
    Mock Write-ImtLog {}
  }

  It 'tracks matched source mailboxes for export targeting' {
    $runContext = [pscustomobject]@{
      Start = [datetime]'2024-10-01T00:00:00'
      End = [datetime]'2025-09-30T23:59:59'
      Timestamp = '20260313_120000'
      OutputDir = $env:TEMP
      Inputs = [pscustomobject]@{
        SubjectLike = $null
        Keywords = @()
        Recipients = @('riveracarolyn929@gmail.com')
        OutboundOnly = $true
        HasAttachmentOnly = $true
        SearchDumpsterDirectly = $false
        IncludeArchive = $false
        DetailedMailboxEvidence = $false
        SearchAllMailboxes = $true
        SourceMailboxes = @()
      }
    }

    Mock Get-Mailbox {
      @(
        [pscustomobject]@{
          Identity = 'user01'
          DistinguishedName = 'dn-user01'
          PrimarySmtpAddress = 'user01@example.org'
          RecipientTypeDetails = 'UserMailbox'
        }
        [pscustomobject]@{
          Identity = 'shared01'
          DistinguishedName = 'dn-shared01'
          PrimarySmtpAddress = 'shared01@example.org'
          RecipientTypeDetails = 'SharedMailbox'
        }
      )
    }

    Mock Search-Mailbox {
      param(
        $Identity,
        [string]$SearchQuery,
        [switch]$EstimateResultOnly
      )

      if ($EstimateResultOnly -and $Identity -eq 'user01' -and $SearchQuery -match 'riveracarolyn929@gmail.com') {
        return [pscustomobject]@{
          ResultItemsCount = 4
          ResultItemsSize = '12 KB'
        }
      }

      if ($EstimateResultOnly) {
        return [pscustomobject]@{
          ResultItemsCount = 0
          ResultItemsSize = '0 B'
        }
      }
    }

    $result = Invoke-ImtDirectMailboxSearch -RunContext $runContext -BaseTargetAddresses @() -EffectiveSenderFilters @()

    $result.Status | Should Be 'OK'
    @($result.Data.MatchedSourceMailboxAddresses).Count | Should Be 1
    $result.Data.MatchedSourceMailboxAddresses[0] | Should Be 'user01@example.org'
  }
}
