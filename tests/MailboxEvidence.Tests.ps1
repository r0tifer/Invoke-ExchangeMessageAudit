Set-StrictMode -Version Latest

$repoRoot = Split-Path -Path $PSScriptRoot -Parent
. (Join-Path $repoRoot 'src\Models\New-ResultObjects.ps1')
. (Join-Path $repoRoot 'src\Reporting\Export-MailboxEvidenceReports.ps1')

Describe 'Export-ImtMailboxEvidenceReports' {
  It 'marks transport correlation when message ids match' {
    $tempDir = Join-Path $env:TEMP ("imt-evidence-tests-{0}" -f ([guid]::NewGuid().ToString('N')))
    New-Item -ItemType Directory -Path $tempDir | Out-Null

    try {
      $runContext = [pscustomobject]@{
        OutputDir = $tempDir
        Timestamp = '20260313_120000'
        Inputs = [pscustomobject]@{
          DetailedMailboxEvidence = $true
        }
      }

      $evidenceRows = @(
        [pscustomobject]@{
          SourceMailbox = 'rachel.aumavae@example.org'
          MailboxLocation = 'Primary'
          SentTime = [datetime]'2025-01-15T18:00:00Z'
          From = 'rachel.aumavae@example.org'
          To = 'riveracarolyn929@gmail.com'
          Cc = $null
          Subject = 'Eligibility letter'
          InternetMessageId = '<msg-01@example.org>'
          HasAttachments = $true
          AttachmentCount = 1
          ItemSize = 1024
          EvidenceFolder = 'IMT_Evidence_20260313_120000\rachel.aumavae@example.org'
        }
      )

      $trackingResults = @(
        [pscustomobject]@{
          Sender = 'rachel.aumavae@example.org'
          Recipients = @('riveracarolyn929@gmail.com')
          Timestamp = [datetime]'2025-01-15T18:03:00Z'
          MessageId = '<msg-01@example.org>'
        }
      )

      $result = Export-ImtMailboxEvidenceReports -RunContext $runContext -EvidenceRows $evidenceRows -TrackingResults $trackingResults

      $result.Status | Should Be 'OK'
      $result.Data.EvidenceRows[0].TransportCorrelated | Should Be $true
      $result.Data.EvidenceRows[0].TrackingMessageId | Should Be '<msg-01@example.org>'
      Test-Path -LiteralPath $result.Data.EvidenceCsv | Should Be $true
    } finally {
      Remove-Item -LiteralPath $tempDir -Recurse -Force -ErrorAction SilentlyContinue
    }
  }

  It 'returns WARN when evidence exists but tracking results are unavailable' {
    $tempDir = Join-Path $env:TEMP ("imt-evidence-tests-{0}" -f ([guid]::NewGuid().ToString('N')))
    New-Item -ItemType Directory -Path $tempDir | Out-Null

    try {
      $runContext = [pscustomobject]@{
        OutputDir = $tempDir
        Timestamp = '20260313_120001'
        Inputs = [pscustomobject]@{
          DetailedMailboxEvidence = $true
        }
      }

      $evidenceRows = @(
        [pscustomobject]@{
          SourceMailbox = 'joshua.stein@example.org'
          MailboxLocation = 'Primary'
          SentTime = [datetime]'2025-02-10T17:30:00Z'
          From = 'joshua.stein@example.org'
          To = 'laserino77@gmail.com'
          Cc = $null
          Subject = 'Eligibility letter'
          InternetMessageId = '<msg-02@example.org>'
          HasAttachments = $true
          AttachmentCount = 1
          ItemSize = 2048
          EvidenceFolder = 'IMT_Evidence_20260313_120001\joshua.stein@example.org'
        }
      )

      $result = Export-ImtMailboxEvidenceReports -RunContext $runContext -EvidenceRows $evidenceRows -TrackingResults @()

      $result.Status | Should Be 'WARN'
      $result.Errors[0] | Should Match 'Transport correlation skipped'
    } finally {
      Remove-Item -LiteralPath $tempDir -Recurse -Force -ErrorAction SilentlyContinue
    }
  }
}
