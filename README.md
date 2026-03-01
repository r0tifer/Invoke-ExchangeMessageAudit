# Invoke-ExchangeMessageAudit

Modular Exchange mail-tracing and audit orchestration for Exchange 2013/2016/2019.

## Purpose

`Invoke-ExchangeMessageAudit` is a PowerShell-based audit tool that helps Exchange administrators trace message flow, validate export readiness, run mailbox-level search estimates, and produce evidence-ready CSV/log artifacts.

The project is intentionally modular:

- One orchestrator script runs the workflow end-to-end.
- Each subscript in `src/` has a single authoritative role.
- Logging and step reporting are centralized and output-level controlled.

## Intent

This project is designed for operational audit and investigation use cases:

- Message delivery investigations (who sent/received, when, and how).
- Compliance and legal prep (keyword-based trace and mailbox export request creation).
- Environment health checks for retention visibility and export prerequisites.
- Repeatable run logging with durable artifacts for review and handoff.

## Core Functionality

- Identity resolution for participants and sender filters.
- Transport topology discovery across Exchange transport services.
- Message tracking query and filtering (date window, participants, subject, keywords, failures).
- Optional retention snapshot collection and export.
- Optional mailbox export preflight and export request creation.
- Optional direct mailbox search estimates (`Search-Mailbox`).
- Optional combined keyword summaries (transport + mailbox estimates).
- Optional full message trail tracing by `MessageId` or latest result.
- Final run summary with step statuses and timing.

## Auditing Capabilities

The module generates auditable outputs for evidence and troubleshooting:

- Main message tracking CSV export.
- Keyword hit summaries (overall and by mailbox).
- Direct mailbox search summary and keyword hit exports.
- Retention snapshot CSV.
- Mailbox export request summary CSV.
- Message trail CSV (full-hop trace).
- Step log (`MTL_Steps_*.log`) with structured event records.
- Optional transcript log (`MTL_RunTranscript_*.log`).

## Output Levels

Use `-OutputLevel` to control terminal verbosity:

- `INFO`: step start, step result, final summary.
- `DEBUG`: includes detailed progress diagnostics.
- `WARN`: warnings/failures plus final summary.
- `ERROR`: errors/failures plus final summary.
- `CRITICAL`: fatal-only console output plus minimal summary.

## Repository Layout

- `Invoke-ExchangeMessageAudit.ps1`: Orchestrator entrypoint.
- `src/Core`: Run context and input validation.
- `src/Logging`: Central logging system.
- `src/Identity`: Participant/sender/mailbox resolution.
- `src/Exchange`: Topology and retention snapshot logic.
- `src/Tracking`: Message tracking audit and trail trace.
- `src/Export`: Export preflight and mailbox export request creation.
- `src/MailboxSearch`: Direct mailbox estimate logic.
- `src/Reporting`: CSV/report composition and final summary.
- `src/Models`: Shared result object contracts.
- `scripts`: Utility scripts (for example, approved verb guard).
- `tests`: Pester test suite.

## Prerequisites

- Windows PowerShell 5.1 or PowerShell 7+.
- Exchange Management Shell context for Exchange cmdlets.
- RBAC/permissions appropriate for:
  - `Get-MessageTrackingLog`
  - `Get-Recipient` / `Get-Mailbox`
  - `New-MailboxExportRequest` (when export is used)
  - `Search-Mailbox` (when direct mailbox search is used)
- UNC export path and server/share permissions when using export features.

## Pull Down Full Repository and Subcomponents

```powershell
git clone https://github.com/r0tifer/Invoke-ExchangeMessageAudit
cd Invoke-ExchangeMessageAudit

```

If your remote uses a different folder name, `cd` into that folder instead.

## Usage

1. Open Exchange Management Shell (recommended) or PowerShell with Exchange snap-ins/modules available.
2. Navigate to the cloned repository root.
3. Unblock downloaded scripts.

```powershell
Get-ChildItem -Recurse -Filter *.ps1 | Unblock-File
```

4. Set execution policy for current process only.

```powershell
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
```

5. Validate that the orchestrator and modules are present.

```powershell
Get-ChildItem .\src -Recurse -Filter *.ps1
```

6. Run usage/help output.

```powershell
.\Invoke-ExchangeMessageAudit.ps1
```

## Quick Start

Example trace with participants and keyword filters:

```powershell
.\Invoke-ExchangeMessageAudit.ps1 `
  -Participants 'user1@contoso.org','user2@contoso.org' `
  -StartDate '2025-01-01 00:00:00' `
  -EndDate '2025-01-31 23:59:59' `
  -Keywords 'audit','invoice' `
  -OutputDir 'C:\Temp' `
  -OutputLevel INFO
```

Preflight-only export readiness check:

```powershell
.\Invoke-ExchangeMessageAudit.ps1 `
  -Participants 'user1@contoso.org' `
  -ExportPstRoot '\\fileserver\PSTExports' `
  -PreflightOnly `
  -OutputLevel DEBUG
```

## Validation and CI Guard

Run tests locally:

```powershell
Invoke-Pester -Path .\tests
```

Run approved-verb guard locally:

```powershell
Import-Module PSScriptAnalyzer -Force
.\scripts\Test-ApprovedVerbs.ps1
```

GitHub Actions includes a CI workflow that enforces approved PowerShell verbs (`PSUseApprovedVerbs`).

## Notes

- `Invoke-ExchangeMessageAudit.ps1` dot-sources all subcomponents in `src/`.
- Keep file structure intact; moving module files without updating the orchestrator will break execution.
- Some features are environment-dependent and will skip or fail if Exchange cmdlets are unavailable.

