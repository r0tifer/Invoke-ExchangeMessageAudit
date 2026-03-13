# Invoke-ExchangeMessageAudit

A modular Exchange message tracing and audit workflow for Exchange 2013 / 2016 / 2019.

Built because doing this manually over and over is painful.

---

## Why This Exists

If you've ever had to:

- Prove who sent what and when
- Trace a message across transport hops
- Validate export readiness before legal comes knocking
- Run keyword checks before creating PST exports
- Or just figure out why Exchange did something weird

…you already know it’s never just one command.

This project ties those moving parts together into a single repeatable workflow. Not magic. Just structure, logging, and less guesswork.

---

## What This Actually Does

This is not just a message tracking wrapper.

It orchestrates:

- Identity resolution (participants, senders, mailbox validation)
- Transport topology discovery
- Message tracking queries with filtering
- Keyword-based audit logic
- Optional mailbox search estimates
- Optional export preflight checks
- Optional mailbox export request creation
- Message trail tracing by MessageId
- Structured CSV artifacts and run logs

It produces outputs you can hand to:

- Compliance
- Legal
- Leadership
- Or the next poor soul who inherits your ticket

---

## Design Philosophy

I kept this modular on purpose.

- One orchestrator script runs the show
- Each module in src/ has a single job
- Logging is centralized
- Output verbosity is controlled
- Artifacts are structured and predictable

You can follow the flow. You can debug it. You can extend it without everything exploding.

---

## Common Use Cases

- Delivery investigations
- Legal discovery prep
- Compliance keyword audits
- Retention visibility checks
- Export readiness validation
- Repeatable audit documentation

If you're tired of cobbling together 6 cmdlets every time, this helps.

---

## Output Artifacts

Every run generates structured artifacts. Not just console noise.

Depending on options used, you’ll get:

- Primary message tracking CSV
- Keyword summaries (overall + per mailbox)
- Direct mailbox search exports
- Retention snapshot CSV
- Mailbox export request summary
- Full message trail trace CSV
- Step log (MTL_Steps_*.log)
- Optional transcript log

Even if you tone down console output, the logs are still there.

---

## Output Levels

Control verbosity with `-OutputLevel`.

- INFO – Normal step progress + summary
- DEBUG – Detailed diagnostics
- WARN – Warnings and summary
- ERROR – Errors and summary
- CRITICAL – Fatal-only output

Logs still capture everything regardless.

---

## Available Options

| Option | Type | What It Does |
| --- | --- | --- |
| `-Participants` | `string[]` | Mailboxes/users to trace as senders or recipients. |
| `-Recipient` | `string` | Target recipient address for recipient-focused tracking. |
| `-Recipients` | `string[]` | Multiple target recipient addresses for recipient-focused tracking and mailbox audit searches. |
| `-SenderAddress` (`-Sender`) | `string` | Single sender filter. |
| `-Senders` (`-SenderList`) | `string[]` | Multiple sender filters. |
| `-SourceMailboxes` | `string[]` | Explicit mailbox scope for mailbox searches/exports when the sender mailbox list is known. |
| `-DaysBack` | `int` | Relative lookback window when explicit dates are not provided. Default: `90`. |
| `-StartDate` | `datetime` | Start of date range (must be used with `-EndDate`). |
| `-EndDate` | `datetime` | End of date range (must be used with `-StartDate`). |
| `-OutputDir` | `string` | Output artifact directory. Default: `C:\Temp`. |
| `-LogDir` | `string` | Optional log directory for step log and transcript. Defaults to `-OutputDir`. |
| `-SubjectLike` | `string` | Subject contains filter. |
| `-Keywords` | `string[]` | Keywords for tracking and direct mailbox keyword summaries. |
| `-HasAttachmentOnly` | `switch` | Restrict tracking/search/export logic to messages with attachments. |
| `-OnlyProblems` | `switch` | Keep only problematic transport events (fail/defer style events/statuses). |
| `-TraceMessageId` | `string` | Explicit Message-Id to trail trace. |
| `-TraceLatest` | `switch` | Automatically choose latest tracked Message-Id for trail tracing. |
| `-SkipRetentionCheck` | `switch` | Skip retention snapshot and retention export steps. |
| `-PromptForMailboxExport` | `switch` | Prompt interactively to create mailbox export requests. |
| `-ExportLocatedEmails` | `switch` | Automatically create mailbox export requests (no prompt). |
| `-ExportPstRoot` | `string` | UNC root path for PST exports and preflight checks. |
| `-IncludeArchive` | `switch` | Include archive mailbox export requests in addition to primary mailbox. |
| `-SkipDagPathValidation` | `switch` | Skip mailbox-server remote path validation during export preflight. |
| `-PreflightOnly` | `switch` | Run identity/transport/preflight checks only and skip tracking/export/search steps. |
| `-SearchAllMailboxes` | `switch` | Search all local user/shared mailboxes for mailbox-audit workflows. |
| `-SearchMailboxesDirectly` | `switch` | Force direct mailbox estimate search step even without keywords. |
| `-OutboundOnly` | `switch` | Restrict mailbox search/export logic to sent-item time windows only. |
| `-DetailedMailboxEvidence` | `switch` | Copy matching items to an evidence mailbox and produce a consolidated message-level evidence CSV. |
| `-EvidenceMailbox` | `string` | Target mailbox used to hold copied evidence items for `-DetailedMailboxEvidence`. |
| `-DisableTranscriptLog` | `switch` | Disable transcript logging (step logging behavior remains as implemented by logger settings). |
| `-SearchDumpsterDirectly` | `switch` | Include dumpster when running direct mailbox estimate queries. |
| `-ExpandExportScopeFromMatchedTraffic` | `switch` | Add matched sender/recipient traffic addresses to mailbox export target scope. |
| `-OutputLevel` | `string` | Console verbosity: `DEBUG`, `INFO`, `WARN`, `ERROR`, `CRITICAL`. Default: `INFO`. |

---

## Repo Layout

```
Invoke-ExchangeMessageAudit.psd1  # Module manifest
Invoke-ExchangeMessageAudit.psm1  # Module root (loads src/)
Invoke-ExchangeMessageAudit.ps1   # Compatibility launcher script

src/
  Orchestration/   # Public module entrypoint
  Core/            # Context + validation
  Logging/         # Logging engine
  Identity/        # Mailbox + participant resolution
  Exchange/        # Topology + retention snapshot
  Tracking/        # Message tracking + trail tracing
  Export/          # Export preflight + PST creation
  MailboxSearch/   # Direct mailbox estimate logic
  Reporting/       # CSV + summary composition
  Models/          # Shared result objects

scripts/           # Utility scripts
tests/             # Pester test suite
```

Don’t move folders around unless you update the module root. It dot-sources `src/` during import.

---

## Requirements

- Windows PowerShell 5.1 or PowerShell 7+
- Exchange Management Shell context (or Exchange cmdlets loaded)
- Proper RBAC permissions for:
  - Get-MessageTrackingLog
  - Get-Recipient
  - Get-Mailbox
  - Search-Mailbox (if used)
  - New-MailboxExportRequest (if used)
- Valid UNC path and permissions for PST exports

If Exchange cmdlets aren’t available, certain features will skip or fail. That’s expected.

---

## Getting Started

You’ve got two ways to run this:

- Clone it and run directly (quick + dirty)
- Install it like a proper PowerShell module (recommended)

---

## Example Commands

Basic participant trace with keyword filtering:

```powershell
Invoke-ExchangeMessageAudit `
  -Participants "alex.rivera@example.org","jamie.chen@example.org" `
  -DaysBack 30 `
  -Keywords "grant","audit" `
  -OutputDir "C:\Temp\ExchangeAudit" `
  -OutputLevel INFO
```

Explicit date range with sender + recipient filtering:

```powershell
Invoke-ExchangeMessageAudit `
  -Recipient "jamie.chen@example.org" `
  -Senders "alex.rivera@example.org","notifications@example.org" `
  -StartDate "2026-01-01 00:00:00" `
  -EndDate "2026-01-31 23:59:59" `
  -OnlyProblems `
  -OutputDir "C:\Temp\ExchangeAudit" `
  -OutputLevel INFO
```

Preflight-only export validation:

```powershell
Invoke-ExchangeMessageAudit `
  -Participants "alex.rivera@example.org","jamie.chen@example.org" `
  -PreflightOnly `
  -ExportPstRoot "\\fileserver01\PSTExports" `
  -OutputDir "C:\Temp\ExchangeAudit" `
  -OutputLevel DEBUG
```

Tracking + mailbox export (non-interactive):

```powershell
Invoke-ExchangeMessageAudit `
  -Participants "alex.rivera@example.org","jamie.chen@example.org" `
  -StartDate "2025-10-01 00:00:00" `
  -EndDate "2026-02-28 23:59:59" `
  -Keywords "eligibility","provider","childcare" `
  -HasAttachmentOnly `
  -ExportLocatedEmails `
  -ExportPstRoot "\\fileserver01\PSTExports" `
  -OutputDir "C:\Temp\ExchangeAudit" `
  -OutputLevel INFO
```

Org-wide outbound attachment audit with detailed mailbox evidence:

```powershell
Invoke-ExchangeMessageAudit `
  -Recipients "Riveracarolyn929@gmail.com","laserino77@gmail.com","manaluuraq@live.com","libertyfrances1@icloud.com" `
  -SourceMailboxes "Rachel Aumavae","Joshua Stein" `
  -StartDate "2024-10-01 00:00:00" `
  -EndDate "2025-09-30 23:59:59" `
  -HasAttachmentOnly `
  -OutboundOnly `
  -DetailedMailboxEvidence `
  -EvidenceMailbox "Discovery Search Mailbox" `
  -OutputDir "C:\Temp\ExchangeAudit" `
  -OutputLevel INFO
```

Trail trace by explicit Message-Id:

```powershell
Invoke-ExchangeMessageAudit `
  -Participants "alex.rivera@example.org" `
  -DaysBack 14 `
  -TraceMessageId "<f4b7d62d-67c1-4fb8-b955-0fc9e2adf98b@example.org>" `
  -OutputDir "C:\Temp\ExchangeAudit" `
  -OutputLevel INFO
```

---

### Option 1 – Clone and Run (Quick Method)

```powershell
git clone https://github.com/r0tifer/Invoke-ExchangeMessageAudit
cd Invoke-ExchangeMessageAudit
```

Unblock scripts:

```powershell
Get-ChildItem -Recurse -Filter *.ps1 | Unblock-File
```

Temporarily allow execution for this session:

```powershell
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
```

Run it:

```powershell
.\Invoke-ExchangeMessageAudit.ps1 -Participants "alex.rivera@example.org"
```

That works fine. But if you plan to use it more than once… install it properly.

---

## Option 2 – Install as a PowerShell Module (Recommended)

### Install for Current User (No Admin Required)

Install into your personal module path:

```powershell
$modulePath = Join-Path $HOME "Documents\PowerShell\Modules\Invoke-ExchangeMessageAudit"
git clone https://github.com/r0tifer/Invoke-ExchangeMessageAudit $modulePath
```

If you're on Windows PowerShell 5.1:

```powershell
$modulePath = Join-Path $HOME "Documents\WindowsPowerShell\Modules\Invoke-ExchangeMessageAudit"
```

Then:

```powershell
Import-Module Invoke-ExchangeMessageAudit -Force
Get-Command Invoke-ExchangeMessageAudit
```

Now you can call it from anywhere:

```powershell
Invoke-ExchangeMessageAudit
```

---

### Install System-Wide (All Users)

Requires admin rights.

```powershell
$modulePath = "C:\Program Files\PowerShell\Modules\Invoke-ExchangeMessageAudit"
git clone https://github.com/r0tifer/Invoke-ExchangeMessageAudit $modulePath
```

If you're on Windows PowerShell 5.1:

```powershell
$modulePath = "C:\Program Files\WindowsPowerShell\Modules\Invoke-ExchangeMessageAudit"
git clone https://github.com/r0tifer/Invoke-ExchangeMessageAudit $modulePath
```

Then:

```powershell
Import-Module Invoke-ExchangeMessageAudit -Force
Get-Command Invoke-ExchangeMessageAudit
```

Restart PowerShell and verify:

```powershell
Get-Module -ListAvailable Invoke-ExchangeMessageAudit
```

Then run it normally:

```powershell
Invoke-ExchangeMessageAudit
```

---

## Updating the Module

```powershell
cd <module install path>
git pull
```

Reload:

```powershell
Import-Module Invoke-ExchangeMessageAudit -Force
```

---

## Testing

Run Pester tests:

```powershell
Invoke-Pester -Path .\tests
```

Run approved verb validation:

```powershell
Import-Module PSScriptAnalyzer -Force
.\scripts\Test-ApprovedVerbs.ps1
```

GitHub Actions enforces approved PowerShell verbs in CI.

---

## Contributing

If you want to improve this:

- Keep modules single-purpose
- Keep logging centralized
- Don't mix output formatting into core logic
- Add tests

PRs welcome. Just don’t turn it into a 4,000 line monolith.

---

## License

This project is licensed under the MIT License. See `LICENSE`.

---

## Final Notes

This was built for real-world Exchange investigations. It’s not flashy. It’s not SaaS. It’s just structured automation around things admins already do — but in a way thats repeatable and defensible.

If it saves you an hour on a compliance ticket, it did its job.
