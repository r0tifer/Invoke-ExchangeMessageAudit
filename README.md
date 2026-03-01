# Invoke-ExchangeMessageAudit

A modular Exchange message tracing and audit workflow for Exchange 2013
/ 2016 / 2019.

Built because doing this manually over and over is painful.

------------------------------------------------------------------------

## Why This Exists

If you've ever had to:

-   Prove who sent what and when\
-   Trace a message across transport hops\
-   Validate export readiness before legal comes knocking\
-   Run keyword checks before creating PST exports\
-   Or just figure out why Exchange did something weird

...you already know it's never just one command.

This project ties those moving parts together into a single repeatable
workflow. Not magic. Just structure, logging, and less guesswork.

------------------------------------------------------------------------

## What This Actually Does

This is not just a message tracking wrapper.

It orchestrates:

-   Identity resolution (participants, senders, mailbox validation)
-   Transport topology discovery
-   Message tracking queries with filtering
-   Keyword-based audit logic
-   Optional mailbox search estimates
-   Optional export preflight checks
-   Optional mailbox export request creation
-   Message trail tracing by MessageId
-   Structured CSV artifacts and run logs

It produces outputs you can hand to:

-   Compliance
-   Legal
-   Leadership
-   Or the next poor soul who inherits your ticket

------------------------------------------------------------------------

## Design Philosophy

I kept this modular on purpose.

-   One orchestrator script runs the show.
-   Each module in src/ has a single job.
-   Logging is centralized.
-   Output verbosity is controlled.
-   Artifacts are structured and predictable.

You can follow the flow. You can debug it. You can extend it without
everything exploding.

------------------------------------------------------------------------

## Common Use Cases

-   Delivery investigations\
-   Legal discovery prep\
-   Compliance keyword audits\
-   Retention visibility checks\
-   Export readiness validation\
-   Repeatable audit documentation

If you're tired of cobbling together 6 cmdlets every time, this helps.

------------------------------------------------------------------------

## Output Artifacts

Every run generates structured artifacts. Not just console noise.

Depending on options used, you'll get:

-   Primary message tracking CSV\
-   Keyword summaries (overall + per mailbox)\
-   Direct mailbox search exports\
-   Retention snapshot CSV\
-   Mailbox export request summary\
-   Full message trail trace CSV\
-   Step log (MTL_Steps\_\*.log)\
-   Optional transcript log

Even if you tone down console output, the logs are still there.

------------------------------------------------------------------------

## Output Levels

Control verbosity with -OutputLevel.

-   INFO -- Normal step progress + summary\
-   DEBUG -- Detailed diagnostics\
-   WARN -- Warnings and summary\
-   ERROR -- Errors and summary\
-   CRITICAL -- Fatal-only output

Logs still capture everything regardless.

------------------------------------------------------------------------

## Repo Layout

Invoke-ExchangeMessageAudit.ps1 \# Main orchestrator

src/ Core/ \# Context + validation Logging/ \# Logging engine Identity/
\# Mailbox + participant resolution Exchange/ \# Topology + retention
snapshot Tracking/ \# Message tracking + trail tracing Export/ \# Export
preflight + PST creation MailboxSearch/ \# Direct mailbox estimate logic
Reporting/ \# CSV + summary composition Models/ \# Shared result objects

scripts/ \# Utility scripts tests/ \# Pester test suite

Don't move folders around unless you update the orchestrator. It
dot-sources everything.

------------------------------------------------------------------------

## Requirements

-   Windows PowerShell 5.1 or PowerShell 7+
-   Exchange Management Shell context (or Exchange cmdlets loaded)
-   Proper RBAC permissions for:
    -   Get-MessageTrackingLog
    -   Get-Recipient
    -   Get-Mailbox
    -   Search-Mailbox (if used)
    -   New-MailboxExportRequest (if used)
-   Valid UNC path and permissions for PST exports

If Exchange cmdlets aren't available, certain features will skip or
fail. That's expected.

------------------------------------------------------------------------

## Getting Started

You've got two ways to run this:

-   Clone it and run directly (quick + dirty)
-   Install it like a proper PowerShell module (recommended)

------------------------------------------------------------------------

### Option 1 -- Clone and Run (Quick Method)

git clone https://github.com/r0tifer/Invoke-ExchangeMessageAudit\
cd Invoke-ExchangeMessageAudit

Unblock scripts:

Get-ChildItem -Recurse -Filter \*.ps1 \| Unblock-File

Temporarily allow execution for this session:

Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass

Run it:

.`\Invoke`{=tex}-ExchangeMessageAudit.ps1

That works fine. But if you plan to use it more than once... install it
properly.

------------------------------------------------------------------------

## Option 2 -- Install as a PowerShell Module (Recommended)

### Install for Current User (No Admin Required)

Install into your personal module path:

\$modulePath = Join-Path \$HOME
"Documents`\PowerShell`{=tex}`\Modules`{=tex}`\Invoke`{=tex}-ExchangeMessageAudit"\
git clone https://github.com/r0tifer/Invoke-ExchangeMessageAudit
\$modulePath

If you're on Windows PowerShell 5.1:

\$modulePath = Join-Path \$HOME
"Documents`\WindowsPowerShell`{=tex}`\Modules`{=tex}`\Invoke`{=tex}-ExchangeMessageAudit"

Then:

Import-Module Invoke-ExchangeMessageAudit -Force

Now you can call it from anywhere:

Invoke-ExchangeMessageAudit

------------------------------------------------------------------------

### Install System-Wide (All Users)

Requires admin rights.

\$modulePath =
"C:`\Program `{=tex}Files`\PowerShell`{=tex}`\Modules`{=tex}`\Invoke`{=tex}-ExchangeMessageAudit"\
git clone https://github.com/r0tifer/Invoke-ExchangeMessageAudit
\$modulePath

Restart PowerShell and verify:

Get-Module -ListAvailable Invoke-ExchangeMessageAudit

Then run it normally:

Invoke-ExchangeMessageAudit

------------------------------------------------------------------------

## Updating the Module

cd `<module install path>`{=html}\
git pull

Reload:

Import-Module Invoke-ExchangeMessageAudit -Force

------------------------------------------------------------------------

## Testing

Run Pester tests:

Invoke-Pester -Path .`\tests  `{=tex}

Run approved verb validation:

Import-Module PSScriptAnalyzer -Force\
.`\scripts`{=tex}`\Test`{=tex}-ApprovedVerbs.ps1

GitHub Actions enforces approved PowerShell verbs in CI.

------------------------------------------------------------------------

## Contributing

If you want to improve this:

-   Keep modules single-purpose\
-   Keep logging centralized\
-   Don't mix output formatting into core logic\
-   Add tests

PRs welcome. Just don't turn it into a 4,000 line monolith.

------------------------------------------------------------------------

## Final Notes

This was built for real-world Exchange investigations. It's not flashy.
It's not SaaS. It's just structured automation around things admins
already do --- but in a way thats repeatable and defensible.

If it saves you an hour on a compliance ticket, it did its job.