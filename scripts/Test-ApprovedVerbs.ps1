Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$results = Invoke-ScriptAnalyzer -Path . -Recurse -IncludeRule PSUseApprovedVerbs

if ($results) {
  Write-Host 'Approved verb guard failed. The following command names use non-approved verbs:' -ForegroundColor Red
  $results |
    Select-Object RuleName, Severity, ScriptName, Line, Message |
    Format-Table -AutoSize |
    Out-String |
    Write-Host
  exit 1
}

Write-Host 'Approved verb guard passed: all discovered PowerShell command names use approved verbs.' -ForegroundColor Green
