Set-StrictMode -Version Latest

$script:ImtLogger = [ordered]@{
  Initialized = $false
  OutputLevel = 'INFO'
  DisableTranscriptLog = $false
  StepLogPath = $null
  TranscriptPath = $null
  TranscriptStarted = $false
}

function Get-ImtLogLevelValue {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory=$true)][ValidateSet('DEBUG','INFO','WARN','ERROR','CRITICAL')][string]$Level
  )

  switch ($Level.ToUpperInvariant()) {
    'DEBUG' { return 10 }
    'INFO' { return 20 }
    'WARN' { return 30 }
    'ERROR' { return 40 }
    'CRITICAL' { return 50 }
    default { return 20 }
  }
}

function Test-ImtShouldEmitConsole {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory=$true)][ValidateSet('DEBUG','INFO','WARN','ERROR','CRITICAL')][string]$CurrentOutputLevel,
    [Parameter(Mandatory=$true)][ValidateSet('DEBUG','INFO','WARN','ERROR','CRITICAL')][string]$MessageLevel,
    [Parameter(Mandatory=$true)][ValidateSet('Start','Progress','Result','Summary')][string]$EventType
  )

  $configured = $CurrentOutputLevel.ToUpperInvariant()
  $messagePriority = Get-ImtLogLevelValue -Level $MessageLevel

  if ($EventType -eq 'Summary') {
    return $true
  }

  switch ($configured) {
    'DEBUG' {
      return $true
    }
    'INFO' {
      if ($messagePriority -ge (Get-ImtLogLevelValue -Level 'WARN')) {
        return $true
      }
      return $EventType -in @('Start','Result')
    }
    'WARN' {
      return $messagePriority -ge (Get-ImtLogLevelValue -Level 'WARN')
    }
    'ERROR' {
      return $messagePriority -ge (Get-ImtLogLevelValue -Level 'ERROR')
    }
    'CRITICAL' {
      return $messagePriority -ge (Get-ImtLogLevelValue -Level 'CRITICAL')
    }
    default {
      return $true
    }
  }
}

function Initialize-ImtLogger {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory=$true)][ValidateSet('DEBUG','INFO','WARN','ERROR','CRITICAL')][string]$OutputLevel,
    [switch]$DisableTranscriptLog,
    [Parameter(Mandatory=$true)][string]$StepLogPath,
    [Parameter(Mandatory=$true)][string]$TranscriptPath
  )

  $script:ImtLogger.OutputLevel = $OutputLevel.ToUpperInvariant()
  $script:ImtLogger.DisableTranscriptLog = [bool]$DisableTranscriptLog
  $script:ImtLogger.StepLogPath = $StepLogPath
  $script:ImtLogger.TranscriptPath = $TranscriptPath
  $script:ImtLogger.TranscriptStarted = $false
  $script:ImtLogger.Initialized = $true

  if (-not $script:ImtLogger.DisableTranscriptLog) {
    "Timestamp`tLevel`tStep`tEventType`tMessage" | Out-File -FilePath $script:ImtLogger.StepLogPath -Encoding UTF8
    try {
      Start-Transcript -Path $script:ImtLogger.TranscriptPath -ErrorAction Stop | Out-Null
      $script:ImtLogger.TranscriptStarted = $true
    } catch {
      Write-ImtLog -Level WARN -Step 'Logging' -EventType Result -Message ("Transcript start failed: {0}" -f $_.Exception.Message)
    }
  }

  $detail = if ($script:ImtLogger.DisableTranscriptLog) {
    'Transcript logging disabled.'
  } else {
    "Transcript={0}; Steps={1}" -f $script:ImtLogger.TranscriptPath, $script:ImtLogger.StepLogPath
  }
  Write-ImtLog -Level INFO -Step 'Logging' -EventType Result -Message $detail
}

function Get-ImtLogColor {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory=$true)][ValidateSet('DEBUG','INFO','WARN','ERROR','CRITICAL')][string]$Level,
    [Parameter(Mandatory=$true)][ValidateSet('Start','Progress','Result','Summary')][string]$EventType
  )

  if ($EventType -eq 'Summary') { return 'Cyan' }

  switch ($Level.ToUpperInvariant()) {
    'DEBUG' { return 'DarkGray' }
    'INFO' { return 'Gray' }
    'WARN' { return 'Yellow' }
    'ERROR' { return 'Red' }
    'CRITICAL' { return 'Magenta' }
    default { return 'White' }
  }
}

function Write-ImtLog {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory=$true)][ValidateSet('DEBUG','INFO','WARN','ERROR','CRITICAL')][string]$Level,
    [Parameter(Mandatory=$true)][string]$Step,
    [Parameter(Mandatory=$true)][string]$Message,
    [Parameter(Mandatory=$true)][ValidateSet('Start','Progress','Result','Summary')][string]$EventType
  )

  $normalizedLevel = $Level.ToUpperInvariant()
  $timestamp = (Get-Date).ToString('o')
  $cleanMessage = ($Message -replace "(\r|\n)", ' ')

  if ($script:ImtLogger.Initialized -and -not $script:ImtLogger.DisableTranscriptLog -and $script:ImtLogger.StepLogPath) {
    $line = "{0}`t{1}`t{2}`t{3}`t{4}" -f $timestamp, $normalizedLevel, $Step, $EventType, $cleanMessage
    Add-Content -Path $script:ImtLogger.StepLogPath -Value $line -Encoding UTF8
  }

  $outputLevel = if ($script:ImtLogger.Initialized) { $script:ImtLogger.OutputLevel } else { 'INFO' }
  if (Test-ImtShouldEmitConsole -CurrentOutputLevel $outputLevel -MessageLevel $normalizedLevel -EventType $EventType) {
    $color = Get-ImtLogColor -Level $normalizedLevel -EventType $EventType
    $prefix = "[{0}] [{1}] [{2}]" -f $normalizedLevel, $Step, $EventType
    Write-Host ("{0} {1}" -f $prefix, $cleanMessage) -ForegroundColor $color
  }
}

function Complete-ImtLogger {
  [CmdletBinding()]
  param()

  if ($script:ImtLogger.Initialized -and $script:ImtLogger.TranscriptStarted) {
    try {
      Stop-Transcript | Out-Null
    } catch {
      # no-op
    }
    $script:ImtLogger.TranscriptStarted = $false
  }
}

function Get-ImtLoggerState {
  [CmdletBinding()]
  param()

  [pscustomobject]@{
    Initialized = $script:ImtLogger.Initialized
    OutputLevel = $script:ImtLogger.OutputLevel
    DisableTranscriptLog = $script:ImtLogger.DisableTranscriptLog
    StepLogPath = $script:ImtLogger.StepLogPath
    TranscriptPath = $script:ImtLogger.TranscriptPath
    TranscriptStarted = $script:ImtLogger.TranscriptStarted
  }
}
