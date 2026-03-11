param(
  [string]$ConfigPath,
  [string]$DefaultDbPath,
  [string]$DefaultHost = "127.0.0.1",
  [int]$DefaultPort = 8787
)

$ErrorActionPreference = "Stop"

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$script:InstallDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$script:BackendScriptPath = Join-Path $script:InstallDir "tds-street-backend.py"
$script:IndexerLauncherPath = Join-Path $script:InstallDir "tds-street-indexer-launcher.cmd"
$script:AutoSetupLauncherPath = Join-Path $script:InstallDir "tds-local-streets-auto-setup-launcher.cmd"
$script:ManagerIconPath = Join-Path $script:InstallDir "tds-street-backend-manager.ico"
$script:ManagerDataDir = Join-Path ([Environment]::GetFolderPath("LocalApplicationData")) "TDS-Pak\StreetBackendManager"
$script:RuntimeLogPath = Join-Path $script:ManagerDataDir "backend-runtime.log"
$script:RuntimeErrorLogPath = Join-Path $script:ManagerDataDir "backend-runtime-error.log"
$script:ManagedBackendProcess = $null
$script:RuntimeLogOffset = 0
$script:RuntimeErrorLogOffset = 0
$script:LastHealthSummary = ""
$script:DefaultTdsPakUrl = "https://vdiaztds.github.io/"
$script:Theme = @{
  PageBg = [System.Drawing.Color]::FromArgb(238, 245, 252)
  HeaderBg = [System.Drawing.Color]::FromArgb(252, 254, 255)
  CardBorder = [System.Drawing.Color]::FromArgb(204, 218, 236)
  Accent = [System.Drawing.Color]::FromArgb(23, 92, 196)
  AccentSoft = [System.Drawing.Color]::FromArgb(232, 241, 255)
  TextStrong = [System.Drawing.Color]::FromArgb(25, 40, 63)
  TextMuted = [System.Drawing.Color]::FromArgb(83, 102, 127)
  InputBorder = [System.Drawing.Color]::FromArgb(193, 207, 225)
  LogBg = [System.Drawing.Color]::FromArgb(17, 24, 39)
  LogFg = [System.Drawing.Color]::FromArgb(229, 241, 255)
  Success = [System.Drawing.Color]::FromArgb(26, 135, 84)
  Danger = [System.Drawing.Color]::FromArgb(181, 56, 56)
}

if (-not (Test-Path -LiteralPath $script:ManagerDataDir)) {
  New-Item -Path $script:ManagerDataDir -ItemType Directory -Force | Out-Null
}

if (-not $ConfigPath) {
  $ConfigPath = Join-Path $script:ManagerDataDir "manager-config.json"
}
$script:ConfigPath = $ConfigPath

function Get-PreferredDbPath {
  if ($DefaultDbPath -and (Test-Path -LiteralPath $DefaultDbPath)) {
    return $DefaultDbPath
  }
  $docs = [Environment]::GetFolderPath("MyDocuments")
  $candidate = if ($docs) { Join-Path $docs "TDS-Pak-Street-Data\tds-streets.sqlite" } else { "" }
  if ($candidate -and (Test-Path -LiteralPath $candidate)) {
    return $candidate
  }
  $fallback = Join-Path $script:InstallDir "tds-streets.sqlite"
  return $fallback
}

function Find-PythonRuntime {
  $candidates = @(
    @{ Command = "py"; PrefixArgs = @("-3") },
    @{ Command = "python"; PrefixArgs = @() },
    @{ Command = "python3"; PrefixArgs = @() }
  )

  foreach ($candidate in $candidates) {
    $cmd = Get-Command -Name $candidate.Command -ErrorAction SilentlyContinue | Select-Object -First 1
    if (-not $cmd) { continue }
    $exe = if ($cmd.Path) { $cmd.Path } else { $candidate.Command }
    try {
      & $exe @($candidate.PrefixArgs + @("-c", "import sys; print(sys.version)")) *> $null
      if ($LASTEXITCODE -eq 0) {
        return @{
          Command = $exe
          PrefixArgs = @($candidate.PrefixArgs)
        }
      }
    } catch {}
  }
  return $null
}

function Load-ManagerConfig {
  $default = @{
    db_path = (Get-PreferredDbPath)
    host = $DefaultHost
    port = [Math]::Max(1, [Math]::Min(65535, [int]$DefaultPort))
    auto_start = $false
    tds_pak_url = $script:DefaultTdsPakUrl
  }

  if (-not (Test-Path -LiteralPath $script:ConfigPath)) {
    return $default
  }

  try {
    $raw = Get-Content -LiteralPath $script:ConfigPath -Raw
    $parsed = $raw | ConvertFrom-Json
    if (-not $parsed) { return $default }
    return @{
      db_path = if ($parsed.db_path) { [string]$parsed.db_path } else { $default.db_path }
      host = if ($parsed.host) { [string]$parsed.host } else { $default.host }
      port = if ($parsed.port) { [int]$parsed.port } else { $default.port }
      auto_start = [bool]$parsed.auto_start
      tds_pak_url = (Resolve-TdsPakUrl (if ($parsed.tds_pak_url) { [string]$parsed.tds_pak_url } else { $default.tds_pak_url }))
    }
  } catch {
    return $default
  }
}

function Save-ManagerConfig([hashtable]$config) {
  $payload = @{
    db_path = [string]$config.db_path
    host = [string]$config.host
    port = [int]$config.port
    auto_start = [bool]$config.auto_start
    tds_pak_url = [string]$config.tds_pak_url
    updated_at = (Get-Date).ToString("s")
  }
  $json = $payload | ConvertTo-Json -Depth 4
  Set-Content -LiteralPath $script:ConfigPath -Value $json -Encoding UTF8
}

function Resolve-TdsPakUrl([string]$inputUrl) {
  $candidate = [string]$inputUrl
  if (-not $candidate) {
    return $script:DefaultTdsPakUrl
  }

  try {
    $uri = [Uri]$candidate
    if ($uri.IsLoopback) {
      return $script:DefaultTdsPakUrl
    }
  } catch {
    if ($candidate -match "(?i)^(https?:\/\/)?(localhost|127\.0\.0\.1)(:\d+)?(\/|$)") {
      return $script:DefaultTdsPakUrl
    }
  }

  return $candidate
}

function Get-BackendHealth([string]$backendHost, [int]$port) {
  $uri = "http://$backendHost`:$port/api/health"
  try {
    $response = Invoke-RestMethod -Method Get -Uri $uri -TimeoutSec 2
    $sourceName = if ($null -ne $response.source_name) { [string]$response.source_name } else { "" }
    $rowCount = 0
    if ($null -ne $response.row_count) {
      try { $rowCount = [int]$response.row_count } catch { $rowCount = 0 }
    }
    $updatedAt = if ($null -ne $response.updated_at) { [string]$response.updated_at } else { "" }
    return @{
      ok = $true
      uri = $uri
      has_index = [bool]$response.has_index
      source_name = $sourceName
      row_count = $rowCount
      updated_at = $updatedAt
      error = ""
    }
  } catch {
    return @{
      ok = $false
      uri = $uri
      has_index = $false
      source_name = ""
      row_count = 0
      updated_at = ""
      error = [string]$_.Exception.Message
    }
  }
}

function Is-ManagedBackendRunning {
  if (-not $script:ManagedBackendProcess) { return $false }
  try {
    return -not $script:ManagedBackendProcess.HasExited
  } catch {
    return $false
  }
}

function Get-BackendPortFromUi {
  $port = 0
  if ($script:PortTextBox) {
    [void][int]::TryParse([string]$script:PortTextBox.Text, [ref]$port)
  }
  if ($port -lt 1 -or $port -gt 65535) {
    $port = 8787
  }
  return $port
}

function Get-ProcessCommandLineById([int]$processId) {
  if ($processId -le 0) { return "" }
  try {
    $proc = Get-CimInstance -ClassName Win32_Process -Filter "ProcessId=$processId" -ErrorAction Stop
    return [string]$proc.CommandLine
  } catch {
    return ""
  }
}

function Test-BackendProcessId([int]$processId) {
  if ($processId -le 0) { return $false }
  $cmd = Get-ProcessCommandLineById -processId $processId
  if (-not $cmd) { return $false }
  return [regex]::IsMatch($cmd, "(?i)(^|[\\/""'\s])tds-street-backend\.py([""'\s]|$)")
}

function Find-ListeningProcessByPort([int]$port) {
  if ($port -le 0) { return $null }

  try {
    $conn = Get-NetTCPConnection -LocalPort $port -State Listen -ErrorAction Stop | Select-Object -First 1
    if ($conn -and $conn.OwningProcess -gt 0) {
      return [int]$conn.OwningProcess
    }
  } catch {}

  try {
    $regex = "^\s*TCP\s+\S+:$port\s+\S+\s+LISTENING\s+(\d+)\s*$"
    $line = netstat -ano -p tcp | Select-String -Pattern $regex | Select-Object -First 1
    if ($line -and $line.Matches.Count -gt 0) {
      return [int]$line.Matches[0].Groups[1].Value
    }
  } catch {}

  return $null
}

function Find-BackendProcessIds([int]$port) {
  $set = @{}

  if (Is-ManagedBackendRunning) {
    try {
      $set[[int]$script:ManagedBackendProcess.Id] = $true
    } catch {}
  }

  $pidByPort = Find-ListeningProcessByPort -port $port
  if ($pidByPort) {
    $includePidByPort = $false
    if (Test-BackendProcessId -processId $pidByPort) {
      $includePidByPort = $true
    } else {
      try {
        $procByPort = Get-Process -Id ([int]$pidByPort) -ErrorAction Stop
        $procNameByPort = [string]$procByPort.ProcessName
        if ($procNameByPort -match "(?i)^pythonw?$|^py$") {
          # Fallback: command line can be unavailable in some environments.
          $includePidByPort = $true
        }
      } catch {}
    }
    if ($includePidByPort) {
      $set[[int]$pidByPort] = $true
    }
  }

  try {
    $pythonProcs = Get-CimInstance -ClassName Win32_Process -Filter "Name='python.exe' OR Name='pythonw.exe' OR Name='py.exe'" -ErrorAction Stop
    foreach ($proc in $pythonProcs) {
      $procId = [int]$proc.ProcessId
      $cmd = [string]$proc.CommandLine
      if ($procId -gt 0 -and $cmd -and ([regex]::IsMatch($cmd, "(?i)(^|[\\/""'\s])tds-street-backend\.py([""'\s]|$)"))) {
        $set[$procId] = $true
      }
    }
  } catch {}

  return @($set.Keys | ForEach-Object { [int]$_ } | Sort-Object -Unique)
}

function Find-AllBackendProcessIds {
  $set = @{}

  $portsToCheck = New-Object System.Collections.Generic.HashSet[int]
  [void]$portsToCheck.Add(8787)
  [void]$portsToCheck.Add((Get-BackendPortFromUi))
  try {
    $cfg = Get-CurrentConfigFromUi
    if ($cfg -and $cfg.port) {
      [void]$portsToCheck.Add([int]$cfg.port)
    }
  } catch {}

  foreach ($port in $portsToCheck) {
    foreach ($procId in (Find-BackendProcessIds -port $port)) {
      $set[[int]$procId] = $true
    }
  }

  try {
    $allProcs = Get-CimInstance -ClassName Win32_Process -ErrorAction Stop
    foreach ($proc in $allProcs) {
      $procId = [int]$proc.ProcessId
      if ($procId -le 0) { continue }
      $cmd = [string]$proc.CommandLine
      if (-not $cmd) { continue }
      if (
        [regex]::IsMatch($cmd, "(?i)(^|[\\/""'\s])tds-street-backend\.py([""'\s]|$)") -or
        [regex]::IsMatch($cmd, "(?i)(^|[\\/""'\s])tds-street-backend-launcher\.cmd([""'\s]|$)")
      ) {
        $set[$procId] = $true
      }
    }
  } catch {}

  return @($set.Keys | ForEach-Object { [int]$_ } | Sort-Object -Unique)
}

$script:LogTextBox = $null
$script:StatusLabel = $null
$script:HealthLabel = $null
$script:HostTextBox = $null
$script:PortTextBox = $null
$script:DbPathTextBox = $null
$script:TdsPakUrlTextBox = $null
$script:AutoStartCheckBox = $null

function Set-ActionButtonStyle(
  [System.Windows.Forms.Button]$Button,
  [System.Drawing.Color]$Background,
  [System.Drawing.Color]$Foreground
) {
  if (-not $Button) { return }
  $Button.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
  $Button.FlatAppearance.BorderSize = 0
  $Button.FlatAppearance.MouseOverBackColor = [System.Drawing.Color]::FromArgb(
    [Math]::Max(0, $Background.R - 12),
    [Math]::Max(0, $Background.G - 12),
    [Math]::Max(0, $Background.B - 12)
  )
  $Button.FlatAppearance.MouseDownBackColor = [System.Drawing.Color]::FromArgb(
    [Math]::Max(0, $Background.R - 24),
    [Math]::Max(0, $Background.G - 24),
    [Math]::Max(0, $Background.B - 24)
  )
  $Button.BackColor = $Background
  $Button.ForeColor = $Foreground
  $Button.Font = New-Object System.Drawing.Font("Segoe UI Semibold", 9.5, [System.Drawing.FontStyle]::Regular)
  $Button.Cursor = [System.Windows.Forms.Cursors]::Hand
}

function Set-SecondaryButtonStyle([System.Windows.Forms.Button]$Button) {
  if (-not $Button) { return }
  $Button.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
  $Button.FlatAppearance.BorderSize = 1
  $Button.FlatAppearance.BorderColor = $script:Theme.CardBorder
  $Button.FlatAppearance.MouseOverBackColor = [System.Drawing.Color]::FromArgb(242, 247, 253)
  $Button.FlatAppearance.MouseDownBackColor = [System.Drawing.Color]::FromArgb(231, 239, 249)
  $Button.BackColor = [System.Drawing.Color]::FromArgb(255, 255, 255)
  $Button.ForeColor = $script:Theme.TextStrong
  $Button.Font = New-Object System.Drawing.Font("Segoe UI", 9.0, [System.Drawing.FontStyle]::Regular)
  $Button.Cursor = [System.Windows.Forms.Cursors]::Hand
}

function Set-InputControlStyle([System.Windows.Forms.Control]$Control) {
  if (-not $Control) { return }
  if ($Control -is [System.Windows.Forms.TextBox]) {
    $Control.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
    $Control.BackColor = [System.Drawing.Color]::FromArgb(255, 255, 255)
    $Control.ForeColor = $script:Theme.TextStrong
    $Control.Font = New-Object System.Drawing.Font("Segoe UI", 9.25, [System.Drawing.FontStyle]::Regular)
  } elseif ($Control -is [System.Windows.Forms.CheckBox]) {
    $Control.ForeColor = $script:Theme.TextStrong
    $Control.Font = New-Object System.Drawing.Font("Segoe UI Semibold", 9.0, [System.Drawing.FontStyle]::Regular)
  }
}

function Write-ManagerLog([string]$message) {
  if (-not $script:LogTextBox) { return }
  $stamp = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
  $script:LogTextBox.AppendText("[$stamp] $message`r`n")
  $script:LogTextBox.SelectionStart = $script:LogTextBox.Text.Length
  $script:LogTextBox.ScrollToCaret()
}

function Read-BackendRuntimeLogTail {
  if (-not $script:LogTextBox) { return }
  if (Test-Path -LiteralPath $script:RuntimeLogPath) {
    try {
      $info = Get-Item -LiteralPath $script:RuntimeLogPath -ErrorAction Stop
      if ($script:RuntimeLogOffset -gt $info.Length) {
        $script:RuntimeLogOffset = 0
      }

      $stream = [System.IO.File]::Open($script:RuntimeLogPath, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::ReadWrite)
      try {
        $null = $stream.Seek($script:RuntimeLogOffset, [System.IO.SeekOrigin]::Begin)
        $reader = New-Object System.IO.StreamReader($stream)
        $delta = $reader.ReadToEnd()
        $script:RuntimeLogOffset = $stream.Position
        if ($delta) {
          $normalized = $delta -replace "`r?`n", "`r`n"
          $script:LogTextBox.AppendText($normalized)
          $script:LogTextBox.SelectionStart = $script:LogTextBox.Text.Length
          $script:LogTextBox.ScrollToCaret()
        }
      } finally {
        $stream.Dispose()
      }
    } catch {}
  }

  if (Test-Path -LiteralPath $script:RuntimeErrorLogPath) {
    try {
      $errInfo = Get-Item -LiteralPath $script:RuntimeErrorLogPath -ErrorAction Stop
      if ($script:RuntimeErrorLogOffset -gt $errInfo.Length) {
        $script:RuntimeErrorLogOffset = 0
      }

      $errStream = [System.IO.File]::Open($script:RuntimeErrorLogPath, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::ReadWrite)
      try {
        $null = $errStream.Seek($script:RuntimeErrorLogOffset, [System.IO.SeekOrigin]::Begin)
        $errReader = New-Object System.IO.StreamReader($errStream)
        $errDelta = $errReader.ReadToEnd()
        $script:RuntimeErrorLogOffset = $errStream.Position
        if ($errDelta) {
          $normalizedErr = $errDelta -replace "`r?`n", "`r`n"
          $prefixedErr = ($normalizedErr -split "`r`n" | Where-Object { $_ -ne "" } | ForEach-Object { "[stderr] $_" }) -join "`r`n"
          if ($prefixedErr) {
            $script:LogTextBox.AppendText($prefixedErr + "`r`n")
            $script:LogTextBox.SelectionStart = $script:LogTextBox.Text.Length
            $script:LogTextBox.ScrollToCaret()
          }
        }
      } finally {
        $errStream.Dispose()
      }
    } catch {}
  }
}

function Update-BackendStatusUi([bool]$forceLog = $false) {
  if (-not $script:StatusLabel -or -not $script:HealthLabel -or -not $script:HostTextBox -or -not $script:PortTextBox) {
    return
  }

  $backendHost = [string]$script:HostTextBox.Text
  $port = 0
  [void][int]::TryParse([string]$script:PortTextBox.Text, [ref]$port)
  if ($port -le 0) { $port = 8787 }
  $health = Get-BackendHealth -backendHost $backendHost -port $port

  if ($health.ok) {
    $indexState = if ($health.has_index) { "Index Loaded" } else { "No Index" }
    $summary = "Online | $indexState | Rows: $($health.row_count.ToString('N0'))"
    $script:StatusLabel.Text = "Backend Status: $summary"
    $script:StatusLabel.ForeColor = $script:Theme.Success
    $sourceText = if ($health.source_name) { $health.source_name } else { "Unknown source" }
    $script:HealthLabel.Text = "Health: $($health.uri) | Source: $sourceText"
    if ($forceLog -or $script:LastHealthSummary -ne $summary) {
      Write-ManagerLog "Health check OK: $summary"
      $script:LastHealthSummary = $summary
    }
    return
  }

  $summary = "Offline"
  $script:StatusLabel.Text = "Backend Status: $summary"
  $script:StatusLabel.ForeColor = $script:Theme.Danger
  $script:HealthLabel.Text = "Health: $($health.uri) | Error: $($health.error)"
  if ($forceLog -or $script:LastHealthSummary -ne $summary) {
    Write-ManagerLog "Health check failed: $($health.error)"
    $script:LastHealthSummary = $summary
  }
}

function Get-CurrentConfigFromUi {
  $backendHost = [string]$script:HostTextBox.Text
  $port = 0
  [void][int]::TryParse([string]$script:PortTextBox.Text, [ref]$port)
  if ($port -lt 1 -or $port -gt 65535) { $port = 8787 }
  return @{
    db_path = [string]$script:DbPathTextBox.Text
    host = if ($backendHost) { $backendHost } else { "127.0.0.1" }
    port = $port
    auto_start = [bool]$script:AutoStartCheckBox.Checked
    tds_pak_url = (Resolve-TdsPakUrl -inputUrl ([string]$script:TdsPakUrlTextBox.Text))
  }
}

function Start-ManagedBackend {
  if (Is-ManagedBackendRunning) {
    Write-ManagerLog "Backend already started by this app."
    Update-BackendStatusUi -forceLog:$false
    return
  }

  if (-not (Test-Path -LiteralPath $script:BackendScriptPath)) {
    [System.Windows.Forms.MessageBox]::Show(
      "Could not find backend script: $script:BackendScriptPath",
      "TDS PAK Backend Manager",
      [System.Windows.Forms.MessageBoxButtons]::OK,
      [System.Windows.Forms.MessageBoxIcon]::Error
    ) | Out-Null
    return
  }

  $cfg = Get-CurrentConfigFromUi
  Save-ManagerConfig -config $cfg

  $existingBackendPids = Find-BackendProcessIds -port ([int]$cfg.port)
  if ($existingBackendPids.Count -gt 0) {
    $pidText = ($existingBackendPids | ForEach-Object { $_.ToString() }) -join ", "
    $primaryPid = [int]$existingBackendPids[0]
    try {
      $script:ManagedBackendProcess = Get-Process -Id $primaryPid -ErrorAction Stop
    } catch {
      $script:ManagedBackendProcess = $null
    }
    Write-ManagerLog "Backend already running (PID: $pidText) on port $($cfg.port)."
    Update-BackendStatusUi -forceLog:$true
    return
  }

  $dbPath = [string]$cfg.db_path
  if (-not $dbPath) {
    [System.Windows.Forms.MessageBox]::Show(
      "Select a SQLite index path before starting backend.",
      "TDS PAK Backend Manager",
      [System.Windows.Forms.MessageBoxButtons]::OK,
      [System.Windows.Forms.MessageBoxIcon]::Warning
    ) | Out-Null
    return
  }

  if (-not (Test-Path -LiteralPath $dbPath)) {
    $answer = [System.Windows.Forms.MessageBox]::Show(
      "SQLite index was not found at:`r`n$dbPath`r`n`r`nStart backend anyway?",
      "TDS PAK Backend Manager",
      [System.Windows.Forms.MessageBoxButtons]::YesNo,
      [System.Windows.Forms.MessageBoxIcon]::Question
    )
    if ($answer -ne [System.Windows.Forms.DialogResult]::Yes) {
      return
    }
  }

  $python = Find-PythonRuntime
  if (-not $python) {
    [System.Windows.Forms.MessageBox]::Show(
      "Python was not found. Install Python from python.org and retry.",
      "TDS PAK Backend Manager",
      [System.Windows.Forms.MessageBoxButtons]::OK,
      [System.Windows.Forms.MessageBoxIcon]::Error
    ) | Out-Null
    return
  }

  try {
    Set-Content -LiteralPath $script:RuntimeLogPath -Value "" -Encoding UTF8
    Set-Content -LiteralPath $script:RuntimeErrorLogPath -Value "" -Encoding UTF8
    $script:RuntimeLogOffset = 0
    $script:RuntimeErrorLogOffset = 0
  } catch {}

  $argList = @()
  $argList += @($python.PrefixArgs)
  $argList += @(
    $script:BackendScriptPath,
    "--db", $dbPath,
    "--host", [string]$cfg.host,
    "--port", [string]$cfg.port
  )

  try {
    $proc = Start-Process -FilePath $python.Command `
      -ArgumentList $argList `
      -PassThru `
      -WindowStyle Hidden `
      -RedirectStandardOutput $script:RuntimeLogPath `
      -RedirectStandardError $script:RuntimeErrorLogPath
    $script:ManagedBackendProcess = $proc
    Write-ManagerLog "Started backend process (PID $($proc.Id)) on http://$($cfg.host):$($cfg.port)"
  } catch {
    [System.Windows.Forms.MessageBox]::Show(
      "Failed to start backend:`r`n$($_.Exception.Message)",
      "TDS PAK Backend Manager",
      [System.Windows.Forms.MessageBoxButtons]::OK,
      [System.Windows.Forms.MessageBoxIcon]::Error
    ) | Out-Null
    Write-ManagerLog "Failed to start backend: $($_.Exception.Message)"
    return
  }

  Start-Sleep -Milliseconds 450
  Read-BackendRuntimeLogTail
  Update-BackendStatusUi -forceLog:$true
}

function Stop-ManagedBackend {
  $port = Get-BackendPortFromUi
  $backendPids = Find-BackendProcessIds -port $port
  if (-not $backendPids -or $backendPids.Count -eq 0) {
    Write-ManagerLog "No running backend process was found to stop."
    $script:ManagedBackendProcess = $null
    Update-BackendStatusUi -forceLog:$false
    return
  }

  $result = Stop-BackendProcessIds -processIds $backendPids -actionLabel "Stop"
  if ($result.StoppedCount -gt 0) {
    Write-ManagerLog "Stop complete. Terminated PID(s): $($result.StoppedText)"
  }
  if ($result.AlreadyStoppedCount -gt 0) {
    Write-ManagerLog "Stop note: PID(s) already stopped: $($result.AlreadyStoppedText)"
  }
  if ($result.FailedCount -gt 0) {
    Write-ManagerLog "Stop attempted with failures: $($result.FailedText)"
  }

  Start-Sleep -Milliseconds 300
  Read-BackendRuntimeLogTail
  Update-BackendStatusUi -forceLog:$true
}

function Stop-BackendProcessIds(
  [int[]]$processIds,
  [string]$actionLabel = "Stop"
) {
  $stopped = New-Object System.Collections.Generic.List[int]
  $alreadyStopped = New-Object System.Collections.Generic.List[int]
  $failed = New-Object System.Collections.Generic.List[string]

  foreach ($backendProcId in ($processIds | Select-Object -Unique)) {
    $targetPid = 0
    if (-not [int]::TryParse([string]$backendProcId, [ref]$targetPid)) { continue }
    if ($targetPid -le 0) { continue }
    $existingProc = $null
    try {
      $existingProc = Get-Process -Id $targetPid -ErrorAction SilentlyContinue
    } catch {}
    if (-not $existingProc) {
      [void]$alreadyStopped.Add($targetPid)
      Write-ManagerLog "${actionLabel}: backend process (PID $targetPid) was already stopped."
      continue
    }
    try {
      Stop-Process -Id $targetPid -Force -ErrorAction Stop
      [void]$stopped.Add($targetPid)
      Write-ManagerLog "${actionLabel}: stopped backend process (PID $targetPid)."
    } catch {
      $message = [string]$_.Exception.Message
      if ($message -match "(?i)cannot find a process with the process identifier|no process.+found") {
        [void]$alreadyStopped.Add($targetPid)
        Write-ManagerLog "${actionLabel}: backend process (PID $targetPid) exited before it could be terminated."
        continue
      }
      [void]$failed.Add("$targetPid ($message)")
      Write-ManagerLog "${actionLabel}: failed to stop backend process (PID $targetPid): $message"
    }
  }

  $script:ManagedBackendProcess = $null
  $stoppedText = ($stopped | ForEach-Object { $_.ToString() } | Sort-Object -Unique) -join ", "
  $alreadyStoppedText = ($alreadyStopped | ForEach-Object { $_.ToString() } | Sort-Object -Unique) -join ", "
  $failedText = ($failed | ForEach-Object { $_.ToString() }) -join " | "
  return @{
    StoppedCount = $stopped.Count
    AlreadyStoppedCount = $alreadyStopped.Count
    FailedCount = $failed.Count
    StoppedText = $stoppedText
    AlreadyStoppedText = $alreadyStoppedText
    FailedText = $failedText
  }
}

function Force-StopAllBackendInstances {
  $answer = [System.Windows.Forms.MessageBox]::Show(
    "Force Stop All will terminate every detected TDS street backend instance, including orphaned processes from old managers/installations.`r`n`r`nContinue?",
    "TDS PAK Backend Manager",
    [System.Windows.Forms.MessageBoxButtons]::YesNo,
    [System.Windows.Forms.MessageBoxIcon]::Warning
  )
  if ($answer -ne [System.Windows.Forms.DialogResult]::Yes) {
    Write-ManagerLog "Force Stop All canceled by user."
    return
  }

  $allPids = Find-AllBackendProcessIds
  if (-not $allPids -or $allPids.Count -eq 0) {
    Write-ManagerLog "Force Stop All: no TDS backend-related processes were detected."
    $script:ManagedBackendProcess = $null
    Update-BackendStatusUi -forceLog:$true
    return
  }

  $pidText = ($allPids | ForEach-Object { $_.ToString() }) -join ", "
  Write-ManagerLog "Force Stop All: detected PID(s): $pidText"

  $result = Stop-BackendProcessIds -processIds $allPids -actionLabel "Force Stop All"
  if ($result.StoppedCount -gt 0) {
    Write-ManagerLog "Force Stop All complete. Terminated PID(s): $($result.StoppedText)"
  }
  if ($result.AlreadyStoppedCount -gt 0) {
    Write-ManagerLog "Force Stop All note: PID(s) already stopped: $($result.AlreadyStoppedText)"
  }
  if ($result.FailedCount -gt 0) {
    Write-ManagerLog "Force Stop All had failures: $($result.FailedText)"
    [System.Windows.Forms.MessageBox]::Show(
      "Some processes could not be terminated.`r`n`r`n$($result.FailedText)",
      "TDS PAK Backend Manager",
      [System.Windows.Forms.MessageBoxButtons]::OK,
      [System.Windows.Forms.MessageBoxIcon]::Warning
    ) | Out-Null
  } elseif ($result.StoppedCount -gt 0 -or $result.AlreadyStoppedCount -gt 0) {
    $summaryLines = New-Object System.Collections.Generic.List[string]
    [void]$summaryLines.Add("Force Stop All completed successfully.")
    if ($result.StoppedCount -gt 0) {
      [void]$summaryLines.Add("")
      [void]$summaryLines.Add("Terminated PID(s): $($result.StoppedText)")
    }
    if ($result.AlreadyStoppedCount -gt 0) {
      [void]$summaryLines.Add("")
      [void]$summaryLines.Add("Already stopped PID(s): $($result.AlreadyStoppedText)")
    }
    [System.Windows.Forms.MessageBox]::Show(
      ($summaryLines -join "`r`n"),
      "TDS PAK Backend Manager",
      [System.Windows.Forms.MessageBoxButtons]::OK,
      [System.Windows.Forms.MessageBoxIcon]::Information
    ) | Out-Null
  }

  Start-Sleep -Milliseconds 300
  Read-BackendRuntimeLogTail
  Update-BackendStatusUi -forceLog:$true
}

$config = Load-ManagerConfig

$form = New-Object System.Windows.Forms.Form
$form.Text = "TDS PAK Street Backend Manager"
$form.StartPosition = "CenterScreen"
$form.Size = New-Object System.Drawing.Size(1100, 760)
$form.MinimumSize = New-Object System.Drawing.Size(980, 680)
$form.BackColor = $script:Theme.PageBg
$form.Font = New-Object System.Drawing.Font("Segoe UI", 9.25, [System.Drawing.FontStyle]::Regular)

if (Test-Path -LiteralPath $script:ManagerIconPath) {
  try {
    $form.Icon = New-Object System.Drawing.Icon($script:ManagerIconPath)
  } catch {}
}

$topPanel = New-Object System.Windows.Forms.Panel
$topPanel.Dock = "Top"
$topPanel.Height = 222
$topPanel.Padding = New-Object System.Windows.Forms.Padding(14, 12, 14, 12)
$topPanel.BackColor = $script:Theme.HeaderBg
$topPanel.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
$form.Controls.Add($topPanel)

$titleLabel = New-Object System.Windows.Forms.Label
$titleLabel.Text = "Street Backend Control Center"
$titleLabel.Font = New-Object System.Drawing.Font("Segoe UI Semibold", 15, [System.Drawing.FontStyle]::Bold)
$titleLabel.AutoSize = $true
$titleLabel.Location = New-Object System.Drawing.Point(12, 8)
$titleLabel.ForeColor = $script:Theme.TextStrong
$topPanel.Controls.Add($titleLabel)

$subtitleLabel = New-Object System.Windows.Forms.Label
$subtitleLabel.Text = "Start and monitor your local TDS street backend. This app is designed to support future database add-ons."
$subtitleLabel.AutoSize = $true
$subtitleLabel.Location = New-Object System.Drawing.Point(14, 38)
$subtitleLabel.ForeColor = $script:Theme.TextMuted
$topPanel.Controls.Add($subtitleLabel)

$dbLabel = New-Object System.Windows.Forms.Label
$dbLabel.Text = "SQLite Index:"
$dbLabel.AutoSize = $true
$dbLabel.Location = New-Object System.Drawing.Point(14, 72)
$dbLabel.ForeColor = $script:Theme.TextStrong
$topPanel.Controls.Add($dbLabel)

$script:DbPathTextBox = New-Object System.Windows.Forms.TextBox
$script:DbPathTextBox.Location = New-Object System.Drawing.Point(116, 68)
$script:DbPathTextBox.Size = New-Object System.Drawing.Size(732, 26)
$script:DbPathTextBox.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
$script:DbPathTextBox.Text = [string]$config.db_path
$topPanel.Controls.Add($script:DbPathTextBox)

$browseDbBtn = New-Object System.Windows.Forms.Button
$browseDbBtn.Text = "Browse..."
$browseDbBtn.Size = New-Object System.Drawing.Size(95, 28)
$browseDbBtn.Location = New-Object System.Drawing.Point(856, 66)
$browseDbBtn.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Right
$topPanel.Controls.Add($browseDbBtn)

$hostLabel = New-Object System.Windows.Forms.Label
$hostLabel.Text = "Host:"
$hostLabel.AutoSize = $true
$hostLabel.Location = New-Object System.Drawing.Point(14, 108)
$hostLabel.ForeColor = $script:Theme.TextStrong
$topPanel.Controls.Add($hostLabel)

$script:HostTextBox = New-Object System.Windows.Forms.TextBox
$script:HostTextBox.Location = New-Object System.Drawing.Point(116, 104)
$script:HostTextBox.Size = New-Object System.Drawing.Size(178, 26)
$script:HostTextBox.Text = [string]$config.host
$topPanel.Controls.Add($script:HostTextBox)

$portLabel = New-Object System.Windows.Forms.Label
$portLabel.Text = "Port:"
$portLabel.AutoSize = $true
$portLabel.Location = New-Object System.Drawing.Point(306, 108)
$portLabel.ForeColor = $script:Theme.TextStrong
$topPanel.Controls.Add($portLabel)

$script:PortTextBox = New-Object System.Windows.Forms.TextBox
$script:PortTextBox.Location = New-Object System.Drawing.Point(348, 104)
$script:PortTextBox.Size = New-Object System.Drawing.Size(96, 26)
$script:PortTextBox.Text = [string]$config.port
$topPanel.Controls.Add($script:PortTextBox)

$script:AutoStartCheckBox = New-Object System.Windows.Forms.CheckBox
$script:AutoStartCheckBox.Text = "Auto-start backend when this app opens"
$script:AutoStartCheckBox.AutoSize = $true
$script:AutoStartCheckBox.Location = New-Object System.Drawing.Point(466, 107)
$script:AutoStartCheckBox.Checked = [bool]$config.auto_start
$topPanel.Controls.Add($script:AutoStartCheckBox)

$urlLabel = New-Object System.Windows.Forms.Label
$urlLabel.Text = "TDS PAK URL:"
$urlLabel.AutoSize = $true
$urlLabel.Location = New-Object System.Drawing.Point(14, 144)
$urlLabel.ForeColor = $script:Theme.TextStrong
$topPanel.Controls.Add($urlLabel)

$script:TdsPakUrlTextBox = New-Object System.Windows.Forms.TextBox
$script:TdsPakUrlTextBox.Location = New-Object System.Drawing.Point(116, 140)
$script:TdsPakUrlTextBox.Size = New-Object System.Drawing.Size(470, 26)
$script:TdsPakUrlTextBox.Text = (Resolve-TdsPakUrl -inputUrl ([string]$config.tds_pak_url))
$topPanel.Controls.Add($script:TdsPakUrlTextBox)

$startBtn = New-Object System.Windows.Forms.Button
$startBtn.Text = "Start Backend"
$startBtn.Size = New-Object System.Drawing.Size(126, 32)
$startBtn.Location = New-Object System.Drawing.Point(14, 178)
$topPanel.Controls.Add($startBtn)

$stopBtn = New-Object System.Windows.Forms.Button
$stopBtn.Text = "Stop Backend"
$stopBtn.Size = New-Object System.Drawing.Size(126, 32)
$stopBtn.Location = New-Object System.Drawing.Point(146, 178)
$topPanel.Controls.Add($stopBtn)

$forceStopAllBtn = New-Object System.Windows.Forms.Button
$forceStopAllBtn.Text = "Force Stop All"
$forceStopAllBtn.Size = New-Object System.Drawing.Size(130, 32)
$forceStopAllBtn.Location = New-Object System.Drawing.Point(278, 178)
$topPanel.Controls.Add($forceStopAllBtn)

$checkBtn = New-Object System.Windows.Forms.Button
$checkBtn.Text = "Check Now"
$checkBtn.Size = New-Object System.Drawing.Size(110, 32)
$checkBtn.Location = New-Object System.Drawing.Point(412, 178)
$topPanel.Controls.Add($checkBtn)

$saveBtn = New-Object System.Windows.Forms.Button
$saveBtn.Text = "Save Settings"
$saveBtn.Size = New-Object System.Drawing.Size(114, 32)
$saveBtn.Location = New-Object System.Drawing.Point(528, 178)
$topPanel.Controls.Add($saveBtn)

$openTdsPakBtn = New-Object System.Windows.Forms.Button
$openTdsPakBtn.Text = "Open TDS PAK"
$openTdsPakBtn.Size = New-Object System.Drawing.Size(126, 32)
$openTdsPakBtn.Location = New-Object System.Drawing.Point(648, 178)
$topPanel.Controls.Add($openTdsPakBtn)

$openHealthBtn = New-Object System.Windows.Forms.Button
$openHealthBtn.Text = "Open Health URL"
$openHealthBtn.Size = New-Object System.Drawing.Size(130, 32)
$openHealthBtn.Location = New-Object System.Drawing.Point(776, 178)
$topPanel.Controls.Add($openHealthBtn)

$script:StatusLabel = New-Object System.Windows.Forms.Label
$script:StatusLabel.AutoSize = $true
$script:StatusLabel.Location = New-Object System.Drawing.Point(918, 184)
$script:StatusLabel.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
$script:StatusLabel.ForeColor = $script:Theme.TextStrong
$script:StatusLabel.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Right
$topPanel.Controls.Add($script:StatusLabel)

$script:HealthLabel = New-Object System.Windows.Forms.Label
$script:HealthLabel.AutoSize = $true
$script:HealthLabel.Location = New-Object System.Drawing.Point(594, 144)
$script:HealthLabel.ForeColor = $script:Theme.TextMuted
$script:HealthLabel.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Right
$topPanel.Controls.Add($script:HealthLabel)

$tabs = New-Object System.Windows.Forms.TabControl
$tabs.Dock = "Fill"
$tabs.Font = New-Object System.Drawing.Font("Segoe UI Semibold", 9.5, [System.Drawing.FontStyle]::Regular)
$tabs.Padding = New-Object System.Drawing.Point(16, 6)
$form.Controls.Add($tabs)

$backendTab = New-Object System.Windows.Forms.TabPage
$backendTab.Text = "Backend Activity"
$backendTab.BackColor = [System.Drawing.Color]::FromArgb(248, 251, 255)
$tabs.TabPages.Add($backendTab)

$script:LogTextBox = New-Object System.Windows.Forms.TextBox
$script:LogTextBox.Multiline = $true
$script:LogTextBox.ScrollBars = "Both"
$script:LogTextBox.WordWrap = $false
$script:LogTextBox.ReadOnly = $true
$script:LogTextBox.Dock = "Fill"
$script:LogTextBox.Font = New-Object System.Drawing.Font("Consolas", 9.25)
$script:LogTextBox.BackColor = $script:Theme.LogBg
$script:LogTextBox.ForeColor = $script:Theme.LogFg
$script:LogTextBox.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
$backendTab.Controls.Add($script:LogTextBox)

$addonsTab = New-Object System.Windows.Forms.TabPage
$addonsTab.Text = "Database Add-ons"
$addonsTab.BackColor = [System.Drawing.Color]::FromArgb(248, 251, 255)
$tabs.TabPages.Add($addonsTab)

$addonsIntro = New-Object System.Windows.Forms.Label
$addonsIntro.Text = "Future add-ons for landfill, transfer station, and custom local data packages can be managed here."
$addonsIntro.AutoSize = $true
$addonsIntro.Location = New-Object System.Drawing.Point(14, 16)
$addonsIntro.ForeColor = $script:Theme.TextMuted
$addonsTab.Controls.Add($addonsIntro)

$runIndexerBtn = New-Object System.Windows.Forms.Button
$runIndexerBtn.Text = "Open Street Indexer"
$runIndexerBtn.Size = New-Object System.Drawing.Size(170, 32)
$runIndexerBtn.Location = New-Object System.Drawing.Point(18, 52)
$addonsTab.Controls.Add($runIndexerBtn)

$runAutoSetupBtn = New-Object System.Windows.Forms.Button
$runAutoSetupBtn.Text = "Open Full Auto Setup"
$runAutoSetupBtn.Size = New-Object System.Drawing.Size(170, 32)
$runAutoSetupBtn.Location = New-Object System.Drawing.Point(194, 52)
$addonsTab.Controls.Add($runAutoSetupBtn)

$placeholderBox = New-Object System.Windows.Forms.GroupBox
$placeholderBox.Text = "Planned Add-on Slots"
$placeholderBox.Location = New-Object System.Drawing.Point(18, 98)
$placeholderBox.Size = New-Object System.Drawing.Size(560, 170)
$placeholderBox.ForeColor = $script:Theme.TextStrong
$placeholderBox.Font = New-Object System.Drawing.Font("Segoe UI Semibold", 9.0, [System.Drawing.FontStyle]::Regular)
$addonsTab.Controls.Add($placeholderBox)

$slotText = New-Object System.Windows.Forms.Label
$slotText.Text = "- Landfill metadata refresh`r`n- Transfer station catalog`r`n- Fleet yard / depot layers`r`n- Custom company GIS package connectors"
$slotText.AutoSize = $true
$slotText.Location = New-Object System.Drawing.Point(16, 28)
$slotText.ForeColor = $script:Theme.TextMuted
$placeholderBox.Controls.Add($slotText)

Set-InputControlStyle -Control $script:DbPathTextBox
Set-InputControlStyle -Control $script:HostTextBox
Set-InputControlStyle -Control $script:PortTextBox
Set-InputControlStyle -Control $script:TdsPakUrlTextBox
Set-InputControlStyle -Control $script:AutoStartCheckBox

Set-ActionButtonStyle -Button $startBtn -Background $script:Theme.Accent -Foreground ([System.Drawing.Color]::White)
Set-ActionButtonStyle -Button $stopBtn -Background $script:Theme.Danger -Foreground ([System.Drawing.Color]::White)
Set-ActionButtonStyle -Button $forceStopAllBtn -Background $script:Theme.Danger -Foreground ([System.Drawing.Color]::White)
Set-SecondaryButtonStyle -Button $checkBtn
Set-SecondaryButtonStyle -Button $saveBtn
Set-SecondaryButtonStyle -Button $openTdsPakBtn
Set-SecondaryButtonStyle -Button $openHealthBtn
Set-SecondaryButtonStyle -Button $browseDbBtn
Set-SecondaryButtonStyle -Button $runIndexerBtn
Set-SecondaryButtonStyle -Button $runAutoSetupBtn

$toolTip = New-Object System.Windows.Forms.ToolTip
$toolTip.IsBalloon = $true
$toolTip.ToolTipTitle = "TDS PAK Backend Manager"
$toolTip.SetToolTip($startBtn, "Starts tds-street-backend.py with the settings shown above.")
$toolTip.SetToolTip($stopBtn, "Stops the backend process started by this manager.")
$toolTip.SetToolTip($forceStopAllBtn, "Force-stops all detected TDS street backend instances, including orphaned old installs.")
$toolTip.SetToolTip($checkBtn, "Runs an immediate health check against /api/health.")
$toolTip.SetToolTip($script:DbPathTextBox, "Path to the SQLite street index used by the backend.")

$browseDbBtn.Add_Click({
  $dlg = New-Object System.Windows.Forms.OpenFileDialog
  $dlg.Title = "Select SQLite Streets Index"
  $dlg.Filter = "SQLite DB (*.sqlite;*.db)|*.sqlite;*.db|All Files (*.*)|*.*"
  $dlg.Multiselect = $false
  $dlg.CheckFileExists = $true
  if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
    $script:DbPathTextBox.Text = $dlg.FileName
  }
})

$saveBtn.Add_Click({
  $cfg = Get-CurrentConfigFromUi
  Save-ManagerConfig -config $cfg
  Write-ManagerLog "Settings saved."
  Update-BackendStatusUi -forceLog:$true
})

$startBtn.Add_Click({
  Start-ManagedBackend
})

$stopBtn.Add_Click({
  Stop-ManagedBackend
})

$forceStopAllBtn.Add_Click({
  Force-StopAllBackendInstances
})

$checkBtn.Add_Click({
  Update-BackendStatusUi -forceLog:$true
})

$openTdsPakBtn.Add_Click({
  $cfg = Get-CurrentConfigFromUi
  Save-ManagerConfig -config $cfg
  $url = Resolve-TdsPakUrl -inputUrl ([string]$cfg.tds_pak_url)
  $script:TdsPakUrlTextBox.Text = $url
  try {
    Start-Process $url | Out-Null
  } catch {
    [System.Windows.Forms.MessageBox]::Show(
      "Could not open URL: $url",
      "TDS PAK Backend Manager",
      [System.Windows.Forms.MessageBoxButtons]::OK,
      [System.Windows.Forms.MessageBoxIcon]::Warning
    ) | Out-Null
  }
})

$openHealthBtn.Add_Click({
  $backendHost = [string]$script:HostTextBox.Text
  if (-not $backendHost) { $backendHost = "127.0.0.1" }
  $port = 0
  [void][int]::TryParse([string]$script:PortTextBox.Text, [ref]$port)
  if ($port -le 0) { $port = 8787 }
  if ($port -eq 5500) {
    Write-ManagerLog "Open Health URL detected dev-server port 5500; using backend port 8787 instead."
    $port = 8787
    try { $script:PortTextBox.Text = "8787" } catch {}
  }
  $url = "http://$backendHost`:$port/api/health"
  try {
    Start-Process $url | Out-Null
  } catch {}
})

$runIndexerBtn.Add_Click({
  if (-not (Test-Path -LiteralPath $script:IndexerLauncherPath)) {
    [System.Windows.Forms.MessageBox]::Show(
      "Could not find indexer launcher:`r`n$script:IndexerLauncherPath",
      "TDS PAK Backend Manager",
      [System.Windows.Forms.MessageBoxButtons]::OK,
      [System.Windows.Forms.MessageBoxIcon]::Warning
    ) | Out-Null
    return
  }
  Start-Process -FilePath $script:IndexerLauncherPath | Out-Null
  Write-ManagerLog "Opened Street Indexer launcher."
})

$runAutoSetupBtn.Add_Click({
  if (-not (Test-Path -LiteralPath $script:AutoSetupLauncherPath)) {
    [System.Windows.Forms.MessageBox]::Show(
      "Could not find auto setup launcher:`r`n$script:AutoSetupLauncherPath",
      "TDS PAK Backend Manager",
      [System.Windows.Forms.MessageBoxButtons]::OK,
      [System.Windows.Forms.MessageBoxIcon]::Warning
    ) | Out-Null
    return
  }
  Start-Process -FilePath $script:AutoSetupLauncherPath | Out-Null
  Write-ManagerLog "Opened full auto setup launcher."
})

$timer = New-Object System.Windows.Forms.Timer
$timer.Interval = 2200
$timer.Add_Tick({
  Read-BackendRuntimeLogTail
  Update-BackendStatusUi -forceLog:$false
})

$form.Add_Shown({
  Write-ManagerLog "Backend manager loaded."
  Write-ManagerLog "Install directory: $script:InstallDir"
  Write-ManagerLog "Config path: $script:ConfigPath"
  Update-BackendStatusUi -forceLog:$true
  if ($script:AutoStartCheckBox.Checked) {
    Write-ManagerLog "Auto-start is enabled. Starting backend..."
    Start-ManagedBackend
  }
  $timer.Start()
})

$form.Add_FormClosing({
  $cfg = Get-CurrentConfigFromUi
  Save-ManagerConfig -config $cfg
  Read-BackendRuntimeLogTail
})

[void]$form.ShowDialog()
