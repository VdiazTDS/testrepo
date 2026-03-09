param(
  [string]$DataDir,
  [string]$ZipPath,
  [string]$GeoJsonPath,
  [string]$DbPath,
  [switch]$ForceDownload,
  [switch]$SkipDownload,
  [switch]$SkipConvert,
  [switch]$SkipIndex,
  [switch]$SkipBackendStart,
  [switch]$OpenTdsPakAfterSetup,
  [string]$TdsPakUrl = "http://127.0.0.1:5500/tds-pak.html"
)

$ErrorActionPreference = "Stop"

$downloadUrl = "https://download.geofabrik.de/north-america/us/texas-latest-free.shp.zip"
$backendUrl = "http://127.0.0.1:8787/api/health"

function Write-Step([string]$Message) {
  Write-Host "[TDS Auto Setup] $Message"
}

function Show-ZipFilePicker {
  try {
    Add-Type -AssemblyName System.Windows.Forms | Out-Null
    $dialog = New-Object System.Windows.Forms.OpenFileDialog
    $dialog.Title = "Select Texas streets ZIP"
    $dialog.Filter = "ZIP files (*.zip)|*.zip|All files (*.*)|*.*"
    $dialog.Multiselect = $false
    $dialog.CheckFileExists = $true
    $result = $dialog.ShowDialog()
    if ($result -ne [System.Windows.Forms.DialogResult]::OK) {
      return $null
    }
    return $dialog.FileName
  } catch {
    return $null
  }
}

function Find-AutoZipCandidate([string]$DataDirPath) {
  $candidates = @()
  $userHome = [Environment]::GetFolderPath("UserProfile")
  $downloads = if ($userHome) { Join-Path $userHome "Downloads" } else { "" }
  $scanDirs = @()
  if ($DataDirPath -and (Test-Path -LiteralPath $DataDirPath)) { $scanDirs += $DataDirPath }
  if ($downloads -and (Test-Path -LiteralPath $downloads)) { $scanDirs += $downloads }

  foreach ($dir in $scanDirs) {
    $matches = Get-ChildItem -Path $dir -File -Filter *.zip -ErrorAction SilentlyContinue | Where-Object {
      $n = $_.Name.ToLowerInvariant()
      return (
        $n -eq "texas-latest-free.shp.zip" -or
        $n -like "texas-*-free.shp.zip" -or
        ($n -like "*texas*" -and $n -like "*free*" -and $n -like "*.shp.zip")
      )
    }
    if ($matches) {
      $candidates += $matches
    }
  }

  if (-not $candidates -or $candidates.Count -eq 0) {
    return $null
  }

  $best = $candidates |
    Sort-Object -Property @{Expression = "LastWriteTime"; Descending = $true}, @{Expression = "Length"; Descending = $true} |
    Select-Object -First 1

  if (-not $best) { return $null }
  return $best.FullName
}

function Test-GeoJsonLooksAttributeEmpty([string]$GeoJsonPathValue) {
  if (-not (Test-Path -LiteralPath $GeoJsonPathValue)) {
    return $false
  }
  $stream = $null
  $reader = $null
  try {
    $stream = [System.IO.File]::Open($GeoJsonPathValue, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::ReadWrite)
    $bufSize = [Math]::Min(786432, [int]$stream.Length)
    if ($bufSize -le 0) { return $false }
    $bytes = New-Object byte[] $bufSize
    $read = $stream.Read($bytes, 0, $bufSize)
    if ($read -le 0) { return $false }
    $text = [System.Text.Encoding]::UTF8.GetString($bytes, 0, $read)
    $matches = [System.Text.RegularExpressions.Regex]::Matches($text, '"highway":"([^"]*)"')
    if ($matches.Count -eq 0) { return $false }
    $known = 0
    foreach ($m in $matches) {
      $v = [string]$m.Groups[1].Value
      if ($v -and $v -ne "Unknown") {
        $known += 1
      }
    }
    return ($known -eq 0)
  } catch {
    return $false
  } finally {
    if ($reader) { $reader.Dispose() }
    if ($stream) { $stream.Dispose() }
  }
}

function Get-PythonInvocation {
  $candidates = @(
    @{ Command = "py"; PrefixArgs = @("-3") },
    @{ Command = "python"; PrefixArgs = @() },
    @{ Command = "python3"; PrefixArgs = @() }
  )

  foreach ($candidate in $candidates) {
    $cmdName = [string]$candidate.Command
    $prefix = @($candidate.PrefixArgs)
    $found = Get-Command $cmdName -ErrorAction SilentlyContinue
    if (-not $found) { continue }
    try {
      & $cmdName @($prefix + @("-c", "import sys; print(sys.version)")) *> $null
      if ($LASTEXITCODE -eq 0) {
        return @{
          Command = $cmdName
          PrefixArgs = $prefix
        }
      }
    } catch {}
  }
  return $null
}

function Test-BackendHealth {
  try {
    $response = Invoke-RestMethod -Uri $backendUrl -Method Get -TimeoutSec 3
    return ($null -ne $response)
  } catch {
    return $false
  }
}

function Wait-ForBackend([int]$TimeoutSeconds) {
  $start = Get-Date
  while (((Get-Date) - $start).TotalSeconds -lt $TimeoutSeconds) {
    if (Test-BackendHealth) { return $true }
    Start-Sleep -Milliseconds 600
  }
  return $false
}

try {
  $scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
  if (-not $scriptDir) {
    throw "Could not resolve script directory."
  }

  if (-not $DataDir) {
    $docs = [Environment]::GetFolderPath("MyDocuments")
    if (-not $docs) {
      $docs = $scriptDir
    }
    $DataDir = Join-Path $docs "TDS-Pak-Street-Data"
  }

  if (-not (Test-Path -LiteralPath $DataDir)) {
    New-Item -Path $DataDir -ItemType Directory -Force | Out-Null
  }

  if (-not $ZipPath) {
    $ZipPath = Find-AutoZipCandidate -DataDirPath $DataDir
    if ($ZipPath) {
      Write-Step "Auto-detected ZIP source: $ZipPath"
    } else {
      $ZipPath = Join-Path $DataDir "texas-latest-free.shp.zip"
    }
  }
  if (-not $GeoJsonPath) {
    $GeoJsonPath = Join-Path $DataDir "texas-roads.geojson"
  }
  if (-not $DbPath) {
    $DbPath = Join-Path $DataDir "tds-streets.sqlite"
  }

  Write-Step "Data directory: $DataDir"
  Write-Step "ZIP file: $ZipPath"
  Write-Step "GeoJSON file: $GeoJsonPath"
  Write-Step "SQLite index: $DbPath"

  if (-not $SkipDownload) {
    if ($ForceDownload -or -not (Test-Path -LiteralPath $ZipPath)) {
      Write-Step "Downloading Texas streets ZIP (this is large)..."
      Invoke-WebRequest -Uri $downloadUrl -OutFile $ZipPath
    } else {
      Write-Step "Using existing ZIP file."
    }
  } else {
    Write-Step "Skipping ZIP download by request."
  }

  if (-not (Test-Path -LiteralPath $ZipPath)) {
    Write-Step "ZIP file not found. Please pick your downloaded Texas ZIP now."
    $pickedZip = Show-ZipFilePicker
    if ($pickedZip) {
      $ZipPath = $pickedZip
      Write-Step "Using selected ZIP: $ZipPath"
    } else {
      throw "ZIP file not found: $ZipPath"
    }
  }

  $converterScript = Join-Path $scriptDir "tds-streets-offline-converter.ps1"
  if (-not (Test-Path -LiteralPath $converterScript)) {
    throw "Converter script not found: $converterScript"
  }

  $shouldConvert = -not $SkipConvert
  if ($shouldConvert -and -not $ForceDownload -and (Test-Path -LiteralPath $GeoJsonPath)) {
    try {
      $zipTime = (Get-Item -LiteralPath $ZipPath).LastWriteTimeUtc
      $geoTime = (Get-Item -LiteralPath $GeoJsonPath).LastWriteTimeUtc
      if ($geoTime -ge $zipTime) {
        $shouldConvert = $false
      }
    } catch {}
  }
  if (-not $ForceDownload -and (Test-GeoJsonLooksAttributeEmpty -GeoJsonPathValue $GeoJsonPath)) {
    Write-Step "Existing GeoJSON appears to have placeholder street attributes. Rebuilding conversion."
    $shouldConvert = $true
  }

  if ($shouldConvert) {
    Write-Step "Converting ZIP -> GeoJSON..."
    powershell -NoProfile -ExecutionPolicy Bypass -STA -File $converterScript -ZipPath $ZipPath -OutputGeoJson $GeoJsonPath
    if ($LASTEXITCODE -ne 0) {
      throw "Conversion failed (exit code $LASTEXITCODE)."
    }
  } elseif ($SkipConvert) {
    Write-Step "Skipping conversion by request."
  } else {
    Write-Step "GeoJSON already up-to-date. Skipping conversion."
  }

  if (-not (Test-Path -LiteralPath $GeoJsonPath)) {
    throw "GeoJSON output was not created: $GeoJsonPath"
  }
  if (Test-GeoJsonLooksAttributeEmpty -GeoJsonPathValue $GeoJsonPath) {
    throw (
      "Converted GeoJSON still has placeholder street attributes. " +
      "Install Python or Node.js and rerun setup, then verify conversion output."
    )
  }

  $python = Get-PythonInvocation
  if (-not $python) {
    throw "Python was not found. Install Python from https://www.python.org/downloads/"
  }

  $indexerScript = Join-Path $scriptDir "tds-street-indexer.py"
  if (-not (Test-Path -LiteralPath $indexerScript)) {
    throw "Indexer script not found: $indexerScript"
  }

  $shouldIndex = -not $SkipIndex
  if ($shouldIndex -and (Test-Path -LiteralPath $DbPath) -and (Test-Path -LiteralPath $GeoJsonPath)) {
    try {
      $dbTime = (Get-Item -LiteralPath $DbPath).LastWriteTimeUtc
      $geoTime = (Get-Item -LiteralPath $GeoJsonPath).LastWriteTimeUtc
      if ($dbTime -ge $geoTime) {
        $shouldIndex = $false
      }
    } catch {}
  }

  if ($shouldIndex) {
    Write-Step "Indexing GeoJSON -> SQLite + RTree (this can take a while)..."
    & $python.Command @($python.PrefixArgs + @($indexerScript, $GeoJsonPath, "--db", $DbPath, "--source-name", "Texas streets"))
    if ($LASTEXITCODE -ne 0) {
      throw "Indexer failed (exit code $LASTEXITCODE)."
    }
  } elseif ($SkipIndex) {
    Write-Step "Skipping index build by request."
  } else {
    Write-Step "SQLite index already up-to-date. Skipping index build."
  }

  if (-not (Test-Path -LiteralPath $DbPath)) {
    throw "Indexed database was not created: $DbPath"
  }

  if (-not $SkipBackendStart) {
    $backendLauncher = Join-Path $scriptDir "tds-street-backend-launcher.cmd"
    if (-not (Test-Path -LiteralPath $backendLauncher)) {
      throw "Backend launcher not found: $backendLauncher"
    }

    if (Test-BackendHealth) {
      Write-Step "Backend already running at $backendUrl"
    } else {
      Write-Step "Starting backend in a separate terminal window..."
      Start-Process -FilePath $backendLauncher -ArgumentList @("`"$DbPath`"")
      if (-not (Wait-ForBackend -TimeoutSeconds 20)) {
        Write-Step "Backend started, but health endpoint did not respond within 20s."
      } else {
        Write-Step "Backend is online."
      }
    }
  }

  Write-Host ""
  Write-Host "Setup complete." -ForegroundColor Green
  Write-Host "1) Open TDS PAK."
  Write-Host "2) Click 'Check Backend'."
  Write-Host "3) Turn on 'Street Segments (Local Source)'."
  Write-Host ""
  Write-Host "Backend URL: http://127.0.0.1:8787"
  Write-Host "DB path: $DbPath"

  if ($OpenTdsPakAfterSetup) {
    try {
      Write-Step "Opening TDS PAK in browser..."
      Start-Process $TdsPakUrl
    } catch {
      Write-Step "Could not open browser automatically. Open this URL manually: $TdsPakUrl"
    }
  }
}
catch {
  Write-Host ""
  Write-Host "ERROR: $($_.Exception.Message)" -ForegroundColor Red
  Write-Host ""
  Write-Host "Troubleshooting:"
  Write-Host "1) Ensure Python is installed (python.org)."
  Write-Host "2) Ensure these files are together: converter, indexer, backend scripts."
  Write-Host "3) Retry with: -ForceDownload if ZIP is corrupted."
  exit 1
}
