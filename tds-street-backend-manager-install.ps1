param(
  [string]$InstallDir,
  [switch]$Quiet,
  [switch]$LaunchManager
)

$ErrorActionPreference = "Stop"
Add-Type -AssemblyName System.Drawing
Add-Type @"
using System;
using System.Runtime.InteropServices;
public static class TdsIconNative {
  [DllImport("user32.dll", CharSet = CharSet.Auto)]
  public static extern bool DestroyIcon(IntPtr handle);
}
"@
$script:ManagerIconVersion = "2026.03.10.1"

function Write-Install([string]$Message) {
  Write-Host "[TDS Backend Manager Installer] $Message"
}

function New-Shortcut([string]$ShortcutPath, [string]$TargetPath, [string]$WorkingDir, [string]$Description) {
  $shell = New-Object -ComObject WScript.Shell
  $shortcut = $shell.CreateShortcut($ShortcutPath)
  $shortcut.TargetPath = $TargetPath
  $shortcut.WorkingDirectory = $WorkingDir
  $shortcut.Description = $Description
  $iconPath = Join-Path $WorkingDir "tds-street-backend-manager.ico"
  if (Test-Path -LiteralPath $iconPath) {
    $shortcut.IconLocation = "$iconPath,0"
  } else {
    $shortcut.IconLocation = "$env:SystemRoot\System32\shell32.dll,25"
  }
  $shortcut.Save()
}

function New-RoundedRectPath([System.Drawing.RectangleF]$Rect, [float]$Radius) {
  $diameter = [Math]::Max(1, [Math]::Min([Math]::Min($Rect.Width, $Rect.Height), $Radius * 2))
  $path = New-Object System.Drawing.Drawing2D.GraphicsPath
  $arcRect = New-Object System.Drawing.RectangleF($Rect.X, $Rect.Y, $diameter, $diameter)
  $path.AddArc($arcRect, 180, 90)
  $arcRect.X = $Rect.Right - $diameter
  $path.AddArc($arcRect, 270, 90)
  $arcRect.Y = $Rect.Bottom - $diameter
  $path.AddArc($arcRect, 0, 90)
  $arcRect.X = $Rect.X
  $path.AddArc($arcRect, 90, 90)
  $path.CloseFigure()
  return $path
}

function Ensure-ManagerIcon([string]$InstallDirPath) {
  $iconPath = Join-Path $InstallDirPath "tds-street-backend-manager.ico"
  $pngPath = Join-Path $InstallDirPath "tds-street-backend-manager.png"
  $versionPath = Join-Path $InstallDirPath "tds-street-backend-manager-icon.version"

  if ((Test-Path -LiteralPath $iconPath) -and (Test-Path -LiteralPath $versionPath)) {
    try {
      $existingVersion = (Get-Content -LiteralPath $versionPath -Raw).Trim()
      if ($existingVersion -eq $script:ManagerIconVersion) {
        return $iconPath
      }
    } catch {}
  }

  foreach ($path in @($iconPath, $pngPath, $versionPath)) {
    if (Test-Path -LiteralPath $path) {
      try { Remove-Item -LiteralPath $path -Force -ErrorAction Stop } catch {}
    }
  }

  $size = 256
  $bmp = New-Object System.Drawing.Bitmap($size, $size)
  $g = [System.Drawing.Graphics]::FromImage($bmp)
  $hIcon = [IntPtr]::Zero
  $icon = $null
  try {
    $g.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
    $g.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
    $g.CompositingQuality = [System.Drawing.Drawing2D.CompositingQuality]::HighQuality
    $g.Clear([System.Drawing.Color]::Transparent)

    $bgRect = New-Object System.Drawing.RectangleF(10, 10, 236, 236)
    $bgPath = New-RoundedRectPath -Rect $bgRect -Radius 42
    try {
      $brush = New-Object System.Drawing.Drawing2D.LinearGradientBrush(
        (New-Object System.Drawing.PointF(0, 0)),
        (New-Object System.Drawing.PointF(256, 256)),
        ([System.Drawing.Color]::FromArgb(34, 93, 201)),
        ([System.Drawing.Color]::FromArgb(35, 170, 220))
      )
      $g.FillPath($brush, $bgPath)
      $brush.Dispose()

      $rackPen = New-Object System.Drawing.Pen([System.Drawing.Color]::FromArgb(240, 248, 255), 8)
      $rackPen.LineJoin = [System.Drawing.Drawing2D.LineJoin]::Round
      $rackBrush = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::FromArgb(242, 249, 255))
      for ($i = 0; $i -lt 3; $i++) {
        $y = 64 + ($i * 48)
        $slot = New-Object System.Drawing.RectangleF(56, $y, 144, 34)
        $slotPath = New-RoundedRectPath -Rect $slot -Radius 11
        $g.FillPath($rackBrush, $slotPath)
        $g.DrawPath($rackPen, $slotPath)
        $slotPath.Dispose()
      }
      $rackBrush.Dispose()
      $rackPen.Dispose()

      $statusBrush = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::FromArgb(65, 214, 129))
      $g.FillEllipse($statusBrush, 202, 164, 26, 26)
      $statusBrush.Dispose()

      $signalPen = New-Object System.Drawing.Pen([System.Drawing.Color]::FromArgb(242, 249, 255), 8)
      $signalPen.StartCap = [System.Drawing.Drawing2D.LineCap]::Round
      $signalPen.EndCap = [System.Drawing.Drawing2D.LineCap]::Round
      $g.DrawLine($signalPen, 72, 38, 184, 38)
      $g.DrawLine($signalPen, 196, 38, 218, 38)
      $g.DrawLine($signalPen, 196, 38, 196, 58)
      $signalPen.Dispose()
    } finally {
      $bgPath.Dispose()
    }

    $bmp.Save($pngPath, [System.Drawing.Imaging.ImageFormat]::Png)

    $hIcon = $bmp.GetHicon()
    $icon = [System.Drawing.Icon]::FromHandle($hIcon)
    $stream = [System.IO.File]::Open($iconPath, [System.IO.FileMode]::Create, [System.IO.FileAccess]::Write)
    try {
      $icon.Save($stream)
    } finally {
      $stream.Dispose()
    }

    Set-Content -LiteralPath $versionPath -Value $script:ManagerIconVersion -Encoding ASCII
  } finally {
    if ($icon) { $icon.Dispose() }
    if ($hIcon -ne [IntPtr]::Zero) { [void][TdsIconNative]::DestroyIcon($hIcon) }
    if ($g) { $g.Dispose() }
    if ($bmp) { $bmp.Dispose() }
  }
  return $iconPath
}

try {
  $scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
  if (-not $scriptDir) {
    throw "Could not resolve installer directory."
  }

  if (-not $InstallDir) {
    $InstallDir = Join-Path $env:LOCALAPPDATA "TDS-Pak\StreetBackendManager"
  }
  if (-not (Test-Path -LiteralPath $InstallDir)) {
    New-Item -Path $InstallDir -ItemType Directory -Force | Out-Null
  }

  $filesToInstall = @(
    "tds-street-backend-manager.ps1",
    "tds-street-backend-manager-launcher.cmd",
    "tds-local-streets-auto-setup.ps1",
    "tds-local-streets-auto-setup-launcher.cmd",
    "tds-local-backend-quickstart.txt",
    "tds-street-backend.py",
    "tds-street-backend-launcher.cmd",
    "tds-street-indexer.py",
    "tds-street-indexer-launcher.cmd",
    "tds-streets-offline-converter.ps1",
    "tds-streets-offline-converter-launcher.cmd"
  )

  $copied = 0
  foreach ($fileName in $filesToInstall) {
    $sourcePath = Join-Path $scriptDir $fileName
    if (-not (Test-Path -LiteralPath $sourcePath)) {
      Write-Install "Skipping missing file: $fileName"
      continue
    }
    Copy-Item -LiteralPath $sourcePath -Destination (Join-Path $InstallDir $fileName) -Force
    $copied += 1
  }

  $launcherPath = Join-Path $InstallDir "tds-street-backend-manager-launcher.cmd"
  if (-not (Test-Path -LiteralPath $launcherPath)) {
    throw "Manager launcher was not copied to install directory: $launcherPath"
  }

  $iconPath = Ensure-ManagerIcon -InstallDirPath $InstallDir

  $desktopPath = [Environment]::GetFolderPath("Desktop")
  if ($desktopPath) {
    $desktopShortcut = Join-Path $desktopPath "TDS PAK Street Backend Manager.lnk"
    New-Shortcut -ShortcutPath $desktopShortcut -TargetPath $launcherPath -WorkingDir $InstallDir -Description "Open TDS PAK Street Backend Manager"
  }

  $programsPath = [Environment]::GetFolderPath("Programs")
  if ($programsPath) {
    $menuFolder = Join-Path $programsPath "TDS PAK"
    if (-not (Test-Path -LiteralPath $menuFolder)) {
      New-Item -Path $menuFolder -ItemType Directory -Force | Out-Null
    }
    $menuShortcut = Join-Path $menuFolder "Street Backend Manager.lnk"
    New-Shortcut -ShortcutPath $menuShortcut -TargetPath $launcherPath -WorkingDir $InstallDir -Description "Open TDS PAK Street Backend Manager"
  }

  Write-Install "Installed/updated files: $copied"
  Write-Install "Install directory: $InstallDir"
  if ($iconPath) {
    Write-Install "Icon file: $iconPath"
  }
  Write-Install "Desktop shortcut: TDS PAK Street Backend Manager"

  if ($LaunchManager) {
    Write-Install "Opening Street Backend Manager..."
    Start-Process -FilePath $launcherPath
  }

  if (-not $Quiet) {
    Write-Host ""
    Write-Host "Install complete." -ForegroundColor Green
    Write-Host "Open 'TDS PAK Street Backend Manager' from Desktop or Start Menu."
    Write-Host ""
  }
}
catch {
  Write-Host ""
  Write-Host "ERROR: $($_.Exception.Message)" -ForegroundColor Red
  exit 1
}
