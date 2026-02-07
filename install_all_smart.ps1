# install_all_smart.ps1
# SMART installer (anti-stuck):
# - Ensure winget
# - Install/upgrade Python Install Manager (winget source=winget)
# - Install Python 3.14 (winget source=winget)
# - Add Python + Scripts to USER PATH (auto detect latest Python*)
# - Install Google Chrome (winget source=winget)
# - Upgrade pip (pakai py.exe atau python.exe yang terdeteksi)
# - Install pip packages hanya yang belum ada: pandas, openpyxl, selenium, webdriver-manager
#
# Jalankan:
# Double click run_install.bat

$ErrorActionPreference = "Stop"

function Step($msg) {
  Write-Host ""
  Write-Host "=== $msg ===" -ForegroundColor Cyan
}

function Has-Command($name) {
  return [bool](Get-Command $name -ErrorAction SilentlyContinue)
}

function Winget-Installed($id, $source = "winget") {
  try {
    $out = winget list --id $id --exact --source $source 2>$null
    return ($out -match [regex]::Escape($id))
  }
  catch {
    return $false
  }
}

function Invoke-Winget($args, [int]$timeoutSec = 900) {
  $psi = New-Object System.Diagnostics.ProcessStartInfo
  $psi.FileName = "winget"
  $psi.Arguments = $args
  $psi.RedirectStandardOutput = $true
  $psi.RedirectStandardError = $true
  $psi.UseShellExecute = $false
  $psi.CreateNoWindow = $true

  $p = New-Object System.Diagnostics.Process
  $p.StartInfo = $psi
  $null = $p.Start()

  $sw = [Diagnostics.Stopwatch]::StartNew()
  while (-not $p.HasExited) {
    Start-Sleep -Milliseconds 250

    while (-not $p.StandardOutput.EndOfStream) {
      $line = $p.StandardOutput.ReadLine()
      if ($line) { Write-Host $line }
    }
    while (-not $p.StandardError.EndOfStream) {
      $line = $p.StandardError.ReadLine()
      if ($line) { Write-Host $line -ForegroundColor DarkRed }
    }

    if ($sw.Elapsed.TotalSeconds -ge $timeoutSec) {
      try { $p.Kill() } catch {}
      throw "winget timeout > ${timeoutSec}s untuk args: $args"
    }
  }

  while (-not $p.StandardOutput.EndOfStream) { $line = $p.StandardOutput.ReadLine(); if ($line) { Write-Host $line } }
  while (-not $p.StandardError.EndOfStream) { $line = $p.StandardError.ReadLine(); if ($line) { Write-Host $line -ForegroundColor DarkRed } }

  if ($p.ExitCode -ne 0) {
    throw "winget gagal (ExitCode=$($p.ExitCode)) args: $args"
  }
}

function Ensure-Winget {
  Step "Check winget"
  if (-not (Has-Command "winget")) {
    Write-Host "Winget tidak ditemukan. Install 'App Installer' dari Microsoft Store dulu." -ForegroundColor Red
    exit 1
  }
  Write-Host "OK: winget tersedia."
}

function Winget-Install-IfMissing($id, $nameForLog, $source = "winget") {
  if (Winget-Installed $id $source) {
    Write-Host "Skip: $nameForLog sudah terpasang ($id)."
    return
  }

  Write-Host "Install: $nameForLog ($id) ..." -ForegroundColor Yellow
  $args = "install -e --id $id --source $source --accept-package-agreements --accept-source-agreements"
  Invoke-Winget $args 900
  Write-Host "Done: $nameForLog" -ForegroundColor Green
}

# -------------------------
# PY RESOLVER (anti-stuck)
# -------------------------
$script:PYRUN = $null

function Find-PyLauncherPath {
  $candidates = @(
    "$Env:WINDIR\py.exe",
    "$Env:WINDIR\System32\py.exe",
    "$Env:LOCALAPPDATA\Programs\Python\Launcher\py.exe",
    "$Env:ProgramFiles\Python\Launcher\py.exe",
    "$Env:ProgramFiles(x86)\Python\Launcher\py.exe"
  )
  foreach ($c in $candidates) {
    if (Test-Path $c) { return $c }
  }
  $cmd = Get-Command py -ErrorAction SilentlyContinue
  if ($cmd) { return $cmd.Source }
  return $null
}

function Find-PythonExeFromLocalAppData {
  $pyBase = Join-Path $env:LOCALAPPDATA "Programs\Python"
  if (-not (Test-Path $pyBase)) { return $null }

  $pythonExe = Get-ChildItem -Path $pyBase -Recurse -Filter python.exe -ErrorAction SilentlyContinue |
  Where-Object { $_.FullName -match '\\Python\d+\\python\.exe$' } |
  Sort-Object FullName -Descending |
  Select-Object -First 1

  if ($pythonExe) { return $pythonExe.FullName }
  return $null
}

function Ensure-PyRunner {
  Step "Resolve Python runner (py or python) - anti stuck"

  $pyExe = Find-PyLauncherPath
  if ($pyExe) {
    $script:PYRUN = $pyExe
    Write-Host "OK: py launcher ditemukan: $pyExe" -ForegroundColor Green
    & $script:PYRUN -V
    return
  }

  $pythonExe = Find-PythonExeFromLocalAppData
  if ($pythonExe) {
    $script:PYRUN = $pythonExe
    Write-Host "WARN: py launcher tidak ditemukan, pakai python.exe: $pythonExe" -ForegroundColor Yellow
    & $script:PYRUN -V
    return
  }

  # Last fallback: coba python dari PATH (kalau ada)
  $cmd = Get-Command python -ErrorAction SilentlyContinue
  if ($cmd) {
    $script:PYRUN = $cmd.Source
    Write-Host "WARN: pakai python dari PATH: $($cmd.Source)" -ForegroundColor Yellow
    & $script:PYRUN -V
    return
  }

  throw "Tidak menemukan py.exe maupun python.exe. Kemungkinan install Python gagal, diblok policy, atau lokasi instalasi berbeda."
}

function Ensure-Python {
  Step "Ensure Python Install Manager"
  Winget-Install-IfMissing "Python.PythonInstallManager" "Python Install Manager" "winget"

  Step "Ensure Python 3.14"
  Winget-Install-IfMissing "Python.Python.3.14" "Python 3.14" "winget"

  # jangan exit kalau 'py' belum kebaca di sesi ini.
  # langsung resolve runner yang benar (py.exe atau python.exe).
  Ensure-PyRunner
}

function Ensure-Py-Path {
  Step "Ensure Python PATH (User)"

  $pyBase = Join-Path $env:LOCALAPPDATA "Programs\Python"
  if (-not (Test-Path $pyBase)) {
    Write-Host "Python base folder not found: $pyBase"
    return
  }

  $pyDirs = Get-ChildItem -Path $pyBase -Directory -ErrorAction SilentlyContinue |
  Where-Object { $_.Name -match '^Python\d+$' }

  if (-not $pyDirs -or $pyDirs.Count -eq 0) {
    Write-Host "No PythonXX folder found under $pyBase"
    return
  }

  $latestPy = $pyDirs |
  Sort-Object @{ Expression = { [int]($_.Name -replace 'Python', '') } } -Descending |
  Select-Object -First 1

  $pyPath = $latestPy.FullName
  $pyScripts = Join-Path $pyPath "Scripts"

  Write-Host "Detected latest Python folder: $pyPath"

  $userPath = [Environment]::GetEnvironmentVariable("PATH", "User")
  if ($null -eq $userPath) { $userPath = "" }

  $changed = $false

  if ($userPath -notmatch [regex]::Escape($pyPath)) {
    Write-Host "Add Python to PATH: $pyPath"
    $userPath = ($userPath.TrimEnd(';') + ";" + $pyPath)
    $changed = $true
  }

  if ((Test-Path $pyScripts) -and ($userPath -notmatch [regex]::Escape($pyScripts))) {
    Write-Host "Add Python Scripts to PATH: $pyScripts"
    $userPath = ($userPath.TrimEnd(';') + ";" + $pyScripts)
    $changed = $true
  }

  if ($changed) {
    [Environment]::SetEnvironmentVariable("PATH", $userPath, "User")
    $env:PATH = [Environment]::GetEnvironmentVariable("PATH", "Machine") + ";" +
    [Environment]::GetEnvironmentVariable("PATH", "User")
    Write-Host "PATH Updated (User level) + refreshed for current session." -ForegroundColor Green
  }
  else {
    Write-Host "PATH already OK."
  }

  # Re-resolve runner setelah PATH di-update (kadang baru kebaca)
  Ensure-PyRunner
}

function Ensure-Chrome {
  Step "Ensure Google Chrome"

  if (Winget-Installed "Google.Chrome" "winget") {
    Write-Host "Skip: Google Chrome sudah terpasang (winget)."
    return
  }

  $chromePaths = @(
    "$Env:ProgramFiles\Google\Chrome\Application\chrome.exe",
    "$Env:ProgramFiles(x86)\Google\Chrome\Application\chrome.exe"
  )
  foreach ($p in $chromePaths) {
    if (Test-Path $p) {
      Write-Host "Skip: Google Chrome terdeteksi di $p"
      return
    }
  }

  Write-Host "Install: Google Chrome ..." -ForegroundColor Yellow
  Invoke-Winget "install -e --id Google.Chrome --source winget --accept-package-agreements --accept-source-agreements" 900
  Write-Host "Done: Google Chrome" -ForegroundColor Green
}

function Ensure-Pip {
  Step "Upgrade pip"
  if (-not $script:PYRUN) { Ensure-PyRunner }

  & $script:PYRUN -m pip install --upgrade pip
  & $script:PYRUN -m pip --version
}

function Pip-Package-Installed($pkg) {
  try {
    if (-not $script:PYRUN) { Ensure-PyRunner }
    & $script:PYRUN -m pip show $pkg 1>$null 2>$null
    return $true
  }
  catch { return $false }
}

function Ensure-Pip-Packages {
  Step "Install pip packages (only missing)"
  if (-not $script:PYRUN) { Ensure-PyRunner }

  $packages = @("pandas", "openpyxl", "selenium", "webdriver-manager")

  foreach ($p in $packages) {
    if (Pip-Package-Installed $p) {
      Write-Host "Skip: $p sudah terpasang."
    }
    else {
      Write-Host "Install: $p ..." -ForegroundColor Yellow
      & $script:PYRUN -m pip install $p
      Write-Host "Done: $p" -ForegroundColor Green
    }
  }

  Step "Verifikasi versi"
  & $script:PYRUN -m pip show pandas openpyxl selenium webdriver-manager | Select-String "Name|Version"
}

try {
  Step "START SMART INSTALLER"
  Ensure-Winget
  Ensure-Python
  Ensure-Py-Path
  Ensure-Chrome
  Ensure-Pip
  Ensure-Pip-Packages

  Step "SELESAI"
  Write-Host "Semua dependency sudah siap." -ForegroundColor Green
}
catch {
  Write-Host ""
  Write-Host "ERROR: $($_.Exception.Message)" -ForegroundColor Red
  Write-Host "Kalau ini error permission/winget, coba buka PowerShell 'Run as Administrator' lalu jalankan lagi." -ForegroundColor Yellow
  exit 1
}

Write-Host ""
pause
