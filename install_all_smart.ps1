# install_all_smart.ps1
# SMART installer (anti-stuck):
# - Ensure winget
# - Install/upgrade Python Install Manager (winget source=winget)
# - Install Python 3.14 (winget source=winget)
# - Add Python + Scripts to USER PATH (auto detect latest Python*)
# - Install Google Chrome (winget source=winget)
# - Upgrade pip (pakai py -m pip)
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

# -------------------------
# Winget helpers (anti stuck)
# -------------------------
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
  # Jalankan winget dengan output live + timeout (anti hang)
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

  # flush sisa output/error
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

  # Anti prompt & anti nyangkut msstore: pakai source winget + accept agreements
  $args = "install -e --id $id --source $source --accept-package-agreements --accept-source-agreements"

  # Kadang --silent bikin terlihat "diam", jadi default: tanpa --silent biar output jalan.
  # Kalau mau silent, ganti baris ini:
  # $args += " --silent"
  Invoke-Winget $args 900

  Write-Host "Done: $nameForLog" -ForegroundColor Green
}

function Ensure-Python {
  Step "Ensure Python Install Manager"
  Winget-Install-IfMissing "Python.PythonInstallManager" "Python Install Manager" "winget"

  Step "Ensure Python 3.14"
  Winget-Install-IfMissing "Python.Python.3.14" "Python 3.14" "winget"

  Step "Check py launcher"
  if (-not (Has-Command "py")) {
    Write-Host "Perintah 'py' belum terdeteksi. Tutup PowerShell dan buka lagi, lalu jalankan ulang script." -ForegroundColor Yellow
    Write-Host "Kalau masih tidak ada, pastikan instalasi Python membuat py launcher tersedia." -ForegroundColor Yellow
    exit 1
  }

  Write-Host "Python OK:"
  py -V
}

function Ensure-Chrome {
  Step "Ensure Google Chrome"

  if (Winget-Installed "Google.Chrome" "winget") {
    Write-Host "Skip: Google Chrome sudah terpasang (winget)."
    return
  }

  # Fallback: cek file chrome.exe
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
  py -m pip install --upgrade pip
  py -m pip --version
}

function Pip-Package-Installed($pkg) {
  try {
    py -m pip show $pkg 1>$null 2>$null
    return $true
  }
  catch {
    return $false
  }
}

function Ensure-Pip-Packages {
  Step "Install pip packages (only missing)"
  $packages = @("pandas", "openpyxl", "selenium", "webdriver-manager")

  foreach ($p in $packages) {
    if (Pip-Package-Installed $p) {
      Write-Host "Skip: $p sudah terpasang."
    }
    else {
      Write-Host "Install: $p ..." -ForegroundColor Yellow
      py -m pip install $p
      Write-Host "Done: $p" -ForegroundColor Green
    }
  }

  Step "Freeze (verifikasi)"
  py -m pip show pandas openpyxl selenium webdriver-manager | Select-String "Name|Version"
}

# -------------------------
# Ensure PATH (anti error "-and")
# -------------------------
function Ensure-Py-Path {
  Step "Ensure Python PATH (User)"

  $pyBase = Join-Path $env:LOCALAPPDATA "Programs\Python"
  if (-not (Test-Path $pyBase)) {
    Write-Host "Python base folder not found: $pyBase"
    return
  }

  # Ambil folder Python*
  $pyDirs = Get-ChildItem -Path $pyBase -Directory -ErrorAction SilentlyContinue |
  Where-Object { $_.Name -match '^Python\d+$' }

  if (-not $pyDirs -or $pyDirs.Count -eq 0) {
    Write-Host "No PythonXX folder found under $pyBase"
    return
  }

  # Pilih yang terbesar (Python314 > Python313 > ...)
  $latestPy = $pyDirs |
  Sort-Object @{ Expression = { [int]($_.Name -replace 'Python', '') } } -Descending |
  Select-Object -First 1

  $pyPath = $latestPy.FullName
  $pyScripts = Join-Path $pyPath "Scripts"

  Write-Host "Detected latest Python folder: $pyPath"

  $userPath = [Environment]::GetEnvironmentVariable("PATH", "User")
  if ($null -eq $userPath) { $userPath = "" }

  $changed = $false

  # pakai regex escape biar aman kalau ada karakter spesial
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

    # refresh PATH untuk sesi PowerShell saat ini
    $env:PATH = [Environment]::GetEnvironmentVariable("PATH", "Machine") + ";" +
    [Environment]::GetEnvironmentVariable("PATH", "User")

    Write-Host "PATH Updated (User level) + refreshed for current session." -ForegroundColor Green
  }
  else {
    Write-Host "PATH already OK."
  }
}

# =========================
# MAIN
# =========================
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
  Write-Host "Kamu bisa lanjut jalankan script scraping-mu." -ForegroundColor Green
}
catch {
  Write-Host ""
  Write-Host "ERROR: $($_.Exception.Message)" -ForegroundColor Red
  Write-Host "Kalau ini error permission/winget, coba buka PowerShell 'Run as Administrator' lalu jalankan lagi." -ForegroundColor Yellow
  exit 1
}

Write-Host ""
pause


$ErrorActionPreference = "Stop"

function Step($msg) {
  Write-Host ""
  Write-Host "=== $msg ===" -ForegroundColor Cyan
}

function Has-Command($name) {
  return [bool](Get-Command $name -ErrorAction SilentlyContinue)
}

function Winget-Installed($id) {
  try {
    $out = winget list --id $id --exact 2>$null
    return ($out -match $id)
  }
  catch {
    return $false
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

function Winget-Install-IfMissing($id, $nameForLog) {
  if (Winget-Installed $id) {
    Write-Host "Skip: $nameForLog sudah terpasang ($id)."
    return
  }
  Write-Host "Install: $nameForLog ($id) ..."
  winget install -e --id $id --silent --accept-package-agreements --accept-source-agreements
  Write-Host "Done: $nameForLog"
}

function Ensure-Python {
  Step "Ensure Python Install Manager"
  Winget-Install-IfMissing "Python.PythonInstallManager" "Python Install Manager"

  Step "Ensure Python 3.14"
  # Kamu bisa ganti 3.12 -> 3.13 kalau mau, tapi 3.12 biasanya paling aman untuk library.
  Winget-Install-IfMissing "Python.Python.3.14" "Python 3.14"

  Step "Check py launcher"
  if (-not (Has-Command "py")) {
    Write-Host "Perintah 'py' belum terdeteksi. Tutup PowerShell dan buka lagi, lalu jalankan ulang script." -ForegroundColor Yellow
    Write-Host "Kalau masih tidak ada, pastikan Python terpasang dengan opsi 'Add to PATH'." -ForegroundColor Yellow
    exit 1
  }

  Write-Host "Python OK:"
  py -V
}

function Ensure-Chrome {
  Step "Ensure Google Chrome"
  # Cara paling stabil: cek via winget id
  if (Winget-Installed "Google.Chrome") {
    Write-Host "Skip: Google Chrome sudah terpasang."
    return
  }

  # Fallback: cek file chrome.exe (kadang Chrome diinstall dari sumber lain)
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

  Write-Host "Install: Google Chrome ..."
  winget install -e --id Google.Chrome --silent --accept-package-agreements --accept-source-agreements
  Write-Host "Done: Google Chrome"
}



function Ensure-Pip {
  Step "Upgrade pip"
  py -m pip install --upgrade pip
  py -m pip --version
}

function Pip-Package-Installed($pkg) {
  try {
    py -m pip show $pkg 1>$null 2>$null
    return $true
  }
  catch {
    return $false
  }
}

function Ensure-Pip-Packages {
  Step "Install pip packages (only missing)"
  $packages = @("pandas", "openpyxl", "selenium", "webdriver-manager")

  foreach ($p in $packages) {
    if (Pip-Package-Installed $p) {
      Write-Host "Skip: $p sudah terpasang."
    }
    else {
      Write-Host "Install: $p ..."
      py -m pip install $p
      Write-Host "Done: $p"
    }
  }

  Step "Freeze (verifikasi)"
  py -m pip show pandas openpyxl selenium webdriver-manager | Select-String "Name|Version"
}

function Ensure-Py-Path {
  Write-Host ""
  Write-Host "=== Ensure Python PATH ==="

  $pyBase = Join-Path $env:LOCALAPPDATA "Programs\Python"

  if (Test-Path $pyBase) {

    # Ambil semua folder Python*
    $pyDirs = Get-ChildItem -Path $pyBase -Directory -ErrorAction SilentlyContinue |
    Where-Object { $_.Name -match '^Python\d+$' }

    if (-not $pyDirs -or $pyDirs.Count -eq 0) {
      Write-Host "No PythonXX folder found under $pyBase"
    }
    else {
      # Pilih yang angkanya paling besar (Python314 > Python313 > Python312 ...)
      $latestPy = $pyDirs |
      Sort-Object @{ Expression = { [int]($_.Name -replace 'Python', '') } } -Descending |
      Select-Object -First 1

      $pyPath = $latestPy.FullName
      $pyScripts = Join-Path $pyPath "Scripts"

      Write-Host "Detected latest Python folder: $pyPath"

      $userPath = [Environment]::GetEnvironmentVariable("PATH", "User")
      if ($null -eq $userPath) { $userPath = "" }

      $changed = $false

      if ($userPath -notlike "*$pyPath*") {
        Write-Host "Add Python to PATH: $pyPath"
        $userPath = ($userPath.TrimEnd(';') + ";" + $pyPath)
        $changed = $true
      }

      if ((Test-Path $pyScripts) -and ($userPath -notlike "*$pyScripts*")) {
        Write-Host "Add Python Scripts to PATH: $pyScripts"
        $userPath = ($userPath.TrimEnd(';') + ";" + $pyScripts)
        $changed = $true
      }

      if ($changed) {
        [Environment]::SetEnvironmentVariable("PATH", $userPath, "User")

        # refresh PATH untuk sesi PowerShell saat ini (tanpa restart)
        $env:PATH = [Environment]::GetEnvironmentVariable("PATH", "Machine") + ";" +
        [Environment]::GetEnvironmentVariable("PATH", "User")

        Write-Host "PATH Updated (User level) + refreshed for current session."
      }
      else {
        Write-Host "PATH already OK."
      }
    }
  }
  else {
    Write-Host "Python base folder not found: $pyBase"
  }

}
# =========================
# MAIN
# =========================
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
  Write-Host "Kamu bisa lanjut jalankan script scraping-mu." -ForegroundColor Green
}
catch {
  Write-Host ""
  Write-Host "ERROR: $($_.Exception.Message)" -ForegroundColor Red
  Write-Host "Kalau ini error permission/winget, coba buka PowerShell 'Run as Administrator' lalu jalankan lagi." -ForegroundColor Yellow
  exit 1
}

Write-Host ""
pause
