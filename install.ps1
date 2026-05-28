# PM CV Formatter — Installer & Launcher
# Called by start.bat with -ExecutionPolicy Bypass

try {

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$AppDir = Split-Path -Parent $MyInvocation.MyCommand.Path

# ── Progress window ───────────────────────────────────────────────────────────
$form                  = New-Object System.Windows.Forms.Form
$form.Text             = "PM CV Formatter"
$form.ClientSize       = New-Object System.Drawing.Size(460, 170)
$form.StartPosition    = "CenterScreen"
$form.FormBorderStyle  = "FixedSingle"
$form.MaximizeBox      = $false
$form.MinimizeBox      = $false
$form.BackColor        = [System.Drawing.Color]::FromArgb(15, 16, 18)
$form.TopMost          = $true

$gold  = [System.Drawing.Color]::FromArgb(200, 169, 97)
$white = [System.Drawing.Color]::FromArgb(232, 233, 234)
$gray  = [System.Drawing.Color]::FromArgb(120, 122, 126)

$titleLbl            = New-Object System.Windows.Forms.Label
$titleLbl.Text       = "PATRICK MORGAN  ·  CV Formatter"
$titleLbl.Font       = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
$titleLbl.ForeColor  = $gold
$titleLbl.Location   = New-Object System.Drawing.Point(24, 20)
$titleLbl.Size       = New-Object System.Drawing.Size(412, 22)
$form.Controls.Add($titleLbl)

$statusLbl           = New-Object System.Windows.Forms.Label
$statusLbl.Text      = "Starting..."
$statusLbl.Font      = New-Object System.Drawing.Font("Segoe UI", 9)
$statusLbl.ForeColor = $white
$statusLbl.Location  = New-Object System.Drawing.Point(24, 54)
$statusLbl.Size      = New-Object System.Drawing.Size(412, 18)
$form.Controls.Add($statusLbl)

$bar          = New-Object System.Windows.Forms.ProgressBar
$bar.Location = New-Object System.Drawing.Point(24, 82)
$bar.Size     = New-Object System.Drawing.Size(412, 16)
$bar.Minimum  = 0
$bar.Maximum  = 100
$bar.Value    = 0
$form.Controls.Add($bar)

$subLbl           = New-Object System.Windows.Forms.Label
$subLbl.Text      = ""
$subLbl.Font      = New-Object System.Drawing.Font("Segoe UI", 8)
$subLbl.ForeColor = $gray
$subLbl.Location  = New-Object System.Drawing.Point(24, 108)
$subLbl.Size      = New-Object System.Drawing.Size(412, 18)
$form.Controls.Add($subLbl)

$form.Show()
$form.Refresh()

function Step($msg, $sub, $pct) {
    $statusLbl.Text = $msg
    $subLbl.Text    = $sub
    $bar.Value      = [Math]::Min([int]$pct, 100)
    $form.Refresh()
    [System.Windows.Forms.Application]::DoEvents()
}

function Fail($msg) {
    Step $msg "Close this window and contact your IT administrator." 0
    [System.Windows.Forms.MessageBox]::Show($msg, "PM CV Formatter", "OK", "Error") | Out-Null
    $form.Close()
    exit 1
}

# ── If server already running just open browser and exit ─────────────────────
Step "Checking if already running..." "" 5
try {
    $ping = Invoke-WebRequest -Uri "http://localhost:5000" -UseBasicParsing -TimeoutSec 1 -ErrorAction Stop
    if ($ping.StatusCode -eq 200) {
        Start-Process "http://localhost:5000"
        $form.Close()
        exit 0
    }
} catch {}

# ── Find pythonw.exe ──────────────────────────────────────────────────────────
Step "Checking Python installation..." "" 10

$pythonw = $null
$patterns = @(
    "$env:LOCALAPPDATA\Programs\Python\Python3*\pythonw.exe",
    "$env:ProgramFiles\Python3*\pythonw.exe",
    "${env:ProgramFiles(x86)}\Python3*\pythonw.exe"
)

foreach ($p in $patterns) {
    $hit = Get-Item $p -ErrorAction SilentlyContinue |
           Sort-Object Name -Descending |
           Select-Object -First 1
    if ($hit) { $pythonw = $hit.FullName; break }
}

if (-not $pythonw) {
    $pycmd = (Get-Command python -ErrorAction SilentlyContinue).Source
    if ($pycmd) {
        $pw = $pycmd -replace 'python\.exe$', 'pythonw.exe'
        if (Test-Path $pw) { $pythonw = $pw }
    }
}

# ── Install Python if not found — try winget first, then direct download ────────
if (-not $pythonw) {

    # ── Attempt 1: winget (fast, no download) ────────────────────────────────
    $winget = (Get-Command winget -ErrorAction SilentlyContinue).Source
    if ($winget) {
        Step "Installing Python 3.12..." "Using Windows Package Manager — please wait..." 18
        $bar.Style = [System.Windows.Forms.ProgressBarStyle]::Marquee
        $bar.MarqueeAnimationSpeed = 30
        $form.Refresh(); [System.Windows.Forms.Application]::DoEvents()

        Start-Process $winget -ArgumentList "install Python.Python.3.12 --silent --accept-source-agreements --accept-package-agreements" -Wait -WindowStyle Hidden

        $bar.Style = [System.Windows.Forms.ProgressBarStyle]::Continuous
        $bar.Value = 30

        foreach ($p in $patterns) {
            $hit = Get-Item $p -ErrorAction SilentlyContinue | Sort-Object Name -Descending | Select-Object -First 1
            if ($hit) { $pythonw = $hit.FullName; break }
        }
    }

    # ── Attempt 2: direct download from python.org ───────────────────────────
    if (-not $pythonw) {
        Step "Downloading Python 3.12..." "Downloading installer (~25 MB) — please wait..." 18
        $bar.Style = [System.Windows.Forms.ProgressBarStyle]::Marquee
        $bar.MarqueeAnimationSpeed = 30
        $form.Refresh(); [System.Windows.Forms.Application]::DoEvents()

        $pyUrl       = "https://www.python.org/ftp/python/3.12.7/python-3.12.7-amd64.exe"
        $pyInstaller = "$env:TEMP\python-installer.exe"

        try {
            [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
            Invoke-WebRequest -Uri $pyUrl -OutFile $pyInstaller -UseBasicParsing -ErrorAction Stop

            Step "Installing Python 3.12..." "Running installer silently..." 25
            Start-Process $pyInstaller -ArgumentList "/quiet InstallAllUsers=0 PrependPath=1 Include_pip=1 Include_launcher=0" -Wait

            Remove-Item $pyInstaller -Force -ErrorAction SilentlyContinue
        } catch {
            # Download failed — nothing more we can do
        }

        $bar.Style = [System.Windows.Forms.ProgressBarStyle]::Continuous
        $bar.Value = 30

        foreach ($p in $patterns) {
            $hit = Get-Item $p -ErrorAction SilentlyContinue | Sort-Object Name -Descending | Select-Object -First 1
            if ($hit) { $pythonw = $hit.FullName; break }
        }
    }
}

if (-not $pythonw) {
    Step "Python could not be installed automatically." "Please install Python from python.org then run start.bat again." 0
    [System.Windows.Forms.MessageBox]::Show(
        "Python could not be installed automatically.`n`nPlease visit python.org, download and install Python 3.12, then run start.bat again.`n`nMake sure to tick 'Add Python to PATH' during installation.",
        "PM CV Formatter — Setup",
        "OK", "Error") | Out-Null
    Start-Process "https://www.python.org/downloads/"
    $form.Close()
    exit 1
}

$pyDir = Split-Path $pythonw -Parent
$pip   = Join-Path $pyDir "Scripts\pip.exe"
if (-not (Test-Path $pip)) { $pip = "pip" }

# ── Install pip dependencies ──────────────────────────────────────────────────
Step "Installing dependencies..." "Flask · python-docx · PyMuPDF · anthropic" 40

& $pip install -r "$AppDir\requirements.txt" --quiet --no-warn-script-location 2>&1 | Out-Null

# Ensure pymupdf explicitly (sometimes not in requirements on older installs)
& $pip install pymupdf --quiet --no-warn-script-location 2>&1 | Out-Null

Step "Dependencies ready." "" 65

# ── Desktop shortcut (first run only) ────────────────────────────────────────
Step "Setting up desktop shortcut..." "" 72

$flagFile = "$env:APPDATA\pm-cv-formatter.flag"
if (-not (Test-Path $flagFile)) {
    try {
        $desktop  = [Environment]::GetFolderPath("Desktop")
        $wsh      = New-Object -ComObject WScript.Shell
        $lnk      = $wsh.CreateShortcut("$desktop\PM CV Formatter.lnk")
        $lnk.TargetPath       = "$AppDir\start.bat"
        $lnk.WorkingDirectory = $AppDir
        $lnk.Description      = "Patrick Morgan CV Formatter"
        $lnk.IconLocation     = "shell32.dll,14"
        $lnk.Save()
        New-Item -Path $flagFile -ItemType File -Force | Out-Null
    } catch {}
}

# ── Start the Flask server ────────────────────────────────────────────────────
Step "Starting server..." "" 82

$env:PYTHONUTF8 = "1"
Start-Process -FilePath $pythonw `
              -ArgumentList "`"$AppDir\app.py`"" `
              -WorkingDirectory $AppDir `
              -WindowStyle Hidden

# ── Wait for server to respond (up to 12 seconds) ────────────────────────────
Step "Waiting for server to start..." "" 88

$ready  = $false
$tries  = 0
while ($tries -lt 15 -and -not $ready) {
    Start-Sleep -Milliseconds 800
    $pct = 88 + [int]($tries * 0.8)
    Step "Waiting for server to start..." "Attempt $($tries + 1) of 15..." $pct
    try {
        $r = Invoke-WebRequest -Uri "http://localhost:5000" -UseBasicParsing -TimeoutSec 1 -ErrorAction Stop
        if ($r.StatusCode -eq 200) { $ready = $true }
    } catch {}
    $tries++
}

if (-not $ready) {
    Fail "Server did not start in time. Try running start.bat again."
}

# ── Open browser ──────────────────────────────────────────────────────────────
Step "Opening PM CV Formatter..." "" 100
Start-Process "http://localhost:5000"
Start-Sleep -Milliseconds 700
$form.Close()

} catch {
    # Show any unexpected error as a popup so it's never silently swallowed
    $errMsg = $_.Exception.Message
    $errLine = $_.InvocationInfo.ScriptLineNumber
    try { $form.Close() } catch {}
    Add-Type -AssemblyName System.Windows.Forms -ErrorAction SilentlyContinue
    [System.Windows.Forms.MessageBox]::Show(
        "Setup failed with an unexpected error.`n`nError: $errMsg`nLine: $errLine`n`nPlease screenshot this and send to Will.",
        "PM CV Formatter — Error", "OK", "Error") | Out-Null
}
