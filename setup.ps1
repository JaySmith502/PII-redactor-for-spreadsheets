# =============================================================================
# setup.ps1 - Automated first-time setup for the XLSX PII Scrubber Web App
# Run once from the project root:  .\setup.ps1
# =============================================================================

$ErrorActionPreference = "Stop"

function Write-Step { param($msg) Write-Host "`n>> $msg" -ForegroundColor Cyan }
function Write-OK   { param($msg) Write-Host "   [OK] $msg" -ForegroundColor Green }
function Write-Warn { param($msg) Write-Host "   [!]  $msg" -ForegroundColor Yellow }
function Write-Fail { param($msg) Write-Host "   [X]  $msg" -ForegroundColor Red }

$script:PythonCommand = $null
$script:PythonArgs = @()

function Test-PythonCandidate {
    param(
        [string]$Command,
        [string[]]$Arguments = @()
    )

    try {
        & $Command @Arguments -c "import sys; sys.exit(0 if sys.version_info >= (3, 10) else 1)" | Out-Null
        return ($LASTEXITCODE -eq 0)
    } catch {
        return $false
    }
}

function Set-PythonCommand {
    $candidates = @(
        @{ Command = "python"; Args = @() },
        @{ Command = "py"; Args = @("-3") },
        @{ Command = "python3"; Args = @() }
    )

    foreach ($candidate in $candidates) {
        if (Get-Command $candidate.Command -ErrorAction SilentlyContinue) {
            if (Test-PythonCandidate -Command $candidate.Command -Arguments $candidate.Args) {
                $script:PythonCommand = $candidate.Command
                $script:PythonArgs = $candidate.Args
                return $true
            }
        }
    }

    return $false
}

function Invoke-Python {
    param([Parameter(ValueFromRemainingArguments = $true)][string[]]$Arguments)
    & $script:PythonCommand @script:PythonArgs @Arguments
}

function Get-PythonVersionText {
    return (& $script:PythonCommand @script:PythonArgs --version 2>&1)
}

function Install-Python {
    if (-not (Get-Command winget -ErrorAction SilentlyContinue)) {
        Write-Fail "Python 3.10+ was not found, and winget is not available for automatic installation."
        Write-Host "   Install Python from https://www.python.org/downloads/windows/" -ForegroundColor White
        Write-Host "   During installation, check 'Add python.exe to PATH', then run .\setup.ps1 again." -ForegroundColor White
        exit 1
    }

    Write-Step "Installing Python 3.12 with winget..."
    winget install --id Python.Python.3.12 --exact --scope user --accept-package-agreements --accept-source-agreements
    if ($LASTEXITCODE -ne 0) {
        Write-Warn "User-scope install failed. Trying the default winget install..."
        winget install --id Python.Python.3.12 --exact --accept-package-agreements --accept-source-agreements
    }

    if ($LASTEXITCODE -ne 0) {
        Write-Fail "Python installation failed. Install Python 3.10+ manually, then run .\setup.ps1 again."
        exit 1
    }
}

function Test-SecretKeyValue {
    param([string]$Key)

    if (-not $Key) {
        return $false
    }

    $script = @"
import base64
import sys

value = sys.argv[1].strip()
padding = "=" * (-len(value) % 4)
try:
    decoded = base64.urlsafe_b64decode(value + padding)
except Exception:
    sys.exit(1)
sys.exit(0 if len(decoded) == 64 else 1)
"@

    $scriptPath = Join-Path ([System.IO.Path]::GetTempPath()) "pii_scrubber_key_check.py"
    Set-Content $scriptPath $script -NoNewline
    try {
        Invoke-Python $scriptPath $Key | Out-Null
        return ($LASTEXITCODE -eq 0)
    } finally {
        Remove-Item $scriptPath -ErrorAction SilentlyContinue
    }
}

function Invoke-KeyCopyAction {
    param([string]$Key)

    Write-Host ""
    Write-Host "   Copy Key action" -ForegroundColor Cyan
    Write-Host "   Press C to copy the generated key to the local Windows clipboard." -ForegroundColor White
    Write-Host "   Press S to show the key once in this local terminal." -ForegroundColor White
    Write-Host "   Press Enter to skip." -ForegroundColor White

    $choice = Read-Host "   Choice"
    if ($choice -match "^[Cc]$") {
        try {
            Set-Clipboard -Value $Key
            Write-OK "SECRET_KEY copied to the local Windows clipboard"
        } catch {
            Write-Warn "Clipboard copy failed. Run PowerShell locally and copy the key from .env if needed."
        }
    } elseif ($choice -match "^[Ss]$") {
        Write-Host ""
        Write-Host "   Generated SECRET_KEY:" -ForegroundColor Yellow
        Write-Host "   $Key" -ForegroundColor Yellow
        Write-Warn "Do not paste this key into chat, email, tickets, or logs."
    } else {
        Write-Warn "Skipped key copy/display."
    }
}

Write-Host ""
Write-Host "=============================================" -ForegroundColor Magenta
Write-Host "  XLSX PII Scrubber - First-time Setup" -ForegroundColor Magenta
Write-Host "=============================================" -ForegroundColor Magenta

# ---------------------------------------------------------------------------
# 1. Check Python version
# ---------------------------------------------------------------------------
Write-Step "Checking Python version..."
if (-not (Set-PythonCommand)) {
    Write-Warn "Python 3.10+ was not found. Setup will try to install it now."
    Install-Python
    if (-not (Set-PythonCommand)) {
        Write-Fail "Python was installed, but this terminal cannot find it yet."
        Write-Host "   Close this terminal, open a new one in this folder, and run .\setup.ps1 again." -ForegroundColor White
        exit 1
    }
}
$pyVer = Get-PythonVersionText
if ($pyVer -notmatch "Python (\d+)\.(\d+)") {
    Write-Fail "Python was found, but its version output could not be read."
    exit 1
}
Write-OK "$pyVer found"

# ---------------------------------------------------------------------------
# 2. Install dependencies from requirements.txt
# ---------------------------------------------------------------------------
Write-Step "Installing dependencies from requirements.txt..."
Invoke-Python -m pip install -r requirements.txt --quiet
if ($LASTEXITCODE -ne 0) {
    Write-Fail "pip install failed. Check the internet connection and requirements.txt."
    exit 1
}
Write-OK "All packages installed"

# ---------------------------------------------------------------------------
# 3. Create .env if it does not exist
# ---------------------------------------------------------------------------
Write-Step "Checking .env..."
if (Test-Path ".env") {
    Write-Warn ".env already exists and will not be overwritten."
} else {
    if (-not (Test-Path ".env.example")) {
        Write-Fail ".env.example was not found. Make sure you run this script from the project root folder."
        exit 1
    }
    Copy-Item ".env.example" ".env"
    Write-OK ".env created from .env.example"
}

# ---------------------------------------------------------------------------
# 4. Check whether SECRET_KEY is set and valid
# ---------------------------------------------------------------------------
Write-Step "Checking SECRET_KEY in .env..."
$envContent = Get-Content ".env" -Raw
$keyLine = ($envContent -split "`n") | Where-Object { $_ -match "^\s*SECRET_KEY\s*=" }

$keyValue = ""
if ($keyLine) {
    $keyValue = ($keyLine -split "=", 2)[1].Trim().Trim('"').Trim("'")
}

$placeholder = "paste_the_complete_generated_key_here"
$keyIsValid = Test-SecretKeyValue -Key $keyValue
$needNewKey = (-not $keyLine) -or ($keyValue -eq "") -or ($keyValue -eq $placeholder) -or (-not $keyIsValid)

if ($needNewKey) {
    if ($keyLine -and $keyValue -and $keyValue -ne $placeholder -and (-not $keyIsValid)) {
        Write-Warn "The current SECRET_KEY is not valid for this app. Creating a new 64-byte key..."
    } else {
        Write-Warn "SECRET_KEY is not set. Creating a new key..."
    }

    $newKey = Invoke-Python -c "import base64,secrets; print(base64.urlsafe_b64encode(secrets.token_bytes(64)).decode('ascii').rstrip('='))"
    if ($LASTEXITCODE -ne 0 -or -not $newKey) {
        Write-Fail "Failed to create SECRET_KEY with Python."
        exit 1
    }

    # Replace the SECRET_KEY line in .env, or add it if it does not exist.
    if ($keyLine) {
        $newContent = $envContent -replace '(?m)^\s*SECRET_KEY\s*=.*$', "SECRET_KEY=`"$newKey`""
    } else {
        $newContent = $envContent.TrimEnd() + "`nSECRET_KEY=`"$newKey`"`n"
    }
    Set-Content ".env" $newContent -NoNewline
    Write-OK "New SECRET_KEY generated and saved to .env"
    Write-Host ""
    Write-Host "   IMPORTANT: Save this key in a secure place." -ForegroundColor Yellow
    Write-Host "   Do not change this key while encrypted files still need to be restored." -ForegroundColor Yellow
    Invoke-KeyCopyAction -Key $newKey
} else {
    Write-OK "SECRET_KEY is already available"
}

# ---------------------------------------------------------------------------
# 5. Validate configuration with _check_env_key.py if available
# ---------------------------------------------------------------------------
Write-Step "Validating configuration..."
if (Test-Path "_check_env_key.py") {
    Invoke-Python _check_env_key.py
    if ($LASTEXITCODE -ne 0) {
        Write-Fail "_check_env_key.py reported an invalid configuration. Check .env and try again."
        exit 1
    }
    Write-OK "Configuration validation passed"
} else {
    Write-Warn "_check_env_key.py was not found, skipping validation."
}

# ---------------------------------------------------------------------------
# 6. Done - show how to run the app
# ---------------------------------------------------------------------------
Write-Host ""
Write-Host "=============================================" -ForegroundColor Green
Write-Host "  Setup complete. The app is ready to run." -ForegroundColor Green
Write-Host "=============================================" -ForegroundColor Green
Write-Host ""
Write-Host "  To start the development server:" -ForegroundColor White
$runPrefix = ($script:PythonCommand + " " + ($script:PythonArgs -join " ")).Trim()
Write-Host "    $runPrefix -m uvicorn app:app --host 127.0.0.1 --port 8000 --reload" -ForegroundColor DarkCyan
Write-Host ""
Write-Host "  To start the production server:" -ForegroundColor White
Write-Host "    $runPrefix -m uvicorn app:app --host 127.0.0.1 --port 8000 --workers 1" -ForegroundColor DarkCyan
Write-Host ""
Write-Host "  Open this in a browser: http://127.0.0.1:8000" -ForegroundColor White
Write-Host ""
