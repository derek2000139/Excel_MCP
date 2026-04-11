# ExcelForge One-Click Installer

$Host.UI.RawUI.BackgroundColor = "Black"
$Host.UI.RawUI.ForegroundColor = "White"

try {
    $ErrorActionPreference = "Stop"

    function Write-Pass {
        param([string]$Message)
        Write-Host "  " -NoNewline
        Write-Host "[OK]" -ForegroundColor Green -NoNewline
        Write-Host " $Message"
    }

    function Write-Fail {
        param([string]$Message)
        Write-Host "  " -NoNewline
        Write-Host "[FAIL]" -ForegroundColor Red -NoNewline
        Write-Host " $Message"
    }

    function Write-Warn {
        param([string]$Message)
        Write-Host "  " -NoNewline
        Write-Host "[WARN]" -ForegroundColor Yellow -NoNewline
        Write-Host " $Message"
    }

    function Write-Info {
        param([string]$Message)
        Write-Host "[INFO] $Message" -ForegroundColor Cyan
    }

    $scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
    Set-Location $scriptDir

    Write-Host ""
    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host "  ExcelForge Installer" -ForegroundColor Cyan
    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host ""

    # Check OS
    Write-Info "Checking system requirements..."
    if (-not $IsWindows -and $env:OS -ne "Windows_NT") {
        Write-Fail "Windows OS required"
        Read-Host "Press Enter to exit"
        exit 1
    }
    Write-Pass "Windows OS"

    # Check Excel
    try {
        $excel = Get-ItemProperty "HKLM:\Software\Microsoft\Windows\CurrentVersion\App Paths\excel.exe" -ErrorAction SilentlyContinue
        if (-not $excel) {
            $excel = Get-ItemProperty "HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\App Paths\excel.exe" -ErrorAction SilentlyContinue
        }
        if ($excel) {
            Write-Pass "Microsoft Excel"
        } else {
            Write-Warn "Microsoft Excel (optional)"
            $continue = Read-Host "Continue without Excel? (y/N)"
            if ($continue -ne 'y' -and $continue -ne 'Y') {
                exit 1
            }
        }
    } catch {
        Write-Warn "Microsoft Excel (check failed)"
    }

    # Check uv
    try {
        $uvVersion = uv --version 2>&1
        Write-Pass "uv package manager"
        
        # Check Python version via uv
        try {
            $pythonVersion = uv run python --version 2>&1
            if ($pythonVersion -match "Python (\d+)\.(\d+)") {
                $major = [int]$matches[1]
                $minor = [int]$matches[2]
                if ($major -lt 3 -or ($major -eq 3 -and $minor -lt 11)) {
                    Write-Fail "Python $major.$minor (requires 3.11+)"
                    Read-Host "Press Enter to exit"
                    exit 1
                }
                Write-Pass "Python $major.$minor"
            } else {
                throw "Cannot parse Python version"
            }
        } catch {
            Write-Fail "Python environment"
            Write-Info "Try: uv python install 3.11"
            Read-Host "Press Enter to exit"
            exit 1
        }
    } catch {
        Write-Info "uv not found, installing..."
        try {
            pip install uv 2>$null
            Write-Pass "uv package manager (installed)"
            
            try {
                uv python install 3.11 2>$null
                Write-Pass "Python 3.11 (installed)"
            } catch {
                Write-Warn "Python 3.11 (manual install needed)"
            }
        } catch {
            Write-Fail "uv installation"
            Read-Host "Press Enter to exit"
            exit 1
        }
    }

    # Check project files
    if (Test-Path "pyproject.toml") {
        Write-Pass "Project files"
    } else {
        Write-Fail "Project files (pyproject.toml not found)"
        Read-Host "Press Enter to exit"
        exit 1
    }

    # Install dependencies
    Write-Host ""
    Write-Info "Installing dependencies..."
    uv sync
    if ($?) {
        Write-Pass "Dependencies installed"
    } else {
        Write-Fail "Dependencies installation"
        Read-Host "Press Enter to exit"
        exit 1
    }

    # Verify installation
    Write-Host ""
    Write-Info "Verifying installation..."
    try {
        $testResult = uv run python -c "import excelforge; print('OK')" 2>$null
        if ($testResult -match "OK") {
            Write-Pass "Python environment"
        } else {
            throw "Import failed"
        }
    } catch {
        Write-Fail "Python environment verification"
        Read-Host "Press Enter to exit"
        exit 1
    }
    
    Write-Host ""
    Write-Host "========================================" -ForegroundColor Green
    Write-Host "  Installation Complete!" -ForegroundColor Green
    Write-Host "========================================" -ForegroundColor Green
    Write-Host ""
    Write-Host "Quick Start:"
    Write-Host "  uv run python -m excelforge.gateway.host --config excel-mcp.yaml --profile basic_edit"
    Write-Host ""

    Read-Host "Press Enter to exit"

} catch {
    Write-Host ""
    Write-Host "========================================" -ForegroundColor Red
    Write-Host "  Installation Failed" -ForegroundColor Red
    Write-Host "========================================" -ForegroundColor Red
    Write-Host ""
    Write-Host "Error: $_" -ForegroundColor Red
    Write-Host ""
    Read-Host "Press Enter to exit"
    exit 1
}
