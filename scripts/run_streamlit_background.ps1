param(
    [int]$Port = 8501,
    [string]$BindAddress = "127.0.0.1"
)

$ErrorActionPreference = "Stop"

function Test-TcpPortOpen {
    param(
        [string]$TargetHost,
        [int]$TargetPort
    )

    $client = New-Object System.Net.Sockets.TcpClient
    try {
        $iar = $client.BeginConnect($TargetHost, $TargetPort, $null, $null)
        $connected = $iar.AsyncWaitHandle.WaitOne(600)
        if (-not $connected) {
            return $false
        }
        $client.EndConnect($iar) | Out-Null
        return $true
    } catch {
        return $false
    } finally {
        $client.Close()
    }
}

$projectRoot = Split-Path -Parent $PSScriptRoot
$appPath = Join-Path $projectRoot "app.py"
$logDir = Join-Path $projectRoot "logs"
$outLog = Join-Path $logDir "streamlit.out.log"
$errLog = Join-Path $logDir "streamlit.err.log"

$pythonExe = $null
$pythonArgsPrefix = @()
$venvPython = Join-Path $projectRoot ".venv\Scripts\python.exe"
if (Test-Path $venvPython) {
    $pythonExe = $venvPython
} else {
    $pythonCmd = Get-Command "python" -ErrorAction SilentlyContinue
    if ($pythonCmd) {
        $pythonExe = $pythonCmd.Source
    } else {
        $pyLauncher = Get-Command "py" -ErrorAction SilentlyContinue
        if ($pyLauncher) {
            $pythonExe = $pyLauncher.Source
            $pythonArgsPrefix = @("-3")
        }
    }
}

if (-not (Test-Path $appPath)) {
    throw "app.py not found: $appPath"
}
if (-not $pythonExe) {
    throw "Python executable not found. Install Python 3 and ensure python or py is available."
}

New-Item -ItemType Directory -Path $logDir -Force | Out-Null

if (Test-TcpPortOpen -TargetHost $BindAddress -TargetPort $Port) {
    Write-Output "Streamlit is already running at http://$BindAddress`:$Port"
    exit 0
}

$args = @()
$args += $pythonArgsPrefix
$args += @(
    "-m", "streamlit", "run", $appPath,
    "--server.port", "$Port",
    "--server.address", $BindAddress,
    "--server.headless", "true",
    "--browser.gatherUsageStats", "false"
)

Start-Process `
    -FilePath $pythonExe `
    -ArgumentList $args `
    -WorkingDirectory $projectRoot `
    -WindowStyle Hidden `
    -RedirectStandardOutput $outLog `
    -RedirectStandardError $errLog | Out-Null

Start-Sleep -Seconds 2

if (Test-TcpPortOpen -TargetHost $BindAddress -TargetPort $Port) {
    Write-Output "Started Streamlit at http://$BindAddress`:$Port"
    exit 0
}

Write-Output "Streamlit process started, but port check has not passed yet."
exit 0
