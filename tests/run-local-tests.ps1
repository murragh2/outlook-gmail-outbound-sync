# Scripted End-to-End Local Test Harness for Azure Functions
$ErrorActionPreference = "Stop"

$projectDir = Join-Path $PSScriptRoot "..\src\OutlookGmailSync"
$localSettingsPath = Join-Path $projectDir "local.settings.json"
$stdoutPath = Join-Path $PSScriptRoot "func-stdout.log"
$stderrPath = Join-Path $PSScriptRoot "func-stderr.log"

Write-Host "================================================================================" -ForegroundColor Cyan
Write-Host "[LOCAL TEST HARNESS] Starting Azure Functions Local E2E Test Suite" -ForegroundColor Cyan
Write-Host "================================================================================" -ForegroundColor Cyan

if (-not (Test-Path $localSettingsPath)) {
    Write-Host "[WARNING] local.settings.json not found in $projectDir." -ForegroundColor Yellow
}

# Clear previous log files
if (Test-Path $stdoutPath) { Remove-Item -Path $stdoutPath -Force }
if (Test-Path $stderrPath) { Remove-Item -Path $stderrPath -Force }

Write-Host "[1/4] Launching local Azure Functions host (func start)..." -ForegroundColor Green
$funcProcess = Start-Process -FilePath "func" -ArgumentList "start", "--port", "7071" -WorkingDirectory $projectDir -RedirectStandardOutput $stdoutPath -RedirectStandardError $stderrPath -PassThru

try {
    Write-Host "[2/4] Waiting for local host to become healthy (http://localhost:7071/admin/host/status)..." -ForegroundColor Green
    $healthy = $false
    $timeout = 30
    $elapsed = 0

    while (-not $healthy -and $elapsed -lt $timeout) {
        Start-Sleep -Seconds 2
        $elapsed += 2
        try {
            $status = Invoke-RestMethod -Uri "http://localhost:7071/admin/host/status" -Method Get -TimeoutSec 3 -ErrorAction SilentlyContinue
            if ($status -and ($status.state -eq "Running" -or $status.state -eq "1")) {
                $healthy = $true
                Write-Host " -> Host is RUNNING! (Took ${elapsed}s)" -ForegroundColor Green
            }
        } catch {
            Write-Host " -> Waiting for host... (${elapsed}s)" -ForegroundColor Gray
        }
    }

    if (-not $healthy) {
        throw "Local Azure Functions host failed to reach Running state within ${timeout}s timeout."
    }

    Write-Host "[3/4] Triggering SyncTimerFunction via admin endpoint..." -ForegroundColor Green
    $triggerUri = "http://localhost:7071/admin/functions/SyncTimerFunction"
    $response = Invoke-WebRequest -Uri $triggerUri -Method Post -Body "{}" -ContentType "application/json" -UseBasicParsing

    if ($response.StatusCode -ne 202 -and $response.StatusCode -ne 200) {
        throw "Unexpected HTTP response status from admin endpoint: $($response.StatusCode)"
    }

    Write-Host "[4/4] Monitoring worker execution logs..." -ForegroundColor Green
    Start-Sleep -Seconds 5

    $logContent = ""
    if (Test-Path $stdoutPath) { $logContent += Get-Content -Path $stdoutPath -Raw }
    if (Test-Path $stderrPath) { $logContent += Get-Content -Path $stderrPath -Raw }

    Write-Host "--------------------------------------------------------------------------------" -ForegroundColor Gray
    Write-Host "Worker Output Logs:" -ForegroundColor Gray
    Write-Host $logContent -ForegroundColor Gray
    Write-Host "--------------------------------------------------------------------------------" -ForegroundColor Gray

    if ($logContent -match "InvalidOperationException" -or $logContent -match "\[SYNC FAILED\]" -or $logContent -match "System\.Exception" -or $logContent -match "expired or is missing") {
        Write-Host "================================================================================" -ForegroundColor Red
        Write-Host "[TEST FAILED ❌] Local worker failed during sync execution (Missing/Expired Access Token)." -ForegroundColor Red
        Write-Host "================================================================================" -ForegroundColor Red
        exit 1
    } elseif ($logContent -match "\[SYNC SUMMARY\]" -or $logContent -match "\[INITIAL SYNC COMPLETE\]") {
        Write-Host "================================================================================" -ForegroundColor Green
        Write-Host "[TEST PASSED ✅] Local Function App executed sync cycle successfully!" -ForegroundColor Green
        Write-Host "================================================================================" -ForegroundColor Green
        exit 0
    } else {
        Write-Host "================================================================================" -ForegroundColor Yellow
        Write-Host "[TEST INCOMPLETE ⚠️] Worker completed invocation without producing explicit summary." -ForegroundColor Yellow
        Write-Host "================================================================================" -ForegroundColor Yellow
    }
}
finally {
    Write-Host "[TEARDOWN] Stopping background Azure Functions host process..." -ForegroundColor Yellow
    if ($funcProcess -and -not $funcProcess.HasExited) {
        Stop-Process -Id $funcProcess.Id -Force -ErrorAction SilentlyContinue
    }
    Write-Host "[TEARDOWN] Complete." -ForegroundColor Yellow
}
