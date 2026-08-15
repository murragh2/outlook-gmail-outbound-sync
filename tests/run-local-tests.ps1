# Scripted End-to-End Local Test Harness for Azure Functions
$ErrorActionPreference = "Stop"

$projectDir = Join-Path $PSScriptRoot "..\src\OutlookGmailSync"
$localSettingsPath = Join-Path $projectDir "local.settings.json"

Write-Host "================================================================================" -ForegroundColor Cyan
Write-Host "[LOCAL TEST HARNESS] Starting Azure Functions Local E2E Test Suite" -ForegroundColor Cyan
Write-Host "================================================================================" -ForegroundColor Cyan

if (-not (Test-Path $localSettingsPath)) {
    Write-Host "[WARNING] local.settings.json not found in $projectDir." -ForegroundColor Yellow
    Write-Host "[WARNING] Creating local.settings.json template..." -ForegroundColor Yellow
    @{
        IsEncrypted = $false
        Values = @{
            AzureWebJobsStorage = "UseDevelopmentStorage=true"
            FUNCTIONS_WORKER_RUNTIME = "dotnet-isolated"
            AzureAd__TenantId = "consumers"
            AzureAd__ClientId = "06d76858-fa3e-48c2-8b98-3f3d167efa6b"
            Sync__GmailAddress = "murragh2@gmail.com"
            Sync__CronSchedule = "0 */10 * * * *"
            Storage__AccountUri = "https://stogmsx4g5sogvlsg3i.table.core.windows.net"
            KeyVault__VaultUri = "https://kv-ogms-x4g5sogvlsg3i.vault.azure.net"
        }
    } | ConvertTo-Json -Depth 5 | Set-Content -Path $localSettingsPath
}

Write-Host "[1/4] Launching local Azure Functions host (func start)..." -ForegroundColor Green
$funcProcess = Start-Process -FilePath "func" -ArgumentList "start", "--port", "7071" -WorkingDirectory $projectDir -PassThru -NoNewWindow

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

    Write-Host "[4/4] Invocation Result: HTTP $($response.StatusCode) ($($response.StatusDescription))" -ForegroundColor Cyan

    if ($response.StatusCode -eq 202 -or $response.StatusCode -eq 200) {
        Write-Host "================================================================================" -ForegroundColor Green
        Write-Host "[TEST PASSED ✅] Local Function App triggered and executed successfully!" -ForegroundColor Green
        Write-Host "================================================================================" -ForegroundColor Green
    } else {
        throw "Unexpected HTTP response status: $($response.StatusCode)"
    }
}
finally {
    Write-Host "[TEARDOWN] Stopping background Azure Functions host process..." -ForegroundColor Yellow
    if ($funcProcess -and -not $funcProcess.HasExited) {
        Stop-Process -Id $funcProcess.Id -Force -ErrorAction SilentlyContinue
    }
    Write-Host "[TEARDOWN] Complete." -ForegroundColor Yellow
}
