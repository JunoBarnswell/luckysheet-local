[CmdletBinding()]
param(
    [string]$Uri = 'http://127.0.0.1:8082/health',
    [int]$TimeoutSeconds = 60
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest
$deadline = [DateTime]::UtcNow.AddSeconds($TimeoutSeconds)
do {
    try {
        $response = Invoke-WebRequest -UseBasicParsing -Uri $Uri -TimeoutSec 5
        if ($response.StatusCode -ge 200 -and $response.StatusCode -lt 300) {
            Write-Host "Health check passed: $Uri ($($response.StatusCode))"
            exit 0
        }
    }
    catch {
        # The service may still be binding its port; retry until the bounded deadline.
    }
    Start-Sleep -Seconds 1
} while ([DateTime]::UtcNow -lt $deadline)

Write-Error "HEALTH_CHECK_FAILED: $Uri did not return a successful response within $TimeoutSeconds seconds."
exit 1
