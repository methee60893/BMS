[CmdletBinding()]
param(
    [switch]$RestoreOnly,
    [switch]$DryRun,
    [switch]$RunSql,
    [switch]$StartMockSap,
    [switch]$SkipSwitchConfig,
    [switch]$PreflightOnly,
    [string]$SqlServer = "YOUR_SERVER",
    [string]$TestDatabase = "BMS_TEST",
    [string]$DbUser = "sa",
    [string]$DbPassword = "YOUR_PASSWORD",
    [string]$MockSapBaseUrl = "http://127.0.0.1:18080",
    [string]$MockSapUsername = "mock-user",
    [string]$MockSapPassword = "mock-password"
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$root = Split-Path -Parent $MyInvocation.MyCommand.Path
$stepScript = Join-Path $root "run_test_step_by_step.ps1"
$testSqlDir = Join-Path $root "database\test"
$createDbSql = Join-Path $root "database\00_create_database_test.sql"
$stateDir = Join-Path $root ".codex-test-state"
$logPath = Join-Path $stateDir "run_full_isolated_test.log"
$mockServerScript = Join-Path $root "test_support\mock_sap_server.py"
$webConfigPath = Join-Path $root "Web.config"
$verifyPythonScript = Join-Path $root "database\compare_excel_to_db.py"
$sqlFiles = @(
    (Join-Path $testSqlDir "01_create_tables_test.sql"),
    (Join-Path $testSqlDir "02_create_views_test.sql"),
    (Join-Path $testSqlDir "03_create_stored_procedures_test.sql"),
    (Join-Path $testSqlDir "04_seed_master_data_template_test.sql"),
    (Join-Path $testSqlDir "05_post_deploy_verify_test.sql")
)

function Invoke-StepScript {
    param(
        [string]$Action,
        [hashtable]$ExtraArgs = @{}
    )

    $argList = @(
        "-NoProfile",
        "-ExecutionPolicy", "Bypass",
        "-File", $stepScript,
        $Action,
        "-SqlServer", $SqlServer,
        "-TestDatabase", $TestDatabase,
        "-DbUser", $DbUser,
        "-DbPassword", $DbPassword,
        "-MockSapBaseUrl", $MockSapBaseUrl,
        "-MockSapUsername", $MockSapUsername,
        "-MockSapPassword", $MockSapPassword
    )

    foreach ($key in $ExtraArgs.Keys) {
        $argList += "-$key"
        $argList += [string]$ExtraArgs[$key]
    }

    & powershell @argList
}

function Ensure-StateDir {
    if (-not (Test-Path -LiteralPath $stateDir)) {
        New-Item -ItemType Directory -Path $stateDir | Out-Null
    }
    if (-not (Test-Path -LiteralPath $logPath)) {
        New-Item -ItemType File -Path $logPath -Force | Out-Null
    }
}

function Write-Log {
    param(
        [string]$Message,
        [string]$Level = "INFO"
    )

    Ensure-StateDir
    $line = "[{0}] [{1}] {2}" -f (Get-Date).ToString("yyyy-MM-dd HH:mm:ss"), $Level, $Message
    [System.IO.File]::AppendAllText($logPath, $line + [Environment]::NewLine, [System.Text.UTF8Encoding]::new($false))
}

function Ensure-Prerequisites {
    Write-Log "Checking prerequisites"
    if (-not (Test-Path -LiteralPath $stepScript)) {
        Write-Log "Missing script: $stepScript" "ERROR"
        throw "Missing script: $stepScript"
    }
    if (-not (Get-Command python -ErrorAction SilentlyContinue)) {
        Write-Log "python command not found" "ERROR"
        throw "python command not found."
    }
}

function Test-PortInUse {
    param([int]$Port)
    $listener = [System.Net.Sockets.TcpListener]::new([System.Net.IPAddress]::Loopback, $Port)
    try {
        $listener.Start()
        return $false
    }
    catch {
        return $true
    }
    finally {
        try { $listener.Stop() } catch {}
    }
}

function Test-SqlConnection {
    if (-not (Get-Command sqlcmd -ErrorAction SilentlyContinue)) {
        return [pscustomobject]@{ Success = $false; Message = "sqlcmd not found in PATH" }
    }
    if ($SqlServer -eq "YOUR_SERVER") {
        return [pscustomobject]@{ Success = $false; Message = "SqlServer is still placeholder YOUR_SERVER" }
    }

    $args = @("-S", $SqlServer, "-E", "-Q", "SELECT 1 AS Ping")
    try {
        $output = & sqlcmd @args 2>&1 | Out-String
        if ($LASTEXITCODE -eq 0) {
            return [pscustomobject]@{ Success = $true; Message = "SQL connection OK" }
        }
        $msg = $output.Trim()
        if ([string]::IsNullOrWhiteSpace($msg)) {
            $msg = "Unable to connect to SQL Server with current sqlcmd settings"
        }
        return [pscustomobject]@{ Success = $false; Message = $msg }
    }
    catch {
        return [pscustomobject]@{ Success = $false; Message = $_.Exception.Message }
    }
}

function Invoke-Preflight {
    Write-Log "Running preflight checks"
    Write-Host "Preflight checks"
    Write-Host "---------------"

    $checks = New-Object System.Collections.Generic.List[object]

    $checks.Add([pscustomobject]@{
        Name = "PowerShell step script"
        Status = if (Test-Path -LiteralPath $stepScript) { "PASS" } else { "FAIL" }
        Detail = $stepScript
    })
    $checks.Add([pscustomobject]@{
        Name = "Web.config"
        Status = if (Test-Path -LiteralPath $webConfigPath) { "PASS" } else { "FAIL" }
        Detail = $webConfigPath
    })
    $checks.Add([pscustomobject]@{
        Name = "Mock SAP script"
        Status = if (Test-Path -LiteralPath $mockServerScript) { "PASS" } else { "FAIL" }
        Detail = $mockServerScript
    })
    $checks.Add([pscustomobject]@{
        Name = "Python"
        Status = if (Get-Command python -ErrorAction SilentlyContinue) { "PASS" } else { "FAIL" }
        Detail = "python command"
    })
    $checks.Add([pscustomobject]@{
        Name = "sqlcmd"
        Status = if (Get-Command sqlcmd -ErrorAction SilentlyContinue) { "PASS" } else { "WARN" }
        Detail = "sqlcmd command"
    })
    $checks.Add([pscustomobject]@{
        Name = "Compare script"
        Status = if (Test-Path -LiteralPath $verifyPythonScript) { "PASS" } else { "FAIL" }
        Detail = $verifyPythonScript
    })

    $mockUri = [System.Uri]$MockSapBaseUrl
    $portBusy = Test-PortInUse -Port $mockUri.Port
    $checks.Add([pscustomobject]@{
        Name = "Mock SAP port"
        Status = if ($portBusy) { "WARN" } else { "PASS" }
        Detail = if ($portBusy) { "Port $($mockUri.Port) already in use" } else { "Port $($mockUri.Port) available" }
    })

    $sqlResult = Test-SqlConnection
    $checks.Add([pscustomobject]@{
        Name = "SQL connectivity"
        Status = if ($sqlResult.Success) { "PASS" } else { "WARN" }
        Detail = $sqlResult.Message
    })

    foreach ($check in $checks) {
        Write-Host ("[{0}] {1} - {2}" -f $check.Status, $check.Name, $check.Detail)
        Write-Log ("Preflight {0}: {1} - {2}" -f $check.Status, $check.Name, $check.Detail)
    }

    $failed = @($checks | Where-Object { $_.Status -eq "FAIL" })
    if ($failed.Count -gt 0) {
        Write-Log "Preflight found blocking failures" "ERROR"
        throw "Preflight failed. Fix FAIL items before continuing."
    }
}

function Ensure-TestSql {
    Write-Log "Generating test SQL files through step script"
    Invoke-StepScript -Action "generate-test-sql"
    foreach ($file in $sqlFiles) {
        if (-not (Test-Path -LiteralPath $file)) {
            Write-Log "Missing generated SQL file: $file" "ERROR"
            throw "Missing generated SQL file: $file"
        }
    }
    Write-Log "Verified generated test SQL files"
}

function Invoke-SqlFile {
    param(
        [string]$Database,
        [string]$InputFile
    )

    if (-not (Get-Command sqlcmd -ErrorAction SilentlyContinue)) {
        Write-Log "sqlcmd not found in PATH" "ERROR"
        throw "sqlcmd not found in PATH."
    }

    if ($SqlServer -eq "YOUR_SERVER") {
        Write-Log "SqlServer placeholder still in use" "ERROR"
        throw "Please specify -SqlServer before using -RunSql."
    }

    $args = @("-S", $SqlServer, "-E")
    if ($Database) {
        $args += @("-d", $Database)
    }
    $args += @("-i", $InputFile)

    if ($DryRun) {
        Write-Host ("DRY RUN: sqlcmd " + ($args -join " "))
        Write-Log ("DRY RUN sqlcmd " + ($args -join " "))
        return
    }

    Write-Log ("Executing sqlcmd " + ($args -join " "))
    & sqlcmd @args
    if ($LASTEXITCODE -ne 0) {
        Write-Log "sqlcmd failed for file: $InputFile" "ERROR"
        throw "sqlcmd failed for file: $InputFile"
    }
    Write-Log "sqlcmd completed for file: $InputFile"
}

function Run-SqlSequence {
    Write-Log "Running SQL sequence"
    Invoke-SqlFile -Database "" -InputFile $createDbSql
    foreach ($file in $sqlFiles) {
        Invoke-SqlFile -Database $TestDatabase -InputFile $file
    }
    Write-Log "Completed SQL sequence"
}

function Show-ManualSummary {
    Write-Log "Showing manual summary"
    Write-Host ""
    Write-Host "Log file: $logPath"
    Write-Host "Next steps:"
    Write-Host "1. Open the web app using your normal Visual Studio / IIS Express flow."
    Write-Host "2. Test Draft OTB, Approve OTB, Draft PO, Sync PO, Matching, Export."
    Write-Host "3. When finished, restore normal environment with:"
    Write-Host "   powershell -ExecutionPolicy Bypass -File `"$PSCommandPath`" -RestoreOnly"
}

Ensure-Prerequisites
Invoke-Preflight

if ($RestoreOnly) {
    Write-Log "RestoreOnly mode requested"
    if ($DryRun) {
        Write-Host "DRY RUN: restore normal environment"
        Write-Log "DRY RUN restore normal environment"
    }
    else {
        Invoke-StepScript -Action "restore"
        Write-Log "Restore command executed"
    }
    exit 0
}

if ($PreflightOnly) {
    Write-Host ""
    Write-Host "Preflight completed successfully."
    exit 0
}

Write-Host "Step 1: current status"
Write-Log "Step 1 current status"
Invoke-StepScript -Action "status"

Write-Host ""
Write-Host "Step 2: generate test SQL files"
Write-Log "Step 2 generate test SQL files"
Ensure-TestSql

Write-Host ""
Write-Host "Step 3: SQL commands"
Write-Log "Step 3 print SQL commands"
Invoke-StepScript -Action "print-sql"

if ($RunSql) {
    Write-Host ""
    Write-Host "Step 4: execute SQL scripts"
    if ($DryRun) {
        Write-Host "RunSql + DryRun enabled. SQL commands will be printed only."
        Write-Log "RunSql + DryRun enabled. SQL commands printed only."
    }
    else {
        Write-Log "Step 4 execute SQL scripts"
    }
    Run-SqlSequence
}

if (-not $SkipSwitchConfig) {
    Write-Host ""
    Write-Host "Step 5: switch Web.config to test environment"
    if ($DryRun) {
        Write-Host "DRY RUN: switch Web.config to test environment"
        Write-Log "DRY RUN switch Web.config to test environment"
    }
    else {
        Invoke-StepScript -Action "switch-to-test"
        Write-Log "Executed switch-to-test"
    }
}

if ($StartMockSap) {
    Write-Host ""
    Write-Host "Step 6: start mock SAP"
    if ($DryRun) {
        Write-Host "DRY RUN: start mock SAP at $MockSapBaseUrl"
        Write-Log "DRY RUN start mock SAP at $MockSapBaseUrl"
    }
    else {
        Invoke-StepScript -Action "start-mock-sap"
        Write-Log "Executed start-mock-sap"
    }
}

Show-ManualSummary
