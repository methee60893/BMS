[CmdletBinding()]
param(
    [Parameter(Position = 0)]
    [ValidateSet("status", "backup", "generate-test-sql", "switch-to-test", "start-mock-sap", "print-sql", "restore")]
    [string]$Action = "status",

    [string]$SqlServer = "YOUR_SERVER",
    [string]$TestDatabase = "BMS_TEST",
    [string]$DbUser = "sa",
    [string]$DbPassword = "YOUR_PASSWORD",
    [string]$MockSapBaseUrl = "http://127.0.0.1:18080",
    [string]$MockSapUsername = "mock-user",
    [string]$MockSapPassword = "mock-password",
    [switch]$AllowProductionServer
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$root = Split-Path -Parent $MyInvocation.MyCommand.Path
$webConfigPath = Join-Path $root "Web.config"
$stateDir = Join-Path $root ".codex-test-state"
$backupWebConfigPath = Join-Path $stateDir "Web.config.backup"
$snapshotPath = Join-Path $stateDir "snapshot.json"
$logPath = Join-Path $stateDir "run_test_step_by_step.log"
$mockServerScript = Join-Path $root "test_support\mock_sap_server.py"
$safetyScript = Join-Path $root "test_support\isolated_test_safety.ps1"
$generateTestScript = Join-Path $root "database\generate_test_db_scripts.py"
$testSqlDir = Join-Path $root "database\test"

. $safetyScript

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

function Load-Xml([string]$path) {
    $xml = New-Object System.Xml.XmlDocument
    $xml.PreserveWhitespace = $true
    $xml.Load($path)
    return $xml
}

function Save-Xml($xml, [string]$path) {
    $settings = New-Object System.Xml.XmlWriterSettings
    $settings.Indent = $true
    $settings.Encoding = [System.Text.UTF8Encoding]::new($false)
    $writer = [System.Xml.XmlWriter]::Create($path, $settings)
    try {
        $xml.Save($writer)
    }
    finally {
        $writer.Dispose()
    }
}

function Get-ConnectionNode($xml, [string]$name) {
    return $xml.SelectSingleNode("/configuration/connectionStrings/add[@name='$name']")
}

function Get-AppSettingNode($xml, [string]$key) {
    return $xml.SelectSingleNode("/configuration/appSettings/add[@key='$key']")
}

function Mask-ConnectionStringSecret {
    param(
        [AllowNull()]
        [string]$ConnectionString
    )

    if ([string]::IsNullOrWhiteSpace($ConnectionString)) {
        return $ConnectionString
    }

    $masked = $ConnectionString -replace '(?i)(Password|Pwd)\s*=\s*[^;]*', '$1=***'
    $masked = $masked -replace '(?i)(User ID|User Id|UID)\s*=\s*[^;]*', '$1=***'
    return $masked
}

function Backup-WebConfig {
    Ensure-StateDir
    if (-not (Test-Path -LiteralPath $webConfigPath)) {
        Write-Log "Web.config not found at $webConfigPath" "ERROR"
        throw "Web.config not found at $webConfigPath"
    }
    if (-not (Test-Path -LiteralPath $backupWebConfigPath)) {
        Copy-Item -LiteralPath $webConfigPath -Destination $backupWebConfigPath -Force
        Write-Log "Backup created at $backupWebConfigPath"
        Write-Host "Backup created: $backupWebConfigPath"
    }
    else {
        Write-Log "Backup already exists at $backupWebConfigPath"
        Write-Host "Backup already exists: $backupWebConfigPath"
    }
}

function Save-Snapshot {
    param(
        [string]$BmsConnectionString,
        [string]$SapBaseUrl,
        [string]$SapUsername
    )

    Ensure-StateDir
    $payload = [ordered]@{
        UpdatedAt = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
        BmsConnectionString = (Mask-ConnectionStringSecret -ConnectionString $BmsConnectionString)
        SapBaseUrl = $SapBaseUrl
        SapUsername = $SapUsername
        TestDatabase = $TestDatabase
    } | ConvertTo-Json -Depth 4
    Set-Content -LiteralPath $snapshotPath -Value $payload -Encoding UTF8
}

function Show-Status {
    Write-Log "Checking current status"
    Write-Host "Workspace: $root"
    Write-Host "Web.config: $webConfigPath"
    Write-Host "Backup exists: $(Test-Path -LiteralPath $backupWebConfigPath)"
    Write-Host "Snapshot exists: $(Test-Path -LiteralPath $snapshotPath)"
    Write-Host "Log file: $logPath"

    if (Test-Path -LiteralPath $webConfigPath) {
        $xml = Load-Xml $webConfigPath
        $bmsConn = Get-ConnectionNode $xml "BMSConnectionString"
        $sapBase = Get-AppSettingNode $xml "SAPAPI_BASEURL"
        Write-Host "Current BMSConnectionString:"
        Write-Host ("  " + (Mask-ConnectionStringSecret -ConnectionString $bmsConn.GetAttribute("connectionString")))
        Write-Host "Current SAPAPI_BASEURL:"
        Write-Host ("  " + $sapBase.GetAttribute("value"))
    }
}

function Generate-TestSql {
    Write-Log "Generating test SQL files"
    & python $generateTestScript
    if ($LASTEXITCODE -ne 0) {
        Write-Log "Failed to generate test SQL files" "ERROR"
        throw "Failed to generate test SQL files."
    }
    Write-Host "Generated test SQL in: $testSqlDir"
    Write-Log "Generated test SQL files in $testSqlDir"
}

function Switch-ToTest {
    Assert-IsolatedTestTarget -WebConfigPath $webConfigPath -SqlServer $SqlServer -TestDatabase $TestDatabase -AllowProductionServer:$AllowProductionServer
    Backup-WebConfig
    Write-Log "Switching Web.config to test environment"

    $xml = Load-Xml $webConfigPath
    $bmsConn = Get-ConnectionNode $xml "BMSConnectionString"
    if ($null -eq $bmsConn) {
        Write-Log "BMSConnectionString node not found" "ERROR"
        throw "BMSConnectionString node not found."
    }

    $testConnString = "Data Source=$SqlServer;Initial Catalog=$TestDatabase;Persist Security Info=True;User ID=$DbUser;Password=$DbPassword;TrustServerCertificate=true;"
    $bmsConn.SetAttribute("connectionString", $testConnString)

    $sapBase = Get-AppSettingNode $xml "SAPAPI_BASEURL"
    $sapUser = Get-AppSettingNode $xml "SAPAPI_USERNAME"
    $sapPass = Get-AppSettingNode $xml "SAPAPI_PASSWORD"

    if ($null -eq $sapBase -or $null -eq $sapUser -or $null -eq $sapPass) {
        Write-Log "SAP appSettings nodes are incomplete in Web.config" "ERROR"
        throw "SAP appSettings nodes are incomplete in Web.config."
    }

    $sapBase.SetAttribute("value", $MockSapBaseUrl)
    $sapUser.SetAttribute("value", $MockSapUsername)
    $sapPass.SetAttribute("value", $MockSapPassword)

    Save-Xml $xml $webConfigPath
    Save-Snapshot -BmsConnectionString $testConnString -SapBaseUrl $MockSapBaseUrl -SapUsername $MockSapUsername

    Write-Host "Switched Web.config to test environment."
    Write-Host "  Database : $TestDatabase"
    Write-Host "  SQL Server : $SqlServer"
    Write-Host "  Mock SAP : $MockSapBaseUrl"
    Write-Log "Web.config switched to test environment. Database=$TestDatabase, SqlServer=$SqlServer, MockSap=$MockSapBaseUrl"
}

function Restore-Normal {
    if (-not (Test-Path -LiteralPath $backupWebConfigPath)) {
        Write-Log "Restore requested but backup not found" "ERROR"
        throw "Backup Web.config not found. Nothing to restore."
    }

    Copy-Item -LiteralPath $backupWebConfigPath -Destination $webConfigPath -Force
    Remove-Item -LiteralPath $backupWebConfigPath -Force
    if (Test-Path -LiteralPath $snapshotPath) {
        Remove-Item -LiteralPath $snapshotPath -Force
    }

    Write-Host "Restored Web.config to the original environment."
    Write-Log "Web.config restored to original environment"
}

function Test-PortInUse {
    param(
        [int]$Port
    )

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

function Start-MockSap {
    $uri = [System.Uri]$MockSapBaseUrl
    $port = $uri.Port
    $host = $uri.Host
    if (Test-PortInUse -Port $port) {
        Write-Host "Mock SAP port $port is already in use. Skip starting a new server."
        Write-Log "Mock SAP port $port already in use. Skipped starting new server." "WARN"
        return
    }

    $cmd = "python `"$mockServerScript`" --host $host --port $port"
    Start-Process powershell -ArgumentList "-NoExit", "-Command", $cmd | Out-Null
    Write-Host "Mock SAP started in a new PowerShell window at $MockSapBaseUrl"
    Write-Log "Started mock SAP in new PowerShell window at $MockSapBaseUrl"
}

function Print-SqlSteps {
    Assert-IsolatedTestTarget -WebConfigPath $webConfigPath -SqlServer $SqlServer -TestDatabase $TestDatabase -AllowProductionServer:$AllowProductionServer
    Write-Log "Printing SQL execution steps"
    $testCreateDb = Join-Path $root "database\00_create_database_test.sql"
    $sqlFiles = @(
        (Join-Path $testSqlDir "01_create_tables_test.sql"),
        (Join-Path $testSqlDir "02_create_views_test.sql"),
        (Join-Path $testSqlDir "03_create_stored_procedures_test.sql"),
        (Join-Path $testSqlDir "04_seed_master_data_template_test.sql"),
        (Join-Path $testSqlDir "05_post_deploy_verify_test.sql")
    )

    Write-Host "Run these commands in order:"
    Write-Host "sqlcmd -S $SqlServer -E -i `"$testCreateDb`""
    foreach ($sqlFile in $sqlFiles) {
        Write-Host "sqlcmd -S $SqlServer -E -d $TestDatabase -i `"$sqlFile`""
    }
}

switch ($Action) {
    "status" { Show-Status }
    "backup" { Backup-WebConfig }
    "generate-test-sql" { Generate-TestSql }
    "switch-to-test" { Switch-ToTest }
    "start-mock-sap" { Start-MockSap }
    "print-sql" { Print-SqlSteps }
    "restore" { Restore-Normal }
}
