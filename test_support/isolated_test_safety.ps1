Set-StrictMode -Version Latest

function Get-IsolatedTestWebConfigTarget {
    param(
        [Parameter(Mandatory = $true)]
        [string]$WebConfigPath
    )

    if (-not (Test-Path -LiteralPath $WebConfigPath)) {
        return $null
    }

    $xml = New-Object System.Xml.XmlDocument
    $xml.PreserveWhitespace = $true
    $xml.Load($WebConfigPath)

    $node = $xml.SelectSingleNode("/configuration/connectionStrings/add[@name='BMSConnectionString']")
    if ($null -eq $node) {
        return $null
    }

    $connectionString = $node.GetAttribute("connectionString")
    if ([string]::IsNullOrWhiteSpace($connectionString)) {
        return $null
    }

    $builder = [System.Data.SqlClient.SqlConnectionStringBuilder]::new($connectionString)

    return [pscustomobject]@{
        DataSource = $builder.DataSource
        InitialCatalog = $builder.InitialCatalog
    }
}

function Normalize-IsolatedTestSqlServer {
    param(
        [AllowNull()]
        [string]$SqlServer
    )

    if ([string]::IsNullOrWhiteSpace($SqlServer)) {
        return ""
    }

    $normalized = $SqlServer.Trim().ToLowerInvariant()
    if ($normalized.StartsWith("tcp:")) {
        $normalized = $normalized.Substring(4)
    }

    return ($normalized -replace ',\d+$', '')
}

function Test-IsolatedTestCatalogName {
    param(
        [AllowNull()]
        [string]$DatabaseName
    )

    if ([string]::IsNullOrWhiteSpace($DatabaseName)) {
        return $false
    }

    return $DatabaseName -match '(?i)(test|uat|dev|sandbox)'
}

function Assert-IsolatedTestTarget {
    param(
        [Parameter(Mandatory = $true)]
        [string]$WebConfigPath,

        [Parameter(Mandatory = $true)]
        [string]$SqlServer,

        [Parameter(Mandatory = $true)]
        [string]$TestDatabase,

        [switch]$AllowProductionServer
    )

    if ([string]::IsNullOrWhiteSpace($SqlServer) -or $SqlServer -eq "YOUR_SERVER") {
        throw "Unsafe test target: please set -SqlServer to a non-production SQL Server before switching config or running SQL."
    }

    if ([string]::IsNullOrWhiteSpace($TestDatabase)) {
        throw "Unsafe test target: -TestDatabase is required."
    }

    if ($TestDatabase.Trim().Equals("BMS", [System.StringComparison]::OrdinalIgnoreCase)) {
        throw "Unsafe test target: -TestDatabase cannot be BMS. Use a separate database such as BMS_TEST."
    }

    $currentTarget = Get-IsolatedTestWebConfigTarget -WebConfigPath $WebConfigPath
    if ($null -eq $currentTarget) {
        return
    }

    if ($TestDatabase.Trim().Equals($currentTarget.InitialCatalog, [System.StringComparison]::OrdinalIgnoreCase) -and
        -not (Test-IsolatedTestCatalogName -DatabaseName $currentTarget.InitialCatalog)) {
        throw "Unsafe test target: -TestDatabase matches the current BMSConnectionString database '$($currentTarget.InitialCatalog)'. Use a separate test database."
    }

    $currentServer = Normalize-IsolatedTestSqlServer -SqlServer $currentTarget.DataSource
    $targetServer = Normalize-IsolatedTestSqlServer -SqlServer $SqlServer
    $currentIsAlreadyTargetDb = $TestDatabase.Trim().Equals($currentTarget.InitialCatalog, [System.StringComparison]::OrdinalIgnoreCase)

    if (-not $AllowProductionServer -and
        -not $currentIsAlreadyTargetDb -and
        $currentServer -ne "" -and
        $currentServer -eq $targetServer) {
        throw "Unsafe test target: -SqlServer matches the current BMSConnectionString server '$($currentTarget.DataSource)'. Use a local/UAT SQL Server, or pass -AllowProductionServer only after DBA approval."
    }
}
