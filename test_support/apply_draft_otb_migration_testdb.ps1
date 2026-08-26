param(
    [switch]$Apply
)

$ErrorActionPreference = 'Stop'
$expectedServer = '10.3.162.155'
$expectedDatabase = 'BMS'
$repoRoot = Split-Path -Parent $PSScriptRoot
$configPath = Join-Path $repoRoot 'Web.config'
$migrationPath = Join-Path $repoRoot 'database\07_create_draft_otb_background_jobs.sql'

if (-not $Apply) {
    throw 'This script changes the guarded Test database. Re-run with -Apply after reviewing the target.'
}

[xml]$config = [IO.File]::ReadAllText($configPath)
$connectionNode = @($config.configuration.connectionStrings.add) |
    Where-Object { $_.name -eq 'BMSConnectionString' } |
    Select-Object -First 1
if ($null -eq $connectionNode) {
    throw 'BMSConnectionString was not found in Web.config.'
}

$builder = [Data.SqlClient.SqlConnectionStringBuilder]::new([string]$connectionNode.connectionString)
if (-not [string]::Equals($builder.DataSource.Trim(), $expectedServer, [StringComparison]::OrdinalIgnoreCase)) {
    throw "Refusing migration: expected Test DB server $expectedServer."
}
if (-not [string]::Equals($builder.InitialCatalog.Trim(), $expectedDatabase, [StringComparison]::OrdinalIgnoreCase)) {
    throw "Refusing migration: expected database $expectedDatabase."
}
$builder.Encrypt = $false

function Invoke-SqlScalar {
    param(
        [Data.SqlClient.SqlConnection]$Connection,
        [string]$Sql,
        [Data.SqlClient.SqlTransaction]$Transaction = $null
    )
    $command = [Data.SqlClient.SqlCommand]::new($Sql, $Connection, $Transaction)
    try {
        $command.CommandTimeout = 600
        return $command.ExecuteScalar()
    } finally {
        $command.Dispose()
    }
}

function Invoke-MigrationBatches {
    param(
        [Data.SqlClient.SqlConnection]$Connection,
        [string]$Sql
    )
    $batches = [regex]::Split($Sql, '(?im)^\s*GO\s*(?:--.*)?$')
    foreach ($batch in $batches) {
        if ([string]::IsNullOrWhiteSpace($batch)) { continue }
        $command = [Data.SqlClient.SqlCommand]::new($batch, $Connection)
        try {
            $command.CommandTimeout = 600
            [void]$command.ExecuteNonQuery()
        } finally {
            $command.Dispose()
        }
    }
}

$connection = [Data.SqlClient.SqlConnection]::new($builder.ConnectionString)
try {
    $connection.Open()
    $identitySql = @'
SELECT DB_NAME(), CONVERT(nvarchar(128), CONNECTIONPROPERTY('local_net_address'));
'@
    $identityCommand = [Data.SqlClient.SqlCommand]::new($identitySql, $connection)
    try {
        $reader = $identityCommand.ExecuteReader()
        try {
            if (-not $reader.Read()) { throw 'Could not verify the connected database.' }
            $actualDatabase = $reader.GetString(0)
            $actualServer = if ($reader.IsDBNull(1)) { '' } else { $reader.GetString(1) }
        } finally {
            $reader.Dispose()
        }
    } finally {
        $identityCommand.Dispose()
    }

    if (-not [string]::Equals($actualDatabase, $expectedDatabase, [StringComparison]::OrdinalIgnoreCase) -or
        -not [string]::Equals($actualServer, $expectedServer, [StringComparison]::OrdinalIgnoreCase)) {
        throw 'Connected database identity did not match the guarded Test DB target.'
    }

    $migrationSql = [IO.File]::ReadAllText($migrationPath)
    Invoke-MigrationBatches -Connection $connection -Sql $migrationSql
    Invoke-MigrationBatches -Connection $connection -Sql $migrationSql

    $verificationSql = @'
SELECT COUNT_BIG(*)
FROM
(
    SELECT N'table-job' AS Item WHERE OBJECT_ID(N'dbo.Draft_OTB_Background_Job', N'U') IS NOT NULL
    UNION ALL SELECT N'table-claim' WHERE OBJECT_ID(N'dbo.Draft_OTB_Approval_Claim', N'U') IS NOT NULL
    UNION ALL SELECT N'ack-at' WHERE COL_LENGTH(N'dbo.Draft_OTB_Background_Job', N'AcknowledgedAt') IS NOT NULL
    UNION ALL SELECT N'ack-by' WHERE COL_LENGTH(N'dbo.Draft_OTB_Background_Job', N'AcknowledgedBy') IS NOT NULL
    UNION ALL SELECT N'idempotency-index' WHERE EXISTS
        (SELECT 1 FROM sys.indexes WHERE object_id = OBJECT_ID(N'dbo.Draft_OTB_Background_Job') AND name = N'UX_Draft_OTB_Job_Idempotency')
    UNION ALL SELECT N'claim-group-index' WHERE EXISTS
        (SELECT 1 FROM sys.indexes WHERE object_id = OBJECT_ID(N'dbo.Draft_OTB_Approval_Claim') AND name = N'IX_Draft_OTB_Approval_Claim_Group')
    UNION ALL SELECT N'template-search-index' WHERE EXISTS
        (SELECT 1 FROM sys.indexes WHERE object_id = OBJECT_ID(N'dbo.Template_Upload_Draft_OTB') AND name = N'IX_Template_Upload_Draft_OTB_Search')
) verified;
'@
    $verifiedCount = [int64](Invoke-SqlScalar -Connection $connection -Sql $verificationSql)
    if ($verifiedCount -ne 7) {
        throw "Migration verification found $verifiedCount of 7 required objects."
    }

    $jobId = [Guid]::NewGuid()
    $requestId = 'codex-test-' + [Guid]::NewGuid().ToString('N')
    $businessKey = 'codex-test-' + [Guid]::NewGuid().ToString('N')
    $sha = [Security.Cryptography.SHA256]::Create()
    try {
        $hashBytes = $sha.ComputeHash([Text.Encoding]::UTF8.GetBytes($businessKey))
        $businessHash = -join ($hashBytes | ForEach-Object { $_.ToString('x2') })
    } finally {
        $sha.Dispose()
    }

    $transaction = $connection.BeginTransaction()
    try {
        $smokeSql = @'
INSERT dbo.Draft_OTB_Background_Job
    (JobID, JobType, ClientRequestID, RequestedBy, [Status], Stage, ProgressPercent,
     TotalRows, ProcessedRows, SuccessRows, ErrorRows, CreatedAt, UpdatedAt, FinishedAt)
VALUES
    (@JobID, N'Upload', @RequestID, N'codex-test', N'ValidationFailed', N'Smoke', 100,
     1, 1, 0, 1, SYSUTCDATETIME(), SYSUTCDATETIME(), SYSUTCDATETIME());

INSERT dbo.Draft_OTB_Approval_Claim
    (JobID, RunNo, BusinessKey, BusinessKeyHash, [Type], [Year], [Month],
     Category, Company, Segment, Brand, Vendor, ClaimedAt)
VALUES
    (@JobID, -2147483647, @BusinessKey, @BusinessHash, N'Original', 2099, 12,
     N'ZZ', N'ZZ', N'ZZ', N'ZZ', N'ZZ', SYSUTCDATETIME());

UPDATE dbo.Draft_OTB_Background_Job
SET AcknowledgedAt = SYSUTCDATETIME(), AcknowledgedBy = N'codex-test'
WHERE JobID = @JobID AND [Status] = N'ValidationFailed';

SELECT COUNT_BIG(*)
FROM dbo.Draft_OTB_Background_Job j
INNER JOIN dbo.Draft_OTB_Approval_Claim c ON c.JobID = j.JobID
WHERE j.JobID = @JobID AND j.AcknowledgedAt IS NOT NULL AND j.AcknowledgedBy = N'codex-test';
'@
        $smokeCommand = [Data.SqlClient.SqlCommand]::new($smokeSql, $connection, $transaction)
        try {
            [void]$smokeCommand.Parameters.Add('@JobID', [Data.SqlDbType]::UniqueIdentifier)
            $smokeCommand.Parameters['@JobID'].Value = $jobId
            [void]$smokeCommand.Parameters.Add('@RequestID', [Data.SqlDbType]::NVarChar, 100)
            $smokeCommand.Parameters['@RequestID'].Value = $requestId
            [void]$smokeCommand.Parameters.Add('@BusinessKey', [Data.SqlDbType]::NVarChar, 500)
            $smokeCommand.Parameters['@BusinessKey'].Value = $businessKey
            [void]$smokeCommand.Parameters.Add('@BusinessHash', [Data.SqlDbType]::Char, 64)
            $smokeCommand.Parameters['@BusinessHash'].Value = $businessHash
            $smokeCommand.CommandTimeout = 600
            $smokeCount = [int64]$smokeCommand.ExecuteScalar()
        } finally {
            $smokeCommand.Dispose()
        }
        if ($smokeCount -ne 1) { throw 'Transactional job/claim/acknowledgement smoke test failed.' }
    } finally {
        $transaction.Rollback()
        $transaction.Dispose()
    }

    $postRollbackSql = 'SELECT COUNT_BIG(*) FROM dbo.Draft_OTB_Background_Job WHERE JobID = @JobID;'
    $postRollbackCommand = [Data.SqlClient.SqlCommand]::new($postRollbackSql, $connection)
    try {
        [void]$postRollbackCommand.Parameters.Add('@JobID', [Data.SqlDbType]::UniqueIdentifier)
        $postRollbackCommand.Parameters['@JobID'].Value = $jobId
        if ([int64]$postRollbackCommand.ExecuteScalar() -ne 0) {
            throw 'Smoke-test data remained after rollback.'
        }
    } finally {
        $postRollbackCommand.Dispose()
    }

    foreach ($tableName in @('dbo.Draft_OTB_Background_Job', 'dbo.Draft_OTB_Approval_Claim')) {
        $checkCommand = [Data.SqlClient.SqlCommand]::new("DBCC CHECKCONSTRAINTS ('$tableName') WITH ALL_CONSTRAINTS;", $connection)
        try {
            $checkCommand.CommandTimeout = 600
            $checkReader = $checkCommand.ExecuteReader()
            try {
                if ($checkReader.Read()) { throw "Constraint violations were found in $tableName." }
            } finally {
                $checkReader.Dispose()
            }
        } finally {
            $checkCommand.Dispose()
        }
    }

    Write-Output "PASS: migration applied twice and verified on Test DB $expectedServer/$expectedDatabase; rollback smoke left no rows."
} finally {
    $connection.Dispose()
}
