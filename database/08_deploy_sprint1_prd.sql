/*
    KBMS Sprint 1 - production database deployment
    Intended target: TCP 10.3.152.155 / database BMS

    Deployment order:
      1. Back up BMS and stop new Draft OTB upload/approval activity.
      2. Run this script while connected by TCP to 10.3.152.155.
      3. Deploy application commit d7b8452 or later.

    Existing business-data impact:
      - No value is changed in Template_Upload_Draft_OTB, OTB_Transaction,
        OTB_Switching_Transaction, or Draft_PO_Transaction.
      - Two operational tables and supporting indexes are created.
      - A non-unique Year/Month index is added to Template_Upload_Draft_OTB
        only when an index with the expected deployment name does not exist.
      - The script fails before DDL when active Draft keys are invalid/duplicated
        or a background/claim job is still active.

    The script is idempotent and may be run again after a successful deploy.
    It intentionally refuses connections whose local SQL listener address is
    not exactly the production IP. A NULL address means the connection is not
    TCP (for example Shared Memory) and is also refused.
*/

USE [BMS];
GO

SET NOCOUNT ON;
IF @@TRANCOUNT <> 0
BEGIN
    /* RAISERROR intentionally runs before XACT_ABORT so the script does not
       roll back unrelated work owned by the caller's existing transaction. */
    RAISERROR('Refusing Sprint 1 deployment: run the script in a new session with no ambient transaction.', 16, 1);
    RETURN;
END;
SET IMPLICIT_TRANSACTIONS OFF;
SET XACT_ABORT ON;
SET LOCK_TIMEOUT 30000;

DECLARE @ExpectedDatabase sysname = N'BMS';
DECLARE @ExpectedLocalAddress nvarchar(128) = N'10.3.152.155';
DECLARE @ActualLocalAddress nvarchar(128) = CONVERT(nvarchar(128), CONNECTIONPROPERTY('local_net_address'));
DECLARE @CompatibilityLevel int = (SELECT compatibility_level FROM sys.databases WHERE database_id = DB_ID());
DECLARE @ProductMajorVersion int = TRY_CONVERT(int, SERVERPROPERTY('ProductMajorVersion'));

IF DB_NAME() <> @ExpectedDatabase
    THROW 51100, 'Refusing Sprint 1 deployment: the selected database is not BMS.', 1;

IF @ActualLocalAddress IS NULL OR @ActualLocalAddress <> @ExpectedLocalAddress
    THROW 51101, 'Refusing Sprint 1 deployment: connect by TCP to production listener 10.3.152.155.', 1;

IF @ProductMajorVersion IS NULL OR @ProductMajorVersion < 12 OR @CompatibilityLevel < 120
    THROW 51102, 'Refusing Sprint 1 deployment: SQL Server 2014+ and database compatibility level 120+ are required.', 1;

IF OBJECT_ID(N'dbo.Template_Upload_Draft_OTB', N'U') IS NULL
    THROW 51103, 'Refusing Sprint 1 deployment: dbo.Template_Upload_Draft_OTB does not exist.', 1;

IF COL_LENGTH(N'dbo.Template_Upload_Draft_OTB', N'RunNo') IS NULL
   OR COL_LENGTH(N'dbo.Template_Upload_Draft_OTB', N'Type') IS NULL
   OR COL_LENGTH(N'dbo.Template_Upload_Draft_OTB', N'Year') IS NULL
   OR COL_LENGTH(N'dbo.Template_Upload_Draft_OTB', N'Month') IS NULL
   OR COL_LENGTH(N'dbo.Template_Upload_Draft_OTB', N'Category') IS NULL
   OR COL_LENGTH(N'dbo.Template_Upload_Draft_OTB', N'Company') IS NULL
   OR COL_LENGTH(N'dbo.Template_Upload_Draft_OTB', N'Segment') IS NULL
   OR COL_LENGTH(N'dbo.Template_Upload_Draft_OTB', N'Brand') IS NULL
   OR COL_LENGTH(N'dbo.Template_Upload_Draft_OTB', N'Vendor') IS NULL
   OR COL_LENGTH(N'dbo.Template_Upload_Draft_OTB', N'OTBStatus') IS NULL
    THROW 51104, 'Refusing Sprint 1 deployment: the Draft OTB base table is missing required columns.', 1;

IF EXISTS
(
    SELECT 1
    FROM sys.columns c
    INNER JOIN sys.types t ON t.user_type_id = c.user_type_id
    WHERE c.object_id = OBJECT_ID(N'dbo.Template_Upload_Draft_OTB')
      AND
      (
          (c.name IN (N'RunNo', N'Year', N'Month') AND t.name <> N'int')
          OR
          (c.name IN (N'Type', N'Category', N'Company', N'Segment', N'Brand', N'Vendor', N'OTBStatus')
           AND t.name NOT IN (N'nvarchar', N'varchar', N'nchar', N'char'))
      )
)
    THROW 51105, 'Refusing Sprint 1 deployment: unsupported Draft OTB key column types were found.', 1;

/* Fail closed before changing schema. These bounds match application staging
   and durable-claim columns. No existing business row is changed. */
IF EXISTS
(
    SELECT 1
    FROM dbo.Template_Upload_Draft_OTB
    WHERE (OTBStatus IS NULL OR CONVERT(nvarchar(30), OTBStatus) = N'Draft')
      AND
      (
          [Type] IS NULL OR LEN(CONVERT(nvarchar(max), [Type])) = 0 OR LEN(CONVERT(nvarchar(max), [Type])) > 20
          OR [Year] IS NULL OR [Year] <= 0
          OR [Month] IS NULL OR [Month] NOT BETWEEN 1 AND 12
          OR Category IS NULL OR LEN(CONVERT(nvarchar(max), Category)) = 0 OR LEN(CONVERT(nvarchar(max), Category)) > 20
          OR Company IS NULL OR LEN(CONVERT(nvarchar(max), Company)) = 0 OR LEN(CONVERT(nvarchar(max), Company)) > 20
          OR Segment IS NULL OR LEN(CONVERT(nvarchar(max), Segment)) = 0 OR LEN(CONVERT(nvarchar(max), Segment)) > 20
          OR Brand IS NULL OR LEN(CONVERT(nvarchar(max), Brand)) = 0 OR LEN(CONVERT(nvarchar(max), Brand)) > 30
          OR Vendor IS NULL OR LEN(CONVERT(nvarchar(max), Vendor)) = 0 OR LEN(CONVERT(nvarchar(max), Vendor)) > 30
      )
)
    THROW 51106, 'Preflight failed: an active Draft OTB row has a NULL, blank, out-of-range, or overlength business key.', 1;

IF EXISTS
(
    SELECT 1
    FROM dbo.Template_Upload_Draft_OTB
    WHERE OTBStatus IS NULL OR CONVERT(nvarchar(30), OTBStatus) = N'Draft'
    GROUP BY CONVERT(nvarchar(20), [Type]), [Year], [Month],
             CONVERT(nvarchar(20), Company), CONVERT(nvarchar(20), Category),
             CONVERT(nvarchar(20), Segment), CONVERT(nvarchar(30), Brand),
             CONVERT(nvarchar(30), Vendor)
    HAVING COUNT_BIG(*) > 1
)
    THROW 51107, 'Preflight failed: duplicate active Draft OTB business keys must be resolved before deployment.', 1;

/* A deploy/IIS recycle must not interrupt an active job. Dynamic SQL keeps
   this preflight valid on both first install and idempotent reruns. */
IF OBJECT_ID(N'dbo.Draft_OTB_Background_Job', N'U') IS NOT NULL
BEGIN
    EXEC sys.sp_executesql N'
        IF EXISTS
        (
            SELECT 1
            FROM dbo.Draft_OTB_Background_Job
            WHERE [Status] IS NULL OR [Status] NOT IN
                  (N''Completed'', N''ValidationFailed'', N''Failed'', N''ReconciliationRequired'')
        )
            THROW 51108, ''Preflight failed: a Draft OTB background job is still active.'', 1;

        IF EXISTS
        (
            SELECT JobType, RequestedBy, ClientRequestID
            FROM dbo.Draft_OTB_Background_Job
            GROUP BY JobType, RequestedBy, ClientRequestID
            HAVING COUNT_BIG(*) > 1
        )
            THROW 51109, ''Preflight failed: duplicate background-job idempotency keys were found.'', 1;';
END;

IF OBJECT_ID(N'dbo.Draft_OTB_Approval_Claim', N'U') IS NOT NULL
BEGIN
    EXEC sys.sp_executesql N'
        IF EXISTS (SELECT 1 FROM dbo.Draft_OTB_Approval_Claim)
            THROW 51110, ''Preflight failed: approval/reconciliation claims still exist. Resolve them before deployment.'', 1;';
END;

PRINT N'PASS: production target and Sprint 1 data preflight checks succeeded.';

BEGIN TRY
    BEGIN TRANSACTION;

    DECLARE @DeploymentLockResult int;
    EXEC @DeploymentLockResult = sys.sp_getapplock
        @Resource = N'KBMS:Sprint1:ProductionSchemaDeployment',
        @LockMode = N'Exclusive',
        @LockOwner = N'Transaction',
        @LockTimeout = 0;
    IF @DeploymentLockResult < 0
        THROW 51119, 'Deployment failed: another Sprint 1 schema deployment is already running.', 1;

    /* Close the preflight-to-DDL race. Existing operational tables remain
       exclusively locked until commit, so no worker can start while their
       schema and indexes are being checked or upgraded. */
    IF OBJECT_ID(N'dbo.Draft_OTB_Background_Job', N'U') IS NOT NULL
    BEGIN
        EXEC sys.sp_executesql N'
            IF EXISTS
            (
                SELECT 1
                FROM dbo.Draft_OTB_Background_Job WITH (TABLOCKX, HOLDLOCK)
                WHERE [Status] IS NULL OR [Status] NOT IN
                      (N''Completed'', N''ValidationFailed'', N''Failed'', N''ReconciliationRequired'')
            )
                THROW 51108, ''Deployment failed: a Draft OTB background job became active during deployment.'', 1;';
    END;

    IF OBJECT_ID(N'dbo.Draft_OTB_Approval_Claim', N'U') IS NOT NULL
    BEGIN
        EXEC sys.sp_executesql N'
            IF EXISTS (SELECT 1 FROM dbo.Draft_OTB_Approval_Claim WITH (TABLOCKX, HOLDLOCK))
                THROW 51110, ''Deployment failed: an approval/reconciliation claim appeared during deployment.'', 1;';
    END;

    IF OBJECT_ID(N'dbo.Draft_OTB_Background_Job', N'U') IS NULL
    BEGIN
        EXEC sys.sp_executesql N'
            CREATE TABLE dbo.Draft_OTB_Background_Job
            (
                JobID uniqueidentifier NOT NULL,
                JobType nvarchar(30) NOT NULL,
                ClientRequestID nvarchar(100) NOT NULL,
                RequestedBy nvarchar(100) NOT NULL,
                [Status] nvarchar(30) NOT NULL,
                Stage nvarchar(100) NOT NULL,
                ProgressPercent tinyint NOT NULL CONSTRAINT DF_Draft_OTB_Job_Progress DEFAULT (0),
                TotalRows int NOT NULL CONSTRAINT DF_Draft_OTB_Job_Total DEFAULT (0),
                ProcessedRows int NOT NULL CONSTRAINT DF_Draft_OTB_Job_Processed DEFAULT (0),
                SuccessRows int NOT NULL CONSTRAINT DF_Draft_OTB_Job_Success DEFAULT (0),
                ErrorRows int NOT NULL CONSTRAINT DF_Draft_OTB_Job_Error DEFAULT (0),
                [Message] nvarchar(max) NULL,
                PayloadJson nvarchar(max) NULL,
                ResultJson nvarchar(max) NULL,
                CreatedAt datetime2(3) NOT NULL CONSTRAINT DF_Draft_OTB_Job_Created DEFAULT (sysutcdatetime()),
                UpdatedAt datetime2(3) NOT NULL CONSTRAINT DF_Draft_OTB_Job_Updated DEFAULT (sysutcdatetime()),
                StartedAt datetime2(3) NULL,
                FinishedAt datetime2(3) NULL,
                AcknowledgedAt datetime2(3) NULL,
                AcknowledgedBy nvarchar(100) NULL,
                CONSTRAINT PK_Draft_OTB_Background_Job PRIMARY KEY CLUSTERED (JobID),
                CONSTRAINT CK_Draft_OTB_Job_Progress CHECK (ProgressPercent BETWEEN 0 AND 100),
                CONSTRAINT CK_Draft_OTB_Job_Counts CHECK
                    (TotalRows >= 0 AND ProcessedRows >= 0 AND SuccessRows >= 0 AND ErrorRows >= 0)
            );';
    END;

    IF COL_LENGTH(N'dbo.Draft_OTB_Background_Job', N'AcknowledgedAt') IS NULL
        EXEC sys.sp_executesql N'ALTER TABLE dbo.Draft_OTB_Background_Job ADD AcknowledgedAt datetime2(3) NULL;';
    IF COL_LENGTH(N'dbo.Draft_OTB_Background_Job', N'AcknowledgedBy') IS NULL
        EXEC sys.sp_executesql N'ALTER TABLE dbo.Draft_OTB_Background_Job ADD AcknowledgedBy nvarchar(100) NULL;';

    IF COL_LENGTH(N'dbo.Draft_OTB_Background_Job', N'JobID') IS NULL
       OR COL_LENGTH(N'dbo.Draft_OTB_Background_Job', N'JobType') IS NULL
       OR COL_LENGTH(N'dbo.Draft_OTB_Background_Job', N'ClientRequestID') IS NULL
       OR COL_LENGTH(N'dbo.Draft_OTB_Background_Job', N'RequestedBy') IS NULL
       OR COL_LENGTH(N'dbo.Draft_OTB_Background_Job', N'Status') IS NULL
       OR COL_LENGTH(N'dbo.Draft_OTB_Background_Job', N'Stage') IS NULL
       OR COL_LENGTH(N'dbo.Draft_OTB_Background_Job', N'ProgressPercent') IS NULL
       OR COL_LENGTH(N'dbo.Draft_OTB_Background_Job', N'TotalRows') IS NULL
       OR COL_LENGTH(N'dbo.Draft_OTB_Background_Job', N'ProcessedRows') IS NULL
       OR COL_LENGTH(N'dbo.Draft_OTB_Background_Job', N'SuccessRows') IS NULL
       OR COL_LENGTH(N'dbo.Draft_OTB_Background_Job', N'ErrorRows') IS NULL
       OR COL_LENGTH(N'dbo.Draft_OTB_Background_Job', N'Message') IS NULL
       OR COL_LENGTH(N'dbo.Draft_OTB_Background_Job', N'PayloadJson') IS NULL
       OR COL_LENGTH(N'dbo.Draft_OTB_Background_Job', N'ResultJson') IS NULL
       OR COL_LENGTH(N'dbo.Draft_OTB_Background_Job', N'CreatedAt') IS NULL
       OR COL_LENGTH(N'dbo.Draft_OTB_Background_Job', N'UpdatedAt') IS NULL
       OR COL_LENGTH(N'dbo.Draft_OTB_Background_Job', N'StartedAt') IS NULL
       OR COL_LENGTH(N'dbo.Draft_OTB_Background_Job', N'FinishedAt') IS NULL
       OR COL_LENGTH(N'dbo.Draft_OTB_Background_Job', N'AcknowledgedAt') IS NULL
       OR COL_LENGTH(N'dbo.Draft_OTB_Background_Job', N'AcknowledgedBy') IS NULL
        THROW 51111, 'Deployment failed: the background-job table has an incompatible shape.', 1;

    IF EXISTS
    (
        SELECT 1
        FROM
        (
            VALUES
                (CAST(N'JobID' AS sysname), CAST(N'uniqueidentifier' AS sysname), CAST(16 AS smallint), CAST(NULL AS tinyint), CAST(0 AS bit)),
                (N'JobType', N'nvarchar', 60, NULL, 0),
                (N'ClientRequestID', N'nvarchar', 200, NULL, 0),
                (N'RequestedBy', N'nvarchar', 200, NULL, 0),
                (N'Status', N'nvarchar', 60, NULL, 0),
                (N'Stage', N'nvarchar', 200, NULL, 0),
                (N'ProgressPercent', N'tinyint', 1, NULL, 0),
                (N'TotalRows', N'int', 4, NULL, 0),
                (N'ProcessedRows', N'int', 4, NULL, 0),
                (N'SuccessRows', N'int', 4, NULL, 0),
                (N'ErrorRows', N'int', 4, NULL, 0),
                (N'Message', N'nvarchar', -1, NULL, 1),
                (N'PayloadJson', N'nvarchar', -1, NULL, 1),
                (N'ResultJson', N'nvarchar', -1, NULL, 1),
                (N'CreatedAt', N'datetime2', 7, 3, 0),
                (N'UpdatedAt', N'datetime2', 7, 3, 0),
                (N'StartedAt', N'datetime2', 7, 3, 1),
                (N'FinishedAt', N'datetime2', 7, 3, 1),
                (N'AcknowledgedAt', N'datetime2', 7, 3, 1),
                (N'AcknowledgedBy', N'nvarchar', 200, NULL, 1)
        ) expected(ColumnName, TypeName, MaxLength, ScaleValue, IsNullable)
        LEFT JOIN sys.columns c
               ON c.object_id = OBJECT_ID(N'dbo.Draft_OTB_Background_Job')
              AND c.name = expected.ColumnName
        LEFT JOIN sys.types t ON t.user_type_id = c.user_type_id
        WHERE c.column_id IS NULL
           OR t.name <> expected.TypeName
           OR c.max_length <> expected.MaxLength
           OR (expected.ScaleValue IS NOT NULL AND c.scale <> expected.ScaleValue)
           OR c.is_nullable <> expected.IsNullable
    )
        THROW 51120, 'Deployment failed: background-job column types, lengths, or nullability do not match Sprint 1.', 1;

    IF NOT EXISTS
       (
           SELECT 1
           FROM sys.key_constraints kc
           INNER JOIN sys.indexes i
                   ON i.object_id = kc.parent_object_id
                  AND i.index_id = kc.unique_index_id
           WHERE kc.parent_object_id = OBJECT_ID(N'dbo.Draft_OTB_Background_Job')
             AND kc.name = N'PK_Draft_OTB_Background_Job'
             AND kc.[type] = N'PK'
             AND i.[type] = 1 AND i.is_unique = 1 AND i.has_filter = 0
             AND i.ignore_dup_key = 0 AND i.is_disabled = 0 AND i.is_hypothetical = 0
             AND (SELECT COUNT(*)
                  FROM sys.index_columns ic
                  WHERE ic.object_id = i.object_id
                    AND ic.index_id = i.index_id
                    AND ic.key_ordinal > 0) = 1
             AND EXISTS
                 (
                     SELECT 1
                     FROM sys.index_columns ic
                     INNER JOIN sys.columns c
                             ON c.object_id = ic.object_id
                            AND c.column_id = ic.column_id
                     WHERE ic.object_id = i.object_id
                       AND ic.index_id = i.index_id
                       AND ic.key_ordinal = 1
                       AND ic.is_descending_key = 0
                       AND c.name = N'JobID'
                 )
       )
       OR NOT EXISTS
          (
              SELECT 1
              FROM sys.check_constraints
              WHERE parent_object_id = OBJECT_ID(N'dbo.Draft_OTB_Background_Job')
                AND name = N'CK_Draft_OTB_Job_Progress'
                AND is_disabled = 0
                AND is_not_trusted = 0
                AND LOWER(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(
                    definition, N'[', N''), N']', N''), N'(', N''), N')', N''), N' ', N''),
                    CHAR(9), N''), CHAR(13), N''), CHAR(10), N'')) =
                    N'progresspercent>=0andprogresspercent<=100'
          )
       OR NOT EXISTS
          (
              SELECT 1
              FROM sys.check_constraints
              WHERE parent_object_id = OBJECT_ID(N'dbo.Draft_OTB_Background_Job')
                AND name = N'CK_Draft_OTB_Job_Counts'
                AND is_disabled = 0
                AND is_not_trusted = 0
                AND LOWER(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(
                    definition, N'[', N''), N']', N''), N'(', N''), N')', N''), N' ', N''),
                    CHAR(9), N''), CHAR(13), N''), CHAR(10), N'')) =
                    N'totalrows>=0andprocessedrows>=0andsuccessrows>=0anderrorrows>=0'
          )
        THROW 51112, 'Deployment failed: required background-job constraints are missing.', 1;

    IF OBJECT_ID(N'dbo.Draft_OTB_Approval_Claim', N'U') IS NULL
    BEGIN
        EXEC sys.sp_executesql N'
            CREATE TABLE dbo.Draft_OTB_Approval_Claim
            (
                ClaimID bigint IDENTITY(1,1) NOT NULL,
                JobID uniqueidentifier NOT NULL,
                RunNo int NOT NULL,
                BusinessKey nvarchar(500) NOT NULL,
                BusinessKeyHash char(64) NOT NULL,
                [Type] nvarchar(20) NOT NULL,
                [Year] int NOT NULL,
                [Month] int NOT NULL,
                Category nvarchar(20) NOT NULL,
                Company nvarchar(20) NOT NULL,
                Segment nvarchar(20) NOT NULL,
                Brand nvarchar(30) NOT NULL,
                Vendor nvarchar(30) NOT NULL,
                ClaimedAt datetime2(3) NOT NULL CONSTRAINT DF_Draft_OTB_Approval_Claimed DEFAULT (sysutcdatetime()),
                CONSTRAINT PK_Draft_OTB_Approval_Claim PRIMARY KEY CLUSTERED (ClaimID),
                CONSTRAINT FK_Draft_OTB_Approval_Claim_Job FOREIGN KEY (JobID)
                    REFERENCES dbo.Draft_OTB_Background_Job (JobID) ON DELETE CASCADE,
                CONSTRAINT UQ_Draft_OTB_Approval_Claim_RunNo UNIQUE (RunNo),
                CONSTRAINT UQ_Draft_OTB_Approval_Claim_BusinessKeyHash UNIQUE (BusinessKeyHash)
            );';
    END;

    IF COL_LENGTH(N'dbo.Draft_OTB_Approval_Claim', N'Type') IS NULL
        EXEC sys.sp_executesql N'ALTER TABLE dbo.Draft_OTB_Approval_Claim ADD [Type] nvarchar(20) NULL;';
    IF COL_LENGTH(N'dbo.Draft_OTB_Approval_Claim', N'Year') IS NULL
        EXEC sys.sp_executesql N'ALTER TABLE dbo.Draft_OTB_Approval_Claim ADD [Year] int NULL;';
    IF COL_LENGTH(N'dbo.Draft_OTB_Approval_Claim', N'Month') IS NULL
        EXEC sys.sp_executesql N'ALTER TABLE dbo.Draft_OTB_Approval_Claim ADD [Month] int NULL;';
    IF COL_LENGTH(N'dbo.Draft_OTB_Approval_Claim', N'Category') IS NULL
        EXEC sys.sp_executesql N'ALTER TABLE dbo.Draft_OTB_Approval_Claim ADD Category nvarchar(20) NULL;';
    IF COL_LENGTH(N'dbo.Draft_OTB_Approval_Claim', N'Company') IS NULL
        EXEC sys.sp_executesql N'ALTER TABLE dbo.Draft_OTB_Approval_Claim ADD Company nvarchar(20) NULL;';
    IF COL_LENGTH(N'dbo.Draft_OTB_Approval_Claim', N'Segment') IS NULL
        EXEC sys.sp_executesql N'ALTER TABLE dbo.Draft_OTB_Approval_Claim ADD Segment nvarchar(20) NULL;';
    IF COL_LENGTH(N'dbo.Draft_OTB_Approval_Claim', N'Brand') IS NULL
        EXEC sys.sp_executesql N'ALTER TABLE dbo.Draft_OTB_Approval_Claim ADD Brand nvarchar(30) NULL;';
    IF COL_LENGTH(N'dbo.Draft_OTB_Approval_Claim', N'Vendor') IS NULL
        EXEC sys.sp_executesql N'ALTER TABLE dbo.Draft_OTB_Approval_Claim ADD Vendor nvarchar(30) NULL;';

    /* The outer preflight and transactional TABLOCKX guarantee that this table
       is empty, so legacy nullable columns can be hardened without data DML. */

    IF EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID(N'dbo.Draft_OTB_Approval_Claim') AND name = N'Type' AND is_nullable = 1)
        EXEC sys.sp_executesql N'ALTER TABLE dbo.Draft_OTB_Approval_Claim ALTER COLUMN [Type] nvarchar(20) NOT NULL;';
    IF EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID(N'dbo.Draft_OTB_Approval_Claim') AND name = N'Year' AND is_nullable = 1)
        EXEC sys.sp_executesql N'ALTER TABLE dbo.Draft_OTB_Approval_Claim ALTER COLUMN [Year] int NOT NULL;';
    IF EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID(N'dbo.Draft_OTB_Approval_Claim') AND name = N'Month' AND is_nullable = 1)
        EXEC sys.sp_executesql N'ALTER TABLE dbo.Draft_OTB_Approval_Claim ALTER COLUMN [Month] int NOT NULL;';
    IF EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID(N'dbo.Draft_OTB_Approval_Claim') AND name = N'Category' AND is_nullable = 1)
        EXEC sys.sp_executesql N'ALTER TABLE dbo.Draft_OTB_Approval_Claim ALTER COLUMN Category nvarchar(20) NOT NULL;';
    IF EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID(N'dbo.Draft_OTB_Approval_Claim') AND name = N'Company' AND is_nullable = 1)
        EXEC sys.sp_executesql N'ALTER TABLE dbo.Draft_OTB_Approval_Claim ALTER COLUMN Company nvarchar(20) NOT NULL;';
    IF EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID(N'dbo.Draft_OTB_Approval_Claim') AND name = N'Segment' AND is_nullable = 1)
        EXEC sys.sp_executesql N'ALTER TABLE dbo.Draft_OTB_Approval_Claim ALTER COLUMN Segment nvarchar(20) NOT NULL;';
    IF EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID(N'dbo.Draft_OTB_Approval_Claim') AND name = N'Brand' AND is_nullable = 1)
        EXEC sys.sp_executesql N'ALTER TABLE dbo.Draft_OTB_Approval_Claim ALTER COLUMN Brand nvarchar(30) NOT NULL;';
    IF EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID(N'dbo.Draft_OTB_Approval_Claim') AND name = N'Vendor' AND is_nullable = 1)
        EXEC sys.sp_executesql N'ALTER TABLE dbo.Draft_OTB_Approval_Claim ALTER COLUMN Vendor nvarchar(30) NOT NULL;';

    IF EXISTS
    (
        SELECT 1
        FROM
        (
            VALUES
                (CAST(N'ClaimID' AS sysname), CAST(N'bigint' AS sysname), CAST(8 AS smallint), CAST(NULL AS tinyint), CAST(0 AS bit)),
                (N'JobID', N'uniqueidentifier', 16, NULL, 0),
                (N'RunNo', N'int', 4, NULL, 0),
                (N'BusinessKey', N'nvarchar', 1000, NULL, 0),
                (N'BusinessKeyHash', N'char', 64, NULL, 0),
                (N'Type', N'nvarchar', 40, NULL, 0),
                (N'Year', N'int', 4, NULL, 0),
                (N'Month', N'int', 4, NULL, 0),
                (N'Category', N'nvarchar', 40, NULL, 0),
                (N'Company', N'nvarchar', 40, NULL, 0),
                (N'Segment', N'nvarchar', 40, NULL, 0),
                (N'Brand', N'nvarchar', 60, NULL, 0),
                (N'Vendor', N'nvarchar', 60, NULL, 0),
                (N'ClaimedAt', N'datetime2', 7, 3, 0)
        ) expected(ColumnName, TypeName, MaxLength, ScaleValue, IsNullable)
        LEFT JOIN sys.columns c
               ON c.object_id = OBJECT_ID(N'dbo.Draft_OTB_Approval_Claim')
              AND c.name = expected.ColumnName
        LEFT JOIN sys.types t ON t.user_type_id = c.user_type_id
        WHERE c.column_id IS NULL
           OR t.name <> expected.TypeName
           OR c.max_length <> expected.MaxLength
           OR (expected.ScaleValue IS NOT NULL AND c.scale <> expected.ScaleValue)
           OR c.is_nullable <> expected.IsNullable
    )
       OR NOT EXISTS
          (
              SELECT 1
              FROM sys.identity_columns
              WHERE object_id = OBJECT_ID(N'dbo.Draft_OTB_Approval_Claim')
                AND name = N'ClaimID'
                AND is_identity = 1
                AND is_not_for_replication = 0
                AND CONVERT(bigint, seed_value) = 1
                AND CONVERT(bigint, increment_value) = 1
          )
        THROW 51121, 'Deployment failed: approval-claim column types, lengths, identity, or nullability do not match Sprint 1.', 1;

    IF NOT EXISTS
       (
           SELECT 1
           FROM sys.key_constraints kc
           INNER JOIN sys.indexes i
                   ON i.object_id = kc.parent_object_id
                  AND i.index_id = kc.unique_index_id
           WHERE kc.parent_object_id = OBJECT_ID(N'dbo.Draft_OTB_Approval_Claim')
             AND kc.name = N'PK_Draft_OTB_Approval_Claim'
             AND kc.[type] = N'PK'
             AND i.[type] = 1 AND i.is_unique = 1 AND i.has_filter = 0
             AND i.ignore_dup_key = 0 AND i.is_disabled = 0 AND i.is_hypothetical = 0
             AND (SELECT COUNT(*) FROM sys.index_columns ic
                  WHERE ic.object_id = i.object_id AND ic.index_id = i.index_id AND ic.key_ordinal > 0) = 1
             AND EXISTS
                 (
                     SELECT 1
                     FROM sys.index_columns ic
                     INNER JOIN sys.columns c ON c.object_id = ic.object_id AND c.column_id = ic.column_id
                     WHERE ic.object_id = i.object_id AND ic.index_id = i.index_id
                       AND ic.key_ordinal = 1 AND ic.is_descending_key = 0 AND c.name = N'ClaimID'
                 )
       )
       OR NOT EXISTS
       (
           SELECT 1
           FROM sys.foreign_keys fk
           WHERE fk.parent_object_id = OBJECT_ID(N'dbo.Draft_OTB_Approval_Claim')
             AND fk.referenced_object_id = OBJECT_ID(N'dbo.Draft_OTB_Background_Job')
             AND fk.name = N'FK_Draft_OTB_Approval_Claim_Job'
             AND fk.delete_referential_action = 1
             AND fk.update_referential_action = 0
             AND fk.is_not_for_replication = 0
             AND fk.is_disabled = 0
             AND fk.is_not_trusted = 0
             AND (SELECT COUNT(*) FROM sys.foreign_key_columns fkc
                  WHERE fkc.constraint_object_id = fk.object_id) = 1
             AND EXISTS
                 (
                     SELECT 1
                     FROM sys.foreign_key_columns fkc
                     INNER JOIN sys.columns pc
                             ON pc.object_id = fkc.parent_object_id
                            AND pc.column_id = fkc.parent_column_id
                     INNER JOIN sys.columns rc
                             ON rc.object_id = fkc.referenced_object_id
                            AND rc.column_id = fkc.referenced_column_id
                     WHERE fkc.constraint_object_id = fk.object_id
                       AND fkc.constraint_column_id = 1
                       AND pc.name = N'JobID'
                       AND rc.name = N'JobID'
                 )
       )
       OR NOT EXISTS
       (
           SELECT 1
           FROM sys.key_constraints kc
           INNER JOIN sys.indexes i ON i.object_id = kc.parent_object_id AND i.index_id = kc.unique_index_id
           WHERE kc.parent_object_id = OBJECT_ID(N'dbo.Draft_OTB_Approval_Claim')
             AND kc.name = N'UQ_Draft_OTB_Approval_Claim_RunNo'
             AND kc.[type] = N'UQ'
             AND i.[type] = 2 AND i.is_unique = 1 AND i.has_filter = 0
             AND i.ignore_dup_key = 0 AND i.is_disabled = 0 AND i.is_hypothetical = 0
             AND (SELECT COUNT(*) FROM sys.index_columns ic
                  WHERE ic.object_id = i.object_id AND ic.index_id = i.index_id AND ic.key_ordinal > 0) = 1
             AND EXISTS
                 (
                     SELECT 1
                     FROM sys.index_columns ic
                     INNER JOIN sys.columns c ON c.object_id = ic.object_id AND c.column_id = ic.column_id
                     WHERE ic.object_id = i.object_id AND ic.index_id = i.index_id
                       AND ic.key_ordinal = 1 AND ic.is_descending_key = 0 AND c.name = N'RunNo'
                 )
       )
       OR NOT EXISTS
       (
           SELECT 1
           FROM sys.key_constraints kc
           INNER JOIN sys.indexes i ON i.object_id = kc.parent_object_id AND i.index_id = kc.unique_index_id
           WHERE kc.parent_object_id = OBJECT_ID(N'dbo.Draft_OTB_Approval_Claim')
             AND kc.name = N'UQ_Draft_OTB_Approval_Claim_BusinessKeyHash'
             AND kc.[type] = N'UQ'
             AND i.[type] = 2 AND i.is_unique = 1 AND i.has_filter = 0
             AND i.ignore_dup_key = 0 AND i.is_disabled = 0 AND i.is_hypothetical = 0
             AND (SELECT COUNT(*) FROM sys.index_columns ic
                  WHERE ic.object_id = i.object_id AND ic.index_id = i.index_id AND ic.key_ordinal > 0) = 1
             AND EXISTS
                 (
                     SELECT 1
                     FROM sys.index_columns ic
                     INNER JOIN sys.columns c ON c.object_id = ic.object_id AND c.column_id = ic.column_id
                     WHERE ic.object_id = i.object_id AND ic.index_id = i.index_id
                       AND ic.key_ordinal = 1 AND ic.is_descending_key = 0 AND c.name = N'BusinessKeyHash'
                 )
       )
        THROW 51114, 'Deployment failed: required approval-claim constraints are missing.', 1;

    IF NOT EXISTS
    (
        SELECT 1 FROM sys.indexes
        WHERE object_id = OBJECT_ID(N'dbo.Draft_OTB_Approval_Claim')
          AND name = N'IX_Draft_OTB_Approval_Claim_Job'
    )
        EXEC sys.sp_executesql N'
            CREATE NONCLUSTERED INDEX IX_Draft_OTB_Approval_Claim_Job
                ON dbo.Draft_OTB_Approval_Claim (JobID)
                INCLUDE (RunNo, BusinessKeyHash, ClaimedAt);';

    IF NOT EXISTS
    (
        SELECT 1 FROM sys.indexes
        WHERE object_id = OBJECT_ID(N'dbo.Draft_OTB_Approval_Claim')
          AND name = N'IX_Draft_OTB_Approval_Claim_Group'
    )
        EXEC sys.sp_executesql N'
            CREATE NONCLUSTERED INDEX IX_Draft_OTB_Approval_Claim_Group
                ON dbo.Draft_OTB_Approval_Claim (Company, [Year], [Month], Category)
                INCLUDE (JobID, RunNo, [Type], Segment, Brand, Vendor, ClaimedAt);';

    IF NOT EXISTS
    (
        SELECT 1 FROM sys.indexes
        WHERE object_id = OBJECT_ID(N'dbo.Draft_OTB_Background_Job')
          AND name = N'UX_Draft_OTB_Job_Idempotency'
    )
        EXEC sys.sp_executesql N'
            CREATE UNIQUE NONCLUSTERED INDEX UX_Draft_OTB_Job_Idempotency
                ON dbo.Draft_OTB_Background_Job (JobType, RequestedBy, ClientRequestID);';

    IF NOT EXISTS
    (
        SELECT 1 FROM sys.indexes
        WHERE object_id = OBJECT_ID(N'dbo.Draft_OTB_Background_Job')
          AND name = N'IX_Draft_OTB_Job_User_Status'
    )
        EXEC sys.sp_executesql N'
            CREATE NONCLUSTERED INDEX IX_Draft_OTB_Job_User_Status
                ON dbo.Draft_OTB_Background_Job (RequestedBy, JobType, [Status], CreatedAt DESC)
                INCLUDE (Stage, ProgressPercent, TotalRows, ProcessedRows, SuccessRows, ErrorRows, UpdatedAt, AcknowledgedAt);';

    /* Lock and recheck the business table inside the DDL transaction so data
       cannot change between preflight and application of the support index. */
    EXEC sys.sp_executesql N'
        IF EXISTS
        (
            SELECT 1
            FROM dbo.Template_Upload_Draft_OTB WITH (TABLOCKX, HOLDLOCK)
            WHERE (OTBStatus IS NULL OR CONVERT(nvarchar(30), OTBStatus) = N''Draft'')
              AND
              (
                  [Type] IS NULL OR LEN(CONVERT(nvarchar(max), [Type])) = 0 OR LEN(CONVERT(nvarchar(max), [Type])) > 20
                  OR [Year] IS NULL OR [Year] <= 0
                  OR [Month] IS NULL OR [Month] NOT BETWEEN 1 AND 12
                  OR Category IS NULL OR LEN(CONVERT(nvarchar(max), Category)) = 0 OR LEN(CONVERT(nvarchar(max), Category)) > 20
                  OR Company IS NULL OR LEN(CONVERT(nvarchar(max), Company)) = 0 OR LEN(CONVERT(nvarchar(max), Company)) > 20
                  OR Segment IS NULL OR LEN(CONVERT(nvarchar(max), Segment)) = 0 OR LEN(CONVERT(nvarchar(max), Segment)) > 20
                  OR Brand IS NULL OR LEN(CONVERT(nvarchar(max), Brand)) = 0 OR LEN(CONVERT(nvarchar(max), Brand)) > 30
                  OR Vendor IS NULL OR LEN(CONVERT(nvarchar(max), Vendor)) = 0 OR LEN(CONVERT(nvarchar(max), Vendor)) > 30
              )
        )
            THROW 51115, ''Deployment failed: an active Draft OTB key became invalid during deployment.'', 1;

        IF EXISTS
        (
            SELECT 1
            FROM dbo.Template_Upload_Draft_OTB WITH (TABLOCKX, HOLDLOCK)
            WHERE OTBStatus IS NULL OR CONVERT(nvarchar(30), OTBStatus) = N''Draft''
            GROUP BY CONVERT(nvarchar(20), [Type]), [Year], [Month],
                     CONVERT(nvarchar(20), Company), CONVERT(nvarchar(20), Category),
                     CONVERT(nvarchar(20), Segment), CONVERT(nvarchar(30), Brand),
                     CONVERT(nvarchar(30), Vendor)
            HAVING COUNT_BIG(*) > 1
        )
            THROW 51115, ''Deployment failed: an active Draft OTB key became duplicated during deployment.'', 1;';

    IF NOT EXISTS
    (
        SELECT 1 FROM sys.indexes
        WHERE object_id = OBJECT_ID(N'dbo.Template_Upload_Draft_OTB')
          AND name = N'IX_Template_Upload_Draft_OTB_Search'
    )
        EXEC sys.sp_executesql N'
            CREATE NONCLUSTERED INDEX IX_Template_Upload_Draft_OTB_Search
                ON dbo.Template_Upload_Draft_OTB ([Year], [Month]);';

    IF NOT EXISTS
       (
           SELECT 1 FROM sys.indexes
           WHERE object_id = OBJECT_ID(N'dbo.Draft_OTB_Approval_Claim')
             AND name = N'IX_Draft_OTB_Approval_Claim_Job'
             AND [type] = 2 AND is_unique = 0 AND has_filter = 0 AND ignore_dup_key = 0
             AND is_disabled = 0 AND is_hypothetical = 0
       )
       OR (SELECT COUNT(*) FROM sys.index_columns
           WHERE object_id = OBJECT_ID(N'dbo.Draft_OTB_Approval_Claim')
             AND index_id = INDEXPROPERTY(OBJECT_ID(N'dbo.Draft_OTB_Approval_Claim'), N'IX_Draft_OTB_Approval_Claim_Job', 'IndexID')
             AND key_ordinal > 0) <> 1
       OR (SELECT COUNT(*) FROM sys.index_columns
           WHERE object_id = OBJECT_ID(N'dbo.Draft_OTB_Approval_Claim')
             AND index_id = INDEXPROPERTY(OBJECT_ID(N'dbo.Draft_OTB_Approval_Claim'), N'IX_Draft_OTB_Approval_Claim_Job', 'IndexID')
             AND is_included_column = 1) <> 3
       OR NOT EXISTS
          (
              SELECT 1
              FROM sys.index_columns ic
              INNER JOIN sys.columns c ON c.object_id = ic.object_id AND c.column_id = ic.column_id
              WHERE ic.object_id = OBJECT_ID(N'dbo.Draft_OTB_Approval_Claim')
                AND ic.index_id = INDEXPROPERTY(OBJECT_ID(N'dbo.Draft_OTB_Approval_Claim'), N'IX_Draft_OTB_Approval_Claim_Job', 'IndexID')
                AND ic.key_ordinal = 1 AND ic.is_descending_key = 0 AND c.name = N'JobID'
          )
       OR EXISTS
          (
              SELECT 1
              FROM (VALUES (N'RunNo'), (N'BusinessKeyHash'), (N'ClaimedAt')) required(ColumnName)
              WHERE NOT EXISTS
              (
                  SELECT 1
                  FROM sys.index_columns ic
                  INNER JOIN sys.columns c ON c.object_id = ic.object_id AND c.column_id = ic.column_id
                  WHERE ic.object_id = OBJECT_ID(N'dbo.Draft_OTB_Approval_Claim')
                    AND ic.index_id = INDEXPROPERTY(OBJECT_ID(N'dbo.Draft_OTB_Approval_Claim'), N'IX_Draft_OTB_Approval_Claim_Job', 'IndexID')
                    AND ic.is_included_column = 1
                    AND c.name = required.ColumnName
              )
          )
        THROW 51122, 'Deployment failed: IX_Draft_OTB_Approval_Claim_Job has an incompatible definition.', 1;

    /* The exact group-index shape is a correctness requirement because the
       application uses an explicit INDEX hint for durable claim locking. */
    IF NOT EXISTS
       (
           SELECT 1 FROM sys.indexes
           WHERE object_id = OBJECT_ID(N'dbo.Draft_OTB_Approval_Claim')
             AND name = N'IX_Draft_OTB_Approval_Claim_Group'
             AND [type] = 2 AND is_unique = 0 AND has_filter = 0 AND ignore_dup_key = 0
             AND is_disabled = 0 AND is_hypothetical = 0
       )
       OR (SELECT COUNT(*) FROM sys.index_columns
           WHERE object_id = OBJECT_ID(N'dbo.Draft_OTB_Approval_Claim')
             AND index_id = INDEXPROPERTY(OBJECT_ID(N'dbo.Draft_OTB_Approval_Claim'), N'IX_Draft_OTB_Approval_Claim_Group', 'IndexID')
             AND key_ordinal > 0) <> 4
       OR (SELECT COUNT(*) FROM sys.index_columns
           WHERE object_id = OBJECT_ID(N'dbo.Draft_OTB_Approval_Claim')
             AND index_id = INDEXPROPERTY(OBJECT_ID(N'dbo.Draft_OTB_Approval_Claim'), N'IX_Draft_OTB_Approval_Claim_Group', 'IndexID')
             AND is_included_column = 1) <> 7
       OR NOT EXISTS (SELECT 1 FROM sys.index_columns ic INNER JOIN sys.columns c ON c.object_id = ic.object_id AND c.column_id = ic.column_id
                       WHERE ic.object_id = OBJECT_ID(N'dbo.Draft_OTB_Approval_Claim') AND ic.index_id = INDEXPROPERTY(OBJECT_ID(N'dbo.Draft_OTB_Approval_Claim'), N'IX_Draft_OTB_Approval_Claim_Group', 'IndexID') AND ic.key_ordinal = 1 AND ic.is_descending_key = 0 AND c.name = N'Company')
       OR NOT EXISTS (SELECT 1 FROM sys.index_columns ic INNER JOIN sys.columns c ON c.object_id = ic.object_id AND c.column_id = ic.column_id
                       WHERE ic.object_id = OBJECT_ID(N'dbo.Draft_OTB_Approval_Claim') AND ic.index_id = INDEXPROPERTY(OBJECT_ID(N'dbo.Draft_OTB_Approval_Claim'), N'IX_Draft_OTB_Approval_Claim_Group', 'IndexID') AND ic.key_ordinal = 2 AND ic.is_descending_key = 0 AND c.name = N'Year')
       OR NOT EXISTS (SELECT 1 FROM sys.index_columns ic INNER JOIN sys.columns c ON c.object_id = ic.object_id AND c.column_id = ic.column_id
                       WHERE ic.object_id = OBJECT_ID(N'dbo.Draft_OTB_Approval_Claim') AND ic.index_id = INDEXPROPERTY(OBJECT_ID(N'dbo.Draft_OTB_Approval_Claim'), N'IX_Draft_OTB_Approval_Claim_Group', 'IndexID') AND ic.key_ordinal = 3 AND ic.is_descending_key = 0 AND c.name = N'Month')
       OR NOT EXISTS (SELECT 1 FROM sys.index_columns ic INNER JOIN sys.columns c ON c.object_id = ic.object_id AND c.column_id = ic.column_id
                       WHERE ic.object_id = OBJECT_ID(N'dbo.Draft_OTB_Approval_Claim') AND ic.index_id = INDEXPROPERTY(OBJECT_ID(N'dbo.Draft_OTB_Approval_Claim'), N'IX_Draft_OTB_Approval_Claim_Group', 'IndexID') AND ic.key_ordinal = 4 AND ic.is_descending_key = 0 AND c.name = N'Category')
       OR EXISTS
          (
              SELECT 1
              FROM (VALUES (N'JobID'), (N'RunNo'), (N'Type'), (N'Segment'), (N'Brand'), (N'Vendor'), (N'ClaimedAt')) required(ColumnName)
              WHERE NOT EXISTS
              (
                  SELECT 1
                  FROM sys.index_columns ic
                  INNER JOIN sys.columns c ON c.object_id = ic.object_id AND c.column_id = ic.column_id
                  WHERE ic.object_id = OBJECT_ID(N'dbo.Draft_OTB_Approval_Claim')
                    AND ic.index_id = INDEXPROPERTY(OBJECT_ID(N'dbo.Draft_OTB_Approval_Claim'), N'IX_Draft_OTB_Approval_Claim_Group', 'IndexID')
                    AND ic.is_included_column = 1
                    AND c.name = required.ColumnName
              )
          )
        THROW 51116, 'Deployment failed: IX_Draft_OTB_Approval_Claim_Group has an incompatible definition.', 1;

    IF NOT EXISTS
    (
        SELECT 1 FROM sys.indexes
        WHERE object_id = OBJECT_ID(N'dbo.Draft_OTB_Background_Job')
          AND name = N'UX_Draft_OTB_Job_Idempotency'
          AND [type] = 2 AND is_unique = 1 AND has_filter = 0 AND ignore_dup_key = 0
          AND is_disabled = 0 AND is_hypothetical = 0
    )
       OR (SELECT COUNT(*) FROM sys.index_columns
           WHERE object_id = OBJECT_ID(N'dbo.Draft_OTB_Background_Job')
             AND index_id = INDEXPROPERTY(OBJECT_ID(N'dbo.Draft_OTB_Background_Job'), N'UX_Draft_OTB_Job_Idempotency', 'IndexID')
             AND key_ordinal > 0) <> 3
       OR EXISTS
          (
              SELECT 1 FROM sys.index_columns
              WHERE object_id = OBJECT_ID(N'dbo.Draft_OTB_Background_Job')
                AND index_id = INDEXPROPERTY(OBJECT_ID(N'dbo.Draft_OTB_Background_Job'), N'UX_Draft_OTB_Job_Idempotency', 'IndexID')
                AND is_included_column = 1
          )
       OR NOT EXISTS (SELECT 1 FROM sys.index_columns ic INNER JOIN sys.columns c ON c.object_id = ic.object_id AND c.column_id = ic.column_id
                       WHERE ic.object_id = OBJECT_ID(N'dbo.Draft_OTB_Background_Job') AND ic.index_id = INDEXPROPERTY(OBJECT_ID(N'dbo.Draft_OTB_Background_Job'), N'UX_Draft_OTB_Job_Idempotency', 'IndexID') AND ic.key_ordinal = 1 AND ic.is_descending_key = 0 AND c.name = N'JobType')
       OR NOT EXISTS (SELECT 1 FROM sys.index_columns ic INNER JOIN sys.columns c ON c.object_id = ic.object_id AND c.column_id = ic.column_id
                       WHERE ic.object_id = OBJECT_ID(N'dbo.Draft_OTB_Background_Job') AND ic.index_id = INDEXPROPERTY(OBJECT_ID(N'dbo.Draft_OTB_Background_Job'), N'UX_Draft_OTB_Job_Idempotency', 'IndexID') AND ic.key_ordinal = 2 AND ic.is_descending_key = 0 AND c.name = N'RequestedBy')
       OR NOT EXISTS (SELECT 1 FROM sys.index_columns ic INNER JOIN sys.columns c ON c.object_id = ic.object_id AND c.column_id = ic.column_id
                       WHERE ic.object_id = OBJECT_ID(N'dbo.Draft_OTB_Background_Job') AND ic.index_id = INDEXPROPERTY(OBJECT_ID(N'dbo.Draft_OTB_Background_Job'), N'UX_Draft_OTB_Job_Idempotency', 'IndexID') AND ic.key_ordinal = 3 AND ic.is_descending_key = 0 AND c.name = N'ClientRequestID')
        THROW 51117, 'Deployment failed: the background-job idempotency index is missing or invalid.', 1;

    IF NOT EXISTS
       (
           SELECT 1 FROM sys.indexes
           WHERE object_id = OBJECT_ID(N'dbo.Draft_OTB_Background_Job')
             AND name = N'IX_Draft_OTB_Job_User_Status'
             AND [type] = 2 AND is_unique = 0 AND has_filter = 0 AND ignore_dup_key = 0
             AND is_disabled = 0 AND is_hypothetical = 0
       )
       OR (SELECT COUNT(*) FROM sys.index_columns
           WHERE object_id = OBJECT_ID(N'dbo.Draft_OTB_Background_Job')
             AND index_id = INDEXPROPERTY(OBJECT_ID(N'dbo.Draft_OTB_Background_Job'), N'IX_Draft_OTB_Job_User_Status', 'IndexID')
             AND key_ordinal > 0) <> 4
       OR (SELECT COUNT(*) FROM sys.index_columns
           WHERE object_id = OBJECT_ID(N'dbo.Draft_OTB_Background_Job')
             AND index_id = INDEXPROPERTY(OBJECT_ID(N'dbo.Draft_OTB_Background_Job'), N'IX_Draft_OTB_Job_User_Status', 'IndexID')
             AND is_included_column = 1) <> 8
       OR NOT EXISTS (SELECT 1 FROM sys.index_columns ic INNER JOIN sys.columns c ON c.object_id = ic.object_id AND c.column_id = ic.column_id
                      WHERE ic.object_id = OBJECT_ID(N'dbo.Draft_OTB_Background_Job') AND ic.index_id = INDEXPROPERTY(OBJECT_ID(N'dbo.Draft_OTB_Background_Job'), N'IX_Draft_OTB_Job_User_Status', 'IndexID') AND ic.key_ordinal = 1 AND ic.is_descending_key = 0 AND c.name = N'RequestedBy')
       OR NOT EXISTS (SELECT 1 FROM sys.index_columns ic INNER JOIN sys.columns c ON c.object_id = ic.object_id AND c.column_id = ic.column_id
                      WHERE ic.object_id = OBJECT_ID(N'dbo.Draft_OTB_Background_Job') AND ic.index_id = INDEXPROPERTY(OBJECT_ID(N'dbo.Draft_OTB_Background_Job'), N'IX_Draft_OTB_Job_User_Status', 'IndexID') AND ic.key_ordinal = 2 AND ic.is_descending_key = 0 AND c.name = N'JobType')
       OR NOT EXISTS (SELECT 1 FROM sys.index_columns ic INNER JOIN sys.columns c ON c.object_id = ic.object_id AND c.column_id = ic.column_id
                      WHERE ic.object_id = OBJECT_ID(N'dbo.Draft_OTB_Background_Job') AND ic.index_id = INDEXPROPERTY(OBJECT_ID(N'dbo.Draft_OTB_Background_Job'), N'IX_Draft_OTB_Job_User_Status', 'IndexID') AND ic.key_ordinal = 3 AND ic.is_descending_key = 0 AND c.name = N'Status')
       OR NOT EXISTS (SELECT 1 FROM sys.index_columns ic INNER JOIN sys.columns c ON c.object_id = ic.object_id AND c.column_id = ic.column_id
                      WHERE ic.object_id = OBJECT_ID(N'dbo.Draft_OTB_Background_Job') AND ic.index_id = INDEXPROPERTY(OBJECT_ID(N'dbo.Draft_OTB_Background_Job'), N'IX_Draft_OTB_Job_User_Status', 'IndexID') AND ic.key_ordinal = 4 AND ic.is_descending_key = 1 AND c.name = N'CreatedAt')
       OR EXISTS
          (
              SELECT 1
              FROM (VALUES (N'Stage'), (N'ProgressPercent'), (N'TotalRows'), (N'ProcessedRows'),
                           (N'SuccessRows'), (N'ErrorRows'), (N'UpdatedAt'), (N'AcknowledgedAt')) required(ColumnName)
              WHERE NOT EXISTS
              (
                  SELECT 1
                  FROM sys.index_columns ic
                  INNER JOIN sys.columns c ON c.object_id = ic.object_id AND c.column_id = ic.column_id
                  WHERE ic.object_id = OBJECT_ID(N'dbo.Draft_OTB_Background_Job')
                    AND ic.index_id = INDEXPROPERTY(OBJECT_ID(N'dbo.Draft_OTB_Background_Job'), N'IX_Draft_OTB_Job_User_Status', 'IndexID')
                    AND ic.is_included_column = 1
                    AND c.name = required.ColumnName
              )
          )
        THROW 51123, 'Deployment failed: IX_Draft_OTB_Job_User_Status has an incompatible definition.', 1;

    IF NOT EXISTS
    (
        SELECT 1 FROM sys.indexes
        WHERE object_id = OBJECT_ID(N'dbo.Template_Upload_Draft_OTB')
          AND name = N'IX_Template_Upload_Draft_OTB_Search'
          AND [type] = 2 AND is_unique = 0 AND has_filter = 0 AND ignore_dup_key = 0
          AND is_disabled = 0 AND is_hypothetical = 0
    )
       OR (SELECT COUNT(*) FROM sys.index_columns
           WHERE object_id = OBJECT_ID(N'dbo.Template_Upload_Draft_OTB')
             AND index_id = INDEXPROPERTY(OBJECT_ID(N'dbo.Template_Upload_Draft_OTB'), N'IX_Template_Upload_Draft_OTB_Search', 'IndexID')
             AND key_ordinal > 0) <> 2
       OR EXISTS
          (
              SELECT 1 FROM sys.index_columns
              WHERE object_id = OBJECT_ID(N'dbo.Template_Upload_Draft_OTB')
                AND index_id = INDEXPROPERTY(OBJECT_ID(N'dbo.Template_Upload_Draft_OTB'), N'IX_Template_Upload_Draft_OTB_Search', 'IndexID')
                AND is_included_column = 1
          )
       OR NOT EXISTS (SELECT 1 FROM sys.index_columns ic INNER JOIN sys.columns c ON c.object_id = ic.object_id AND c.column_id = ic.column_id
                      WHERE ic.object_id = OBJECT_ID(N'dbo.Template_Upload_Draft_OTB') AND ic.index_id = INDEXPROPERTY(OBJECT_ID(N'dbo.Template_Upload_Draft_OTB'), N'IX_Template_Upload_Draft_OTB_Search', 'IndexID') AND ic.key_ordinal = 1 AND ic.is_descending_key = 0 AND c.name = N'Year')
       OR NOT EXISTS (SELECT 1 FROM sys.index_columns ic INNER JOIN sys.columns c ON c.object_id = ic.object_id AND c.column_id = ic.column_id
                      WHERE ic.object_id = OBJECT_ID(N'dbo.Template_Upload_Draft_OTB') AND ic.index_id = INDEXPROPERTY(OBJECT_ID(N'dbo.Template_Upload_Draft_OTB'), N'IX_Template_Upload_Draft_OTB_Search', 'IndexID') AND ic.key_ordinal = 2 AND ic.is_descending_key = 0 AND c.name = N'Month')
        THROW 51118, 'Deployment failed: IX_Template_Upload_Draft_OTB_Search has an incompatible definition.', 1;

    COMMIT TRANSACTION;
END TRY
BEGIN CATCH
    IF XACT_STATE() <> 0 ROLLBACK TRANSACTION;
    THROW;
END CATCH;

SELECT
    N'PASS' AS DeploymentStatus,
    DB_NAME() AS DatabaseName,
    CONVERT(nvarchar(128), CONNECTIONPROPERTY('local_net_address')) AS LocalSqlAddress,
    (SELECT COUNT_BIG(*) FROM dbo.Draft_OTB_Background_Job) AS BackgroundJobRowsPreserved,
    (SELECT COUNT_BIG(*) FROM dbo.Draft_OTB_Approval_Claim) AS ApprovalClaimRows,
    N'No existing business-data values were changed.' AS DataImpact;
