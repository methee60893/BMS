/*
    Sprint 1 #65/#68 - durable Draft OTB upload/approval jobs.

    Run this script on the BMS test database before deploying the application.
    The table contains operational job state only. Invalid upload/approval jobs
    do not write to Template_Upload_Draft_OTB, OTB_Transaction, or SAP.
*/
SET XACT_ABORT ON;
GO

IF OBJECT_ID(N'dbo.Draft_OTB_Background_Job', N'U') IS NULL
BEGIN
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
    );
END;
GO

/* Idempotent upgrade for server-backed review acknowledgement. */
IF COL_LENGTH(N'dbo.Draft_OTB_Background_Job', N'AcknowledgedAt') IS NULL
    ALTER TABLE dbo.Draft_OTB_Background_Job ADD AcknowledgedAt datetime2(3) NULL;
IF COL_LENGTH(N'dbo.Draft_OTB_Background_Job', N'AcknowledgedBy') IS NULL
    ALTER TABLE dbo.Draft_OTB_Background_Job ADD AcknowledgedBy nvarchar(100) NULL;
GO

IF OBJECT_ID(N'dbo.Draft_OTB_Approval_Claim', N'U') IS NULL
BEGIN
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
    );
END;
GO

/* Idempotent upgrade for environments where the claim table predates
   dimension-level mutation interlocks. */
IF COL_LENGTH(N'dbo.Draft_OTB_Approval_Claim', N'Type') IS NULL
    ALTER TABLE dbo.Draft_OTB_Approval_Claim ADD [Type] nvarchar(20) NULL;
IF COL_LENGTH(N'dbo.Draft_OTB_Approval_Claim', N'Year') IS NULL
    ALTER TABLE dbo.Draft_OTB_Approval_Claim ADD [Year] int NULL;
IF COL_LENGTH(N'dbo.Draft_OTB_Approval_Claim', N'Month') IS NULL
    ALTER TABLE dbo.Draft_OTB_Approval_Claim ADD [Month] int NULL;
IF COL_LENGTH(N'dbo.Draft_OTB_Approval_Claim', N'Category') IS NULL
    ALTER TABLE dbo.Draft_OTB_Approval_Claim ADD Category nvarchar(20) NULL;
IF COL_LENGTH(N'dbo.Draft_OTB_Approval_Claim', N'Company') IS NULL
    ALTER TABLE dbo.Draft_OTB_Approval_Claim ADD Company nvarchar(20) NULL;
IF COL_LENGTH(N'dbo.Draft_OTB_Approval_Claim', N'Segment') IS NULL
    ALTER TABLE dbo.Draft_OTB_Approval_Claim ADD Segment nvarchar(20) NULL;
IF COL_LENGTH(N'dbo.Draft_OTB_Approval_Claim', N'Brand') IS NULL
    ALTER TABLE dbo.Draft_OTB_Approval_Claim ADD Brand nvarchar(30) NULL;
IF COL_LENGTH(N'dbo.Draft_OTB_Approval_Claim', N'Vendor') IS NULL
    ALTER TABLE dbo.Draft_OTB_Approval_Claim ADD Vendor nvarchar(30) NULL;
GO

UPDATE c
SET c.[Type] = d.[Type], c.[Year] = d.[Year], c.[Month] = d.[Month],
    c.Category = d.Category, c.Company = d.Company, c.Segment = d.Segment,
    c.Brand = d.Brand, c.Vendor = d.Vendor
FROM dbo.Draft_OTB_Approval_Claim c
INNER JOIN dbo.Template_Upload_Draft_OTB d ON d.RunNo = c.RunNo
WHERE c.[Type] IS NULL OR c.[Year] IS NULL OR c.[Month] IS NULL OR c.Category IS NULL
   OR c.Company IS NULL OR c.Segment IS NULL OR c.Brand IS NULL OR c.Vendor IS NULL;

IF EXISTS
(
    SELECT 1 FROM dbo.Draft_OTB_Approval_Claim
    WHERE [Type] IS NULL OR [Year] IS NULL OR [Month] IS NULL OR Category IS NULL
       OR Company IS NULL OR Segment IS NULL OR Brand IS NULL OR Vendor IS NULL
)
BEGIN
    ;THROW 51020, 'Existing approval claims could not be backfilled with budget dimensions.', 1;
END;
GO

IF EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID(N'dbo.Draft_OTB_Approval_Claim') AND name = N'Type' AND is_nullable = 1)
    ALTER TABLE dbo.Draft_OTB_Approval_Claim ALTER COLUMN [Type] nvarchar(20) NOT NULL;
IF EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID(N'dbo.Draft_OTB_Approval_Claim') AND name = N'Year' AND is_nullable = 1)
    ALTER TABLE dbo.Draft_OTB_Approval_Claim ALTER COLUMN [Year] int NOT NULL;
IF EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID(N'dbo.Draft_OTB_Approval_Claim') AND name = N'Month' AND is_nullable = 1)
    ALTER TABLE dbo.Draft_OTB_Approval_Claim ALTER COLUMN [Month] int NOT NULL;
IF EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID(N'dbo.Draft_OTB_Approval_Claim') AND name = N'Category' AND is_nullable = 1)
    ALTER TABLE dbo.Draft_OTB_Approval_Claim ALTER COLUMN Category nvarchar(20) NOT NULL;
IF EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID(N'dbo.Draft_OTB_Approval_Claim') AND name = N'Company' AND is_nullable = 1)
    ALTER TABLE dbo.Draft_OTB_Approval_Claim ALTER COLUMN Company nvarchar(20) NOT NULL;
IF EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID(N'dbo.Draft_OTB_Approval_Claim') AND name = N'Segment' AND is_nullable = 1)
    ALTER TABLE dbo.Draft_OTB_Approval_Claim ALTER COLUMN Segment nvarchar(20) NOT NULL;
IF EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID(N'dbo.Draft_OTB_Approval_Claim') AND name = N'Brand' AND is_nullable = 1)
    ALTER TABLE dbo.Draft_OTB_Approval_Claim ALTER COLUMN Brand nvarchar(30) NOT NULL;
IF EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID(N'dbo.Draft_OTB_Approval_Claim') AND name = N'Vendor' AND is_nullable = 1)
    ALTER TABLE dbo.Draft_OTB_Approval_Claim ALTER COLUMN Vendor nvarchar(30) NOT NULL;
GO

IF NOT EXISTS
(
    SELECT 1 FROM sys.indexes
    WHERE object_id = OBJECT_ID(N'dbo.Draft_OTB_Approval_Claim')
      AND name = N'IX_Draft_OTB_Approval_Claim_Job'
)
BEGIN
    CREATE NONCLUSTERED INDEX IX_Draft_OTB_Approval_Claim_Job
        ON dbo.Draft_OTB_Approval_Claim (JobID)
        INCLUDE (RunNo, BusinessKeyHash, ClaimedAt);
END;
GO

IF NOT EXISTS
(
    SELECT 1 FROM sys.indexes
    WHERE object_id = OBJECT_ID(N'dbo.Draft_OTB_Approval_Claim')
      AND name = N'IX_Draft_OTB_Approval_Claim_Group'
)
BEGIN
    CREATE NONCLUSTERED INDEX IX_Draft_OTB_Approval_Claim_Group
        ON dbo.Draft_OTB_Approval_Claim (Company, [Year], [Month], Category)
        INCLUDE (JobID, RunNo, [Type], Segment, Brand, Vendor, ClaimedAt);
END;
GO

IF NOT EXISTS
(
    SELECT 1
    FROM sys.indexes
    WHERE object_id = OBJECT_ID(N'dbo.Draft_OTB_Background_Job')
      AND name = N'UX_Draft_OTB_Job_Idempotency'
)
BEGIN
    CREATE UNIQUE NONCLUSTERED INDEX UX_Draft_OTB_Job_Idempotency
        ON dbo.Draft_OTB_Background_Job (JobType, RequestedBy, ClientRequestID);
END;
GO

/* Refuse deployment when legacy data already violates the one-active-Draft
   invariant. Upload persistence takes an exclusive table lock because some
   deployed databases retain nvarchar(max) dimension columns that cannot be
   index keys. This index is therefore a compatible candidate-row accelerator,
   not the concurrency boundary. */
IF NOT EXISTS
(
    SELECT 1 FROM sys.indexes
    WHERE object_id = OBJECT_ID(N'dbo.Template_Upload_Draft_OTB')
      AND name = N'IX_Template_Upload_Draft_OTB_Search'
)
BEGIN
    IF EXISTS
    (
        SELECT 1
        FROM dbo.Template_Upload_Draft_OTB
        WHERE OTBStatus IS NULL OR OTBStatus = N'Draft'
        GROUP BY [Type], [Year], [Month], Company, Category, Segment, Brand, Vendor
        HAVING COUNT_BIG(*) > 1
    )
    BEGIN
        ;THROW 51021, 'Duplicate active Draft OTB business keys must be resolved before creating the upload lock index.', 1;
    END;

    CREATE NONCLUSTERED INDEX IX_Template_Upload_Draft_OTB_Search
        ON dbo.Template_Upload_Draft_OTB ([Year], [Month]);
END;
GO

IF NOT EXISTS
(
    SELECT 1
    FROM sys.indexes
    WHERE object_id = OBJECT_ID(N'dbo.Draft_OTB_Background_Job')
      AND name = N'IX_Draft_OTB_Job_User_Status'
)
BEGIN
    CREATE NONCLUSTERED INDEX IX_Draft_OTB_Job_User_Status
        ON dbo.Draft_OTB_Background_Job (RequestedBy, JobType, [Status], CreatedAt DESC)
        INCLUDE (Stage, ProgressPercent, TotalRows, ProcessedRows, SuccessRows, ErrorRows, UpdatedAt, AcknowledgedAt);
END;
GO
