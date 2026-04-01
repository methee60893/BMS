USE [BMS_TEST];
GO

SET NOCOUNT ON;
GO

DECLARE @Expected TABLE
(
    TableName sysname NOT NULL,
    ExpectedCount int NOT NULL
);

INSERT INTO @Expected (TableName, ExpectedCount)
VALUES
    (N'MS_Vendor', 333),
    (N'MS_Brand', 626),
    (N'MS_Category', 44),
    (N'MS_User', 97),
    (N'MS_Role', 3),
    (N'Map_User_Role', 95),
    (N'Map_Role_Permission', 30);

WITH ActualCounts AS
(
    SELECT N'MS_Vendor' AS TableName, COUNT(1) AS ActualCount FROM dbo.MS_Vendor
    UNION ALL SELECT N'MS_Brand', COUNT(1) FROM dbo.MS_Brand
    UNION ALL SELECT N'MS_Category', COUNT(1) FROM dbo.MS_Category
    UNION ALL SELECT N'MS_User', COUNT(1) FROM dbo.MS_User
    UNION ALL SELECT N'MS_Role', COUNT(1) FROM dbo.MS_Role
    UNION ALL SELECT N'Map_User_Role', COUNT(1) FROM dbo.Map_User_Role
    UNION ALL SELECT N'Map_Role_Permission', COUNT(1) FROM dbo.Map_Role_Permission
)
SELECT
    e.TableName,
    e.ExpectedCount,
    a.ActualCount,
    CASE WHEN e.ExpectedCount = a.ActualCount THEN N'PASS' ELSE N'FAIL' END AS VerifyStatus
FROM @Expected e
LEFT JOIN ActualCounts a
    ON e.TableName = a.TableName
ORDER BY e.TableName;

SELECT
    CASE
        WHEN EXISTS
        (
            SELECT 1
            FROM @Expected e
            LEFT JOIN ActualCounts a
                ON e.TableName = a.TableName
            WHERE ISNULL(a.ActualCount, -1) <> e.ExpectedCount
        )
        THEN N'FAIL'
        ELSE N'PASS'
    END AS OverallStatus;
GO
