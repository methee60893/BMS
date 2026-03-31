USE [BMS];
GO

CREATE OR ALTER PROCEDURE dbo.SP_Get_Actual_PO_List
    @Year nvarchar(10) = NULL,
    @Month nvarchar(10) = NULL,
    @Company nvarchar(20) = NULL,
    @Category nvarchar(20) = NULL,
    @Segment nvarchar(20) = NULL,
    @Brand nvarchar(30) = NULL,
    @Vendor nvarchar(30) = NULL
AS
BEGIN
    SET NOCOUNT ON;

    SELECT
        a.ActualPO_ID,
        a.Actual_PO_Date,
        a.PO_No AS Actual_PO_No,
        N'Actual' AS PO_Type,
        a.OTB_Year AS PO_Year,
        m.month_name_sh AS PO_Month_Name,
        a.Category_Code,
        c.Category AS Category_Name,
        a.Company_Code,
        co.CompanyNameShort AS Company_Name,
        CASE
            WHEN LEN(ISNULL(a.Segment_Code, N'')) > 2 THEN SUBSTRING(a.Segment_Code, 2, LEN(a.Segment_Code) - 2)
            ELSE a.Segment_Code
        END AS Segment_Code,
        s.SegmentName AS Segment_Name,
        a.Brand_Code,
        b.[Brand Name] AS Brand_Name,
        a.Vendor_Code,
        v.Vendor AS Vendor_Name,
        ISNULL(a.Amount_THB, 0) AS Amount_THB,
        ISNULL(a.Amount_CCY, 0) AS Amount_CCY,
        a.CCY,
        ISNULL(a.Exchange_Rate, 0) AS Exchange_Rate,
        a.Draft_PO_Ref,
        a.[Status],
        a.Matching_Date AS Status_Date,
        a.Remark
    FROM dbo.Actual_PO_Summary a
    LEFT JOIN dbo.MS_Month m
        ON a.OTB_Month = m.month_code
    LEFT JOIN dbo.MS_Category c
        ON a.Category_Code = c.Cate
    LEFT JOIN dbo.MS_Company co
        ON a.Company_Code = co.CompanyCode
    LEFT JOIN dbo.MS_Segment s
        ON CASE
               WHEN LEN(ISNULL(a.Segment_Code, N'')) > 2 THEN SUBSTRING(a.Segment_Code, 2, LEN(a.Segment_Code) - 2)
               ELSE a.Segment_Code
           END = s.SegmentCode
    LEFT JOIN dbo.MS_Brand b
        ON a.Brand_Code = b.[Brand Code]
    LEFT JOIN dbo.MS_Vendor v
        ON a.Vendor_Code = v.VendorCode
       AND CASE
               WHEN LEN(ISNULL(a.Segment_Code, N'')) > 2 THEN SUBSTRING(a.Segment_Code, 2, LEN(a.Segment_Code) - 2)
               ELSE a.Segment_Code
           END = v.SegmentCode
    WHERE (@Year IS NULL OR CONVERT(nvarchar(10), a.OTB_Year) = @Year)
      AND (@Month IS NULL OR CONVERT(nvarchar(10), a.OTB_Month) = @Month)
      AND (@Company IS NULL OR a.Company_Code = @Company)
      AND (@Category IS NULL OR a.Category_Code = @Category)
      AND (@Segment IS NULL OR CASE
                                   WHEN LEN(ISNULL(a.Segment_Code, N'')) > 2 THEN SUBSTRING(a.Segment_Code, 2, LEN(a.Segment_Code) - 2)
                                   ELSE a.Segment_Code
                               END = @Segment)
      AND (@Brand IS NULL OR a.Brand_Code = @Brand)
      AND (@Vendor IS NULL OR a.Vendor_Code = @Vendor)
    ORDER BY a.Actual_PO_Date DESC, a.PO_No DESC;
END;
GO

CREATE OR ALTER PROCEDURE dbo.SP_Search_Approved_OTB
    @Type nvarchar(20) = NULL,
    @Year nvarchar(10) = NULL,
    @Month nvarchar(10) = NULL,
    @Company nvarchar(20) = NULL,
    @Category nvarchar(20) = NULL,
    @Segment nvarchar(20) = NULL,
    @Brand nvarchar(30) = NULL,
    @Vendor nvarchar(30) = NULL,
    @Status nvarchar(30) = NULL,
    @DateFrom datetime = NULL,
    @DateTo datetime = NULL,
    @Version nvarchar(20) = NULL
AS
BEGIN
    SET NOCOUNT ON;

    SELECT
        CreateDate,
        [Type],
        [Year],
        [Month],
        Category,
        CategoryName,
        Company,
        Segment,
        SegmentName,
        Brand,
        BrandName,
        Vendor,
        VendorName,
        Amount,
        RevisedDiff,
        Remark,
        OTBStatus,
        ApprovedDate,
        ActionBy,
        SAPStatus,
        SAPErrorMessage,
        [Version],
        DraftID
    FROM dbo.OTB_Transaction
    WHERE (@Type IS NULL OR [Type] = @Type)
      AND (@Year IS NULL OR CONVERT(nvarchar(10), [Year]) = @Year)
      AND (@Month IS NULL OR CONVERT(nvarchar(10), [Month]) = @Month)
      AND (@Company IS NULL OR Company = @Company)
      AND (@Category IS NULL OR Category = @Category)
      AND (@Segment IS NULL OR Segment = @Segment)
      AND (@Brand IS NULL OR Brand = @Brand)
      AND (@Vendor IS NULL OR Vendor = @Vendor)
      AND (@Status IS NULL OR OTBStatus = @Status)
      AND (@Version IS NULL OR [Version] = @Version)
      AND (@DateFrom IS NULL OR CreateDate >= @DateFrom)
      AND (@DateTo IS NULL OR CreateDate < DATEADD(day, 1, @DateTo))
    ORDER BY CreateDate DESC, [Year] DESC, [Month] DESC;
END;
GO

CREATE OR ALTER PROCEDURE dbo.SP_Search_SWitch_OTB
    @Type nvarchar(20) = NULL,
    @Year nvarchar(10) = NULL,
    @Month nvarchar(10) = NULL,
    @Company nvarchar(20) = NULL,
    @Category nvarchar(20) = NULL,
    @Segment nvarchar(20) = NULL,
    @Brand nvarchar(30) = NULL,
    @Vendor nvarchar(30) = NULL
AS
BEGIN
    SET NOCOUNT ON;

    SELECT
        s.CreateDT,
        s.[From] AS [Type],
        s.[Year],
        m1.month_name_sh AS MonthName,
        s.Category,
        c1.Category AS CategoryName,
        co1.CompanyNameShort AS CompanyName,
        s.Segment,
        seg1.SegmentName,
        s.Brand,
        b1.[Brand Name] AS BrandName,
        s.Vendor,
        v1.Vendor AS VendorName,
        s.[To] AS SwitchType,
        s.SwitchYear,
        m2.month_name_sh AS SwitchMonthName,
        s.SwitchCategory,
        co2.CompanyNameShort AS SwitchCompanyName,
        s.SwitchSegment,
        s.SwitchBrand,
        s.SwitchVendor,
        s.BudgetAmount,
        s.CreateBy,
        s.Remark,
        s.OTBStatus
    FROM dbo.OTB_Switching_Transaction s
    LEFT JOIN dbo.MS_Month m1 ON s.[Month] = m1.month_code
    LEFT JOIN dbo.MS_Month m2 ON s.SwitchMonth = m2.month_code
    LEFT JOIN dbo.MS_Category c1 ON s.Category = c1.Cate
    LEFT JOIN dbo.MS_Company co1 ON s.Company = co1.CompanyCode
    LEFT JOIN dbo.MS_Company co2 ON s.SwitchCompany = co2.CompanyCode
    LEFT JOIN dbo.MS_Segment seg1 ON s.Segment = seg1.SegmentCode
    LEFT JOIN dbo.MS_Brand b1 ON s.Brand = b1.[Brand Code]
    LEFT JOIN dbo.MS_Vendor v1 ON s.Vendor = v1.VendorCode AND s.Segment = v1.SegmentCode
    WHERE (@Type IS NULL OR s.[From] = @Type OR s.[To] = @Type)
      AND (@Year IS NULL OR CONVERT(nvarchar(10), s.[Year]) = @Year OR CONVERT(nvarchar(10), s.SwitchYear) = @Year)
      AND (@Month IS NULL OR CONVERT(nvarchar(10), s.[Month]) = @Month OR CONVERT(nvarchar(10), s.SwitchMonth) = @Month)
      AND (@Company IS NULL OR s.Company = @Company OR s.SwitchCompany = @Company)
      AND (@Category IS NULL OR s.Category = @Category OR s.SwitchCategory = @Category)
      AND (@Segment IS NULL OR s.Segment = @Segment OR s.SwitchSegment = @Segment)
      AND (@Brand IS NULL OR s.Brand = @Brand OR s.SwitchBrand = @Brand)
      AND (@Vendor IS NULL OR s.Vendor = @Vendor OR s.SwitchVendor = @Vendor)
    ORDER BY s.CreateDT DESC, s.SwitchingID DESC;
END;
GO

CREATE OR ALTER PROCEDURE dbo.SP_Sync_Actual_PO_Summary
AS
BEGIN
    SET NOCOUNT ON;

    MERGE dbo.Actual_PO_Summary AS T
    USING
    (
        SELECT
            s.PO AS PO_No,
            s.PO_Item,
            s.Otb_Year AS OTB_Year,
            s.Otb_Month AS OTB_Month,
            s.Company_Code,
            s.Category AS Category_Code,
            s.Fund AS Segment_Code,
            s.Brand AS Brand_Code,
            s.Supplier AS Vendor_Code,
            s.Otb_Date AS Actual_PO_Date,
            s.PO_Local_Amount AS Amount_THB,
            s.PO_Amount AS Amount_CCY,
            COALESCE(s.PO_Currency, s.PO_Local_Currency) AS CCY,
            s.Exchange_Rate,
            s.Modified_Date AS Source_Modified_Date,
            s.BMS_Last_Synced,
            CASE
                WHEN ISNULL(s.Deletion_Flag, N'') = N'L' THEN N'Cancelled'
                ELSE NULL
            END AS SyncStatus
        FROM dbo.Actual_PO_Staging s
    ) AS S
        ON T.PO_No = S.PO_No
       AND T.PO_Item = S.PO_Item
    WHEN MATCHED THEN
        UPDATE SET
            T.OTB_Year = S.OTB_Year,
            T.OTB_Month = S.OTB_Month,
            T.Company_Code = S.Company_Code,
            T.Category_Code = S.Category_Code,
            T.Segment_Code = S.Segment_Code,
            T.Brand_Code = S.Brand_Code,
            T.Vendor_Code = S.Vendor_Code,
            T.Actual_PO_Date = S.Actual_PO_Date,
            T.Amount_THB = S.Amount_THB,
            T.Amount_CCY = S.Amount_CCY,
            T.CCY = S.CCY,
            T.Exchange_Rate = S.Exchange_Rate,
            T.Source_Modified_Date = S.Source_Modified_Date,
            T.BMS_Last_Synced = S.BMS_Last_Synced,
            T.[Status] = COALESCE(S.SyncStatus, T.[Status]),
            T.Changed_Date = sysdatetime()
    WHEN NOT MATCHED BY TARGET THEN
        INSERT
        (
            PO_No, PO_Item, OTB_Year, OTB_Month, Company_Code, Category_Code, Segment_Code,
            Brand_Code, Vendor_Code, Actual_PO_Date, Amount_THB, Amount_CCY, CCY, Exchange_Rate,
            [Status], Source_Modified_Date, BMS_Last_Synced
        )
        VALUES
        (
            S.PO_No, S.PO_Item, S.OTB_Year, S.OTB_Month, S.Company_Code, S.Category_Code, S.Segment_Code,
            S.Brand_Code, S.Vendor_Code, S.Actual_PO_Date, S.Amount_THB, S.Amount_CCY, S.CCY, S.Exchange_Rate,
            S.SyncStatus, S.Source_Modified_Date, S.BMS_Last_Synced
        );
END;
GO

CREATE OR ALTER PROCEDURE dbo.SP_Auto_Match_Actual_Draft
    @UpdateBy nvarchar(100) = N'System AutoMatch'
AS
BEGIN
    SET NOCOUNT ON;

    ;WITH DraftPool AS
    (
        SELECT
            d.DraftPO_ID,
            d.DraftPO_No,
            d.PO_Year,
            d.PO_Month,
            d.Company_Code,
            d.Category_Code,
            d.Segment_Code,
            d.Brand_Code,
            d.Vendor_Code,
            d.Amount_THB,
            ROW_NUMBER() OVER
            (
                PARTITION BY d.PO_Year, d.PO_Month, d.Company_Code, d.Category_Code, d.Segment_Code, d.Brand_Code, d.Vendor_Code, d.Amount_THB
                ORDER BY d.Created_Date, d.DraftPO_ID
            ) AS rn
        FROM dbo.Draft_PO_Transaction d
        WHERE ISNULL(d.[Status], N'Draft') IN (N'Draft', N'Edited', N'Matching')
          AND d.Actual_PO_Ref IS NULL
    ),
    ActualPool AS
    (
        SELECT
            a.ActualPO_ID,
            a.PO_No,
            a.OTB_Year,
            a.OTB_Month,
            a.Company_Code,
            a.Category_Code,
            CASE
                WHEN LEN(ISNULL(a.Segment_Code, N'')) > 2 THEN SUBSTRING(a.Segment_Code, 2, LEN(a.Segment_Code) - 2)
                ELSE a.Segment_Code
            END AS Segment_Code_Normalized,
            a.Brand_Code,
            a.Vendor_Code,
            a.Amount_THB,
            ROW_NUMBER() OVER
            (
                PARTITION BY a.OTB_Year, a.OTB_Month, a.Company_Code, a.Category_Code,
                             CASE
                                 WHEN LEN(ISNULL(a.Segment_Code, N'')) > 2 THEN SUBSTRING(a.Segment_Code, 2, LEN(a.Segment_Code) - 2)
                                 ELSE a.Segment_Code
                             END,
                             a.Brand_Code, a.Vendor_Code, a.Amount_THB
                ORDER BY a.Actual_PO_Date, a.ActualPO_ID
            ) AS rn
        FROM dbo.Actual_PO_Summary a
        WHERE ISNULL(a.[Status], N'') NOT IN (N'Matched', N'Cancelled')
          AND a.Draft_PO_Ref IS NULL
    ),
    Pairs AS
    (
        SELECT
            d.DraftPO_ID,
            d.DraftPO_No,
            a.ActualPO_ID,
            a.PO_No
        FROM DraftPool d
        INNER JOIN ActualPool a
            ON d.PO_Year = a.OTB_Year
           AND d.PO_Month = a.OTB_Month
           AND d.Company_Code = a.Company_Code
           AND d.Category_Code = a.Category_Code
           AND d.Segment_Code = a.Segment_Code_Normalized
           AND d.Brand_Code = a.Brand_Code
           AND d.Vendor_Code = a.Vendor_Code
           AND d.Amount_THB = a.Amount_THB
           AND d.rn = a.rn
    )
    UPDATE d
        SET d.Actual_PO_Ref = p.PO_No,
            d.[Status] = N'Matching',
            d.Status_Date = sysdatetime(),
            d.Status_By = @UpdateBy
    FROM dbo.Draft_PO_Transaction d
    INNER JOIN Pairs p
        ON d.DraftPO_ID = p.DraftPO_ID;

    UPDATE a
        SET a.Draft_PO_Ref = p.DraftPO_No,
            a.[Status] = N'Matching',
            a.Matching_Date = sysdatetime(),
            a.Changed_By = @UpdateBy,
            a.Changed_Date = sysdatetime()
    FROM dbo.Actual_PO_Summary a
    INNER JOIN Pairs p
        ON a.ActualPO_ID = p.ActualPO_ID;
END;
GO

CREATE OR ALTER PROCEDURE dbo.SP_Approve_Draft_OTB
    @DraftIDs nvarchar(max),
    @ApprovedBy nvarchar(100),
    @Remark nvarchar(500) = NULL
AS
BEGIN
    SET NOCOUNT ON;

    DECLARE @Ids TABLE (RunNo int PRIMARY KEY);

    INSERT INTO @Ids (RunNo)
    SELECT DISTINCT TRY_CAST(value AS int)
    FROM string_split(@DraftIDs, N',')
    WHERE TRY_CAST(value AS int) IS NOT NULL;

    UPDATE d
        SET d.OTBStatus = N'Approved',
            d.UpdateBy = @ApprovedBy,
            d.UpdateDT = sysdatetime(),
            d.Remark = COALESCE(@Remark, d.Remark)
    FROM dbo.Template_Upload_Draft_OTB d
    INNER JOIN @Ids i
        ON d.RunNo = i.RunNo
    WHERE ISNULL(d.OTBStatus, N'Draft') = N'Draft';

    MERGE dbo.OTB_Transaction AS T
    USING
    (
        SELECT
            d.RunNo,
            d.[Type],
            d.[Year],
            d.[Month],
            d.Category,
            c.Category AS CategoryName,
            d.Company,
            ISNULL(d.SegmentCode, d.Segment) AS Segment,
            s.SegmentName,
            d.Brand,
            b.[Brand Name] AS BrandName,
            d.Vendor,
            v.Vendor AS VendorName,
            d.Amount,
            d.Remark,
            COALESCE(d.[Version], N'A1') AS [Version]
        FROM dbo.Template_Upload_Draft_OTB d
        INNER JOIN @Ids i
            ON d.RunNo = i.RunNo
        LEFT JOIN dbo.MS_Category c
            ON d.Category = c.Cate
        LEFT JOIN dbo.MS_Segment s
            ON ISNULL(d.SegmentCode, d.Segment) = s.SegmentCode
        LEFT JOIN dbo.MS_Brand b
            ON d.Brand = b.[Brand Code]
        LEFT JOIN dbo.MS_Vendor v
            ON d.Vendor = v.VendorCode AND ISNULL(d.SegmentCode, d.Segment) = v.SegmentCode
    ) AS S
        ON T.[Type] = S.[Type]
       AND T.[Year] = S.[Year]
       AND T.[Month] = S.[Month]
       AND T.Category = S.Category
       AND T.Company = S.Company
       AND T.Segment = S.Segment
       AND T.Brand = S.Brand
       AND T.Vendor = S.Vendor
       AND T.[Version] = S.[Version]
    WHEN MATCHED THEN
        UPDATE SET
            T.Amount = S.Amount,
            T.Remark = S.Remark,
            T.OTBStatus = N'Approved',
            T.ApprovedDate = sysdatetime(),
            T.ActionBy = @ApprovedBy,
            T.DraftID = S.RunNo,
            T.CategoryName = S.CategoryName,
            T.SegmentName = S.SegmentName,
            T.BrandName = S.BrandName,
            T.VendorName = S.VendorName
    WHEN NOT MATCHED BY TARGET THEN
        INSERT
        (
            CreateDate, [Type], [Year], [Month], Category, CategoryName, Company, Segment, SegmentName,
            Brand, BrandName, Vendor, VendorName, Amount, RevisedDiff, Remark, OTBStatus, ApprovedDate,
            SAPDate, ActionBy, DraftID, SAPStatus, SAPErrorMessage, [Version]
        )
        VALUES
        (
            sysdatetime(), S.[Type], S.[Year], S.[Month], S.Category, S.CategoryName, S.Company, S.Segment, S.SegmentName,
            S.Brand, S.BrandName, S.Vendor, S.VendorName, S.Amount, CASE WHEN S.[Type] = N'Revise' THEN S.Amount ELSE 0 END,
            S.Remark, N'Approved', sysdatetime(), NULL, @ApprovedBy, S.RunNo, NULL, NULL, S.[Version]
        );

    SELECT
        COUNT(1) AS ApprovedCount,
        N'Success' AS [Status],
        sysdatetime() AS SAPDate
    FROM @Ids;
END;
GO

CREATE OR ALTER PROCEDURE dbo.SP_Deleted_Draft_OTB
    @runNos nvarchar(max)
AS
BEGIN
    SET NOCOUNT ON;

    DECLARE @Ids TABLE (RunNo int PRIMARY KEY);

    INSERT INTO @Ids (RunNo)
    SELECT DISTINCT TRY_CAST(value AS int)
    FROM string_split(@runNos, N',')
    WHERE TRY_CAST(value AS int) IS NOT NULL;

    UPDATE d
        SET d.OTBStatus = N'Cancelled',
            d.UpdateDT = sysdatetime()
    FROM dbo.Template_Upload_Draft_OTB d
    INNER JOIN @Ids i
        ON d.RunNo = i.RunNo
    WHERE ISNULL(d.OTBStatus, N'Draft') IN (N'Draft', N'Waiting', N'Edited');

    SELECT
        @@ROWCOUNT AS DeletedCount,
        N'Success' AS [Status],
        NULL AS [Message];
END;
GO
