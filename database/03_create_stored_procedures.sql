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
            WHEN LEFT(ISNULL(a.Segment_Code, N''), 1) = N'O' AND RIGHT(ISNULL(a.Segment_Code, N''), 1) = N'0' AND LEN(ISNULL(a.Segment_Code, N'')) > 2 THEN SUBSTRING(a.Segment_Code, 2, LEN(a.Segment_Code) - 2)
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
               WHEN LEFT(ISNULL(a.Segment_Code, N''), 1) = N'O' AND RIGHT(ISNULL(a.Segment_Code, N''), 1) = N'0' AND LEN(ISNULL(a.Segment_Code, N'')) > 2 THEN SUBSTRING(a.Segment_Code, 2, LEN(a.Segment_Code) - 2)
               ELSE a.Segment_Code
           END = s.SegmentCode
    LEFT JOIN dbo.MS_Brand b
        ON a.Brand_Code = b.[Brand Code]
    LEFT JOIN dbo.MS_Vendor v
        ON a.Vendor_Code = v.VendorCode
       AND CASE
               WHEN LEFT(ISNULL(a.Segment_Code, N''), 1) = N'O' AND RIGHT(ISNULL(a.Segment_Code, N''), 1) = N'0' AND LEN(ISNULL(a.Segment_Code, N'')) > 2 THEN SUBSTRING(a.Segment_Code, 2, LEN(a.Segment_Code) - 2)
               ELSE a.Segment_Code
           END = v.SegmentCode
    WHERE (@Year IS NULL OR CONVERT(nvarchar(10), a.OTB_Year) = @Year)
      AND (@Month IS NULL OR CONVERT(nvarchar(10), a.OTB_Month) = @Month)
      AND (@Company IS NULL OR a.Company_Code = @Company)
      AND (@Category IS NULL OR a.Category_Code = @Category)
      AND (@Segment IS NULL OR CASE
                                   WHEN LEFT(ISNULL(a.Segment_Code, N''), 1) = N'O' AND RIGHT(ISNULL(a.Segment_Code, N''), 1) = N'0' AND LEN(ISNULL(a.Segment_Code, N'')) > 2 THEN SUBSTRING(a.Segment_Code, 2, LEN(a.Segment_Code) - 2)
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

CREATE OR ALTER PROCEDURE dbo.SP_Get_Draft_OTB_Diff
    @Version nvarchar(20) = NULL,
    @Type nvarchar(20) = NULL,
    @Year nvarchar(10) = NULL,
    @Month nvarchar(10) = NULL,
    @Company nvarchar(20) = NULL,
    @Category nvarchar(20) = NULL,
    @Segment nvarchar(20) = NULL,
    @Brand nvarchar(30) = NULL,
    @Vendor nvarchar(30) = NULL,
    @Status nvarchar(30) = NULL
AS
BEGIN
    SET NOCOUNT ON;

    DECLARE @VersionFilter nvarchar(20) = NULLIF(LTRIM(RTRIM(@Version)), N'');
    DECLARE @TypeFilter nvarchar(20) = NULLIF(LTRIM(RTRIM(@Type)), N'');
    DECLARE @YearFilter nvarchar(10) = NULLIF(LTRIM(RTRIM(@Year)), N'');
    DECLARE @MonthFilter nvarchar(10) = NULLIF(LTRIM(RTRIM(@Month)), N'');
    DECLARE @CompanyFilter nvarchar(20) = NULLIF(LTRIM(RTRIM(@Company)), N'');
    DECLARE @CategoryFilter nvarchar(20) = NULLIF(LTRIM(RTRIM(@Category)), N'');
    DECLARE @SegmentFilter nvarchar(20) = NULLIF(LTRIM(RTRIM(@Segment)), N'');
    DECLARE @BrandFilter nvarchar(30) = NULLIF(LTRIM(RTRIM(@Brand)), N'');
    DECLARE @VendorFilter nvarchar(30) = NULLIF(LTRIM(RTRIM(@Vendor)), N'');
    DECLARE @StatusFilter nvarchar(30) = NULLIF(LTRIM(RTRIM(@Status)), N'');

    ;WITH DraftRows AS
    (
        SELECT
            d.RunNo,
            d.CreateDT,
            d.OTBType,
            d.OTBYear,
            d.OTBMonth,
            d.month_name_sh,
            d.OTBCategory,
            d.CateName,
            d.OTBCompany,
            d.CompanyName,
            d.OTBSegment,
            d.SegmentName,
            d.OTBBrand,
            d.BrandName,
            d.OTBVendor,
            d.Vendor AS VendorName,
            d.Amount,
            d.Batch,
            d.UploadBy,
            d.Remark,
            d.[Version],
            d.OTBStatus
        FROM dbo.View_OTB_Draft d
        WHERE (@VersionFilter IS NULL OR d.[Version] = @VersionFilter)
          AND (@TypeFilter IS NULL OR d.OTBType = @TypeFilter)
          AND (@YearFilter IS NULL OR CONVERT(nvarchar(10), d.OTBYear) = @YearFilter)
          AND (@MonthFilter IS NULL OR CONVERT(nvarchar(10), d.OTBMonth) = @MonthFilter)
          AND (@CompanyFilter IS NULL OR d.OTBCompany = @CompanyFilter)
          AND (@CategoryFilter IS NULL OR d.OTBCategory = @CategoryFilter)
          AND (@SegmentFilter IS NULL OR d.OTBSegment = @SegmentFilter)
          AND (@BrandFilter IS NULL OR d.OTBBrand = @BrandFilter)
          AND (@VendorFilter IS NULL OR d.OTBVendor = @VendorFilter)
          AND (@StatusFilter IS NULL OR d.OTBStatus = @StatusFilter)
    ),
    DraftKeys AS
    (
        SELECT DISTINCT
            OTBYear,
            OTBMonth,
            OTBCategory,
            OTBCompany,
            OTBSegment,
            OTBBrand,
            OTBVendor
        FROM DraftRows
    ),
    BudgetRows AS
    (
        SELECT
            t.[Year] AS OTBYear,
            t.[Month] AS OTBMonth,
            t.Category AS OTBCategory,
            t.Company AS OTBCompany,
            t.Segment AS OTBSegment,
            t.Brand AS OTBBrand,
            t.Vendor AS OTBVendor,
            CASE WHEN t.[Type] = N'Original' THEN ISNULL(t.Amount, 0) ELSE 0 END AS Original,
            CASE WHEN t.[Type] = N'Revise' THEN ISNULL(t.RevisedDiff, 0) ELSE 0 END AS RevDiff,
            CAST(0 AS decimal(18,2)) AS Extra,
            CAST(0 AS decimal(18,2)) AS SwitchIn,
            CAST(0 AS decimal(18,2)) AS BalanceIn,
            CAST(0 AS decimal(18,2)) AS CarryIn,
            CAST(0 AS decimal(18,2)) AS SwitchOut,
            CAST(0 AS decimal(18,2)) AS BalanceOut,
            CAST(0 AS decimal(18,2)) AS CarryOut
        FROM dbo.OTB_Transaction t
        INNER JOIN DraftKeys k
            ON t.[Year] = k.OTBYear
           AND t.[Month] = k.OTBMonth
           AND t.Category = k.OTBCategory
           AND t.Company = k.OTBCompany
           AND t.Segment = k.OTBSegment
           AND t.Brand = k.OTBBrand
           AND t.Vendor = k.OTBVendor
        WHERE t.OTBStatus = N'Approved'

        UNION ALL

        SELECT
            s.[Year] AS OTBYear,
            s.[Month] AS OTBMonth,
            s.Category AS OTBCategory,
            s.Company AS OTBCompany,
            s.Segment AS OTBSegment,
            s.Brand AS OTBBrand,
            s.Vendor AS OTBVendor,
            CAST(0 AS decimal(18,2)) AS Original,
            CAST(0 AS decimal(18,2)) AS RevDiff,
            CASE WHEN s.[From] = N'E' THEN ISNULL(s.BudgetAmount, 0) ELSE 0 END AS Extra,
            CAST(0 AS decimal(18,2)) AS SwitchIn,
            CAST(0 AS decimal(18,2)) AS BalanceIn,
            CAST(0 AS decimal(18,2)) AS CarryIn,
            CASE WHEN s.[From] = N'D' THEN ISNULL(s.BudgetAmount, 0) ELSE 0 END AS SwitchOut,
            CASE WHEN s.[From] = N'I' THEN ISNULL(s.BudgetAmount, 0) ELSE 0 END AS BalanceOut,
            CASE WHEN s.[From] = N'G' THEN ISNULL(s.BudgetAmount, 0) ELSE 0 END AS CarryOut
        FROM dbo.OTB_Switching_Transaction s
        INNER JOIN DraftKeys k
            ON s.[Year] = k.OTBYear
           AND s.[Month] = k.OTBMonth
           AND s.Category = k.OTBCategory
           AND s.Company = k.OTBCompany
           AND s.Segment = k.OTBSegment
           AND s.Brand = k.OTBBrand
           AND s.Vendor = k.OTBVendor
        WHERE s.OTBStatus = N'Approved'

        UNION ALL

        SELECT
            s.SwitchYear AS OTBYear,
            s.SwitchMonth AS OTBMonth,
            s.SwitchCategory AS OTBCategory,
            s.SwitchCompany AS OTBCompany,
            s.SwitchSegment AS OTBSegment,
            s.SwitchBrand AS OTBBrand,
            s.SwitchVendor AS OTBVendor,
            CAST(0 AS decimal(18,2)) AS Original,
            CAST(0 AS decimal(18,2)) AS RevDiff,
            CAST(0 AS decimal(18,2)) AS Extra,
            CASE WHEN s.[To] = N'C' THEN ISNULL(s.BudgetAmount, 0) ELSE 0 END AS SwitchIn,
            CASE WHEN s.[To] = N'H' THEN ISNULL(s.BudgetAmount, 0) ELSE 0 END AS BalanceIn,
            CASE WHEN s.[To] = N'F' THEN ISNULL(s.BudgetAmount, 0) ELSE 0 END AS CarryIn,
            CAST(0 AS decimal(18,2)) AS SwitchOut,
            CAST(0 AS decimal(18,2)) AS BalanceOut,
            CAST(0 AS decimal(18,2)) AS CarryOut
        FROM dbo.OTB_Switching_Transaction s
        INNER JOIN DraftKeys k
            ON s.SwitchYear = k.OTBYear
           AND s.SwitchMonth = k.OTBMonth
           AND s.SwitchCategory = k.OTBCategory
           AND s.SwitchCompany = k.OTBCompany
           AND s.SwitchSegment = k.OTBSegment
           AND s.SwitchBrand = k.OTBBrand
           AND s.SwitchVendor = k.OTBVendor
        WHERE s.OTBStatus = N'Approved'
          AND s.[To] IS NOT NULL
          AND s.SwitchYear IS NOT NULL
          AND s.SwitchMonth IS NOT NULL
          AND NULLIF(LTRIM(RTRIM(s.SwitchCompany)), N'') IS NOT NULL
          AND NULLIF(LTRIM(RTRIM(s.SwitchCategory)), N'') IS NOT NULL
          AND NULLIF(LTRIM(RTRIM(s.SwitchSegment)), N'') IS NOT NULL
          AND NULLIF(LTRIM(RTRIM(s.SwitchBrand)), N'') IS NOT NULL
          AND NULLIF(LTRIM(RTRIM(s.SwitchVendor)), N'') IS NOT NULL
    ),
    BudgetTotals AS
    (
        SELECT
            OTBYear,
            OTBMonth,
            OTBCategory,
            OTBCompany,
            OTBSegment,
            OTBBrand,
            OTBVendor,
            SUM(Original) AS Original,
            SUM(RevDiff) AS RevDiff,
            SUM(Extra) AS Extra,
            SUM(SwitchIn) AS SwitchIn,
            SUM(BalanceIn) AS BalanceIn,
            SUM(CarryIn) AS CarryIn,
            SUM(SwitchOut) AS SwitchOut,
            SUM(BalanceOut) AS BalanceOut,
            SUM(CarryOut) AS CarryOut
        FROM BudgetRows
        GROUP BY OTBYear, OTBMonth, OTBCategory, OTBCompany, OTBSegment, OTBBrand, OTBVendor
    )
    SELECT
        d.RunNo,
        d.CreateDT,
        d.OTBType,
        d.OTBYear,
        d.OTBMonth,
        d.month_name_sh,
        d.OTBCategory,
        d.CateName,
        d.OTBCompany,
        d.CompanyName,
        d.OTBSegment,
        d.SegmentName,
        d.OTBBrand,
        d.BrandName,
        d.OTBVendor,
        d.VendorName AS Vendor,
        d.VendorName,
        d.Amount,
        budget.CurrentApproved,
        d.Amount AS ToBeAmountTHB,
        CAST(ISNULL(d.Amount, 0) - budget.CurrentApproved AS decimal(18,2)) AS Diff,
        d.OTBStatus,
        d.[Version],
        d.Remark,
        d.UploadBy,
        d.Batch
    FROM DraftRows d
    LEFT JOIN BudgetTotals b
        ON d.OTBYear = b.OTBYear
       AND d.OTBMonth = b.OTBMonth
       AND d.OTBCategory = b.OTBCategory
       AND d.OTBCompany = b.OTBCompany
       AND d.OTBSegment = b.OTBSegment
       AND d.OTBBrand = b.OTBBrand
       AND d.OTBVendor = b.OTBVendor
    CROSS APPLY
    (
        SELECT CAST(
            ISNULL(b.Original, 0)
            + ISNULL(b.RevDiff, 0)
            + ISNULL(b.Extra, 0)
            + ISNULL(b.SwitchIn, 0)
            + ISNULL(b.BalanceIn, 0)
            + ISNULL(b.CarryIn, 0)
            - ISNULL(b.SwitchOut, 0)
            - ISNULL(b.BalanceOut, 0)
            - ISNULL(b.CarryOut, 0)
            AS decimal(18,2)
        ) AS CurrentApproved
    ) budget
    ORDER BY d.CreateDT DESC, d.RunNo DESC;
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

    CREATE TABLE #CurrentActualPO
    (
        Summary_ActualPO_ID int NULL,
        PO_No nvarchar(50) NOT NULL,
        OTB_Year_Key nvarchar(4) NOT NULL,
        OTB_Month_Key nvarchar(2) NOT NULL,
        OTB_Year int NOT NULL,
        OTB_Month int NOT NULL,
        Company_Code nvarchar(10) NOT NULL,
        Category_Code nvarchar(20) NOT NULL,
        Segment_Code nvarchar(20) NOT NULL,
        Clean_Segment nvarchar(20) NOT NULL,
        Brand_Code nvarchar(20) NOT NULL,
        Vendor_Code nvarchar(20) NOT NULL,
        Actual_PO_Date datetime NULL,
        Amount_THB decimal(18,2) NOT NULL,
        Amount_CCY decimal(18,2) NOT NULL,
        CCY nvarchar(5) NOT NULL,
        Exchange_Rate decimal(18,6) NOT NULL,
        Source_Modified_Date datetime NULL,
        BMS_Last_Synced datetime NULL,
        SyncStatus nvarchar(20) NULL
    );

    ;WITH NormalizedStaging AS
    (
        SELECT
            k.PO_No,
            k.PO_Item,
            CASE WHEN k.OTB_Year_Text IS NULL THEN NULL ELSE RIGHT(REPLICATE(N'0', 4) + k.OTB_Year_Text, 4) END AS OTB_Year_Key,
            CASE WHEN k.OTB_Month_Text IS NULL THEN NULL ELSE RIGHT(REPLICATE(N'0', 2) + k.OTB_Month_Text, 2) END AS OTB_Month_Key,
            TRY_CONVERT(int, k.OTB_Year_Text) AS OTB_Year,
            TRY_CONVERT(int, k.OTB_Month_Text) AS OTB_Month,
            k.Company_Code,
            k.Category_Code,
            k.Segment_Code,
            CASE
                WHEN LEFT(ISNULL(k.Segment_Code, N''), 1) = N'O'
                 AND RIGHT(ISNULL(k.Segment_Code, N''), 1) = N'0'
                 AND LEN(ISNULL(k.Segment_Code, N'')) > 2
                THEN SUBSTRING(k.Segment_Code, 2, LEN(k.Segment_Code) - 2)
                ELSE k.Segment_Code
            END AS Clean_Segment,
            k.Brand_Code,
            k.Vendor_Code,
            TRY_CONVERT(datetime, s.Otb_Date) AS Actual_PO_Date,
            TRY_CONVERT(decimal(18,2), s.PO_Local_Amount) AS Amount_THB,
            TRY_CONVERT(decimal(18,2), s.PO_Amount) AS Amount_CCY,
            LEFT(COALESCE(NULLIF(s.PO_Currency, N''), NULLIF(s.PO_Local_Currency, N''), N''), 5) AS CCY,
            TRY_CONVERT(decimal(18,6), s.Exchange_Rate) AS Exchange_Rate,
            TRY_CONVERT(datetime, s.Modified_Date) AS Source_Modified_Date,
            TRY_CONVERT(datetime, s.BMS_Last_Synced) AS BMS_Last_Synced,
            TRY_CONVERT(datetime, s.Change_On) AS Source_Change_On,
            TRY_CONVERT(datetime, s.Create_On) AS Source_Create_On,
            CASE WHEN ISNULL(s.Deletion_Flag, N'') = N'L' THEN 1 ELSE 0 END AS IsDeleted
        FROM dbo.Actual_PO_Staging s
        CROSS APPLY
        (
            SELECT
                NULLIF(LTRIM(RTRIM(CONVERT(nvarchar(50), s.PO))), N'') AS PO_No,
                NULLIF(LTRIM(RTRIM(CONVERT(nvarchar(10), s.PO_Item))), N'') AS PO_Item,
                NULLIF(LTRIM(RTRIM(CONVERT(nvarchar(4), s.Otb_Year))), N'') AS OTB_Year_Text,
                NULLIF(LTRIM(RTRIM(CONVERT(nvarchar(2), s.Otb_Month))), N'') AS OTB_Month_Text,
                NULLIF(LTRIM(RTRIM(CONVERT(nvarchar(10), s.Company_Code))), N'') AS Company_Code,
                NULLIF(LTRIM(RTRIM(CONVERT(nvarchar(20), s.Category))), N'') AS Category_Code,
                NULLIF(LTRIM(RTRIM(CONVERT(nvarchar(20), s.Fund))), N'') AS Segment_Code,
                NULLIF(LTRIM(RTRIM(CONVERT(nvarchar(20), s.Brand))), N'') AS Brand_Code,
                NULLIF(LTRIM(RTRIM(CONVERT(nvarchar(20), s.Supplier))), N'') AS Vendor_Code
        ) k
        WHERE k.PO_No IS NOT NULL
          AND k.PO_Item IS NOT NULL
    ),
    RankedStaging AS
    (
        SELECT
            ns.*,
            ROW_NUMBER() OVER
            (
                PARTITION BY
                    ns.PO_No,
                    ns.PO_Item,
                    ns.OTB_Year_Key,
                    ns.OTB_Month_Key,
                    ns.Company_Code,
                    ns.Vendor_Code,
                    ns.Segment_Code,
                    ns.Category_Code,
                    ns.Brand_Code
                ORDER BY
                    ns.BMS_Last_Synced DESC,
                    ns.Source_Modified_Date DESC,
                    ns.Source_Change_On DESC,
                    ns.Source_Create_On DESC,
                    ns.IsDeleted ASC
            ) AS rn
        FROM NormalizedStaging ns
        WHERE ns.OTB_Year IS NOT NULL
          AND ns.OTB_Month IS NOT NULL
          AND ns.Company_Code IS NOT NULL
          AND ns.Category_Code IS NOT NULL
          AND ns.Segment_Code IS NOT NULL
          AND ns.Clean_Segment IS NOT NULL
          AND ns.Brand_Code IS NOT NULL
          AND ns.Vendor_Code IS NOT NULL
    ),
    CurrentItems AS
    (
        SELECT *
        FROM RankedStaging
        WHERE rn = 1
    ),
    AggregatedActualPO AS
    (
        SELECT
            PO_No,
            OTB_Year_Key,
            OTB_Month_Key,
            MAX(OTB_Year) AS OTB_Year,
            MAX(OTB_Month) AS OTB_Month,
            Company_Code,
            Category_Code,
            Segment_Code,
            Clean_Segment,
            Brand_Code,
            Vendor_Code,
            MAX(CASE WHEN IsDeleted = 0 THEN Actual_PO_Date ELSE NULL END) AS Actual_PO_Date,
            CAST(SUM(CASE WHEN IsDeleted = 0 THEN ISNULL(Amount_THB, 0) ELSE 0 END) AS decimal(18,2)) AS Amount_THB,
            CAST(SUM(CASE WHEN IsDeleted = 0 THEN ISNULL(Amount_CCY, 0) ELSE 0 END) AS decimal(18,2)) AS Amount_CCY,
            COALESCE(MAX(CASE WHEN IsDeleted = 0 THEN NULLIF(CCY, N'') ELSE NULL END), MAX(NULLIF(CCY, N'')), N'') AS CCY,
            COALESCE(MAX(CASE WHEN IsDeleted = 0 THEN Exchange_Rate ELSE NULL END), MAX(Exchange_Rate), 0) AS Exchange_Rate,
            MAX(Source_Modified_Date) AS Source_Modified_Date,
            MAX(BMS_Last_Synced) AS BMS_Last_Synced,
            CASE
                WHEN SUM(CASE WHEN IsDeleted = 0 THEN 1 ELSE 0 END) = 0 THEN N'Cancelled'
                ELSE NULL
            END AS SyncStatus
        FROM CurrentItems
        GROUP BY
            PO_No,
            OTB_Year_Key,
            OTB_Month_Key,
            Company_Code,
            Category_Code,
            Segment_Code,
            Clean_Segment,
            Brand_Code,
            Vendor_Code
    )
    INSERT INTO #CurrentActualPO
    (
        PO_No, OTB_Year_Key, OTB_Month_Key, OTB_Year, OTB_Month, Company_Code,
        Category_Code, Segment_Code, Clean_Segment, Brand_Code, Vendor_Code,
        Actual_PO_Date, Amount_THB, Amount_CCY, CCY, Exchange_Rate,
        Source_Modified_Date, BMS_Last_Synced, SyncStatus
    )
    SELECT
        PO_No, OTB_Year_Key, OTB_Month_Key, OTB_Year, OTB_Month, Company_Code,
        Category_Code, Segment_Code, Clean_Segment, Brand_Code, Vendor_Code,
        Actual_PO_Date, Amount_THB, Amount_CCY, CCY, Exchange_Rate,
        Source_Modified_Date, BMS_Last_Synced, SyncStatus
    FROM AggregatedActualPO;

    UPDATE d
        SET d.Actual_PO_Ref = NULL,
            d.[Status] = CASE WHEN d.[Status] = N'Edited' THEN N'Edited' ELSE N'Draft' END,
            d.Status_Date = sysdatetime(),
            d.Status_By = N'System Sync'
    FROM dbo.Draft_PO_Transaction d
    INNER JOIN dbo.Actual_PO_Summary a
        ON d.DraftPO_No = a.Draft_PO_Ref
       AND d.Actual_PO_Ref = a.PO_No
       AND d.PO_Year = a.OTB_Year
       AND d.PO_Month = a.OTB_Month
       AND d.Company_Code = a.Company_Code
       AND d.Category_Code = a.Category_Code
       AND d.Segment_Code = COALESCE(NULLIF(a.Clean_Segment, N''), CASE
                                WHEN LEFT(ISNULL(a.Segment_Code, N''), 1) = N'O'
                                 AND RIGHT(ISNULL(a.Segment_Code, N''), 1) = N'0'
                                 AND LEN(ISNULL(a.Segment_Code, N'')) > 2
                                THEN SUBSTRING(a.Segment_Code, 2, LEN(a.Segment_Code) - 2)
                                ELSE a.Segment_Code
                            END)
       AND d.Brand_Code = a.Brand_Code
       AND d.Vendor_Code = a.Vendor_Code
    INNER JOIN #CurrentActualPO s
        ON a.PO_No = s.PO_No
       AND a.OTB_Year = s.OTB_Year
       AND a.OTB_Month = s.OTB_Month
       AND a.Company_Code = s.Company_Code
       AND a.Category_Code = s.Category_Code
       AND COALESCE(NULLIF(a.Clean_Segment, N''), CASE
                    WHEN LEFT(ISNULL(a.Segment_Code, N''), 1) = N'O'
                     AND RIGHT(ISNULL(a.Segment_Code, N''), 1) = N'0'
                     AND LEN(ISNULL(a.Segment_Code, N'')) > 2
                    THEN SUBSTRING(a.Segment_Code, 2, LEN(a.Segment_Code) - 2)
                    ELSE a.Segment_Code
                  END) = s.Clean_Segment
       AND a.Brand_Code = s.Brand_Code
       AND a.Vendor_Code = s.Vendor_Code
    WHERE a.Draft_PO_Ref IS NOT NULL
      AND ISNULL(d.[Status], N'') IN (N'Draft', N'Edited', N'Matching', N'ForceMatching', N'Matched')
      AND s.SyncStatus = N'Cancelled';

    ;WITH ExistingSummary AS
    (
        SELECT
            a.ActualPO_ID,
            s.PO_No,
            s.OTB_Year,
            s.OTB_Month,
            s.Company_Code,
            s.Category_Code,
            s.Clean_Segment,
            s.Brand_Code,
            s.Vendor_Code,
            ROW_NUMBER() OVER
            (
                PARTITION BY
                    s.PO_No,
                    s.OTB_Year,
                    s.OTB_Month,
                    s.Company_Code,
                    s.Category_Code,
                    s.Clean_Segment,
                    s.Brand_Code,
                    s.Vendor_Code
                ORDER BY
                    CASE WHEN a.Draft_PO_Ref IS NOT NULL THEN 0 ELSE 1 END,
                    CASE WHEN ISNULL(a.[Status], N'') IN (N'Matching', N'ForceMatching', N'Matched') THEN 0 ELSE 1 END,
                    CASE WHEN ISNULL(a.[Status], N'') <> N'Cancelled' THEN 0 ELSE 1 END,
                    a.ActualPO_ID
            ) AS rn
        FROM dbo.Actual_PO_Summary a
        INNER JOIN #CurrentActualPO s
            ON a.PO_No = s.PO_No
           AND a.OTB_Year = s.OTB_Year
           AND a.OTB_Month = s.OTB_Month
           AND a.Company_Code = s.Company_Code
           AND a.Category_Code = s.Category_Code
           AND COALESCE(NULLIF(a.Clean_Segment, N''), CASE
                        WHEN LEFT(ISNULL(a.Segment_Code, N''), 1) = N'O'
                         AND RIGHT(ISNULL(a.Segment_Code, N''), 1) = N'0'
                         AND LEN(ISNULL(a.Segment_Code, N'')) > 2
                        THEN SUBSTRING(a.Segment_Code, 2, LEN(a.Segment_Code) - 2)
                        ELSE a.Segment_Code
                      END) = s.Clean_Segment
           AND a.Brand_Code = s.Brand_Code
           AND a.Vendor_Code = s.Vendor_Code
    )
    UPDATE s
        SET Summary_ActualPO_ID = e.ActualPO_ID
    FROM #CurrentActualPO s
    INNER JOIN ExistingSummary e
        ON s.PO_No = e.PO_No
       AND s.OTB_Year = e.OTB_Year
       AND s.OTB_Month = e.OTB_Month
       AND s.Company_Code = e.Company_Code
       AND s.Category_Code = e.Category_Code
       AND s.Clean_Segment = e.Clean_Segment
       AND s.Brand_Code = e.Brand_Code
       AND s.Vendor_Code = e.Vendor_Code
       AND e.rn = 1;

    ;WITH DuplicateSummary AS
    (
        SELECT
            a.ActualPO_ID,
            ROW_NUMBER() OVER
            (
                PARTITION BY
                    s.PO_No,
                    s.OTB_Year,
                    s.OTB_Month,
                    s.Company_Code,
                    s.Category_Code,
                    s.Clean_Segment,
                    s.Brand_Code,
                    s.Vendor_Code
                ORDER BY
                    CASE WHEN a.ActualPO_ID = s.Summary_ActualPO_ID THEN 0 ELSE 1 END,
                    CASE WHEN a.Draft_PO_Ref IS NOT NULL THEN 0 ELSE 1 END,
                    a.ActualPO_ID
            ) AS rn
        FROM dbo.Actual_PO_Summary a
        INNER JOIN #CurrentActualPO s
            ON a.PO_No = s.PO_No
           AND a.OTB_Year = s.OTB_Year
           AND a.OTB_Month = s.OTB_Month
           AND a.Company_Code = s.Company_Code
           AND a.Category_Code = s.Category_Code
           AND COALESCE(NULLIF(a.Clean_Segment, N''), CASE
                        WHEN LEFT(ISNULL(a.Segment_Code, N''), 1) = N'O'
                         AND RIGHT(ISNULL(a.Segment_Code, N''), 1) = N'0'
                         AND LEN(ISNULL(a.Segment_Code, N'')) > 2
                        THEN SUBSTRING(a.Segment_Code, 2, LEN(a.Segment_Code) - 2)
                        ELSE a.Segment_Code
                      END) = s.Clean_Segment
           AND a.Brand_Code = s.Brand_Code
           AND a.Vendor_Code = s.Vendor_Code
    )
    UPDATE a
        SET a.Amount_THB = 0,
            a.Amount_CCY = 0,
            a.Draft_PO_ID_Ref = NULL,
            a.Draft_PO_Ref = NULL,
            a.Matching_Date = NULL,
            a.[Status] = N'Cancelled',
            a.Status_Date = sysdatetime(),
            a.Status_By = N'System Sync',
            a.Changed_By = N'System Sync',
            a.Changed_date = sysdatetime(),
            a.Updated_By = N'System Sync',
            a.Updated_Date = sysdatetime()
    FROM dbo.Actual_PO_Summary a
    INNER JOIN DuplicateSummary d
        ON a.ActualPO_ID = d.ActualPO_ID
    WHERE d.rn > 1;

    UPDATE t
        SET t.OTB_Year = s.OTB_Year,
            t.OTB_Month = s.OTB_Month,
            t.Company_Code = s.Company_Code,
            t.Category_Code = s.Category_Code,
            t.Segment_Code = s.Segment_Code,
            t.Clean_Segment = s.Clean_Segment,
            t.Brand_Code = s.Brand_Code,
            t.Vendor_Code = s.Vendor_Code,
            t.Actual_PO_Date = s.Actual_PO_Date,
            t.Amount_THB = s.Amount_THB,
            t.Amount_CCY = s.Amount_CCY,
            t.CCY = s.CCY,
            t.Exchange_Rate = s.Exchange_Rate,
            t.Draft_PO_ID_Ref = CASE WHEN s.SyncStatus = N'Cancelled' OR ISNULL(t.[Status], N'') = N'Cancelled' THEN NULL ELSE t.Draft_PO_ID_Ref END,
            t.Draft_PO_Ref = CASE WHEN s.SyncStatus = N'Cancelled' OR ISNULL(t.[Status], N'') = N'Cancelled' THEN NULL ELSE t.Draft_PO_Ref END,
            t.Matching_Date = CASE WHEN s.SyncStatus = N'Cancelled' OR ISNULL(t.[Status], N'') = N'Cancelled' THEN NULL ELSE t.Matching_Date END,
            t.[Status] = CASE
                WHEN s.SyncStatus IS NOT NULL THEN s.SyncStatus
                WHEN ISNULL(t.[Status], N'') IN (N'', N'Cancelled', N'Unmatched') THEN N'Active'
                ELSE t.[Status]
            END,
            t.Status_Date = CASE
                WHEN s.SyncStatus = N'Cancelled' OR ISNULL(t.[Status], N'') IN (N'', N'Cancelled', N'Unmatched') THEN sysdatetime()
                ELSE t.Status_Date
            END,
            t.Status_By = CASE
                WHEN s.SyncStatus = N'Cancelled' OR ISNULL(t.[Status], N'') IN (N'', N'Cancelled', N'Unmatched') THEN N'System Sync'
                ELSE t.Status_By
            END,
            t.Changed_By = CASE WHEN s.SyncStatus = N'Cancelled' THEN N'System Sync' ELSE t.Changed_By END,
            t.Changed_date = CASE WHEN s.SyncStatus = N'Cancelled' THEN sysdatetime() ELSE t.Changed_date END,
            t.Updated_By = N'System Sync',
            t.Updated_Date = sysdatetime()
    FROM dbo.Actual_PO_Summary t
    INNER JOIN #CurrentActualPO s
        ON t.ActualPO_ID = s.Summary_ActualPO_ID;

    INSERT INTO dbo.Actual_PO_Summary
    (
        PO_No, OTB_Year, OTB_Month, Company_Code, Category_Code, Segment_Code,
        Clean_Segment, Brand_Code, Vendor_Code, CCY, Exchange_Rate, Amount_CCY, Amount_THB,
        [Status], Status_Date, Status_By, Created_By, Created_Date, Actual_PO_Date, Updated_By, Updated_Date
    )
    SELECT
        s.PO_No, s.OTB_Year, s.OTB_Month, s.Company_Code, s.Category_Code, s.Segment_Code,
        s.Clean_Segment, s.Brand_Code, s.Vendor_Code, s.CCY, s.Exchange_Rate, s.Amount_CCY, s.Amount_THB,
        COALESCE(s.SyncStatus, N'Active'), sysdatetime(), N'System Sync', N'System Sync', sysdatetime(),
        s.Actual_PO_Date, N'System Sync', sysdatetime()
    FROM #CurrentActualPO s
    WHERE s.Summary_ActualPO_ID IS NULL;
END;
GO

CREATE OR ALTER PROCEDURE dbo.SP_Auto_Match_Actual_Draft
    @UpdateBy nvarchar(100) = N'System AutoMatch'
AS
BEGIN
    SET NOCOUNT ON;
    SET XACT_ABORT ON;

    DECLARE @Now datetime = GETDATE();

    BEGIN TRY
        BEGIN TRANSACTION;

        UPDATE A
            SET A.[Status] = N'Cancelled',
                A.Draft_PO_ID_Ref = NULL,
                A.Draft_PO_Ref = NULL,
                A.Matching_Date = NULL,
                A.Status_Date = @Now,
                A.Status_By = @UpdateBy,
                A.Changed_By = @UpdateBy,
                A.Changed_date = @Now,
                A.Updated_By = @UpdateBy,
                A.Updated_Date = @Now
        FROM dbo.Actual_PO_Summary A WITH (UPDLOCK, HOLDLOCK)
        WHERE ISNULL(A.Amount_THB, 0) = 0
          AND ISNULL(A.[Status], N'') NOT IN (N'Matched', N'Cancelled', N'Deleted');

        UPDATE D
            SET D.[Status] = N'Draft',
                D.Actual_PO_Ref = NULL,
                D.Actual_PO_Date = NULL,
                D.Status_Date = @Now,
                D.Status_By = @UpdateBy
        FROM dbo.Draft_PO_Transaction D WITH (UPDLOCK, HOLDLOCK)
        WHERE D.[Status] IN (N'Matching', N'ForceMatching')
          AND ISNULL(D.Actual_PO_Ref, N'') <> N''
          AND NOT EXISTS
          (
                SELECT 1
                FROM dbo.Actual_PO_Summary A WITH (UPDLOCK, HOLDLOCK)
                WHERE A.PO_No = D.Actual_PO_Ref
                  AND A.[Status] IN (N'Matching', N'ForceMatching', N'Matched')
                  AND (A.Draft_PO_ID_Ref = TRY_CONVERT(int, D.DraftPO_ID)
                       OR ISNULL(A.Draft_PO_Ref, N'') = ISNULL(D.DraftPO_No, N''))
          );

        UPDATE A
            SET A.[Status] = N'Active',
                A.Draft_PO_ID_Ref = NULL,
                A.Draft_PO_Ref = NULL,
                A.Matching_Date = NULL,
                A.Status_Date = @Now,
                A.Status_By = @UpdateBy,
                A.Changed_By = @UpdateBy,
                A.Changed_date = @Now,
                A.Updated_By = @UpdateBy,
                A.Updated_Date = @Now
        FROM dbo.Actual_PO_Summary A WITH (UPDLOCK, HOLDLOCK)
        INNER JOIN dbo.Draft_PO_Transaction D WITH (UPDLOCK, HOLDLOCK)
            ON D.Actual_PO_Ref = A.PO_No
           AND (A.Draft_PO_ID_Ref = TRY_CONVERT(int, D.DraftPO_ID)
                OR ISNULL(A.Draft_PO_Ref, N'') = ISNULL(D.DraftPO_No, N''))
        WHERE D.[Status] = N'Edited'
          AND A.[Status] IN (N'Matching', N'ForceMatching');

        UPDATE D
            SET D.Actual_PO_Ref = NULL,
                D.Actual_PO_Date = NULL,
                D.Status_Date = @Now,
                D.Status_By = @UpdateBy
        FROM dbo.Draft_PO_Transaction D WITH (UPDLOCK, HOLDLOCK)
        WHERE D.[Status] = N'Edited'
          AND ISNULL(D.Actual_PO_Ref, N'') <> N'';

        UPDATE A
            SET A.[Status] = N'Active',
                A.Draft_PO_ID_Ref = NULL,
                A.Draft_PO_Ref = NULL,
                A.Matching_Date = NULL,
                A.Status_Date = @Now,
                A.Status_By = @UpdateBy,
                A.Changed_By = @UpdateBy,
                A.Changed_date = @Now,
                A.Updated_By = @UpdateBy,
                A.Updated_Date = @Now
        FROM dbo.Actual_PO_Summary A WITH (UPDLOCK, HOLDLOCK)
        WHERE A.[Status] IN (N'Matching', N'ForceMatching')
          AND (A.Draft_PO_ID_Ref IS NOT NULL OR ISNULL(A.Draft_PO_Ref, N'') <> N'')
          AND NOT EXISTS
          (
                SELECT 1
                FROM dbo.Draft_PO_Transaction D WITH (UPDLOCK, HOLDLOCK)
                WHERE (A.Draft_PO_ID_Ref = TRY_CONVERT(int, D.DraftPO_ID)
                       OR ISNULL(D.DraftPO_No, N'') = ISNULL(A.Draft_PO_Ref, N''))
                  AND D.[Status] IN (N'Matching', N'ForceMatching', N'Matched')
                  AND ISNULL(D.Actual_PO_Ref, N'') = ISNULL(A.PO_No, N'')
          );

        IF OBJECT_ID(N'tempdb..#Actual_Prep') IS NOT NULL DROP TABLE #Actual_Prep;

        SELECT *,
               ROW_NUMBER() OVER
               (
                   PARTITION BY PO_No, OTB_Year, OTB_Month, Company_Code, Category_Code,
                                Clean_Segment, Brand_Code, Vendor_Code
                   ORDER BY ActualPO_ID
               ) AS LineRN
        INTO #Actual_Prep
        FROM
        (
            SELECT A.ActualPO_ID,
                   A.PO_No,
                   A.Amount_THB,
                   A.OTB_Year,
                   A.OTB_Month,
                   A.Company_Code,
                   A.Category_Code,
                   A.Brand_Code,
                   A.Vendor_Code,
                   A.[Status],
                   A.Actual_PO_Date,
                   COALESCE(NULLIF(A.Clean_Segment, N''), CASE
                       WHEN LEFT(ISNULL(A.Segment_Code, N''), 1) = N'O'
                        AND RIGHT(ISNULL(A.Segment_Code, N''), 1) = N'0'
                        AND LEN(ISNULL(A.Segment_Code, N'')) > 2
                       THEN SUBSTRING(A.Segment_Code, 2, LEN(A.Segment_Code) - 2)
                       ELSE A.Segment_Code
                   END) AS Clean_Segment
            FROM dbo.Actual_PO_Summary A WITH (UPDLOCK, HOLDLOCK)
            WHERE A.[Status] = N'Active'
              AND ISNULL(A.Amount_THB, 0) <> 0
              AND A.Draft_PO_ID_Ref IS NULL
              AND ISNULL(A.Draft_PO_Ref, N'') = N''
        ) T;

        CREATE CLUSTERED INDEX IX_A ON #Actual_Prep(ActualPO_ID);

        IF OBJECT_ID(N'tempdb..#Draft_Prep') IS NOT NULL DROP TABLE #Draft_Prep;

        SELECT D.*,
               ROW_NUMBER() OVER
               (
                   PARTITION BY D.DraftPO_No, D.PO_Year, D.PO_Month, D.Company_Code, D.Category_Code,
                                D.Segment_Code, D.Brand_Code, D.Vendor_Code
                   ORDER BY D.DraftPO_ID
               ) AS LineRN
        INTO #Draft_Prep
        FROM dbo.Draft_PO_Transaction D WITH (UPDLOCK, HOLDLOCK)
        WHERE D.[Status] IN (N'Draft', N'Edited')
          AND ISNULL(D.Amount_THB, 0) <> 0
          AND ISNULL(D.Actual_PO_Ref, N'') = N'';

        CREATE CLUSTERED INDEX IX_D ON #Draft_Prep(DraftPO_ID);

        CREATE TABLE #Final_Matches
        (
            ActualPO_ID int NOT NULL,
            DraftPO_ID bigint NOT NULL,
            DraftPO_No nvarchar(100) NOT NULL,
            ActualPO_No nvarchar(50) NOT NULL,
            Actual_PO_Date datetime NULL,
            Match_Type varchar(10) NOT NULL,
            CONSTRAINT PK_Final_Matches PRIMARY KEY (ActualPO_ID),
            CONSTRAINT UQ_Final_Matches_Draft UNIQUE (DraftPO_ID)
        );

        INSERT INTO #Final_Matches
            (ActualPO_ID, DraftPO_ID, DraftPO_No, ActualPO_No, Actual_PO_Date, Match_Type)
        SELECT A.ActualPO_ID,
               D.DraftPO_ID,
               D.DraftPO_No,
               A.PO_No,
               A.Actual_PO_Date,
               'EXACT'
        FROM #Actual_Prep A
        INNER JOIN #Draft_Prep D
            ON A.PO_No = D.DraftPO_No
           AND A.LineRN = D.LineRN
           AND (A.OTB_Year = D.PO_Year OR (A.OTB_Year IS NULL AND D.PO_Year IS NULL))
           AND (A.OTB_Month = D.PO_Month OR (A.OTB_Month IS NULL AND D.PO_Month IS NULL))
           AND (A.Company_Code = D.Company_Code OR (A.Company_Code IS NULL AND D.Company_Code IS NULL))
           AND (A.Category_Code = D.Category_Code OR (A.Category_Code IS NULL AND D.Category_Code IS NULL))
           AND (A.Clean_Segment = D.Segment_Code OR (A.Clean_Segment IS NULL AND D.Segment_Code IS NULL))
           AND (A.Brand_Code = D.Brand_Code OR (A.Brand_Code IS NULL AND D.Brand_Code IS NULL))
           AND (A.Vendor_Code = D.Vendor_Code OR (A.Vendor_Code IS NULL AND D.Vendor_Code IS NULL));

        INSERT INTO #Final_Matches
            (ActualPO_ID, DraftPO_ID, DraftPO_No, ActualPO_No, Actual_PO_Date, Match_Type)
        SELECT T2.ActualPO_ID,
               T2.DraftPO_ID,
               T2.DraftPO_No,
               T2.PO_No,
               T2.Actual_PO_Date,
               'FUZZY'
        FROM
        (
            SELECT T1.ActualPO_ID,
                   T1.DraftPO_ID,
                   T1.DraftPO_No,
                   T1.PO_No,
                   T1.Actual_PO_Date,
                   ROW_NUMBER() OVER
                   (
                       PARTITION BY T1.ActualPO_ID
                       ORDER BY T1.Diff_Amount ASC, T1.DraftPO_ID ASC
                   ) AS Rank_Actual_Best
            FROM
            (
                SELECT A.ActualPO_ID,
                       D.DraftPO_ID,
                       D.DraftPO_No,
                       A.PO_No,
                       A.Actual_PO_Date,
                       ABS(A.Amount_THB - D.Amount_THB) AS Diff_Amount,
                       ROW_NUMBER() OVER
                       (
                           PARTITION BY D.DraftPO_ID
                           ORDER BY ABS(A.Amount_THB - D.Amount_THB) ASC, A.ActualPO_ID ASC
                       ) AS Rank_Draft_Best
                FROM #Actual_Prep A
                INNER JOIN #Draft_Prep D
                    ON (A.OTB_Year = D.PO_Year OR (A.OTB_Year IS NULL AND D.PO_Year IS NULL))
                   AND (A.OTB_Month = D.PO_Month OR (A.OTB_Month IS NULL AND D.PO_Month IS NULL))
                   AND (A.Company_Code = D.Company_Code OR (A.Company_Code IS NULL AND D.Company_Code IS NULL))
                   AND (A.Category_Code = D.Category_Code OR (A.Category_Code IS NULL AND D.Category_Code IS NULL))
                   AND (A.Clean_Segment = D.Segment_Code OR (A.Clean_Segment IS NULL AND D.Segment_Code IS NULL))
                   AND (A.Brand_Code = D.Brand_Code OR (A.Brand_Code IS NULL AND D.Brand_Code IS NULL))
                   AND (A.Vendor_Code = D.Vendor_Code OR (A.Vendor_Code IS NULL AND D.Vendor_Code IS NULL))
                WHERE NOT EXISTS (SELECT 1 FROM #Final_Matches FM WHERE FM.ActualPO_ID = A.ActualPO_ID)
                  AND NOT EXISTS (SELECT 1 FROM #Final_Matches FM WHERE FM.DraftPO_ID = D.DraftPO_ID)
                  AND NOT EXISTS (SELECT 1 FROM #Actual_Prep A_Check WHERE A_Check.PO_No = D.DraftPO_No)
                  AND ABS(A.Amount_THB - D.Amount_THB) <= ABS(A.Amount_THB) * 0.10
            ) T1
            WHERE T1.Rank_Draft_Best = 1
        ) T2
        WHERE T2.Rank_Actual_Best = 1;

        UPDATE Main_A
            SET Main_A.[Status] = N'Matching',
                Main_A.Draft_PO_ID_Ref = TRY_CONVERT(int, M.DraftPO_ID),
                Main_A.Draft_PO_Ref = M.DraftPO_No,
                Main_A.Matching_Date = @Now,
                Main_A.Status_Date = @Now,
                Main_A.Status_By = @UpdateBy,
                Main_A.Changed_By = @UpdateBy,
                Main_A.Changed_date = @Now,
                Main_A.Updated_By = @UpdateBy,
                Main_A.Updated_Date = @Now
        FROM dbo.Actual_PO_Summary Main_A WITH (UPDLOCK, HOLDLOCK)
        INNER JOIN #Final_Matches M
            ON Main_A.ActualPO_ID = M.ActualPO_ID
        WHERE Main_A.[Status] = N'Active'
          AND Main_A.Draft_PO_ID_Ref IS NULL
          AND ISNULL(Main_A.Draft_PO_Ref, N'') = N'';

        UPDATE Main_D
            SET Main_D.[Status] = N'Matching',
                Main_D.Actual_PO_Ref = M.ActualPO_No,
                Main_D.Actual_PO_Date = M.Actual_PO_Date,
                Main_D.Status_Date = @Now,
                Main_D.Status_By = @UpdateBy
        FROM dbo.Draft_PO_Transaction Main_D WITH (UPDLOCK, HOLDLOCK)
        INNER JOIN #Final_Matches M
            ON Main_D.DraftPO_ID = M.DraftPO_ID
        WHERE Main_D.[Status] IN (N'Draft', N'Edited')
          AND ISNULL(Main_D.Actual_PO_Ref, N'') = N'';

        DROP TABLE #Actual_Prep;
        DROP TABLE #Draft_Prep;
        DROP TABLE #Final_Matches;

        COMMIT TRANSACTION;
    END TRY
    BEGIN CATCH
        IF @@TRANCOUNT > 0 ROLLBACK TRANSACTION;
        THROW;
    END CATCH
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

CREATE OR ALTER PROCEDURE dbo.SP_Sync_User_From_AD
    @Username nvarchar(255),
    @FullName nvarchar(255) = NULL,
    @Email nvarchar(255) = NULL
AS
BEGIN
    SET NOCOUNT ON;

    IF NULLIF(LTRIM(RTRIM(@Username)), N'') IS NULL
    BEGIN
        THROW 50001, 'Username is required.', 1;
    END;

    SET @Username = LTRIM(RTRIM(@Username));
    SET @FullName = NULLIF(LTRIM(RTRIM(ISNULL(@FullName, N''))), N'');
    SET @Email = NULLIF(LTRIM(RTRIM(ISNULL(@Email, N''))), N'');

    DECLARE @ExistingUserID int;
    SELECT TOP (1) @ExistingUserID = UserID
    FROM dbo.MS_User
    WHERE Username = @Username
       OR (@Email IS NOT NULL AND Email = @Email);

    IF @ExistingUserID IS NULL
    BEGIN
        DECLARE @NextUserID int;
        SELECT @NextUserID = ISNULL(MAX(UserID), 0) + 1 FROM dbo.MS_User;

        INSERT INTO dbo.MS_User
        (
            UserID,
            Username,
            FullName,
            Email,
            IsActive,
            LastLoginDate,
            CreateDate
        )
        VALUES
        (
            @NextUserID,
            @Username,
            COALESCE(@FullName, @Username),
            COALESCE(@Email, @Username),
            1,
            SYSDATETIME(),
            SYSDATETIME()
        );
    END
    ELSE
    BEGIN
        UPDATE dbo.MS_User
        SET Username = @Username,
            FullName = COALESCE(@FullName, FullName, @Username),
            Email = COALESCE(@Email, Email),
            IsActive = 1,
            LastLoginDate = SYSDATETIME()
        WHERE UserID = @ExistingUserID;
    END
END;
GO

CREATE OR ALTER PROCEDURE dbo.SP_Get_Users_List
    @SearchText nvarchar(255) = NULL
AS
BEGIN
    SET NOCOUNT ON;

    SET @SearchText = NULLIF(LTRIM(RTRIM(ISNULL(@SearchText, N''))), N'');

    SELECT
        u.UserID,
        u.Username,
        u.FullName,
        u.Email,
        u.IsActive,
        STRING_AGG(r.RoleName, N', ') AS Roles
    FROM dbo.MS_User u
    LEFT JOIN dbo.Map_User_Role ur
        ON u.UserID = ur.UserID
    LEFT JOIN dbo.MS_Role r
        ON ur.RoleID = r.RoleID
    WHERE @SearchText IS NULL
       OR u.Username LIKE N'%' + @SearchText + N'%'
       OR u.FullName LIKE N'%' + @SearchText + N'%'
       OR u.Email LIKE N'%' + @SearchText + N'%'
    GROUP BY
        u.UserID,
        u.Username,
        u.FullName,
        u.Email,
        u.IsActive
    ORDER BY u.Username;
END;
GO

CREATE OR ALTER PROCEDURE dbo.SP_Admin_Save_User
    @UserID int,
    @RoleID int,
    @IsActive bit
AS
BEGIN
    SET NOCOUNT ON;

    IF NOT EXISTS (SELECT 1 FROM dbo.MS_User WHERE UserID = @UserID)
    BEGIN
        THROW 50002, 'UserID not found.', 1;
    END;

    IF NOT EXISTS (SELECT 1 FROM dbo.MS_Role WHERE RoleID = @RoleID)
    BEGIN
        THROW 50003, 'RoleID not found.', 1;
    END;

    UPDATE dbo.MS_User
    SET IsActive = @IsActive
    WHERE UserID = @UserID;

    DELETE FROM dbo.Map_User_Role
    WHERE UserID = @UserID;

    INSERT INTO dbo.Map_User_Role (MapID, UserID, RoleID)
    SELECT
        ISNULL(MAX(MapID), 0) + 1,
        @UserID,
        @RoleID
    FROM dbo.Map_User_Role;
END;
GO
