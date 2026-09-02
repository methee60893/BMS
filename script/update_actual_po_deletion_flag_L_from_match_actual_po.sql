USE [BMS];
GO

SET NOCOUNT ON;
SET XACT_ABORT ON;

BEGIN TRY
    BEGIN TRAN;

    DECLARE @Keys TABLE
    (
        [PO] nvarchar(50) NOT NULL,
        [Otb_Year] int NOT NULL,
        [Otb_Month] nvarchar(2) NOT NULL,
        [Company_Code] nvarchar(10) NOT NULL,
        [Supplier] nvarchar(20) NOT NULL,
        [Fund] nvarchar(20) NOT NULL,
        [Category] nvarchar(20) NOT NULL,
        [Brand] nvarchar(20) NOT NULL
    );

    -- Source: C:\Users\60893\Desktop\BMS\Match Actual PO.xlsx, Sheet2
    -- Mapping:
    --   Company: KPC = 1000, KPD = 2000, KPT = 3000
    --   Fund: 'O' + Segment + '0'
    --   Month converted to zero-padded Otb_Month text
    INSERT INTO @Keys
        ([PO], [Otb_Year], [Otb_Month], [Company_Code], [Supplier], [Fund], [Category], [Brand])
    VALUES
        (N'2011010410', 2026, N'06', N'1000', N'1011544', N'O2000', N'241', N'AVE'),
        (N'2011010410', 2026, N'06', N'1000', N'1011544', N'O2000', N'241', N'MUN'),
        (N'2035011568', 2026, N'03', N'3000', N'1020771', N'O3000', N'381', N'RPF'),
        (N'2035011571', 2026, N'03', N'3000', N'1020771', N'O3000', N'381', N'RPF'),
        (N'2035011574', 2026, N'03', N'3000', N'1020771', N'O3000', N'381', N'RPF'),
        (N'2035011572', 2026, N'03', N'3000', N'1020771', N'O3000', N'381', N'RPF'),
        (N'2035011569', 2026, N'03', N'3000', N'1020771', N'O3000', N'381', N'RPF'),
        (N'2035011570', 2026, N'03', N'3000', N'1020771', N'O3000', N'381', N'RPF'),
        (N'2035011575', 2026, N'03', N'3000', N'1020771', N'O3000', N'381', N'RPF'),
        (N'2011010002', 2026, N'03', N'1000', N'1011173', N'O2000', N'201', N'MCM'),
        (N'2011010003', 2026, N'04', N'1000', N'1010271', N'O2000', N'211', N'BVN'),
        (N'2011010003', 2026, N'04', N'1000', N'1010271', N'O2000', N'211', N'FGR'),
        (N'2021013118', 2026, N'03', N'2000', N'1011567', N'O2000', N'223', N'CGF'),
        (N'2011010003', 2026, N'04', N'1000', N'1010271', N'O2000', N'211', N'GRA'),
        (N'2011010003', 2026, N'04', N'1000', N'1010271', N'O2000', N'211', N'MKY'),
        (N'2011010003', 2026, N'04', N'1000', N'1010271', N'O2000', N'211', N'HDK'),
        (N'2011010003', 2026, N'04', N'1000', N'1010271', N'O2000', N'211', N'GFD'),
        (N'2035011555', 2026, N'03', N'3000', N'1020771', N'O3000', N'381', N'RPF'),
        (N'2035011556', 2026, N'03', N'3000', N'1020771', N'O3000', N'381', N'RPF'),
        (N'2015015885', 2026, N'03', N'1000', N'1021571', N'O3000', N'351', N'APP'),
        (N'2011009995', 2026, N'03', N'1000', N'1020181', N'O2010', N'231', N'SWT');

    SELECT
        k.[PO],
        k.[Otb_Year],
        k.[Otb_Month],
        k.[Company_Code],
        k.[Supplier],
        k.[Fund],
        k.[Category],
        k.[Brand],
        COUNT(t.[PO]) AS [MatchedRows]
    FROM @Keys AS k
    LEFT JOIN [BMS].[dbo].[Actual_PO_Staging] AS t
        ON  t.[PO] = k.[PO]
        AND t.[Otb_Year] = k.[Otb_Year]
        AND t.[Otb_Month] = k.[Otb_Month]
        AND t.[Company_Code] = k.[Company_Code]
        AND t.[Supplier] = k.[Supplier]
        AND t.[Fund] = k.[Fund]
        AND t.[Category] = k.[Category]
        AND t.[Brand] = k.[Brand]
    GROUP BY
        k.[PO],
        k.[Otb_Year],
        k.[Otb_Month],
        k.[Company_Code],
        k.[Supplier],
        k.[Fund],
        k.[Category],
        k.[Brand]
    ORDER BY
        k.[PO],
        k.[Brand];

    UPDATE t
        SET t.[Deletion_Flag] = N'L'
    OUTPUT
        inserted.[PO],
        inserted.[PO_Item],
        inserted.[Otb_Year],
        inserted.[Otb_Month],
        inserted.[Company_Code],
        inserted.[Supplier],
        inserted.[Fund],
        inserted.[Category],
        inserted.[Brand],
        deleted.[Deletion_Flag] AS [Old_Deletion_Flag],
        inserted.[Deletion_Flag] AS [New_Deletion_Flag]
    FROM [BMS].[dbo].[Actual_PO_Staging] AS t
    INNER JOIN @Keys AS k
        ON  t.[PO] = k.[PO]
        AND t.[Otb_Year] = k.[Otb_Year]
        AND t.[Otb_Month] = k.[Otb_Month]
        AND t.[Company_Code] = k.[Company_Code]
        AND t.[Supplier] = k.[Supplier]
        AND t.[Fund] = k.[Fund]
        AND t.[Category] = k.[Category]
        AND t.[Brand] = k.[Brand]
    WHERE ISNULL(t.[Deletion_Flag], N'') <> N'L';

    SELECT @@ROWCOUNT AS [UpdatedRows];

    COMMIT TRAN;
END TRY
BEGIN CATCH
    IF @@TRANCOUNT > 0
        ROLLBACK TRAN;

    THROW;
END CATCH;
GO
