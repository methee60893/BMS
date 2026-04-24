USE [BMS];
GO

IF OBJECT_ID(N'dbo.MS_Segment', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.MS_Segment
    (
        SegmentCode nvarchar(20) NOT NULL,
        SegmentName nvarchar(200) NOT NULL,
        isActive bit NOT NULL CONSTRAINT DF_MS_Segment_isActive DEFAULT (1),
        CONSTRAINT PK_MS_Segment PRIMARY KEY CLUSTERED (SegmentCode)
    );
END;
GO

IF OBJECT_ID(N'dbo.MS_Year', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.MS_Year
    (
        Year_Kept int NOT NULL,
        CONSTRAINT PK_MS_Year PRIMARY KEY CLUSTERED (Year_Kept)
    );
END;
GO

IF OBJECT_ID(N'dbo.MS_Month', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.MS_Month
    (
        month_code int NOT NULL,
        month_name_sh nvarchar(20) NOT NULL,
        CONSTRAINT PK_MS_Month PRIMARY KEY CLUSTERED (month_code)
    );
END;
GO

IF OBJECT_ID(N'dbo.MS_Company', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.MS_Company
    (
        CompanyCode nvarchar(20) NOT NULL,
        CompanyNameShort nvarchar(200) NOT NULL,
        CONSTRAINT PK_MS_Company PRIMARY KEY CLUSTERED (CompanyCode)
    );
END;
GO

IF OBJECT_ID(N'dbo.MS_Category', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.MS_Category
    (
        Cate nvarchar(20) NOT NULL,
        Category nvarchar(200) NOT NULL,
        isActive bit NOT NULL CONSTRAINT DF_MS_Category_isActive DEFAULT (1),
        CONSTRAINT PK_MS_Category PRIMARY KEY CLUSTERED (Cate)
    );
END;
GO

IF OBJECT_ID(N'dbo.MS_Brand', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.MS_Brand
    (
        [Brand Code] nvarchar(30) NOT NULL,
        [Brand Name] nvarchar(200) NOT NULL,
        isActive bit NOT NULL CONSTRAINT DF_MS_Brand_isActive DEFAULT (1),
        CONSTRAINT PK_MS_Brand PRIMARY KEY CLUSTERED ([Brand Code])
    );
END;
GO

IF OBJECT_ID(N'dbo.MS_CCY', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.MS_CCY
    (
        CCY_Code nvarchar(10) NOT NULL,
        CCY_Name nvarchar(100) NOT NULL,
        isActive bit NOT NULL CONSTRAINT DF_MS_CCY_isActive DEFAULT (1),
        CONSTRAINT PK_MS_CCY PRIMARY KEY CLUSTERED (CCY_Code)
    );
END;
GO

IF OBJECT_ID(N'dbo.MS_Version', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.MS_Version
    (
        VersionCode nvarchar(20) NOT NULL,
        OTBTypeCode nvarchar(20) NOT NULL,
        Seq int NOT NULL,
        isActive bit NOT NULL CONSTRAINT DF_MS_Version_isActive DEFAULT (1),
        CONSTRAINT PK_MS_Version PRIMARY KEY CLUSTERED (VersionCode, OTBTypeCode)
    );
END;
GO

IF OBJECT_ID(N'dbo.MS_Vendor', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.MS_Vendor
    (
        VendorId int IDENTITY(1,1) NOT NULL,
        VendorCode nvarchar(30) NOT NULL,
        Vendor nvarchar(200) NOT NULL,
        CCY nvarchar(10) NULL,
        PaymentTermCode nvarchar(30) NULL,
        PaymentTerm nvarchar(100) NULL,
        SegmentCode nvarchar(20) NULL,
        Segment nvarchar(200) NULL,
        Incoterm nvarchar(50) NULL,
        isActive bit NOT NULL CONSTRAINT DF_MS_Vendor_isActive DEFAULT (1),
        CONSTRAINT PK_MS_Vendor PRIMARY KEY CLUSTERED (VendorId),
        CONSTRAINT UQ_MS_Vendor UNIQUE (VendorCode, SegmentCode, PaymentTermCode)
    );
END;
GO

IF OBJECT_ID(N'dbo.MS_User', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.MS_User
    (
        UserID int NOT NULL,
        Username nvarchar(255) NOT NULL,
        FullName nvarchar(255) NULL,
        Email nvarchar(255) NULL,
        IsActive bit NOT NULL CONSTRAINT DF_MS_User_IsActive DEFAULT (1),
        LastLoginDate datetime2(0) NULL,
        CreateDate datetime2(0) NULL,
        CONSTRAINT PK_MS_User PRIMARY KEY CLUSTERED (UserID),
        CONSTRAINT UQ_MS_User_Username UNIQUE (Username)
    );
END;
GO

IF OBJECT_ID(N'dbo.MS_Role', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.MS_Role
    (
        RoleID int NOT NULL,
        RoleName nvarchar(100) NOT NULL,
        [Description] nvarchar(255) NULL,
        CONSTRAINT PK_MS_Role PRIMARY KEY CLUSTERED (RoleID),
        CONSTRAINT UQ_MS_Role_RoleName UNIQUE (RoleName)
    );
END;
GO

IF OBJECT_ID(N'dbo.MS_Menu', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.MS_Menu
    (
        MenuID int NOT NULL,
        MenuName nvarchar(150) NOT NULL,
        PageUrl nvarchar(255) NOT NULL,
        MenuGroup nvarchar(100) NULL,
        SortOrder int NOT NULL CONSTRAINT DF_MS_Menu_SortOrder DEFAULT (0),
        IsActive bit NOT NULL CONSTRAINT DF_MS_Menu_IsActive DEFAULT (1),
        CONSTRAINT PK_MS_Menu PRIMARY KEY CLUSTERED (MenuID),
        CONSTRAINT UQ_MS_Menu_PageUrl UNIQUE (PageUrl)
    );
END;
GO

IF OBJECT_ID(N'dbo.Map_User_Role', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.Map_User_Role
    (
        MapID int NOT NULL,
        UserID int NOT NULL,
        RoleID int NOT NULL,
        CONSTRAINT PK_Map_User_Role PRIMARY KEY CLUSTERED (MapID)
    );
END;
GO

IF OBJECT_ID(N'dbo.Map_Role_Permission', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.Map_Role_Permission
    (
        PermissionID int NOT NULL,
        RoleID int NOT NULL,
        MenuID int NOT NULL,
        CanView bit NOT NULL CONSTRAINT DF_Map_Role_Permission_CanView DEFAULT (0),
        CanEdit bit NOT NULL CONSTRAINT DF_Map_Role_Permission_CanEdit DEFAULT (0),
        CanDelete bit NOT NULL CONSTRAINT DF_Map_Role_Permission_CanDelete DEFAULT (0),
        CanApprove bit NOT NULL CONSTRAINT DF_Map_Role_Permission_CanApprove DEFAULT (0),
        CONSTRAINT PK_Map_Role_Permission PRIMARY KEY CLUSTERED (PermissionID)
    );
END;
GO

IF OBJECT_ID(N'dbo.Template_Upload_Draft_OTB', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.Template_Upload_Draft_OTB
    (
        RunNo int IDENTITY(1,1) NOT NULL,
        [Type] nvarchar(20) NOT NULL,
        [Year] int NOT NULL,
        [Month] int NOT NULL,
        Category nvarchar(20) NOT NULL,
        Company nvarchar(20) NOT NULL,
        Segment nvarchar(20) NOT NULL,
        SegmentCode nvarchar(20) NULL,
        Brand nvarchar(30) NOT NULL,
        Vendor nvarchar(30) NOT NULL,
        Amount decimal(18,2) NOT NULL,
        Batch nvarchar(50) NULL,
        UploadBy nvarchar(100) NULL,
        CreateDT datetime2(0) NOT NULL CONSTRAINT DF_Template_Upload_Draft_OTB_CreateDT DEFAULT (sysdatetime()),
        UpdateBy nvarchar(100) NULL,
        UpdateDT datetime2(0) NULL,
        Remark nvarchar(500) NULL,
        [Version] nvarchar(20) NULL,
        OTBStatus nvarchar(30) NULL CONSTRAINT DF_Template_Upload_Draft_OTB_Status DEFAULT (N'Draft'),
        SAPStatus nvarchar(10) NULL,
        SAPErrorMessage nvarchar(1000) NULL,
        CONSTRAINT PK_Template_Upload_Draft_OTB PRIMARY KEY CLUSTERED (RunNo)
    );
END;
GO

IF OBJECT_ID(N'dbo.OTB_Transaction', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.OTB_Transaction
    (
        OTBTransID bigint IDENTITY(1,1) NOT NULL,
        CreateDate datetime2(0) NOT NULL CONSTRAINT DF_OTB_Transaction_CreateDate DEFAULT (sysdatetime()),
        [Type] nvarchar(20) NOT NULL,
        [Year] int NOT NULL,
        [Month] int NOT NULL,
        Category nvarchar(20) NOT NULL,
        CategoryName nvarchar(200) NULL,
        Company nvarchar(20) NOT NULL,
        Segment nvarchar(20) NOT NULL,
        SegmentName nvarchar(200) NULL,
        Brand nvarchar(30) NOT NULL,
        BrandName nvarchar(200) NULL,
        Vendor nvarchar(30) NOT NULL,
        VendorName nvarchar(200) NULL,
        Amount decimal(18,2) NOT NULL,
        RevisedDiff decimal(18,2) NULL,
        Remark nvarchar(500) NULL,
        OTBStatus nvarchar(30) NOT NULL CONSTRAINT DF_OTB_Transaction_Status DEFAULT (N'Approved'),
        ApprovedDate datetime2(0) NULL,
        SAPDate datetime2(0) NULL,
        ActionBy nvarchar(100) NULL,
        DraftID int NULL,
        SAPStatus nvarchar(10) NULL,
        SAPErrorMessage nvarchar(1000) NULL,
        [Version] nvarchar(20) NOT NULL,
        CONSTRAINT PK_OTB_Transaction PRIMARY KEY CLUSTERED (OTBTransID),
        CONSTRAINT UQ_OTB_Transaction UNIQUE ([Type], [Year], [Month], Category, Company, Segment, Brand, Vendor, [Version])
    );
END;
GO

IF OBJECT_ID(N'dbo.OTB_Switching_Transaction', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.OTB_Switching_Transaction
    (
        SwitchingID bigint IDENTITY(1,1) NOT NULL,
        [Year] int NOT NULL,
        [Month] int NOT NULL,
        Company nvarchar(20) NOT NULL,
        Category nvarchar(20) NOT NULL,
        Segment nvarchar(20) NOT NULL,
        Brand nvarchar(30) NOT NULL,
        Vendor nvarchar(30) NOT NULL,
        [From] nvarchar(5) NOT NULL,
        BudgetAmount decimal(18,2) NOT NULL,
        [Release] decimal(18,2) NOT NULL CONSTRAINT DF_OTB_Switching_Release DEFAULT (0),
        SwitchYear int NULL,
        SwitchMonth int NULL,
        SwitchCompany nvarchar(20) NULL,
        SwitchCategory nvarchar(20) NULL,
        SwitchSegment nvarchar(20) NULL,
        [To] nvarchar(5) NULL,
        SwitchBrand nvarchar(30) NULL,
        SwitchVendor nvarchar(30) NULL,
        OTBStatus nvarchar(30) NOT NULL CONSTRAINT DF_OTB_Switching_Status DEFAULT (N'Approved'),
        Batch nvarchar(50) NULL,
        Remark nvarchar(500) NULL,
        CreateBy nvarchar(100) NULL,
        CreateDT datetime2(0) NOT NULL CONSTRAINT DF_OTB_Switching_CreateDT DEFAULT (sysdatetime()),
        ActionBy nvarchar(100) NULL,
        CONSTRAINT PK_OTB_Switching_Transaction PRIMARY KEY CLUSTERED (SwitchingID)
    );
END;
GO

IF OBJECT_ID(N'dbo.Draft_PO_Transaction', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.Draft_PO_Transaction
    (
        DraftPO_ID bigint IDENTITY(1,1) NOT NULL,
        DraftPO_No nvarchar(50) NOT NULL,
        PO_Year int NOT NULL,
        PO_Month int NOT NULL,
        Company_Code nvarchar(20) NOT NULL,
        Category_Code nvarchar(20) NOT NULL,
        Segment_Code nvarchar(20) NOT NULL,
        Brand_Code nvarchar(30) NOT NULL,
        Vendor_Code nvarchar(30) NOT NULL,
        CCY nvarchar(10) NOT NULL,
        Exchange_Rate decimal(18,6) NOT NULL,
        Amount_CCY decimal(18,2) NOT NULL,
        Amount_THB decimal(18,2) NOT NULL,
        PO_Type nvarchar(20) NOT NULL CONSTRAINT DF_Draft_PO_Transaction_POType DEFAULT (N'Manual'),
        [Status] nvarchar(30) NOT NULL CONSTRAINT DF_Draft_PO_Transaction_Status DEFAULT (N'Draft'),
        Status_Date datetime2(0) NULL,
        Status_By nvarchar(100) NULL,
        Actual_PO_Ref nvarchar(50) NULL,
        Actual_PO_No nvarchar(50) NULL,
        Remark nvarchar(500) NULL,
        Created_By nvarchar(100) NULL,
        Created_Date datetime2(0) NOT NULL CONSTRAINT DF_Draft_PO_Transaction_CreatedDate DEFAULT (sysdatetime()),
        CONSTRAINT PK_Draft_PO_Transaction PRIMARY KEY CLUSTERED (DraftPO_ID)
    );
END;
GO

IF OBJECT_ID(N'dbo.Actual_PO_Staging', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.Actual_PO_Staging
    (
        StagingID bigint IDENTITY(1,1) NOT NULL,
        PO nvarchar(50) NOT NULL,
        PO_Item nvarchar(20) NOT NULL,
        Otb_Year int NULL,
        Otb_Month int NULL,
        Company_Code nvarchar(20) NULL,
        Supplier nvarchar(30) NULL,
        Fund nvarchar(20) NULL,
        Category nvarchar(20) NULL,
        Brand nvarchar(30) NULL,
        Otb_Date datetime2(0) NULL,
        Supplier_Name nvarchar(200) NULL,
        Fund_Name nvarchar(200) NULL,
        Brand_Name nvarchar(200) NULL,
        Category_Name nvarchar(200) NULL,
        PO_Amount decimal(18,2) NULL,
        PO_Currency nvarchar(10) NULL,
        PO_Local_Amount decimal(18,2) NULL,
        PO_Local_Currency nvarchar(10) NULL,
        Exchange_Rate decimal(18,6) NULL,
        Deletion_Flag nvarchar(5) NULL,
        Delivery_Completed_Flag bit NULL,
        Final_Invoice_Flag bit NULL,
        Create_On datetime2(0) NULL,
        Change_On datetime2(0) NULL,
        Modified_Date datetime2(0) NULL,
        BMS_Last_Synced datetime2(0) NULL,
        CONSTRAINT PK_Actual_PO_Staging PRIMARY KEY CLUSTERED (StagingID)
    );
END;
GO

IF OBJECT_ID(N'dbo.Actual_PO_Summary', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.Actual_PO_Summary
    (
        ActualPO_ID bigint IDENTITY(1,1) NOT NULL,
        PO_No nvarchar(50) NOT NULL,
        PO_Item nvarchar(20) NOT NULL,
        OTB_Year int NULL,
        OTB_Month int NULL,
        Company_Code nvarchar(20) NULL,
        Category_Code nvarchar(20) NULL,
        Segment_Code nvarchar(20) NULL,
        Brand_Code nvarchar(30) NULL,
        Vendor_Code nvarchar(30) NULL,
        Actual_PO_Date datetime2(0) NULL,
        Amount_THB decimal(18,2) NULL,
        Amount_CCY decimal(18,2) NULL,
        CCY nvarchar(10) NULL,
        Exchange_Rate decimal(18,6) NULL,
        Draft_PO_Ref nvarchar(50) NULL,
        [Status] nvarchar(30) NULL,
        Matching_Date datetime2(0) NULL,
        Changed_By nvarchar(100) NULL,
        Changed_Date datetime2(0) NULL,
        Remark nvarchar(500) NULL,
        Source_Modified_Date datetime2(0) NULL,
        BMS_Last_Synced datetime2(0) NULL,
        CONSTRAINT PK_Actual_PO_Summary PRIMARY KEY CLUSTERED (ActualPO_ID),
        CONSTRAINT UQ_Actual_PO_Summary UNIQUE (PO_No, PO_Item)
    );
END;
GO

IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = N'IX_Template_Upload_Draft_OTB_Search' AND object_id = OBJECT_ID(N'dbo.Template_Upload_Draft_OTB'))
BEGIN
    CREATE NONCLUSTERED INDEX IX_Template_Upload_Draft_OTB_Search
        ON dbo.Template_Upload_Draft_OTB ([Type], [Year], [Month], Company, Category, Segment, Brand, Vendor)
        INCLUDE (Amount, OTBStatus, [Version], CreateDT);
END;
GO

IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = N'IX_OTB_Transaction_Search' AND object_id = OBJECT_ID(N'dbo.OTB_Transaction'))
BEGIN
    CREATE NONCLUSTERED INDEX IX_OTB_Transaction_Search
        ON dbo.OTB_Transaction ([Year], [Month], Company, Category, Segment, Brand, Vendor, OTBStatus)
        INCLUDE ([Type], Amount, RevisedDiff, [Version], ApprovedDate);
END;
GO

IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = N'IX_OTB_Switching_Search' AND object_id = OBJECT_ID(N'dbo.OTB_Switching_Transaction'))
BEGIN
    CREATE NONCLUSTERED INDEX IX_OTB_Switching_Search
        ON dbo.OTB_Switching_Transaction ([Year], [Month], Company, Category, Segment, Brand, Vendor, OTBStatus)
        INCLUDE ([From], [To], BudgetAmount, SwitchYear, SwitchMonth, SwitchCompany, SwitchCategory, SwitchSegment, SwitchBrand, SwitchVendor, CreateDT);
END;
GO

IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = N'IX_Draft_PO_Transaction_Search' AND object_id = OBJECT_ID(N'dbo.Draft_PO_Transaction'))
BEGIN
    CREATE NONCLUSTERED INDEX IX_Draft_PO_Transaction_Search
        ON dbo.Draft_PO_Transaction (PO_Year, PO_Month, Company_Code, Category_Code, Segment_Code, Brand_Code, Vendor_Code, [Status])
        INCLUDE (DraftPO_No, Amount_THB, Amount_CCY, CCY, Exchange_Rate, Actual_PO_Ref, Created_Date);
END;
GO

IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = N'IX_Actual_PO_Staging_Key' AND object_id = OBJECT_ID(N'dbo.Actual_PO_Staging'))
BEGIN
    CREATE NONCLUSTERED INDEX IX_Actual_PO_Staging_Key
        ON dbo.Actual_PO_Staging (PO, PO_Item, Otb_Year, Otb_Month, Company_Code, Supplier, Fund, Category, Brand)
        INCLUDE (PO_Local_Amount, PO_Amount, PO_Currency, PO_Local_Currency, Exchange_Rate, Deletion_Flag, Modified_Date, BMS_Last_Synced);
END;
GO

IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = N'IX_Actual_PO_Summary_Search' AND object_id = OBJECT_ID(N'dbo.Actual_PO_Summary'))
BEGIN
    CREATE NONCLUSTERED INDEX IX_Actual_PO_Summary_Search
        ON dbo.Actual_PO_Summary (OTB_Year, OTB_Month, Company_Code, Category_Code, Segment_Code, Brand_Code, Vendor_Code, [Status])
        INCLUDE (PO_No, Amount_THB, Amount_CCY, CCY, Exchange_Rate, Draft_PO_Ref, Actual_PO_Date);
END;
GO
