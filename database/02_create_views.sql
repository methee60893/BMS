USE [BMS];
GO

CREATE OR ALTER VIEW dbo.View_OTB_Draft
AS
SELECT
    d.RunNo,
    d.Vendor AS OTBVendor,
    ISNULL(v.Vendor, d.Vendor) AS Vendor,
    d.Company AS OTBCompany,
    ISNULL(co.CompanyNameShort, d.Company) AS CompanyName,
    d.[Month] AS OTBMonth,
    ISNULL(m.month_name_sh, CONVERT(nvarchar(20), d.[Month])) AS month_name_sh,
    d.Category AS OTBCategory,
    ISNULL(c.Category, d.Category) AS CateName,
    d.[Type] AS OTBType,
    d.[Year] AS OTBYear,
    d.Brand AS OTBBrand,
    ISNULL(b.[Brand Name], d.Brand) AS BrandName,
    ISNULL(s.SegmentName, ISNULL(d.SegmentCode, d.Segment)) AS SegmentName,
    ISNULL(d.SegmentCode, d.Segment) AS OTBSegment,
    d.Amount,
    d.Batch,
    d.CreateDT,
    d.UploadBy,
    d.Remark,
    d.[Version],
    ISNULL(d.OTBStatus, N'Draft') AS OTBStatus,
    d.SAPStatus,
    d.SAPErrorMessage
FROM dbo.Template_Upload_Draft_OTB d
LEFT JOIN dbo.MS_Month m
    ON d.[Month] = m.month_code
LEFT JOIN dbo.MS_Category c
    ON d.Category = c.Cate
LEFT JOIN dbo.MS_Company co
    ON d.Company = co.CompanyCode
LEFT JOIN dbo.MS_Segment s
    ON ISNULL(d.SegmentCode, d.Segment) = s.SegmentCode
LEFT JOIN dbo.MS_Brand b
    ON d.Brand = b.[Brand Code]
LEFT JOIN dbo.MS_Vendor v
    ON d.Vendor = v.VendorCode
   AND ISNULL(d.SegmentCode, d.Segment) = v.SegmentCode;
GO

CREATE OR ALTER VIEW dbo.View_UserRole
AS
SELECT
    u.UserID,
    u.Username,
    u.FullName,
    u.Email,
    u.IsActive,
    ur.RoleID,
    r.RoleName
FROM dbo.MS_User u
LEFT JOIN dbo.Map_User_Role ur
    ON u.UserID = ur.UserID
LEFT JOIN dbo.MS_Role r
    ON ur.RoleID = r.RoleID;
GO
