Imports System
Imports System.Configuration
Imports System.Data
Imports System.Data.Odbc
Imports System.Data.SqlClient
Imports System.Globalization
Imports System.IO
Imports System.Runtime.InteropServices
Imports System.Text
Imports System.Web
Imports System.Web.Script.Serialization
Imports ExcelDataReader
Imports Newtonsoft.Json
Imports OfficeOpenXml.FormulaParsing.Ranges

Public Class MasterDataHandler
    Implements IHttpHandler

    Private Shared connectionString As String = ConfigurationManager.ConnectionStrings("BMSConnectionString")?.ConnectionString

    Sub ProcessRequest(ByVal context As HttpContext) Implements IHttpHandler.ProcessRequest

        context.Response.Clear()
        context.Response.ContentType = "text/html"
        context.Response.ContentEncoding = Encoding.UTF8

        Try
            Dim dt As DataTable = Nothing


            If context.Request("action") = "SegmentList" Then
                context.Response.Write(GetSegmentList())
            ElseIf context.Request("action") = "YearList" Then
                context.Response.Write(GetYearList())
            ElseIf context.Request("action") = "MonthList" Then
                context.Response.Write(GetMonthList())
            ElseIf context.Request("action") = "CompanyList" Then
                context.Response.Write(GetCompanyList())
            ElseIf context.Request("action") = "CategoryList" Then
                context.Response.Write(GetCategoryList())
            ElseIf context.Request("action") = "BrandList" Then
                context.Response.Write(GetBrandList())
            ElseIf context.Request("action") = "VendorList" Then
                context.Response.Write(GetVendorList())
            ElseIf context.Request("action") = "VendorListChg" Then
                Dim segmentCode As String = If(String.IsNullOrWhiteSpace(context.Request.Form("segmentCode")),
                                   "",
                                   context.Request.Form("segmentCode").Trim())
                context.Response.Write(GetVendorListwithfilter(segmentCode))
            ElseIf context.Request("action") = "SegmentMSList" Then
                context.Response.Write(GetMSSegmentList())
            ElseIf context.Request("action") = "YearMSList" Then
                context.Response.Write(GetMSYearList())
            ElseIf context.Request("action") = "MonthMSList" Then
                context.Response.Write(GetMSMonthList())
            ElseIf context.Request("action") = "CompanyMSList" Then
                context.Response.Write(GetMSCompanyList())
            ElseIf context.Request("action") = "CategoryMSList" Then
                context.Response.Write(GetMSCategoryList())
            ElseIf context.Request("action") = "BrandMSList" Then
                context.Response.Write(GetMSBrandList())
            ElseIf context.Request("action") = "VendorMSList" Then
                context.Response.Write(GetMSVendorList())
            ElseIf context.Request("action") = "CCYMSList" Then
                context.Response.Write(GetMSCCYListFromMainMaster())
            ElseIf context.Request("action") = "VersionMSList" Then
                Dim typeCode As String = If(String.IsNullOrWhiteSpace(context.Request.Form("OTBtype")),
                                       "",
                                       context.Request.Form("OTBtype").Trim())
                context.Response.Write(GetMSVersionList(typeCode))
            ElseIf context.Request("action") = "VendorMSListChg" Then
                Dim segmentCode As String = If(String.IsNullOrWhiteSpace(context.Request.Form("segmentCode")),
                                       "",
                                       context.Request.Form("segmentCode").Trim())
                context.Response.Write(GetMSVendorListwithfilter(segmentCode))
            ElseIf context.Request("action") = "CCYMSListChg" Then
                Dim vendorCode As String = If(String.IsNullOrWhiteSpace(context.Request.Form("vendorCode")),
                                       "",
                                       context.Request.Form("vendorCode").Trim())
                context.Response.Write(GetMSCCYListwithfilter(vendorCode))
            ElseIf context.Request("action") = "getvendormslist_json" Then
                HandleVendorSearch(context)
            ElseIf context.Request("action") = "getvendorbysegment_json" Then
                HandleVendorSearchBySegment(context)
            ElseIf context.Request("action") = "getVendorListHtml" Then
                GetVendorListHtml(context)
            ElseIf context.Request("action") = "saveVendor" Then
                SaveVendor(context)
            ElseIf context.Request("action") = "deleteVendor" Then
                DeleteVendor(context)

                ' ===== START: NEW CATEGORY ACTIONS =====
            ElseIf context.Request("action") = "getcategorylisthtml" Then
                GetCategoryListHtml(context)
            ElseIf context.Request("action") = "savecategory" Then
                SaveCategory(context)
            ElseIf context.Request("action") = "deletecategory" Then
                DeleteCategory(context)
                ' ===== END: NEW CATEGORY ACTIONS =====

                ' ===== START: NEW BRAND ACTIONS =====
            ElseIf context.Request("action") = "getbrandlisthtml" Then
                GetBrandListHtml(context)
            ElseIf context.Request("action") = "savebrand" Then
                SaveBrand(context)
            ElseIf context.Request("action") = "deletebrand" Then
                DeleteBrand(context)
                ' ===== END: NEW BRAND ACTIONS =====
            End If

        Catch ex As Exception
            context.Response.StatusCode = 500
            context.Response.Write($"<div class='alert alert-danger'>Error: {HttpUtility.HtmlEncode(ex.Message)}</div>")
        End Try
    End Sub

    Private Sub HandleVendorSearch(context As HttpContext)
        Dim searchTerm As String = If(context.Request("search"), "").Trim()
        Dim dt As DataTable = GetVendorDataAsJSON(searchTerm)
        ReturnSelect2Json(context, dt)
    End Sub

    Private Sub HandleVendorSearchBySegment(context As HttpContext)
        Dim searchTerm As String = If(context.Request("search"), "").Trim()
        Dim segmentCode As String = If(context.Request("segmentCode"), "").Trim()
        Dim dt As DataTable = GetVendorDataBySegmentAsJSON(searchTerm, segmentCode)
        ReturnSelect2Json(context, dt)
    End Sub

    ' (เพิ่ม) Helper ดึงข้อมูล Vendor (รองรับการค้นหาและ Paging)
    Private Function GetVendorDataAsJSON(searchTerm As String, Optional segmentCode As String = "") As DataTable
        Dim dt As New DataTable()
        ' (ปรับ Query ให้รองรับการค้นหา (LIKE) และ Paging (OFFSET/FETCH) เพื่อประสิทธิภาพ)
        Dim query As String = "
            SELECT [VendorCode], [Vendor] AS VendorName 
            FROM [MS_Vendor] 
            WHERE 1=1"

        If Not String.IsNullOrEmpty(segmentCode) Then
            query &= " AND [SegmentCode] = @SegmentCode"
        End If

        If Not String.IsNullOrEmpty(searchTerm) Then
            query &= " AND ([VendorCode] LIKE @Search OR [Vendor] LIKE @Search)"
        End If

        query &= " ORDER BY [VendorCode] OFFSET 0 ROWS FETCH NEXT 50 ROWS ONLY" ' (Paging: ดึงทีละ 50)

        Using conn As New SqlConnection(connectionString)
            Using cmd As New SqlCommand(query, conn)
                If Not String.IsNullOrEmpty(segmentCode) Then
                    cmd.Parameters.AddWithValue("@SegmentCode", segmentCode)
                End If
                If Not String.IsNullOrEmpty(searchTerm) Then
                    cmd.Parameters.AddWithValue("@Search", "%" & searchTerm & "%")
                End If
                conn.Open()
                Using reader As SqlDataReader = cmd.ExecuteReader()
                    dt.Load(reader)
                End Using
            End Using
        End Using
        Return dt
    End Function

    Private Function GetVendorDataBySegmentAsJSON(searchTerm As String, Optional segmentCode As String = "") As DataTable
        Dim dt As New DataTable()
        ' (ปรับ Query ให้รองรับการค้นหา (LIKE) และ Paging (OFFSET/FETCH) เพื่อประสิทธิภาพ)
        Dim query As String = "
            SELECT [SegmentCode], [SegmentName]
            FROM [MS_Segment] 
            WHERE 1=1"

        If Not String.IsNullOrEmpty(segmentCode) Then
            query &= " AND [SegmentCode] = @SegmentCode"
        End If

        If Not String.IsNullOrEmpty(searchTerm) Then
            query &= " AND ([SegmentCode] LIKE @Search OR [SegmentName] LIKE @Search)"
        End If

        query &= " ORDER BY [SegmentCode] OFFSET 0 ROWS FETCH NEXT 50 ROWS ONLY" ' (Paging: ดึงทีละ 50)

        Using conn As New SqlConnection(connectionString)
            Using cmd As New SqlCommand(query, conn)
                If Not String.IsNullOrEmpty(segmentCode) Then
                    cmd.Parameters.AddWithValue("@SegmentCode", segmentCode)
                End If
                If Not String.IsNullOrEmpty(searchTerm) Then
                    cmd.Parameters.AddWithValue("@Search", "%" & searchTerm & "%")
                End If
                conn.Open()
                Using reader As SqlDataReader = cmd.ExecuteReader()
                    dt.Load(reader)
                End Using
            End Using
        End Using
        Return dt
    End Function

    ' (เพิ่ม) Helper แปลง DataTable เป็น JSON ที่ Select2 ต้องการ
    Private Sub ReturnSelect2Json(context As HttpContext, dt As DataTable)
        Dim results = New List(Of Object)()
        For Each row As DataRow In dt.Rows
            results.Add(New With {
                .id = row("VendorCode").ToString(),
                .text = row("VendorCode").ToString() & " - " & row("VendorName").ToString()
            })
        Next

        context.Response.ContentType = "application/json"
        context.Response.Write(JsonConvert.SerializeObject(New With {.results = results}))
    End Sub

    Private Function GetSegmentList() As String
        Dim dt As New DataTable()
        Using conn As New SqlConnection(connectionString)
            conn.Open()
            Dim query As String = "SELECT DISTINCT [dbo].[MS_Segment].SegmentName, [dbo].[MS_Segment].SegmentCode FROM [BMS].[dbo].[Template_Upload_Draft_OTB]
                                   INNER JOIN [dbo].[MS_Segment] ON  [BMS].[dbo].[Template_Upload_Draft_OTB].[Segment] = [dbo].[MS_Segment].SegmentCode
                                   ORDER BY [dbo].[MS_Segment].SegmentCode ASC
                                    "
            Using cmd As New SqlCommand(query, conn)
                Using reader As SqlDataReader = cmd.ExecuteReader()
                    dt.Load(reader)
                End Using
            End Using
        End Using
        Return GenerateHtmlSegmentDropdown(dt)
    End Function

    Private Function GetMSVersionList(Optional OTBTypeCode As String = "Original") As String
        Dim dt As New DataTable()
        Using conn As New SqlConnection(connectionString)
            conn.Open()
            Dim query As String = "SELECT  [VersionCode]
                                          ,[OTBTypeCode]
                                          ,[Seq]
                                      FROM [BMS].[dbo].[MS_Version]
                                        WHERE [OTBTypeCode] = @OTBTypeCode
                                       ORDER BY [Seq] ASC
                                    "
            Using cmd As New SqlCommand(query, conn)
                cmd.Parameters.AddWithValue("@OTBTypeCode", OTBTypeCode)
                Using reader As SqlDataReader = cmd.ExecuteReader()
                    dt.Load(reader)
                End Using
            End Using
        End Using
        Return GenerateHtmlVersionDropdown(dt)
    End Function

    Private Function GetYearList() As String
        Dim dt As New DataTable()
        Using conn As New SqlConnection(connectionString)
            conn.Open()
            Dim query As String = "SELECT DISTINCT [Year] FROM [BMS].[dbo].[Template_Upload_Draft_OTB]"
            Using cmd As New SqlCommand(query, conn)
                Using reader As SqlDataReader = cmd.ExecuteReader()
                    dt.Load(reader)
                End Using
            End Using
        End Using
        Return GenerateHtmlYearDropdown(dt)
    End Function

    Private Function GetMonthList() As String
        Dim dt As New DataTable()
        Using conn As New SqlConnection(connectionString)
            conn.Open()
            Dim query As String = "SELECT DISTINCT [dbo].[MS_Month].[month_name_sh],[dbo].[MS_Month].[month_code] FROM [BMS].[dbo].[Template_Upload_Draft_OTB]
                                    INNER JOIN [dbo].[MS_Month] ON  [BMS].[dbo].[Template_Upload_Draft_OTB].[Month] =  [dbo].[MS_Month].[month_code]
                                    ORDER BY [dbo].[MS_Month].[month_code] ASC
                                    "
            Using cmd As New SqlCommand(query, conn)
                Using reader As SqlDataReader = cmd.ExecuteReader()
                    dt.Load(reader)
                End Using
            End Using
        End Using
        Return GenerateHtmlMonthDropdown(dt)
    End Function

    Private Function GetCompanyList() As String
        Dim dt As New DataTable()
        Using conn As New SqlConnection(connectionString)
            conn.Open()
            Dim query As String = "SELECT DISTINCT [dbo].[MS_Company].[CompanyNameShort], [dbo].[MS_Company].[CompanyCode] FROM [BMS].[dbo].[Template_Upload_Draft_OTB]
                                    INNER JOIN [dbo].[MS_Company] ON  [BMS].[dbo].[Template_Upload_Draft_OTB].[Company] = [dbo].[MS_Company].[CompanyCode]
                                    ORDER BY [dbo].[MS_Company].[CompanyCode] ASC
                                    "
            Using cmd As New SqlCommand(query, conn)
                Using reader As SqlDataReader = cmd.ExecuteReader()
                    dt.Load(reader)
                End Using
            End Using
        End Using
        Return GenerateHtmlCompanyDropdown(dt)
    End Function

    Private Function GetCategoryList() As String
        Dim dt As New DataTable()
        Using conn As New SqlConnection(connectionString)
            conn.Open()
            Dim query As String = "SELECT DISTINCT [dbo].[MS_Category].[Category], [dbo].[MS_Category].Cate FROM [BMS].[dbo].[Template_Upload_Draft_OTB]
                                    INNER JOIN [dbo].[MS_Category] ON  [BMS].[dbo].[Template_Upload_Draft_OTB].[Category] =  [dbo].[MS_Category].Cate
                                    WHERE [isActive] = 1        
                                    ORDER BY [dbo].[MS_Category].Cate ASC
                                    "
            Using cmd As New SqlCommand(query, conn)
                Using reader As SqlDataReader = cmd.ExecuteReader()
                    dt.Load(reader)
                End Using
            End Using
        End Using
        Return GenerateHtmlCategoryDropdown(dt)
    End Function

    Private Function GetBrandList() As String
        Dim dt As New DataTable()
        Using conn As New SqlConnection(connectionString)
            conn.Open()
            Dim query As String = "SELECT DISTINCT [dbo].[MS_Brand].[Brand Name], [dbo].[MS_Brand].[Brand Code] FROM [BMS].[dbo].[Template_Upload_Draft_OTB]
                                    INNER JOIN [dbo].[MS_Brand] ON  [BMS].[dbo].[Template_Upload_Draft_OTB].[Brand] = [dbo].[MS_Brand].[Brand Code]
                                    WHERE [isActive] = 1
                                    ORDER BY [dbo].[MS_Brand].[Brand Code] ASC
                                    " ' ปรับเปลี่ยนตามตารางจริง
            Using cmd As New SqlCommand(query, conn)
                Using reader As SqlDataReader = cmd.ExecuteReader()
                    dt.Load(reader)
                End Using
            End Using
        End Using
        Return GenerateHtmlBrandDropdown(dt)
    End Function

    Private Function GetVendorList() As String
        Dim dt As New DataTable()
        Using conn As New SqlConnection(connectionString)
            conn.Open()
            Dim query As String = "SELECT DISTINCT [dbo].[MS_Vendor].[VendorCode], [dbo].[MS_Vendor].[Vendor] FROM [BMS].[dbo].[Template_Upload_Draft_OTB]
                                    INNER JOIN [dbo].[MS_Vendor] ON  [BMS].[dbo].[Template_Upload_Draft_OTB].[Vendor] = [dbo].[MS_Vendor].[VendorCode]
                                    WHERE [isActive] = 1
                                    ORDER BY [dbo].[MS_Vendor].[VendorCode] ASC
                                    "
            Using cmd As New SqlCommand(query, conn)
                Using reader As SqlDataReader = cmd.ExecuteReader()
                    dt.Load(reader)
                End Using
            End Using
        End Using
        Return GenerateHtmlVendorDropdown(dt)
    End Function

    Private Function GetVendorListwithfilter(segmentCode As String) As String
        Dim dt As New DataTable()
        Using conn As New SqlConnection(connectionString)
            conn.Open()
            Dim query As String = "SELECT DISTINCT  [dbo].[MS_Vendor].[VendorCode], [dbo].[MS_Vendor].[Vendor] FROM [BMS].[dbo].[Template_Upload_Draft_OTB]
                                    INNER JOIN [dbo].[MS_Vendor] ON  [BMS].[dbo].[Template_Upload_Draft_OTB].[Vendor] = [dbo].[MS_Vendor].[VendorCode]
                                    WHERE [isActive] = 1 AND [BMS].[dbo].[Template_Upload_Draft_OTB].[SegmentCode] = @segmentCode
                                    ORDER BY [dbo].[MS_Vendor].[VendorCode] ASC
                                    "
            Using cmd As New SqlCommand(query, conn)
                cmd.Parameters.AddWithValue("@segmentCode", segmentCode)
                Using reader As SqlDataReader = cmd.ExecuteReader()
                    dt.Load(reader)
                End Using
            End Using
        End Using
        Return GenerateHtmlVendorDropdown(dt)
    End Function

    Private Function GetMSSegmentList() As String
        Dim dt As New DataTable()
        Using conn As New SqlConnection(connectionString)
            conn.Open()
            Dim query As String = "SELECT [dbo].[MS_Segment].SegmentName, [dbo].[MS_Segment].SegmentCode FROM [dbo].[MS_Segment]
                                    "
            Using cmd As New SqlCommand(query, conn)
                Using reader As SqlDataReader = cmd.ExecuteReader()
                    dt.Load(reader)
                End Using
            End Using
        End Using
        Return GenerateHtmlSegmentDropdown(dt)
    End Function

    Private Function GetMSYearList() As String
        Dim dt As New DataTable()
        Using conn As New SqlConnection(connectionString)
            conn.Open()
            Dim query As String = "SELECT DISTINCT [Year_Kept] AS 'Year' FROM [BMS].[dbo].[MS_Year]"
            Using cmd As New SqlCommand(query, conn)
                Using reader As SqlDataReader = cmd.ExecuteReader()
                    dt.Load(reader)
                End Using
            End Using
        End Using
        Return GenerateHtmlYearDropdown(dt)
    End Function

    Private Function GetMSMonthList() As String
        Dim dt As New DataTable()
        Using conn As New SqlConnection(connectionString)
            conn.Open()
            Dim query As String = "SELECT [dbo].[MS_Month].[month_name_sh],[dbo].[MS_Month].[month_code] FROM [dbo].[MS_Month]
                                    ORDER BY [dbo].[MS_Month].[month_code] ASC
                                    "
            Using cmd As New SqlCommand(query, conn)
                Using reader As SqlDataReader = cmd.ExecuteReader()
                    dt.Load(reader)
                End Using
            End Using
        End Using
        Return GenerateHtmlMonthDropdown(dt)
    End Function

    Private Function GetMSCompanyList() As String
        Dim dt As New DataTable()
        Using conn As New SqlConnection(connectionString)
            conn.Open()
            Dim query As String = "SELECT  [dbo].[MS_Company].[CompanyNameShort], [dbo].[MS_Company].[CompanyCode] FROM [dbo].[MS_Company] 
                                    ORDER BY [dbo].[MS_Company].[CompanyCode] ASC
                                    "
            Using cmd As New SqlCommand(query, conn)
                Using reader As SqlDataReader = cmd.ExecuteReader()
                    dt.Load(reader)
                End Using
            End Using
        End Using
        Return GenerateHtmlCompanyDropdown(dt)
    End Function

    Private Function GetMSCategoryList() As String
        Dim dt As New DataTable()
        Using conn As New SqlConnection(connectionString)
            conn.Open()
            Dim query As String = "SELECT  [dbo].[MS_Category].[Category], [dbo].[MS_Category].Cate FROM [dbo].[MS_Category]  WHERE [isActive] = 1
                                    ORDER BY [dbo].[MS_Category].Cate ASC
                                    "
            Using cmd As New SqlCommand(query, conn)
                Using reader As SqlDataReader = cmd.ExecuteReader()
                    dt.Load(reader)
                End Using
            End Using
        End Using
        Return GenerateHtmlCategoryDropdown(dt)
    End Function

    Private Function GetMSBrandList() As String
        Dim dt As New DataTable()
        Using conn As New SqlConnection(connectionString)
            conn.Open()
            Dim query As String = "SELECT [dbo].[MS_Brand].[Brand Name], [dbo].[MS_Brand].[Brand Code] FROM  [dbo].[MS_Brand]  WHERE [isActive] = 1
                                    ORDER BY [dbo].[MS_Brand].[Brand Code] ASC
                                    " ' ปรับเปลี่ยนตามตารางจริง
            Using cmd As New SqlCommand(query, conn)
                Using reader As SqlDataReader = cmd.ExecuteReader()
                    dt.Load(reader)
                End Using
            End Using
        End Using
        Return GenerateHtmlBrandDropdown(dt)
    End Function

    Private Function GetMSVendorList() As String
        Dim dt As New DataTable()
        Using conn As New SqlConnection(connectionString)
            conn.Open()
            Dim query As String = "SELECT DISTINCT [dbo].[MS_Vendor].[Vendor], [dbo].[MS_Vendor].[VendorCode] FROM [MS_Vendor] WHERE [isActive] = 1 
                                    ORDER BY [dbo].[MS_Vendor].[VendorCode] ASC
                                    "
            Using cmd As New SqlCommand(query, conn)
                Using reader As SqlDataReader = cmd.ExecuteReader()
                    dt.Load(reader)
                End Using
            End Using
        End Using
        Return GenerateHtmlVendorDropdown(dt)
    End Function

    Private Function GetMSVendorListwithfilter(segmentCode As String) As String
        Dim dt As New DataTable()
        Using conn As New SqlConnection(connectionString)
            conn.Open()
            Dim query As String = "SELECT DISTINCT [dbo].[MS_Vendor].[Vendor], [dbo].[MS_Vendor].[VendorCode] FROM  [dbo].[MS_Vendor]
                                    WHERE [isActive] = 1 
                                    AND  [BMS].[dbo].[MS_Vendor].[SegmentCode] = @segmentCode
                                    ORDER BY [dbo].[MS_Vendor].[VendorCode] ASC
                                    "
            Using cmd As New SqlCommand(query, conn)
                cmd.Parameters.AddWithValue("@segmentCode", segmentCode)
                Using reader As SqlDataReader = cmd.ExecuteReader()
                    dt.Load(reader)
                End Using
            End Using
        End Using
        Return GenerateHtmlVendorDropdown(dt)
    End Function

    Private Function GetMSCCYListwithfilter(vendorCode As String) As String
        Dim dt As New DataTable()
        Using conn As New SqlConnection(connectionString)
            conn.Open()
            Dim query As String = "SELECT  DISTINCT [CCY]
                                   FROM [BMS].[dbo].[MS_Vendor]
                                   WHERE [isActive] = 1 
                                   AND (@vendorCode IS NULL OR @vendorCode = '' OR [VendorCode] = @vendorCode)
                                   ORDER BY [dbo].[MS_Vendor].[CCY] ASC
                                    "
            Using cmd As New SqlCommand(query, conn)
                cmd.Parameters.AddWithValue("@vendorCode", vendorCode)
                Using reader As SqlDataReader = cmd.ExecuteReader()
                    dt.Load(reader)
                End Using
            End Using
        End Using
        Return GenerateHtmlCCYDropdown(dt)
    End Function

    Private Function GetMSCCYListFromMainMaster() As String
        Dim dt As New DataTable()
        Using conn As New SqlConnection(connectionString)
            conn.Open()
            Dim query As String = "SELECT [CCY_Code] As 'CCY'
                                      FROM [BMS].[dbo].[MS_CCY]
                                      ORDER BY [CCY] ASC
                                    "
            Using cmd As New SqlCommand(query, conn)
                Using reader As SqlDataReader = cmd.ExecuteReader()
                    dt.Load(reader)
                End Using
            End Using
        End Using
        Return GenerateHtmlCCYDropdown(dt)
    End Function

    Private Function GenerateHtmlVersionDropdown(dt As DataTable) As String

        Dim sb As New StringBuilder()
        sb.Append("<option value=''>-- กรุณาเลือก Version --</option>")
        For i As Integer = 0 To dt.Rows.Count - 1
            Dim versionCode As String = If(dt.Rows(i)("VersionCode") IsNot DBNull.Value, dt.Rows(i)("VersionCode").ToString(), "")
            sb.AppendFormat("<option value='{0}'>{1}</option>",
                       HttpUtility.HtmlEncode(versionCode),
                       HttpUtility.HtmlEncode(versionCode))
        Next

        Return sb.ToString()
    End Function

    Private Function GenerateHtmlSegmentDropdown(dt As DataTable) As String

        Dim sb As New StringBuilder()
        sb.Append("<option value=''>-- กรุณาเลือก Segment --</option>")
        For i As Integer = 0 To dt.Rows.Count - 1
            Dim segmentCode As String = If(dt.Rows(i)("SegmentCode") IsNot DBNull.Value, dt.Rows(i)("SegmentCode").ToString(), "")
            Dim segmentName As String = If(dt.Rows(i)("SegmentName") IsNot DBNull.Value, dt.Rows(i)("SegmentName").ToString(), "")
            sb.AppendFormat("<option value='{0}'>{1} - {2}</option>",
                       HttpUtility.HtmlEncode(segmentCode),
                       HttpUtility.HtmlEncode(segmentCode),
                       HttpUtility.HtmlEncode(segmentName))
        Next

        Return sb.ToString()
    End Function

    Private Function GenerateHtmlYearDropdown(dt As DataTable) As String

        Dim sb As New StringBuilder()
        sb.Append("<option value=''>-- กรุณาเลือก Year --</option>")
        For i As Integer = 0 To dt.Rows.Count - 1
            Dim Year_Kept As String = If(dt.Rows(i)("Year") IsNot DBNull.Value, dt.Rows(i)("Year").ToString(), "")

            sb.AppendFormat("<option value='{0}'>{1}</option>",
                       HttpUtility.HtmlEncode(Year_Kept),
                       HttpUtility.HtmlEncode(Year_Kept))
        Next

        Return sb.ToString()
    End Function

    Private Function GenerateHtmlMonthDropdown(dt As DataTable) As String

        Dim sb As New StringBuilder()
        sb.Append("<option value=''>-- กรุณาเลือก Month --</option>")
        For i As Integer = 0 To dt.Rows.Count - 1
            Dim month_code As String = If(dt.Rows(i)("month_code") IsNot DBNull.Value, dt.Rows(i)("month_code").ToString(), "")
            Dim month_name_sh As String = If(dt.Rows(i)("month_name_sh") IsNot DBNull.Value, dt.Rows(i)("month_name_sh").ToString(), "")
            sb.AppendFormat("<option value='{0}'>{1}</option>",
                       HttpUtility.HtmlEncode(month_code),
                       HttpUtility.HtmlEncode(month_name_sh))
        Next

        Return sb.ToString()
    End Function

    Private Function GenerateHtmlCompanyDropdown(dt As DataTable) As String

        Dim sb As New StringBuilder()
        sb.Append("<option value=''>-- กรุณาเลือก Company --</option>")
        For i As Integer = 0 To dt.Rows.Count - 1
            Dim CompanyCode As String = If(dt.Rows(i)("CompanyCode") IsNot DBNull.Value, dt.Rows(i)("CompanyCode").ToString(), "")
            Dim CompanyNameShort As String = If(dt.Rows(i)("CompanyNameShort") IsNot DBNull.Value, dt.Rows(i)("CompanyNameShort").ToString(), "")
            sb.AppendFormat("<option value='{0}'>{1}</option>",
                       HttpUtility.HtmlEncode(CompanyCode),
                       HttpUtility.HtmlEncode(CompanyNameShort))
        Next

        Return sb.ToString()
    End Function


    Private Function GenerateHtmlCategoryDropdown(dt As DataTable) As String

        Dim sb As New StringBuilder()
        If dt.Rows.Count.Equals(0) Then
            sb.Append("<option value=''>-- ไม่พบ Category --</option>")
        Else
            sb.Append("<option value=''>-- กรุณาเลือก Category --</option>")
            For i As Integer = 0 To dt.Rows.Count - 1
                Dim cate As String = If(dt.Rows(i)("Cate") IsNot DBNull.Value, dt.Rows(i)("Cate").ToString(), "")
                Dim category As String = If(dt.Rows(i)("Category") IsNot DBNull.Value, dt.Rows(i)("Category").ToString(), "")
                sb.AppendFormat("<option value='{0}'>{1} - {2}</option>",
                           HttpUtility.HtmlEncode(cate),
                           HttpUtility.HtmlEncode(cate),
                           HttpUtility.HtmlEncode(category))
            Next
        End If
        Return sb.ToString()
    End Function

    Private Function GenerateHtmlBrandDropdown(dt As DataTable) As String

        Dim sb As New StringBuilder()
        If dt.Rows.Count.Equals(0) Then
            sb.Append("<option value=''>-- ไม่พบ Brand --</option>")
        Else
            sb.Append("<option value=''>-- กรุณาเลือก Brand --</option>")
            For i As Integer = 0 To dt.Rows.Count - 1
                Dim brandCode As String = If(dt.Rows(i)("Brand Code") IsNot DBNull.Value, dt.Rows(i)("Brand Code").ToString(), "")
                Dim brandName As String = If(dt.Rows(i)("Brand Name") IsNot DBNull.Value, dt.Rows(i)("Brand Name").ToString(), "")
                sb.AppendFormat("<option value='{0}'>{1} - {2}</option>",
                           HttpUtility.HtmlEncode(brandCode),
                           HttpUtility.HtmlEncode(brandCode),
                           HttpUtility.HtmlEncode(brandName))
            Next
        End If
        Return sb.ToString()
    End Function

    Private Function GenerateHtmlVendorDropdown(dt As DataTable) As String

        Dim sb As New StringBuilder()
        If dt.Rows.Count.Equals(0) Then
            sb.Append("<option value=''>-- ไม่พบ Vendor --</option>")
        Else
            sb.Append("<option value=''>-- กรุณาเลือก Vendor --</option>")
            For i As Integer = 0 To dt.Rows.Count - 1
                Dim vendorCode As String = If(dt.Rows(i)("VendorCode") IsNot DBNull.Value, dt.Rows(i)("VendorCode").ToString(), "")
                Dim vendorName As String = If(dt.Rows(i)("Vendor") IsNot DBNull.Value, dt.Rows(i)("Vendor").ToString(), "")
                sb.AppendFormat("<option value='{0}'>{1} - {2}</option>",
                           HttpUtility.HtmlEncode(vendorCode),
                           HttpUtility.HtmlEncode(vendorCode),
                           HttpUtility.HtmlEncode(vendorName))
            Next
        End If

        Return sb.ToString()
    End Function

    Private Function GenerateHtmlCCYDropdown(dt As DataTable) As String

        Dim sb As New StringBuilder()
        If dt.Rows.Count.Equals(0) Then
            sb.Append("<option value=''>-- กรุณาเลือก Vendor ก่อน --</option>")
        Else
            sb.Append("<option value=''>-- กรุณาเลือก CCY --</option>")
            For i As Integer = 0 To dt.Rows.Count - 1
                Dim CCY As String = If(dt.Rows(i)("CCY") IsNot DBNull.Value, dt.Rows(i)("CCY").ToString(), "")

                sb.AppendFormat("<option value='{0}'>{1}</option>",
                           HttpUtility.HtmlEncode(CCY),
                           HttpUtility.HtmlEncode(CCY))
            Next
        End If

        Return sb.ToString()
    End Function


    ' 1. ฟังก์ชันสำหรับดึงข้อมูล Vendor (สร้าง HTML ของ <tbody>)
    Private Sub GetVendorListHtml(context As HttpContext)
        Dim searchCode As String = context.Request.Form("searchCode")
        Dim searchName As String = context.Request.Form("searchName")
        Dim segmentCode As String = context.Request.Form("segmentCode")

        ' (ย้าย Logic มาจาก BindGridView ใน master_vendor.aspx.vb)
        Dim dt As DataTable = GetVendorData(searchCode, searchName, segmentCode)

        ' (สร้าง HTML <tbody>)
        Dim sb As New StringBuilder()
        If dt.Rows.Count = 0 Then
            ' (MODIFIED) Colspan increased
            sb.Append("<tr><td colspan='10' class='text-center text-muted'>No data found.</td></tr>")
        Else
            For Each row As DataRow In dt.Rows
                ' (START) ADDED: Read isActive
                Dim isActive As Boolean = If(row("isActive") IsNot DBNull.Value, Convert.ToBoolean(row("isActive")), False)
                ' (END) ADDED

                sb.Append("<tr>")
                sb.AppendFormat("<td>{0}</td>", HttpUtility.HtmlEncode(row("VendorCode")))
                sb.AppendFormat("<td>{0}</td>", HttpUtility.HtmlEncode(row("Vendor")))
                sb.AppendFormat("<td>{0}</td>", HttpUtility.HtmlEncode(row("CCY")))
                sb.AppendFormat("<td>{0}</td>", HttpUtility.HtmlEncode(row("PaymentTermCode")))
                sb.AppendFormat("<td>{0}</td>", HttpUtility.HtmlEncode(row("PaymentTerm")))
                sb.AppendFormat("<td>{0}</td>", HttpUtility.HtmlEncode(row("SegmentCode")))
                sb.AppendFormat("<td>{0}</td>", HttpUtility.HtmlEncode(row("Segment")))
                sb.AppendFormat("<td>{0}</td>", HttpUtility.HtmlEncode(row("Incoterm")))

                ' (START) ADDED: Status column
                sb.AppendFormat("<td>{0}</td>", If(isActive, "<span class='badge bg-success'>Active</span>", "<span class='badge bg-secondary'>Inactive</span>"))
                ' (END) ADDED

                ' (สร้างปุ่ม Edit/Delete ใหม่)
                sb.Append("<td class='text-center'>")
                sb.AppendFormat("<button type='button' class='btn btn-edit btn-sm me-1 btn-edit-vendor' " &
                            "data-vendorid='{0}' data-code='{1}' data-name='{2}' data-ccy='{3}' data-term-code='{4}' " &
                            "data-term='{5}' data-seg-code='{6}' data-seg='{7}' data-incoterm='{8}' data-active='{9}' >" &
                            "<i class='bi bi-pencil'></i> Edit</button>",
                            HttpUtility.HtmlAttributeEncode(row("VendorId")),
                            HttpUtility.HtmlAttributeEncode(row("VendorCode")),
                            HttpUtility.HtmlAttributeEncode(row("Vendor")),
                            HttpUtility.HtmlAttributeEncode(row("CCY")),
                            HttpUtility.HtmlAttributeEncode(row("PaymentTermCode")),
                            HttpUtility.HtmlAttributeEncode(row("PaymentTerm")),
                            HttpUtility.HtmlAttributeEncode(row("SegmentCode")),
                            HttpUtility.HtmlAttributeEncode(row("Segment")),
                            HttpUtility.HtmlAttributeEncode(row("Incoterm")),
                            isActive.ToString().ToLower()) ' (MODIFIED) Added data-active

                sb.AppendFormat("<button type='button' class='btn btn-delete btn-sm btn-delete-vendor' " &
                            "data-vendorid='{0}' data-code='{1}' data-name='{2}'><i class='bi bi-trash'></i> Delete</button>",
                            HttpUtility.HtmlAttributeEncode(row("VendorId")),
                            HttpUtility.HtmlAttributeEncode(row("VendorCode")),
                            HttpUtility.HtmlAttributeEncode(row("Vendor")))
                sb.Append("</td>")

                sb.Append("</tr>")
            Next
        End If

        context.Response.ContentType = "text/html"
        context.Response.Write(sb.ToString())
    End Sub

    ' 2. ฟังก์ชันสำหรับ Save (Create/Update)
    Private Sub SaveVendor(context As HttpContext)
        context.Response.ContentType = "application/json"
        Dim serializer As New JavaScriptSerializer()
        Dim response As New Dictionary(Of String, Object)

        Try
            ' (ดึงข้อมูลจาก AJAX)
            Dim editMode As String = context.Request.Form("editMode")
            Dim code As String = context.Request.Form("code")
            ' *** FIX: รับ OriginalCode สำหรับ Edit Mode ***
            Dim vendorId As Int64 = context.Request.Form("vendorId")

            Dim name As String = context.Request.Form("name")
            Dim ccy As String = context.Request.Form("ccy")
            Dim paymentTermCode As String = context.Request.Form("paymentTermCode")
            Dim paymentTerm As String = context.Request.Form("paymentTerm")
            Dim segmentCode As String = context.Request.Form("segmentCode")
            Dim segment As String = context.Request.Form("segment")
            Dim incoterm As String = context.Request.Form("incoterm")
            ' (START) ADDED: Read isActive
            Dim isActive As Boolean = False
            Boolean.TryParse(context.Request.Form("isActive"), isActive)
            ' (END) ADDED

            ' (ย้าย Logic มาจาก btnCreate_Click และ gvVendor_RowUpdating)
            Using conn As New SqlConnection(connectionString)
                conn.Open()
                Dim query As String = ""

                If editMode = "create" Then
                    ' (Logic จาก btnCreate_Click)
                    If CheckVendorCodeAndSectionExists(code, segmentCode) Then
                        Throw New Exception($"Vendor Code '{code}' With SegmentCode '{segmentCode}' already exists!")
                    End If

                    ' (MODIFIED) Added isActive
                    query = "INSERT INTO MS_Vendor ([VendorCode], [Vendor], [CCY], [PaymentTermCode], [PaymentTerm], [SegmentCode], [Segment], [Incoterm], [isActive]) " &
                        "VALUES (@code, @name, @ccy, @paymentTermCode, @paymentTerm, @segmentCode, @segment, @incoterm, @isActive)"
                Else
                    ' (Logic จาก gvVendor_RowUpdating)
                    ' *** FIX: ใช้ originalCode ใน WHERE clause ***
                    ' (MODIFIED) Added isActive
                    query = "UPDATE MS_Vendor SET [VendorCode] = @code, [Vendor] = @name, [CCY] = @ccy, [PaymentTermCode] = @paymentTermCode, " &
                        "[PaymentTerm] = @paymentTerm, [SegmentCode] = @segmentCode, [Segment] = @segment, [Incoterm] = @incoterm, [isActive] = @isActive " &
                        "WHERE [VendorId] = @VendorId"
                End If

                Using cmd As New SqlCommand(query, conn)
                    cmd.Parameters.AddWithValue("@code", code)
                    If editMode = "edit" Then
                        cmd.Parameters.AddWithValue("@VendorId", vendorId)
                    End If
                    cmd.Parameters.AddWithValue("@name", If(String.IsNullOrEmpty(name), String.Empty, name))
                    cmd.Parameters.AddWithValue("@ccy", If(String.IsNullOrEmpty(ccy), String.Empty, ccy))
                    cmd.Parameters.AddWithValue("@paymentTermCode", If(String.IsNullOrEmpty(paymentTermCode), String.Empty, paymentTermCode))
                    cmd.Parameters.AddWithValue("@paymentTerm", If(String.IsNullOrEmpty(paymentTerm), String.Empty, paymentTerm))
                    cmd.Parameters.AddWithValue("@segmentCode", If(String.IsNullOrEmpty(segmentCode), String.Empty, segmentCode))
                    cmd.Parameters.AddWithValue("@segment", If(String.IsNullOrEmpty(segment), String.Empty, segment))
                    cmd.Parameters.AddWithValue("@incoterm", If(String.IsNullOrEmpty(incoterm), String.Empty, incoterm))
                    cmd.Parameters.AddWithValue("@isActive", isActive) ' (MODIFIED) Added parameter

                    cmd.ExecuteNonQuery()
                End Using
            End Using

            response("success") = True
            response("message") = "Vendor saved successfully!"

        Catch ex As Exception
            response("success") = False
            response("message") = ex.Message
        End Try

        context.Response.Write(serializer.Serialize(response))
    End Sub

    ' 3. ฟังก์ชันสำหรับ Delete
    Private Sub DeleteVendor(context As HttpContext)
        context.Response.ContentType = "application/json"
        Dim serializer As New JavaScriptSerializer()
        Dim response As New Dictionary(Of String, Object)

        Try
            Dim vendorId As String = context.Request.Form("vendorId")

            ' (ย้าย Logic มาจาก gvVendor_RowDeleting)
            Using conn As New SqlConnection(connectionString)
                conn.Open()
                Dim query As String = "DELETE FROM MS_Vendor WHERE [VendorId] = @vendorId"
                Using cmd As New SqlCommand(query, conn)
                    cmd.Parameters.AddWithValue("@vendorId", vendorId)
                    cmd.ExecuteNonQuery()
                End Using
            End Using

            response("success") = True
            response("message") = "Vendor deleted successfully!"

        Catch ex As SqlException When ex.Number = 547 ' Foreign key constraint
            response("success") = False
            response("message") = "Cannot delete this vendor. It may be referenced by other records."
        Catch ex As Exception
            response("success") = False
            response("message") = ex.Message
        End Try

        context.Response.Write(serializer.Serialize(response))
    End Sub

    ' 4. (Helper) ฟังก์ชันสำหรับดึงข้อมูล (จาก BindGridView)
    Private Function GetVendorData(searchCode As String, searchName As String, segmentCode As String) As DataTable
        Dim dt As New DataTable()
        ' (Logic เดียวกับ BindGridView เดิม)
        Dim connectionString As String = ConfigurationManager.ConnectionStrings("BMSConnectionString").ConnectionString
        ' (MODIFIED) Added isActive
        Dim query As String = "SELECT [VendorId], [VendorCode], [Vendor], [CCY], [PaymentTermCode], [PaymentTerm], [SegmentCode], [Segment], [Incoterm], ISNULL([isActive], 0) AS [isActive] FROM [MS_Vendor] WHERE 1=1"

        If Not String.IsNullOrEmpty(searchCode) Then
            query &= " AND [VendorCode] LIKE @Code"
        End If
        If Not String.IsNullOrEmpty(searchName) Then
            query &= " AND [Vendor] LIKE @Name"
        End If
        If Not String.IsNullOrEmpty(segmentCode) Then
            query &= " AND [SegmentCode] = @SegmentCode"
        End If
        query &= " ORDER BY [VendorCode]"

        Using conn As New SqlConnection(connectionString)
            Using cmd As New SqlCommand(query, conn)
                If Not String.IsNullOrEmpty(searchCode) Then
                    cmd.Parameters.AddWithValue("@Code", "%" & searchCode & "%")
                End If
                If Not String.IsNullOrEmpty(searchName) Then
                    cmd.Parameters.AddWithValue("@Name", "%" & searchName & "%")
                End If
                If Not String.IsNullOrEmpty(segmentCode) Then
                    cmd.Parameters.AddWithValue("@SegmentCode", segmentCode)
                End If
                conn.Open()
                Dim adapter As New SqlDataAdapter(cmd)
                adapter.Fill(dt)
            End Using
        End Using
        Return dt
    End Function

    ' 5. (Helper) ฟังก์ชันตรวจสอบ (จาก CheckVendorCodeAndSectionExists)
    Private Function CheckVendorCodeAndSectionExists(vendorCode As String, segmentCode As String) As Boolean
        Dim connectionString As String = ConfigurationManager.ConnectionStrings("BMSConnectionString").ConnectionString
        Dim query As String = "SELECT COUNT(*) FROM MS_Vendor WHERE [VendorCode] = @code AND [SegmentCode] = @segmentCode"
        Using conn As New SqlConnection(connectionString)
            Using cmd As New SqlCommand(query, conn)
                cmd.Parameters.AddWithValue("@code", vendorCode)
                cmd.Parameters.AddWithValue("@segmentCode", segmentCode)
                conn.Open()
                Dim count As Integer = Convert.ToInt32(cmd.ExecuteScalar())
                Return count > 0
            End Using
        End Using
    End Function


    ' =================================================================
    ' ===== START: NEW CATEGORY FUNCTIONS =============================
    ' =================================================================

    ' 1. ดึงข้อมูล Category (HTML)
    Private Sub GetCategoryListHtml(context As HttpContext)
        Dim searchCode As String = context.Request.Form("searchCode")
        Dim searchName As String = context.Request.Form("searchName")

        Dim dt As DataTable = GetCategoryData(searchCode, searchName)

        Dim sb As New StringBuilder()
        If dt.Rows.Count = 0 Then
            ' (MODIFIED) Colspan increased
            sb.Append("<tr><td colspan='4' class='text-center text-muted'>No data found.</td></tr>")
        Else
            For Each row As DataRow In dt.Rows
                ' (START) ADDED: Read isActive
                Dim isActive As Boolean = If(row("isActive") IsNot DBNull.Value, Convert.ToBoolean(row("isActive")), False)
                ' (END) ADDED

                sb.Append("<tr>")
                sb.AppendFormat("<td>{0}</td>", HttpUtility.HtmlEncode(row("Cate")))
                sb.AppendFormat("<td>{0}</td>", HttpUtility.HtmlEncode(row("Category")))

                ' (START) ADDED: Status column
                sb.AppendFormat("<td>{0}</td>", If(isActive, "<span class='badge bg-success'>Active</span>", "<span class='badge bg-secondary'>Inactive</span>"))
                ' (END) ADDED

                sb.Append("<td class='text-center'>")
                sb.AppendFormat("<button type='button' class='btn btn-edit btn-sm me-1 btn-edit-category' " &
                            "data-code='{0}' data-name='{1}' data-active='{2}'>" &
                            "<i class='bi bi-pencil'></i> Edit</button>",
                            HttpUtility.HtmlAttributeEncode(row("Cate")),
                            HttpUtility.HtmlAttributeEncode(row("Category")),
                            isActive.ToString().ToLower()) ' (MODIFIED) Added data-active

                sb.AppendFormat("<button type='button' class='btn btn-delete btn-sm btn-delete-category' " &
                            "data-code='{0}' data-name='{1}'><i class='bi bi-trash'></i> Delete</button>",
                            HttpUtility.HtmlAttributeEncode(row("Cate")),
                            HttpUtility.HtmlAttributeEncode(row("Category")))
                sb.Append("</td>")
                sb.Append("</tr>")
            Next
        End If
        context.Response.ContentType = "text/html"
        context.Response.Write(sb.ToString())
    End Sub

    ' 2. บันทึก Category
    Private Sub SaveCategory(context As HttpContext)
        context.Response.ContentType = "application/json"
        Dim serializer As New JavaScriptSerializer()
        Dim response As New Dictionary(Of String, Object)

        Try
            Dim editMode As String = context.Request.Form("editMode")
            Dim code As String = context.Request.Form("code")
            Dim originalCode As String = If(editMode = "edit", context.Request.Form("originalCode"), code)
            Dim name As String = context.Request.Form("name")
            ' (START) ADDED: Read isActive
            Dim isActive As Boolean = False
            Boolean.TryParse(context.Request.Form("isActive"), isActive)
            ' (END) ADDED

            If String.IsNullOrEmpty(code) OrElse String.IsNullOrEmpty(name) Then
                Throw New Exception("Category Code and Category Name are required!")
            End If

            Using conn As New SqlConnection(connectionString)
                conn.Open()
                Dim query As String = ""

                If editMode = "create" Then
                    If CheckCategoryCodeExists(code) Then
                        Throw New Exception($"Category Code '{code}' already exists!")
                    End If
                    ' (MODIFIED) Added isActive
                    query = "INSERT INTO MS_Category ([Cate], [Category], [isActive]) VALUES (@code, @name, @isActive)"
                Else
                    ' (MODIFIED) Added isActive
                    query = "UPDATE MS_Category SET [Cate] = @code, [Category] = @name, [isActive] = @isActive WHERE [Cate] = @originalCode"
                End If

                Using cmd As New SqlCommand(query, conn)
                    cmd.Parameters.AddWithValue("@code", code)
                    If editMode = "edit" Then
                        cmd.Parameters.AddWithValue("@originalCode", originalCode)
                    End If
                    cmd.Parameters.AddWithValue("@name", name)
                    cmd.Parameters.AddWithValue("@isActive", isActive) ' (MODIFIED) Added parameter
                    cmd.ExecuteNonQuery()
                End Using
            End Using

            response("success") = True
            response("message") = "Category saved successfully!"
        Catch ex As Exception
            response("success") = False
            response("message") = ex.Message
        End Try
        context.Response.Write(serializer.Serialize(response))
    End Sub

    ' 3. ลบ Category
    Private Sub DeleteCategory(context As HttpContext)
        context.Response.ContentType = "application/json"
        Dim serializer As New JavaScriptSerializer()
        Dim response As New Dictionary(Of String, Object)

        Try
            Dim categoryCode As String = context.Request.Form("categoryCode")

            Using conn As New SqlConnection(connectionString)
                conn.Open()
                Dim query As String = "DELETE FROM MS_Category WHERE [Cate] = @code"
                Using cmd As New SqlCommand(query, conn)
                    cmd.Parameters.AddWithValue("@code", categoryCode)
                    cmd.ExecuteNonQuery()
                End Using
            End Using

            response("success") = True
            response("message") = "Category deleted successfully!"
        Catch ex As SqlException When ex.Number = 547 ' Foreign key constraint
            response("success") = False
            response("message") = "Cannot delete this category. It may be referenced by other records."
        Catch ex As Exception
            response("success") = False
            response("message") = ex.Message
        End Try
        context.Response.Write(serializer.Serialize(response))
    End Sub

    ' 4. (Helper) ดึงข้อมูล Category
    Private Function GetCategoryData(Optional searchCode As String = "", Optional searchName As String = "") As DataTable
        Dim dt As New DataTable()
        ' (MODIFIED) Added isActive
        Dim query As String = "SELECT [Cate], [Category], ISNULL([isActive], 0) AS [isActive] FROM [MS_Category] WHERE 1=1"

        If Not String.IsNullOrEmpty(searchCode) Then
            query &= " AND [Cate] LIKE @Code"
        End If
        If Not String.IsNullOrEmpty(searchName) Then
            query &= " AND [Category] LIKE @Name"
        End If
        query &= " ORDER BY [Cate]"

        Using conn As New SqlConnection(connectionString)
            Using cmd As New SqlCommand(query, conn)
                If Not String.IsNullOrEmpty(searchCode) Then
                    cmd.Parameters.AddWithValue("@Code", "%" & searchCode & "%")
                End If
                If Not String.IsNullOrEmpty(searchName) Then
                    cmd.Parameters.AddWithValue("@Name", "%" & searchName & "%")
                End If
                conn.Open()
                Dim adapter As New SqlDataAdapter(cmd)
                adapter.Fill(dt)
            End Using
        End Using
        Return dt
    End Function

    ' 5. (Helper) ตรวจสอบ Category Code
    Private Function CheckCategoryCodeExists(categoryCode As String) As Boolean
        Dim query As String = "SELECT COUNT(*) FROM MS_Category WHERE [Cate] = @code"
        Using conn As New SqlConnection(connectionString)
            Using cmd As New SqlCommand(query, conn)
                cmd.Parameters.AddWithValue("@code", categoryCode)
                conn.Open()
                Dim count As Integer = Convert.ToInt32(cmd.ExecuteScalar())
                Return count > 0
            End Using
        End Using
    End Function
    ' =================================================================
    ' ===== END: NEW CATEGORY FUNCTIONS ===============================
    ' =================================================================


    ' =================================================================
    ' ===== START: NEW BRAND FUNCTIONS ================================
    ' =================================================================

    ' 1. ดึงข้อมูล Brand (HTML)
    Private Sub GetBrandListHtml(context As HttpContext)
        Dim searchCode As String = context.Request.Form("searchCode")
        Dim searchName As String = context.Request.Form("searchName")

        Dim dt As DataTable = GetBrandData(searchCode, searchName)

        Dim sb As New StringBuilder()
        If dt.Rows.Count = 0 Then
            sb.Append("<tr><td colspan='4' class='text-center text-muted'>No data found.</td></tr>")
        Else
            For Each row As DataRow In dt.Rows

                Dim isActive As Boolean = If(row("isActive") IsNot DBNull.Value, Convert.ToBoolean(row("isActive")), False)

                sb.Append("<tr>")
                sb.AppendFormat("<td>{0}</td>", HttpUtility.HtmlEncode(row("Brand Code")))
                sb.AppendFormat("<td>{0}</td>", HttpUtility.HtmlEncode(row("Brand Name")))
                sb.AppendFormat("<td>{0}</td>", If(isActive, "<span class='badge bg-success'>Active</span>", "<span class='badge bg-secondary'>Inactive</span>"))
                sb.Append("<td class='text-center'>")
                sb.AppendFormat("<button type='button' class='btn btn-edit btn-sm me-1 btn-edit-brand' " &
                            "data-code='{0}' data-name='{1}' data-active='{2}'>" &
                            "<i class='bi bi-pencil'></i> Edit</button>",
                            HttpUtility.HtmlAttributeEncode(row("Brand Code")),
                            HttpUtility.HtmlAttributeEncode(row("Brand Name")),
                            isActive.ToString().ToLower())

                sb.AppendFormat("<button type='button' class='btn btn-delete btn-sm btn-delete-brand' " &
                            "data-code='{0}' data-name='{1}'><i class='bi bi-trash'></i> Delete</button>",
                            HttpUtility.HtmlAttributeEncode(row("Brand Code")),
                            HttpUtility.HtmlAttributeEncode(row("Brand Name")))
                sb.Append("</td>")
                sb.Append("</tr>")
            Next
        End If
        context.Response.ContentType = "text/html"
        context.Response.Write(sb.ToString())
    End Sub

    ' 2. บันทึก Brand
    Private Sub SaveBrand(context As HttpContext)
        context.Response.ContentType = "application/json"
        Dim serializer As New JavaScriptSerializer()
        Dim response As New Dictionary(Of String, Object)

        Try
            Dim editMode As String = context.Request.Form("editMode")
            Dim code As String = context.Request.Form("code")
            Dim originalCode As String = If(editMode = "edit", context.Request.Form("originalCode"), code)
            Dim name As String = context.Request.Form("name")

            Dim isActive As Boolean = False
            Boolean.TryParse(context.Request.Form("isActive"), isActive)

            If String.IsNullOrEmpty(code) OrElse String.IsNullOrEmpty(name) Then
                Throw New Exception("Brand Code and Brand Name are required!")
            End If

            Using conn As New SqlConnection(connectionString)
                conn.Open()
                Dim query As String = ""

                If editMode = "create" Then
                    If CheckBrandCodeExists(code) Then
                        Throw New Exception($"Brand Code '{code}' already exists!")
                    End If
                    query = "INSERT INTO MS_Brand ([Brand Code], [Brand Name], [isActive]) VALUES (@code, @name, @isActive)"
                Else
                    query = "UPDATE MS_Brand SET [Brand Code] = @code, [Brand Name] = @name, [isActive] = @isActive WHERE [Brand Code] = @originalCode"
                End If

                Using cmd As New SqlCommand(query, conn)
                    cmd.Parameters.AddWithValue("@code", code)
                    If editMode = "edit" Then
                        cmd.Parameters.AddWithValue("@originalCode", originalCode)
                    End If
                    cmd.Parameters.AddWithValue("@name", name)
                    cmd.Parameters.AddWithValue("@isActive", isActive)
                    cmd.ExecuteNonQuery()
                End Using
            End Using

            response("success") = True
            response("message") = "Brand saved successfully!"
        Catch ex As Exception
            response("success") = False
            response("message") = ex.Message
        End Try
        context.Response.Write(serializer.Serialize(response))
    End Sub

    ' 3. ลบ Brand
    Private Sub DeleteBrand(context As HttpContext)
        context.Response.ContentType = "application/json"
        Dim serializer As New JavaScriptSerializer()
        Dim response As New Dictionary(Of String, Object)

        Try
            Dim brandCode As String = context.Request.Form("brandCode")

            Using conn As New SqlConnection(connectionString)
                conn.Open()
                Dim query As String = "DELETE FROM MS_Brand WHERE [Brand Code] = @code"
                Using cmd As New SqlCommand(query, conn)
                    cmd.Parameters.AddWithValue("@code", brandCode)
                    cmd.ExecuteNonQuery()
                End Using
            End Using

            response("success") = True
            response("message") = "Brand deleted successfully!"
        Catch ex As SqlException When ex.Number = 547 ' Foreign key constraint
            response("success") = False
            response("message") = "Cannot delete this brand. It may be referenced by other records."
        Catch ex As Exception
            response("success") = False
            response("message") = ex.Message
        End Try
        context.Response.Write(serializer.Serialize(response))
    End Sub

    ' 4. (Helper) ดึงข้อมูล Brand
    Private Function GetBrandData(Optional searchCode As String = "", Optional searchName As String = "") As DataTable
        Dim dt As New DataTable()
        Dim query As String = "SELECT [Brand Code], [Brand Name], [isActive] FROM [MS_Brand] WHERE 1=1"

        If Not String.IsNullOrEmpty(searchCode) Then
            query &= " AND [Brand Code] LIKE @Code"
        End If
        If Not String.IsNullOrEmpty(searchName) Then
            query &= " AND [Brand Name] LIKE @Name"
        End If
        query &= " ORDER BY [Brand Code]"

        Using conn As New SqlConnection(connectionString)
            Using cmd As New SqlCommand(query, conn)
                If Not String.IsNullOrEmpty(searchCode) Then
                    cmd.Parameters.AddWithValue("@Code", "%" & searchCode & "%")
                End If
                If Not String.IsNullOrEmpty(searchName) Then
                    cmd.Parameters.AddWithValue("@Name", "%" & searchName & "%")
                End If
                conn.Open()
                Dim adapter As New SqlDataAdapter(cmd)
                adapter.Fill(dt)
            End Using
        End Using
        Return dt
    End Function

    ' 5. (Helper) ตรวจสอบ Brand Code
    Private Function CheckBrandCodeExists(brandCode As String) As Boolean
        Dim query As String = "SELECT COUNT(*) FROM MS_Brand WHERE [Brand Code] = @code"
        Using conn As New SqlConnection(connectionString)
            Using cmd As New SqlCommand(query, conn)
                cmd.Parameters.AddWithValue("@code", brandCode)
                conn.Open()
                Dim count As Integer = Convert.ToInt32(cmd.ExecuteScalar())
                Return count > 0
            End Using
        End Using
    End Function
    ' =================================================================
    ' ===== END: NEW BRAND FUNCTIONS ==================================
    ' =================================================================

    ReadOnly Property IsReusable() As Boolean Implements IHttpHandler.IsReusable
        Get
            Return False
        End Get
    End Property

End Class