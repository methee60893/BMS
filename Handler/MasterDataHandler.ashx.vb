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
Imports ExcelDataReader
Imports Newtonsoft.Json

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
                context.Response.Write(GetMSCCYListwithfilter(""))
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
                                    WHERE [BMS].[dbo].[Template_Upload_Draft_OTB].[SegmentCode] = @segmentCode
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
            Dim query As String = "SELECT  [dbo].[MS_Category].[Category], [dbo].[MS_Category].Cate FROM [dbo].[MS_Category]
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
            Dim query As String = "SELECT [dbo].[MS_Brand].[Brand Name], [dbo].[MS_Brand].[Brand Code] FROM  [dbo].[MS_Brand]
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
            Dim query As String = "SELECT DISTINCT [dbo].[MS_Vendor].[Vendor], [dbo].[MS_Vendor].[VendorCode] FROM [MS_Vendor]
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
                                    WHERE [BMS].[dbo].[MS_Vendor].[SegmentCode] = @segmentCode
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
                                     WHERE (@vendorCode IS NULL OR @vendorCode = '' OR [VendorCode] = @vendorCode)
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



    ReadOnly Property IsReusable() As Boolean Implements IHttpHandler.IsReusable
        Get
            Return False
        End Get
    End Property

End Class