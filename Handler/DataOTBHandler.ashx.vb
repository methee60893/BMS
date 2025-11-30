Imports System
Imports System.Data
Imports System.Data.SqlClient
Imports System.Drawing
Imports System.Globalization
Imports System.IO
Imports System.Text
Imports System.Web
Imports System.Web.Script.Serialization
Imports System.Web.Services.Description
Imports ExcelDataReader
Imports Newtonsoft.Json
Imports OfficeOpenXml
Imports OfficeOpenXml.Style
Imports System.Threading.Tasks

Public Class DataOTBHandler
    Implements IHttpHandler

    Private Shared connectionString As String = ConfigurationManager.ConnectionStrings("BMSConnectionString")?.ConnectionString
    ' *** NEW: Private Class for strong typing in Summary Report ***
    Private Class SummaryRawItem
        Public Property Company As String
        Public Property CompanyName As String
        Public Property Year As String
        Public Property Month As String
        Public Property MonthName As String
        Public Property Category As String
        Public Property CategoryName As String
        Public Property Segment As String
        Public Property SegmentName As String
        Public Property TotalBudget As Decimal
        Public Property TotalActualDraft As Decimal
    End Class
    Sub ProcessRequest(ByVal context As HttpContext) Implements IHttpHandler.ProcessRequest

        Try
            ' *** MODIFIED: Moved action check to the top ***
            Dim action As String = If(context.Request("action"), "").ToLower().Trim()

            If action = "exportdraftotb" Then
                ' *** NEW: Call dedicated export function ***
                HandleExportDraftOTB(context)
            ElseIf action = "exportapprovedotb" Then
                HandleExportApprovedOTB(context)
            ElseIf action = "exportswitchingtxn" Then
                HandleExportSwitchingOTB(context)
            ElseIf action = "exportdraftotbsum" Then
                ExportDraftOTBSum(context)
            ElseIf action = "exportotbmovement" Then
                HandleExportOTBMovement(context)
            ElseIf action = "exportsummarycategory" Then
                HandleExportSummaryCategory(context)
            Else
                ' *** MOVED: The rest of the logic into an Else block ***
                context.Response.Clear()
                context.Response.ContentType = "text/html"
                context.Response.ContentEncoding = Encoding.UTF8
                Dim dt As DataTable = Nothing

                Dim OTBtype As String = If(String.IsNullOrWhiteSpace(context.Request.Form("OTBtype")), "", context.Request.Form("OTBtype").Trim())
                Dim OTByear As String = If(String.IsNullOrWhiteSpace(context.Request.Form("OTByear")), "", context.Request.Form("OTByear").Trim())
                Dim OTBmonth As String = If(String.IsNullOrWhiteSpace(context.Request.Form("OTBmonth")), "", context.Request.Form("OTBmonth").Trim())
                Dim OTBCompany As String = If(String.IsNullOrWhiteSpace(context.Request.Form("OTBCompany")), "", context.Request.Form("OTBCompany").Trim())
                Dim OTBCategory As String = If(String.IsNullOrWhiteSpace(context.Request.Form("OTBCategory")), "", context.Request.Form("OTBCategory").Trim())
                Dim OTBSegment As String = If(String.IsNullOrWhiteSpace(context.Request.Form("OTBSegment")), "", context.Request.Form("OTBSegment").Trim())
                Dim OTBBrand As String = If(String.IsNullOrWhiteSpace(context.Request.Form("OTBBrand")), "", context.Request.Form("OTBBrand").Trim())
                Dim OTBVendor As String = If(String.IsNullOrWhiteSpace(context.Request.Form("OTBVendor")), "", context.Request.Form("OTBVendor").Trim())
                Dim OTBVersion As String = If(String.IsNullOrWhiteSpace(context.Request.Form("OTBVersion")), "", context.Request.Form("OTBVersion").Trim())

                If context.Request("action") = "obtlistbyfilter" Then
                    dt = GetOTBDraftDataWithFilter(OTBtype, OTByear, OTBmonth, OTBCompany, OTBCategory, OTBSegment, OTBBrand, OTBVendor)
                    context.Response.Write(GenerateHtmlDraftTable(dt))
                ElseIf context.Request("action") = "approveDraftOTB" Then
                    ApproveDraftOTB(context)
                ElseIf context.Request("action") = "deleteDraftOTB" Then
                    DeleteDraftOTB(context)
                ElseIf context.Request("action") = "obtApprovelistbyfilter" Then
                    dt = GetOTBApproveDataWithFilter(OTBtype, OTByear, OTBmonth, OTBCompany, OTBCategory, OTBSegment, OTBBrand, OTBVendor, OTBVersion)
                    context.Response.Write(GenerateHtmlApprovedTable(dt))
                ElseIf context.Request("action") = "obtswitchlistbyfilter" Then
                    dt = GetOTBSwitchDataWithFilter(OTBtype, OTByear, OTBmonth, OTBCompany, OTBCategory, OTBSegment, OTBBrand, OTBVendor)
                    context.Response.Write(GenerateHtmlSwitchable(dt))
                End If
            End If
        Catch ex As Exception
            context.Response.StatusCode = 500
            context.Response.Write($"<div class='alert alert-danger'>Error: {HttpUtility.HtmlEncode(ex.Message)}</div>")
        End Try

    End Sub

    ' *** NEW: Function for Switching Export ***
    Private Sub HandleExportSwitchingOTB(ByVal context As HttpContext)
        ' 1. Read filters from QueryString
        Dim OTBtype As String = If(String.IsNullOrWhiteSpace(context.Request.QueryString("OTBtype")), "", context.Request.QueryString("OTBtype").Trim())
        Dim OTByear As String = If(String.IsNullOrWhiteSpace(context.Request.QueryString("OTByear")), "", context.Request.QueryString("OTByear").Trim())
        Dim OTBmonth As String = If(String.IsNullOrWhiteSpace(context.Request.QueryString("OTBmonth")), "", context.Request.QueryString("OTBmonth").Trim())
        Dim OTBCompany As String = If(String.IsNullOrWhiteSpace(context.Request.QueryString("OTBCompany")), "", context.Request.QueryString("OTBCompany").Trim())
        Dim OTBCategory As String = If(String.IsNullOrWhiteSpace(context.Request.QueryString("OTBCategory")), "", context.Request.QueryString("OTBCategory").Trim())
        Dim OTBSegment As String = If(String.IsNullOrWhiteSpace(context.Request.QueryString("OTBSegment")), "", context.Request.QueryString("OTBSegment").Trim())
        Dim OTBBrand As String = If(String.IsNullOrWhiteSpace(context.Request.QueryString("OTBBrand")), "", context.Request.QueryString("OTBBrand").Trim())
        Dim OTBVendor As String = If(String.IsNullOrWhiteSpace(context.Request.QueryString("OTBVendor")), "", context.Request.QueryString("OTBVendor").Trim())

        ' 2. Get Raw Data
        Dim dtRaw As DataTable = GetOTBSwitchDataWithFilter(OTBtype, OTByear, OTBmonth, OTBCompany, OTBCategory, OTBSegment, OTBBrand, OTBVendor)

        ' 3. Format data for export
        Dim dtExport As DataTable = FormatSwitchDataForExport(dtRaw)

        ' 4. Call generic export function
        ExportDataTableToExcel(context, dtExport, "Switching_OTB_" & DateTime.Now.ToString("yyyyMMdd_HHmmss") & ".xlsx")
    End Sub

    ' *** NEW: Helper to format Switching data for export ***
    Private Function FormatSwitchDataForExport(dtRaw As DataTable) As DataTable
        Dim dtExport As New DataTable("SwitchingOTB")

        ' Add headers matching the complex table structure
        dtExport.Columns.Add("Create Date", GetType(String))
        dtExport.Columns.Add("Type (Out)", GetType(String))
        dtExport.Columns.Add("Year (Out)", GetType(String))
        dtExport.Columns.Add("Month (Out)", GetType(String))
        dtExport.Columns.Add("Category (Out)", GetType(String))
        dtExport.Columns.Add("Category Name (Out)", GetType(String))
        dtExport.Columns.Add("Company (Out)", GetType(String))
        dtExport.Columns.Add("Segment (Out)", GetType(String))
        dtExport.Columns.Add("Segment Name (Out)", GetType(String))
        dtExport.Columns.Add("Brand (Out)", GetType(String))
        dtExport.Columns.Add("Brand Name (Out)", GetType(String))
        dtExport.Columns.Add("Vendor (Out)", GetType(String))
        dtExport.Columns.Add("Vendor Name (Out)", GetType(String))

        dtExport.Columns.Add("Type (In)", GetType(String))
        dtExport.Columns.Add("Year (In)", GetType(String))
        dtExport.Columns.Add("Month (In)", GetType(String))
        dtExport.Columns.Add("Category (In)", GetType(String))
        dtExport.Columns.Add("Company (In)", GetType(String))
        dtExport.Columns.Add("Segment (In)", GetType(String))
        dtExport.Columns.Add("Brand (In)", GetType(String))
        dtExport.Columns.Add("Vendor (In)", GetType(String))

        dtExport.Columns.Add("Amount (THB)", GetType(Decimal))
        dtExport.Columns.Add("Create By", GetType(String))

        For Each row As DataRow In dtRaw.Rows
            dtExport.Rows.Add(
                If(row("CreateDT") IsNot DBNull.Value, Convert.ToDateTime(row("CreateDT")).ToString("dd/MM/yyyy HH:mm"), ""),
                If(row("Type") IsNot DBNull.Value, row("Type").ToString(), ""),
                If(row("Year") IsNot DBNull.Value, row("Year").ToString(), ""),
                If(row("MonthName") IsNot DBNull.Value, row("MonthName").ToString(), ""),
                If(row("Category") IsNot DBNull.Value, row("Category").ToString(), ""),
                If(row("CategoryName") IsNot DBNull.Value, row("CategoryName").ToString(), ""),
                If(row("CompanyName") IsNot DBNull.Value, row("CompanyName").ToString(), ""),
                If(row("Segment") IsNot DBNull.Value, row("Segment").ToString(), ""),
                If(row("SegmentName") IsNot DBNull.Value, row("SegmentName").ToString(), ""),
                If(row("Brand") IsNot DBNull.Value, row("Brand").ToString(), ""),
                If(row("BrandName") IsNot DBNull.Value, row("BrandName").ToString(), ""),
                If(row("Vendor") IsNot DBNull.Value, row("Vendor").ToString(), ""),
                If(row("VendorName") IsNot DBNull.Value, row("VendorName").ToString(), ""),
                If(row("SwitchType") IsNot DBNull.Value, row("SwitchType").ToString(), ""), ' Assuming column name from SP is 'ToType'
                If(row("SwitchYear") IsNot DBNull.Value, row("SwitchYear").ToString(), ""),
                If(row("SwitchMonthName") IsNot DBNull.Value, row("SwitchMonthName").ToString(), ""),
                If(row("SwitchCategory") IsNot DBNull.Value, row("SwitchCategory").ToString(), ""),
                If(row("SwitchCompanyName") IsNot DBNull.Value, row("SwitchCompanyName").ToString(), ""),
                If(row("SwitchSegment") IsNot DBNull.Value, row("SwitchSegment").ToString(), ""),
                If(row("SwitchBrand") IsNot DBNull.Value, row("SwitchBrand").ToString(), ""),
                If(row("SwitchVendor") IsNot DBNull.Value, row("SwitchVendor").ToString(), ""),
                If(row("BudgetAmount") IsNot DBNull.Value, Convert.ToDecimal(row("BudgetAmount")), 0),
                If(row("CreateBy") IsNot DBNull.Value, row("CreateBy").ToString(), ""))
        Next

        Return dtExport
    End Function


    ' *** (Refactored) Function for Draft Export ***
    Private Sub HandleExportDraftOTB(ByVal context As HttpContext)
        ' (Note: We use QueryString because the JS call is a GET request for file download)
        Dim OTBtype As String = If(String.IsNullOrWhiteSpace(context.Request.QueryString("OTBtype")), "", context.Request.QueryString("OTBtype").Trim())
        Dim OTByear As String = If(String.IsNullOrWhiteSpace(context.Request.QueryString("OTByear")), "", context.Request.QueryString("OTByear").Trim())
        Dim OTBmonth As String = If(String.IsNullOrWhiteSpace(context.Request.QueryString("OTBmonth")), "", context.Request.QueryString("OTBmonth").Trim())
        Dim OTBCompany As String = If(String.IsNullOrWhiteSpace(context.Request.QueryString("OTBCompany")), "", context.Request.QueryString("OTBCompany").Trim())
        Dim OTBCategory As String = If(String.IsNullOrWhiteSpace(context.Request.QueryString("OTBCategory")), "", context.Request.QueryString("OTBCategory").Trim())
        Dim OTBSegment As String = If(String.IsNullOrWhiteSpace(context.Request.QueryString("OTBSegment")), "", context.Request.QueryString("OTBSegment").Trim())
        Dim OTBBrand As String = If(String.IsNullOrWhiteSpace(context.Request.QueryString("OTBBrand")), "", context.Request.QueryString("OTBBrand").Trim())
        Dim OTBVendor As String = If(String.IsNullOrWhiteSpace(context.Request.QueryString("OTBVendor")), "", context.Request.QueryString("OTBVendor").Trim())

        ' 1. Get Raw Data
        Dim dtRaw As DataTable = GetOTBDraftDataWithFilter(OTBtype, OTByear, OTBmonth, OTBCompany, OTBCategory, OTBSegment, OTBBrand, OTBVendor)
        Dim budgetCalculator As New OTBBudgetCalculator()
        ' 2. Create Export-formatted DataTable (to match the HTML table)
        Dim dtExport As New DataTable("DraftOTB")
        dtExport.Columns.Add("Create Date", GetType(String))
        dtExport.Columns.Add("Type", GetType(String))
        dtExport.Columns.Add("Year", GetType(String))
        dtExport.Columns.Add("Month", GetType(String))
        dtExport.Columns.Add("Category", GetType(String))
        dtExport.Columns.Add("Category Name", GetType(String))
        dtExport.Columns.Add("Company", GetType(String))
        dtExport.Columns.Add("Segment", GetType(String))
        dtExport.Columns.Add("Segment Name", GetType(String))
        dtExport.Columns.Add("Brand", GetType(String))
        dtExport.Columns.Add("Brand Name", GetType(String))
        dtExport.Columns.Add("Vendor", GetType(String))
        dtExport.Columns.Add("Vendor Name", GetType(String))
        dtExport.Columns.Add("Current Approved", GetType(Decimal))
        dtExport.Columns.Add("TO-BE Amount (THB)", GetType(Decimal))
        dtExport.Columns.Add("Diff", GetType(Decimal))
        dtExport.Columns.Add("Status", GetType(String))
        dtExport.Columns.Add("Version", GetType(String))
        dtExport.Columns.Add("Remark", GetType(String))

        ' 3. Populate dtExport with calculated fields (mirroring GenerateHtmlDraftTable)
        For Each row As DataRow In dtRaw.Rows
            Dim OTBYear_Calc As String = If(row("OTBYear") IsNot DBNull.Value, row("OTBYear").ToString(), "")
            Dim OTBMonth_Calc As String = If(row("OTBMonth") IsNot DBNull.Value, row("OTBMonth").ToString(), "")
            Dim OTBCategory_Calc As String = If(row("OTBCategory") IsNot DBNull.Value, row("OTBCategory").ToString(), "")
            Dim OTBCompany_Calc As String = If(row("OTBCompany") IsNot DBNull.Value, row("OTBCompany").ToString(), "")
            Dim OTBSegment_Calc As String = If(row("OTBSegment") IsNot DBNull.Value, row("OTBSegment").ToString(), "")
            Dim OTBBrand_Calc As String = If(row("OTBBrand") IsNot DBNull.Value, row("OTBBrand").ToString(), "")
            Dim OTBVendor_Calc As String = If(row("OTBVendor") IsNot DBNull.Value, row("OTBVendor").ToString(), "")

            Dim amountValue As Decimal = 0
            Decimal.TryParse(If(row("Amount") IsNot DBNull.Value, row("Amount").ToString(), ""), amountValue)

            Dim currentBudget As Decimal = budgetCalculator.CalculateCurrentApprovedBudget(OTBYear_Calc, OTBMonth_Calc, OTBCategory_Calc, OTBCompany_Calc, OTBSegment_Calc, OTBBrand_Calc, OTBVendor_Calc)
            Dim diffAmount As Decimal = amountValue - currentBudget
            Dim OTBStatus As String = If(row("OTBStatus") IsNot DBNull.Value, row("OTBStatus").ToString(), "Draft")

            dtExport.Rows.Add(
                If(row("CreateDT") IsNot DBNull.Value, row("CreateDT").ToString(), ""),
                If(row("OTBType") IsNot DBNull.Value, row("OTBType").ToString(), ""),
                OTBYear_Calc,
                If(row("month_name_sh") IsNot DBNull.Value, row("month_name_sh").ToString(), ""),
                OTBCategory_Calc,
                If(row("CateName") IsNot DBNull.Value, row("CateName").ToString(), ""),
                If(row("CompanyName") IsNot DBNull.Value, row("CompanyName").ToString(), ""),
                OTBSegment_Calc,
                If(row("SegmentName") IsNot DBNull.Value, row("SegmentName").ToString(), ""),
                OTBBrand_Calc,
                If(row("BrandName") IsNot DBNull.Value, row("BrandName").ToString(), ""),
                OTBVendor_Calc,
                If(row("Vendor") IsNot DBNull.Value, row("Vendor").ToString(), ""),
                currentBudget,
                amountValue,
                diffAmount,
                OTBStatus,
                If(row("Version") IsNot DBNull.Value, row("Version").ToString(), ""),
                If(row("Remark") IsNot DBNull.Value, row("Remark").ToString(), "")
            )
        Next

        ' 4. Call Export Function
        ExportDataTableToExcel(context, dtExport, "Draft_OTB_" & DateTime.Now.ToString("yyyyMMdd_HHmmss") & ".xlsx")
    End Sub

    ' *** NEW: Function for Approved Export ***
    Private Sub HandleExportApprovedOTB(ByVal context As HttpContext)
        ' 1. Read filters from QueryString
        Dim OTBtype As String = If(String.IsNullOrWhiteSpace(context.Request.QueryString("OTBtype")), "", context.Request.QueryString("OTBtype").Trim())
        Dim OTByear As String = If(String.IsNullOrWhiteSpace(context.Request.QueryString("OTByear")), "", context.Request.QueryString("OTByear").Trim())
        Dim OTBmonth As String = If(String.IsNullOrWhiteSpace(context.Request.QueryString("OTBmonth")), "", context.Request.QueryString("OTBmonth").Trim())
        Dim OTBCompany As String = If(String.IsNullOrWhiteSpace(context.Request.QueryString("OTBCompany")), "", context.Request.QueryString("OTBCompany").Trim())
        Dim OTBCategory As String = If(String.IsNullOrWhiteSpace(context.Request.QueryString("OTBCategory")), "", context.Request.QueryString("OTBCategory").Trim())
        Dim OTBSegment As String = If(String.IsNullOrWhiteSpace(context.Request.QueryString("OTBSegment")), "", context.Request.QueryString("OTBSegment").Trim())
        Dim OTBBrand As String = If(String.IsNullOrWhiteSpace(context.Request.QueryString("OTBBrand")), "", context.Request.QueryString("OTBBrand").Trim())
        Dim OTBVendor As String = If(String.IsNullOrWhiteSpace(context.Request.QueryString("OTBVendor")), "", context.Request.QueryString("OTBVendor").Trim())
        Dim OTBVersion As String = If(String.IsNullOrWhiteSpace(context.Request.QueryString("OTBVersion")), "", context.Request.QueryString("OTBVersion").Trim())

        ' 2. Get Raw Data (using the same function as the 'View' action)
        Dim dtRaw As DataTable = GetOTBApproveDataWithFilter(OTBtype, OTByear, OTBmonth, OTBCompany, OTBCategory, OTBSegment, OTBBrand, OTBVendor, OTBVersion)

        ' 3. Format data for export (using existing function from ApprovedOTBManager)
        Dim dtExport As DataTable = ApprovedOTBManager.ExportToDataTable(dtRaw)

        ' 4. Call generic export function
        ExportDataTableToExcel(context, dtExport, "Approved_OTB_" & DateTime.Now.ToString("yyyyMMdd_HHmmss") & ".xlsx")
    End Sub


    ' *** RENAMED: from ExportDraftToExcel to ExportDataTableToExcel ***
    Private Sub ExportDataTableToExcel(context As HttpContext, dt As DataTable, filename As String)
        ' Set the license context for EPPlus
        ExcelPackage.License.SetNonCommercialOrganization("KingPower")

        Using package As New ExcelPackage()
            Dim worksheet = package.Workbook.Worksheets.Add("Data")
            worksheet.Cells("A1").LoadFromDataTable(dt, True)

            ' Format header
            Using range = worksheet.Cells(1, 1, 1, dt.Columns.Count)
                range.Style.Font.Bold = True
                range.Style.Fill.PatternType = ExcelFillStyle.Solid
                range.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(13, 110, 253)) ' Header Blue
                range.Style.Font.Color.SetColor(System.Drawing.Color.White)
            End Using

            ' Format number columns (Generic formatting for potential decimals)
            For i As Integer = 1 To dt.Columns.Count
                If dt.Columns(i - 1).DataType Is GetType(Decimal) Or dt.Columns(i - 1).DataType Is GetType(Double) Then
                    worksheet.Column(i).Style.Numberformat.Format = "#,##0.00"
                End If
            Next

            ' Auto-fit columns
            worksheet.Cells.AutoFitColumns()

            ' Download
            context.Response.Clear()
            context.Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            context.Response.AddHeader("content-disposition", $"attachment; filename={filename}")
            context.Response.BinaryWrite(package.GetAsByteArray())
            context.Response.Flush()
            context.ApplicationInstance.CompleteRequest()
        End Using
    End Sub

    ' --- (START) NEW FUNCTION FOR SUMMARY EXPORT ---
    ''' <summary>
    ''' สร้างไฟล์ Excel สรุป OTB Plan ตามรูปแบบในรูปภาพ
    ''' </summary>
    Private Sub ExportDraftOTBSum(context As HttpContext)
        ' 1. รับ Filters (จาก QueryString)
        Dim year As Integer = 0
        Integer.TryParse(context.Request.QueryString("OTByear"), year)
        Dim company As String = If(String.IsNullOrWhiteSpace(context.Request.QueryString("OTBCompany")), "", context.Request.QueryString("OTBCompany").Trim())
        Dim segment As String = If(String.IsNullOrWhiteSpace(context.Request.QueryString("OTBSegment")), "", context.Request.QueryString("OTBSegment").Trim())

        ' (บังคับต้องมี Year)
        If year = 0 Then Throw New Exception("Year is required.")

        ' 2. ดึงข้อมูลจาก SP ใหม่
        Dim dt As New DataTable()
        Using conn As New SqlConnection(connectionString)
            conn.Open()
            Using cmd As New SqlCommand("SP_Get_OTB_Summary_Report", conn)
                cmd.CommandType = CommandType.StoredProcedure
                cmd.Parameters.AddWithValue("@Year", year)
                cmd.Parameters.AddWithValue("@Company", If(String.IsNullOrEmpty(company), DBNull.Value, company))
                cmd.Parameters.AddWithValue("@Segment", If(String.IsNullOrEmpty(segment), DBNull.Value, segment))

                Using adapter As New SqlDataAdapter(cmd)
                    adapter.Fill(dt)
                End Using
            End Using
        End Using

        ' 3. สร้างไฟล์ Excel
        ExcelPackage.License.SetNonCommercialOrganization("KingPower")

        Using package As New ExcelPackage()
            Dim ws = package.Workbook.Worksheets.Add("OTB Plan Summary")

            ws.Cells("A1").Value = "Categories"
            ws.Cells("A1:B2").Merge = True
            ws.Cells("A1:B2").Style.Fill.PatternType = ExcelFillStyle.Solid
            ws.Cells("A1:B2").Style.Fill.BackgroundColor.SetColor(Color.LightGray)

            ' Header C-N (OTB Plan) - สีฟ้า
            ws.Cells("C1").Value = "OTB Amount (Current Approved / Actual PO)"
            ws.Cells("C1:N1").Merge = True
            ws.Cells("C1:N1").Style.Fill.PatternType = ExcelFillStyle.Solid
            ws.Cells("C1:N1").Style.Fill.BackgroundColor.SetColor(Color.FromArgb(189, 215, 238)) ' Light Blue

            ' Header O-Z (TO-BE Amount) - สีส้ม
            ws.Cells("O1").Value = "TO-BE Amount (Revised)"
            ws.Cells("O1:Z1").Merge = True
            ws.Cells("O1:Z1").Style.Fill.PatternType = ExcelFillStyle.Solid
            ws.Cells("O1:Z1").Style.Fill.BackgroundColor.SetColor(Color.FromArgb(248, 203, 173)) ' Light Orange

            ' Header AA-AL (Diff) - สีเขียว
            ws.Cells("AA1").Value = "Diff"
            ws.Cells("AA1:AL1").Merge = True
            ws.Cells("AA1:AL1").Style.Fill.PatternType = ExcelFillStyle.Solid
            ws.Cells("AA1:AL1").Style.Fill.BackgroundColor.SetColor(Color.FromArgb(226, 239, 218)) ' Light Green

            ' --- สร้าง Headers แถวที่ 2 (Months) ---
            Dim months() As String = {"Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"}


            ' Loop 3 รอบ สำหรับ 3 Section
            For i As Integer = 0 To 2
                For m As Integer = 0 To 11
                    Dim colIndex As Integer = 3 + (i * 12) + m ' เริ่มที่คอลัมน์ C (3)
                    ws.Cells(2, colIndex).Value = months(m)

                    ' ใส่สีพื้นหลังให้ตรงกับ Section ด้านบน
                    ws.Cells(2, colIndex).Style.Fill.PatternType = ExcelFillStyle.Solid
                    If i = 0 Then ws.Cells(2, colIndex).Style.Fill.BackgroundColor.SetColor(Color.FromArgb(189, 215, 238))
                    If i = 1 Then ws.Cells(2, colIndex).Style.Fill.BackgroundColor.SetColor(Color.FromArgb(248, 203, 173))
                    If i = 2 Then ws.Cells(2, colIndex).Style.Fill.BackgroundColor.SetColor(Color.FromArgb(226, 239, 218))
                Next
            Next

            ' จัดรูปแบบ Header
            ws.Cells("A1:AL2").Style.HorizontalAlignment = ExcelHorizontalAlignment.Center
            ws.Cells("A1:AL2").Style.VerticalAlignment = ExcelVerticalAlignment.Center
            ws.Cells("A1:AL2").Style.Font.Bold = True
            ws.Cells("A1:AL2").Style.Border.Top.Style = ExcelBorderStyle.Thin
            ws.Cells("A1:AL2").Style.Border.Bottom.Style = ExcelBorderStyle.Thin
            ws.Cells("A1:AL2").Style.Border.Left.Style = ExcelBorderStyle.Thin
            ws.Cells("A1:AL2").Style.Border.Right.Style = ExcelBorderStyle.Thin

            ' --- ใส่ข้อมูล ---
            If dt.Rows.Count > 0 Then
                ws.Cells("A3").LoadFromDataTable(dt, False)

                ' Format Numbers
                ws.Cells(3, 3, dt.Rows.Count + 2, 38).Style.Numberformat.Format = "#,##0.00"
            End If

            ws.Cells.AutoFitColumns()

            ' --- ส่งไฟล์กลับ ---
            context.Response.Clear()
            context.Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            context.Response.AddHeader("content-disposition", $"attachment; filename=OTB_Plan_Summary_{year}_{DateTime.Now.ToString("yyyyMMdd")}.xlsx")
            context.Response.BinaryWrite(package.GetAsByteArray())
            context.Response.Flush()
            context.ApplicationInstance.CompleteRequest()

        End Using

    End Sub
    ' --- (END) NEW FUNCTION FOR SUMMARY EXPORT ---


    Private Function GenerateHtmlDraftTable(dt As DataTable) As String
        Dim sb As New StringBuilder()
        Dim budgetCalculator As New OTBBudgetCalculator()
        If dt.Rows.Count = 0 Then
            sb.Append("<tr><td colspan='21' class='text-center text-muted'>No Draft OTB records found</td></tr>")
        Else
            For i As Integer = 0 To dt.Rows.Count - 1

                Dim RunNo As String = If(dt.Rows(i)("RunNo") IsNot DBNull.Value, dt.Rows(i)("RunNo").ToString(), "")
                Dim CreateDT As String = If(dt.Rows(i)("CreateDT") IsNot DBNull.Value, dt.Rows(i)("CreateDT").ToString(), "")
                Dim OTBType As String = If(dt.Rows(i)("OTBType") IsNot DBNull.Value, dt.Rows(i)("OTBType").ToString(), "")
                Dim OTBYear As String = If(dt.Rows(i)("OTBYear") IsNot DBNull.Value, dt.Rows(i)("OTBYear").ToString(), "")
                Dim OTBMonth As String = If(dt.Rows(i)("OTBMonth") IsNot DBNull.Value, dt.Rows(i)("OTBMonth").ToString(), "")
                Dim MonthName As String = If(dt.Rows(i)("month_name_sh") IsNot DBNull.Value, dt.Rows(i)("month_name_sh").ToString(), "")
                Dim CateName As String = If(dt.Rows(i)("CateName") IsNot DBNull.Value, dt.Rows(i)("CateName").ToString(), "")
                Dim OTBCategory As String = If(dt.Rows(i)("OTBCategory") IsNot DBNull.Value, dt.Rows(i)("OTBCategory").ToString(), "")
                Dim CompanyName As String = If(dt.Rows(i)("CompanyName") IsNot DBNull.Value, dt.Rows(i)("CompanyName").ToString(), "")
                Dim OTBCompany As String = If(dt.Rows(i)("OTBCompany") IsNot DBNull.Value, dt.Rows(i)("OTBCompany").ToString(), "")
                Dim SegmentName As String = If(dt.Rows(i)("SegmentName") IsNot DBNull.Value, dt.Rows(i)("SegmentName").ToString(), "")
                Dim OTBSegment As String = If(dt.Rows(i)("OTBSegment") IsNot DBNull.Value, dt.Rows(i)("OTBSegment").ToString(), "")
                Dim OTBBrand As String = If(dt.Rows(i)("OTBBrand") IsNot DBNull.Value, dt.Rows(i)("OTBBrand").ToString(), "")
                Dim BrandName As String = If(dt.Rows(i)("BrandName") IsNot DBNull.Value, dt.Rows(i)("BrandName").ToString(), "")
                Dim OTBVendor As String = If(dt.Rows(i)("OTBVendor") IsNot DBNull.Value, dt.Rows(i)("OTBVendor").ToString(), "")
                Dim Vendor As String = If(dt.Rows(i)("Vendor") IsNot DBNull.Value, dt.Rows(i)("Vendor").ToString(), "")
                Dim Amount As String = "0.00"
                Dim CurrentBudgetAmount As String = "0.00"
                Dim Diff As String = "0.00"
                Dim amountValue As Decimal
                If Decimal.TryParse(If(dt.Rows(i)("Amount") IsNot DBNull.Value, dt.Rows(i)("Amount").ToString(), ""), amountValue) Then
                    Amount = amountValue.ToString("N2")
                End If
                Dim currentBudget As Decimal = budgetCalculator.CalculateCurrentApprovedBudget(OTBYear, OTBMonth, OTBCategory, OTBCompany, OTBSegment, OTBBrand, OTBVendor)
                CurrentBudgetAmount = currentBudget.ToString("N2")

                Dim diffamout As Decimal = amountValue - currentBudget
                Diff = diffamout.ToString("N2")

                Dim Batch As String = If(dt.Rows(i)("Batch") IsNot DBNull.Value, dt.Rows(i)("Batch").ToString(), "")
                Dim Remark As String = If(dt.Rows(i)("Remark") IsNot DBNull.Value, dt.Rows(i)("Remark").ToString(), "")
                Dim Version As String = If(dt.Rows(i)("Version") IsNot DBNull.Value, dt.Rows(i)("Version").ToString(), "")
                Dim OTBStatus As String = If(dt.Rows(i)("OTBStatus") IsNot DBNull.Value, dt.Rows(i)("OTBStatus").ToString(), "")
                Dim CreateBy As String = If(dt.Rows(i)("UploadBy") IsNot DBNull.Value, dt.Rows(i)("UploadBy").ToString(), "")


                sb.AppendFormat("<tr>
                                    <td><input type=""checkbox"" id=""checkselect{0}"" name=""checkselect"" class=""form-check-input"" checked></td>
                                    <td>{1}</td>
                                    <td>{2}</td>
                                    <td>{3}</td>
                                    <td>{4}</td>
                                    <td>{5}</td>
                                    <td>{6}</td>
                                    <td>{7}</td>
                                    <td>{8}</td>
                                    <td>{9}</td>
                                    <td>{10}</td>
                                    <td>{11}</td>
                                    <td>{12}</td>
                                    <td>{13}</td>    
                                    <td class=""text-end"">{14}</td>
                                    <td class=""text-end"">{15}</td>
                                    <td class=""text-end"">{16}</td>
                                    <td>{17}</td>
                                    <td>{18}</td>
                                    <td>{19}</td>
                                    <td>{20}</td>
                                </tr>",
                            HttpUtility.HtmlEncode(RunNo),
                            HttpUtility.HtmlEncode(CreateDT),
                            HttpUtility.HtmlEncode(OTBType),
                            HttpUtility.HtmlEncode(OTBYear),
                            HttpUtility.HtmlEncode(MonthName),
                            HttpUtility.HtmlEncode(OTBCategory),
                            HttpUtility.HtmlEncode(CateName),
                            HttpUtility.HtmlEncode(CompanyName),
                            HttpUtility.HtmlEncode(OTBSegment),
                            HttpUtility.HtmlEncode(SegmentName),
                            HttpUtility.HtmlEncode(OTBBrand),
                            HttpUtility.HtmlEncode(BrandName),
                            HttpUtility.HtmlEncode(OTBVendor),
                            HttpUtility.HtmlEncode(Vendor),
                            HttpUtility.HtmlEncode(CurrentBudgetAmount),
                            HttpUtility.HtmlEncode(Amount),
                            HttpUtility.HtmlEncode(Diff),
                            If(OTBStatus.Equals("Draft"), "<span class=""badge-draft"">Draft</span>", "<span class=""badge-approved"">Approved</span>"),
                            HttpUtility.HtmlEncode(Version),
                            HttpUtility.HtmlEncode(Remark),
                            HttpUtility.HtmlEncode(CreateBy))
            Next
        End If
        Return sb.ToString()
    End Function
    Private Function GenerateHtmlApprovedTable(dt As DataTable) As String
        Dim sb As New StringBuilder()


        If dt.Rows.Count = 0 Then
            sb.Append("<tr><td colspan='22' class='text-center text-muted'>No approved OTB records found</td></tr>")
        Else
            For Each row As DataRow In dt.Rows
                sb.Append("<tr>")

                ' Create Date
                sb.AppendFormat("<td class='date-cell'>{0}</td>",
                           If(row("CreateDate") IsNot DBNull.Value, Convert.ToDateTime(row("CreateDate")).ToString("dd/MM/yyyy HH:mm"), ""))

                ' Version
                sb.AppendFormat("<td class='version-cell'>{0}</td>", HttpUtility.HtmlEncode(If(row("Version") IsNot DBNull.Value, row("Version").ToString(), "")))

                ' Type
                Dim typeValue As String = If(row("Type") IsNot DBNull.Value, row("Type").ToString(), "")
                Dim typeClass As String = If(typeValue = "Original", "type-original", "type-revise")
                sb.AppendFormat("<td class='text-center {0}'>{1}</td>", typeClass, HttpUtility.HtmlEncode(typeValue))

                ' Year & Month
                sb.AppendFormat("<td class='text-center'>{0}</td>", row("Year"))
                sb.AppendFormat("<td class='text-center'>{0}</td>", GetMonthName(row("Month")))

                ' Category
                sb.AppendFormat("<td class='text-center'>{0}</td>", HttpUtility.HtmlEncode(If(row("Category") IsNot DBNull.Value, row("Category").ToString(), "")))
                sb.AppendFormat("<td>{0}</td>", HttpUtility.HtmlEncode(If(row("CategoryName") IsNot DBNull.Value, row("CategoryName").ToString(), "")))

                ' Company
                sb.AppendFormat("<td class='text-center'>{0}</td>", HttpUtility.HtmlEncode(If(row("CompanyName") IsNot DBNull.Value, row("CompanyName").ToString(), "")))

                ' Segment
                sb.AppendFormat("<td class='text-center'>{0}</td>", HttpUtility.HtmlEncode(If(row("Segment") IsNot DBNull.Value, row("Segment").ToString(), "")))
                sb.AppendFormat("<td>{0}</td>", HttpUtility.HtmlEncode(If(row("SegmentName") IsNot DBNull.Value, row("SegmentName").ToString(), "")))

                ' Brand
                sb.AppendFormat("<td class='text-center'>{0}</td>", HttpUtility.HtmlEncode(If(row("Brand") IsNot DBNull.Value, row("Brand").ToString(), "")))
                sb.AppendFormat("<td>{0}</td>", HttpUtility.HtmlEncode(If(row("BrandName") IsNot DBNull.Value, row("BrandName").ToString(), "")))

                ' Vendor
                sb.AppendFormat("<td class='text-center'>{0}</td>", HttpUtility.HtmlEncode(If(row("Vendor") IsNot DBNull.Value, row("Vendor").ToString(), "")))
                sb.AppendFormat("<td>{0}</td>", HttpUtility.HtmlEncode(If(row("VendorName") IsNot DBNull.Value, row("VendorName").ToString(), "")))

                ' Amount
                Dim amount As Decimal = If(row("Amount") IsNot DBNull.Value, Convert.ToDecimal(row("Amount")), 0)
                sb.AppendFormat("<td class='amount-cell'>{0}</td>", amount.ToString("N2"))

                ' Revised Diff
                If row("RevisedDiff") IsNot DBNull.Value Then
                    Dim revDiff As Decimal = Convert.ToDecimal(row("RevisedDiff"))
                    Dim diffClass As String = If(revDiff >= 0, "text-success", "text-danger")
                    sb.AppendFormat("<td class='amount-cell {0}'>{1}</td>", diffClass, revDiff.ToString("N2"))
                Else
                    sb.Append("<td class='amount-cell'>-</td>")
                End If

                ' Remark
                sb.AppendFormat("<td class='small'>{0}</td>", HttpUtility.HtmlEncode(If(row("Remark") IsNot DBNull.Value, row("Remark").ToString(), "")))

                ' Status
                sb.AppendFormat("<td class='text-center status-approved'>{0}</td>", HttpUtility.HtmlEncode(If(row("OTBStatus") IsNot DBNull.Value, row("OTBStatus").ToString(), "")))

                ' Approved Date
                sb.AppendFormat("<td class='date-cell'>{0}</td>",
                           If(row("ApprovedDate") IsNot DBNull.Value, Convert.ToDateTime(row("ApprovedDate")).ToString("dd/MM/yyyy HH:mm"), ""))
                ' Create By
                sb.AppendFormat("<td>{0}</td>", HttpUtility.HtmlEncode(If(row("CreateBy") IsNot DBNull.Value, row("CreateBy").ToString(), "")))
                ' Action By
                sb.AppendFormat("<td>{0}</td>", HttpUtility.HtmlEncode(If(row("ActionBy") IsNot DBNull.Value, row("ActionBy").ToString(), "")))
                ' SAP Status
                sb.AppendFormat("<td>{0}</td>", HttpUtility.HtmlEncode(If(row("SAPStatus") IsNot DBNull.Value, row("SAPStatus").ToString(), "")))

                sb.Append("</tr>")
            Next
        End If
        Return sb.ToString()
    End Function

    Private Function GenerateHtmlSwitchable(dt As DataTable) As String
        Dim sb As New StringBuilder()

        If dt.Rows.Count = 0 Then
            sb.Append("<tr><td colspan='30' class='text-center text-muted'>No switch OTB records found</td></tr>")
        Else
            For Each row As DataRow In dt.Rows
                sb.Append("<tr>")

                ' Create Date
                sb.AppendFormat("<td class='date-cell'>{0}</td>",
                       If(row("CreateDT") IsNot DBNull.Value, Convert.ToDateTime(row("CreateDT")).ToString("dd/MM/yyyy HH:mm"), ""))

                ' Type (Source)
                Dim typeValue As String = If(row("Type") IsNot DBNull.Value, row("Type").ToString(), "")
                Dim typeClass As String = ""
                Select Case typeValue
                    Case "Switch out"
                        typeClass = "type-switch-out"
                    Case "Carry out"
                        typeClass = "type-carry-out"
                    Case "Balance out"
                        typeClass = "type-balance-out"
                    Case "Extra"
                        typeClass = "type-extra"
                    Case Else
                        typeClass = "type-default"
                End Select
                sb.AppendFormat("<td class='text-center {0}'>{1}</td>", typeClass, HttpUtility.HtmlEncode(typeValue))
                ' Year & Month (Source) - เพิ่มการตรวจสอบ DBNull
                sb.AppendFormat("<td class='text-center'>{0}</td>",
                       If(row("Year") IsNot DBNull.Value, row("Year").ToString(), ""))
                sb.AppendFormat("<td class='text-center'>{0}</td>",
                       HttpUtility.HtmlEncode(If(row("MonthName") IsNot DBNull.Value, row("MonthName").ToString(), "")))
                ' Category (Source)
                sb.AppendFormat("<td class='text-center'>{0}</td>",
                       HttpUtility.HtmlEncode(If(row("Category") IsNot DBNull.Value, row("Category").ToString(), "")))
                sb.AppendFormat("<td>{0}</td>",
                       HttpUtility.HtmlEncode(If(row("CategoryName") IsNot DBNull.Value, row("CategoryName").ToString(), "")))
                ' Company (Source)
                'sb.AppendFormat("<td class='text-center'>{0}</td>",
                '       HttpUtility.HtmlEncode(If(row("Company") IsNot DBNull.Value, row("Company").ToString(), "")))
                sb.AppendFormat("<td>{0}</td>",
                       HttpUtility.HtmlEncode(If(row("CompanyName") IsNot DBNull.Value, row("CompanyName").ToString(), "")))



                ' Segment (Source)
                sb.AppendFormat("<td class='text-center'>{0}</td>",
                       HttpUtility.HtmlEncode(If(row("Segment") IsNot DBNull.Value, row("Segment").ToString(), "")))
                sb.AppendFormat("<td>{0}</td>",
                       HttpUtility.HtmlEncode(If(row("SegmentName") IsNot DBNull.Value, row("SegmentName").ToString(), "")))

                ' Brand (Source)
                sb.AppendFormat("<td class='text-center'>{0}</td>",
                       HttpUtility.HtmlEncode(If(row("Brand") IsNot DBNull.Value, row("Brand").ToString(), "")))
                sb.AppendFormat("<td>{0}</td>",
                       HttpUtility.HtmlEncode(If(row("BrandName") IsNot DBNull.Value, row("BrandName").ToString(), "")))

                ' Vendor (Source)
                sb.AppendFormat("<td class='text-center'>{0}</td>",
                       HttpUtility.HtmlEncode(If(row("Vendor") IsNot DBNull.Value, row("Vendor").ToString(), "")))
                sb.AppendFormat("<td>{0}</td>",
                       HttpUtility.HtmlEncode(If(row("VendorName") IsNot DBNull.Value, row("VendorName").ToString(), "")))



                ' Type (Source)
                Dim switchtypeValue As String = If(row("SwitchType") IsNot DBNull.Value, row("SwitchType").ToString(), "")
                Dim switchtypeClass As String = ""
                Select Case typeValue
                    Case "Switch In"
                        switchtypeClass = "type-switch-in"
                    Case "Carry In"
                        switchtypeClass = "type-carry-in"
                    Case "Balance In"
                        switchtypeClass = "type-balance-in"
                    Case Else
                        switchtypeClass = "type-default"
                End Select
                sb.AppendFormat("<td class='text-center {0}'>{1}</td>", switchtypeClass, HttpUtility.HtmlEncode(switchtypeValue))

                ' Switch Year & Month (Target) - เพิ่มการตรวจสอบ DBNull
                sb.AppendFormat("<td class='text-center'>{0}</td>",
                       If(row("SwitchYear") IsNot DBNull.Value, row("SwitchYear").ToString(), ""))
                sb.AppendFormat("<td class='text-center'>{0}</td>",
                       HttpUtility.HtmlEncode(If(row("SwitchMonthName") IsNot DBNull.Value, row("SwitchMonthName").ToString(), "")))
                ' Switch Category (Target)
                sb.AppendFormat("<td class='text-center'>{0}</td>",
                       HttpUtility.HtmlEncode(If(row("SwitchCategory") IsNot DBNull.Value, row("SwitchCategory").ToString(), "")))
                sb.AppendFormat("<td>{0}</td>",
                       HttpUtility.HtmlEncode(If(row("SwitchCategoryName") IsNot DBNull.Value, row("SwitchCategoryName").ToString(), "")))
                ' Switch Company (Target)
                sb.AppendFormat("<td class='text-center'>{0}</td>",
                       HttpUtility.HtmlEncode(If(row("SwitchCompanyName") IsNot DBNull.Value, row("SwitchCompanyName").ToString(), "")))


                ' Switch Segment (Target)
                sb.AppendFormat("<td class='text-center'>{0}</td>",
                       HttpUtility.HtmlEncode(If(row("SwitchSegment") IsNot DBNull.Value, row("SwitchSegment").ToString(), "")))
                sb.AppendFormat("<td>{0}</td>",
                       HttpUtility.HtmlEncode(If(row("SwitchSegmentName") IsNot DBNull.Value, row("SwitchSegmentName").ToString(), "")))

                ' Switch Brand (Target)
                sb.AppendFormat("<td class='text-center'>{0}</td>",
                       HttpUtility.HtmlEncode(If(row("SwitchBrand") IsNot DBNull.Value, row("SwitchBrand").ToString(), "")))
                sb.AppendFormat("<td class='text-center'>{0}</td>",
                       HttpUtility.HtmlEncode(If(row("SwitchBrandName") IsNot DBNull.Value, row("SwitchBrandName").ToString(), "")))

                ' Switch Vendor (Target)
                sb.AppendFormat("<td class='text-center'>{0}</td>",
                       HttpUtility.HtmlEncode(If(row("SwitchVendor") IsNot DBNull.Value, row("SwitchVendor").ToString(), "")))
                sb.AppendFormat("<td class='text-center'>{0}</td>",
                       HttpUtility.HtmlEncode(If(row("SwitchVendorName") IsNot DBNull.Value, row("SwitchVendorName").ToString(), "")))

                ' Budget Amount
                Dim budgetAmount As Decimal = If(row("BudgetAmount") IsNot DBNull.Value, Convert.ToDecimal(row("BudgetAmount")), 0)
                sb.AppendFormat("<td class='amount-cell'>{0}</td>", budgetAmount.ToString("N2"))
                ' Batch

                ' Create By
                sb.AppendFormat("<td>{0}</td>",
                       HttpUtility.HtmlEncode(If(row("CreateBy") IsNot DBNull.Value, row("CreateBy").ToString(), "")))

                sb.Append("</tr>")
            Next
        End If
        Return sb.ToString()
    End Function

    Private Function GetOTBData() As DataTable
        Dim dt As New DataTable()
        Using conn As New SqlConnection(connectionString)
            conn.Open()
            Dim query As String = "SELECT * FROM [Template_Upload_Draft_OTB]" ' ปรับเปลี่ยนตามตารางจริง
            Using cmd As New SqlCommand(query, conn)
                Using reader As SqlDataReader = cmd.ExecuteReader()
                    dt.Load(reader)
                End Using
            End Using
        End Using
        Return dt
    End Function

    Private Function GetOTBDraftDataWithFilter(OTBtype As String, OTByear As String, OTBmonth As String, OTBCompany As String, OTBCategory As String, OTBSegment As String, OTBBrand As String, OTBVendor As String) As DataTable
        Dim dt As New DataTable()
        Using conn As New SqlConnection(connectionString)
            conn.Open()
            Dim query As String = "SELECT  [RunNo]
                                          ,[OTBVendor]
                                          ,[Vendor]
                                          ,[OTBCompany]
                                          ,[CompanyName]
                                          ,[OTBMonth]
                                          ,[month_name_sh]
                                          ,[OTBCategory]
                                          ,[CateName]
                                          ,[OTBType]
                                          ,[OTBYear]
                                          ,[OTBBrand]
                                          ,[BrandName]
                                          ,[SegmentName]
                                          ,[OTBSegment]
                                          ,[Amount]
                                          ,[Batch]
                                          ,[CreateDT]
                                          ,[UploadBy]
                                          ,[Remark]
                                          ,[Version]
                                          ,[OTBStatus]
                                      FROM [BMS].[dbo].[View_OTB_Draft]
                                      WHERE (@OTBtype = '' OR OTBType = @OTBtype)
                                      AND (@OTByear = '' OR OTBYear = @OTByear)
                                      AND (@OTBmonth = '' OR OTBMonth = @OTBmonth)
                                      AND (@OTBCompany = '' OR OTBCompany = @OTBCompany)
                                      AND (@OTBCategory = '' OR OTBCategory = @OTBCategory)
                                      AND (@OTBSegment = '' OR OTBSegment = @OTBSegment)
                                      AND (@OTBBrand = '' OR OTBBrand = @OTBBrand)
                                      AND (@OTBVendor = '' OR OTBVendor = @OTBVendor)
                                      ORDER BY CreateDT DESC
                                    "
            Using cmd As New SqlCommand(query, conn)
                cmd.Parameters.AddWithValue("@OTBtype", OTBtype)
                cmd.Parameters.AddWithValue("@OTByear", OTByear)
                cmd.Parameters.AddWithValue("@OTBmonth", OTBmonth)
                cmd.Parameters.AddWithValue("@OTBCompany", OTBCompany)
                cmd.Parameters.AddWithValue("@OTBCategory", OTBCategory)
                cmd.Parameters.AddWithValue("@OTBSegment", OTBSegment)
                cmd.Parameters.AddWithValue("@OTBBrand", OTBBrand)
                cmd.Parameters.AddWithValue("@OTBVendor", OTBVendor)
                Using reader As SqlDataReader = cmd.ExecuteReader()
                    dt.Load(reader)
                End Using
            End Using
        End Using
        Return dt
    End Function

    Private Function GetOTBApproveDataWithFilter(OTBtype As String, OTByear As String, OTBmonth As String, OTBCompany As String, OTBCategory As String, OTBSegment As String, OTBBrand As String, OTBVendor As String, OTBVersion As String) As DataTable
        Dim dt As New DataTable()
        dt = ApprovedOTBManager.SearchApprovedOTB(OTBtype, OTByear, OTBmonth, OTBCompany, OTBCategory, OTBSegment, OTBBrand, OTBVendor, OTBVersion)

        Return dt
    End Function

    Private Function GetOTBSwitchDataWithFilter(OTBtype As String, OTByear As String, OTBmonth As String, OTBCompany As String, OTBCategory As String, OTBSegment As String, OTBBrand As String, OTBVendor As String) As DataTable
        Dim dt As New DataTable()
        dt = ApprovedOTBManager.SearchSwitchOTB(OTBtype, OTByear, OTBmonth, OTBCompany, OTBCategory, OTBSegment, OTBBrand, OTBVendor)

        Return dt
    End Function

    ' ===================================================================
    ' ===== START: REPLACEMENT LOGIC FOR ApproveDraftOTB (With MERGE) ===
    ' ===================================================================

    Private Sub ApproveDraftOTB(context As HttpContext)
        context.Response.ContentType = "application/json"
        Dim approvedBy As String = If(String.IsNullOrWhiteSpace(context.Request.Form("approvedBy")), "System", context.Request.Form("approvedBy").Trim())
        Dim remark As String = If(String.IsNullOrWhiteSpace(context.Request.Form("remark")), Nothing, context.Request.Form("remark").Trim())
        Dim responseJson As New Dictionary(Of String, Object)
        Dim masterinstance As New MasterDataUtil
        Try
            Dim budgetCalculator As New OTBBudgetCalculator()
            ' 1. Get selected IDs (Same as original code)
            Dim idsString As String = If(String.IsNullOrWhiteSpace(context.Request.Form("runNos")), "[]", context.Request.Form("runNos").Trim())
            If String.IsNullOrEmpty(idsString) OrElse idsString = "[]" Then
                Throw New Exception("No records selected for approval.")
            End If

            ' 2. Convert IDs to List(Of Integer) (Same as original code)
            Dim runNos As New List(Of Integer)
            Try
                Dim jsonArray As List(Of String) = JsonConvert.DeserializeObject(Of List(Of String))(idsString)
                For Each idStr As String In jsonArray
                    Dim id As Integer
                    If Integer.TryParse(idStr.Trim(), id) Then
                        runNos.Add(id)
                    End If
                Next
            Catch jsonEx As Exception
                For Each idStr As String In idsString.Split(","c)
                    Dim id As Integer
                    If Integer.TryParse(idStr.Trim().Replace("""", ""), id) Then
                        runNos.Add(id)
                    End If
                Next
            End Try

            If runNos.Count = 0 Then
                Throw New Exception("No valid RunNos provided.")
            End If

            ' 3. Get data for selected RunNos (This now includes RunNo in the DataTable)
            Dim draftData As DataTable = GetOTBDraftDataByRunNos(runNos)
            If draftData.Rows.Count = 0 Then
                Throw New Exception("Could not find draft records to approve.")
            End If

            ' 4. Build list to send to SAP API and create Key-to-RunNo map
            Dim plansToUpload As New List(Of OtbPlanUploadItem)()
            Dim sapKeyToRunNoMap As New Dictionary(Of String, Integer)
            Dim runNoToDataRowMap As New Dictionary(Of Integer, DataRow)

            For Each row As DataRow In draftData.Rows
                Dim currentRunNo As Integer = Convert.ToInt32(row("RunNo"))
                Dim OTBType As String = If(row("OTBType") IsNot DBNull.Value, row("OTBType").ToString(), "Original")
                Dim OTBVersion As String = If(row("Version") IsNot DBNull.Value, row("Version").ToString(), "A1")
                Dim OTBYear As String = If(row("OTBYear") IsNot DBNull.Value, row("OTBYear").ToString(), "")
                Dim OTBMonth As String = If(row("OTBMonth") IsNot DBNull.Value, row("OTBMonth").ToString(), "")
                Dim OTBCategory As String = If(row("OTBCategory") IsNot DBNull.Value, row("OTBCategory").ToString(), "")
                Dim OTBCompany As String = If(row("OTBCompany") IsNot DBNull.Value, row("OTBCompany").ToString(), "")
                Dim OTBSegment As String = If(row("OTBSegment") IsNot DBNull.Value, row("OTBSegment").ToString(), "")
                Dim OTBBrand As String = If(row("OTBBrand") IsNot DBNull.Value, row("OTBBrand").ToString(), "")
                Dim OTBVendor As String = If(row("OTBVendor") IsNot DBNull.Value, row("OTBVendor").ToString(), "")
                Dim toBeAmount As Decimal = 0
                Decimal.TryParse(If(row("Amount") IsNot DBNull.Value, row("Amount").ToString(), ""), toBeAmount)

                Dim amountToSendToSap As Decimal
                Dim amountStr As String = ""
                If OTBVersion.StartsWith("R", StringComparison.OrdinalIgnoreCase) OrElse OTBType.Equals("Revise", StringComparison.OrdinalIgnoreCase) Then
                    ' Version "Rn" (Revise): ให้ส่งยอด Diff
                    Dim currentBudget As Decimal = budgetCalculator.CalculateCurrentApprovedBudget(
                        OTBYear, OTBMonth, OTBCategory, OTBCompany, OTBSegment, OTBBrand, OTBVendor
                    )
                    ' ยอด Diff = ยอดใหม่ (To-Be) - ยอดที่ Approved ปัจจุบัน
                    amountToSendToSap = toBeAmount - currentBudget
                Else
                    ' Version "A1" (Original): ให้ส่งยอดเต็ม
                    amountToSendToSap = toBeAmount
                End If

                amountStr = amountToSendToSap.ToString("F2")

                Dim sapKey As String = String.Join("|",
                    OTBVersion,
                    OTBCompany,
                    OTBCategory,
                    OTBVendor,
                    OTBSegment,
                    OTBBrand,
                    amountStr, ' <-- [สำคัญ] ใช้ยอดที่คำนวณใหม่
                    OTBYear,
                    OTBMonth
                )

                ' Add to maps
                If Not sapKeyToRunNoMap.ContainsKey(sapKey) Then
                    sapKeyToRunNoMap.Add(sapKey, currentRunNo)
                End If
                If Not runNoToDataRowMap.ContainsKey(currentRunNo) Then
                    runNoToDataRowMap.Add(currentRunNo, row)
                End If

                plansToUpload.Add(New OtbPlanUploadItem With {
                    .Version = OTBVersion,
                    .CompCode = OTBCompany,
                    .Category = OTBCategory,
                    .VendorCode = OTBVendor,
                    .SegmentCode = OTBSegment,
                    .BrandCode = OTBBrand,
                    .Amount = amountStr, ' <--- [สำคัญ] ใช้ยอดที่คำนวณใหม่
                    .Year = OTBYear,
                    .Month = OTBMonth,
                    .Remark = If(row("Remark") IsNot DBNull.Value, row("Remark").ToString(), "")
                })
            Next

            ' 5. Call SAP API
            Dim sapResponse As SapApiResponse(Of SapUploadResultItem) = Task.Run(Async Function()
                                                                                     Return Await SapApiHelper.UploadOtbPlanAsync(plansToUpload)
                                                                                 End Function).Result

            ' 6. Check for catastrophic failure (no response)
            If sapResponse Is Nothing Then
                Throw New Exception("No response received from SAP API. Approval aborted.")
            End If

            ' 7. (NEW) Build Detailed Result List REGARDLESS of success/failure
            Dim detailedResults As New List(Of Dictionary(Of String, Object))
            Dim sapSuccessResults As New List(Of SapUploadResultItem) ' List for DB update

            If sapResponse.Results Is Nothing Then
                Throw New Exception("SAP response was successful, but returned no results array. Approval aborted.")
            End If

            For Each sapResult As SapUploadResultItem In sapResponse.Results
                ' Re-create the key from the SAP result to find its RunNo
                Dim sapKey As String = String.Join("|",
                    sapResult.Version,
                    sapResult.CompCode,
                    sapResult.Category,
                    sapResult.VendorCode,
                    sapResult.SegmentCode,
                    sapResult.BrandCode,
                    sapResult.Amount,
                    sapResult.Year,
                    sapResult.Month
                )

                Dim runNoToFind As Integer = -1
                If sapKeyToRunNoMap.ContainsKey(sapKey) Then
                    runNoToFind = sapKeyToRunNoMap(sapKey)
                End If

                Dim resultRow As New Dictionary(Of String, Object)

                If runNoToFind <> -1 AndAlso runNoToDataRowMap.ContainsKey(runNoToFind) Then
                    ' Found matching Draft data
                    Dim draftRow As DataRow = runNoToDataRowMap(runNoToFind)

                    resultRow.Add("OTBYear", draftRow("OTBYear"))
                    resultRow.Add("OTBMonth", draftRow("OTBMonth"))
                    resultRow.Add("OTBCategory", draftRow("OTBCategory"))
                    resultRow.Add("CateName", draftRow("CateName"))
                    resultRow.Add("CompanyName", draftRow("CompanyName"))
                    resultRow.Add("OTBSegment", draftRow("OTBSegment"))
                    resultRow.Add("SegmentName", draftRow("SegmentName"))
                    resultRow.Add("OTBBrand", draftRow("OTBBrand"))
                    resultRow.Add("BrandName", draftRow("BrandName"))
                    resultRow.Add("OTBVendor", draftRow("OTBVendor"))
                    resultRow.Add("Vendor", draftRow("Vendor"))
                    resultRow.Add("Amount", draftRow("Amount"))
                    resultRow.Add("Remark", draftRow("Remark"))
                Else
                    ' Data mismatch - should not happen, but good to handle
                    resultRow.Add("OTBYear", sapResult.Year)
                    resultRow.Add("OTBMonth", sapResult.Month)
                    resultRow.Add("OTBCategory", sapResult.Category)
                    resultRow.Add("CateName", masterinstance.GetCategoryName(sapResult.Category))
                    resultRow.Add("CompanyName", masterinstance.GetCompanyName(sapResult.CompCode))
                    resultRow.Add("OTBSegment", sapResult.SegmentCode)
                    resultRow.Add("SegmentName", masterinstance.GetSegmentName(sapResult.SegmentCode))
                    resultRow.Add("OTBBrand", sapResult.BrandCode)
                    resultRow.Add("BrandName", masterinstance.GetBrandName(sapResult.BrandCode))
                    resultRow.Add("OTBVendor", sapResult.VendorCode)
                    resultRow.Add("Vendor", masterinstance.GetVendorName(sapResult.VendorCode))
                    resultRow.Add("Amount", sapResult.Amount)
                    resultRow.Add("Remark", sapResult.Remark)
                End If

                ' Add SAP Results
                resultRow.Add("SAP_MessageType", sapResult.MessageType)
                resultRow.Add("SAP_Message", sapResult.Message)
                detailedResults.Add(resultRow)

                ' If this item was successful, add it to the list for DB update
                If sapResult.MessageType.Equals("S", StringComparison.OrdinalIgnoreCase) Then
                    sapSuccessResults.Add(sapResult)
                End If
            Next

            ' 8. Check for Full Success vs. Partial/Total Failure
            If sapResponse.Status.Total = sapResponse.Status.Success AndAlso sapSuccessResults.Count = sapResponse.Status.Total Then
                ' *** FULL SUCCESS ***
                ' Proceed with Database Update
                Dim updateCount As Integer = 0
                Using conn As New SqlConnection(connectionString)
                    conn.Open()
                    Using transaction As SqlTransaction = conn.BeginTransaction()
                        Try
                            For Each successResult In sapSuccessResults
                                ' Re-create the key to find RunNo
                                Dim sapKey As String = String.Join("|",
                                    successResult.Version, successResult.CompCode, successResult.Category,
                                    successResult.VendorMap, successResult.SegmentCodeMap, successResult.BrandCode,
                                    successResult.Amount, successResult.Year, successResult.Month
                                )

                                If sapKeyToRunNoMap.ContainsKey(sapKey) Then
                                    Dim runNoToUpdate As Integer = sapKeyToRunNoMap(sapKey)

                                    ' 1. (Existing Code) Update Template_Upload_Draft_OTB
                                    Dim updateQuery As String = "
                                        UPDATE [dbo].[Template_Upload_Draft_OTB]
                                        SET 
                                            [OTBStatus] = @OTBStatus,
                                            [UpdateBy] = @ApprovedBy,
                                            [UpdateDT] = GETDATE(),
                                            [SAPStatus] = @SAPStatus,
                                            [SAPErrorMessage] = @SAPErrorMessage
                                        WHERE 
                                            [RunNo] = @RunNo
                                            AND (OTBStatus IS NULL OR OTBStatus = 'Draft')
                                    "

                                    Using cmdUpdate As New SqlCommand(updateQuery, conn, transaction)
                                        cmdUpdate.Parameters.AddWithValue("@OTBStatus", "Approved")
                                        cmdUpdate.Parameters.AddWithValue("@ApprovedBy", approvedBy)
                                        cmdUpdate.Parameters.AddWithValue("@SAPStatus", successResult.MessageType)
                                        cmdUpdate.Parameters.AddWithValue("@SAPErrorMessage", If(String.IsNullOrEmpty(successResult.Message), DBNull.Value, successResult.Message))
                                        cmdUpdate.Parameters.AddWithValue("@RunNo", runNoToUpdate)
                                        updateCount += cmdUpdate.ExecuteNonQuery()
                                    End Using

                                    ' 2. *** MODIFIED LOGIC: UPSERT into OTB_Transaction using MERGE ***
                                    If runNoToDataRowMap.ContainsKey(runNoToUpdate) Then
                                        Dim approvedRow As DataRow = runNoToDataRowMap(runNoToUpdate)

                                        ' Calculate RevisedDiff
                                        Dim calc_Year As String = approvedRow("OTBYear").ToString()
                                        Dim calc_Month As String = approvedRow("OTBMonth").ToString()
                                        Dim calc_Category As String = approvedRow("OTBCategory").ToString()
                                        Dim calc_Company As String = approvedRow("OTBCompany").ToString()
                                        Dim calc_Segment As String = approvedRow("OTBSegment").ToString()
                                        Dim calc_Brand As String = approvedRow("OTBBrand").ToString()
                                        Dim calc_Vendor As String = approvedRow("OTBVendor").ToString()
                                        Dim calc_Amount As Decimal = Convert.ToDecimal(approvedRow("Amount"))
                                        Dim calc_Type As String = approvedRow("OTBType").ToString()
                                        Dim calc_Version As String = approvedRow("Version").ToString()

                                        Dim revisedDiffValue As Decimal = 0
                                        If calc_Type.Equals("Revise", StringComparison.OrdinalIgnoreCase) Then
                                            Dim currentBudget As Decimal = budgetCalculator.CalculateCurrentApprovedBudget(
                                                calc_Year, calc_Month, calc_Category, calc_Company, calc_Segment, calc_Brand, calc_Vendor)
                                            revisedDiffValue = calc_Amount - currentBudget
                                        End If

                                        ' (Schema based on data_BMS.png)
                                        Dim mergeQuery As String = "
                                            MERGE INTO [dbo].[OTB_Transaction] AS T
                                            USING (
                                                SELECT 
                                                    @Type AS [Type], @Year AS [Year], @Month AS [Month], 
                                                    @Category AS [Category], @Company AS [Company], @Segment AS [Segment], 
                                                    @Brand AS [Brand], @Vendor AS [Vendor], @Version AS [Version]
                                            ) AS S
                                            ON (
                                                T.[Type] = S.[Type] AND
                                                T.[Year] = S.[Year] AND
                                                T.[Month] = S.[Month] AND
                                                T.[Category] = S.[Category] AND
                                                T.[Company] = S.[Company] AND
                                                T.[Segment] = S.[Segment] AND
                                                T.[Brand] = S.[Brand] AND
                                                T.[Vendor] = S.[Vendor] AND
                                                T.[Version] = S.[Version]
                                            )
                                            WHEN MATCHED THEN
                                                UPDATE SET
                                                    T.[Amount] = @Amount,
                                                    T.[RevisedDiff] = @RevisedDiff,
                                                    T.[Remark] = @Remark,
                                                    T.[ApprovedDate] = GETDATE(),
                                                    T.[SAPDate] = GETDATE(),
                                                    T.[ActionBy] = @ActionBy,
                                                    T.[DraftID] = @DraftID,
                                                    T.[SAPStatus] = @SAPStatus,
                                                    T.[SAPErrorMessage] = @SAPErrorMessage,
                                                    T.[CategoryName] = @CategoryName,
                                                    T.[SegmentName] = @SegmentName,
                                                    T.[BrandName] = @BrandName,
                                                    T.[VendorName] = @VendorName,
                                                    T.[OTBStatus] = 'Approved'
                                            WHEN NOT MATCHED BY TARGET THEN
                                                INSERT (
                                                    [CreateDate], [Type], [Year], [Month], [Category], [CategoryName],
                                                    [Company], [Segment], [SegmentName], [Brand], [BrandName],
                                                    [Vendor], [VendorName], [Amount], [RevisedDiff], [Remark],
                                                    [OTBStatus], [ApprovedDate], [SAPDate], [ActionBy], [DraftID],
                                                    [SAPStatus], [SAPErrorMessage], [Version]
                                                )
                                                VALUES (
                                                    GETDATE(), @Type, @Year, @Month, @Category, @CategoryName,
                                                    @Company, @Segment, @SegmentName, @Brand, @BrandName,
                                                    @Vendor, @VendorName, @Amount, @RevisedDiff, @Remark,
                                                    'Approved', GETDATE(), GETDATE(), @ActionBy, @DraftID,
                                                    @SAPStatus, @SAPErrorMessage, @Version
                                                );
                                        "

                                        Using cmdMerge As New SqlCommand(mergeQuery, conn, transaction)
                                            ' Key Parameters (for ON clause)
                                            cmdMerge.Parameters.AddWithValue("@Type", approvedRow("OTBType"))
                                            cmdMerge.Parameters.AddWithValue("@Year", approvedRow("OTBYear"))
                                            cmdMerge.Parameters.AddWithValue("@Month", approvedRow("OTBMonth"))
                                            cmdMerge.Parameters.AddWithValue("@Category", approvedRow("OTBCategory"))
                                            cmdMerge.Parameters.AddWithValue("@Company", approvedRow("OTBCompany"))
                                            cmdMerge.Parameters.AddWithValue("@Segment", approvedRow("OTBSegment"))
                                            cmdMerge.Parameters.AddWithValue("@Brand", approvedRow("OTBBrand"))
                                            cmdMerge.Parameters.AddWithValue("@Vendor", approvedRow("OTBVendor"))
                                            cmdMerge.Parameters.AddWithValue("@Version", approvedRow("Version"))

                                            ' Data Parameters (for INSERT/UPDATE)
                                            cmdMerge.Parameters.AddWithValue("@Amount", calc_Amount)
                                            cmdMerge.Parameters.AddWithValue("@RevisedDiff", revisedDiffValue)
                                            cmdMerge.Parameters.AddWithValue("@Remark", approvedRow("Remark"))
                                            cmdMerge.Parameters.AddWithValue("@ActionBy", approvedBy)
                                            cmdMerge.Parameters.AddWithValue("@DraftID", runNoToUpdate)
                                            cmdMerge.Parameters.AddWithValue("@SAPStatus", successResult.MessageType)
                                            cmdMerge.Parameters.AddWithValue("@SAPErrorMessage", If(String.IsNullOrEmpty(successResult.Message), DBNull.Value, successResult.Message))


                                            ' Parameters for UPDATE/INSERT (Names)
                                            cmdMerge.Parameters.AddWithValue("@CategoryName", approvedRow("CateName"))
                                            cmdMerge.Parameters.AddWithValue("@SegmentName", approvedRow("SegmentName"))
                                            cmdMerge.Parameters.AddWithValue("@BrandName", approvedRow("BrandName"))
                                            cmdMerge.Parameters.AddWithValue("@VendorName", approvedRow("Vendor"))

                                            cmdMerge.ExecuteNonQuery()
                                        End Using
                                    Else
                                        Throw New Exception($"Critical Error: Could not find original DataRow for RunNo '{runNoToUpdate}'.")
                                    End If
                                Else
                                    Throw New Exception($"Critical Error: Could not map SAP success key '{sapKey}' back to a RunNo.")
                                End If
                            Next

                            transaction.Commit()

                            ' 9. Send success response (Full Success)
                            responseJson("success") = True
                            responseJson("action") = "preview"
                            responseJson("message") = $"Successfully approved and updated {updateCount} / {sapSuccessResults.Count} records in the database."
                            responseJson("detailedResults") = detailedResults

                        Catch ex As Exception
                            transaction.Rollback()
                            Throw New Exception("Database update failed after SAP success: " & ex.Message)
                        End Try
                    End Using
                End Using
            Else
                ' *** PARTIAL OR TOTAL FAILURE ***
                ' Do NOT update database.
                ' 9. Send failure response (Partial/Total Failure)
                responseJson("success") = False
                responseJson("action") = "preview"
                responseJson("message") = $"SAP processing failed or was incomplete. Total: {sapResponse.Status.Total}, Success: {sapResponse.Status.Success}, Error: {sapResponse.Status.ErrorCount}. No records were updated in the database."
                responseJson("detailedResults") = detailedResults
            End If

            context.Response.Write(JsonConvert.SerializeObject(responseJson))

        Catch ex As Exception
            ' Catch all errors (from validation, SAP call, or DB update)
            responseJson("success") = False
            responseJson("action") = "error" ' General error
            responseJson("message") = "Error approving records: " & ex.Message
            context.Response.StatusCode = 500
            context.Response.Write(JsonConvert.SerializeObject(responseJson))
        End Try
    End Sub
    ' ===================================================================
    ' ===== END: REPLACEMENT LOGIC ======================================
    ' ===================================================================

    Private Sub DeleteDraftOTB(context As HttpContext)
        Try

            context.Response.ContentType = "application/json"
            Dim runNoJson As String = If(String.IsNullOrWhiteSpace(context.Request.Form("runNos")), "", context.Request.Form("runNos").Trim())

            If String.IsNullOrEmpty(runNoJson) Then
                Dim errorResponse As New With {
                .success = False,
                .message = "No records selected for approval"
            }
                context.Response.Write(JsonConvert.SerializeObject(errorResponse))
                Return
            End If

            Dim runNos As New List(Of Integer)
            Try

                Dim jsonArray As List(Of String) = JsonConvert.DeserializeObject(Of List(Of String))(runNoJson)

                For Each idStr As String In jsonArray
                    Dim id As Integer
                    If Integer.TryParse(idStr.Trim(), id) Then
                        runNos.Add(id)
                    End If
                Next
            Catch jsonEx As Exception
                For Each idStr As String In runNoJson.Split(","c)
                    Dim id As Integer
                    If Integer.TryParse(idStr.Trim().Replace("""", ""), id) Then
                        runNos.Add(id)
                    End If
                Next
            End Try

            If runNos.Count = 0 Then
                Dim errorResponse As New With {
                .success = False,
                .message = "No valid records to delete"
            }
                context.Response.Write(JsonConvert.SerializeObject(errorResponse))
                Return
            End If

            Dim result As Dictionary(Of String, Object) = ApprovedOTBManager.DeleteDraftOTB(runNos)
            Dim response As New With {
                   .success = If(result("Status").ToString() = "Success", True, False),
                   .message = If(result.ContainsKey("Message"), result("Message").ToString(), $"Successfully approved {result("DeletedCount")} records"),
                   .deletedCount = result("DeletedCount")
            }
            context.Response.Write(JsonConvert.SerializeObject(response))
        Catch ex As Exception
            Dim errorResponse As New With {
                .success = False,
                .message = "Error deleting draft OTB: " + ex.Message
            }
            context.Response.Write(JsonConvert.SerializeObject(errorResponse))

        End Try
    End Sub

    Private Function GetMonthName(month As Object) As String
        If month Is DBNull.Value Then Return ""

        Dim monthInt As Integer = Convert.ToInt32(month)
        Select Case monthInt
            Case 1 : Return "Jan"
            Case 2 : Return "Feb"
            Case 3 : Return "Mar"
            Case 4 : Return "Apr"
            Case 5 : Return "May"
            Case 6 : Return "Jun"
            Case 7 : Return "Jul"
            Case 8 : Return "Aug"
            Case 9 : Return "Sep"
            Case 10 : Return "Oct"
            Case 11 : Return "Nov"
            Case 12 : Return "Dec"
            Case Else : Return monthInt.ToString()
        End Select
    End Function

    ' <summary>
    ' (ฟังก์ชันใหม่) ดึงข้อมูล Draft OTB จาก List ของ RunNo
    ' </summary>
    ' ===================================================================
    ' ===== START: MODIFIED FUNCTION GetOTBDraftDataByRunNos ==========
    ' ===================================================================
    Private Function GetOTBDraftDataByRunNos(runNos As List(Of Integer)) As DataTable
        Dim dt As New DataTable()
        If runNos Is Nothing OrElse runNos.Count = 0 Then
            Return dt
        End If

        Dim paramNames As New List(Of String)()
        For i As Integer = 0 To runNos.Count - 1
            paramNames.Add($"@p{i}")
        Next

        ' ===== MODIFIED QUERY (เพิ่ม [RunNo] และ Join Master Data เพื่อ Preview) =====
        Dim query As String = $"
            SELECT DISTINCT
                d.RunNo, d.Version, d.OTBCompany, d.OTBCategory, d.OTBVendor, 
                d.OTBSegment, d.OTBBrand, d.Amount, d.OTBYear, 
                d.OTBMonth, d.Remark, d.OTBType,
                c.Category AS CateName,
                co.CompanyNameShort AS CompanyName,
                s.SegmentName,
                b.[Brand Name] AS BrandName,
                v.Vendor
            FROM [BMS].[dbo].[View_OTB_Draft] d
            LEFT JOIN [dbo].[MS_Category] c ON d.OTBCategory = c.Cate
            LEFT JOIN [dbo].[MS_Company] co ON d.OTBCompany = co.CompanyCode
            LEFT JOIN [dbo].[MS_Segment] s ON d.OTBSegment = s.SegmentCode
            LEFT JOIN [dbo].[MS_Brand] b ON d.OTBBrand = b.[Brand Code]
            LEFT JOIN [dbo].[MS_Vendor] v ON d.OTBVendor = v.VendorCode AND d.OTBSegment = v.SegmentCode
            WHERE d.RunNo IN ({String.Join(",", paramNames)})"
        ' ===== END: MODIFIED QUERY =====

        Using conn As New SqlConnection(connectionString)
            conn.Open()
            Using cmd As New SqlCommand(query, conn)
                For i As Integer = 0 To runNos.Count - 1
                    cmd.Parameters.AddWithValue(paramNames(i), runNos(i))
                Next

                Using reader As SqlDataReader = cmd.ExecuteReader()
                    dt.Load(reader)
                End Using
            End Using
        End Using
        Return dt
    End Function
    ' ===================================================================
    ' ===== END: MODIFIED FUNCTION ======================================
    ' ===================================================================

    Private Sub HandleExportOTBMovement(context As HttpContext)
        ' 1. รับ Parameters
        Dim year As String = context.Request.QueryString("OTByear")
        Dim month As String = context.Request.QueryString("OTBmonth")
        Dim company As String = context.Request.QueryString("OTBCompany")
        Dim category As String = context.Request.QueryString("OTBCategory")
        Dim segment As String = context.Request.QueryString("OTBSegment")
        Dim brand As String = context.Request.QueryString("OTBBrand")
        Dim vendor As String = context.Request.QueryString("OTBVendor")
        Dim masterinstance As New MasterDataUtil
        If String.IsNullOrEmpty(year) Then Throw New Exception("Year is required.")

        ' 2. เตรียมข้อมูล
        ' 2.1 เรียก Calculator (ซึ่งจะโหลด Approved Transaction ทั้งหมดเข้า Memory)
        Dim budgetCalc As New OTBBudgetCalculator()

        ' 2.2 โหลด Draft PO และ Actual PO เข้า Memory เพื่อความเร็วในการ Lookup
        Dim dtDraftPO As DataTable = LoadDraftPOForReport(year, month, company, category, segment, brand, vendor)
        Dim dtActualPO As DataTable = LoadActualPOForReport(year, month, company, category, segment, brand, vendor)

        '

        ' 2.3 หา Distinct Keys ทั้งหมดที่มีการเคลื่อนไหว (จาก OTB Transaction, Draft PO, Actual PO)
        ' เพื่อให้มั่นใจว่าแสดงครบทุกบรรทัดที่มีข้อมูล แม้จะไม่มี Budget แต่มี PO
        Dim dtKeys As DataTable = GetDistinctKeysForReport(year, month, company, category, segment, brand, vendor)

        ' 3. สร้าง DataTable ผลลัพธ์
        Dim dtExport As New DataTable("OTBMovement")
        dtExport.Columns.Add("Year")
        dtExport.Columns.Add("Month")
        dtExport.Columns.Add("Cate")
        dtExport.Columns.Add("CateName")
        dtExport.Columns.Add("Company")
        dtExport.Columns.Add("CompanyName")
        dtExport.Columns.Add("Segment")
        dtExport.Columns.Add("SegmentName")
        dtExport.Columns.Add("Brand")
        dtExport.Columns.Add("BrandName")
        dtExport.Columns.Add("Vendor")
        dtExport.Columns.Add("VendorName")

        dtExport.Columns.Add("Budget Approved", GetType(Decimal))
        dtExport.Columns.Add("Revised Diff", GetType(Decimal))
        dtExport.Columns.Add("Extra", GetType(Decimal))

        dtExport.Columns.Add("Switch in", GetType(Decimal))
        dtExport.Columns.Add("Balance in", GetType(Decimal))
        dtExport.Columns.Add("Carry in", GetType(Decimal))

        dtExport.Columns.Add("Switch out", GetType(Decimal))
        dtExport.Columns.Add("Balance out", GetType(Decimal))
        dtExport.Columns.Add("Carry out", GetType(Decimal))

        dtExport.Columns.Add("Total Budget Approved", GetType(Decimal))
        dtExport.Columns.Add("Actual PO", GetType(Decimal))
        dtExport.Columns.Add("Draft PO", GetType(Decimal))
        dtExport.Columns.Add("Total Actual + Draft PO", GetType(Decimal))
        dtExport.Columns.Add("Remaining", GetType(Decimal))

        ' 4. Loop คำนวณ
        For Each rowKey As DataRow In dtKeys.Rows
            Dim kYear As String = rowKey("Year").ToString()
            Dim kMonth As String = rowKey("Month").ToString() ' เป็นตัวเลข
            Dim kCate As String = rowKey("Category").ToString()
            Dim kCateName As String = masterinstance.GetCategoryName(rowKey("Category").ToString())
            Dim kComp As String = rowKey("Company").ToString()
            Dim kCompName As String = masterinstance.GetCompanyName(rowKey("Company").ToString())
            Dim kSeg As String = rowKey("Segment").ToString()
            Dim kSegName As String = masterinstance.GetSegmentName(rowKey("Segment").ToString())
            Dim kBrand As String = rowKey("Brand").ToString()
            Dim kBrandName As String = masterinstance.GetBrandName(rowKey("Brand").ToString())
            Dim kVendor As String = rowKey("Vendor").ToString()
            Dim kVendorName As String = masterinstance.GetVendorName(rowKey("Vendor").ToString())
            Dim kMonthName As String = GetMonthName(kMonth)

            ' 4.1 คำนวณ Budget (ใช้ Logic ของ OTBBudgetCalculator ที่มีอยู่แล้ว)
            Dim budgetBreakdown = budgetCalc.GetBudgetBreakdown(kYear, kMonth, kCate, kComp, kSeg, kBrand, kVendor)

            ' 4.2 คำนวณ PO (ใช้ Compute จาก DataTable ใน Memory)
            Dim filterPO As String = $"Year = '{kYear}' AND Month = '{kMonth}' AND Category = '{kCate}' AND Company = '{kComp}' AND Segment = '{kSeg}' AND Brand = '{kBrand}' AND Vendor = '{kVendor}'"

            Dim sumDraftObj As Object = dtDraftPO.Compute("SUM(Amount)", filterPO)
            Dim sumActualObj As Object = dtActualPO.Compute("SUM(Amount)", filterPO)

            Dim sumDraft As Decimal = If(IsDBNull(sumDraftObj), 0, Convert.ToDecimal(sumDraftObj))
            Dim sumActual As Decimal = If(IsDBNull(sumActualObj), 0, Convert.ToDecimal(sumActualObj))

            Dim totalBudget As Decimal = budgetBreakdown("Total")
            Dim totalUsage As Decimal = sumDraft + sumActual
            Dim remaining As Decimal = totalBudget - totalUsage

            ' 4.3 Add Row (เฉพาะแถวที่มีค่าอย่างน้อย 1 อย่าง)
            If totalBudget <> 0 OrElse totalUsage <> 0 Then
                dtExport.Rows.Add(
                    kYear,
                    kMonthName,
                    kCate,
                    kCateName,
                    kComp,
                    kCompName,
                    kSeg,
                    kSegName,
                    kBrand,
                    kBrandName,
                    kVendor,
                    kVendorName,
                    budgetBreakdown("Original"),
                    budgetBreakdown("RevDiff"),
                    budgetBreakdown("Extra"),
                    budgetBreakdown("SwitchIn"),
                    budgetBreakdown("BalanceIn"),
                    budgetBreakdown("CarryIn"),
                    budgetBreakdown("SwitchOut"),
                    budgetBreakdown("BalanceOut"),
                    budgetBreakdown("CarryOut"),
                    totalBudget,
                    sumActual,
                    sumDraft,
                    totalUsage,
                    remaining
                )
            End If
        Next

        ' 5. สร้าง Excel File
        GenerateExcelOTBMovement(context, dtExport, $"OTB_Movement_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx")
    End Sub

    ' --- Helper Methods for Data Loading ---
    Private Function LoadDraftPOForReport(year As String, month As String, company As String, category As String, segment As String, brand As String, vendor As String) As DataTable
        Dim dt As New DataTable()
        Using conn As New SqlConnection(connectionString)
            conn.Open()
            ' ดึง Draft PO ทั้งหมดที่ยังไม่ Cancelled และยังไม่ Match กับ Actual (หรือจะรวมหมดแล้วแต่ Business Rule)
            ' ปกติถ้า Match แล้ว มูลค่าจะไปอยู่ที่ Actual PO, ถ้ายังไม่ Match อยู่ที่ Draft
            ' กรณีนี้เราจะดึงเฉพาะ Draft ที่ Status != 'Cancelled' และ Actual_PO_Ref IS NULL
            Dim query As String = "
                SELECT PO_Year AS Year, PO_Month AS Month, Company_Code AS Company, Category_Code AS Category, 
                       Segment_Code AS Segment, Brand_Code AS Brand, Vendor_Code AS Vendor, Amount_THB AS Amount
                FROM [BMS].[dbo].[Draft_PO_Transaction]
                WHERE ISNULL(Status, '') IN ('Draft', 'Edited')
                AND (@Year IS NULL OR PO_Year = @Year)
                AND (@Month IS NULL OR PO_Month = @Month)
                AND (@Company IS NULL OR Company_Code = @Company)
                AND (@Category IS NULL OR Category_Code = @Category)
                AND (@Segment IS NULL OR Segment_Code = @Segment)
                AND (@Brand IS NULL OR Brand_Code = @Brand)
                AND (@Vendor IS NULL OR Vendor_Code = @Vendor)
                "
            ' (Add filters parameters as needed logic similar to other functions)
            Using cmd As New SqlCommand(query, conn)
                cmd.Parameters.AddWithValue("@Year", If(String.IsNullOrEmpty(year), DBNull.Value, year))
                cmd.Parameters.AddWithValue("@Month", If(String.IsNullOrEmpty(month), DBNull.Value, month))
                cmd.Parameters.AddWithValue("@Company", If(String.IsNullOrEmpty(company), DBNull.Value, company))
                cmd.Parameters.AddWithValue("@Category", If(String.IsNullOrEmpty(category), DBNull.Value, category))
                cmd.Parameters.AddWithValue("@Segment", If(String.IsNullOrEmpty(segment), DBNull.Value, segment))
                cmd.Parameters.AddWithValue("@Brand", If(String.IsNullOrEmpty(brand), DBNull.Value, brand))
                cmd.Parameters.AddWithValue("@Vendor", If(String.IsNullOrEmpty(vendor), DBNull.Value, vendor))
                Using da As New SqlDataAdapter(cmd)
                    da.Fill(dt)
                End Using
            End Using
        End Using
        Return dt
    End Function

    Private Function LoadActualPOForReport(year As String, month As String, company As String, category As String, segment As String, brand As String, vendor As String) As DataTable
        Dim dt As New DataTable()
        Using conn As New SqlConnection(connectionString)
            conn.Open()
            ' ดึง Actual PO (จาก Staging หรือ View)
            Dim query As String = "
                SELECT Otb_Year AS Year, Otb_Month AS Month, Company_Code AS Company, Category_Code As Category, 
                       SUBSTRING(Segment_Code, 2, CASE WHEN LEN(ISNULL(Segment_Code, '')) > 2 THEN LEN(Segment_Code) - 2 ELSE 0 END) As Segment,Brand_Code As Brand,Vendor_Code As Vendor, Amount_THB AS Amount
                FROM [BMS].[dbo].[Actual_PO_Summary]
                WHERE ISNULL(Status, '') IN ('Matching','Matched' )
                AND (@Year IS NULL OR Otb_Year = @Year)
                AND (@Month IS NULL OR Otb_Month = @Month)
                AND (@Company IS NULL OR Company_Code = @Company)
                AND (@Category IS NULL OR Category_Code = @Category)
                AND (@Segment IS NULL OR SUBSTRING(Segment_Code, 2, CASE WHEN LEN(ISNULL(Segment_Code, '')) > 2 THEN LEN(Segment_Code) - 2 ELSE 0 END) = @Segment)
                AND (@Brand IS NULL OR Brand_Code = @Brand)
                AND (@Vendor IS NULL OR Vendor_Code = @Vendor)
                "
            Using cmd As New SqlCommand(query, conn)
                cmd.Parameters.AddWithValue("@Year", If(String.IsNullOrEmpty(year), DBNull.Value, year))
                cmd.Parameters.AddWithValue("@Month", If(String.IsNullOrEmpty(month), DBNull.Value, month))
                cmd.Parameters.AddWithValue("@Company", If(String.IsNullOrEmpty(company), DBNull.Value, company))
                cmd.Parameters.AddWithValue("@Category", If(String.IsNullOrEmpty(category), DBNull.Value, category))
                cmd.Parameters.AddWithValue("@Segment", If(String.IsNullOrEmpty(segment), DBNull.Value, segment))
                cmd.Parameters.AddWithValue("@Brand", If(String.IsNullOrEmpty(brand), DBNull.Value, brand))
                cmd.Parameters.AddWithValue("@Vendor", If(String.IsNullOrEmpty(vendor), DBNull.Value, vendor))
                Using da As New SqlDataAdapter(cmd)
                    da.Fill(dt)
                End Using
            End Using
        End Using
        Return dt
    End Function

    Private Function GetDistinctKeysForReport(year As String, month As String, company As String, category As String, segment As String, brand As String, vendor As String) As DataTable
        ' รวม Key จากทุกตารางเพื่อให้ได้รายการครบถ้วน
        Dim dt As New DataTable()
        Using conn As New SqlConnection(connectionString)
            conn.Open()
            Dim query As String = "
                SELECT DISTINCT Year, Month, Category, Company, Segment, Brand, Vendor 
                FROM [dbo].[OTB_Transaction] 
                WHERE OTBStatus='Approved' 
                AND (@Year IS NULL OR Year = @Year)
                AND (@Month IS NULL OR Month = @Month)
                AND (@Company IS NULL OR Company = @Company)
                AND (@Category IS NULL OR Category = @Category)
                AND (@Segment IS NULL OR Segment = @Segment)
                AND (@Brand IS NULL OR Brand = @Brand)
                AND (@Vendor IS NULL OR Vendor = @Vendor)

                UNION

                SELECT DISTINCT Year, Month, Category, Company, Segment, Brand, Vendor 
                FROM [dbo].[OTB_Switching_Transaction] 
                WHERE OTBStatus='Approved' 
                AND (@Year IS NULL OR Year = @Year)
                AND (@Month IS NULL OR Month = @Month)
                AND (@Company IS NULL OR Company = @Company)
                AND (@Category IS NULL OR Category = @Category)
                AND (@Segment IS NULL OR Segment = @Segment)
                AND (@Brand IS NULL OR Brand = @Brand)
                AND (@Vendor IS NULL OR Vendor = @Vendor)

                UNION

                SELECT DISTINCT SwitchYear AS Year, SwitchMonth AS Month, SwitchCategory AS Category, SwitchCompany AS Company, SwitchSegment AS Segment, SwitchBrand AS Brand, SwitchVendor AS Vendor 
                FROM [dbo].[OTB_Switching_Transaction] 
                WHERE OTBStatus='Approved' AND [To] IS NOT NULL 
                AND (@Year IS NULL OR SwitchYear = @Year)
                AND (@Month IS NULL OR SwitchMonth = @Month)
                AND (@Company IS NULL OR SwitchCompany = @Company)
                AND (@Category IS NULL OR SwitchCategory = @Category)
                AND (@Segment IS NULL OR SwitchSegment = @Segment)
                AND (@Brand IS NULL OR SwitchBrand = @Brand)
                AND (@Vendor IS NULL OR SwitchVendor = @Vendor)

                SELECT DISTINCT PO_Year, PO_Month, Category_Code, Company_Code, Segment_Code, Brand_Code, Vendor_Code 
                FROM [dbo].[Draft_PO_Transaction] 
                WHERE ISNULL(Status,'')<>'Cancelled' 
                AND (@Year IS NULL OR PO_Year = @Year)
                AND (@Month IS NULL OR PO_Month = @Month)
                AND (@Company IS NULL OR Company_Code = @Company)
                AND (@Category IS NULL OR Category_Code = @Category)
                AND (@Segment IS NULL OR Segment_Code = @Segment)
                AND (@Brand IS NULL OR Brand_Code = @Brand)
                AND (@Vendor IS NULL OR Vendor_Code = @Vendor)

                UNION

                SELECT DISTINCT Otb_Year, Otb_Month, Category, Company_Code, Fund, Brand, Supplier 
                FROM [dbo].[Actual_PO_Staging] 
                WHERE ISNULL(Deletion_Flag,'')<>'L' 
                AND (@Year IS NULL OR Otb_Year = @Year)
                AND (@Month IS NULL OR Otb_Month = @Month)
                AND (@Company IS NULL OR Company_Code = @Company)
                AND (@Category IS NULL OR Category = @Category)
                AND (@Segment IS NULL OR Fund = @Segment)
                AND (@Brand IS NULL OR Brand = @Brand)
                AND (@Vendor IS NULL OR Supplier = @Vendor)
            "
            Using cmd As New SqlCommand(query, conn)
                cmd.Parameters.AddWithValue("@Year", If(String.IsNullOrEmpty(year), DBNull.Value, year))
                cmd.Parameters.AddWithValue("@Month", If(String.IsNullOrEmpty(month), DBNull.Value, month))
                cmd.Parameters.AddWithValue("@Company", If(String.IsNullOrEmpty(company), DBNull.Value, company))
                cmd.Parameters.AddWithValue("@Category", If(String.IsNullOrEmpty(category), DBNull.Value, category))
                cmd.Parameters.AddWithValue("@Segment", If(String.IsNullOrEmpty(segment), DBNull.Value, segment))
                cmd.Parameters.AddWithValue("@Brand", If(String.IsNullOrEmpty(brand), DBNull.Value, brand))
                cmd.Parameters.AddWithValue("@Vendor", If(String.IsNullOrEmpty(vendor), DBNull.Value, vendor))
                Using da As New SqlDataAdapter(cmd)
                    da.Fill(dt)
                End Using
            End Using
        End Using
        Return dt
    End Function

    ' --- Excel Generation ---
    Private Sub GenerateExcelOTBMovement(context As HttpContext, dt As DataTable, filename As String)
        ExcelPackage.License.SetNonCommercialOrganization("KingPower")

        Using package As New ExcelPackage()
            Dim ws = package.Workbook.Worksheets.Add("OTB Movement")

            ' Load Data starting from Row 3 (เผื่อ Header 2 บรรทัดตามภาพ)
            ws.Cells("A3").LoadFromDataTable(dt, False)

            ' --- สร้าง Headers ตามภาพ Template ---
            ' Row 1: Group Headers (เช่น ค่า + เท่านั้น, ค่า - เท่านั้น)
            ' Note: Column Index เริ่มที่ 1
            ' Columns: 
            ' 1-7: Keys (Year..Vendor)
            ' 8: Budget Approved
            ' 9: Revised Diff
            ' 10: Extra
            ' 11: Switch In
            ' 12: Balance In
            ' 13: Carry In
            ' 14: Switch Out
            ' 15: Balance Out
            ' 16: Carry Out
            ' 17: Total Budget Approved
            ' 18: Actual PO
            ' 19: Draft PO
            ' 20: Total Actual+Draft
            ' 21: Remaining

            Dim headers() As String = {"Year", "Month", "Cate", "Category", "Company", "CompanyName", "Segment", "SegmentName", "Brand", "BrandName", "Vendor", "VendorName",
                                       "Budget Approved", "Revised Diff", "Extra",
                                       "Switch in", "Balance in", "Carry in",
                                       "Switch out", "Balance out", "Carry out",
                                       "Total Budget Approved", "Actual PO", "Draft PO", "Total Actual + Draft PO", "Remaining"}

            ' Set Row 2 Headers
            For i As Integer = 0 To headers.Length - 1
                ws.Cells(2, i + 1).Value = headers(i)
            Next

            ' Set Row 1 Group Headers (Label สีแดงๆ ในภาพ)
            ws.Cells(1, 13).Value = "ค่า + เท่านั้น" ' Budget Approved (จริงๆ อันนี้อาจเป็นค่าตั้งต้น)
            ws.Cells(1, 14).Value = "ค่ามีได้ทั้ง +,-" ' Revised
            ws.Cells(1, 15).Value = "ค่า + เท่านั้น" ' Extra
            ws.Cells(1, 16).Value = "ค่า + เท่านั้น" ' Switch In
            ws.Cells(1, 17).Value = "ค่า + เท่านั้น" ' Balance In
            ws.Cells(1, 18).Value = "ค่า + เท่านั้น" ' Carry In
            ws.Cells(1, 19).Value = "ค่า - เท่านั้น" ' Switch Out
            ws.Cells(1, 20).Value = "ค่า - เท่านั้น" ' Balance Out
            ws.Cells(1, 21).Value = "ค่า - เท่านั้น" ' Carry Out

            ' --- Styling ---
            ' Header Row 2 (Blue Background)
            Using rng = ws.Cells(2, 1, 2, 26)
                rng.Style.Font.Bold = True
                rng.Style.Fill.PatternType = ExcelFillStyle.Solid
                rng.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(0, 90, 160)) ' KPG Blue
                rng.Style.Font.Color.SetColor(Color.White)
                rng.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center
            End Using

            ' Total Budget Column (Green)
            Using rng = ws.Cells(2, 22, dt.Rows.Count + 2, 22)
                rng.Style.Fill.PatternType = ExcelFillStyle.Solid
                rng.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(146, 208, 80)) ' Light Green
            End Using

            ' Total Usage Column (Green)
            Using rng = ws.Cells(2, 25, dt.Rows.Count + 2, 25)
                rng.Style.Fill.PatternType = ExcelFillStyle.Solid
                rng.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(146, 208, 80))
            End Using

            ' PO Columns (Blue Header background for Actual/Draft/Remaining based on image)
            ws.Cells(2, 23, 2, 24).Style.Fill.BackgroundColor.SetColor(Color.FromArgb(0, 90, 160))

            ' Number Format
            ws.Cells(3, 13, dt.Rows.Count + 2, 26).Style.Numberformat.Format = "#,##0.00"

            ' AutoFit
            ws.Cells.AutoFitColumns()

            ' Response
            context.Response.Clear()
            context.Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            context.Response.AddHeader("content-disposition", $"attachment; filename={filename}")
            context.Response.BinaryWrite(package.GetAsByteArray())
            context.Response.Flush()
            context.ApplicationInstance.CompleteRequest()
        End Using
    End Sub

    ' ==================================================================
    ' ===== NEW: SUMMARY OTB BY CATEGORY REPORT ========================
    ' ==================================================================
    Private Sub HandleExportSummaryCategory(context As HttpContext)
        ' 1. รับ Parameters
        Dim year As String = context.Request.QueryString("OTByear")
        Dim month As String = context.Request.QueryString("OTBmonth")
        Dim company As String = context.Request.QueryString("OTBCompany")
        Dim category As String = context.Request.QueryString("OTBCategory")
        Dim segment As String = context.Request.QueryString("OTBSegment")
        Dim masterinstance As New MasterDataUtil
        ' *สำคัญ* เราต้องดึงข้อมูลระดับ Brand/Vendor ทั้งหมดภายใต้ Filter นี้มาคำนวณก่อน แล้วค่อย Group รวม
        Dim brand As String = context.Request.QueryString("OTBBrand")
        Dim vendor As String = context.Request.QueryString("OTBVendor")


        If String.IsNullOrEmpty(year) Then Throw New Exception("Year is required.")

        ' 2. เตรียม Calculator และโหลดข้อมูลดิบ
        Dim budgetCalc As New OTBBudgetCalculator()
        Dim dtDraftPO As DataTable = LoadDraftPOForReport(year, month, company, category, segment, brand, vendor)
        Dim dtActualPO As DataTable = LoadActualPOForReport(year, month, company, category, segment, brand, vendor)

        ' ดึง Key ทั้งหมดที่มีข้อมูล (ระดับละเอียด)
        Dim dtKeys As DataTable = GetDistinctKeysForReport(year, month, company, category, segment, brand, vendor)

        ' 3. สร้าง List ชั่วคราวเพื่อเก็บข้อมูลระดับละเอียดก่อน Group
        Dim rawList As New List(Of SummaryRawItem)

        For Each rowKey As DataRow In dtKeys.Rows
            Dim kYear As String = rowKey("Year").ToString()
            Dim kMonth As String = rowKey("Month").ToString()
            Dim kCate As String = rowKey("Category").ToString()
            Dim kCateName As String = masterinstance.GetCategoryName(rowKey("Category").ToString())
            Dim kComp As String = rowKey("Company").ToString()
            Dim kCompName As String = masterinstance.GetCompanyName(rowKey("Company").ToString())
            Dim kSeg As String = rowKey("Segment").ToString()
            Dim kSegName As String = masterinstance.GetSegmentName(rowKey("Segment").ToString())
            Dim kBrand As String = rowKey("Brand").ToString()
            Dim kBrandName As String = masterinstance.GetBrandName(rowKey("Brand").ToString())
            Dim kVendor As String = rowKey("Vendor").ToString()
            Dim kVendorName As String = masterinstance.GetVendorName(rowKey("Vendor").ToString())
            Dim kMonthName As String = GetMonthName(kMonth)

            ' 3.1 คำนวณ Budget (ระดับละเอียด)
            Dim budgetBreakdown = budgetCalc.GetBudgetBreakdown(kYear, kMonth, kCate, kComp, kSeg, kBrand, kVendor)
            Dim totalBudgetApproved As Decimal = budgetBreakdown("Total")

            ' 3.2 คำนวณ PO (ระดับละเอียด)
            Dim filterPO As String = $"Year = '{kYear}' AND Month = '{kMonth}' AND Category = '{kCate}' AND Company = '{kComp}' AND Segment = '{kSeg}' AND Brand = '{kBrand}' AND Vendor = '{kVendor}'"
            Dim sumDraft As Decimal = If(IsDBNull(dtDraftPO.Compute("SUM(Amount)", filterPO)), 0, Convert.ToDecimal(dtDraftPO.Compute("SUM(Amount)", filterPO)))
            Dim sumActual As Decimal = If(IsDBNull(dtActualPO.Compute("SUM(Amount)", filterPO)), 0, Convert.ToDecimal(dtActualPO.Compute("SUM(Amount)", filterPO)))

            ' เก็บลง List
            rawList.Add(New SummaryRawItem With {
                .Year = kYear,
                .Month = kMonth,
                .MonthName = kMonthName,
                .Category = kCate,
                .CategoryName = kCateName,
                .Company = kComp,
                .CompanyName = kCompName,
                .Segment = kSeg,
                .SegmentName = kSegName,
                .TotalBudget = totalBudgetApproved,
                .TotalActualDraft = sumDraft + sumActual
            })
        Next

        ' 4. (สำคัญ) Group By Category & Segment ด้วย LINQ
        ' *** FIXED: Group.Sum จะใช้ Decimal overload อัตโนมัติเพราะ rawList เป็น Strongly Typed ***
        Dim groupedQuery = From item In rawList
                           Group item By Key = New With {
                               Key .Company = item.Company,
                               Key .Year = item.Year,
                               Key .MonthName = item.MonthName,
                               Key .Category = item.Category,
                               Key .Segment = item.Segment
                           } Into Group
                           Select New With {
                               .Company = Key.Company,
                               .Year = Key.Year,
                               .Month = Key.MonthName,
                               .Category = Key.Category,
                               .Segment = Key.Segment,
                               .TotalBudgetApproved = Group.Sum(Function(x) x.TotalBudget),
                               .TotalActualDraft = Group.Sum(Function(x) x.TotalActualDraft),
                               .Remaining = Group.Sum(Function(x) x.TotalBudget) - Group.Sum(Function(x) x.TotalActualDraft)
                           }

        ' 5. สร้าง DataTable สำหรับ Export
        Dim dtExport As New DataTable("SummaryOTB")
        dtExport.Columns.Add("Company")
        dtExport.Columns.Add("CompanyName")
        dtExport.Columns.Add("Year")
        dtExport.Columns.Add("Month")
        dtExport.Columns.Add("Cate")
        dtExport.Columns.Add("Category")
        dtExport.Columns.Add("Segment")
        dtExport.Columns.Add("SegmentName")
        dtExport.Columns.Add("Total Budget Approved", GetType(Decimal))
        dtExport.Columns.Add("Total Actual + Draft PO", GetType(Decimal))
        dtExport.Columns.Add("Remaining", GetType(Decimal))

        For Each row In groupedQuery
            ' Filter out rows with zero values everywhere if desired, or keep all
            If row.TotalBudgetApproved <> 0 OrElse row.TotalActualDraft <> 0 Then
                dtExport.Rows.Add(
                    row.Company,
                    masterinstance.GetCompanyName(row.Company),
                    row.Year,
                    row.Month,
                    row.Category,
                    masterinstance.GetCategoryName(row.Category),
                    row.Segment,
                    masterinstance.GetSegmentName(row.Segment),
                    row.TotalBudgetApproved,
                    row.TotalActualDraft,
                    row.Remaining
                )
            End If
        Next

        ' 6. Generate Excel
        GenerateExcelSummaryCategory(context, dtExport, $"Summary_OTB_Category_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx")
    End Sub

    Private Sub GenerateExcelSummaryCategory(context As HttpContext, dt As DataTable, filename As String)
        ExcelPackage.License.SetNonCommercialOrganization("KingPower")

        Using package As New ExcelPackage()
            Dim ws = package.Workbook.Worksheets.Add("Summary OTB")

            ' ใส่ Header แบบ Custom ตามภาพ
            ws.Cells("A1").Value = "Summary OTB by category"
            ws.Cells("A1").Style.Font.Bold = True
            ws.Cells("A1").Style.Font.Size = 14

            ' Load Data starting from Row 3
            ws.Cells("A3").LoadFromDataTable(dt, True)

            ' Style Header Row (Row 3)
            Using rng = ws.Cells(3, 1, 3, dt.Columns.Count)
                rng.Style.Font.Bold = True
                rng.Style.Fill.PatternType = ExcelFillStyle.Solid
                rng.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(0, 90, 160)) ' Blue
                rng.Style.Font.Color.SetColor(Color.White)
                rng.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center
            End Using

            ' Style Specific Headers (Budget, Actual, Remaining) - Green Header based on image
            Using rng = ws.Cells(3, 9, 3, 10) ' Total Budget, Total Actual
                rng.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(146, 208, 80)) ' Light Green
                rng.Style.Font.Color.SetColor(Color.White)
            End Using
            Using rng = ws.Cells(3, 11, 3, 11) ' Remaining
                rng.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(0, 90, 160)) ' Blue
                rng.Style.Font.Color.SetColor(Color.White)
            End Using

            ' Format Numbers
            Using rng = ws.Cells(4, 9, dt.Rows.Count + 3, 11)
                rng.Style.Numberformat.Format = "#,##0.00"
            End Using

            ws.Cells.AutoFitColumns()

            context.Response.Clear()
            context.Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            context.Response.AddHeader("content-disposition", $"attachment; filename={filename}")
            context.Response.BinaryWrite(package.GetAsByteArray())
            context.Response.Flush()
            context.ApplicationInstance.CompleteRequest()
        End Using
    End Sub

    ReadOnly Property IsReusable() As Boolean Implements IHttpHandler.IsReusable
        Get
            Return False
        End Get
    End Property

End Class