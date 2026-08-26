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
Imports System.Collections.Generic
Imports ExcelDataReader
Imports Newtonsoft.Json
Imports OfficeOpenXml
Imports OfficeOpenXml.Style
Imports System.Threading.Tasks
Imports System.Web.Hosting
Imports System.Web.SessionState
Imports System.Linq
Imports System.Security.Cryptography
Imports System.Text.RegularExpressions

Public Class DataOTBHandler
    Implements IHttpHandler, IReadOnlySessionState

    Private Shared connectionString As String = ConfigurationManager.ConnectionStrings("BMSConnectionString")?.ConnectionString
    Private Const MaxRunNoQueryBatchSize As Integer = 900
    ' SAP can wait for 60 minutes and the following database stage allows up
    ' to roughly 60 minutes of bounded commands. Keep stale recovery beyond
    ' either live stage so a status poll cannot terminate an active worker.
    Private Const PostSapStaleMinutes As Integer = 120
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
        Public Property TotalActualPO As Decimal
        Public Property TotalDraftPO As Decimal

        Public ReadOnly Property TotalActualDraft As Decimal
            Get
                Return TotalActualPO + TotalDraftPO
            End Get
        End Property
    End Class

    Private Class ReportBudgetRow
        Public Property Year As String
        Public Property Month As String
        Public Property Category As String
        Public Property Company As String
        Public Property Segment As String
        Public Property Brand As String
        Public Property Vendor As String
        Public Property Original As Decimal
        Public Property RevDiff As Decimal
        Public Property Extra As Decimal
        Public Property SwitchIn As Decimal
        Public Property BalanceIn As Decimal
        Public Property CarryIn As Decimal
        Public Property SwitchOut As Decimal
        Public Property BalanceOut As Decimal
        Public Property CarryOut As Decimal

        Public ReadOnly Property SignedSwitchOut As Decimal
            Get
                Return -Math.Abs(SwitchOut)
            End Get
        End Property

        Public ReadOnly Property SignedBalanceOut As Decimal
            Get
                Return -Math.Abs(BalanceOut)
            End Get
        End Property

        Public ReadOnly Property SignedCarryOut As Decimal
            Get
                Return -Math.Abs(CarryOut)
            End Get
        End Property

        Public ReadOnly Property Total As Decimal
            Get
                Return Original + RevDiff + Extra + SwitchIn + BalanceIn + CarryIn +
                    SignedSwitchOut + SignedBalanceOut + SignedCarryOut
            End Get
        End Property
    End Class

    Private Class ReportUsageRow
        Public Property Year As String
        Public Property Month As String
        Public Property Category As String
        Public Property Company As String
        Public Property Segment As String
        Public Property Brand As String
        Public Property Vendor As String
        Public Property DraftPO As Decimal
        Public Property ActualPO As Decimal

        Public ReadOnly Property TotalUsage As Decimal
            Get
                Return DraftPO + ActualPO
            End Get
        End Property
    End Class

    Private Class ApprovalPreparedRow
        Public Property RunNo As Integer
        Public Property Data As DataRow
        Public Property SapItem As OtbPlanUploadItem
        Public Property SapKey As String
        Public Property BusinessKey As String
        Public Property CurrentApproved As Decimal
    End Class

    Private Class ApprovalMappedResult
        Public Property RunNo As Integer
        Public Property SapResult As SapUploadResultItem
    End Class

    Private Class ApprovalGroupAccumulator
        Public Property Company As String
        Public Property Year As String
        Public Property Month As String
        Public Property Category As String
        Public Property Revised As Decimal
        Public Property SelectedCurrent As Decimal
    End Class

    Private Class ApprovalPreviewSnapshot
        Public Property Groups As List(Of DraftOtbApprovalPreviewRow)
        Public Property PreviewHash As String
        Public Property CurrentApprovedByBusinessKey As Dictionary(Of String, Decimal)
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
            ElseIf action = "approvalpreview" Then
                GetApprovalPreview(context)
            ElseIf action = "approvaljobstatus" Then
                GetApprovalJobStatus(context)
            ElseIf action = "latestapprovaljob" Then
                GetLatestApprovalJob(context)
            ElseIf action = "acknowledgeapprovaljob" Then
                AcknowledgeApprovalJob(context)
            ElseIf action = "approvedraftotb" Then
                StartApprovalJob(context)
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
                ElseIf action = "deletedraftotb" Then
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
            context.Response.StatusCode = 200
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
        Dim hasStoredProcDiff As Boolean = dtRaw.Columns.Contains("CurrentApproved") AndAlso dtRaw.Columns.Contains("Diff")
        Dim budgetCalculator As OTBBudgetCalculator = Nothing
        If Not hasStoredProcDiff Then
            budgetCalculator = New OTBBudgetCalculator()
        End If
        Dim budgetCache As New Dictionary(Of String, Decimal)(StringComparer.OrdinalIgnoreCase)
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

            Dim amountValue As Decimal = GetDecimalValue(row, If(dtRaw.Columns.Contains("ToBeAmountTHB"), "ToBeAmountTHB", "Amount"))

            Dim budgetKey As String = BuildOTBBudgetKey(OTBYear_Calc, OTBMonth_Calc, OTBCategory_Calc, OTBCompany_Calc, OTBSegment_Calc, OTBBrand_Calc, OTBVendor_Calc)
            Dim currentBudget As Decimal
            Dim diffAmount As Decimal
            If hasStoredProcDiff Then
                currentBudget = GetDecimalValue(row, "CurrentApproved")
                diffAmount = GetDecimalValue(row, "Diff")
            Else
                If Not budgetCache.TryGetValue(budgetKey, currentBudget) Then
                    currentBudget = budgetCalculator.CalculateCurrentApprovedBudget(OTBYear_Calc, OTBMonth_Calc, OTBCategory_Calc, OTBCompany_Calc, OTBSegment_Calc, OTBBrand_Calc, OTBVendor_Calc)
                    budgetCache(budgetKey) = currentBudget
                End If
                diffAmount = amountValue - currentBudget
            End If
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
                GetStringValue(row, If(dtRaw.Columns.Contains("Vendor"), "Vendor", "VendorName")),
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
            MarkDownloadReady(context)
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
            MarkDownloadReady(context)
            context.Response.BinaryWrite(package.GetAsByteArray())
            context.Response.Flush()
            context.ApplicationInstance.CompleteRequest()

        End Using

    End Sub
    ' --- (END) NEW FUNCTION FOR SUMMARY EXPORT ---


    Private Function GenerateHtmlDraftTable(dt As DataTable) As String
        Dim sb As New StringBuilder()
        Dim hasStoredProcDiff As Boolean = dt.Columns.Contains("CurrentApproved") AndAlso dt.Columns.Contains("Diff")
        Dim budgetCalculator As OTBBudgetCalculator = Nothing
        If Not hasStoredProcDiff Then
            budgetCalculator = New OTBBudgetCalculator()
        End If
        Dim budgetCache As New Dictionary(Of String, Decimal)(StringComparer.OrdinalIgnoreCase)
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
                Dim Vendor As String = GetStringValue(dt.Rows(i), If(dt.Columns.Contains("Vendor"), "Vendor", "VendorName"))
                Dim Amount As String = "0.00"
                Dim CurrentBudgetAmount As String = "0.00"
                Dim Diff As String = "0.00"
                Dim amountValue As Decimal = GetDecimalValue(dt.Rows(i), If(dt.Columns.Contains("ToBeAmountTHB"), "ToBeAmountTHB", "Amount"))
                Amount = amountValue.ToString("N2")
                Dim budgetKey As String = BuildOTBBudgetKey(OTBYear, OTBMonth, OTBCategory, OTBCompany, OTBSegment, OTBBrand, OTBVendor)
                Dim currentBudget As Decimal = 0
                Dim diffamout As Decimal
                If hasStoredProcDiff Then
                    currentBudget = GetDecimalValue(dt.Rows(i), "CurrentApproved")
                    diffamout = GetDecimalValue(dt.Rows(i), "Diff")
                Else
                    If Not budgetCache.TryGetValue(budgetKey, currentBudget) Then
                        currentBudget = budgetCalculator.CalculateCurrentApprovedBudget(OTBYear, OTBMonth, OTBCategory, OTBCompany, OTBSegment, OTBBrand, OTBVendor)
                        budgetCache(budgetKey) = currentBudget
                    End If
                    diffamout = amountValue - currentBudget
                End If
                CurrentBudgetAmount = currentBudget.ToString("N2")

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

    Private Function BuildOTBBudgetKey(year As String, month As String, category As String, company As String, segment As String, brand As String, vendor As String) As String
        Return String.Join("|", New String() {year, month, category, company, segment, brand, vendor})
    End Function

    Private Function GetStringValue(row As DataRow, columnName As String) As String
        If row Is Nothing OrElse row.Table Is Nothing OrElse Not row.Table.Columns.Contains(columnName) OrElse row(columnName) Is DBNull.Value Then
            Return ""
        End If

        Return row(columnName).ToString()
    End Function

    Private Function GetDecimalValue(row As DataRow, columnName As String) As Decimal
        Dim value As Decimal = 0D
        If row Is Nothing OrElse row.Table Is Nothing OrElse Not row.Table.Columns.Contains(columnName) OrElse row(columnName) Is DBNull.Value Then
            Return value
        End If

        Decimal.TryParse(row(columnName).ToString(), value)
        Return value
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
            sb.Append("<tr><td colspan='27' class='text-center text-muted'>No switch OTB records found</td></tr>")
        Else
            For Each row As DataRow In dt.Rows
                sb.Append("<tr>")

                ' Create Date
                sb.AppendFormat("<td class='date-cell'>{0}</td>",
                       If(row("CreateDT") IsNot DBNull.Value, Convert.ToDateTime(row("CreateDT")).ToString("dd/MM/yyyy HH:mm:ss"), ""))

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

    Private Function GetOTBDraftDataWithFilter(OTBtype As String, OTByear As String, OTBmonth As String, OTBCompany As String, OTBCategory As String, OTBSegment As String, OTBBrand As String, OTBVendor As String, Optional OTBVersion As String = "") As DataTable
        Dim dt As New DataTable()
        Using conn As New SqlConnection(connectionString)
            conn.Open()
            Using cmd As New SqlCommand("dbo.SP_Get_Draft_OTB_Diff", conn)
                cmd.CommandType = CommandType.StoredProcedure
                AddOptionalNVarCharParameter(cmd, "@Version", OTBVersion, 20)
                AddOptionalNVarCharParameter(cmd, "@Type", OTBtype, 20)
                AddOptionalNVarCharParameter(cmd, "@Year", OTByear, 10)
                AddOptionalNVarCharParameter(cmd, "@Month", OTBmonth, 10)
                AddOptionalNVarCharParameter(cmd, "@Company", OTBCompany, 20)
                AddOptionalNVarCharParameter(cmd, "@Category", OTBCategory, 20)
                AddOptionalNVarCharParameter(cmd, "@Segment", OTBSegment, 20)
                AddOptionalNVarCharParameter(cmd, "@Brand", OTBBrand, 30)
                AddOptionalNVarCharParameter(cmd, "@Vendor", OTBVendor, 30)

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

    Private Sub GetApprovalPreview(context As HttpContext)
        context.Response.ContentType = "application/json"
        Try
            EnsureApprovalPermission(context, True)
            Dim runNos As List(Of Integer) = ParseApprovalRunNos(context)
            Dim snapshot As ApprovalPreviewSnapshot = BuildApprovalPreview(runNos)
            context.Response.Write(JsonConvert.SerializeObject(New With {
                .success = True,
                .totalRows = runNos.Count,
                .groups = snapshot.Groups,
                .previewHash = snapshot.PreviewHash
            }))
        Catch ex As Exception
            context.Response.StatusCode = 200
            context.Response.Write(JsonConvert.SerializeObject(New With {
                .success = False,
                .message = ex.Message
            }))
        End Try
    End Sub

    Private Sub StartApprovalJob(context As HttpContext)
        context.Response.ContentType = "application/json"
        Try
            EnsureAjaxMutationRequest(context)
            Dim approvedBy As String = EnsureApprovalPermission(context, True)
            Dim runNos As List(Of Integer) = ParseApprovalRunNos(context)
            Dim previewHash As String = If(context.Request.Form("previewHash"), "").Trim().ToLowerInvariant()
            If Not Regex.IsMatch(previewHash, "\A[0-9a-f]{64}\z", RegexOptions.CultureInvariant) Then
                Throw New Exception("A valid approval preview hash is required. Please preview the selected rows again.")
            End If
            Dim clientRequestId As String = If(context.Request.Form("clientRequestId"), "").Trim()
            If String.IsNullOrWhiteSpace(clientRequestId) Then
                Throw New Exception("A client request ID is required. Please refresh the page and preview again.")
            End If
            If clientRequestId.Length > 100 Then
                Throw New Exception("The client request ID is too long. Please refresh the page and preview again.")
            End If
            Dim payload As New DraftOtbApprovalJobPayload With {
                .RunNos = runNos,
                .ApprovedBy = approvedBy,
                .PreviewHash = previewHash
            }
            Dim job As DraftOtbJobRecord = DraftOtbJobStore.CreateOrGet(
                DraftOtbJobTypes.Approval,
                approvedBy,
                clientRequestId,
                JsonConvert.SerializeObject(payload),
                runNos.Count)

            If String.Equals(job.Status, DraftOtbJobStatuses.Queued, StringComparison.OrdinalIgnoreCase) Then
                QueueApprovalJob(job.JobId)
            End If
            WriteApprovalJobJson(context, DraftOtbJobStore.GetJobMetadata(job.JobId, approvedBy))
        Catch ex As Exception
            context.Response.StatusCode = 200
            context.Response.Write(JsonConvert.SerializeObject(New With {
                .success = False,
                .message = ex.Message
            }))
        End Try
    End Sub

    Private Shared Sub EnsureAjaxMutationRequest(context As HttpContext)
        If context Is Nothing OrElse Not String.Equals(context.Request.HttpMethod, "POST", StringComparison.OrdinalIgnoreCase) Then
            Throw New UnauthorizedAccessException("This action must be submitted with a same-origin POST request.")
        End If
        If Not String.Equals(context.Request.Headers("X-Requested-With"), "XMLHttpRequest", StringComparison.OrdinalIgnoreCase) Then
            Throw New UnauthorizedAccessException("The request could not be verified. Refresh the page and try again.")
        End If
    End Sub

    Private Sub GetApprovalJobStatus(context As HttpContext)
        context.Response.ContentType = "application/json"
        SetJobResponseNoStore(context)
        Try
            Dim currentUser As String = EnsureApprovalPermission(context, False)
            Dim jobId As Guid
            If Not Guid.TryParse(If(context.Request("jobId"), ""), jobId) Then Throw New Exception("A valid approval Job ID is required.")
            Dim job As DraftOtbJobRecord = DraftOtbJobStore.GetJobMetadata(jobId, currentUser)
            If job Is Nothing Then Throw New Exception("Approval job was not found.")
            If IsStalePostSapApproval(job) AndAlso
               DraftOtbJobStore.TryMarkStaleApprovalReconciliation(job.JobId, currentUser, DateTime.UtcNow.AddMinutes(-PostSapStaleMinutes)) Then
                job = DraftOtbJobStore.GetJobMetadata(jobId, currentUser)
            End If
            If DraftOtbJobStatuses.IsTerminal(job.Status) Then
                job = DraftOtbJobStore.GetJob(jobId, currentUser)
            End If
            WriteApprovalJobJson(context, job)
        Catch ex As Exception
            context.Response.StatusCode = 200
            context.Response.Write(JsonConvert.SerializeObject(New With {
                .success = False,
                .message = ex.Message,
                .errorCode = GetApprovalJobErrorCode(ex),
                .retryable = IsRetryableApprovalJobError(ex)
            }))
        End Try
    End Sub

    Private Sub GetLatestApprovalJob(context As HttpContext)
        context.Response.ContentType = "application/json"
        SetJobResponseNoStore(context)
        Try
            Dim currentUser As String = EnsureApprovalPermission(context, False)
            Dim clientRequestId As String = If(context.Request("clientRequestId"), "").Trim()
            If clientRequestId.Length > 100 Then Throw New Exception("The client request ID is too long.")
            Dim job As DraftOtbJobRecord = Nothing
            If Not String.IsNullOrWhiteSpace(clientRequestId) Then
                job = DraftOtbJobStore.GetByClientRequestMetadata(DraftOtbJobTypes.Approval, currentUser, clientRequestId)
            Else
                job = DraftOtbJobStore.GetLatestActiveMetadata(DraftOtbJobTypes.Approval, currentUser)
            End If
            If job Is Nothing Then
                context.Response.Write(JsonConvert.SerializeObject(New With {.success = True, .job = CType(Nothing, Object)}))
                Return
            End If
            If IsStalePostSapApproval(job) AndAlso
               DraftOtbJobStore.TryMarkStaleApprovalReconciliation(job.JobId, currentUser, DateTime.UtcNow.AddMinutes(-PostSapStaleMinutes)) Then
                job = DraftOtbJobStore.GetJobMetadata(job.JobId, currentUser)
            End If
            If DraftOtbJobStatuses.IsTerminal(job.Status) Then
                job = DraftOtbJobStore.GetJob(job.JobId, currentUser)
            End If
            WriteApprovalJobJson(context, job)
        Catch ex As Exception
            context.Response.StatusCode = 200
            context.Response.Write(JsonConvert.SerializeObject(New With {
                .success = False,
                .message = ex.Message,
                .errorCode = GetApprovalJobErrorCode(ex),
                .retryable = IsRetryableApprovalJobError(ex)
            }))
        End Try
    End Sub

    Private Sub AcknowledgeApprovalJob(context As HttpContext)
        context.Response.ContentType = "application/json"
        SetJobResponseNoStore(context)
        Try
            EnsureAjaxMutationRequest(context)
            Dim currentUser As String = EnsureApprovalPermission(context, False)
            Dim jobId As Guid
            If Not Guid.TryParse(If(context.Request.Form("jobId"), "").Trim(), jobId) Then
                Throw New Exception("A valid approval Job ID is required.")
            End If
            If Not DraftOtbJobStore.Acknowledge(jobId, currentUser, DraftOtbJobTypes.Approval) Then
                Throw New Exception("The approval job was not found or is not ready to acknowledge.")
            End If
            context.Response.Write(JsonConvert.SerializeObject(New With {
                .success = True,
                .jobId = jobId.ToString("D"),
                .acknowledged = True
            }))
        Catch ex As Exception
            context.Response.StatusCode = 200
            context.Response.Write(JsonConvert.SerializeObject(New With {.success = False, .message = ex.Message}))
        End Try
    End Sub

    Private Shared Sub SetJobResponseNoStore(context As HttpContext)
        context.Response.Cache.SetCacheability(HttpCacheability.NoCache)
        context.Response.Cache.SetNoStore()
        context.Response.Cache.SetExpires(DateTime.UtcNow.AddYears(-1))
        context.Response.AppendHeader("Pragma", "no-cache")
    End Sub

    Private Shared Function GetApprovalJobErrorCode(ex As Exception) As String
        If TypeOf ex Is UnauthorizedAccessException Then Return "AUTH"
        If ex IsNot Nothing AndAlso ex.Message.IndexOf("not found", StringComparison.OrdinalIgnoreCase) >= 0 Then Return "NOT_FOUND"
        Return "TRANSIENT"
    End Function

    Private Shared Function IsRetryableApprovalJobError(ex As Exception) As Boolean
        Dim code As String = GetApprovalJobErrorCode(ex)
        Return code <> "AUTH" AndAlso code <> "NOT_FOUND"
    End Function

    Private Shared Function IsStalePostSapApproval(job As DraftOtbJobRecord) As Boolean
        If job Is Nothing OrElse job.UpdatedAt >= DateTime.UtcNow.AddMinutes(-PostSapStaleMinutes) Then Return False
        Return String.Equals(job.Status, DraftOtbJobStatuses.SendingToSap, StringComparison.OrdinalIgnoreCase) OrElse
               String.Equals(job.Status, DraftOtbJobStatuses.SavingApproval, StringComparison.OrdinalIgnoreCase)
    End Function

    Private Sub WriteApprovalJobJson(context As HttpContext, job As DraftOtbJobRecord)
        If job Is Nothing Then Throw New Exception("Approval job was not found.")
        Dim detailedResults As Object = Nothing
        If DraftOtbJobStatuses.IsTerminal(job.Status) AndAlso Not String.IsNullOrWhiteSpace(job.ResultJson) Then
            detailedResults = JsonConvert.DeserializeObject(job.ResultJson)
        End If
        context.Response.Write(JsonConvert.SerializeObject(New With {
            .success = True,
            .jobId = job.JobId.ToString("D"),
            .status = job.Status,
            .stage = job.Stage,
            .progress = job.ProgressPercent,
            .totalRows = job.TotalRows,
            .processedRows = job.ProcessedRows,
            .successRows = job.SuccessRows,
            .errorRows = job.ErrorRows,
            .message = job.Message,
            .detailedResults = detailedResults,
            .acknowledged = job.AcknowledgedAt.HasValue,
            .acknowledgedAt = job.AcknowledgedAt,
            .updatedAt = job.UpdatedAt
        }))
    End Sub

    Private Function ParseApprovalRunNos(context As HttpContext) As List(Of Integer)
        Dim idsString As String = If(context.Request.Form("runNos"), context.Request("runNos"))
        If String.IsNullOrWhiteSpace(idsString) Then Throw New Exception("No records selected for approval.")
        Dim runNos As List(Of Integer) = ParseRunNos(idsString)
        If runNos.Count = 0 Then Throw New Exception("No valid RunNos were provided.")
        If runNos.Count > 15000 Then Throw New Exception($"A maximum of 15,000 rows can be approved in one job. Selected: {runNos.Count:N0}.")
        Return runNos
    End Function

    Private Function GetApprovalRequestUser(context As HttpContext) As String
        Dim value As String = Nothing
        If context IsNot Nothing AndAlso context.Session IsNot Nothing Then
            value = Convert.ToString(context.Session("user"))
        End If
        If String.IsNullOrWhiteSpace(value) Then
            Throw New UnauthorizedAccessException("Your login session has expired. Please sign in again.")
        End If
        value = value.Trim()
        If value.Length > 100 Then Throw New UnauthorizedAccessException("The signed-in user ID is invalid.")
        Return value
    End Function

    Private Function EnsureApprovalPermission(context As HttpContext, requireApprove As Boolean) As String
        Dim username As String = GetApprovalRequestUser(context)
        Dim roleName As Object = If(context.Session Is Nothing, Nothing, context.Session("UserRole"))
        Dim rights As PermissionHelper.UserRights = PermissionHelper.GetPermission(username, "draftOTB.aspx", roleName)
        If Not rights.CanView OrElse (requireApprove AndAlso Not rights.CanApprove) Then
            Throw New UnauthorizedAccessException(If(requireApprove,
                "You do not have permission to approve Draft OTB.",
                "You do not have permission to view Draft OTB approval jobs."))
        End If
        Return username
    End Function

    Private Shared Sub QueueApprovalJob(jobId As Guid)
        HostingEnvironment.QueueBackgroundWorkItem(
            Sub(cancellationToken)
                Try
                    Dim worker As New DataOTBHandler()
                    worker.ProcessApprovalJob(jobId)
                Catch ex As Exception
                    Dim current As DraftOtbJobRecord = Nothing
                    Try
                        current = DraftOtbJobStore.GetJob(jobId)
                    Catch
                        ' A later status poll converts stale post-SAP stages to
                        ' ReconciliationRequired once the database is available.
                        Return
                    End Try
                    If current IsNot Nothing AndAlso DraftOtbJobStatuses.IsTerminal(current.Status) Then
                        ' Never revive or release claims for a terminal state. In
                        ' particular, ReconciliationRequired intentionally keeps
                        ' its durable claims until an operator resolves the SAP outcome.
                        Return
                    End If
                    Dim uncertainSapOutcome As Boolean = current IsNot Nothing AndAlso
                        (String.Equals(current.Status, DraftOtbJobStatuses.SendingToSap, StringComparison.OrdinalIgnoreCase) OrElse
                         String.Equals(current.Status, DraftOtbJobStatuses.SavingApproval, StringComparison.OrdinalIgnoreCase))
                    If Not uncertainSapOutcome Then
                        Try
                            DraftOtbJobStore.ReleaseApprovalClaims(jobId)
                        Catch
                            ' Preserve the original failure. There should be no SAP side-effect in pre-SAP stages.
                        End Try
                    End If
                    Try
                        DraftOtbJobStore.Fail(jobId, "Approval failed: " & ex.Message, uncertainSapOutcome)
                    Catch
                        ' Preserve the original state/evidence. Stale post-SAP
                        ' statuses are recovered by the status endpoint.
                    End Try
                End Try
            End Sub)
    End Sub

    Private Function BuildApprovalPreview(runNos As List(Of Integer)) As ApprovalPreviewSnapshot
        Dim draftData As DataTable = GetOTBDraftDataByRunNos(runNos)
        EnsureEveryDraftWasLoaded(runNos, draftData)

        Dim calculator As New OTBBudgetCalculator()
        Dim groups As New Dictionary(Of String, ApprovalGroupAccumulator)(StringComparer.OrdinalIgnoreCase)
        Dim selectedBudgetKeys As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
        Dim currentByDetail As New Dictionary(Of String, Decimal)(StringComparer.OrdinalIgnoreCase)
        Dim currentByBusinessKey As New Dictionary(Of String, Decimal)(StringComparer.OrdinalIgnoreCase)

        For Each row As DataRow In draftData.AsEnumerable().OrderBy(Function(item) Convert.ToInt32(item("RunNo")))
            Dim company As String = GetRequiredDraftText(row, "OTBCompany", "Company")
            Dim year As String = GetRequiredDraftText(row, "OTBYear", "Year")
            Dim month As String = GetRequiredDraftText(row, "OTBMonth", "Month")
            Dim category As String = GetRequiredDraftText(row, "OTBCategory", "Category")
            Dim groupKey As String = BuildApprovalGroupKey(company, year, month, category)
            Dim group As ApprovalGroupAccumulator = Nothing
            If Not groups.TryGetValue(groupKey, group) Then
                group = New ApprovalGroupAccumulator With {
                    .Company = company, .Year = year, .Month = month, .Category = category
                }
                groups(groupKey) = group
            End If

            Dim revised As Decimal
            If Not Decimal.TryParse(GetRequiredDraftText(row, "Amount", "Amount"), revised) Then
                Throw New Exception($"RunNo {row("RunNo")}: Amount is not a valid number.")
            End If
            group.Revised += revised

            Dim segment As String = GetRequiredDraftText(row, "OTBSegment", "Segment")
            Dim brand As String = GetRequiredDraftText(row, "OTBBrand", "Brand")
            Dim vendor As String = GetRequiredDraftText(row, "OTBVendor", "Vendor")
            Dim detailKey As String = String.Join("|", New String() {year, month, category, company, segment, brand, vendor})
            Dim currentApproved As Decimal
            If Not currentByDetail.TryGetValue(detailKey, currentApproved) Then
                currentApproved = calculator.CalculateCurrentApprovedBudget(year, month, category, company, segment, brand, vendor)
                currentByDetail(detailKey) = currentApproved
            End If
            currentByBusinessKey(BuildApprovalBusinessKey(year, month, company, category, segment, brand, vendor)) = currentApproved
            If selectedBudgetKeys.Add(detailKey) Then
                group.SelectedCurrent += currentApproved
            End If
        Next

        Dim groupTotals As Dictionary(Of String, Decimal) = GetCurrentApprovalGroupTotals(groups.Values.ToList())
        Dim result As New List(Of DraftOtbApprovalPreviewRow)()
        Dim canonical As New StringBuilder()
        canonical.Append("approval-preview-v1|")
        For Each row As DataRow In draftData.AsEnumerable().OrderBy(Function(item) Convert.ToInt32(item("RunNo")))
            Dim detailKey As String = String.Join("|", New String() {
                GetDraftText(row, "OTBYear"), GetDraftText(row, "OTBMonth"), GetDraftText(row, "OTBCategory"),
                GetDraftText(row, "OTBCompany"), GetDraftText(row, "OTBSegment"), GetDraftText(row, "OTBBrand"), GetDraftText(row, "OTBVendor")})
            Dim currentApproved As Decimal = If(currentByDetail.ContainsKey(detailKey), currentByDetail(detailKey), 0D)
            AppendApprovalHashFields(canonical, New String() {
                Convert.ToInt32(row("RunNo")).ToString(CultureInfo.InvariantCulture),
                GetDraftText(row, "Version"), GetDraftText(row, "OTBType"),
                GetDraftText(row, "OTBYear"), GetDraftText(row, "OTBMonth"),
                GetDraftText(row, "OTBCompany"), GetDraftText(row, "OTBCategory"),
                GetDraftText(row, "OTBSegment"), GetDraftText(row, "OTBBrand"), GetDraftText(row, "OTBVendor"),
                CanonicalApprovalDecimal(GetDraftText(row, "Amount")), GetDraftText(row, "OTBStatus"),
                GetDraftText(row, "Remark"), GetDraftText(row, "CateName"), GetDraftText(row, "CompanyName"),
                GetDraftText(row, "SegmentName"), GetDraftText(row, "BrandName"), GetDraftText(row, "Vendor"),
                currentApproved.ToString("0.00############################", CultureInfo.InvariantCulture)
            })
        Next
        For Each pair In groups.OrderBy(Function(item) item.Value.Company).
                                  ThenBy(Function(item) item.Value.Year).
                                  ThenBy(Function(item) item.Value.Month).
                                  ThenBy(Function(item) item.Value.Category)
            Dim group As ApprovalGroupAccumulator = pair.Value
            Dim oldTotal As Decimal = 0D
            groupTotals.TryGetValue(pair.Key, oldTotal)
            Dim unchangedOld As Decimal = oldTotal - group.SelectedCurrent
            Dim previewRow As New DraftOtbApprovalPreviewRow With {
                .Company = group.Company,
                .Year = group.Year,
                .Month = group.Month,
                .Category = group.Category,
                .Revised = group.Revised,
                .Diff = oldTotal - group.Revised,
                .TotalBudget = group.Revised + unchangedOld
            }
            result.Add(previewRow)
            AppendApprovalHashFields(canonical, New String() {
                previewRow.Company, previewRow.Year, previewRow.Month, previewRow.Category,
                oldTotal.ToString("0.00############################", CultureInfo.InvariantCulture),
                group.SelectedCurrent.ToString("0.00############################", CultureInfo.InvariantCulture),
                previewRow.Revised.ToString("0.00############################", CultureInfo.InvariantCulture),
                previewRow.Diff.ToString("0.00############################", CultureInfo.InvariantCulture),
                previewRow.TotalBudget.ToString("0.00############################", CultureInfo.InvariantCulture)
            })
        Next
        Return New ApprovalPreviewSnapshot With {
            .Groups = result,
            .PreviewHash = ComputeSha256Hex(canonical.ToString()),
            .CurrentApprovedByBusinessKey = currentByBusinessKey
        }
    End Function

    Private Shared Sub AppendApprovalHashFields(builder As StringBuilder, values As IEnumerable(Of String))
        For Each value As String In values
            Dim normalized As String = If(value, "").Trim()
            builder.Append(normalized.Length.ToString(CultureInfo.InvariantCulture)).Append(":"c).Append(normalized).Append("|"c)
        Next
        builder.Append(ChrW(10))
    End Sub

    Private Shared Function CanonicalApprovalDecimal(value As String) As String
        Dim parsed As Decimal
        If TryParseApprovalDecimal(value, parsed) Then
            Return parsed.ToString("0.00############################", CultureInfo.InvariantCulture)
        End If
        Return If(value, "").Trim()
    End Function

    Private Shared Function ComputeSha256Hex(value As String) As String
        Using sha As SHA256 = SHA256.Create()
            Dim bytes As Byte() = sha.ComputeHash(Encoding.UTF8.GetBytes(If(value, "")))
            Return BitConverter.ToString(bytes).Replace("-", "").ToLowerInvariant()
        End Using
    End Function

    Private Function GetCurrentApprovalGroupTotals(groups As List(Of ApprovalGroupAccumulator)) As Dictionary(Of String, Decimal)
        Dim result As New Dictionary(Of String, Decimal)(StringComparer.OrdinalIgnoreCase)
        If groups Is Nothing OrElse groups.Count = 0 Then Return result

        Dim groupTable As New DataTable()
        groupTable.Columns.Add("Company", GetType(String))
        groupTable.Columns.Add("Year", GetType(Integer))
        groupTable.Columns.Add("Month", GetType(Integer))
        groupTable.Columns.Add("Category", GetType(String))
        For Each group As ApprovalGroupAccumulator In groups
            groupTable.Rows.Add(group.Company, ParseRequiredInteger(group.Year, "Year"), ParseRequiredInteger(group.Month, "Month"), group.Category)
        Next

        Using conn As New SqlConnection(connectionString)
            conn.Open()
            Using createCmd As New SqlCommand("
                CREATE TABLE #ApprovalGroups
                (
                    Company nvarchar(20) NOT NULL,
                    [Year] int NOT NULL,
                    [Month] int NOT NULL,
                    Category nvarchar(20) NOT NULL,
                    PRIMARY KEY (Company, [Year], [Month], Category)
                )", conn)
                createCmd.ExecuteNonQuery()
            End Using
            Using bulk As New SqlBulkCopy(conn)
                bulk.DestinationTableName = "#ApprovalGroups"
                For Each col As DataColumn In groupTable.Columns
                    bulk.ColumnMappings.Add(col.ColumnName, col.ColumnName)
                Next
                bulk.WriteToServer(groupTable)
            End Using

            Dim sql As String = "
                ;WITH BudgetMovement AS
                (
                    SELECT t.Company, t.[Year], t.[Month], t.Category,
                           SUM(CASE WHEN t.[Type] = N'Original' THEN ISNULL(t.Amount, 0)
                                    WHEN t.[Type] = N'Revise' THEN ISNULL(t.RevisedDiff, 0)
                                    ELSE 0 END) AS Amount
                    FROM dbo.OTB_Transaction t
                    INNER JOIN #ApprovalGroups g ON g.Company = t.Company AND g.[Year] = t.[Year]
                        AND g.[Month] = t.[Month] AND g.Category = t.Category
                    WHERE t.OTBStatus = N'Approved'
                    GROUP BY t.Company, t.[Year], t.[Month], t.Category

                    UNION ALL

                    SELECT s.Company, s.[Year], s.[Month], s.Category,
                           SUM(CASE WHEN s.[From] = N'E' THEN ISNULL(s.BudgetAmount, 0)
                                    WHEN s.[From] IN (N'D', N'I', N'G') THEN -ISNULL(s.BudgetAmount, 0)
                                    ELSE 0 END) AS Amount
                    FROM dbo.OTB_Switching_Transaction s
                    INNER JOIN #ApprovalGroups g ON g.Company = s.Company AND g.[Year] = s.[Year]
                        AND g.[Month] = s.[Month] AND g.Category = s.Category
                    WHERE s.OTBStatus = N'Approved'
                    GROUP BY s.Company, s.[Year], s.[Month], s.Category

                    UNION ALL

                    SELECT s.SwitchCompany, s.SwitchYear, s.SwitchMonth, s.SwitchCategory,
                           SUM(CASE WHEN s.[To] IN (N'C', N'F', N'H') THEN ISNULL(s.BudgetAmount, 0) ELSE 0 END) AS Amount
                    FROM dbo.OTB_Switching_Transaction s
                    INNER JOIN #ApprovalGroups g ON g.Company = s.SwitchCompany AND g.[Year] = s.SwitchYear
                        AND g.[Month] = s.SwitchMonth AND g.Category = s.SwitchCategory
                    WHERE s.OTBStatus = N'Approved'
                    GROUP BY s.SwitchCompany, s.SwitchYear, s.SwitchMonth, s.SwitchCategory
                )
                SELECT Company, [Year], [Month], Category, SUM(Amount) AS TotalBudget
                FROM BudgetMovement
                GROUP BY Company, [Year], [Month], Category"
            Using cmd As New SqlCommand(sql, conn)
                cmd.CommandTimeout = 300
                Using reader As SqlDataReader = cmd.ExecuteReader()
                    While reader.Read()
                        Dim key As String = BuildApprovalGroupKey(reader("Company").ToString(), reader("Year").ToString(), reader("Month").ToString(), reader("Category").ToString())
                        result(key) = If(reader("TotalBudget") Is DBNull.Value, 0D, Convert.ToDecimal(reader("TotalBudget")))
                    End While
                End Using
            End Using
        End Using
        Return result
    End Function

    Private Shared Function BuildApprovalGroupKey(company As String, year As String, month As String, category As String) As String
        Return String.Join("|", New String() {company, year, month, category})
    End Function

    Public Shared Function CalculateApprovalPreviewValues(oldTotalBudget As Decimal,
                                                           selectedCurrentBudget As Decimal,
                                                           revisedTarget As Decimal) As DraftOtbApprovalPreviewRow
        Return New DraftOtbApprovalPreviewRow With {
            .Revised = revisedTarget,
            .Diff = oldTotalBudget - revisedTarget,
            .TotalBudget = revisedTarget + (oldTotalBudget - selectedCurrentBudget)
        }
    End Function

    Private Sub ProcessApprovalJob(jobId As Guid)
        If Not DraftOtbJobStore.TryStart(jobId, DraftOtbJobStatuses.Queued, DraftOtbJobStatuses.Validating, "Loading selected draft rows") Then Return
        Dim job As DraftOtbJobRecord = DraftOtbJobStore.GetJob(jobId)
        If job Is Nothing Then Throw New Exception("Approval job was not found.")
        Dim payload As DraftOtbApprovalJobPayload = JsonConvert.DeserializeObject(Of DraftOtbApprovalJobPayload)(job.PayloadJson)
        If payload Is Nothing OrElse payload.RunNos Is Nothing Then Throw New Exception("Approval payload is missing.")
        If payload.RunNos.Count = 0 OrElse payload.RunNos.Count > 15000 Then Throw New Exception("Approval jobs must contain between 1 and 15,000 rows.")
        If Not Regex.IsMatch(If(payload.PreviewHash, ""), "\A[0-9a-f]{64}\z", RegexOptions.CultureInvariant Or RegexOptions.IgnoreCase) Then
            Throw New Exception("Approval preview hash is missing or invalid. SAP was not called.")
        End If

        Dim draftData As DataTable = GetOTBDraftDataByRunNos(payload.RunNos)
        Dim validationResults As New List(Of Dictionary(Of String, Object))()
        Dim loadedIds As New HashSet(Of Integer)(draftData.AsEnumerable().Select(Function(row) Convert.ToInt32(row("RunNo"))))
        For Each requestedId As Integer In payload.RunNos
            If Not loadedIds.Contains(requestedId) Then
                validationResults.Add(CreateApprovalErrorResult(requestedId, Nothing, "Draft row was not found."))
            End If
        Next

        Dim calculator As New OTBBudgetCalculator()
        Dim validator As New OTBValidate()
        Dim preparedByKey As New Dictionary(Of String, ApprovalPreparedRow)(StringComparer.OrdinalIgnoreCase)
        Dim preparedByRunNo As New Dictionary(Of Integer, ApprovalPreparedRow)()
        Dim preparedByBusinessKey As New Dictionary(Of String, ApprovalPreparedRow)(StringComparer.OrdinalIgnoreCase)
        For i As Integer = 0 To draftData.Rows.Count - 1
            Dim row As DataRow = draftData.Rows(i)
            Dim currentRunNo As Integer = Convert.ToInt32(row("RunNo"))
            Dim rowErrors As List(Of String) = ValidateApprovalDraftRow(row, validator)
            Dim prepared As ApprovalPreparedRow = Nothing
            If rowErrors.Count = 0 Then
                Try
                    prepared = PrepareApprovalRow(row, calculator)
                    If preparedByBusinessKey.ContainsKey(prepared.BusinessKey) Then
                        rowErrors.Add("Another selected row has the same approval budget dimensions. Amount and version do not make it a separate approval key.")
                    Else
                        preparedByBusinessKey(prepared.BusinessKey) = prepared
                        preparedByKey(prepared.SapKey) = prepared
                        preparedByRunNo(currentRunNo) = prepared
                    End If
                Catch ex As Exception
                    rowErrors.Add(ex.Message)
                End Try
            End If
            If rowErrors.Count > 0 Then
                validationResults.Add(CreateApprovalErrorResult(currentRunNo, row, String.Join(" | ", rowErrors)))
            End If

            If (i + 1) Mod 100 = 0 OrElse i = draftData.Rows.Count - 1 Then
                Dim progress As Integer = 5 + CInt(Math.Floor((i + 1) * 35.0 / Math.Max(1, draftData.Rows.Count)))
                ' Success/error counts represent committed business results. Keep both
                ' at zero until a terminal validation result or the atomic DB commit.
                DraftOtbJobStore.UpdateProgress(jobId, DraftOtbJobStatuses.Validating, "Validating every selected row", progress, i + 1, 0, 0,
                                                expectedStatus:=DraftOtbJobStatuses.Validating)
            End If
        Next

        If validationResults.Count > 0 Then
            DraftOtbJobStore.Complete(jobId,
                DraftOtbJobStatuses.ValidationFailed,
                "Validation failed",
                $"Validation failed for {validationResults.Count:N0} rows. SAP was not called and no business data was changed.",
                payload.RunNos.Count,
                0,
                validationResults.Count,
                JsonConvert.SerializeObject(validationResults))
            Return
        End If

        Dim approvalDimensions As New List(Of OTBSwitchBudgetGuard.BudgetDimension)()
        Dim claims As New List(Of DraftOtbApprovalClaim)()
        For Each runNo As Integer In payload.RunNos
            Dim prepared As ApprovalPreparedRow = preparedByRunNo(runNo)
            Dim dimension As OTBSwitchBudgetGuard.BudgetDimension = OTBSwitchBudgetGuard.CreateDimension(
                GetRequiredDraftText(prepared.Data, "OTBYear", "Year"),
                GetRequiredDraftText(prepared.Data, "OTBMonth", "Month"),
                GetRequiredDraftText(prepared.Data, "OTBCompany", "Company"),
                GetRequiredDraftText(prepared.Data, "OTBCategory", "Category"),
                GetRequiredDraftText(prepared.Data, "OTBSegment", "Segment"),
                GetRequiredDraftText(prepared.Data, "OTBBrand", "Brand"),
                GetRequiredDraftText(prepared.Data, "OTBVendor", "Vendor"))
            approvalDimensions.Add(dimension)
            claims.Add(New DraftOtbApprovalClaim With {
                .RunNo = runNo,
                .BusinessKey = prepared.BusinessKey,
                .BusinessKeyHash = ComputeSha256Hex(prepared.BusinessKey),
                .OtbType = GetRequiredDraftText(prepared.Data, "OTBType", "Type"),
                .Year = dimension.Year,
                .Month = dimension.Month,
                .Category = dimension.Category,
                .Company = dimension.Company,
                .Segment = dimension.Segment,
                .Brand = dimension.Brand,
                .Vendor = dimension.Vendor
            })
        Next

        ' A canonical group lock serializes approval preview calculation with all
        ' approved-budget movement flows. Durable detail claims protect every RunNo.
        Using approvalLockConn As New SqlConnection(connectionString)
            approvalLockConn.Open()
            Using approvalLockHandle As IDisposable = OTBSwitchBudgetGuard.AcquireBudgetLocks(
                approvalLockConn, OTBSwitchBudgetGuard.BuildGroupLockResources(approvalDimensions))
                Dim claimConflict As String = Nothing
                If Not DraftOtbJobStore.TryAcquireApprovalClaims(jobId, claims, claimConflict) Then
                    DraftOtbJobStore.Complete(jobId,
                        DraftOtbJobStatuses.ValidationFailed,
                        "Approval already in progress or no longer Draft",
                        claimConflict,
                        payload.RunNos.Count, 0, payload.RunNos.Count)
                    Return
                End If

                ' Recompute only after both interlocks are held so the hash covers the
                ' exact Draft rows and live group totals that will be sent to SAP.
                Dim currentSnapshot As ApprovalPreviewSnapshot = BuildApprovalPreview(payload.RunNos)
                If Not String.Equals(currentSnapshot.PreviewHash, payload.PreviewHash, StringComparison.OrdinalIgnoreCase) Then
                    DraftOtbJobStore.ReleaseApprovalClaims(jobId)
                    DraftOtbJobStore.Complete(jobId,
                        DraftOtbJobStatuses.ValidationFailed,
                        "Approval preview is stale",
                        "The selected Draft OTB or current approved budget changed after preview. Please preview and confirm again. SAP was not called.",
                        payload.RunNos.Count, 0, payload.RunNos.Count)
                    Return
                End If

        ' Use the exact current-approved values that participated in the matched hash
        ' for both the SAP revision delta and the later database RevisedDiff.
        preparedByKey.Clear()
        For Each runNo As Integer In payload.RunNos
            Dim prepared As ApprovalPreparedRow = preparedByRunNo(runNo)
            Dim currentApproved As Decimal
            If Not currentSnapshot.CurrentApprovedByBusinessKey.TryGetValue(prepared.BusinessKey, currentApproved) Then
                Throw New Exception($"RunNo {runNo}: current approved budget snapshot is missing.")
            End If
            prepared.CurrentApproved = currentApproved
            Dim target As Decimal
            If Not TryParseApprovalDecimal(GetRequiredDraftText(prepared.Data, "Amount", "Amount"), target) Then
                Throw New Exception($"RunNo {runNo}: Amount must be numeric.")
            End If
            Dim typeValue As String = GetRequiredDraftText(prepared.Data, "OTBType", "Type")
            Dim versionValue As String = GetRequiredDraftText(prepared.Data, "Version", "Version")
            Dim sapAmount As Decimal = If(versionValue.StartsWith("R", StringComparison.OrdinalIgnoreCase) OrElse
                                             typeValue.Equals("Revise", StringComparison.OrdinalIgnoreCase),
                                             target - currentApproved, target)
            prepared.SapItem.Amount = sapAmount.ToString("F2", CultureInfo.InvariantCulture)
            prepared.SapKey = BuildSapApprovalKey(versionValue,
                GetRequiredDraftText(prepared.Data, "OTBCompany", "Company"),
                GetRequiredDraftText(prepared.Data, "OTBCategory", "Category"),
                GetRequiredDraftText(prepared.Data, "OTBVendor", "Vendor"),
                GetRequiredDraftText(prepared.Data, "OTBSegment", "Segment"),
                GetRequiredDraftText(prepared.Data, "OTBBrand", "Brand"),
                prepared.SapItem.Amount,
                GetRequiredDraftText(prepared.Data, "OTBYear", "Year"),
                GetRequiredDraftText(prepared.Data, "OTBMonth", "Month"))
            preparedByKey.Add(prepared.SapKey, prepared)
        Next

                Dim plans As List(Of OtbPlanUploadItem) = payload.RunNos.Select(Function(runNo) preparedByRunNo(runNo).SapItem).ToList()
                DraftOtbJobStore.UpdateProgress(jobId, DraftOtbJobStatuses.PreparingSap, "Preparing one all-or-nothing SAP request", 48, payload.RunNos.Count, 0, 0,
                                                expectedStatus:=DraftOtbJobStatuses.Validating)
                DraftOtbJobStore.UpdateProgress(jobId, DraftOtbJobStatuses.SendingToSap, "Waiting for SAP", 55, payload.RunNos.Count, 0, 0,
                                                expectedStatus:=DraftOtbJobStatuses.PreparingSap)

        Dim sapResponse As SapApiResponse(Of SapUploadResultItem) = Task.Run(Async Function()
                                                                                 Return Await SapApiHelper.UploadOtbPlanAsync(plans)
                                                                             End Function).Result
        If sapResponse Is Nothing OrElse sapResponse.Status Is Nothing OrElse sapResponse.Results Is Nothing Then
            Throw New Exception("SAP returned an incomplete response. The outcome must be reconciled before retrying.")
        End If

        Dim detailedResults As New List(Of Dictionary(Of String, Object))()
        Dim mappedSuccess As New List(Of ApprovalMappedResult)()
        Dim mappingErrors As Integer = 0
        For i As Integer = 0 To sapResponse.Results.Count - 1
            Dim sapResult As SapUploadResultItem = sapResponse.Results(i)
            Dim prepared As ApprovalPreparedRow = FindPreparedRow(sapResult, preparedByKey)
            If prepared Is Nothing Then
                mappingErrors += 1
                detailedResults.Add(CreateSapUnmappedResult(sapResult))
            Else
                detailedResults.Add(CreateApprovalSapResult(prepared.Data, sapResult))
                If String.Equals(sapResult.MessageType, "S", StringComparison.OrdinalIgnoreCase) Then
                    mappedSuccess.Add(New ApprovalMappedResult With {.RunNo = prepared.RunNo, .SapResult = sapResult})
                End If
            End If
            If (i + 1) Mod 100 = 0 OrElse i = sapResponse.Results.Count - 1 Then
                Dim progress As Integer = 58 + CInt(Math.Floor((i + 1) * 14.0 / Math.Max(1, sapResponse.Results.Count)))
                DraftOtbJobStore.UpdateProgress(jobId, DraftOtbJobStatuses.SendingToSap, "Checking SAP response", progress, i + 1, 0, 0,
                                                expectedStatus:=DraftOtbJobStatuses.SendingToSap)
            End If
        Next

        Dim fullSuccess As Boolean = sapResponse.Status.Total = payload.RunNos.Count AndAlso
                                     sapResponse.Status.Success = payload.RunNos.Count AndAlso
                                     sapResponse.Status.ErrorCount = 0 AndAlso
                                     sapResponse.Results.Count = payload.RunNos.Count AndAlso
                                     mappedSuccess.Count = payload.RunNos.Count AndAlso
                                     mappingErrors = 0
        If Not fullSuccess Then
            Dim sapAcceptedWholeBatch As Boolean = sapResponse.Status.Total = payload.RunNos.Count AndAlso
                                                   sapResponse.Status.Success = payload.RunNos.Count AndAlso
                                                   sapResponse.Status.ErrorCount = 0
            Dim explicitlyRejectedWholeBatch As Boolean = SapExplicitlyRejectedWholeApprovalBatch(sapResponse, payload.RunNos.Count)
            If explicitlyRejectedWholeBatch Then
                DraftOtbJobStore.ReleaseApprovalClaims(jobId)
            End If
            Dim failureStatus As String = If(explicitlyRejectedWholeBatch,
                                             DraftOtbJobStatuses.Failed,
                                             DraftOtbJobStatuses.ReconciliationRequired)
            Dim failureStage As String = If(explicitlyRejectedWholeBatch,
                                            "SAP rejected the approval batch",
                                            "Manual reconciliation required")
            Dim failureMessage As String
            If sapAcceptedWholeBatch Then
                failureMessage = "SAP accepted the complete batch, but one or more response rows could not be mapped safely. Do not retry; reconcile this Job ID before changing database data."
            ElseIf explicitlyRejectedWholeBatch Then
                failureMessage = $"SAP explicitly rejected the complete batch. Total: {sapResponse.Status.Total}, success: {sapResponse.Status.Success}, errors: {sapResponse.Status.ErrorCount}. Claims were released and no database rows were approved."
            Else
                failureMessage = $"SAP returned a partial or ambiguous result. Total: {sapResponse.Status.Total}, success: {sapResponse.Status.Success}, errors: {sapResponse.Status.ErrorCount}. Do not retry; reconcile this Job ID."
            End If
            DraftOtbJobStore.Complete(jobId,
                failureStatus,
                failureStage,
                failureMessage,
                payload.RunNos.Count,
                0,
                Math.Max(1, payload.RunNos.Count - mappedSuccess.Count),
                JsonConvert.SerializeObject(detailedResults),
                expectedStatus:=DraftOtbJobStatuses.SendingToSap)
            Return
        End If

        ' Persist the complete row-level SAP evidence before the business-data
        ' transaction. If the DB save later fails, ReconciliationRequired still
        ' has the response needed to compare SAP against BMS safely.
        Dim detailedResultsJson As String = JsonConvert.SerializeObject(detailedResults)
        DraftOtbJobStore.SetResult(jobId, detailedResultsJson, DraftOtbJobStatuses.SendingToSap)
        DraftOtbJobStore.UpdateProgress(jobId, DraftOtbJobStatuses.SavingApproval, "SAP accepted all rows; saving one database transaction", 75, payload.RunNos.Count, 0, 0,
                                        expectedStatus:=DraftOtbJobStatuses.SendingToSap)
        SaveApprovedRows(jobId, mappedSuccess, preparedByRunNo, payload.ApprovedBy,
                         detailedResultsJson)
            End Using
        End Using
    End Sub

    Private Shared Function SapExplicitlyRejectedWholeApprovalBatch(response As SapApiResponse(Of SapUploadResultItem), expectedRows As Integer) As Boolean
        If response Is Nothing OrElse response.Status Is Nothing OrElse response.Results Is Nothing Then Return False
        If response.Status.Total <> expectedRows OrElse response.Status.Success <> 0 OrElse response.Status.ErrorCount <> expectedRows Then Return False
        If response.Results.Count <> expectedRows Then Return False
        Return response.Results.All(Function(item) item IsNot Nothing AndAlso Not String.Equals(item.MessageType, "S", StringComparison.OrdinalIgnoreCase))
    End Function

    Private Sub EnsureEveryDraftWasLoaded(runNos As List(Of Integer), draftData As DataTable)
        Dim loaded As New HashSet(Of Integer)(draftData.AsEnumerable().Select(Function(row) Convert.ToInt32(row("RunNo"))))
        Dim missing As List(Of Integer) = runNos.Where(Function(runNo) Not loaded.Contains(runNo)).ToList()
        If missing.Count > 0 Then
            Throw New Exception($"{missing.Count:N0} selected draft row(s) were not found or are no longer available.")
        End If
    End Sub

    Private Function ValidateApprovalDraftRow(row As DataRow, validator As OTBValidate) As List(Of String)
        Dim errors As New List(Of String)()
        Dim status As String = GetDraftText(row, "OTBStatus")
        If Not status.Equals("Draft", StringComparison.OrdinalIgnoreCase) Then
            errors.Add("Status must be Draft.")
        End If
        For Each field As String In New String() {"Version", "OTBCompany", "OTBCategory", "OTBVendor", "OTBSegment", "OTBBrand", "OTBYear", "OTBMonth", "Amount", "OTBType"}
            If String.IsNullOrWhiteSpace(GetDraftText(row, field)) Then errors.Add(field & " is required.")
        Next

        Dim typeValue As String = GetDraftText(row, "OTBType")
        AddApprovalValidationError(errors, validator.ValidateType(typeValue))
        Dim versionValue As String = GetDraftText(row, "Version")
        If Not Regex.IsMatch(versionValue, "\A(?:A1|R(?:[1-9]|1[0-5]))\z", RegexOptions.CultureInvariant) Then
            errors.Add("Version must be A1 or R1 through R15.")
        End If

        Dim yearValue As Integer
        If Not Integer.TryParse(GetDraftText(row, "OTBYear"), NumberStyles.Integer, CultureInfo.InvariantCulture, yearValue) Then
            errors.Add("Year must be numeric.")
        Else
            AddApprovalValidationError(errors, validator.ValidateYear(yearValue))
        End If
        Dim monthValue As Short
        If Not Short.TryParse(GetDraftText(row, "OTBMonth"), NumberStyles.Integer, CultureInfo.InvariantCulture, monthValue) Then
            errors.Add("Month must be numeric.")
        Else
            AddApprovalValidationError(errors, validator.ValidateMonth(monthValue))
        End If
        Dim amountValue As Decimal
        If Not TryParseApprovalDecimal(GetDraftText(row, "Amount"), amountValue) Then
            errors.Add("Amount must be numeric.")
        Else
            AddApprovalValidationError(errors, validator.ValidateAmount(amountValue))
            If amountValue > 9999999999999999.99D OrElse amountValue < -9999999999999999.99D Then
                errors.Add("Amount exceeds the database decimal(18,2) range.")
            ElseIf Decimal.Round(amountValue, 2, MidpointRounding.AwayFromZero) <> amountValue Then
                errors.Add("Amount must not contain more than 2 decimal places.")
            End If
        End If
        AddApprovalValidationError(errors, validator.ValidateCategory(GetDraftText(row, "OTBCategory")))
        AddApprovalValidationError(errors, validator.ValidateCompany(GetDraftText(row, "OTBCompany")))
        AddApprovalValidationError(errors, validator.ValidateSegment(GetDraftText(row, "OTBSegment")))
        AddApprovalValidationError(errors, validator.ValidateBrand(GetDraftText(row, "OTBBrand")))
        AddApprovalValidationError(errors, validator.ValidateVendor(GetDraftText(row, "OTBVendor")))
        ValidateApprovalTextLength(errors, row, "OTBCompany", 20, "Company")
        ValidateApprovalTextLength(errors, row, "OTBCategory", 20, "Category")
        ValidateApprovalTextLength(errors, row, "OTBSegment", 20, "Segment")
        ValidateApprovalTextLength(errors, row, "OTBBrand", 30, "Brand")
        ValidateApprovalTextLength(errors, row, "OTBVendor", 30, "Vendor")
        ValidateApprovalTextLength(errors, row, "Version", 20, "Version")
        ValidateApprovalTextLength(errors, row, "OTBType", 20, "Type")
        ValidateApprovalTextLength(errors, row, "Remark", 500, "Remark")
        ValidateApprovalTextLength(errors, row, "CateName", 200, "Category name")
        ValidateApprovalTextLength(errors, row, "SegmentName", 200, "Segment name")
        ValidateApprovalTextLength(errors, row, "BrandName", 200, "Brand name")
        ValidateApprovalTextLength(errors, row, "Vendor", 200, "Vendor name")
        Return errors.Distinct(StringComparer.OrdinalIgnoreCase).ToList()
    End Function

    Private Shared Sub ValidateApprovalTextLength(errors As List(Of String), row As DataRow,
                                                   fieldName As String, maxLength As Integer,
                                                   displayName As String)
        Dim value As String = GetDraftText(row, fieldName)
        If value.Length > maxLength Then errors.Add($"{displayName} is too long ({value.Length}/{maxLength} characters).")
    End Sub

    Private Shared Sub AddApprovalValidationError(errors As List(Of String), validationText As String)
        If Not String.IsNullOrWhiteSpace(validationText) Then errors.Add(validationText.Trim())
    End Sub

    Private Shared Function TryParseApprovalDecimal(value As String, ByRef parsed As Decimal) As Boolean
        Return Decimal.TryParse(value, NumberStyles.Number, CultureInfo.InvariantCulture, parsed) OrElse
               Decimal.TryParse(value, NumberStyles.Number, CultureInfo.CurrentCulture, parsed)
    End Function

    Private Function PrepareApprovalRow(row As DataRow, calculator As OTBBudgetCalculator) As ApprovalPreparedRow
        Dim runNo As Integer = Convert.ToInt32(row("RunNo"))
        Dim version As String = GetRequiredDraftText(row, "Version", "Version")
        Dim company As String = GetRequiredDraftText(row, "OTBCompany", "Company")
        Dim category As String = GetRequiredDraftText(row, "OTBCategory", "Category")
        Dim vendor As String = GetRequiredDraftText(row, "OTBVendor", "Vendor")
        Dim segment As String = GetRequiredDraftText(row, "OTBSegment", "Segment")
        Dim brand As String = GetRequiredDraftText(row, "OTBBrand", "Brand")
        Dim year As String = GetRequiredDraftText(row, "OTBYear", "Year")
        Dim month As String = GetRequiredDraftText(row, "OTBMonth", "Month")
        Dim type As String = GetRequiredDraftText(row, "OTBType", "Type")
        Dim target As Decimal
        If Not TryParseApprovalDecimal(GetRequiredDraftText(row, "Amount", "Amount"), target) Then Throw New Exception("Amount must be numeric.")
        Dim currentApproved As Decimal = calculator.CalculateCurrentApprovedBudget(year, month, category, company, segment, brand, vendor)
        Dim amountToSap As Decimal = target
        If version.StartsWith("R", StringComparison.OrdinalIgnoreCase) OrElse type.Equals("Revise", StringComparison.OrdinalIgnoreCase) Then
            amountToSap = target - currentApproved
        End If
        Dim amount As String = amountToSap.ToString("F2", CultureInfo.InvariantCulture)
        Dim item As New OtbPlanUploadItem With {
            .Version = version, .CompCode = company, .Category = category,
            .VendorCode = vendor, .SegmentCode = segment, .BrandCode = brand,
            .Amount = amount, .Year = year, .Month = month,
            .Remark = GetDraftText(row, "Remark")
        }
        Return New ApprovalPreparedRow With {
            .RunNo = runNo,
            .Data = row,
            .SapItem = item,
            .SapKey = BuildSapApprovalKey(version, company, category, vendor, segment, brand, amount, year, month),
            .BusinessKey = BuildApprovalBusinessKey(year, month, company, category, segment, brand, vendor),
            .CurrentApproved = currentApproved
        }
    End Function

    Private Shared Function BuildApprovalBusinessKey(year As String, month As String, company As String,
                                                     category As String, segment As String, brand As String,
                                                     vendor As String) As String
        Dim yearValue As Integer
        Dim monthValue As Integer
        Integer.TryParse(year, NumberStyles.Integer, CultureInfo.InvariantCulture, yearValue)
        Integer.TryParse(month, NumberStyles.Integer, CultureInfo.InvariantCulture, monthValue)
        Dim canonical As New StringBuilder("approval-business-v1|")
        AppendApprovalHashFields(canonical, New String() {
            yearValue.ToString(CultureInfo.InvariantCulture),
            monthValue.ToString(CultureInfo.InvariantCulture),
            If(company, "").Trim().ToUpperInvariant(),
            If(category, "").Trim().ToUpperInvariant(),
            If(segment, "").Trim().ToUpperInvariant(),
            If(brand, "").Trim().ToUpperInvariant(),
            If(vendor, "").Trim().ToUpperInvariant()
        })
        Return canonical.ToString()
    End Function

    Private Function FindPreparedRow(result As SapUploadResultItem, preparedByKey As Dictionary(Of String, ApprovalPreparedRow)) As ApprovalPreparedRow
        Dim candidates As New List(Of String) From {
            BuildSapApprovalKey(result.Version, result.CompCode, result.Category, result.VendorCode, result.SegmentCode, result.BrandCode, result.Amount, result.Year, result.Month),
            BuildSapApprovalKey(result.Version, result.CompCode, result.Category, result.VendorMap, result.SegmentCodeMap, result.BrandCode, result.Amount, result.Year, result.Month)
        }
        For Each key As String In candidates.Distinct(StringComparer.OrdinalIgnoreCase)
            Dim prepared As ApprovalPreparedRow = Nothing
            If preparedByKey.TryGetValue(key, prepared) Then Return prepared
        Next
        Return Nothing
    End Function

    Private Shared Function BuildSapApprovalKey(version As String, company As String, category As String,
                                                 vendor As String, segment As String, brand As String,
                                                 amount As String, year As String, month As String) As String
        Dim numericAmount As Decimal
        Dim canonicalAmount As String = If(Decimal.TryParse(amount, NumberStyles.Any, CultureInfo.InvariantCulture, numericAmount),
                                           numericAmount.ToString("F2", CultureInfo.InvariantCulture), If(amount, "").Trim())
        Return String.Join("|", New String() {version, company, category, vendor, segment, brand, canonicalAmount, year, month})
    End Function

    Private Function CreateApprovalErrorResult(runNo As Integer, row As DataRow, message As String) As Dictionary(Of String, Object)
        Dim result As Dictionary(Of String, Object) = CreateApprovalResultBase(row)
        result("RunNo") = runNo
        result("SAP_MessageType") = "E"
        result("SAP_Message") = message
        Return result
    End Function

    Private Function CreateApprovalSapResult(row As DataRow, sapResult As SapUploadResultItem) As Dictionary(Of String, Object)
        Dim result As Dictionary(Of String, Object) = CreateApprovalResultBase(row)
        result("SAP_MessageType") = sapResult.MessageType
        result("SAP_Message") = sapResult.Message
        Return result
    End Function

    Private Function CreateApprovalResultBase(row As DataRow) As Dictionary(Of String, Object)
        Dim result As New Dictionary(Of String, Object)()
        For Each name As String In New String() {"OTBYear", "OTBMonth", "OTBCategory", "CateName", "CompanyName", "OTBSegment", "SegmentName", "OTBBrand", "BrandName", "OTBVendor", "Vendor", "Amount", "Remark"}
            result(name) = If(row Is Nothing OrElse Not row.Table.Columns.Contains(name) OrElse row(name) Is DBNull.Value, "", row(name))
        Next
        Return result
    End Function

    Private Function CreateSapUnmappedResult(sapResult As SapUploadResultItem) As Dictionary(Of String, Object)
        Return New Dictionary(Of String, Object) From {
            {"OTBYear", sapResult.Year}, {"OTBMonth", sapResult.Month}, {"OTBCategory", sapResult.Category},
            {"OTBSegment", sapResult.SegmentCode}, {"OTBBrand", sapResult.BrandCode}, {"OTBVendor", sapResult.VendorCode},
            {"Amount", sapResult.Amount}, {"Remark", sapResult.Remark}, {"SAP_MessageType", "E"},
            {"SAP_Message", "SAP response could not be mapped back to a selected Draft OTB row. " & If(sapResult.Message, "")}
        }
    End Function

    Private Shared Function GetDraftText(row As DataRow, name As String) As String
        If row Is Nothing OrElse row.Table Is Nothing OrElse Not row.Table.Columns.Contains(name) OrElse row(name) Is DBNull.Value Then Return ""
        Return row(name).ToString().Trim()
    End Function

    Private Function GetRequiredDraftText(row As DataRow, name As String, displayName As String) As String
        Dim value As String = GetDraftText(row, name)
        If String.IsNullOrWhiteSpace(value) Then Throw New Exception(displayName & " is required.")
        Return value
    End Function

    Private Function SaveApprovedRows(jobId As Guid,
                                      mappedResults As List(Of ApprovalMappedResult),
                                      preparedByRunNo As Dictionary(Of Integer, ApprovalPreparedRow),
                                      approvedBy As String,
                                      resultJson As String) As Integer
        If mappedResults Is Nothing OrElse mappedResults.Count = 0 Then Return 0
        Dim stage As New DataTable()
        stage.Columns.Add("RunNo", GetType(Integer))
        stage.Columns.Add("Type", GetType(String))
        stage.Columns.Add("Year", GetType(Integer))
        stage.Columns.Add("Month", GetType(Integer))
        stage.Columns.Add("Category", GetType(String))
        stage.Columns.Add("Company", GetType(String))
        stage.Columns.Add("Segment", GetType(String))
        stage.Columns.Add("Brand", GetType(String))
        stage.Columns.Add("Vendor", GetType(String))
        stage.Columns.Add("Version", GetType(String))
        stage.Columns.Add("Amount", GetType(Decimal))
        stage.Columns.Add("RevisedDiff", GetType(Decimal))
        stage.Columns.Add("Remark", GetType(String))
        stage.Columns.Add("CategoryName", GetType(String))
        stage.Columns.Add("SegmentName", GetType(String))
        stage.Columns.Add("BrandName", GetType(String))
        stage.Columns.Add("VendorName", GetType(String))
        stage.Columns.Add("SAPStatus", GetType(String))
        stage.Columns.Add("SAPErrorMessage", GetType(String))

        For Each mapped As ApprovalMappedResult In mappedResults
            Dim prepared As ApprovalPreparedRow = Nothing
            If Not preparedByRunNo.TryGetValue(mapped.RunNo, prepared) Then
                Throw New Exception($"Could not map approved RunNo {mapped.RunNo} to its source row.")
            End If
            Dim source As DataRow = prepared.Data
            Dim amount As Decimal
            If Not TryParseApprovalDecimal(GetRequiredDraftText(source, "Amount", "Amount"), amount) Then
                Throw New Exception($"RunNo {mapped.RunNo}: Amount must be numeric.")
            End If
            Dim typeValue As String = GetRequiredDraftText(source, "OTBType", "Type")
            Dim revisedDiff As Decimal = If(typeValue.Equals("Revise", StringComparison.OrdinalIgnoreCase),
                                               amount - prepared.CurrentApproved, 0D)
            stage.Rows.Add(
                mapped.RunNo,
                ApprovalStageText(typeValue, 20, "Type"),
                ParseRequiredInteger(GetRequiredDraftText(source, "OTBYear", "Year"), "Year"),
                ParseRequiredInteger(GetRequiredDraftText(source, "OTBMonth", "Month"), "Month"),
                ApprovalStageText(GetRequiredDraftText(source, "OTBCategory", "Category"), 20, "Category"),
                ApprovalStageText(GetRequiredDraftText(source, "OTBCompany", "Company"), 20, "Company"),
                ApprovalStageText(GetRequiredDraftText(source, "OTBSegment", "Segment"), 20, "Segment"),
                ApprovalStageText(GetRequiredDraftText(source, "OTBBrand", "Brand"), 30, "Brand"),
                ApprovalStageText(GetRequiredDraftText(source, "OTBVendor", "Vendor"), 30, "Vendor"),
                ApprovalStageText(GetRequiredDraftText(source, "Version", "Version"), 20, "Version"),
                amount, revisedDiff,
                ApprovalStageNullableText(GetDraftText(source, "Remark"), 500, "Remark"),
                ApprovalStageNullableText(GetDraftText(source, "CateName"), 200, "CategoryName"),
                ApprovalStageNullableText(GetDraftText(source, "SegmentName"), 200, "SegmentName"),
                ApprovalStageNullableText(GetDraftText(source, "BrandName"), 200, "BrandName"),
                ApprovalStageNullableText(GetDraftText(source, "Vendor"), 200, "VendorName"),
                ApprovalStageText(mapped.SapResult.MessageType, 10, "SAPStatus"),
                ApprovalStageTruncatedNullableText(mapped.SapResult.Message, 1000))
        Next

        DraftOtbJobStore.UpdateProgress(jobId, DraftOtbJobStatuses.SavingApproval,
                                        "Staging approval rows for one database transaction", 82,
                                        mappedResults.Count, 0, 0,
                                        expectedStatus:=DraftOtbJobStatuses.SavingApproval)
        Using conn As New SqlConnection(connectionString)
            conn.Open()
            Using transaction As SqlTransaction = conn.BeginTransaction(IsolationLevel.Serializable)
                Try
                    Using createCmd As New SqlCommand("
                        CREATE TABLE #ApprovalStage
                        (
                            RunNo int NOT NULL PRIMARY KEY,
                            [Type] nvarchar(20) NOT NULL,
                            [Year] int NOT NULL,
                            [Month] int NOT NULL,
                            Category nvarchar(20) NOT NULL,
                            Company nvarchar(20) NOT NULL,
                            Segment nvarchar(20) NOT NULL,
                            Brand nvarchar(30) NOT NULL,
                            Vendor nvarchar(30) NOT NULL,
                            [Version] nvarchar(20) NOT NULL,
                            Amount decimal(18,2) NOT NULL,
                            RevisedDiff decimal(18,2) NOT NULL,
                            Remark nvarchar(500) NULL,
                            CategoryName nvarchar(200) NULL,
                            SegmentName nvarchar(200) NULL,
                            BrandName nvarchar(200) NULL,
                            VendorName nvarchar(200) NULL,
                            SAPStatus nvarchar(10) NOT NULL,
                            SAPErrorMessage nvarchar(1000) NULL,
                            UNIQUE ([Type], [Year], [Month], Category, Company, Segment, Brand, Vendor, [Version])
                        )", conn, transaction)
                        createCmd.CommandTimeout = 600
                        createCmd.ExecuteNonQuery()
                    End Using
                    Using bulk As New SqlBulkCopy(conn, SqlBulkCopyOptions.CheckConstraints Or SqlBulkCopyOptions.KeepNulls, transaction)
                        bulk.DestinationTableName = "#ApprovalStage"
                        bulk.BulkCopyTimeout = 600
                        bulk.BatchSize = 2000
                        For Each column As DataColumn In stage.Columns
                            bulk.ColumnMappings.Add(column.ColumnName, column.ColumnName)
                        Next
                        bulk.WriteToServer(stage)
                    End Using

                    Dim updateCount As Integer
                    Using updateCmd As New SqlCommand("
                        UPDATE d WITH (UPDLOCK, HOLDLOCK)
                        SET d.OTBStatus = N'Approved', d.UpdateBy = @ApprovedBy, d.UpdateDT = GETDATE(),
                            d.SAPStatus = s.SAPStatus, d.SAPErrorMessage = s.SAPErrorMessage
                        FROM dbo.Template_Upload_Draft_OTB d
                        INNER JOIN #ApprovalStage s ON s.RunNo = d.RunNo
                        WHERE d.OTBStatus IS NULL OR d.OTBStatus = N'Draft'", conn, transaction)
                        updateCmd.CommandTimeout = 600
                        AddNVarCharParameter(updateCmd, "@ApprovedBy", approvedBy, 100, "ApprovedBy")
                        updateCount = updateCmd.ExecuteNonQuery()
                    End Using
                    If updateCount <> mappedResults.Count Then
                        Throw New Exception($"Only {updateCount:N0} of {mappedResults.Count:N0} rows were still Draft. The entire database transaction was rolled back.")
                    End If

                    Dim mergeCount As Integer
                    Using mergeCmd As New SqlCommand("
                        DECLARE @MergeActions TABLE (ActionName nvarchar(10) NOT NULL);
                        MERGE dbo.OTB_Transaction WITH (HOLDLOCK) AS T
                        USING #ApprovalStage AS S
                        ON T.[Type] = S.[Type] AND T.[Year] = S.[Year] AND T.[Month] = S.[Month]
                           AND T.Category = S.Category AND T.Company = S.Company AND T.Segment = S.Segment
                           AND T.Brand = S.Brand AND T.Vendor = S.Vendor AND T.[Version] = S.[Version]
                        WHEN MATCHED THEN UPDATE SET
                            T.Amount = S.Amount, T.RevisedDiff = S.RevisedDiff, T.Remark = S.Remark,
                            T.ApprovedDate = GETDATE(), T.SAPDate = GETDATE(), T.ActionBy = @ApprovedBy,
                            T.DraftID = S.RunNo, T.SAPStatus = S.SAPStatus, T.SAPErrorMessage = S.SAPErrorMessage,
                            T.CategoryName = S.CategoryName, T.SegmentName = S.SegmentName,
                            T.BrandName = S.BrandName, T.VendorName = S.VendorName, T.OTBStatus = N'Approved'
                        WHEN NOT MATCHED THEN INSERT
                            (CreateDate, [Type], [Year], [Month], Category, CategoryName, Company,
                             Segment, SegmentName, Brand, BrandName, Vendor, VendorName, Amount,
                             RevisedDiff, Remark, OTBStatus, ApprovedDate, SAPDate, ActionBy, DraftID,
                             SAPStatus, SAPErrorMessage, [Version])
                        VALUES
                            (GETDATE(), S.[Type], S.[Year], S.[Month], S.Category, S.CategoryName, S.Company,
                             S.Segment, S.SegmentName, S.Brand, S.BrandName, S.Vendor, S.VendorName, S.Amount,
                             S.RevisedDiff, S.Remark, N'Approved', GETDATE(), GETDATE(), @ApprovedBy, S.RunNo,
                             S.SAPStatus, S.SAPErrorMessage, S.[Version])
                        OUTPUT $action INTO @MergeActions;
                        SELECT COUNT_BIG(*) FROM @MergeActions;", conn, transaction)
                        mergeCmd.CommandTimeout = 600
                        AddNVarCharParameter(mergeCmd, "@ApprovedBy", approvedBy, 100, "ApprovedBy")
                        mergeCount = Convert.ToInt32(mergeCmd.ExecuteScalar())
                    End Using
                    If mergeCount <> mappedResults.Count Then
                        Throw New Exception($"Expected to save {mappedResults.Count:N0} approved transactions but SQL affected {mergeCount:N0}. The entire database transaction was rolled back.")
                    End If

                    Using completeCmd As New SqlCommand("
                        UPDATE dbo.Draft_OTB_Background_Job
                        SET Status = @Status, Stage = @Stage, ProgressPercent = 100,
                            TotalRows = @TotalRows, ProcessedRows = @TotalRows,
                            SuccessRows = @TotalRows, ErrorRows = 0,
                            Message = @Message, ResultJson = @ResultJson,
                            FinishedAt = SYSUTCDATETIME(), UpdatedAt = SYSUTCDATETIME()
                        WHERE JobID = @JobID AND Status = @ExpectedStatus", conn, transaction)
                        completeCmd.CommandTimeout = 600
                        completeCmd.Parameters.Add("@JobID", SqlDbType.UniqueIdentifier).Value = jobId
                        AddNVarCharParameter(completeCmd, "@Status", DraftOtbJobStatuses.Completed, 30, "Job status")
                        AddNVarCharParameter(completeCmd, "@Stage", "Approval complete", 100, "Job stage")
                        completeCmd.Parameters.Add("@TotalRows", SqlDbType.Int).Value = mappedResults.Count
                        AddNVarCharParameter(completeCmd, "@ExpectedStatus", DraftOtbJobStatuses.SavingApproval, 30, "Expected job status")
                        Dim messageParameter As SqlParameter = completeCmd.Parameters.Add("@Message", SqlDbType.NVarChar, -1)
                        messageParameter.Value = $"SAP and database approval completed for all {mappedResults.Count:N0} rows."
                        Dim resultParameter As SqlParameter = completeCmd.Parameters.Add("@ResultJson", SqlDbType.NVarChar, -1)
                        resultParameter.Value = If(resultJson Is Nothing, CType(DBNull.Value, Object), resultJson)
                        If completeCmd.ExecuteNonQuery() <> 1 Then
                            Throw New Exception("The approval job status changed before commit. The entire database transaction was rolled back.")
                        End If
                    End Using

                    Using releaseCmd As New SqlCommand("DELETE FROM dbo.Draft_OTB_Approval_Claim WHERE JobID = @JobID", conn, transaction)
                        releaseCmd.CommandTimeout = 600
                        releaseCmd.Parameters.Add("@JobID", SqlDbType.UniqueIdentifier).Value = jobId
                        If releaseCmd.ExecuteNonQuery() <> mappedResults.Count Then
                            Throw New Exception("Not every approval claim could be released. The entire database transaction was rolled back.")
                        End If
                    End Using
                    transaction.Commit()
                    Return updateCount
                Catch
                    transaction.Rollback()
                    Throw
                End Try
            End Using
        End Using
    End Function

    Private Shared Function ApprovalStageText(value As String, maxLength As Integer, fieldName As String) As String
        Dim result As String = If(value, "").Trim()
        If String.IsNullOrEmpty(result) Then Throw New Exception(fieldName & " is required.")
        If result.Length > maxLength Then Throw New Exception($"{fieldName} is too long for database column ({result.Length}/{maxLength} characters).")
        Return result
    End Function

    Private Shared Function ApprovalStageNullableText(value As String, maxLength As Integer, fieldName As String) As Object
        Dim result As String = If(value, "").Trim()
        If result.Length = 0 Then Return DBNull.Value
        If result.Length > maxLength Then Throw New Exception($"{fieldName} is too long for database column ({result.Length}/{maxLength} characters).")
        Return result
    End Function

    Private Shared Function ApprovalStageTruncatedNullableText(value As String, maxLength As Integer) As Object
        Dim result As String = If(value, "").Trim()
        If result.Length = 0 Then Return DBNull.Value
        Return If(result.Length <= maxLength, result, result.Substring(0, maxLength))
    End Function

    Private Sub LegacyApproveDraftOTB(context As HttpContext)
        context.Response.ContentType = "application/json"
        Dim approvedBy As String = EnsureApprovalPermission(context, True)
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

            ' 2. Convert IDs to List(Of Integer)
            Dim runNos As List(Of Integer) = ParseRunNos(idsString)

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
                                        AddNVarCharParameter(cmdUpdate, "@OTBStatus", "Approved", 30, "OTBStatus")
                                        AddNVarCharParameter(cmdUpdate, "@ApprovedBy", approvedBy, 100, "ApprovedBy")
                                        AddNVarCharParameter(cmdUpdate, "@SAPStatus", successResult.MessageType, 10, "SAPStatus")
                                        AddNVarCharParameter(cmdUpdate, "@SAPErrorMessage", If(String.IsNullOrEmpty(successResult.Message), DBNull.Value, successResult.Message), 1000, "SAPErrorMessage")
                                        cmdUpdate.Parameters.Add("@RunNo", SqlDbType.Int).Value = runNoToUpdate
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
                                        Dim calc_YearNumber As Integer = ParseRequiredInteger(calc_Year, "Year")
                                        Dim calc_MonthNumber As Integer = ParseRequiredInteger(calc_Month, "Month")

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
                                            AddNVarCharParameter(cmdMerge, "@Type", calc_Type, 20, "Type")
                                            cmdMerge.Parameters.Add("@Year", SqlDbType.Int).Value = calc_YearNumber
                                            cmdMerge.Parameters.Add("@Month", SqlDbType.Int).Value = calc_MonthNumber
                                            AddNVarCharParameter(cmdMerge, "@Category", calc_Category, 20, "Category")
                                            AddNVarCharParameter(cmdMerge, "@Company", calc_Company, 20, "Company")
                                            AddNVarCharParameter(cmdMerge, "@Segment", calc_Segment, 20, "Segment")
                                            AddNVarCharParameter(cmdMerge, "@Brand", calc_Brand, 30, "Brand")
                                            AddNVarCharParameter(cmdMerge, "@Vendor", calc_Vendor, 30, "Vendor")
                                            AddNVarCharParameter(cmdMerge, "@Version", calc_Version, 20, "Version")

                                            ' Data Parameters (for INSERT/UPDATE)
                                            cmdMerge.Parameters.Add("@Amount", SqlDbType.Decimal).Value = calc_Amount
                                            cmdMerge.Parameters("@Amount").Precision = 18
                                            cmdMerge.Parameters("@Amount").Scale = 2
                                            cmdMerge.Parameters.Add("@RevisedDiff", SqlDbType.Decimal).Value = revisedDiffValue
                                            cmdMerge.Parameters("@RevisedDiff").Precision = 18
                                            cmdMerge.Parameters("@RevisedDiff").Scale = 2
                                            AddNVarCharParameter(cmdMerge, "@Remark", approvedRow("Remark"), 500, "Remark")
                                            AddNVarCharParameter(cmdMerge, "@ActionBy", approvedBy, 100, "ActionBy")
                                            cmdMerge.Parameters.Add("@DraftID", SqlDbType.Int).Value = runNoToUpdate
                                            AddNVarCharParameter(cmdMerge, "@SAPStatus", successResult.MessageType, 10, "SAPStatus")
                                            AddNVarCharParameter(cmdMerge, "@SAPErrorMessage", If(String.IsNullOrEmpty(successResult.Message), DBNull.Value, successResult.Message), 1000, "SAPErrorMessage")


                                            ' Parameters for UPDATE/INSERT (Names)
                                            AddNVarCharParameter(cmdMerge, "@CategoryName", approvedRow("CateName"), 200, "CategoryName")
                                            AddNVarCharParameter(cmdMerge, "@SegmentName", approvedRow("SegmentName"), 200, "SegmentName")
                                            AddNVarCharParameter(cmdMerge, "@BrandName", approvedRow("BrandName"), 200, "BrandName")
                                            AddNVarCharParameter(cmdMerge, "@VendorName", approvedRow("Vendor"), 200, "VendorName")

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
            context.Response.StatusCode = 200
            context.Response.Write(JsonConvert.SerializeObject(responseJson))
        End Try
    End Sub

    Private Function ParseRunNos(idsString As String) As List(Of Integer)
        Dim runNos As New List(Of Integer)()
        Dim seen As New HashSet(Of Integer)()
        Dim jsonArray As List(Of Object) = Nothing
        Try
            jsonArray = JsonConvert.DeserializeObject(Of List(Of Object))(idsString)
        Catch jsonEx As Exception
            jsonArray = Nothing
        End Try

        If jsonArray IsNot Nothing Then
            For Each idValue As Object In jsonArray
                AddRunNo(If(idValue, "").ToString(), runNos, seen)
            Next
        Else
            For Each idStr As String In idsString.Split(","c)
                AddRunNo(idStr.Replace("""", ""), runNos, seen)
            Next
        End If

        Return runNos
    End Function

    Private Sub AddRunNo(idText As String, runNos As List(Of Integer), seen As HashSet(Of Integer))
        Dim id As Integer
        If Not Integer.TryParse(If(idText, "").Trim(), id) OrElse id <= 0 Then
            Throw New InvalidOperationException("Every RunNo must be a positive whole number. Nothing was changed.")
        End If
        If Not seen.Contains(id) Then
            seen.Add(id)
            runNos.Add(id)
        End If
    End Sub

    Private Function ParseRequiredInteger(value As String, fieldName As String) As Integer
        Dim parsed As Integer
        If Integer.TryParse(If(value, "").Trim(), parsed) Then
            Return parsed
        End If

        Throw New Exception(fieldName & " must be a valid number.")
    End Function

    Private Sub AddNVarCharParameter(cmd As SqlCommand, parameterName As String, value As Object, size As Integer, displayName As String)
        Dim parameter As SqlParameter = cmd.Parameters.Add(parameterName, SqlDbType.NVarChar, size)
        If value Is Nothing OrElse value Is DBNull.Value Then
            parameter.Value = DBNull.Value
            Return
        End If

        Dim text As String = value.ToString()
        If text.Length > size Then
            Throw New Exception($"{displayName} is too long for database column ({text.Length}/{size} characters).")
        End If

        parameter.Value = text
    End Sub

    Private Sub MarkDownloadReady(context As HttpContext)
        Dim token As String = If(context.Request.QueryString("_downloadToken"), "")
        If String.IsNullOrWhiteSpace(token) Then Return

        Dim cookie As New HttpCookie("BMSDownloadToken", token.Trim()) With {
            .HttpOnly = False,
            .Path = "/",
            .Expires = DateTime.Now.AddMinutes(5)
        }
        context.Response.Cookies.Set(cookie)
    End Sub

    ' ===================================================================
    ' ===== END: REPLACEMENT LOGIC ======================================
    ' ===================================================================

    Private Sub DeleteDraftOTB(context As HttpContext)
        context.Response.ContentType = "application/json"
        Try
            EnsureAjaxMutationRequest(context)
            Dim currentUser As String = GetApprovalRequestUser(context)
            Dim roleName As Object = If(context.Session Is Nothing, Nothing, context.Session("UserRole"))
            Dim rights As PermissionHelper.UserRights = PermissionHelper.GetPermission(currentUser, "draftOTB.aspx", roleName)
            If Not rights.CanView OrElse Not rights.CanDelete Then
                Throw New UnauthorizedAccessException("You do not have permission to cancel Draft OTB rows.")
            End If

            Dim runNoText As String = If(context.Request.Form("runNos"), context.Request("runNos"))
            If String.IsNullOrWhiteSpace(runNoText) Then runNoText = If(context.Request.Form("runNo"), context.Request("runNo"))
            If String.IsNullOrWhiteSpace(runNoText) Then Throw New Exception("No Draft OTB rows were selected.")
            Dim runNos As List(Of Integer) = ParseRunNos(runNoText)
            If runNos.Count = 0 Then Throw New Exception("No valid RunNos were provided.")
            If runNos.Count > 15000 Then Throw New Exception($"A maximum of 15,000 Draft OTB rows can be cancelled at once. Selected: {runNos.Count:N0}.")

            Dim stage As New DataTable()
            stage.Columns.Add("RunNo", GetType(Integer))
            For Each runNo As Integer In runNos
                stage.Rows.Add(runNo)
            Next

            Dim cancelledCount As Integer = 0
            Using conn As New SqlConnection(connectionString)
                conn.Open()
                Using transaction As SqlTransaction = conn.BeginTransaction(IsolationLevel.Serializable)
                    Try
                        Using timeoutCmd As New SqlCommand("SET LOCK_TIMEOUT 30000", conn, transaction)
                            timeoutCmd.ExecuteNonQuery()
                        End Using
                        Using createCmd As New SqlCommand("CREATE TABLE #DeleteDraftRunNos (RunNo int NOT NULL PRIMARY KEY);", conn, transaction)
                            createCmd.ExecuteNonQuery()
                        End Using
                        Using bulk As New SqlBulkCopy(conn, SqlBulkCopyOptions.CheckConstraints, transaction)
                            bulk.DestinationTableName = "#DeleteDraftRunNos"
                            bulk.BatchSize = Math.Min(runNos.Count, 2000)
                            bulk.BulkCopyTimeout = 300
                            bulk.ColumnMappings.Add("RunNo", "RunNo")
                            bulk.WriteToServer(stage)
                        End Using

                        ' Global lock order: durable approval claim before Draft rows.
                        Using claimCmd As New SqlCommand("
                            SELECT TOP (1) c.RunNo
                            FROM dbo.Draft_OTB_Approval_Claim c WITH (UPDLOCK, HOLDLOCK)
                            INNER JOIN #DeleteDraftRunNos s ON s.RunNo = c.RunNo;", conn, transaction)
                            claimCmd.CommandTimeout = 300
                            Dim claimed As Object = claimCmd.ExecuteScalar()
                            If claimed IsNot Nothing AndAlso claimed IsNot DBNull.Value Then
                                Throw New InvalidOperationException("One or more selected rows are claimed by an approval or reconciliation job. Nothing was cancelled.")
                            End If
                        End Using

                        Using eligibleCmd As New SqlCommand("
                            SELECT COUNT_BIG(*)
                            FROM dbo.Template_Upload_Draft_OTB d WITH (UPDLOCK, HOLDLOCK)
                            INNER JOIN #DeleteDraftRunNos s ON s.RunNo = d.RunNo
                            WHERE ISNULL(d.OTBStatus, N'Draft') IN (N'Draft', N'Waiting', N'Edited');", conn, transaction)
                            eligibleCmd.CommandTimeout = 300
                            If Convert.ToInt64(eligibleCmd.ExecuteScalar()) <> runNos.Count Then
                                Throw New InvalidOperationException("One or more selected rows do not exist or can no longer be cancelled. Nothing was changed.")
                            End If
                        End Using

                        Using cancelCmd As New SqlCommand("
                            UPDATE d
                            SET d.OTBStatus = N'Cancelled', d.UpdateBy = @CurrentUser, d.UpdateDT = GETDATE()
                            FROM dbo.Template_Upload_Draft_OTB d
                            INNER JOIN #DeleteDraftRunNos s ON s.RunNo = d.RunNo
                            WHERE ISNULL(d.OTBStatus, N'Draft') IN (N'Draft', N'Waiting', N'Edited');", conn, transaction)
                            AddNVarCharParameter(cancelCmd, "@CurrentUser", currentUser, 100, "CurrentUser")
                            cancelCmd.CommandTimeout = 300
                            cancelledCount = cancelCmd.ExecuteNonQuery()
                        End Using
                        If cancelledCount <> runNos.Count Then
                            Throw New InvalidOperationException("Not every selected row could be cancelled. The entire change was rolled back.")
                        End If
                        transaction.Commit()
                    Catch
                        Try
                            transaction.Rollback()
                        Catch
                            ' Preserve the original SQL failure/commit acknowledgement error.
                        End Try
                        Throw
                    End Try
                End Using
            End Using
            context.Response.Write(JsonConvert.SerializeObject(New With {
                .success = True,
                .message = $"Successfully cancelled {cancelledCount:N0} Draft OTB rows.",
                .deletedCount = cancelledCount
            }))
        Catch ex As Exception
            context.Response.StatusCode = 200
            context.Response.Write(JsonConvert.SerializeObject(New With {
                .success = False,
                .message = ex.Message
            }))
        End Try
    End Sub

    Private Function GetMonthName(month As Object) As String
        If month Is Nothing OrElse month Is DBNull.Value Then Return ""

        Dim monthInt As Integer
        If Not Integer.TryParse(month.ToString(), monthInt) Then Return ""

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

        Using conn As New SqlConnection(connectionString)
            conn.Open()

            For batchStart As Integer = 0 To runNos.Count - 1 Step MaxRunNoQueryBatchSize
                Dim batchCount As Integer = Math.Min(MaxRunNoQueryBatchSize, runNos.Count - batchStart)
                Dim paramNames As New List(Of String)()
                For i As Integer = 0 To batchCount - 1
                    paramNames.Add("@p" & i.ToString())
                Next

                ' Keep each SELECT well below SQL Server's 2,100 parameter limit.
                Dim query As String = $"
                    SELECT DISTINCT
                        d.RunNo, d.Version, d.OTBCompany, d.OTBCategory, d.OTBVendor,
                        d.OTBSegment, d.OTBBrand, d.Amount, d.OTBYear,
                        d.OTBMonth, d.Remark, d.OTBType, d.OTBStatus,
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

                Using cmd As New SqlCommand(query, conn)
                    For i As Integer = 0 To batchCount - 1
                        cmd.Parameters.Add(paramNames(i), SqlDbType.Int).Value = runNos(batchStart + i)
                    Next

                    Using adapter As New SqlDataAdapter(cmd)
                        adapter.Fill(dt)
                    End Using
                End Using
            Next
        End Using

        Return dt
    End Function
    ' ===================================================================
    ' ===== END: MODIFIED FUNCTION ======================================
    ' ===================================================================

    Private Sub HandleExportOTBMovement(context As HttpContext)
        Dim year As String = NormalizeReportFilter(context.Request.QueryString("OTByear"))
        Dim month As String = NormalizeReportFilter(context.Request.QueryString("OTBmonth"))
        Dim company As String = NormalizeReportFilter(context.Request.QueryString("OTBCompany"))
        Dim category As String = NormalizeReportFilter(context.Request.QueryString("OTBCategory"))
        Dim segment As String = NormalizeReportFilter(context.Request.QueryString("OTBSegment"))
        Dim brand As String = NormalizeReportFilter(context.Request.QueryString("OTBBrand"))
        Dim vendor As String = NormalizeReportFilter(context.Request.QueryString("OTBVendor"))
        Dim masterinstance As New MasterDataUtil

        If String.IsNullOrEmpty(year) Then Throw New Exception("Year is required.")

        Dim budgetByKey As Dictionary(Of String, ReportBudgetRow) = LoadBudgetBreakdownForReport(year, month, company, category, segment, brand, vendor)
        Dim usageByKey As Dictionary(Of String, ReportUsageRow) = LoadPOUsageForReport(year, month, company, category, segment, brand, vendor)
        Dim reportKeys As List(Of String) = BuildSortedReportKeys(budgetByKey, usageByKey)

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

        For Each reportKey As String In reportKeys
            Dim budget As ReportBudgetRow = Nothing
            Dim usage As ReportUsageRow = Nothing
            Dim hasBudget As Boolean = budgetByKey.TryGetValue(reportKey, budget)
            Dim hasUsage As Boolean = usageByKey.TryGetValue(reportKey, usage)

            Dim kYear As String = If(hasBudget, budget.Year, usage.Year)
            Dim kMonth As String = If(hasBudget, budget.Month, usage.Month)
            Dim kCate As String = If(hasBudget, budget.Category, usage.Category)
            Dim kComp As String = If(hasBudget, budget.Company, usage.Company)
            Dim kSeg As String = If(hasBudget, budget.Segment, usage.Segment)
            Dim kBrand As String = If(hasBudget, budget.Brand, usage.Brand)
            Dim kVendor As String = If(hasBudget, budget.Vendor, usage.Vendor)

            Dim kCateName As String = masterinstance.GetCategoryName(kCate)
            Dim kCompName As String = masterinstance.GetCompanyName(kComp)
            Dim kSegName As String = masterinstance.GetSegmentName(kSeg)
            Dim kBrandName As String = masterinstance.GetBrandName(kBrand)
            Dim kVendorName As String = masterinstance.GetVendorName(kVendor)
            Dim kMonthName As String = GetMonthName(kMonth)

            Dim original As Decimal = If(hasBudget, budget.Original, 0D)
            Dim revDiff As Decimal = If(hasBudget, budget.RevDiff, 0D)
            Dim extra As Decimal = If(hasBudget, budget.Extra, 0D)
            Dim switchIn As Decimal = If(hasBudget, budget.SwitchIn, 0D)
            Dim balanceIn As Decimal = If(hasBudget, budget.BalanceIn, 0D)
            Dim carryIn As Decimal = If(hasBudget, budget.CarryIn, 0D)
            Dim switchOut As Decimal = If(hasBudget, budget.SignedSwitchOut, 0D)
            Dim balanceOut As Decimal = If(hasBudget, budget.SignedBalanceOut, 0D)
            Dim carryOut As Decimal = If(hasBudget, budget.SignedCarryOut, 0D)
            Dim totalBudget As Decimal = If(hasBudget, budget.Total, 0D)
            Dim sumDraft As Decimal = If(hasUsage, usage.DraftPO, 0D)
            Dim sumActual As Decimal = If(hasUsage, usage.ActualPO, 0D)
            Dim totalUsage As Decimal = sumDraft + sumActual
            Dim remaining As Decimal = totalBudget - totalUsage

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
                    original,
                    revDiff,
                    extra,
                    switchIn,
                    balanceIn,
                    carryIn,
                    switchOut,
                    balanceOut,
                    carryOut,
                    totalBudget,
                    sumActual,
                    sumDraft,
                    totalUsage,
                    remaining
                )
        Next

        GenerateExcelOTBMovement(context, dtExport, $"OTB_Movement_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx")
    End Sub

    Private Function NormalizeReportFilter(value As String) As String
        If String.IsNullOrWhiteSpace(value) Then Return Nothing
        Return value.Trim()
    End Function

    Private Function NormalizeReportKeyPart(value As String) As String
        Return If(value, "").Trim().ToUpperInvariant()
    End Function

    Private Function BuildReportKey(year As String, month As String, category As String, company As String,
                                    segment As String, brand As String, vendor As String) As String
        Return String.Join("|", New String() {
            NormalizeReportKeyPart(year),
            NormalizeReportKeyPart(month),
            NormalizeReportKeyPart(category),
            NormalizeReportKeyPart(company),
            NormalizeReportKeyPart(segment),
            NormalizeReportKeyPart(brand),
            NormalizeReportKeyPart(vendor)
        })
    End Function

    Private Function GetReportString(reader As SqlDataReader, columnName As String) As String
        Dim value As Object = reader(columnName)
        If value Is DBNull.Value Then Return ""
        Return value.ToString().Trim()
    End Function

    Private Function GetReportDecimal(reader As SqlDataReader, columnName As String) As Decimal
        Dim value As Object = reader(columnName)
        If value Is DBNull.Value Then Return 0D
        Return Convert.ToDecimal(value)
    End Function

    Private Sub AddReportFilterParameters(cmd As SqlCommand, year As String, month As String, company As String,
                                          category As String, segment As String, brand As String, vendor As String)
        AddOptionalIntParameter(cmd, "@Year", year, "Year")
        AddOptionalIntParameter(cmd, "@Month", month, "Month")
        AddOptionalNVarCharParameter(cmd, "@Company", company, 20)
        AddOptionalNVarCharParameter(cmd, "@Category", category, 20)
        AddOptionalNVarCharParameter(cmd, "@Segment", segment, 20)
        AddOptionalNVarCharParameter(cmd, "@Brand", brand, 30)
        AddOptionalNVarCharParameter(cmd, "@Vendor", vendor, 30)
    End Sub

    Private Sub AddOptionalIntParameter(cmd As SqlCommand, name As String, value As String, label As String)
        Dim parsedValue As Integer
        Dim parameter As SqlParameter = cmd.Parameters.Add(name, SqlDbType.Int)
        If String.IsNullOrWhiteSpace(value) Then
            parameter.Value = DBNull.Value
        ElseIf Integer.TryParse(value, parsedValue) Then
            parameter.Value = parsedValue
        Else
            Throw New Exception(label & " is invalid.")
        End If
    End Sub

    Private Sub AddOptionalNVarCharParameter(cmd As SqlCommand, name As String, value As String, size As Integer)
        Dim parameter As SqlParameter = cmd.Parameters.Add(name, SqlDbType.NVarChar, size)
        If String.IsNullOrWhiteSpace(value) Then
            parameter.Value = DBNull.Value
        Else
            parameter.Value = value.Trim()
        End If
    End Sub

    Private Function BuildSortedReportKeys(budgetByKey As Dictionary(Of String, ReportBudgetRow),
                                           usageByKey As Dictionary(Of String, ReportUsageRow)) As List(Of String)
        Dim allKeys As New Dictionary(Of String, Boolean)(StringComparer.OrdinalIgnoreCase)

        For Each key As String In budgetByKey.Keys
            If Not allKeys.ContainsKey(key) Then allKeys.Add(key, True)
        Next

        For Each key As String In usageByKey.Keys
            If Not allKeys.ContainsKey(key) Then allKeys.Add(key, True)
        Next

        Dim sortedKeys As New List(Of String)(allKeys.Keys)
        sortedKeys.Sort()
        Return sortedKeys
    End Function

    Private Function LoadBudgetBreakdownForReport(year As String, month As String, company As String, category As String,
                                                  segment As String, brand As String, vendor As String) As Dictionary(Of String, ReportBudgetRow)
        Dim result As New Dictionary(Of String, ReportBudgetRow)(StringComparer.OrdinalIgnoreCase)

        Dim query As String = "
            WITH BudgetRows AS (
                SELECT
                    CONVERT(nvarchar(10), [Year]) AS [Year],
                    CONVERT(nvarchar(10), [Month]) AS [Month],
                    Category,
                    Company,
                    Segment,
                    Brand,
                    Vendor,
                    CASE WHEN [Type] = 'Original' THEN ISNULL(Amount, 0) ELSE 0 END AS Original,
                    CASE WHEN [Type] = 'Revise' THEN ISNULL(RevisedDiff, 0) ELSE 0 END AS RevDiff,
                    CAST(0 AS decimal(18,2)) AS Extra,
                    CAST(0 AS decimal(18,2)) AS SwitchIn,
                    CAST(0 AS decimal(18,2)) AS BalanceIn,
                    CAST(0 AS decimal(18,2)) AS CarryIn,
                    CAST(0 AS decimal(18,2)) AS SwitchOut,
                    CAST(0 AS decimal(18,2)) AS BalanceOut,
                    CAST(0 AS decimal(18,2)) AS CarryOut
                FROM [BMS].[dbo].[OTB_Transaction]
                WHERE OTBStatus = 'Approved'
                  AND (@Year IS NULL OR [Year] = @Year)
                  AND (@Month IS NULL OR [Month] = @Month)
                  AND (@Company IS NULL OR Company = @Company)
                  AND (@Category IS NULL OR Category = @Category)
                  AND (@Segment IS NULL OR Segment = @Segment)
                  AND (@Brand IS NULL OR Brand = @Brand)
                  AND (@Vendor IS NULL OR Vendor = @Vendor)

                UNION ALL

                SELECT
                    CONVERT(nvarchar(10), [Year]) AS [Year],
                    CONVERT(nvarchar(10), [Month]) AS [Month],
                    Category,
                    Company,
                    Segment,
                    Brand,
                    Vendor,
                    CAST(0 AS decimal(18,2)) AS Original,
                    CAST(0 AS decimal(18,2)) AS RevDiff,
                    CASE WHEN [From] = 'E' THEN ISNULL(BudgetAmount, 0) ELSE 0 END AS Extra,
                    CAST(0 AS decimal(18,2)) AS SwitchIn,
                    CAST(0 AS decimal(18,2)) AS BalanceIn,
                    CAST(0 AS decimal(18,2)) AS CarryIn,
                    CASE WHEN [From] = 'D' THEN ISNULL(BudgetAmount, 0) ELSE 0 END AS SwitchOut,
                    CASE WHEN [From] = 'I' THEN ISNULL(BudgetAmount, 0) ELSE 0 END AS BalanceOut,
                    CASE WHEN [From] = 'G' THEN ISNULL(BudgetAmount, 0) ELSE 0 END AS CarryOut
                FROM [BMS].[dbo].[OTB_Switching_Transaction]
                WHERE OTBStatus = 'Approved'
                  AND (@Year IS NULL OR [Year] = @Year)
                  AND (@Month IS NULL OR [Month] = @Month)
                  AND (@Company IS NULL OR Company = @Company)
                  AND (@Category IS NULL OR Category = @Category)
                  AND (@Segment IS NULL OR Segment = @Segment)
                  AND (@Brand IS NULL OR Brand = @Brand)
                  AND (@Vendor IS NULL OR Vendor = @Vendor)

                UNION ALL

                SELECT
                    CONVERT(nvarchar(10), SwitchYear) AS [Year],
                    CONVERT(nvarchar(10), SwitchMonth) AS [Month],
                    SwitchCategory AS Category,
                    SwitchCompany AS Company,
                    SwitchSegment AS Segment,
                    SwitchBrand AS Brand,
                    SwitchVendor AS Vendor,
                    CAST(0 AS decimal(18,2)) AS Original,
                    CAST(0 AS decimal(18,2)) AS RevDiff,
                    CAST(0 AS decimal(18,2)) AS Extra,
                    CASE WHEN [To] = 'C' THEN ISNULL(BudgetAmount, 0) ELSE 0 END AS SwitchIn,
                    CASE WHEN [To] = 'H' THEN ISNULL(BudgetAmount, 0) ELSE 0 END AS BalanceIn,
                    CASE WHEN [To] = 'F' THEN ISNULL(BudgetAmount, 0) ELSE 0 END AS CarryIn,
                    CAST(0 AS decimal(18,2)) AS SwitchOut,
                    CAST(0 AS decimal(18,2)) AS BalanceOut,
                    CAST(0 AS decimal(18,2)) AS CarryOut
                FROM [BMS].[dbo].[OTB_Switching_Transaction]
                WHERE OTBStatus = 'Approved'
                  AND [To] IS NOT NULL
                  AND SwitchYear IS NOT NULL
                  AND SwitchMonth IS NOT NULL
                  AND NULLIF(LTRIM(RTRIM(SwitchCompany)), '') IS NOT NULL
                  AND NULLIF(LTRIM(RTRIM(SwitchCategory)), '') IS NOT NULL
                  AND NULLIF(LTRIM(RTRIM(SwitchSegment)), '') IS NOT NULL
                  AND NULLIF(LTRIM(RTRIM(SwitchBrand)), '') IS NOT NULL
                  AND NULLIF(LTRIM(RTRIM(SwitchVendor)), '') IS NOT NULL
                  AND (@Year IS NULL OR SwitchYear = @Year)
                  AND (@Month IS NULL OR SwitchMonth = @Month)
                  AND (@Company IS NULL OR SwitchCompany = @Company)
                  AND (@Category IS NULL OR SwitchCategory = @Category)
                  AND (@Segment IS NULL OR SwitchSegment = @Segment)
                  AND (@Brand IS NULL OR SwitchBrand = @Brand)
                  AND (@Vendor IS NULL OR SwitchVendor = @Vendor)
            )
            SELECT
                [Year],
                [Month],
                Category,
                Company,
                Segment,
                Brand,
                Vendor,
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
            GROUP BY [Year], [Month], Category, Company, Segment, Brand, Vendor"

        Using conn As New SqlConnection(connectionString)
            conn.Open()
            Using cmd As New SqlCommand(query, conn)
                AddReportFilterParameters(cmd, year, month, company, category, segment, brand, vendor)
                Using reader As SqlDataReader = cmd.ExecuteReader()
                    While reader.Read()
                        Dim item As New ReportBudgetRow With {
                            .Year = GetReportString(reader, "Year"),
                            .Month = GetReportString(reader, "Month"),
                            .Category = GetReportString(reader, "Category"),
                            .Company = GetReportString(reader, "Company"),
                            .Segment = GetReportString(reader, "Segment"),
                            .Brand = GetReportString(reader, "Brand"),
                            .Vendor = GetReportString(reader, "Vendor"),
                            .Original = GetReportDecimal(reader, "Original"),
                            .RevDiff = GetReportDecimal(reader, "RevDiff"),
                            .Extra = GetReportDecimal(reader, "Extra"),
                            .SwitchIn = GetReportDecimal(reader, "SwitchIn"),
                            .BalanceIn = GetReportDecimal(reader, "BalanceIn"),
                            .CarryIn = GetReportDecimal(reader, "CarryIn"),
                            .SwitchOut = GetReportDecimal(reader, "SwitchOut"),
                            .BalanceOut = GetReportDecimal(reader, "BalanceOut"),
                            .CarryOut = GetReportDecimal(reader, "CarryOut")
                        }
                        result(BuildReportKey(item.Year, item.Month, item.Category, item.Company, item.Segment, item.Brand, item.Vendor)) = item
                    End While
                End Using
            End Using
        End Using

        Return result
    End Function

    Private Function LoadPOUsageForReport(year As String, month As String, company As String, category As String,
                                          segment As String, brand As String, vendor As String) As Dictionary(Of String, ReportUsageRow)
        Dim result As New Dictionary(Of String, ReportUsageRow)(StringComparer.OrdinalIgnoreCase)

        Dim actualSegmentExpression As String = "COALESCE(NULLIF(a.Clean_Segment, ''), CASE WHEN LEFT(ISNULL(a.Segment_Code, ''), 1) = 'O' AND RIGHT(ISNULL(a.Segment_Code, ''), 1) = '0' AND LEN(ISNULL(a.Segment_Code, '')) > 2 THEN SUBSTRING(a.Segment_Code, 2, LEN(a.Segment_Code) - 2) ELSE ISNULL(a.Segment_Code, '') END)"
        Dim query As String = "
            WITH UsageRows AS (
                SELECT
                    CONVERT(nvarchar(10), d.PO_Year) AS [Year],
                    CONVERT(nvarchar(10), d.PO_Month) AS [Month],
                    d.Category_Code AS Category,
                    d.Company_Code AS Company,
                    d.Segment_Code AS Segment,
                    d.Brand_Code AS Brand,
                    d.Vendor_Code AS Vendor,
                    SUM(ISNULL(d.Amount_THB, 0)) AS DraftPO,
                    CAST(0 AS decimal(18,2)) AS ActualPO
                FROM [BMS].[dbo].[Draft_PO_Transaction] d
                WHERE ISNULL(d.[Status], 'Draft') NOT IN ('Matched', 'Cancelled', 'Canceled')
                  AND (@Year IS NULL OR d.PO_Year = @Year)
                  AND (@Month IS NULL OR d.PO_Month = @Month)
                  AND (@Company IS NULL OR d.Company_Code = @Company)
                  AND (@Category IS NULL OR d.Category_Code = @Category)
                  AND (@Segment IS NULL OR d.Segment_Code = @Segment)
                  AND (@Brand IS NULL OR d.Brand_Code = @Brand)
                  AND (@Vendor IS NULL OR d.Vendor_Code = @Vendor)
                GROUP BY d.PO_Year, d.PO_Month, d.Category_Code, d.Company_Code, d.Segment_Code, d.Brand_Code, d.Vendor_Code

                UNION ALL

                SELECT
                    CONVERT(nvarchar(10), a.OTB_Year) AS [Year],
                    CONVERT(nvarchar(10), a.OTB_Month) AS [Month],
                    a.Category_Code AS Category,
                    a.Company_Code AS Company,
                    " & actualSegmentExpression & " AS Segment,
                    a.Brand_Code AS Brand,
                    a.Vendor_Code AS Vendor,
                    CAST(0 AS decimal(18,2)) AS DraftPO,
                    SUM(ISNULL(a.Amount_THB, 0)) AS ActualPO
                FROM [BMS].[dbo].[Actual_PO_Summary] a
                WHERE ISNULL(a.[Status], '') = 'Matched'
                  AND a.OTB_Year IS NOT NULL
                  AND a.OTB_Month IS NOT NULL
                  AND NULLIF(LTRIM(RTRIM(a.Company_Code)), '') IS NOT NULL
                  AND NULLIF(LTRIM(RTRIM(a.Category_Code)), '') IS NOT NULL
                  AND NULLIF(LTRIM(RTRIM(" & actualSegmentExpression & ")), '') IS NOT NULL
                  AND NULLIF(LTRIM(RTRIM(a.Brand_Code)), '') IS NOT NULL
                  AND NULLIF(LTRIM(RTRIM(a.Vendor_Code)), '') IS NOT NULL
                  AND (@Year IS NULL OR a.OTB_Year = @Year)
                  AND (@Month IS NULL OR a.OTB_Month = @Month)
                  AND (@Company IS NULL OR a.Company_Code = @Company)
                  AND (@Category IS NULL OR a.Category_Code = @Category)
                  AND (@Segment IS NULL OR " & actualSegmentExpression & " = @Segment)
                  AND (@Brand IS NULL OR a.Brand_Code = @Brand)
                  AND (@Vendor IS NULL OR a.Vendor_Code = @Vendor)
                GROUP BY a.OTB_Year, a.OTB_Month, a.Category_Code, a.Company_Code, " & actualSegmentExpression & ", a.Brand_Code, a.Vendor_Code
            )
            SELECT
                [Year],
                [Month],
                Category,
                Company,
                Segment,
                Brand,
                Vendor,
                SUM(DraftPO) AS DraftPO,
                SUM(ActualPO) AS ActualPO
            FROM UsageRows
            GROUP BY [Year], [Month], Category, Company, Segment, Brand, Vendor"

        Using conn As New SqlConnection(connectionString)
            conn.Open()
            Using cmd As New SqlCommand(query, conn)
                AddReportFilterParameters(cmd, year, month, company, category, segment, brand, vendor)
                Using reader As SqlDataReader = cmd.ExecuteReader()
                    While reader.Read()
                        Dim item As New ReportUsageRow With {
                            .Year = GetReportString(reader, "Year"),
                            .Month = GetReportString(reader, "Month"),
                            .Category = GetReportString(reader, "Category"),
                            .Company = GetReportString(reader, "Company"),
                            .Segment = GetReportString(reader, "Segment"),
                            .Brand = GetReportString(reader, "Brand"),
                            .Vendor = GetReportString(reader, "Vendor"),
                            .DraftPO = GetReportDecimal(reader, "DraftPO"),
                            .ActualPO = GetReportDecimal(reader, "ActualPO")
                        }
                        result(BuildReportKey(item.Year, item.Month, item.Category, item.Company, item.Segment, item.Brand, item.Vendor)) = item
                    End While
                End Using
            End Using
        End Using

        Return result
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
            If dt.Rows.Count > 0 Then
                ws.Cells(3, 13, dt.Rows.Count + 2, 26).Style.Numberformat.Format = "#,##0.00"
            End If

            ' AutoFit
            ws.Cells.AutoFitColumns()

            ' Response
            context.Response.Clear()
            context.Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            context.Response.AddHeader("content-disposition", $"attachment; filename={filename}")
            MarkDownloadReady(context)
            context.Response.BinaryWrite(package.GetAsByteArray())
            context.Response.Flush()
            context.ApplicationInstance.CompleteRequest()
        End Using
    End Sub

    ' ==================================================================
    ' ===== NEW: SUMMARY OTB BY CATEGORY REPORT ========================
    ' ==================================================================
    Private Sub HandleExportSummaryCategory(context As HttpContext)
        Dim year As String = NormalizeReportFilter(context.Request.QueryString("OTByear"))
        Dim month As String = NormalizeReportFilter(context.Request.QueryString("OTBmonth"))
        Dim company As String = NormalizeReportFilter(context.Request.QueryString("OTBCompany"))
        Dim category As String = NormalizeReportFilter(context.Request.QueryString("OTBCategory"))
        Dim segment As String = NormalizeReportFilter(context.Request.QueryString("OTBSegment"))
        Dim masterinstance As New MasterDataUtil
        Dim brand As String = NormalizeReportFilter(context.Request.QueryString("OTBBrand"))
        Dim vendor As String = NormalizeReportFilter(context.Request.QueryString("OTBVendor"))

        If String.IsNullOrEmpty(year) Then Throw New Exception("Year is required.")

        Dim budgetByKey As Dictionary(Of String, ReportBudgetRow) = LoadBudgetBreakdownForReport(year, month, company, category, segment, brand, vendor)
        Dim usageByKey As Dictionary(Of String, ReportUsageRow) = LoadPOUsageForReport(year, month, company, category, segment, brand, vendor)
        Dim reportKeys As List(Of String) = BuildSortedReportKeys(budgetByKey, usageByKey)
        Dim summaryByKey As New Dictionary(Of String, SummaryRawItem)(StringComparer.OrdinalIgnoreCase)

        For Each reportKey As String In reportKeys
            Dim budget As ReportBudgetRow = Nothing
            Dim usage As ReportUsageRow = Nothing
            Dim hasBudget As Boolean = budgetByKey.TryGetValue(reportKey, budget)
            Dim hasUsage As Boolean = usageByKey.TryGetValue(reportKey, usage)

            Dim kYear As String = If(hasBudget, budget.Year, usage.Year)
            Dim kMonth As String = If(hasBudget, budget.Month, usage.Month)
            Dim kCate As String = If(hasBudget, budget.Category, usage.Category)
            Dim kComp As String = If(hasBudget, budget.Company, usage.Company)
            Dim kSeg As String = If(hasBudget, budget.Segment, usage.Segment)
            Dim kMonthName As String = GetMonthName(kMonth)

            Dim summaryKey As String = String.Join("|", New String() {
                NormalizeReportKeyPart(kComp),
                NormalizeReportKeyPart(kYear),
                NormalizeReportKeyPart(kMonth),
                NormalizeReportKeyPart(kCate),
                NormalizeReportKeyPart(kSeg)
            })

            If Not summaryByKey.ContainsKey(summaryKey) Then
                summaryByKey.Add(summaryKey, New SummaryRawItem With {
                    .Year = kYear,
                    .Month = kMonth,
                    .MonthName = kMonthName,
                    .Category = kCate,
                    .Company = kComp,
                    .Segment = kSeg,
                    .TotalBudget = 0D,
                    .TotalActualPO = 0D,
                    .TotalDraftPO = 0D
                })
            End If

            summaryByKey(summaryKey).TotalBudget += If(hasBudget, budget.Total, 0D)
            summaryByKey(summaryKey).TotalActualPO += If(hasUsage, usage.ActualPO, 0D)
            summaryByKey(summaryKey).TotalDraftPO += If(hasUsage, usage.DraftPO, 0D)
        Next

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
        dtExport.Columns.Add("Actual PO", GetType(Decimal))
        dtExport.Columns.Add("Draft PO", GetType(Decimal))
        dtExport.Columns.Add("Total Actual + Draft PO", GetType(Decimal))
        dtExport.Columns.Add("Remaining", GetType(Decimal))

        Dim summaryKeys As New List(Of String)(summaryByKey.Keys)
        summaryKeys.Sort()

        For Each summaryKey As String In summaryKeys
            Dim row As SummaryRawItem = summaryByKey(summaryKey)
            If row.TotalBudget <> 0 OrElse row.TotalActualDraft <> 0 Then
                dtExport.Rows.Add(
                    row.Company,
                    masterinstance.GetCompanyName(row.Company),
                    row.Year,
                    row.MonthName,
                    row.Category,
                    masterinstance.GetCategoryName(row.Category),
                    row.Segment,
                    masterinstance.GetSegmentName(row.Segment),
                    row.TotalBudget,
                    row.TotalActualPO,
                    row.TotalDraftPO,
                    row.TotalActualDraft,
                    row.TotalBudget - row.TotalActualDraft
                )
            End If
        Next

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

            ' Highlight the calculated total columns to match the movement report.
            Using rng = ws.Cells(3, 9, 3, 9) ' Total Budget Approved
                rng.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(146, 208, 80)) ' Light Green
                rng.Style.Font.Color.SetColor(Color.White)
            End Using
            Using rng = ws.Cells(3, 12, 3, 12) ' Total Actual + Draft PO
                rng.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(146, 208, 80)) ' Light Green
                rng.Style.Font.Color.SetColor(Color.White)
            End Using

            ' Format Numbers
            If dt.Rows.Count > 0 Then
                Using rng = ws.Cells(4, 9, dt.Rows.Count + 3, 13)
                    rng.Style.Numberformat.Format = "#,##0.00"
                End Using
            End If

            ws.Cells.AutoFitColumns()

            context.Response.Clear()
            context.Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            context.Response.AddHeader("content-disposition", $"attachment; filename={filename}")
            MarkDownloadReady(context)
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
