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

Public Class DataOTBHandler
    Implements IHttpHandler

    Private Shared connectionString As String = ConfigurationManager.ConnectionStrings("BMSConnectionString")?.ConnectionString

    Sub ProcessRequest(ByVal context As HttpContext) Implements IHttpHandler.ProcessRequest

        Try
            ' *** MODIFIED: Moved action check to the top ***
            Dim action As String = If(context.Request("action"), "").ToLower().Trim()

            If action = "exportdraftotb" Then
                ' *** ADDED: Handle Export Action ***
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
                dtExport.Columns.Add("TO-BE Amount (THB)", GetType(Decimal))
                dtExport.Columns.Add("Current Approved", GetType(Decimal))
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

                    Dim currentBudget As Decimal = OTBBudgetCalculator.CalculateCurrentApprovedBudget(OTBYear_Calc, OTBMonth_Calc, OTBCategory_Calc, OTBCompany_Calc, OTBSegment_Calc, OTBBrand_Calc, OTBVendor_Calc)
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
                        amountValue,
                        currentBudget,
                        diffAmount,
                        OTBStatus,
                        If(row("Version") IsNot DBNull.Value, row("Version").ToString(), ""),
                        If(row("Remark") IsNot DBNull.Value, row("Remark").ToString(), "")
                    )
                Next

                ' 4. Call Export Function
                ExportDraftToExcel(context, dtExport, "Draft_OTB_" & DateTime.Now.ToString("yyyyMMdd_HHmmss") & ".xlsx")

            Else
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
                    ' เพิ่มใน ProcessRequest method
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
                ElseIf context.Request("action") = "approveDraftOTB" Then
                    Dim runNosJson As String = If(String.IsNullOrWhiteSpace(context.Request.Form("runNos")), "[]", context.Request.Form("runNos"))
                    Dim approvedBy As String = If(String.IsNullOrWhiteSpace(context.Request.Form("approvedBy")), "unknown", context.Request.Form("approvedBy").Trim())

                    Dim serializer As New JavaScriptSerializer()
                    Dim runNosString As List(Of String) = serializer.Deserialize(Of List(Of String))(runNosJson)


                    Dim runNosInt As New List(Of Integer)()
                    For Each idStr As String In runNosString
                        Dim idInt As Integer
                        If Integer.TryParse(idStr, idInt) Then
                            runNosInt.Add(idInt)
                        End If
                    Next

                    If runNosInt.Count > 0 Then

                        Dim result As Dictionary(Of String, Object) = ApprovedOTBManager.ApproveDraftOTB(runNosInt, approvedBy)


                        If result("Status").ToString() = "Success" Then
                            context.Response.Write("Success")
                        Else
                            Dim errorMessage As String = "Unknown error during approval."
                            If result.ContainsKey("ErrorMessage") Then
                                errorMessage = result("ErrorMessage").ToString()
                            End If
                            Throw New Exception(errorMessage)
                        End If
                    Else
                        Throw New Exception("No valid items selected for approval.")
                    End If
                End If
            End If
        Catch ex As Exception
            context.Response.StatusCode = 500
            context.Response.Write($"<div class='alert alert-danger'>Error: {HttpUtility.HtmlEncode(ex.Message)}</div>")
        End Try

    End Sub



    ' *** ADDED: New function to create and send the Excel file ***
    Private Sub ExportDraftToExcel(context As HttpContext, dt As DataTable, filename As String)
        ' Set the license context for EPPlus
        ExcelPackage.License.SetNonCommercialOrganization("KingPower")


        Using package As New ExcelPackage()
            Dim worksheet = package.Workbook.Worksheets.Add("Draft OTB")
            worksheet.Cells("A1").LoadFromDataTable(dt, True)

            ' Format header
            Using range = worksheet.Cells(1, 1, 1, dt.Columns.Count)
                range.Style.Font.Bold = True
                range.Style.Fill.PatternType = ExcelFillStyle.Solid
                range.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(255, 107, 53)) ' Header Orange
                range.Style.Font.Color.SetColor(System.Drawing.Color.White)
            End Using

            ' Format number columns
            worksheet.Column(15).Style.Numberformat.Format = "#,##0.00" ' TO-BE Amount
            worksheet.Column(16).Style.Numberformat.Format = "#,##0.00" ' Current Approved
            worksheet.Column(17).Style.Numberformat.Format = "#,##0.00" ' Diff

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
    Private Function GenerateHtmlDraftTable(dt As DataTable) As String
        Dim sb As New StringBuilder()

        If dt.Rows.Count = 0 Then
            sb.Append("<tr><td colspan='20' class='text-center text-muted'>No Draft OTB records found</td></tr>")
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
                Dim currentBudget As Decimal = OTBBudgetCalculator.CalculateCurrentApprovedBudget(OTBYear, OTBMonth, OTBCategory, OTBCompany, OTBSegment, OTBBrand, OTBVendor)
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
                            HttpUtility.HtmlEncode(Amount),
                            HttpUtility.HtmlEncode(CurrentBudgetAmount),
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
            sb.Append("<tr><td colspan='21' class='text-center text-muted'>No approved OTB records found</td></tr>")
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
                sb.AppendFormat("<td class='text-center'>{0}</td>", HttpUtility.HtmlEncode(If(row("Company") IsNot DBNull.Value, row("Company").ToString(), "")))

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

                ' SAP Date
                sb.AppendFormat("<td class='date-cell'>{0}</td>",
                           If(row("SAPDate") IsNot DBNull.Value, Convert.ToDateTime(row("SAPDate")).ToString("dd/MM/yyyy HH:mm"), ""))

                ' Action By
                sb.AppendFormat("<td>{0}</td>", HttpUtility.HtmlEncode(If(row("ActionBy") IsNot DBNull.Value, row("ActionBy").ToString(), "")))
                ' SAP Status
                sb.AppendFormat("<td>{0}</td>", HttpUtility.HtmlEncode(If(row("SAPStatus") IsNot DBNull.Value, row("SAPStatus").ToString(), "")))
                ' SAP Message
                sb.AppendFormat("<td>{0}</td>", HttpUtility.HtmlEncode(If(row("SAPErrorMessage") IsNot DBNull.Value, row("SAPErrorMessage").ToString(), "")))

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

                ' Year & Month (Source) - เพิ่มการตรวจสอบ DBNull
                sb.AppendFormat("<td class='text-center'>{0}</td>",
                       If(row("Year") IsNot DBNull.Value, row("Year").ToString(), ""))
                sb.AppendFormat("<td class='text-center'>{0}</td>",
                       HttpUtility.HtmlEncode(If(row("MonthName") IsNot DBNull.Value, row("MonthName").ToString(), "")))

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

                ' Company (Source)
                sb.AppendFormat("<td class='text-center'>{0}</td>",
                       HttpUtility.HtmlEncode(If(row("Company") IsNot DBNull.Value, row("Company").ToString(), "")))
                sb.AppendFormat("<td>{0}</td>",
                       HttpUtility.HtmlEncode(If(row("CompanyName") IsNot DBNull.Value, row("CompanyName").ToString(), "")))

                ' Category (Source)
                sb.AppendFormat("<td class='text-center'>{0}</td>",
                       HttpUtility.HtmlEncode(If(row("Category") IsNot DBNull.Value, row("Category").ToString(), "")))
                sb.AppendFormat("<td>{0}</td>",
                       HttpUtility.HtmlEncode(If(row("CategoryName") IsNot DBNull.Value, row("CategoryName").ToString(), "")))

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

                ' Budget Amount
                Dim budgetAmount As Decimal = If(row("BudgetAmount") IsNot DBNull.Value, Convert.ToDecimal(row("BudgetAmount")), 0)
                sb.AppendFormat("<td class='amount-cell'>{0}</td>", budgetAmount.ToString("N2"))

                ' Release
                Dim releaseAmount As Decimal = If(row("Release") IsNot DBNull.Value, Convert.ToDecimal(row("Release")), 0)
                sb.AppendFormat("<td class='amount-cell'>{0}</td>", releaseAmount.ToString("N2"))

                ' Switch Year & Month (Target) - เพิ่มการตรวจสอบ DBNull
                sb.AppendFormat("<td class='text-center'>{0}</td>",
                       If(row("SwitchYear") IsNot DBNull.Value, row("SwitchYear").ToString(), ""))
                sb.AppendFormat("<td class='text-center'>{0}</td>",
                       HttpUtility.HtmlEncode(If(row("SwitchMonthName") IsNot DBNull.Value, row("SwitchMonthName").ToString(), "")))

                ' Switch Company (Target)
                sb.AppendFormat("<td class='text-center'>{0}</td>",
                       HttpUtility.HtmlEncode(If(row("SwitchCompany") IsNot DBNull.Value, row("SwitchCompany").ToString(), "")))

                ' Switch Category (Target)
                sb.AppendFormat("<td class='text-center'>{0}</td>",
                       HttpUtility.HtmlEncode(If(row("SwitchCategory") IsNot DBNull.Value, row("SwitchCategory").ToString(), "")))

                ' Switch Segment (Target)
                sb.AppendFormat("<td class='text-center'>{0}</td>",
                       HttpUtility.HtmlEncode(If(row("SwitchSegment") IsNot DBNull.Value, row("SwitchSegment").ToString(), "")))

                ' Switch Brand (Target)
                sb.AppendFormat("<td class='text-center'>{0}</td>",
                       HttpUtility.HtmlEncode(If(row("SwitchBrand") IsNot DBNull.Value, row("SwitchBrand").ToString(), "")))

                ' Switch Vendor (Target)
                sb.AppendFormat("<td class='text-center'>{0}</td>",
                       HttpUtility.HtmlEncode(If(row("SwitchVendor") IsNot DBNull.Value, row("SwitchVendor").ToString(), "")))

                ' Batch
                sb.AppendFormat("<td class='text-center'>{0}</td>",
                       HttpUtility.HtmlEncode(If(row("Batch") IsNot DBNull.Value, row("Batch").ToString(), "")))

                ' Remark
                sb.AppendFormat("<td class='small'>{0}</td>",
                       HttpUtility.HtmlEncode(If(row("Remark") IsNot DBNull.Value, row("Remark").ToString(), "")))

                ' Status
                sb.AppendFormat("<td class='text-center status-approved'>{0}</td>",
                       HttpUtility.HtmlEncode(If(row("OTBStatus") IsNot DBNull.Value, row("OTBStatus").ToString(), "")))

                ' Approved Date
                sb.AppendFormat("<td class='date-cell'>{0}</td>",
                       If(row("ApproveDT") IsNot DBNull.Value, Convert.ToDateTime(row("ApproveDT")).ToString("dd/MM/yyyy HH:mm"), ""))

                ' SAP Date
                sb.AppendFormat("<td class='date-cell'>{0}</td>",
                       If(row("SAPDate") IsNot DBNull.Value, Convert.ToDateTime(row("SAPDate")).ToString("dd/MM/yyyy HH:mm"), ""))

                ' Action By
                sb.AppendFormat("<td>{0}</td>",
                       HttpUtility.HtmlEncode(If(row("ActionBy") IsNot DBNull.Value, row("ActionBy").ToString(), "")))

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

    Private Sub ApproveDraftOTB(context As HttpContext)
        Try
            ' Get selected IDs
            Dim idsString As String = If(String.IsNullOrWhiteSpace(context.Request.Form("runNos")), "", context.Request.Form("runNos").Trim())
            Dim approvedBy As String = If(String.IsNullOrWhiteSpace(context.Request.Form("approvedBy")), "System", context.Request.Form("approvedBy").Trim())
            Dim remark As String = If(String.IsNullOrWhiteSpace(context.Request.Form("remark")), Nothing, context.Request.Form("remark").Trim())

            If String.IsNullOrEmpty(idsString) Then
                Dim errorResponse As New With {
                .success = False,
                .message = "No records selected for approval"
            }
                context.Response.Write(JsonConvert.SerializeObject(errorResponse))
                Return
            End If

            ' Convert comma-separated IDs to List
            Dim draftIDs As New List(Of Integer)
            For Each idStr As String In idsString.Split(","c)
                Dim id As Integer
                If Integer.TryParse(idStr.Trim(), id) Then
                    draftIDs.Add(id)
                End If
            Next

            If draftIDs.Count = 0 Then
                Dim errorResponse As New With {
                .success = False,
                .message = "Invalid draft IDs"
            }
                context.Response.Write(JsonConvert.SerializeObject(errorResponse))
                Return
            End If

            ' Call Approve function
            Dim result As Dictionary(Of String, Object) = ApprovedOTBManager.ApproveDraftOTB(draftIDs, approvedBy, remark)

            Dim response As New With {
            .success = If(result("Status").ToString() = "Success", True, False),
            .message = If(result.ContainsKey("ErrorMessage"), result("ErrorMessage").ToString(), $"Successfully approved {result("ApprovedCount")} records"),
            .approvedCount = result("ApprovedCount"),
            .sapDate = If(result.ContainsKey("SAPDate"), result("SAPDate"), Nothing)
        }

            context.Response.Write(JsonConvert.SerializeObject(response))

        Catch ex As Exception
            Dim errorResponse As New With {
            .success = False,
            .message = "Error approving records: " & ex.Message
        }
            context.Response.Write(JsonConvert.SerializeObject(errorResponse))
        End Try
    End Sub

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

            ' Convert comma-separated IDs to List
            Dim runNos As New List(Of Integer)
            Try
                ' Deserialize JSON array ["842","843","844",...]
                Dim jsonArray As List(Of String) = JsonConvert.DeserializeObject(Of List(Of String))(runNoJson)

                For Each idStr As String In jsonArray
                    Dim id As Integer
                    If Integer.TryParse(idStr.Trim(), id) Then
                        runNos.Add(id)
                    End If
                Next
            Catch jsonEx As Exception
                ' ถ้า deserialize ไม่ได้ ลองแบบเก่า (comma-separated)
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

            ' Call Approve function
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

    ReadOnly Property IsReusable() As Boolean Implements IHttpHandler.IsReusable
        Get
            Return False
        End Get
    End Property

End Class