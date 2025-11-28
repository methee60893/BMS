Imports System.Data
Imports System.Data.SqlClient
Imports System.Globalization
Imports System.Text
Imports System.Threading.Tasks
Imports System.Web
Imports System.Web.Services
Imports System.Web.SessionState
Imports Newtonsoft.Json
Imports OfficeOpenXml
Imports OfficeOpenXml.Style

Public Class DataPOHandler
    Implements System.Web.IHttpHandler, IRequiresSessionState

    Private Shared connectionString As String = ConfigurationManager.ConnectionStrings("BMSConnectionString")?.ConnectionString

    Sub ProcessRequest(ByVal context As HttpContext) Implements IHttpHandler.ProcessRequest
        Dim action As String = If(context.Request("action"), "").ToLower().Trim()

        ' --- NEW: Export Action ---
        If action = "exportdraftpo" Then
            ExportDraftPOToExcel(context)
        ElseIf action = "exportactualpo" Then
            ExportActualPOToExcel(context)

            ' --- NEW: Get List Action ---
        ElseIf action = "getactualpolist" Then
            GetActualPOList(context)
        ElseIf action = "savedraftpotxn" Then
            SaveDraftPOTXN(context)
        ElseIf action = "getdraftpolist" Then
            GetDraftPOList(context)
        ElseIf action = "getdraftpodetails" Then
            GetDraftPODetails(context)
        ElseIf action = "savedraftpoedit" Then
            SaveDraftPOEdit(context)
        ElseIf action = "deletedraftpo" Then
            DeleteDraftPO(context)
        ElseIf context.Request("action") = "PoListFilter" Then
            Try
                ' 1. รับค่า Parameters (ตัวอย่าง)
                Dim filterDate As Date = New DateTime(2025, 11, 20)
                Dim top As Integer = 1000
                Dim skip As Integer = 0

                ' 2. เรียก Helper ด้วย Workaround (Task.Run)
                ' *** ผลลัพธ์ที่ได้จะเป็น List(Of SapPOResultItem) ทันที ***
                Dim poList As List(Of SapPOResultItem) = Task.Run(Async Function()
                                                                      Return Await SapApiHelper.GetPOsAsync(filterDate)
                                                                  End Function).Result

                ' 3. ตรวจสอบผลลัพธ์
                If poList Is Nothing Then
                    Throw New Exception("Failed to get PO data from SAP.")
                End If

                ' 4. ส่งกลับเป็น JSON ให้ JavaScript (แนะนำวิธีนี้)
                context.Response.ContentType = "application/json"
                Dim successResponse = New With {
            .success = True,
            .count = poList.Count,
            .data = poList ' ส่ง List ทั้งหมดไปให้ JavaScript
        }
                context.Response.Write(JsonConvert.SerializeObject(successResponse))

            Catch ex As Exception
                context.Response.ContentType = "application/json"
                context.Response.StatusCode = 500
                Dim errorResponse As New With {
            .success = False,
            .message = ex.Message
        }
                context.Response.Write(JsonConvert.SerializeObject(errorResponse))
            End Try
        Else
            context.Response.Clear()
            context.Response.ContentType = "text/html"
            context.Response.ContentEncoding = Encoding.UTF8
            Dim dt As DataTable = Nothing

            Dim draftPONo As String = If(String.IsNullOrWhiteSpace(context.Request.QueryString("draftPONo")), "", context.Request.QueryString("draftPONo").Trim())
            Dim draftYear As String = If(String.IsNullOrWhiteSpace(context.Request.QueryString("draftYear")), "", context.Request.QueryString("draftYear").Trim())
            Dim draftMonth As String = If(String.IsNullOrWhiteSpace(context.Request.QueryString("draftMonth")), "", context.Request.QueryString("draftMonth").Trim())
            Dim draftCategory As String = If(String.IsNullOrWhiteSpace(context.Request.QueryString("draftCategory")), "", context.Request.QueryString("draftCategory").Trim())
            Dim draftCompany As String = If(String.IsNullOrWhiteSpace(context.Request.QueryString("draftCompany")), "", context.Request.QueryString("draftCompany").Trim())
            Dim draftBrand As String = If(String.IsNullOrWhiteSpace(context.Request.QueryString("draftBrand")), "", context.Request.QueryString("draftBrand").Trim())
            Dim draftVendor As String = If(String.IsNullOrWhiteSpace(context.Request.QueryString("draftVendor")), "", context.Request.QueryString("draftVendor").Trim())
            Dim draftAmountTHB As String = If(String.IsNullOrWhiteSpace(context.Request.QueryString("draftAmountTHB")), "", context.Request.QueryString("draftAmountTHB").Trim())
            Dim draftAmountCCY As String = If(String.IsNullOrWhiteSpace(context.Request.QueryString("draftAmountCCY")), "", context.Request.QueryString("draftAmountCCY").Trim())
            Dim draftCCY As String = If(String.IsNullOrWhiteSpace(context.Request.QueryString("draftCCY")), "", context.Request.QueryString("draftCCY").Trim())
            Dim draftExhangeRate As String = If(String.IsNullOrWhiteSpace(context.Request.QueryString("draftExhangeRate")), "", context.Request.QueryString("draftExhangeRate").Trim())
            Dim draftRemark As String = If(String.IsNullOrWhiteSpace(context.Request.QueryString("draftRemark")), "", context.Request.QueryString("draftRemark").Trim())
        End If
    End Sub

    ' =================================================
    ' ===== ADDED: DELETE (CANCEL) DRAFT PO FUNCTION =====
    ' =================================================
    Private Sub DeleteDraftPO(context As HttpContext)
        context.Response.ContentType = "application/json"
        Dim response As New Dictionary(Of String, Object)
        Dim statusBy As String = "System"

        If context.Session IsNot Nothing AndAlso context.Session("user") IsNot Nothing Then
            statusBy = context.Session("user").ToString()
        End If

        Try
            Dim draftPOID As Integer
            If Not Integer.TryParse(context.Request.Form("draftPOID"), draftPOID) Then
                Throw New Exception("Invalid DraftPO ID")
            End If

            Using conn As New SqlConnection(connectionString)
                conn.Open()

                ' ตรวจสอบสถานะก่อนลบ (Double Check)
                Dim checkQuery As String = "SELECT Status FROM [BMS].[dbo].[Draft_PO_Transaction] WHERE DraftPO_ID = @ID"
                Dim currentStatus As String = ""
                Using cmdCheck As New SqlCommand(checkQuery, conn)
                    cmdCheck.Parameters.AddWithValue("@ID", draftPOID)
                    Dim res = cmdCheck.ExecuteScalar()
                    If res IsNot Nothing Then currentStatus = res.ToString()
                End Using

                If currentStatus.Equals("Matched", StringComparison.OrdinalIgnoreCase) OrElse
                   currentStatus.Equals("Cancelled", StringComparison.OrdinalIgnoreCase) Then
                    Throw New Exception($"Cannot cancel. Current status is '{currentStatus}'.")
                End If

                ' Update Status to Cancelled
                Dim query As String = "UPDATE [BMS].[dbo].[Draft_PO_Transaction] " &
                                      "SET [Status] = 'Cancelled', " &
                                      "    [Status_Date] = GETDATE(), " &
                                      "    [Status_By] = @StatusBy " &
                                      "WHERE [DraftPO_ID] = @DraftPOID"

                Using cmd As New SqlCommand(query, conn)
                    cmd.Parameters.AddWithValue("@DraftPOID", draftPOID)
                    cmd.Parameters.AddWithValue("@StatusBy", statusBy)
                    cmd.ExecuteNonQuery()
                End Using
            End Using

            response("success") = True
            response("message") = "Draft PO cancelled successfully."

        Catch ex As Exception
            response("success") = False
            response("message") = ex.Message
        End Try

        context.Response.Write(JsonConvert.SerializeObject(response))
    End Sub
    Private Function GetDraftPODataTable(context As HttpContext, isExport As Boolean) As DataTable
        ' รับค่า Parameter (รองรับทั้ง Form สำหรับ View และ QueryString สำหรับ Export)
        Dim req = context.Request
        Dim year As String = If(isExport, req.QueryString("year"), req.Form("year"))
        Dim month As String = If(isExport, req.QueryString("month"), req.Form("month"))
        Dim company As String = If(isExport, req.QueryString("company"), req.Form("company"))
        Dim category As String = If(isExport, req.QueryString("category"), req.Form("category"))
        Dim segment As String = If(isExport, req.QueryString("segment"), req.Form("segment"))
        Dim brand As String = If(isExport, req.QueryString("brand"), req.Form("brand"))
        Dim vendor As String = If(isExport, req.QueryString("vendor"), req.Form("vendor"))

        Dim dt As New DataTable()
        Using conn As New SqlConnection(connectionString)
            conn.Open()
            Dim query As New StringBuilder()
            query.Append("SELECT DISTINCT  ")
            query.Append("   po.DraftPO_ID, ")
            query.Append("   po.Created_Date, ")
            query.Append("   po.DraftPO_No, ")
            query.Append("   po.PO_Type, ")
            query.Append("   po.PO_Year, ")
            query.Append("   m.month_name_sh AS PO_Month_Name, ")
            query.Append("   po.Category_Code, ")
            query.Append("   c.Category AS Category_Name, ")
            query.Append("   po.Company_Code, ")
            query.Append("   comp.[CompanyNameShort] As Company_Name, ")
            query.Append("   po.Segment_Code, ")
            query.Append("   s.SegmentName AS Segment_Name, ")
            query.Append("   po.Brand_Code, ")
            query.Append("   b.[Brand Name] AS Brand_Name, ")
            query.Append("   po.Vendor_Code, ")
            query.Append("   v.Vendor AS Vendor_Name, ")
            query.Append("   po.Amount_THB, ")
            query.Append("   po.Amount_CCY, ")
            query.Append("   po.CCY, ")
            query.Append("   po.Exchange_Rate, ")
            query.Append("   po.Actual_PO_Ref, ")
            query.Append("   po.Status, ")
            query.Append("   po.Status_Date, ")
            query.Append("   po.Remark, ")
            query.Append("   po.Status_By ")
            query.Append("FROM [BMS].[dbo].[Draft_PO_Transaction] po ")
            query.Append("LEFT JOIN [BMS].[dbo].[MS_Month] m ON po.PO_Month = m.month_code ")
            query.Append("LEFT JOIN [BMS].[dbo].[MS_Company] comp ON po.Company_Code = comp.[CompanyCode]  ")
            query.Append("LEFT JOIN [BMS].[dbo].[MS_Category] c ON po.Category_Code = c.Cate ")
            query.Append("LEFT JOIN [BMS].[dbo].[MS_Segment] s ON po.Segment_Code = s.SegmentCode ")
            query.Append("LEFT JOIN [BMS].[dbo].[MS_Brand] b ON po.Brand_Code = b.[Brand Code] ")
            query.Append("LEFT JOIN [BMS].[dbo].[MS_Vendor] v ON po.Vendor_Code = v.VendorCode AND po.Segment_Code = v.SegmentCode ")
            query.Append("WHERE 1=1 ")

            Using cmd As New SqlCommand()
                If Not String.IsNullOrEmpty(year) Then
                    query.Append("AND po.PO_Year = @Year ")
                    cmd.Parameters.AddWithValue("@Year", year)
                End If
                If Not String.IsNullOrEmpty(month) Then
                    query.Append("AND po.PO_Month = @Month ")
                    cmd.Parameters.AddWithValue("@Month", month)
                End If
                If Not String.IsNullOrEmpty(company) Then
                    query.Append("AND po.Company_Code = @Company ")
                    cmd.Parameters.AddWithValue("@Company", company)
                End If
                If Not String.IsNullOrEmpty(category) Then
                    query.Append("AND po.Category_Code = @Category ")
                    cmd.Parameters.AddWithValue("@Category", category)
                End If
                If Not String.IsNullOrEmpty(segment) Then
                    query.Append("AND po.Segment_Code = @Segment ")
                    cmd.Parameters.AddWithValue("@Segment", segment)
                End If
                If Not String.IsNullOrEmpty(brand) Then
                    query.Append("AND po.Brand_Code = @Brand ")
                    cmd.Parameters.AddWithValue("@Brand", brand)
                End If
                If Not String.IsNullOrEmpty(vendor) Then
                    query.Append("AND po.Vendor_Code = @Vendor ")
                    cmd.Parameters.AddWithValue("@Vendor", vendor)
                End If

                query.Append("ORDER BY po.Created_Date DESC ")

                cmd.CommandText = query.ToString()
                cmd.Connection = conn

                Using reader As SqlDataReader = cmd.ExecuteReader()
                    dt.Load(reader)
                End Using
            End Using
        End Using
        Return dt
    End Function

    ' =================================================
    ' ===== NEW: GetActualPOList FUNCTION =====
    ' =================================================
    Private Sub GetActualPOList(context As HttpContext)
        Try
            Dim year As String = If(String.IsNullOrWhiteSpace(context.Request.Form("year")), Nothing, context.Request.Form("year").Trim())
            Dim month As String = If(String.IsNullOrWhiteSpace(context.Request.Form("month")), Nothing, context.Request.Form("month").Trim())
            Dim company As String = If(String.IsNullOrWhiteSpace(context.Request.Form("company")), Nothing, context.Request.Form("company").Trim())
            Dim category As String = If(String.IsNullOrWhiteSpace(context.Request.Form("category")), Nothing, context.Request.Form("category").Trim())
            Dim segment As String = If(String.IsNullOrWhiteSpace(context.Request.Form("segment")), Nothing, context.Request.Form("segment").Trim())
            Dim brand As String = If(String.IsNullOrWhiteSpace(context.Request.Form("brand")), Nothing, context.Request.Form("brand").Trim())
            Dim vendor As String = If(String.IsNullOrWhiteSpace(context.Request.Form("vendor")), Nothing, context.Request.Form("vendor").Trim())

            Dim dt As DataTable = GetActualPOData(year, month, company, category, segment, brand, vendor)

            ' Generate HTML table body
            Dim html As String = GenerateHtmlActualPOTable(dt)

            context.Response.ContentType = "text/html"
            context.Response.Write(html)

        Catch ex As Exception
            context.Response.StatusCode = 500
            context.Response.Write($"<tr><td colspan='22' class='text-center text-danger'>Error: {HttpUtility.HtmlEncode(ex.Message)}</td></tr>")
        End Try
    End Sub

    ' =================================================
    ' ===== NEW: GetActualPOData (Helper Function) =====
    ' =================================================
    Private Function GetActualPOData(year As String, month As String, company As String, category As String, segment As String, brand As String, vendor As String) As DataTable
        Dim dt As New DataTable()
        Using conn As New SqlConnection(connectionString)
            conn.Open()
            Using cmd As New SqlCommand("SP_Get_Actual_PO_List", conn)
                cmd.CommandType = CommandType.StoredProcedure
                cmd.Parameters.AddWithValue("@Year", If(year, DBNull.Value))
                cmd.Parameters.AddWithValue("@Month", If(month, DBNull.Value))
                cmd.Parameters.AddWithValue("@Company", If(company, DBNull.Value))
                cmd.Parameters.AddWithValue("@Category", If(category, DBNull.Value))
                cmd.Parameters.AddWithValue("@Segment", If(segment, DBNull.Value))
                cmd.Parameters.AddWithValue("@Brand", If(brand, DBNull.Value))
                cmd.Parameters.AddWithValue("@Vendor", If(vendor, DBNull.Value))

                Using reader As SqlDataReader = cmd.ExecuteReader()
                    dt.Load(reader)
                End Using
            End Using
        End Using
        Return dt
    End Function

    ' =================================================
    ' ===== NEW: GenerateHtmlActualPOTable FUNCTION =====
    ' =================================================
    Private Function GenerateHtmlActualPOTable(dt As DataTable) As String
        Dim sb As New StringBuilder()
        Dim masterinstance As New MasterDataUtil
        If dt.Rows.Count = 0 Then
            sb.Append("<tr><td colspan='22' class='text-center text-muted p-4'>No Actual PO records found for the selected filters.</td></tr>")
            Return sb.ToString()
        End If

        For Each row As DataRow In dt.Rows
            Dim status As String = GetDbString(row, "Status")
            Dim statusClass As String = ""
            If status.ToLower() = "cancelled" Then
                statusClass = "class='status-cancelled'"
            End If

            sb.Append("<tr>")
            sb.AppendFormat("<td>{0}</td>", GetDbDate(row, "Actual_PO_Date"))
            sb.AppendFormat("<td>{0}</td>", HttpUtility.HtmlEncode(GetDbString(row, "Actual_PO_No")))
            sb.AppendFormat("<td>{0}</td>", HttpUtility.HtmlEncode(GetDbString(row, "PO_Type")))
            sb.AppendFormat("<td>{0}</td>", HttpUtility.HtmlEncode(GetDbString(row, "PO_Year")))
            sb.AppendFormat("<td>{0}</td>", HttpUtility.HtmlEncode(GetDbString(row, "PO_Month_Name"))) ' Use Month Name
            sb.AppendFormat("<td>{0}</td>", HttpUtility.HtmlEncode(GetDbString(row, "Category_Code")))
            sb.AppendFormat("<td>{0}</td>", HttpUtility.HtmlEncode(GetDbString(row, "Category_Name")))
            sb.AppendFormat("<td>{0}</td>", HttpUtility.HtmlEncode(masterinstance.GetCompanyName(GetDbString(row, "Company_Code"))))
            sb.AppendFormat("<td>{0}</td>", HttpUtility.HtmlEncode(GetDbString(row, "Segment_Code")))
            sb.AppendFormat("<td>{0}</td>", HttpUtility.HtmlEncode(GetDbString(row, "Segment_Name")))
            sb.AppendFormat("<td>{0}</td>", HttpUtility.HtmlEncode(GetDbString(row, "Brand_Code")))
            sb.AppendFormat("<td>{0}</td>", HttpUtility.HtmlEncode(GetDbString(row, "Brand_Name")))
            sb.AppendFormat("<td>{0}</td>", HttpUtility.HtmlEncode(GetDbString(row, "Vendor_Code")))
            sb.AppendFormat("<td>{0}</td>", HttpUtility.HtmlEncode(GetDbString(row, "Vendor_Name")))
            sb.AppendFormat("<td class='text-end'>{0}</td>", GetDbDecimal(row, "Amount_THB").ToString("N2"))
            sb.AppendFormat("<td class='text-end'>{0}</td>", GetDbDecimal(row, "Amount_CCY").ToString("N2"))
            sb.AppendFormat("<td>{0}</td>", HttpUtility.HtmlEncode(GetDbString(row, "CCY")))
            sb.AppendFormat("<td class='text-end'>{0}</td>", GetDbDecimal(row, "Exchange_Rate").ToString("N4")) ' Spec shows 4 decimals
            sb.AppendFormat("<td>{0}</td>", HttpUtility.HtmlEncode(GetDbString(row, "Draft_PO_Ref")))
            sb.AppendFormat("<td>{0}</td>", HttpUtility.HtmlEncode(GetDbString(row, "Remark")))
            sb.AppendFormat("<td {0}>{1}</td>", statusClass, HttpUtility.HtmlEncode(status))
            sb.AppendFormat("<td>{0}</td>", GetDbDate(row, "Status_Date"))
            sb.Append("</tr>")
        Next

        Return sb.ToString()
    End Function

    ' =================================================
    ' ===== NEW: ExportActualPOToExcel FUNCTION =====
    ' =================================================
    Private Sub ExportActualPOToExcel(context As HttpContext)
        ' 1. Get Filters from QueryString (Export is GET request)
        Dim year As String = If(String.IsNullOrWhiteSpace(context.Request.QueryString("year")), Nothing, context.Request.QueryString("year").Trim())
        Dim month As String = If(String.IsNullOrWhiteSpace(context.Request.QueryString("month")), Nothing, context.Request.QueryString("month").Trim())
        Dim company As String = If(String.IsNullOrWhiteSpace(context.Request.QueryString("company")), Nothing, context.Request.QueryString("company").Trim())
        Dim category As String = If(String.IsNullOrWhiteSpace(context.Request.QueryString("category")), Nothing, context.Request.QueryString("category").Trim())
        Dim segment As String = If(String.IsNullOrWhiteSpace(context.Request.QueryString("segment")), Nothing, context.Request.QueryString("segment").Trim())
        Dim brand As String = If(String.IsNullOrWhiteSpace(context.Request.QueryString("brand")), Nothing, context.Request.QueryString("brand").Trim())
        Dim vendor As String = If(String.IsNullOrWhiteSpace(context.Request.QueryString("vendor")), Nothing, context.Request.QueryString("vendor").Trim())

        ' 2. Get Raw Data
        Dim dtRaw As DataTable = GetActualPOData(year, month, company, category, segment, brand, vendor)

        ' 3. Create Export-formatted DataTable (to match the spec columns)
        Dim dtExport As New DataTable("ActualPO")
        dtExport.Columns.Add("Actual PO Date", GetType(String))
        dtExport.Columns.Add("Actual PO no.", GetType(String))
        dtExport.Columns.Add("Type", GetType(String))
        dtExport.Columns.Add("Year", GetType(String))
        dtExport.Columns.Add("Month", GetType(String))
        dtExport.Columns.Add("Category", GetType(String))
        dtExport.Columns.Add("Category name", GetType(String))
        dtExport.Columns.Add("Company", GetType(String))
        dtExport.Columns.Add("Segment", GetType(String))
        dtExport.Columns.Add("Segment name", GetType(String))
        dtExport.Columns.Add("Brand", GetType(String))
        dtExport.Columns.Add("Brand name", GetType(String))
        dtExport.Columns.Add("Vendor", GetType(String))
        dtExport.Columns.Add("Vendor name", GetType(String))
        dtExport.Columns.Add("Amount (THB)", GetType(Decimal))
        dtExport.Columns.Add("Amount (CCY)", GetType(Decimal))
        dtExport.Columns.Add("CCY", GetType(String))
        dtExport.Columns.Add("Ex. Rate", GetType(Decimal))
        dtExport.Columns.Add("Draft PO Ref", GetType(String))
        dtExport.Columns.Add("Status", GetType(String))
        dtExport.Columns.Add("Status date", GetType(String))
        dtExport.Columns.Add("Remark", GetType(String))

        ' 4. Populate dtExport
        For Each row As DataRow In dtRaw.Rows
            dtExport.Rows.Add(
                GetDbDate(row, "Actual_PO_Date"),
                GetDbString(row, "Actual_PO_No"),
                GetDbString(row, "PO_Type"),
                GetDbString(row, "PO_Year"),
                GetDbString(row, "PO_Month_Name"),
                GetDbString(row, "Category_Code"),
                GetDbString(row, "Category_Name"),
                GetDbString(row, "Company_Name"),
                GetDbString(row, "Segment_Code"),
                GetDbString(row, "Segment_Name"),
                GetDbString(row, "Brand_Code"),
                GetDbString(row, "Brand_Name"),
                GetDbString(row, "Vendor_Code"),
                GetDbString(row, "Vendor_Name"),
                GetDbDecimal(row, "Amount_THB"),
                GetDbDecimal(row, "Amount_CCY"),
                GetDbString(row, "CCY"),
                GetDbDecimal(row, "Exchange_Rate"),
                GetDbString(row, "Draft_PO_Ref"),
                GetDbString(row, "Status"),
                GetDbDate(row, "Status_Date"),
                GetDbString(row, "Remark")
            )
        Next

        ' 5. Call Export Function
        ExportToExcel(context, dtExport, "Actual_PO_" & DateTime.Now.ToString("yyyyMMdd_HHmmss") & ".xlsx")
    End Sub

    ' =================================================
    ' ===== NEW: ExportToExcel (Helper Function) =====
    ' =================================================
    Private Sub ExportToExcel(context As HttpContext, dt As DataTable, filename As String)
        ExcelPackage.License.SetNonCommercialOrganization("KingPower")

        Using package As New ExcelPackage()
            Dim worksheet = package.Workbook.Worksheets.Add("ActualPO")
            worksheet.Cells("A1").LoadFromDataTable(dt, True)

            ' Format header
            Using range = worksheet.Cells(1, 1, 1, dt.Columns.Count)
                range.Style.Font.Bold = True
                range.Style.Fill.PatternType = ExcelFillStyle.Solid
                range.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(11, 86, 164)) ' Primary Blue
                range.Style.Font.Color.SetColor(System.Drawing.Color.White)
            End Using

            ' Format number columns (based on new dtExport)
            worksheet.Column(15).Style.Numberformat.Format = "#,##0.00" ' Amount (THB)
            worksheet.Column(16).Style.Numberformat.Format = "#,##0.00" ' Amount (CCY)
            worksheet.Column(18).Style.Numberformat.Format = "#,##0.0000" ' Ex. Rate (4 decimals)

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

    Private Function GetDbDate(row As DataRow, columnName As String, Optional format As String = "dd/MM/yyyy HH:mm") As String
        If row IsNot Nothing AndAlso row.Table.Columns.Contains(columnName) AndAlso row(columnName) IsNot DBNull.Value Then
            Try
                Return Convert.ToDateTime(row(columnName)).ToString(format)
            Catch ex As Exception
                Return row(columnName).ToString() ' Return as string if conversion fails
            End Try
        End If
        Return ""
    End Function

    ' =================================================
    ' ===== ADDED: GetDraftPOList FUNCTION =====
    ' =================================================
    Private Sub GetDraftPOList(context As HttpContext)
        Try
            ' เรียกใช้ฟังก์ชันกลาง (isExport = False เพราะรับค่าจาก Form)
            Dim dt As DataTable = GetDraftPODataTable(context, False)
            Dim jsonResult As String = JsonConvert.SerializeObject(dt, Formatting.None)
            context.Response.Write(jsonResult)
        Catch ex As Exception
            context.Response.StatusCode = 500
            context.Response.Write(JsonConvert.SerializeObject(New With {.success = False, .message = ex.Message}))
        End Try
    End Sub

    Private Sub ExportDraftPOToExcel(context As HttpContext)
        Try
            ' 1. ดึงข้อมูลโดยใช้ Logic เดียวกับ View (isExport = True รับค่าจาก QueryString)
            Dim dtRaw As DataTable = GetDraftPODataTable(context, True)

            ' 2. สร้าง DataTable สำหรับ Export (จัด Format Column)
            Dim dtExport As New DataTable("DraftPO")
            dtExport.Columns.Add("Draft PO Date", GetType(String))
            dtExport.Columns.Add("Draft PO no.", GetType(String))
            dtExport.Columns.Add("Type", GetType(String))
            dtExport.Columns.Add("Year", GetType(String))
            dtExport.Columns.Add("Month", GetType(String))
            dtExport.Columns.Add("Category", GetType(String))
            dtExport.Columns.Add("Category name", GetType(String))
            dtExport.Columns.Add("Company", GetType(String))
            dtExport.Columns.Add("Segment", GetType(String))
            dtExport.Columns.Add("Segment name", GetType(String))
            dtExport.Columns.Add("Brand", GetType(String))
            dtExport.Columns.Add("Brand name", GetType(String))
            dtExport.Columns.Add("Vendor", GetType(String))
            dtExport.Columns.Add("Vendor name", GetType(String))
            dtExport.Columns.Add("Amount (THB)", GetType(Decimal))
            dtExport.Columns.Add("Amount (CCY)", GetType(Decimal))
            dtExport.Columns.Add("CCY", GetType(String))
            dtExport.Columns.Add("Ex. Rate", GetType(Decimal))
            dtExport.Columns.Add("Actual PO Ref", GetType(String))
            dtExport.Columns.Add("Status", GetType(String))
            dtExport.Columns.Add("Status date", GetType(String))
            dtExport.Columns.Add("Remark", GetType(String))
            dtExport.Columns.Add("Action by", GetType(String))

            For Each row As DataRow In dtRaw.Rows
                dtExport.Rows.Add(
                    GetDbDate(row, "Created_Date"),
                    GetDbString(row, "DraftPO_No"),
                    GetDbString(row, "PO_Type"),
                    GetDbString(row, "PO_Year"),
                    GetDbString(row, "PO_Month_Name"),
                    GetDbString(row, "Category_Code"),
                    GetDbString(row, "Category_Name"),
                    GetDbString(row, "Company_Name"),
                    GetDbString(row, "Segment_Code"),
                    GetDbString(row, "Segment_Name"),
                    GetDbString(row, "Brand_Code"),
                    GetDbString(row, "Brand_Name"),
                    GetDbString(row, "Vendor_Code"),
                    GetDbString(row, "Vendor_Name"),
                    GetDbDecimal(row, "Amount_THB"),
                    GetDbDecimal(row, "Amount_CCY"),
                    GetDbString(row, "CCY"),
                    GetDbDecimal(row, "Exchange_Rate"),
                    GetDbString(row, "Actual_PO_Ref"),
                    GetDbString(row, "Status"),
                    GetDbDate(row, "Status_Date"),
                    GetDbString(row, "Remark"),
                    GetDbString(row, "Status_By")
                )
            Next

            ' 3. ใช้ฟังก์ชัน ExportToExcel (ที่มีอยู่แล้ว หรือสร้างใหม่ถ้ายังไม่มีในไฟล์นี้)
            ExportToExcel(context, dtExport, "Draft_PO_" & DateTime.Now.ToString("yyyyMMdd_HHmmss") & ".xlsx")

        Catch ex As Exception
            context.Response.Write("Error exporting data: " & ex.Message)
        End Try
    End Sub

    ' =================================================
    ' ===== ADDED: GetDraftPODetails FUNCTION =====
    ' =================================================
    Private Sub GetDraftPODetails(context As HttpContext)
        Try
            Dim draftPOID As Integer = 0
            Integer.TryParse(context.Request.Form("draftPOID"), draftPOID)

            If draftPOID = 0 Then
                Throw New Exception("Invalid DraftPO_ID.")
            End If

            Dim dt As New DataTable()
            Using conn As New SqlConnection(connectionString)
                conn.Open()
                ' (ไม่ต้อง Join master เพราะเราจะใช้ ID ไป set dropdown)
                Dim query As String = "SELECT * FROM [BMS].[dbo].[Draft_PO_Transaction] WHERE DraftPO_ID = @DraftPOID"
                Using cmd As New SqlCommand(query, conn)
                    cmd.Parameters.AddWithValue("@DraftPOID", draftPOID)
                    Using reader As SqlDataReader = cmd.ExecuteReader()
                        dt.Load(reader)
                    End Using
                End Using
            End Using

            If dt.Rows.Count > 0 Then
                ' --- MODIFIED: Convert DataRow to Dictionary for clean JSON ---
                Dim row As DataRow = dt.Rows(0)
                Dim dict As New Dictionary(Of String, Object)()
                For Each col As DataColumn In row.Table.Columns
                    dict(col.ColumnName) = If(row(col) Is DBNull.Value, Nothing, row(col))
                Next

                Dim jsonResult As String = JsonConvert.SerializeObject(dict, Formatting.None)
                context.Response.Write(jsonResult)
            Else
                Throw New Exception("Draft PO record not found.")
            End If

        Catch ex As Exception
            context.Response.StatusCode = 500
            context.Response.Write(JsonConvert.SerializeObject(New With {
                .success = False,
                .message = "Error fetching details: " & ex.Message
            }))
        End Try
    End Sub


    ' =================================================
    ' ===== ADDED: SaveDraftPOEdit FUNCTION =====
    ' =================================================
    Private Sub SaveDraftPOEdit(context As HttpContext)
        context.Response.ContentType = "application/json"
        Dim response As New Dictionary(Of String, Object)
        Dim statusBy As String = "System" ' Default
        Try
            ' (ต้องตรวจสอบ Session จริง)
            If context.Session IsNot Nothing AndAlso context.Session("user") IsNot Nothing AndAlso Not String.IsNullOrEmpty(context.Session("user").ToString()) Then
                statusBy = context.Session("user").ToString()
            End If

            ' ดึงข้อมูลจาก Form
            Dim draftPOID As Integer = Integer.Parse(context.Request.Form("draftPOID"))
            Dim year As String = context.Request.Form("year")
            Dim month As String = context.Request.Form("month")
            Dim company As String = context.Request.Form("company")
            Dim category As String = context.Request.Form("category")
            Dim segment As String = context.Request.Form("segment")
            Dim brand As String = context.Request.Form("brand")
            Dim vendor As String = context.Request.Form("vendor")
            Dim amtCCY As Decimal = Decimal.Parse(context.Request.Form("amtCCY"), CultureInfo.InvariantCulture)
            Dim ccy As String = context.Request.Form("ccy")
            Dim exRate As Decimal = Decimal.Parse(context.Request.Form("exRate"), CultureInfo.InvariantCulture)
            Dim remark As String = context.Request.Form("remark")
            Dim amtTHB As Decimal = amtCCY * exRate

            Using conn As New SqlConnection(connectionString)
                conn.Open()
                Dim query As String = "UPDATE [BMS].[dbo].[Draft_PO_Transaction] 
                                       SET [PO_Year] = @year, 
                                           [PO_Month] = @month, 
                                           [Company_Code] = @company, 
                                           [Category_Code] = @category, 
                                           [Segment_Code] = @segment, 
                                           [Brand_Code] = @brand, 
                                           [Vendor_Code] = @vendor, 
                                           [CCY] = @ccy, 
                                           [Exchange_Rate] = @exRate, 
                                           [Amount_CCY] = @amtCCY, 
                                           [Amount_THB] = @amtTHB, 
                                           [Status] = 'Edited', 
                                           [Status_Date] = GETDATE(), 
                                           [Status_By] = @statusBy, 
                                           [Remark] = @remark
                                       WHERE [DraftPO_ID] = @draftPOID"

                Using cmd As New SqlCommand(query, conn)
                    cmd.Parameters.AddWithValue("@draftPOID", draftPOID)
                    cmd.Parameters.AddWithValue("@year", year)
                    cmd.Parameters.AddWithValue("@month", month)
                    cmd.Parameters.AddWithValue("@company", company)
                    cmd.Parameters.AddWithValue("@category", category)
                    cmd.Parameters.AddWithValue("@segment", segment)
                    cmd.Parameters.AddWithValue("@brand", brand)
                    cmd.Parameters.AddWithValue("@vendor", vendor)
                    cmd.Parameters.AddWithValue("@amtTHB", amtTHB)
                    cmd.Parameters.AddWithValue("@amtCCY", amtCCY)
                    cmd.Parameters.AddWithValue("@ccy", ccy)
                    cmd.Parameters.AddWithValue("@exRate", exRate)
                    cmd.Parameters.AddWithValue("@statusBy", statusBy)
                    cmd.Parameters.AddWithValue("@remark", If(String.IsNullOrEmpty(remark), DBNull.Value, remark))

                    Dim rowsAffected As Integer = cmd.ExecuteNonQuery()
                    If rowsAffected > 0 Then
                        response("success") = True
                        response("message") = "Draft PO updated successfully."
                    Else
                        Throw New Exception("No rows were updated. Record not found or data unchanged.")
                    End If
                End Using
            End Using
        Catch ex As Exception
            response("success") = False
            response("message") = "Error saving Draft PO: " & ex.Message
        End Try

        context.Response.Write(JsonConvert.SerializeObject(response))
    End Sub

    ' =================================================
    ' ===== ADDED: SAVE DRAFT PO TXN FUNCTION =====
    ' =================================================
    Private Sub SaveDraftPOTXN(context As HttpContext)
        context.Response.ContentType = "application/json"
        Dim response As New Dictionary(Of String, Object)
        Dim createdBy As String = "System" ' Default

        If context.Session("user") IsNot Nothing Then createdBy = context.Session("user").ToString()

        Try


            ' ดึงข้อมูลจาก Form
            Dim year As String = context.Request.Form("year")
            Dim month As String = context.Request.Form("month")
            Dim company As String = context.Request.Form("company")
            Dim category As String = context.Request.Form("category")
            Dim segment As String = context.Request.Form("segment")
            Dim brand As String = context.Request.Form("brand")
            Dim vendor As String = context.Request.Form("vendor")
            Dim pono As String = context.Request.Form("pono")
            Dim amtCCY As Decimal = Decimal.Parse(context.Request.Form("amtCCY"), CultureInfo.InvariantCulture)
            Dim ccy As String = context.Request.Form("ccy")
            Dim exRate As Decimal = Decimal.Parse(context.Request.Form("exRate"), CultureInfo.InvariantCulture)
            Dim remark As String = context.Request.Form("remark")
            Dim amtTHB As Decimal = amtCCY * exRate

            ' 1. เรียก Validation
            Dim Validator As New POValidate()
            Dim errors As Dictionary(Of String, String) = Validator.ValidateDraftPOCreation(
                year, month, company, category, segment, brand, vendor,
                pono, amtCCY, ccy, exRate, amtTHB
            )

            If errors.Count > 0 Then
                response("success") = False
                response("message") = "Validation Failed: " & String.Join(", ", errors.Values)
                context.Response.Write(JsonConvert.SerializeObject(response))
                Return
            End If

            Using conn As New SqlConnection(connectionString)
                conn.Open()
                ' ปรับปรุง Query ให้ตรงกับตาราง [BMS].[dbo].[Draft_PO_Transaction]
                Dim query As String = "INSERT INTO [BMS].[dbo].[Draft_PO_Transaction] 
                                           ([DraftPO_No], [PO_Year], [PO_Month], [Company_Code], 
                                            [Category_Code], [Segment_Code], [Brand_Code], [Vendor_Code], 
                                            [CCY], [Exchange_Rate], [Amount_CCY], [Amount_THB], 
                                            [PO_Type], [Status], [Status_Date], [Status_By], [Actual_PO_Ref], 
                                            [Remark], [Created_By], [Created_Date])
                                     VALUES 
                                           (@pono, @year, @month, @company, 
                                            @category, @segment, @brand, @vendor, 
                                            @ccy, @exRate, @amtCCY, @amtTHB, 
                                            'Manual', 'Draft', GETDATE(), @createdBy, NULL, 
                                            @remark, @createdBy, GETDATE())"

                Using cmd As New SqlCommand(query, conn)
                    ' ปรับปรุงชื่อ Parameters ให้ตรงกับตัวแปร
                    cmd.Parameters.AddWithValue("@pono", pono)
                    cmd.Parameters.AddWithValue("@year", year)
                    cmd.Parameters.AddWithValue("@month", month)
                    cmd.Parameters.AddWithValue("@company", company)
                    cmd.Parameters.AddWithValue("@category", category)
                    cmd.Parameters.AddWithValue("@segment", segment)
                    cmd.Parameters.AddWithValue("@brand", brand)
                    cmd.Parameters.AddWithValue("@vendor", vendor)
                    cmd.Parameters.AddWithValue("@amtTHB", amtTHB)
                    cmd.Parameters.AddWithValue("@amtCCY", amtCCY)
                    cmd.Parameters.AddWithValue("@ccy", ccy)
                    cmd.Parameters.AddWithValue("@exRate", exRate)
                    cmd.Parameters.AddWithValue("@createdBy", createdBy)
                    cmd.Parameters.AddWithValue("@remark", If(String.IsNullOrEmpty(remark), DBNull.Value, remark))
                    ' [Actual_PO_Ref] ถูกตั้งค่าเป็น NULL ใน query โดยตรง

                    cmd.ExecuteNonQuery()
                End Using
            End Using

            response("success") = True
            response("message") = "Draft PO saved successfully."

        Catch ex As Exception
            response("success") = False
            response("message") = "Error saving Draft PO: " & ex.Message
            ' (ควร Log ex.ToString() ไว้ด้วย)
        End Try

        context.Response.Write(JsonConvert.SerializeObject(response))
    End Sub
    ' =================================================

    Private Function GenerateHtmlApprovedTable(dt As DataTable) As String
        Dim sb As New StringBuilder()
        If dt.Rows.Count = 0 Then
            sb.Append("<tr><td colspan='24' class='text-center text-muted'>No approved OTB records found</td></tr>")
        Else
            For Each row As DataRow In dt.Rows
                sb.Append("<tr>")

                ' 1. Create Date (เช่น draftPODate หรือ CreateDT)
                Dim createDate As String = ""
                If row.Table.Columns.Contains("draftPODate") AndAlso row("draftPODate") IsNot DBNull.Value Then
                    createDate = Convert.ToDateTime(row("draftPODate")).ToString("dd/MM/yyyy HH:mm tt")
                ElseIf row.Table.Columns.Contains("CreateDT") AndAlso row("CreateDT") IsNot DBNull.Value Then
                    createDate = Convert.ToDateTime(row("CreateDT")).ToString("dd/MM/yyyy HH:mm tt")
                End If
                sb.AppendFormat("<td>{0}</td>", HttpUtility.HtmlEncode(createDate))

                ' 2. PO No (draftPONo)
                sb.AppendFormat("<td>{0}</td>", HttpUtility.HtmlEncode(GetDbString(row, "draftPONo")))

                ' 3. Status/Type (draftType)
                sb.AppendFormat("<td>{0}</td>", HttpUtility.HtmlEncode(GetDbString(row, "draftType")))

                ' 4. Year (draftYear)
                sb.AppendFormat("<td>{0}</td>", HttpUtility.HtmlEncode(GetDbString(row, "draftYear")))

                ' 5. Month (draftMonth - แปลงจากตัวเลขเป็นชื่อย่อ)
                sb.AppendFormat("<td>{0}</td>", HttpUtility.HtmlEncode(GetMonthName(GetDbString(row, "draftMonth"))))

                ' 6. Cat Code (draftCategory)
                sb.AppendFormat("<td>{0}</td>", HttpUtility.HtmlEncode(GetDbString(row, "draftCategory")))

                ' 7. Cat Name (CategoryName)
                sb.AppendFormat("<td>{0}</td>", HttpUtility.HtmlEncode(GetDbString(row, "CategoryName")))

                ' 8. Comp Code (draftCompany)
                sb.AppendFormat("<td>{0}</td>", HttpUtility.HtmlEncode(GetDbString(row, "draftCompany")))

                ' 9. Seg Code (draftSegment)
                sb.AppendFormat("<td>{0}</td>", HttpUtility.HtmlEncode(GetDbString(row, "draftSegment")))

                ' 10. Seg Name (SegmentName)
                sb.AppendFormat("<td>{0}</td>", HttpUtility.HtmlEncode(GetDbString(row, "SegmentName")))

                ' 11. Brand Code (draftBrand)
                sb.AppendFormat("<td>{0}</td>", HttpUtility.HtmlEncode(GetDbString(row, "draftBrand")))

                ' 12. Brand Name (BrandName)
                sb.AppendFormat("<td>{0}</td>", HttpUtility.HtmlEncode(GetDbString(row, "BrandName")))

                ' 13. Vendor Code (draftVendor)
                sb.AppendFormat("<td>{0}</td>", HttpUtility.HtmlEncode(GetDbString(row, "draftVendor")))

                ' 14. Vendor Name (VendorName)
                sb.AppendFormat("<td>{0}</td>", HttpUtility.HtmlEncode(GetDbString(row, "VendorName")))

                ' 15. Amount THB (AmountTHB)
                sb.AppendFormat("<td class='text-end'>{0}</td>", GetDbDecimal(row, "AmountTHB").ToString("N2"))

                ' 16. Amount CCY (AmountCCY)
                sb.AppendFormat("<td class='text-end'>{0}</td>", GetDbDecimal(row, "AmountCCY").ToString("N2"))

                ' 17. Currency (CCY)
                sb.AppendFormat("<td>{0}</td>", HttpUtility.HtmlEncode(GetDbString(row, "CCY")))

                ' 18. Ex. Rate (ExRate)
                sb.AppendFormat("<td class='text-end'>{0}</td>", GetDbDecimal(row, "ExRate").ToString("N2"))

                ' 19. Actual PO Ref (ActualPORef)
                sb.AppendFormat("<td>{0}</td>", HttpUtility.HtmlEncode(GetDbString(row, "ActualPORef")))

                ' 20. Sub-Status (Status - เช่น Cancelled)
                Dim subStatus As String = GetDbString(row, "Status")
                Dim statusClass As String = ""
                If subStatus.ToLower() = "cancelled" Then
                    statusClass = "class='status-cancelled'"
                End If
                sb.AppendFormat("<td {0}>{1}</td>", statusClass, HttpUtility.HtmlEncode(subStatus))

                ' 21. Sub-Status Date (StatusDate)
                Dim statusDate As String = ""
                If row.Table.Columns.Contains("StatusDate") AndAlso row("StatusDate") IsNot DBNull.Value Then
                    statusDate = Convert.ToDateTime(row("StatusDate")).ToString("dd/MM/yyyy HH:mm tt")
                End If
                sb.AppendFormat("<td>{0}</td>", HttpUtility.HtmlEncode(statusDate))

                ' 22. Remark (Remark)
                sb.AppendFormat("<td>{0}</td>", HttpUtility.HtmlEncode(GetDbString(row, "Remark")))

                ' 23. Action By (ActionBy)
                sb.AppendFormat("<td>{0}</td>", HttpUtility.HtmlEncode(GetDbString(row, "ActionBy")))

                ' 24. Action Button (Static)
                ' (คุณอาจจะต้องเพิ่ม RunNo หรือ ID ให้ปุ่มนี้ ถ้าปุ่ม Edit ต้องทำงาน)
                sb.Append("<td><button class=""btn btn-action""><i class=""bi bi-pencil""></i> Edit</button></td>")

                sb.Append("</tr>")
            Next

        End If
        Return sb.ToString()
    End Function

    ''' <summary>
    ''' (Helper) ดึงค่า String จาก DataRow โดยเช็ค DBNull
    ''' </summary>
    Private Function GetDbString(row As DataRow, columnName As String) As String
        If row IsNot Nothing AndAlso row.Table.Columns.Contains(columnName) AndAlso row(columnName) IsNot DBNull.Value Then
            Return row(columnName).ToString().Trim()
        End If
        Return ""
    End Function

    ''' <summary>
    ''' (Helper) ดึงค่า Decimal จาก DataRow โดยเช็ค DBNull
    ''' </summary>
    Private Function GetDbDecimal(row As DataRow, columnName As String) As Decimal
        Dim val As Decimal = 0
        If row IsNot Nothing AndAlso row.Table.Columns.Contains(columnName) AndAlso row(columnName) IsNot DBNull.Value Then
            Decimal.TryParse(row(columnName).ToString(), val)
        End If
        Return val
    End Function

    ''' <summary>
    ''' (Helper) แปลงเลขเดือนเป็นชื่อย่อ (เช่น "6" -> "Jun")
    ''' </summary>
    Private Function GetMonthName(month As Object) As String
        If month Is DBNull.Value OrElse month Is Nothing Then Return ""

        Dim monthStr As String = month.ToString()
        Dim monthInt As Integer
        If Not Integer.TryParse(monthStr, monthInt) Then
            Return monthStr ' ถ้าเป็นชื่ออยู่แล้ว ก็ส่งกลับเลย
        End If

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