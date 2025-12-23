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
Imports BMS ' Import Namespace ของ Project เพื่อเรียกใช้ POValidate

Public Class DataPOHandler
    Implements System.Web.IHttpHandler, IRequiresSessionState

    Private Shared connectionString As String = ConfigurationManager.ConnectionStrings("BMSConnectionString")?.ConnectionString

    Sub ProcessRequest(ByVal context As HttpContext) Implements IHttpHandler.ProcessRequest
        Dim action As String = If(context.Request("action"), "").ToLower().Trim()

        ' --- Export Actions ---
        If action = "exportdraftpo" Then
            ExportDraftPOToExcel(context)
        ElseIf action = "exportactualpo" Then
            ExportActualPOToExcel(context)

            ' --- Get List Actions ---
        ElseIf action = "getactualpolist" Then
            GetActualPOList(context)
        ElseIf action = "getdraftpolist" Then
            GetDraftPOList(context)
        ElseIf action = "getdraftpodetails" Then
            GetDraftPODetails(context)

            ' --- Transaction Actions (Save/Edit/Delete) ---
        ElseIf action = "savedraftpotxn" Then
            SaveDraftPOTXN(context)
        ElseIf action = "savedraftpoedit" Then
            SaveDraftPOEdit(context)
        ElseIf action = "deletedraftpo" Then
            DeleteDraftPO(context)

            ' --- Special Action (SAP Filter) ---
        ElseIf context.Request("action") = "PoListFilter" Then
            ProcessPoListFilter(context)

        Else
            ' Default Action
            context.Response.ContentType = "text/plain"
            context.Response.Write("Invalid Action")
        End If
    End Sub

    ' =================================================
    ' ===== SAVE DRAFT PO TXN (Create New) =====
    ' =================================================
    Private Sub SaveDraftPOTXN(context As HttpContext)
        context.Response.ContentType = "application/json"
        Dim response As New Dictionary(Of String, Object)
        Dim createdBy As String = "System"

        If context.Session("user") IsNot Nothing Then createdBy = context.Session("user").ToString()

        Try
            ' 1. รับค่าจาก Form
            Dim year As String = context.Request.Form("year")
            Dim month As String = context.Request.Form("month")
            Dim company As String = context.Request.Form("company")
            Dim category As String = context.Request.Form("category")
            Dim segment As String = context.Request.Form("segment")
            Dim brand As String = context.Request.Form("brand")
            Dim vendor As String = context.Request.Form("vendor")
            Dim pono As String = context.Request.Form("pono")
            Dim ccy As String = context.Request.Form("ccy")
            Dim remark As String = context.Request.Form("remark")

            Dim amtCCY As Decimal = 0
            Decimal.TryParse(context.Request.Form("amtCCY"), NumberStyles.Any, CultureInfo.InvariantCulture, amtCCY)

            Dim exRate As Decimal = 0
            Decimal.TryParse(context.Request.Form("exRate"), NumberStyles.Any, CultureInfo.InvariantCulture, exRate)

            Dim amtTHB As Decimal = amtCCY * exRate

            ' 2. สร้าง Item เพื่อส่งไป Validate (ใช้ POValidate ตัวใหม่)
            Dim newItem As New POValidate.DraftPOItem With {
                .RowIndex = 1,
                .DraftPO_ID = 0, ' New Record = 0 เสมอ
                .PO_Year = year,
                .PO_Month = month,
                .Company_Code = company,
                .Category_Code = category,
                .Segment_Code = segment,
                .Brand_Code = brand,
                .Vendor_Code = vendor,
                .PO_No = pono,
                .Currency = ccy,
                .Amount_CCY = amtCCY,
                .ExchangeRate = exRate,
                .Amount_THB = amtTHB
            }

            ' 3. เรียก Validation Logic
            Dim Validator As New POValidate()
            Dim items As New List(Of POValidate.DraftPOItem)
            items.Add(newItem)

            Dim result As POValidate.ValidationResult = Validator.ValidateBatch(items)

            If Not result.IsValid Then
                response("success") = False
                Dim errorMsg As String = result.GlobalError
                If result.RowErrors.ContainsKey(1) Then
                    errorMsg = String.Join(", ", result.RowErrors(1))
                End If
                response("message") = "Validation Failed: " & errorMsg
                context.Response.Write(JsonConvert.SerializeObject(response))
                Return
            End If

            ' 4. ถ้าผ่าน Insert ลง DB
            Using conn As New SqlConnection(connectionString)
                conn.Open()
                Dim query As String = "INSERT INTO [BMS].[dbo].[Draft_PO_Transaction] " &
                                      "([DraftPO_No], [PO_Year], [PO_Month], [Company_Code], [Category_Code], [Segment_Code], [Brand_Code], [Vendor_Code], " &
                                      "[CCY], [Exchange_Rate], [Amount_CCY], [Amount_THB], " &
                                      "[PO_Type], [Status], [Status_Date], [Status_By], [Actual_PO_Ref], " &
                                      "[Remark], [Created_By], [Created_Date]) " &
                                      "VALUES " &
                                      "(@pono, @year, @month, @company, @category, @segment, @brand, @vendor, " &
                                      "@ccy, @exRate, @amtCCY, @amtTHB, " &
                                      "'Manual', 'Draft', GETDATE(), @createdBy, NULL, " &
                                      "@remark, @createdBy, GETDATE())"

                Using cmd As New SqlCommand(query, conn)
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
                    cmd.ExecuteNonQuery()
                End Using
            End Using

            response("success") = True
            response("message") = "Draft PO saved successfully."

        Catch ex As Exception
            response("success") = False
            response("message") = "Error saving Draft PO: " & ex.Message
        End Try

        context.Response.Write(JsonConvert.SerializeObject(response))
    End Sub

    ' =================================================
    ' ===== SAVE DRAFT PO EDIT (Update) =====
    ' =================================================
    Private Sub SaveDraftPOEdit(context As HttpContext)
        context.Response.ContentType = "application/json"
        Dim response As New Dictionary(Of String, Object)
        Dim statusBy As String = "System"

        Try
            If context.Session IsNot Nothing AndAlso context.Session("user") IsNot Nothing Then
                statusBy = context.Session("user").ToString()
            End If

            ' 1. รับค่า
            Dim draftPOID As Integer = Integer.Parse(context.Request.Form("draftPOID"))
            Dim draftPOno As String = context.Request.Form("pono")
            Dim year As String = context.Request.Form("year")
            Dim month As String = context.Request.Form("month")
            Dim company As String = context.Request.Form("company")
            Dim category As String = context.Request.Form("category")
            Dim segment As String = context.Request.Form("segment")
            Dim brand As String = context.Request.Form("brand")
            Dim vendor As String = context.Request.Form("vendor")
            Dim ccy As String = context.Request.Form("ccy")
            Dim remark As String = context.Request.Form("remark")

            Dim amtCCY As Decimal = 0
            Decimal.TryParse(context.Request.Form("amtCCY"), NumberStyles.Any, CultureInfo.InvariantCulture, amtCCY)

            Dim exRate As Decimal = 0
            Decimal.TryParse(context.Request.Form("exRate"), NumberStyles.Any, CultureInfo.InvariantCulture, exRate)

            Dim amtTHB As Decimal = amtCCY * exRate

            ' 2. Validate (สำคัญ: ส่ง ID เพื่อ Exclude Self)
            Dim editItem As New POValidate.DraftPOItem With {
                .RowIndex = 1,
                .DraftPO_ID = draftPOID, ' *** ส่ง ID เดิมเพื่อให้ระบบรู้ว่าเป็น Edit และ Exclude ยอดเดิม ***
                .PO_Year = year,
                .PO_Month = month,
                .Company_Code = company,
                .Category_Code = category,
                .Segment_Code = segment,
                .Brand_Code = brand,
                .Vendor_Code = vendor,
                .PO_No = draftPOno,
                .Currency = ccy,
                .Amount_CCY = amtCCY,
                .ExchangeRate = exRate,
                .Amount_THB = amtTHB
            }

            Dim Validator As New POValidate()
            Dim items As New List(Of POValidate.DraftPOItem)
            items.Add(editItem)

            Dim result As POValidate.ValidationResult = Validator.ValidateBatch(items)

            If Not result.IsValid Then
                response("success") = False
                Dim errorMsg As String = result.GlobalError
                If result.RowErrors.ContainsKey(1) Then
                    errorMsg = String.Join(", ", result.RowErrors(1))
                End If
                response("message") = "Update Failed: " & errorMsg
                context.Response.Write(JsonConvert.SerializeObject(response))
                Return
            End If

            ' 3. Update DB
            Using conn As New SqlConnection(connectionString)
                conn.Open()
                Dim query As String = "UPDATE [BMS].[dbo].[Draft_PO_Transaction] SET " &
                                      "[DraftPO_No] = @draftPono, [PO_Year] = @year, [PO_Month] = @month, " &
                                      "[Company_Code] = @company, [Category_Code] = @category, " &
                                      "[Segment_Code] = @segment, [Brand_Code] = @brand, [Vendor_Code] = @vendor, " &
                                      "[CCY] = @ccy, [Exchange_Rate] = @exRate, [Amount_CCY] = @amtCCY, [Amount_THB] = @amtTHB, " &
                                      "[Status] = 'Edited', [Status_Date] = GETDATE(), [Status_By] = @statusBy, [Remark] = @remark " &
                                      "WHERE [DraftPO_ID] = @draftPOID"

                Using cmd As New SqlCommand(query, conn)
                    cmd.Parameters.AddWithValue("@draftPono", draftPOno)
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
                        Throw New Exception("No rows were updated. Record not found.")
                    End If
                End Using
            End Using

        Catch ex As Exception
            response("success") = False
            response("message") = "Error updating Draft PO: " & ex.Message
        End Try

        context.Response.Write(JsonConvert.SerializeObject(response))
    End Sub

    ' =================================================
    ' ===== DELETE (CANCEL) DRAFT PO =====
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
                ' Check Status First
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

                Dim query As String = "UPDATE [BMS].[dbo].[Draft_PO_Transaction] " &
                                      "SET [Status] = 'Cancelled', [Status_Date] = GETDATE(), [Status_By] = @StatusBy " &
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

    ' =================================================
    ' ===== GET DETAILS (For Edit Modal) =====
    ' =================================================
    Private Sub GetDraftPODetails(context As HttpContext)
        Try
            Dim draftPOID As Integer = 0
            Integer.TryParse(context.Request.Form("draftPOID"), draftPOID)

            If draftPOID = 0 Then Throw New Exception("Invalid DraftPO_ID.")

            Dim dt As New DataTable()
            Using conn As New SqlConnection(connectionString)
                conn.Open()
                Dim query As String = "SELECT * FROM [BMS].[dbo].[Draft_PO_Transaction] WHERE DraftPO_ID = @DraftPOID"
                Using cmd As New SqlCommand(query, conn)
                    cmd.Parameters.AddWithValue("@DraftPOID", draftPOID)
                    Using reader As SqlDataReader = cmd.ExecuteReader()
                        dt.Load(reader)
                    End Using
                End Using
            End Using

            If dt.Rows.Count > 0 Then
                Dim row As DataRow = dt.Rows(0)
                Dim dict As New Dictionary(Of String, Object)()
                For Each col As DataColumn In row.Table.Columns
                    dict(col.ColumnName) = If(row(col) Is DBNull.Value, Nothing, row(col))
                Next
                context.Response.Write(JsonConvert.SerializeObject(dict, Formatting.None))
            Else
                Throw New Exception("Record not found.")
            End If

        Catch ex As Exception
            context.Response.StatusCode = 200 ' ส่ง 200 แต่เป็น error msg
            context.Response.Write(JsonConvert.SerializeObject(New With {.success = False, .message = ex.Message}))
        End Try
    End Sub

    ' =================================================
    ' ===== LIST & EXPORT HELPERS =====
    ' =================================================

    Private Sub GetDraftPOList(context As HttpContext)
        Try
            Dim dt As DataTable = GetDraftPODataTable(context, False)
            Dim jsonResult As String = JsonConvert.SerializeObject(dt, Formatting.None)
            context.Response.Write(jsonResult)
        Catch ex As Exception
            context.Response.StatusCode = 200
            context.Response.Write(JsonConvert.SerializeObject(New With {.success = False, .message = ex.Message}))
        End Try
    End Sub

    Private Sub GetActualPOList(context As HttpContext)
        Try
            Dim year As String = context.Request.Form("year")
            Dim month As String = context.Request.Form("month")
            Dim company As String = context.Request.Form("company")
            Dim category As String = context.Request.Form("category")
            Dim segment As String = context.Request.Form("segment")
            Dim brand As String = context.Request.Form("brand")
            Dim vendor As String = context.Request.Form("vendor")

            Dim dt As DataTable = GetActualPOData(year, month, company, category, segment, brand, vendor)
            Dim html As String = GenerateHtmlActualPOTable(dt)
            context.Response.ContentType = "text/html"
            context.Response.Write(html)
        Catch ex As Exception
            context.Response.Write($"<tr><td colspan='22' class='text-danger'>Error: {ex.Message}</td></tr>")
        End Try
    End Sub

    Private Sub ExportDraftPOToExcel(context As HttpContext)
        Try
            Dim dtRaw As DataTable = GetDraftPODataTable(context, True)
            Dim dtExport As New DataTable("DraftPO")
            ' ... (Define Columns) ...
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
            ExportToExcel(context, dtExport, "Draft_PO_" & DateTime.Now.ToString("yyyyMMdd_HHmmss") & ".xlsx")
        Catch ex As Exception
            context.Response.Write("Error exporting data: " & ex.Message)
        End Try
    End Sub

    Private Sub ExportActualPOToExcel(context As HttpContext)
        Dim year As String = context.Request.QueryString("year")
        Dim dtRaw As DataTable = GetActualPOData(year, context.Request.QueryString("month"), context.Request.QueryString("company"),
                                                context.Request.QueryString("category"), context.Request.QueryString("segment"),
                                                context.Request.QueryString("brand"), context.Request.QueryString("vendor"))

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
        ExportToExcel(context, dtExport, "Actual_PO_" & DateTime.Now.ToString("yyyyMMdd_HHmmss") & ".xlsx")
    End Sub

    ' =================================================
    ' ===== SPECIAL: SAP LIST FILTER =====
    ' =================================================
    Private Sub ProcessPoListFilter(context As HttpContext)
        Try
            ' Example: ใช้ SapApiHelper (Code เดิมของคุณ)
            ' ปรับ Parameter Date ตามต้องการ หรือรับจาก Request
            Dim filterDate As Date = DateTime.Today

            Dim poList As List(Of SapPOResultItem) = Task.Run(Async Function()
                                                                  Return Await SapApiHelper.GetPOsAsync(filterDate)
                                                              End Function).Result

            context.Response.ContentType = "application/json"
            Dim successResponse = New With {
                .success = True,
                .count = If(poList IsNot Nothing, poList.Count, 0),
                .data = poList
            }
            context.Response.Write(JsonConvert.SerializeObject(successResponse))

        Catch ex As Exception
            context.Response.ContentType = "application/json"
            Dim errorResponse As New With {
                .success = False,
                .message = ex.Message
            }
            context.Response.Write(JsonConvert.SerializeObject(errorResponse))
        End Try
    End Sub


    ' =================================================
    ' ===== INTERNAL HELPERS (DB & FORMAT) =====
    ' =================================================

    Private Function GetDraftPODataTable(context As HttpContext, isExport As Boolean) As DataTable
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
            query.Append("SELECT DISTINCT po.DraftPO_ID, po.Created_Date, po.DraftPO_No, po.PO_Type, po.PO_Year, ")
            query.Append("m.month_name_sh AS PO_Month_Name, po.Category_Code, c.Category AS Category_Name, ")
            query.Append("po.Company_Code, comp.[CompanyNameShort] As Company_Name, po.Segment_Code, s.SegmentName AS Segment_Name, ")
            query.Append("po.Brand_Code, b.[Brand Name] AS Brand_Name, po.Vendor_Code, v.Vendor AS Vendor_Name, ")
            query.Append("po.Amount_THB, po.Amount_CCY, po.CCY, po.Exchange_Rate, po.Actual_PO_Ref, po.Status, ")
            query.Append("po.Status_Date, po.Remark, po.Status_By ")
            query.Append("FROM [BMS].[dbo].[Draft_PO_Transaction] po ")
            query.Append("LEFT JOIN [BMS].[dbo].[MS_Month] m ON po.PO_Month = m.month_code ")
            query.Append("LEFT JOIN [BMS].[dbo].[MS_Company] comp ON po.Company_Code = comp.[CompanyCode] ")
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

    Private Function GenerateHtmlActualPOTable(dt As DataTable) As String
        Dim sb As New StringBuilder()
        ' ใช้ Class เดิมที่มีในระบบ (ถ้ามี) หรือเขียน Helper เอง
        ' ในที่นี้สมมติว่า MasterDataUtil มีอยู่แล้วตาม Code เดิม
        Dim masterinstance As New MasterDataUtil

        If dt.Rows.Count = 0 Then
            sb.Append("<tr><td colspan='22' class='text-center text-muted p-4'>No Actual PO records found for the selected filters.</td></tr>")
            Return sb.ToString()
        End If

        For Each row As DataRow In dt.Rows
            Dim status As String = GetDbString(row, "Status")
            Dim statusClass As String = ""
            If status.ToLower() = "cancelled" Then statusClass = "class='status-cancelled'"

            sb.Append("<tr>")
            sb.AppendFormat("<td>{0}</td>", GetDbDate(row, "Actual_PO_Date"))
            sb.AppendFormat("<td>{0}</td>", HttpUtility.HtmlEncode(GetDbString(row, "Actual_PO_No")))
            sb.AppendFormat("<td>{0}</td>", HttpUtility.HtmlEncode(GetDbString(row, "PO_Type")))
            sb.AppendFormat("<td>{0}</td>", HttpUtility.HtmlEncode(GetDbString(row, "PO_Year")))
            sb.AppendFormat("<td>{0}</td>", HttpUtility.HtmlEncode(GetDbString(row, "PO_Month_Name")))
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
            sb.AppendFormat("<td class='text-end'>{0}</td>", GetDbDecimal(row, "Exchange_Rate").ToString("N4"))
            sb.AppendFormat("<td>{0}</td>", HttpUtility.HtmlEncode(GetDbString(row, "Draft_PO_Ref")))
            sb.AppendFormat("<td>{0}</td>", HttpUtility.HtmlEncode(GetDbString(row, "Remark")))
            sb.AppendFormat("<td {0}>{1}</td>", statusClass, HttpUtility.HtmlEncode(status))
            sb.AppendFormat("<td>{0}</td>", GetDbDate(row, "Status_Date"))
            sb.Append("</tr>")
        Next
        Return sb.ToString()
    End Function

    Private Sub ExportToExcel(context As HttpContext, dt As DataTable, filename As String)
        ExcelPackage.License.SetNonCommercialOrganization("KingPower")
        Using package As New ExcelPackage()
            Dim worksheet = package.Workbook.Worksheets.Add("Sheet1")
            worksheet.Cells("A1").LoadFromDataTable(dt, True)
            Using range = worksheet.Cells(1, 1, 1, dt.Columns.Count)
                range.Style.Font.Bold = True
                range.Style.Fill.PatternType = ExcelFillStyle.Solid
                range.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(11, 86, 164))
                range.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(11, 86, 164))
                range.Style.Font.Color.SetColor(System.Drawing.Color.White)
            End Using
            worksheet.Cells.AutoFitColumns()
            context.Response.Clear()
            context.Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            context.Response.AddHeader("content-disposition", $"attachment; filename={filename}")
            context.Response.BinaryWrite(package.GetAsByteArray())
            context.Response.Flush()
            context.ApplicationInstance.CompleteRequest()
        End Using
    End Sub

    ' Helpers for Safe Data Retrieval
    Private Function GetDbString(row As DataRow, columnName As String) As String
        If row IsNot Nothing AndAlso row.Table.Columns.Contains(columnName) AndAlso row(columnName) IsNot DBNull.Value Then
            Return row(columnName).ToString().Trim()
        End If
        Return ""
    End Function

    Private Function GetDbDate(row As DataRow, columnName As String) As String
        If row IsNot Nothing AndAlso row.Table.Columns.Contains(columnName) AndAlso row(columnName) IsNot DBNull.Value Then
            Try
                Return Convert.ToDateTime(row(columnName)).ToString("dd/MM/yyyy HH:mm")
            Catch
                Return row(columnName).ToString()
            End Try
        End If
        Return ""
    End Function

    Private Function GetDbDecimal(row As DataRow, columnName As String) As Decimal
        Dim val As Decimal = 0
        If row IsNot Nothing AndAlso row.Table.Columns.Contains(columnName) AndAlso row(columnName) IsNot DBNull.Value Then
            Decimal.TryParse(row(columnName).ToString(), val)
        End If
        Return val
    End Function

    ReadOnly Property IsReusable() As Boolean Implements IHttpHandler.IsReusable
        Get
            Return False
        End Get
    End Property

End Class