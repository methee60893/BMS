Imports System
Imports System.Web
Imports System.Data
Imports System.Data.SqlClient
Imports System.IO
Imports ExcelDataReader
Imports System.Text
Imports System.Web.Script.Serialization
Imports System.Collections.Generic
Imports System.Globalization

Public Class SwitchUploadHandler
    Implements IHttpHandler, IRequiresSessionState

    Private Shared connectionString As String = ConfigurationManager.ConnectionStrings("BMSConnectionString")?.ConnectionString

    ' ตัวแปรสำหรับเก็บ Master Data
    Private dtMonth As DataTable
    Private dtCompany As DataTable
    Private dtCategory As DataTable
    Private dtSegment As DataTable
    Private dtBrand As DataTable
    Private dtVendor As DataTable

    Private Class SwitchTypeInfo
        Public Property TypeFrom As String
        Public Property TypeTo As String
        Public Property DisplayName As String
    End Class

    Private Class SwitchUploadDbItem
        Public Property FromData As Dictionary(Of String, Object)
        Public Property ToData As Dictionary(Of String, Object)
        Public Property Amount As Decimal
        Public Property TypeFrom As String
        Public Property TypeTo As String
        Public Property Remark As String
        Public Property ActionType As String
        Public Property TypeName As String
    End Class

    Public Sub ProcessRequest(ByVal context As HttpContext) Implements IHttpHandler.ProcessRequest
        context.Response.Clear()
        context.Response.ContentEncoding = Encoding.UTF8

        Dim action As String = If(context.Request("action"), "").ToLowerInvariant()
        Dim uploadBy As String = If(context.Session("user") IsNot Nothing, context.Session("user").ToString(), "System")

        If action = "preview" Then
            HandlePreview(context)
        ElseIf action = "save" Then
            HandleSave(context, uploadBy)
        Else
            context.Response.ContentType = "application/json"
            context.Response.Write(New JavaScriptSerializer().Serialize(New With {.success = False, .message = "Invalid action."}))
        End If
    End Sub

    ' ==========================================
    ' 1. PREVIEW & VALIDATE
    ' ==========================================
    Private Sub HandlePreview(context As HttpContext)
        context.Response.ContentType = "application/json"

        If context.Request.Files.Count = 0 Then
            context.Response.Write(New JavaScriptSerializer().Serialize(New With {.success = False, .message = "No file uploaded."}))
            Return
        End If

        Dim postedFile As HttpPostedFile = context.Request.Files(0)
        Dim tempPath As String = Path.GetTempFileName()
        postedFile.SaveAs(tempPath)

        Try
            ' 1. โหลด Master Data เตรียมไว้
            LoadAllMasterData()

            ' 2. อ่าน Excel
            Dim dt As DataTable = ReadExcel(tempPath)

            ' 3. สร้างตาราง Preview พร้อมชื่อ Description
            Dim result = GeneratePreviewData(dt)

            context.Response.Write(New JavaScriptSerializer().Serialize(result))
        Catch ex As Exception
            context.Response.Write(New JavaScriptSerializer().Serialize(New With {.success = False, .message = "Error: " & ex.Message}))
        Finally
            If File.Exists(tempPath) Then File.Delete(tempPath)
        End Try
    End Sub

    Private Sub LoadAllMasterData()
        Using conn As New SqlConnection(connectionString)
            conn.Open()

            ' Load Month
            dtMonth = New DataTable()
            Using cmd As New SqlCommand("SELECT [month_code], [month_name_sh] FROM [MS_Month]", conn)
                dtMonth.Load(cmd.ExecuteReader())
            End Using

            ' Load Company
            dtCompany = New DataTable()
            Using cmd As New SqlCommand("SELECT [CompanyCode], [CompanyNameShort] FROM [MS_Company]", conn)
                dtCompany.Load(cmd.ExecuteReader())
            End Using

            ' Load Category
            dtCategory = New DataTable()
            Using cmd As New SqlCommand("SELECT [Cate], [Category] FROM [MS_Category]", conn)
                dtCategory.Load(cmd.ExecuteReader())
            End Using

            ' Load Segment
            dtSegment = New DataTable()
            Using cmd As New SqlCommand("SELECT [SegmentCode], [SegmentName] FROM [MS_Segment]", conn)
                dtSegment.Load(cmd.ExecuteReader())
            End Using

            ' Load Brand
            dtBrand = New DataTable()
            Using cmd As New SqlCommand("SELECT [Brand Code], [Brand Name] FROM [MS_Brand]", conn)
                dtBrand.Load(cmd.ExecuteReader())
            End Using

            ' Load Vendor
            dtVendor = New DataTable()
            Using cmd As New SqlCommand("SELECT [VendorCode], [Vendor] FROM [MS_Vendor]", conn)
                dtVendor.Load(cmd.ExecuteReader())
            End Using
        End Using
    End Sub

    Private Function GetName(dt As DataTable, codeCol As String, nameCol As String, codeVal As String) As String
        If String.IsNullOrEmpty(codeVal) OrElse dt Is Nothing Then Return ""
        Dim rows = dt.Select("[" & codeCol & "] = '" & codeVal.Replace("'", "''") & "'")
        If rows.Length > 0 Then
            Return rows(0)(nameCol).ToString()
        End If
        Return "" ' Not found
    End Function

    Private Function ReadExcel(filePath As String) As DataTable
        Using stream = File.Open(filePath, FileMode.Open, FileAccess.Read)
            Using reader = ExcelReaderFactory.CreateReader(stream)
                Dim conf As New ExcelDataSetConfiguration() With {
                    .ConfigureDataTable = Function(__) New ExcelDataTableConfiguration() With {.UseHeaderRow = True}
                }
                Return reader.AsDataSet(conf).Tables(0)
            End Using
        End Using
    End Function

    Private Function GeneratePreviewData(dt As DataTable) As Object

        Dim sb As New StringBuilder()
        Dim budgetCalc As New OTBBudgetCalculator()
        Dim pendingUsage As New Dictionary(Of String, Decimal)()

        ' สร้าง Cache เพื่อเก็บค่าจาก DB ลดการ Query ซ้ำใน Loop
        Dim dbBudgetCache As New Dictionary(Of String, Decimal)()
        Dim dbUsedCache As New Dictionary(Of String, Decimal)()
        Dim povalidate As New POValidate() ' สร้างไว้นอก Loop

        Dim uniqueRows As New HashSet(Of String)()
        Dim allValid As Boolean = True

        sb.Append("<div class='table-responsive' style='max-height: 500px;'>")
        sb.Append("<table class='table table-bordered table-sm table-hover' id='tblBulkPreview' style='font-size: 0.85rem;'>")
        sb.Append("<thead class='table-dark sticky-top'><tr>")
        sb.Append("<th style='width:30px;'>No.</th>")
        sb.Append("<th style='width:60px;'>Function</th>")
        sb.Append("<th style='width:60px;'>Type</th>")
        sb.Append("<th>From Details (Source)</th>")
        sb.Append("<th class='text-end' style='width:100px;'>Amount</th>")
        sb.Append("<th>To Details (Destination)</th>")
        sb.Append("<th>Remark</th>")
        sb.Append("<th style='width:80px;'>Status</th><th>Issue</th>")
        sb.Append("</tr></thead><tbody>")

        For i As Integer = 0 To dt.Rows.Count - 1
            Dim row As DataRow = dt.Rows(i)
            Dim errors As New List(Of String)()

            Dim actionType As String = GetFunctionValue(row)
            Dim amtStr As String = GetVal(row, "Amount")
            Dim amount As Decimal = 0
            TryParseAmount(amtStr, amount)

            Dim fromData As Dictionary(Of String, Object) = BuildFieldData(
                GetVal(row, "Year"), GetVal(row, "Month"), GetVal(row, "Company"), GetVal(row, "Category"),
                GetVal(row, "Segment"), GetVal(row, "Brand"), GetVal(row, "Vendor"))

            Dim toData As Dictionary(Of String, Object) = BuildFieldData(
                GetVal(row, "To_Year"), GetVal(row, "To_Month"), GetVal(row, "To_Company"), GetVal(row, "To_Category"),
                GetVal(row, "To_Segment"), GetVal(row, "To_Brand"), GetVal(row, "To_Vendor"))

            Dim typeInfo As SwitchTypeInfo = DetermineSwitchTypeInfo(actionType, GetDictString(fromData, "Year"), GetDictString(fromData, "Month"), GetDictString(toData, "Year"), GetDictString(toData, "Month"))
            Dim previewToData As Dictionary(Of String, Object) = Nothing
            If actionType.Equals("Switch", StringComparison.OrdinalIgnoreCase) Then previewToData = toData

            Dim previewItem As New SwitchUploadDbItem With {
                .FromData = fromData,
                .ToData = previewToData,
                .Amount = amount,
                .TypeFrom = typeInfo.TypeFrom,
                .TypeTo = typeInfo.TypeTo,
                .Remark = GetVal(row, "Remark"),
                .ActionType = actionType,
                .TypeName = typeInfo.DisplayName
            }

            ValidateMappedItem(previewItem, errors)

            If errors.Count = 0 Then
                ValidateDuplicateInBatch(previewItem, uniqueRows, errors)
                If errors.Count = 0 Then ValidateBudget(previewItem, budgetCalc, povalidate, dbBudgetCache, dbUsedCache, pendingUsage, errors)
            End If

            Dim isError As Boolean = (errors.Count > 0)
            If isError Then allValid = False

            AddMasterNames(previewItem.FromData)
            If previewItem.ToData IsNot Nothing Then AddMasterNames(previewItem.ToData)

            Dim rowDataObj As New Dictionary(Of String, Object) From {
                {"Function", previewItem.ActionType},
                {"Type", previewItem.TypeName},
                {"TypeFrom", previewItem.TypeFrom},
                {"TypeTo", previewItem.TypeTo},
                {"Amount", previewItem.Amount},
                {"Remark", previewItem.Remark},
                {"From", previewItem.FromData},
                {"To", previewItem.ToData}
            }
            Dim rowJson As String = If(isError, "", HttpUtility.HtmlAttributeEncode(New JavaScriptSerializer().Serialize(rowDataObj)))

            Dim fHtml As String = BuildDetailHtml(previewItem.FromData)
            Dim tHtml As String = "-"
            If previewItem.ToData IsNot Nothing Then
                tHtml = BuildDetailHtml(previewItem.ToData)
            End If

            Dim trClass As String = If(isError, "table-danger", "")
            sb.AppendFormat("<tr class='{0}'>", trClass)
            sb.AppendFormat("<td class='text-center'>{0}<input type='hidden' class='row-data' value='{1}'></td>", i + 2, rowJson)

            sb.AppendFormat("<td class='text-center fw-bold'>{0}</td>", Html(actionType))
            sb.AppendFormat("<td class='text-center fw-bold'>{0}</td>", Html(previewItem.TypeName))

            sb.AppendFormat("<td>{0}</td>", fHtml)
            sb.AppendFormat("<td class='text-end fw-bold'>{0:N2}</td>", previewItem.Amount)
            sb.AppendFormat("<td>{0}</td>", tHtml)
            sb.AppendFormat("<td><small>{0}</small></td>", Html(previewItem.Remark))

            ' Status & Error
            Dim status As String = If(isError, "<span class='badge bg-danger'>Fail</span>", "<span class='badge bg-success'>Pass</span>")
            Dim encodedErrors As New List(Of String)()
            For Each errorMessage In errors
                encodedErrors.Add(Html(errorMessage))
            Next
            Dim errText As String = If(errors.Count > 0, String.Join("<br/>", encodedErrors), "")
            sb.AppendFormat("<td class='text-center'>{0}</td><td class='text-danger small'>{1}</td>", status, errText)
            sb.Append("</tr>")
        Next
        sb.Append("</tbody></table></div>")

        If Not allValid Then
            sb.Append("<div class='alert alert-warning mt-2'><i class='bi bi-exclamation-triangle'></i> ข้อมูลบางรายการไม่ถูกต้อง กรุณาแก้ไขไฟล์แล้ว Upload ใหม่ (ต้องผ่านทุกรายการถึงจะบันทึกได้)</div>")
        Else
            sb.Append("<div class='alert alert-success mt-2'><i class='bi bi-check-circle'></i> ข้อมูลถูกต้องครบถ้วน พร้อมบันทึกเข้าฐานข้อมูล</div>")
        End If

        Return New With {
            .success = True,
            .html = sb.ToString(),
            .canSubmit = allValid
        }
    End Function

    Private Function GetVal(row As DataRow, col As String) As String
        Return If(row.Table.Columns.Contains(col) AndAlso row(col) IsNot DBNull.Value, row(col).ToString().Trim(), "")
    End Function

    Private Function GetFunctionValue(row As DataRow) As String
        Dim actionType As String = GetVal(row, "Function")
        If String.IsNullOrEmpty(actionType) Then actionType = GetVal(row, "Type")
        Return NormalizeActionType(actionType)
    End Function

    Private Function NormalizeActionType(value As String) As String
        If String.IsNullOrWhiteSpace(value) Then Return ""

        Select Case value.Trim().ToLowerInvariant()
            Case "switch", "switching", "switch in-out", "switch in out"
                Return "Switch"
            Case "extra", "extra budget"
                Return "Extra"
            Case Else
                Return value.Trim()
        End Select
    End Function

    Private Function BuildFieldData(year As String, month As String, company As String, category As String, segment As String, brand As String, vendor As String) As Dictionary(Of String, Object)
        Return New Dictionary(Of String, Object) From {
            {"Year", year},
            {"Month", month},
            {"Company", company},
            {"Category", category},
            {"Segment", segment},
            {"Brand", brand},
            {"Vendor", vendor}
        }
    End Function

    Private Function GetDictString(data As Dictionary(Of String, Object), key As String) As String
        If data Is Nothing OrElse Not data.ContainsKey(key) OrElse data(key) Is Nothing Then Return ""
        Return data(key).ToString().Trim()
    End Function

    Private Function GetPayloadString(row As Dictionary(Of String, Object), key As String) As String
        If row Is Nothing OrElse Not row.ContainsKey(key) OrElse row(key) Is Nothing Then Return ""
        Return row(key).ToString().Trim()
    End Function

    Private Function GetPayloadData(row As Dictionary(Of String, Object), key As String) As Dictionary(Of String, Object)
        If row Is Nothing OrElse Not row.ContainsKey(key) OrElse row(key) Is Nothing Then
            Return New Dictionary(Of String, Object)()
        End If

        Dim data = TryCast(row(key), Dictionary(Of String, Object))
        If data Is Nothing Then Return New Dictionary(Of String, Object)()
        Return data
    End Function

    Private Function TryParseAmount(value As String, ByRef amount As Decimal) As Boolean
        If Decimal.TryParse(value, NumberStyles.Number, CultureInfo.InvariantCulture, amount) Then Return True
        Return Decimal.TryParse(value, NumberStyles.Number, CultureInfo.CurrentCulture, amount)
    End Function

    Private Function DetermineSwitchTypeInfo(actionType As String, fromYear As String, fromMonth As String, toYear As String, toMonth As String) As SwitchTypeInfo
        If actionType.Equals("Extra", StringComparison.OrdinalIgnoreCase) Then
            Return New SwitchTypeInfo With {.TypeFrom = "E", .TypeTo = Nothing, .DisplayName = "Extra"}
        End If

        If Not actionType.Equals("Switch", StringComparison.OrdinalIgnoreCase) Then
            Return New SwitchTypeInfo With {.TypeFrom = "", .TypeTo = Nothing, .DisplayName = ""}
        End If

        Dim info As New SwitchTypeInfo With {.TypeFrom = "D", .TypeTo = "C", .DisplayName = "Switch"}

        Dim fy, fm, ty, tm As Integer
        If Not Integer.TryParse(fromYear, fy) OrElse Not Integer.TryParse(fromMonth, fm) OrElse
           Not Integer.TryParse(toYear, ty) OrElse Not Integer.TryParse(toMonth, tm) OrElse
           fm < 1 OrElse fm > 12 OrElse tm < 1 OrElse tm > 12 Then
            Return info
        End If

        Dim dFrom As New Date(fy, fm, 1)
        Dim dTo As New Date(ty, tm, 1)
        If dFrom > dTo Then
            info.TypeFrom = "G"
            info.TypeTo = "F"
            info.DisplayName = "Carry"
        ElseIf dFrom < dTo Then
            info.TypeFrom = "I"
            info.TypeTo = "H"
            info.DisplayName = "Balance"
        End If

        Return info
    End Function

    Private Sub AddMasterNames(data As Dictionary(Of String, Object))
        If data Is Nothing Then Return

        data("CompanyName") = GetName(dtCompany, "CompanyCode", "CompanyNameShort", GetDictString(data, "Company"))
        data("VendorName") = GetName(dtVendor, "VendorCode", "Vendor", GetDictString(data, "Vendor"))
        data("CategoryName") = GetName(dtCategory, "Cate", "Category", GetDictString(data, "Category"))
        data("SegmentName") = GetName(dtSegment, "SegmentCode", "SegmentName", GetDictString(data, "Segment"))
        data("BrandName") = GetName(dtBrand, "Brand Code", "Brand Name", GetDictString(data, "Brand"))
    End Sub

    Private Function BuildDetailHtml(data As Dictionary(Of String, Object)) As String
        If data Is Nothing Then Return "-"

        Return "<strong>" & Html(GetDictString(data, "Year")) & "/" & Html(GetDictString(data, "Month")) & "</strong><br/>" &
               "<small>Comp: " & Html(GetDictString(data, "Company")) & ":" & Html(GetDictString(data, "CompanyName")) & "<br/>" &
               "Vend: " & Html(GetDictString(data, "Vendor")) & ":" & Html(GetDictString(data, "VendorName")) & "<br/>" &
               "Brand: " & Html(GetDictString(data, "Brand")) & ":" & Html(GetDictString(data, "BrandName")) & "<br/>" &
               "Seg: " & Html(GetDictString(data, "Segment")) & ":" & Html(GetDictString(data, "SegmentName")) &
               " | Cat: " & Html(GetDictString(data, "Category")) & ":" & Html(GetDictString(data, "CategoryName")) & "</small>"
    End Function

    Private Function Html(value As Object) As String
        If value Is Nothing Then Return ""
        Return HttpUtility.HtmlEncode(value.ToString())
    End Function

    Private Sub ValidateMappedItem(item As SwitchUploadDbItem, errors As List(Of String))
        If item Is Nothing Then
            errors.Add("Invalid row data")
            Return
        End If

        If String.IsNullOrEmpty(item.ActionType) Then
            errors.Add("Missing Type")
        ElseIf Not item.ActionType.Equals("Switch", StringComparison.OrdinalIgnoreCase) AndAlso
               Not item.ActionType.Equals("Extra", StringComparison.OrdinalIgnoreCase) Then
            errors.Add("Invalid Type (use Switch or Extra)")
        End If

        If item.Amount <= 0D Then errors.Add("Invalid Amount")

        ValidateRequiredFields(item.FromData, "Source", errors)
        ValidatePeriodFields(item.FromData, "From", errors)
        ValidateMasterFields(item.FromData, "From", errors)

        If item.ActionType.Equals("Switch", StringComparison.OrdinalIgnoreCase) Then
            ValidateRequiredFields(item.ToData, "Destination", errors)
            ValidatePeriodFields(item.ToData, "To", errors)
            ValidateMasterFields(item.ToData, "To", errors)

            If HasSameSwitchKey(item.FromData, item.ToData) Then
                errors.Add("Source=Dest")
            End If
        End If
    End Sub

    Private Sub ValidateRequiredFields(data As Dictionary(Of String, Object), label As String, errors As List(Of String))
        Dim missing As New List(Of String)()
        Dim fields() As String = {"Year", "Month", "Company", "Category", "Segment", "Brand", "Vendor"}

        For Each fieldName In fields
            If String.IsNullOrEmpty(GetDictString(data, fieldName)) Then missing.Add(fieldName)
        Next

        If missing.Count > 0 Then errors.Add("Missing " & label & ": " & String.Join(", ", missing))
    End Sub

    Private Sub ValidatePeriodFields(data As Dictionary(Of String, Object), label As String, errors As List(Of String))
        If data Is Nothing Then Return

        Dim yearValue As String = GetDictString(data, "Year")
        Dim monthValue As String = GetDictString(data, "Month")
        Dim parsed As Integer

        If Not String.IsNullOrEmpty(yearValue) AndAlso Not Integer.TryParse(yearValue, parsed) Then
            errors.Add("Invalid " & label & " Year (" & yearValue & ")")
        End If

        If Not String.IsNullOrEmpty(monthValue) Then
            If Not Integer.TryParse(monthValue, parsed) OrElse parsed < 1 OrElse parsed > 12 Then
                errors.Add("Invalid " & label & " Month (" & monthValue & ")")
            End If
        End If
    End Sub

    Private Sub ValidateMasterFields(data As Dictionary(Of String, Object), label As String, errors As List(Of String))
        If data Is Nothing Then Return

        ValidateMasterField(data, "Month", dtMonth, "month_code", label & " Month", errors)
        ValidateMasterField(data, "Company", dtCompany, "CompanyCode", label & " Company", errors)
        ValidateMasterField(data, "Category", dtCategory, "Cate", label & " Category", errors)
        ValidateMasterField(data, "Segment", dtSegment, "SegmentCode", label & " Segment", errors)
        ValidateMasterField(data, "Brand", dtBrand, "Brand Code", label & " Brand", errors)
        ValidateMasterField(data, "Vendor", dtVendor, "VendorCode", label & " Vendor", errors)
    End Sub

    Private Sub ValidateMasterField(data As Dictionary(Of String, Object), fieldName As String, dt As DataTable, codeCol As String, label As String, errors As List(Of String))
        Dim value As String = GetDictString(data, fieldName)
        If Not String.IsNullOrEmpty(value) AndAlso Not IsValidMaster(dt, codeCol, value) Then
            errors.Add("Invalid " & label & " (" & value & ")")
        End If
    End Sub

    Private Function HasSameSwitchKey(fromData As Dictionary(Of String, Object), toData As Dictionary(Of String, Object)) As Boolean
        If fromData Is Nothing OrElse toData Is Nothing Then Return False

        Return GetDictString(fromData, "Year") = GetDictString(toData, "Year") AndAlso
               GetDictString(fromData, "Month") = GetDictString(toData, "Month") AndAlso
               GetDictString(fromData, "Company") = GetDictString(toData, "Company") AndAlso
               GetDictString(fromData, "Category") = GetDictString(toData, "Category") AndAlso
               GetDictString(fromData, "Segment") = GetDictString(toData, "Segment") AndAlso
               GetDictString(fromData, "Brand") = GetDictString(toData, "Brand") AndAlso
               GetDictString(fromData, "Vendor") = GetDictString(toData, "Vendor")
    End Function

    Private Sub ValidateDuplicateInBatch(item As SwitchUploadDbItem, uniqueRows As HashSet(Of String), errors As List(Of String))
        Dim rowKey As String = BuildRowKey(item)
        If uniqueRows.Contains(rowKey) Then
            errors.Add("Duplicate Data in File")
        Else
            uniqueRows.Add(rowKey)
        End If
    End Sub

    Private Function BuildRowKey(item As SwitchUploadDbItem) As String
        Dim amountKey As String = item.Amount.ToString("F2", CultureInfo.InvariantCulture)
        Dim f = item.FromData

        If item.ActionType.Equals("Switch", StringComparison.OrdinalIgnoreCase) Then
            Dim t = item.ToData
            Return String.Join("|", New String() {
                "SW", GetDictString(f, "Year"), GetDictString(f, "Month"), GetDictString(f, "Company"),
                GetDictString(f, "Category"), GetDictString(f, "Segment"), GetDictString(f, "Brand"),
                GetDictString(f, "Vendor"), item.TypeFrom, GetDictString(t, "Year"), GetDictString(t, "Month"),
                GetDictString(t, "Company"), GetDictString(t, "Category"), GetDictString(t, "Segment"),
                GetDictString(t, "Brand"), GetDictString(t, "Vendor"), item.TypeTo, amountKey
            })
        End If

        Return String.Join("|", New String() {
            "EX", GetDictString(f, "Year"), GetDictString(f, "Month"), GetDictString(f, "Company"),
            GetDictString(f, "Category"), GetDictString(f, "Segment"), GetDictString(f, "Brand"),
            GetDictString(f, "Vendor"), item.TypeFrom, amountKey
        })
    End Function

    Private Sub ValidateBudget(item As SwitchUploadDbItem,
                               budgetCalc As OTBBudgetCalculator,
                               povalidate As POValidate,
                               dbBudgetCache As Dictionary(Of String, Decimal),
                               dbUsedCache As Dictionary(Of String, Decimal),
                               pendingUsage As Dictionary(Of String, Decimal),
                               errors As List(Of String))
        If Not item.ActionType.Equals("Switch", StringComparison.OrdinalIgnoreCase) Then Return

        Try
            Dim f = item.FromData
            Dim key As String = String.Join("|", New String() {
                GetDictString(f, "Year"), GetDictString(f, "Month"), GetDictString(f, "Category"),
                GetDictString(f, "Company"), GetDictString(f, "Segment"), GetDictString(f, "Brand"), GetDictString(f, "Vendor")
            })

            If Not dbBudgetCache.ContainsKey(key) Then
                dbBudgetCache(key) = budgetCalc.CalculateCurrentApprovedBudget(
                    GetDictString(f, "Year"), GetDictString(f, "Month"), GetDictString(f, "Category"),
                    GetDictString(f, "Company"), GetDictString(f, "Segment"), GetDictString(f, "Brand"), GetDictString(f, "Vendor"))
                dbUsedCache(key) = povalidate.GetUsedBudgetFromDBForOTB(
                    GetDictString(f, "Year"), GetDictString(f, "Month"), GetDictString(f, "Category"),
                    GetDictString(f, "Company"), GetDictString(f, "Segment"), GetDictString(f, "Brand"), GetDictString(f, "Vendor"))
            End If

            Dim usedInBatch As Decimal = If(pendingUsage.ContainsKey(key), pendingUsage(key), 0D)
            Dim available As Decimal = (dbBudgetCache(key) - dbUsedCache(key)) - usedInBatch

            If available < item.Amount Then
                errors.Add($"Over Budget (Remaining: {available:N2}, Used in Batch: {usedInBatch:N2})")
            ElseIf pendingUsage.ContainsKey(key) Then
                pendingUsage(key) += item.Amount
            Else
                pendingUsage.Add(key, item.Amount)
            End If
        Catch ex As Exception
            errors.Add("Budget Error: " & ex.Message)
        End Try
    End Sub

    ' ==========================================
    ' 2. SAVE (BULK DB ONLY)
    ' ==========================================
    Private Sub HandleSave(context As HttpContext, uploadBy As String)
        context.Response.ContentType = "application/json"

        Try
            LoadAllMasterData()

            Dim json As String = context.Request.Form("data")
            Dim rows As List(Of Dictionary(Of String, Object)) = New JavaScriptSerializer().Deserialize(Of List(Of Dictionary(Of String, Object)))(json)
            If rows Is Nothing OrElse rows.Count = 0 Then Throw New Exception("No data found to process.")

            Dim pendingDbItems As List(Of SwitchUploadDbItem) = BuildDbItemsForSave(rows)

            Using conn As New SqlConnection(connectionString)
                conn.Open()
                Using dbTrans As SqlTransaction = conn.BeginTransaction()
                    Try
                        Using cmd As New SqlCommand("", conn, dbTrans)
                            For Each dbItem In pendingDbItems
                                SaveToDB(cmd, dbItem, uploadBy)
                            Next
                        End Using
                        dbTrans.Commit()
                    Catch ex As Exception
                        dbTrans.Rollback()
                        Throw New Exception("Database save failed: " & ex.Message)
                    End Try
                End Using
            End Using

            Dim processResults As New List(Of Object)
            For Each dbItem In pendingDbItems
                processResults.Add(New With {
                    .status = "Success",
                    .message = "Saved to database (SAP bypassed).",
                    .row = New With {
                        .Type = dbItem.TypeName,
                        .Function = dbItem.ActionType,
                        .Amount = dbItem.Amount,
                        .From = dbItem.FromData,
                        .To = dbItem.ToData
                    }
                })
            Next

            context.Response.Write(New JavaScriptSerializer().Serialize(New With {
                .success = True,
                .message = $"Upload & Save Completed Successfully. Saved {pendingDbItems.Count} record(s) to database. SAP was bypassed.",
                .results = processResults
            }))

        Catch ex As Exception
            context.Response.Write(New JavaScriptSerializer().Serialize(New With {.success = False, .message = ex.Message}))
        End Try
    End Sub

    Private Function BuildDbItemsForSave(rows As List(Of Dictionary(Of String, Object))) As List(Of SwitchUploadDbItem)
        Dim result As New List(Of SwitchUploadDbItem)()
        Dim uniqueRows As New HashSet(Of String)()
        Dim budgetCalc As New OTBBudgetCalculator()
        Dim povalidate As New POValidate()
        Dim dbBudgetCache As New Dictionary(Of String, Decimal)()
        Dim dbUsedCache As New Dictionary(Of String, Decimal)()
        Dim pendingUsage As New Dictionary(Of String, Decimal)()

        For i As Integer = 0 To rows.Count - 1
            Dim errors As New List(Of String)()
            Dim dbItem As SwitchUploadDbItem = MapPayloadToDbItem(rows(i))

            ValidateMappedItem(dbItem, errors)
            If errors.Count = 0 Then
                ValidateDuplicateInBatch(dbItem, uniqueRows, errors)
                If errors.Count = 0 Then ValidateBudget(dbItem, budgetCalc, povalidate, dbBudgetCache, dbUsedCache, pendingUsage, errors)
            End If

            If errors.Count > 0 Then
                Throw New Exception("Row " & (i + 2).ToString() & ": " & String.Join("; ", errors))
            End If

            AddMasterNames(dbItem.FromData)
            If dbItem.ToData IsNot Nothing Then AddMasterNames(dbItem.ToData)
            result.Add(dbItem)
        Next

        Return result
    End Function

    Private Function MapPayloadToDbItem(row As Dictionary(Of String, Object)) As SwitchUploadDbItem
        Dim actionType As String = NormalizeActionType(GetPayloadString(row, "Function"))
        If String.IsNullOrEmpty(actionType) Then actionType = NormalizePayloadType(GetPayloadString(row, "Type"))

        Dim amount As Decimal = 0D
        TryParseAmount(GetPayloadString(row, "Amount"), amount)

        Dim fromData As Dictionary(Of String, Object) = GetPayloadData(row, "From")
        Dim toData As Dictionary(Of String, Object) = Nothing
        If actionType.Equals("Switch", StringComparison.OrdinalIgnoreCase) Then toData = GetPayloadData(row, "To")

        Dim safeToData As Dictionary(Of String, Object) = toData
        If safeToData Is Nothing Then safeToData = New Dictionary(Of String, Object)()
        Dim typeInfo As SwitchTypeInfo = DetermineSwitchTypeInfo(actionType, GetDictString(fromData, "Year"), GetDictString(fromData, "Month"), GetDictString(safeToData, "Year"), GetDictString(safeToData, "Month"))

        Return New SwitchUploadDbItem With {
            .FromData = fromData,
            .ToData = toData,
            .Amount = amount,
            .TypeFrom = typeInfo.TypeFrom,
            .TypeTo = typeInfo.TypeTo,
            .Remark = GetPayloadString(row, "Remark"),
            .ActionType = actionType,
            .TypeName = typeInfo.DisplayName
        }
    End Function

    Private Function NormalizePayloadType(value As String) As String
        Dim normalized As String = NormalizeActionType(value)
        If normalized.Equals("Carry", StringComparison.OrdinalIgnoreCase) OrElse
           normalized.Equals("Balance", StringComparison.OrdinalIgnoreCase) Then
            Return "Switch"
        End If

        Return normalized
    End Function

    Private Sub SaveToDB(cmd As SqlCommand, item As SwitchUploadDbItem, user As String)
        cmd.CommandText = "
            MERGE dbo.OTB_Switching_Transaction AS T
            USING (
                SELECT
                    @Y AS [Year], @M AS [Month], @Co AS Company, @Ca AS Category,
                    @Se AS Segment, @Br AS Brand, @Ve AS Vendor, @TyF AS [From],
                    @Amt AS BudgetAmount, @SY AS SwitchYear, @SM AS SwitchMonth,
                    @SCo AS SwitchCompany, @SCa AS SwitchCategory, @SSe AS SwitchSegment,
                    @TyT AS [To], @SBr AS SwitchBrand, @SVe AS SwitchVendor
            ) AS S
            ON T.[Year] = S.[Year]
               AND T.[Month] = S.[Month]
               AND T.Company = S.Company
               AND T.Category = S.Category
               AND T.Segment = S.Segment
               AND T.Brand = S.Brand
               AND T.Vendor = S.Vendor
               AND T.[From] = S.[From]
               AND T.BudgetAmount = S.BudgetAmount
               AND ISNULL(T.SwitchYear, -1) = ISNULL(S.SwitchYear, -1)
               AND ISNULL(T.SwitchMonth, -1) = ISNULL(S.SwitchMonth, -1)
               AND ISNULL(T.SwitchCompany, N'') = ISNULL(S.SwitchCompany, N'')
               AND ISNULL(T.SwitchCategory, N'') = ISNULL(S.SwitchCategory, N'')
               AND ISNULL(T.SwitchSegment, N'') = ISNULL(S.SwitchSegment, N'')
               AND ISNULL(T.[To], N'') = ISNULL(S.[To], N'')
               AND ISNULL(T.SwitchBrand, N'') = ISNULL(S.SwitchBrand, N'')
               AND ISNULL(T.SwitchVendor, N'') = ISNULL(S.SwitchVendor, N'')
            WHEN MATCHED THEN
                UPDATE SET
                    T.[Release] = 0,
                    T.OTBStatus = N'Approved',
                    T.Batch = NULL,
                    T.Remark = @Rem,
                    T.CreateBy = COALESCE(T.CreateBy, @User),
                    T.ActionBy = @User
            WHEN NOT MATCHED BY TARGET THEN
                INSERT (
                    [Year], [Month], Company, Category, Segment, Brand, Vendor, [From], BudgetAmount, [Release],
                    SwitchYear, SwitchMonth, SwitchCompany, SwitchCategory, SwitchSegment, [To], SwitchBrand, SwitchVendor,
                    OTBStatus, Batch, Remark, CreateBy, CreateDT, ActionBy
                )
                VALUES (
                    S.[Year], S.[Month], S.Company, S.Category, S.Segment, S.Brand, S.Vendor, S.[From], S.BudgetAmount, 0,
                    S.SwitchYear, S.SwitchMonth, S.SwitchCompany, S.SwitchCategory, S.SwitchSegment, S.[To], S.SwitchBrand, S.SwitchVendor,
                    N'Approved', NULL, @Rem, @User, GETDATE(), @User
                );"

        cmd.Parameters.Clear()
        AddIntParameter(cmd, "@Y", GetDictString(item.FromData, "Year"))
        AddIntParameter(cmd, "@M", GetDictString(item.FromData, "Month"))
        AddStringParameter(cmd, "@Co", GetDictString(item.FromData, "Company"), 20)
        AddStringParameter(cmd, "@Ca", GetDictString(item.FromData, "Category"), 20)
        AddStringParameter(cmd, "@Se", GetDictString(item.FromData, "Segment"), 20)
        AddStringParameter(cmd, "@Br", GetDictString(item.FromData, "Brand"), 30)
        AddStringParameter(cmd, "@Ve", GetDictString(item.FromData, "Vendor"), 30)
        AddStringParameter(cmd, "@TyF", item.TypeFrom, 5)
        cmd.Parameters.Add("@Amt", SqlDbType.Decimal).Value = item.Amount
        cmd.Parameters("@Amt").Precision = 18
        cmd.Parameters("@Amt").Scale = 2
        AddStringParameter(cmd, "@User", user, 100)
        AddStringParameter(cmd, "@Rem", item.Remark, 500, True)

        If item.ActionType.Equals("Switch", StringComparison.OrdinalIgnoreCase) Then
            AddNullableIntParameter(cmd, "@SY", GetDictString(item.ToData, "Year"))
            AddNullableIntParameter(cmd, "@SM", GetDictString(item.ToData, "Month"))
            AddStringParameter(cmd, "@SCo", GetDictString(item.ToData, "Company"), 20)
            AddStringParameter(cmd, "@SCa", GetDictString(item.ToData, "Category"), 20)
            AddStringParameter(cmd, "@SSe", GetDictString(item.ToData, "Segment"), 20)
            AddStringParameter(cmd, "@TyT", item.TypeTo, 5)
            AddStringParameter(cmd, "@SBr", GetDictString(item.ToData, "Brand"), 30)
            AddStringParameter(cmd, "@SVe", GetDictString(item.ToData, "Vendor"), 30)
        Else
            AddNullableIntParameter(cmd, "@SY", "")
            AddNullableIntParameter(cmd, "@SM", "")
            AddStringParameter(cmd, "@SCo", Nothing, 20, True)
            AddStringParameter(cmd, "@SCa", Nothing, 20, True)
            AddStringParameter(cmd, "@SSe", Nothing, 20, True)
            AddStringParameter(cmd, "@TyT", Nothing, 5, True)
            AddStringParameter(cmd, "@SBr", Nothing, 30, True)
            AddStringParameter(cmd, "@SVe", Nothing, 30, True)
        End If

        cmd.ExecuteNonQuery()
    End Sub

    Private Sub AddIntParameter(cmd As SqlCommand, name As String, value As String)
        cmd.Parameters.Add(name, SqlDbType.Int).Value = Convert.ToInt32(value)
    End Sub

    Private Sub AddNullableIntParameter(cmd As SqlCommand, name As String, value As String)
        Dim parsed As Integer
        Dim parameter = cmd.Parameters.Add(name, SqlDbType.Int)
        If Integer.TryParse(value, parsed) Then
            parameter.Value = parsed
        Else
            parameter.Value = DBNull.Value
        End If
    End Sub

    Private Sub AddStringParameter(cmd As SqlCommand, name As String, value As String, size As Integer, Optional allowNull As Boolean = False)
        Dim parameter = cmd.Parameters.Add(name, SqlDbType.NVarChar, size)
        If allowNull AndAlso String.IsNullOrEmpty(value) Then
            parameter.Value = DBNull.Value
        Else
            parameter.Value = If(value Is Nothing, "", value)
        End If
    End Sub

    Private Function IsValidMaster(dt As DataTable, codeCol As String, codeVal As String) As Boolean
        If dt Is Nothing OrElse String.IsNullOrEmpty(codeVal) Then Return False
        Dim rows = dt.Select("[" & codeCol & "] = '" & codeVal.Replace("'", "''") & "'")
        Return rows.Length > 0
    End Function

    ReadOnly Property IsReusable() As Boolean Implements IHttpHandler.IsReusable
        Get
            Return False
        End Get
    End Property
End Class
