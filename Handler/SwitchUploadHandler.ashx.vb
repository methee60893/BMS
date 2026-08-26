Imports System
Imports System.Web
Imports System.Data
Imports System.Data.SqlClient
Imports System.IO
Imports ExcelDataReader
Imports System.Text
Imports System.Web.Script.Serialization
Imports System.Collections.Generic
Imports System.Threading.Tasks
Imports System.Globalization
Imports System.Linq
Imports System.Web.SessionState

Public Class SwitchUploadHandler
    Implements IHttpHandler, IReadOnlySessionState

    Private Shared connectionString As String = ConfigurationManager.ConnectionStrings("BMSConnectionString")?.ConnectionString
    Private Const MaxMovementAmount As Decimal = 9999999999999999.99D

    ' ตัวแปรสำหรับเก็บ Master Data
    Private dtYear As DataTable
    Private dtMonth As DataTable
    Private dtCompany As DataTable
    Private dtCategory As DataTable
    Private dtSegment As DataTable
    Private dtBrand As DataTable
    Private dtVendor As DataTable

    Public Sub ProcessRequest(ByVal context As HttpContext) Implements IHttpHandler.ProcessRequest
        context.Response.Clear()
        context.Response.ContentType = "application/json"
        context.Response.ContentEncoding = Encoding.UTF8

        Try
            EnsureAjaxMutationRequest(context)
            Dim uploadBy As String = EnsureSwitchEditPermission(context)
            Dim action As String = If(context.Request("action"), "").Trim()
            If action.Equals("preview", StringComparison.OrdinalIgnoreCase) Then
                HandlePreview(context)
            ElseIf action.Equals("save", StringComparison.OrdinalIgnoreCase) Then
                HandleSave(context, uploadBy)
            Else
                Throw New InvalidOperationException("Unsupported bulk OTB movement action.")
            End If
        Catch ex As Exception
            context.Response.StatusCode = 200
            context.Response.Write(New JavaScriptSerializer().Serialize(New With {.success = False, .message = ex.Message}))
        End Try
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
        Dim extension As String = Path.GetExtension(postedFile.FileName).ToLowerInvariant()
        If extension <> ".xlsx" AndAlso extension <> ".xls" Then
            context.Response.Write(New JavaScriptSerializer().Serialize(New With {.success = False, .message = "Only .xlsx and .xls files are supported."}))
            Return
        End If
        If postedFile.ContentLength <= 0 OrElse postedFile.ContentLength > 20 * 1024 * 1024 Then
            context.Response.Write(New JavaScriptSerializer().Serialize(New With {.success = False, .message = "The file must be between 1 byte and 20 MB."}))
            Return
        End If
        Dim tempPath As String = Nothing

        Try
            tempPath = Path.GetTempFileName()
            postedFile.SaveAs(tempPath)

            ' 1. โหลด Master Data เตรียมไว้
            LoadAllMasterData()

            ' 2. อ่าน Excel
            Dim dt As DataTable = ReadExcel(tempPath)
            If dt.Rows.Count <= 0 OrElse dt.Rows.Count > 15000 Then
                Throw New InvalidOperationException($"Bulk OTB files must contain between 1 and 15,000 data rows. Found: {dt.Rows.Count:N0}.")
            End If

            ' 3. สร้างตาราง Preview พร้อมชื่อ Description
            Dim result = GeneratePreviewData(dt)

            context.Response.Write(New JavaScriptSerializer().Serialize(result))
        Catch ex As Exception
            context.Response.Write(New JavaScriptSerializer().Serialize(New With {.success = False, .message = "Error: " & ex.Message}))
        Finally
            Try
                If Not String.IsNullOrWhiteSpace(tempPath) AndAlso File.Exists(tempPath) Then File.Delete(tempPath)
            Catch
                ' Preview result must not be replaced by a best-effort temp-file cleanup error.
            End Try
        End Try
    End Sub

    Private Sub LoadAllMasterData()
        Using conn As New SqlConnection(connectionString)
            conn.Open()

            dtYear = New DataTable()
            Using cmd As New SqlCommand("SELECT [Year_Kept] FROM [MS_Year]", conn)
                dtYear.Load(cmd.ExecuteReader())
            End Using

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
        Dim rows = dt.Select($"[{codeCol}] = '{codeVal.Replace("'", "''")}'")
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
        Dim hasAnyWarnings As Boolean = False

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
            Dim warnings As New List(Of String)()

            ' --- 1. Read Data ---
            Dim _function As String = GetVal(row, "Function")
            Dim type As String = ""
            Dim amtStr As String = GetVal(row, "Amount")
            Dim remark As String = GetVal(row, "Remark")

            ' From
            Dim fYear As String = GetVal(row, "Year")
            Dim fMonth As String = GetVal(row, "Month")
            Dim fComp As String = GetVal(row, "Company")
            Dim fCate As String = GetVal(row, "Category")
            Dim fSeg As String = GetVal(row, "Segment")
            Dim fBrand As String = GetVal(row, "Brand")
            Dim fVend As String = GetVal(row, "Vendor")

            ' To
            Dim tYear As String = GetVal(row, "To_Year")
            Dim tMonth As String = GetVal(row, "To_Month")
            Dim tComp As String = GetVal(row, "To_Company")
            Dim tCate As String = GetVal(row, "To_Category")
            Dim tSeg As String = GetVal(row, "To_Segment")
            Dim tBrand As String = GetVal(row, "To_Brand")
            Dim tVend As String = GetVal(row, "To_Vendor")

            ' --- 2. Validation ---
            Dim amount As Decimal = 0
            Dim isAmountValid As Boolean = False
            Dim hasFractionalAmount As Boolean = False
            Dim hasPeriodWarning As Boolean = False
            If String.IsNullOrEmpty(_function) Then errors.Add("Missing Type")
            If Not String.IsNullOrEmpty(_function) AndAlso
               Not _function.Equals("Switch", StringComparison.OrdinalIgnoreCase) AndAlso
               Not _function.Equals("Extra", StringComparison.OrdinalIgnoreCase) Then
                errors.Add("Invalid Type (use Switch or Extra)")
            End If
            Dim rawAmount As Decimal = 0
            If Decimal.TryParse(amtStr, NumberStyles.Any, CultureInfo.InvariantCulture, rawAmount) OrElse
               Decimal.TryParse(amtStr, NumberStyles.Any, CultureInfo.CurrentCulture, rawAmount) Then
                amount = share_class.RoundAmountForSapRule(rawAmount)
                hasFractionalAmount = (_function.Equals("Switch", StringComparison.OrdinalIgnoreCase) AndAlso
                                       rawAmount <> Decimal.Truncate(rawAmount))
                If amount <= 0 Then
                    errors.Add("Invalid Amount")
                ElseIf amount > MaxMovementAmount Then
                    errors.Add("Amount exceeds the database limit")
                Else
                    isAmountValid = True
                End If
            Else
                errors.Add("Invalid Amount")
            End If

            If hasFractionalAmount Then
                warnings.Add("ยอดมีทศนิยม กรุณาตรวจสอบ")
            End If
            If remark.Length > 500 Then errors.Add("Remark exceeds 500 characters")
            If String.IsNullOrEmpty(fYear) OrElse String.IsNullOrEmpty(fMonth) OrElse String.IsNullOrEmpty(fComp) OrElse String.IsNullOrEmpty(fCate) OrElse String.IsNullOrEmpty(fSeg) OrElse String.IsNullOrEmpty(fBrand) OrElse String.IsNullOrEmpty(fVend) Then errors.Add("Missing Source")

            If Not String.IsNullOrEmpty(fYear) AndAlso Not IsValidMaster(dtYear, "Year_Kept", fYear) Then errors.Add($"Invalid From Year ({fYear})")
            If Not String.IsNullOrEmpty(fMonth) AndAlso Not IsValidMaster(dtMonth, "month_code", fMonth) Then errors.Add($"Invalid From Month ({fMonth})")
            If Not String.IsNullOrEmpty(fComp) AndAlso Not IsValidMaster(dtCompany, "CompanyCode", fComp) Then errors.Add($"Invalid From Company ({fComp})")
            If Not String.IsNullOrEmpty(fCate) AndAlso Not IsValidMaster(dtCategory, "Cate", fCate) Then errors.Add($"Invalid From Category ({fCate})")
            If Not String.IsNullOrEmpty(fSeg) AndAlso Not IsValidMaster(dtSegment, "SegmentCode", fSeg) Then errors.Add($"Invalid From Segment ({fSeg})")
            If Not String.IsNullOrEmpty(fBrand) AndAlso Not IsValidMaster(dtBrand, "Brand Code", fBrand) Then errors.Add($"Invalid From Brand ({fBrand})")
            If Not String.IsNullOrEmpty(fVend) AndAlso Not IsValidMaster(dtVendor, "VendorCode", fVend) Then errors.Add($"Invalid From Vendor ({fVend})")

            If _function.Equals("Switch", StringComparison.OrdinalIgnoreCase) Then
                If String.IsNullOrEmpty(tYear) OrElse String.IsNullOrEmpty(tMonth) OrElse String.IsNullOrEmpty(tComp) OrElse String.IsNullOrEmpty(tCate) OrElse String.IsNullOrEmpty(tSeg) OrElse String.IsNullOrEmpty(tBrand) OrElse String.IsNullOrEmpty(tVend) Then errors.Add("Missing Dest")
                If fYear = tYear AndAlso fMonth = tMonth AndAlso fComp = tComp AndAlso fCate = tCate AndAlso fSeg = tSeg AndAlso fBrand = tBrand AndAlso fVend = tVend Then
                    errors.Add("Source=Dest")
                End If

                If Not String.IsNullOrEmpty(tYear) AndAlso Not IsValidMaster(dtYear, "Year_Kept", tYear) Then errors.Add($"Invalid To Year ({tYear})")
                If Not String.IsNullOrEmpty(tMonth) AndAlso Not IsValidMaster(dtMonth, "month_code", tMonth) Then errors.Add($"Invalid To Month ({tMonth})")
                If Not String.IsNullOrEmpty(tComp) AndAlso Not IsValidMaster(dtCompany, "CompanyCode", tComp) Then errors.Add($"Invalid To Company ({tComp})")
                If Not String.IsNullOrEmpty(tCate) AndAlso Not IsValidMaster(dtCategory, "Cate", tCate) Then errors.Add($"Invalid To Category ({tCate})")
                If Not String.IsNullOrEmpty(tSeg) AndAlso Not IsValidMaster(dtSegment, "SegmentCode", tSeg) Then errors.Add($"Invalid To Segment ({tSeg})")
                If Not String.IsNullOrEmpty(tBrand) AndAlso Not IsValidMaster(dtBrand, "Brand Code", tBrand) Then errors.Add($"Invalid To Brand ({tBrand})")
                If Not String.IsNullOrEmpty(tVend) AndAlso Not IsValidMaster(dtVendor, "VendorCode", tVend) Then errors.Add($"Invalid To Vendor ({tVend})")

                If Not String.IsNullOrEmpty(fYear) AndAlso Not String.IsNullOrEmpty(fMonth) AndAlso
                   Not String.IsNullOrEmpty(tYear) AndAlso Not String.IsNullOrEmpty(tMonth) Then
                    Dim fYearNumber As Integer
                    Dim fMonthNumber As Integer
                    Dim tYearNumber As Integer
                    Dim tMonthNumber As Integer
                    If Integer.TryParse(fYear, fYearNumber) AndAlso Integer.TryParse(fMonth, fMonthNumber) AndAlso
                       Integer.TryParse(tYear, tYearNumber) AndAlso Integer.TryParse(tMonth, tMonthNumber) Then
                        hasPeriodWarning = (fYearNumber <> tYearNumber OrElse fMonthNumber <> tMonthNumber)
                    Else
                        hasPeriodWarning = (Not fYear.Equals(tYear, StringComparison.OrdinalIgnoreCase) OrElse
                                            Not fMonth.Equals(tMonth, StringComparison.OrdinalIgnoreCase))
                    End If
                End If

                If hasPeriodWarning Then
                    warnings.Add("เดือนไม่สอดคล้องกัน")
                End If
            End If

            Dim rowKey As String = ""
            If _function.Equals("Switch", StringComparison.OrdinalIgnoreCase) Then
                rowKey = $"SW|{fYear}|{fMonth}|{fComp}|{fCate}|{fSeg}|{fBrand}|{fVend}|{tYear}|{tMonth}|{tComp}|{tCate}|{tSeg}|{tBrand}|{tVend}|{amount}"
            Else
                ' Extra
                rowKey = $"EX|{fYear}|{fMonth}|{fComp}|{fCate}|{fSeg}|{fBrand}|{fVend}"
            End If

            If uniqueRows.Contains(rowKey) Then
                errors.Add("Duplicate Data in File")
            Else
                uniqueRows.Add(rowKey)
            End If

            ' --- 3. Budget Check ---
            If errors.Count = 0 AndAlso isAmountValid AndAlso _function.Equals("Switch", StringComparison.OrdinalIgnoreCase) Then
                Try
                    Dim key As String = $"{fYear}|{fMonth}|{fCate}|{fComp}|{fSeg}|{fBrand}|{fVend}"

                    If Not dbBudgetCache.ContainsKey(key) Then
                        dbBudgetCache(key) = budgetCalc.CalculateCurrentApprovedBudget(fYear, fMonth, fCate, fComp, fSeg, fBrand, fVend)
                        dbUsedCache(key) = povalidate.GetUsedBudgetFromDBForOTB(fYear, fMonth, fCate, fComp, fSeg, fBrand, fVend)
                    End If

                    Dim currentDbBudget As Decimal = dbBudgetCache(key)
                    Dim usedInDB As Decimal = dbUsedCache(key)

                    Dim usedInBatch As Decimal = If(pendingUsage.ContainsKey(key), pendingUsage(key), 0)
                    Dim available As Decimal = (currentDbBudget - usedInDB) - usedInBatch

                    If available < amount Then
                        errors.Add($"Over Budget (Remaining: {available:N2}, Used in Batch: {usedInBatch:N2})")
                    Else
                        If pendingUsage.ContainsKey(key) Then pendingUsage(key) += amount Else pendingUsage.Add(key, amount)
                    End If
                Catch ex As Exception
                    errors.Add("Budget Error: " & ex.Message)
                End Try
            End If

            Dim isError As Boolean = (errors.Count > 0)
            If isError Then allValid = False
            If warnings.Count > 0 Then hasAnyWarnings = True



            ' --- 4. Serialize Data for Save ---
            Dim rowDataObj As New Dictionary(Of String, Object) From {
                {"Function", _function}, {"Amount", amount}, {"Remark", remark},
                {"From", New Dictionary(Of String, String) From {{"Year", fYear}, {"Month", fMonth}, {"Company", fComp}, {"Category", fCate}, {"Segment", fSeg}, {"Brand", fBrand}, {"Vendor", fVend}}},
                {"To", New Dictionary(Of String, String) From {{"Year", tYear}, {"Month", tMonth}, {"Company", tComp}, {"Category", tCate}, {"Segment", tSeg}, {"Brand", tBrand}, {"Vendor", tVend}}}
            }
            Dim rowJson As String = If(isError, "", HttpUtility.HtmlAttributeEncode(New JavaScriptSerializer().Serialize(rowDataObj)))

            Dim typeName As String = ""
            If _function.Equals("Switch", StringComparison.OrdinalIgnoreCase) Then
                typeName = "Switch"
                Dim fYearNumber As Integer
                Dim fMonthNumber As Integer
                Dim tYearNumber As Integer
                Dim tMonthNumber As Integer
                If Integer.TryParse(fYear, fYearNumber) AndAlso Integer.TryParse(fMonth, fMonthNumber) AndAlso
                   Integer.TryParse(tYear, tYearNumber) AndAlso Integer.TryParse(tMonth, tMonthNumber) AndAlso
                   fMonthNumber >= 1 AndAlso fMonthNumber <= 12 AndAlso tMonthNumber >= 1 AndAlso tMonthNumber <= 12 Then
                    Dim dFrom As New Date(fYearNumber, fMonthNumber, 1)
                    Dim dTo As New Date(tYearNumber, tMonthNumber, 1)
                    If dFrom > dTo Then typeName = "Carry"
                    If dFrom < dTo Then typeName = "Balance"
                End If
            ElseIf _function.Equals("Extra", StringComparison.OrdinalIgnoreCase) Then
                typeName = "Extra"
            End If


            ' --- 5. Prepare Display Strings (With Names) ---
            Dim fCompName As String = GetName(dtCompany, "CompanyCode", "CompanyNameShort", fComp)
            Dim fCateName As String = GetName(dtCategory, "Cate", "Category", fCate)
            Dim fSegName As String = GetName(dtSegment, "SegmentCode", "SegmentName", fSeg)
            Dim fBrandName As String = GetName(dtBrand, "Brand Code", "Brand Name", fBrand)
            Dim fVendName As String = GetName(dtVendor, "VendorCode", "Vendor", fVend)

            ' HTML Block for From
            Dim fHtml As String = $"<strong>{HttpUtility.HtmlEncode(fYear)}/{HttpUtility.HtmlEncode(fMonth)}</strong><br/>" &
                                  $"<small>Comp: {HttpUtility.HtmlEncode(fComp)}:{HttpUtility.HtmlEncode(fCompName)}<br/>" &
                                  $"Vend: {HttpUtility.HtmlEncode(fVend)}:{HttpUtility.HtmlEncode(fVendName)}<br/>" &
                                  $"Brand: {HttpUtility.HtmlEncode(fBrand)}:{HttpUtility.HtmlEncode(fBrandName)}<br/>" &
                                  $"Seg: {HttpUtility.HtmlEncode(fSeg)}:{HttpUtility.HtmlEncode(fSegName)} | " &
                                  $"Cat: {HttpUtility.HtmlEncode(fCate)}:{HttpUtility.HtmlEncode(fCateName)}</small>"

            Dim tHtml As String = "-"
            If _function.Equals("Switch", StringComparison.OrdinalIgnoreCase) Then
                Dim tCompName As String = GetName(dtCompany, "CompanyCode", "CompanyNameShort", tComp)
                Dim tCateName As String = GetName(dtCategory, "Cate", "Category", tCate)
                Dim tSegName As String = GetName(dtSegment, "SegmentCode", "SegmentName", tSeg)
                Dim tBrandName As String = GetName(dtBrand, "Brand Code", "Brand Name", tBrand)
                Dim tVendName As String = GetName(dtVendor, "VendorCode", "Vendor", tVend)

                tHtml = $"<strong>{HttpUtility.HtmlEncode(tYear)}/{HttpUtility.HtmlEncode(tMonth)}</strong><br/>" &
                        $"<small>Comp: {HttpUtility.HtmlEncode(tComp)}:{HttpUtility.HtmlEncode(tCompName)}<br/>" &
                        $"Vend: {HttpUtility.HtmlEncode(tVend)}:{HttpUtility.HtmlEncode(tVendName)}<br/>" &
                        $"Brand: {HttpUtility.HtmlEncode(tBrand)}:{HttpUtility.HtmlEncode(tBrandName)}<br/>" &
                        $"Seg: {HttpUtility.HtmlEncode(tSeg)}:{HttpUtility.HtmlEncode(tSegName)} | " &
                        $"Cat: {HttpUtility.HtmlEncode(tCate)}:{HttpUtility.HtmlEncode(tCateName)}</small>"
            End If

            ' --- 6. Render Row ---
            Dim hasWarning As Boolean = (warnings.Count > 0)
            Dim trClass As String = If(isError, "table-danger", If(hasWarning, "table-warning", ""))
            Dim periodWarningCellClass As String = If(hasPeriodWarning, "table-warning", "")
            Dim amountWarningCellClass As String = If(hasFractionalAmount, "table-warning", "")
            sb.AppendFormat("<tr class='{0}'>", trClass)
            sb.AppendFormat("<td class='text-center'>{0}<input type='hidden' class='row-data' value='{1}'></td>", i + 2, rowJson)

            sb.AppendFormat("<td class='text-center fw-bold'>{0}</td>", HttpUtility.HtmlEncode(_function))
            sb.AppendFormat("<td class='text-center fw-bold'>{0}</td>", HttpUtility.HtmlEncode(typeName))

            ' From Column
            sb.AppendFormat("<td class='{0}'>{1}</td>", periodWarningCellClass, fHtml)

            ' Amount Column
            sb.AppendFormat("<td class='text-end fw-bold {0}'>{1:N2}</td>", amountWarningCellClass, amount)

            ' To Column
            sb.AppendFormat("<td class='{0}'>{1}</td>", periodWarningCellClass, tHtml)

            ' Remark Column
            sb.AppendFormat("<td><small>{0}</small></td>", HttpUtility.HtmlEncode(remark))

            ' Status & Error
            Dim status As String
            If isError Then
                status = "<span class='badge bg-danger'>Fail</span>"
            ElseIf hasWarning Then
                status = "<span class='badge bg-warning text-dark'>Warning</span>"
            Else
                status = "<span class='badge bg-success'>Pass</span>"
            End If
            Dim encodedErrors As New List(Of String)()
            For Each errorMessage In errors
                encodedErrors.Add(HttpUtility.HtmlEncode(errorMessage))
            Next
            Dim encodedWarnings As New List(Of String)()
            For Each warningMessage In warnings
                encodedWarnings.Add(HttpUtility.HtmlEncode(warningMessage))
            Next
            Dim issueParts As New List(Of String)()
            If encodedErrors.Count > 0 Then issueParts.Add("<span class='text-danger'>" & String.Join("<br/>", encodedErrors) & "</span>")
            If encodedWarnings.Count > 0 Then issueParts.Add("<span class='text-warning fw-semibold'>" & String.Join("<br/>", encodedWarnings) & "</span>")
            sb.AppendFormat("<td class='text-center'>{0}</td><td class='small'>{1}</td>", status, String.Join("<br/>", issueParts))
            sb.Append("</tr>")
        Next
        sb.Append("</tbody></table></div>")

        If Not allValid Then
            sb.Append("<div class='alert alert-warning mt-2'><i class='bi bi-exclamation-triangle'></i> ข้อมูลบางรายการไม่ถูกต้อง กรุณาแก้ไขไฟล์แล้ว Upload ใหม่ (ต้องผ่านทุกรายการถึงจะบันทึกได้)</div>")
        ElseIf hasAnyWarnings Then
            sb.Append("<div class='alert alert-warning mt-2'><i class='bi bi-exclamation-triangle'></i> ข้อมูลผ่านการตรวจสอบและพร้อมบันทึก กรุณาตรวจสอบรายการสีเหลืองก่อนส่ง SAP</div>")
        Else
            sb.Append("<div class='alert alert-success mt-2'><i class='bi bi-check-circle'></i> ข้อมูลถูกต้องครบถ้วน พร้อมบันทึกไปยัง SAP</div>")
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

    ' ==========================================
    ' 2. SAVE (BULK SAP & DB) - UPDATED VERSION
    ' ==========================================
    Private Sub HandleSave(context As HttpContext, uploadBy As String)
        context.Response.ContentType = "application/json"

        Try
            ' 0. โหลด Master Data เตรียมไว้สำหรับดึงชื่อ (CompanyName, VendorName, ฯลฯ)
            LoadAllMasterData()

            Dim json As String = context.Request.Form("data")
            Dim rows As List(Of Dictionary(Of String, Object)) = New JavaScriptSerializer().Deserialize(Of List(Of Dictionary(Of String, Object)))(json)
            If rows Is Nothing OrElse rows.Count = 0 OrElse rows.Count > 15000 Then
                Throw New InvalidOperationException("Bulk OTB save requires between 1 and 15,000 rows.")
            End If

            ' Revalidate the submitted JSON against live PO usage. Preview data is not trusted here.
            ' Source locks serialize concurrent single/bulk saves for the same OTB dimension.
            Using budgetLockConn As New SqlConnection(connectionString)
                budgetLockConn.Open()
                Dim affectedDimensions As New List(Of OTBSwitchBudgetGuard.BudgetDimension)()
                Dim submittedKeys As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
                For Each submittedRow In rows
                    If submittedRow Is Nothing Then Throw New InvalidOperationException("Every submitted row must be an object.")
                    Dim functionValue As Object = Nothing
                    If Not submittedRow.TryGetValue("Function", functionValue) Then Throw New InvalidOperationException("Every row must include Function.")
                    Dim functionName As String = Convert.ToString(functionValue).Trim()
                    If Not functionName.Equals("Switch", StringComparison.OrdinalIgnoreCase) AndAlso
                       Not functionName.Equals("Extra", StringComparison.OrdinalIgnoreCase) Then
                        Throw New InvalidOperationException("Function must be Switch or Extra. SAP was not called.")
                    End If
                    Dim amountValue As Object = Nothing
                    If Not submittedRow.TryGetValue("Amount", amountValue) Then Throw New InvalidOperationException("Every row must include Amount.")
                    Dim submittedAmount As Decimal = share_class.ParseAndRoundAmount(Convert.ToString(amountValue, CultureInfo.InvariantCulture))
                    Dim submittedRemark As String = ""
                    If submittedRow.ContainsKey("Remark") Then submittedRemark = Convert.ToString(submittedRow("Remark"))
                    ValidateMovementPersistenceFields(submittedAmount, submittedRemark)

                    Dim sourceValue As Object = Nothing
                    If Not submittedRow.TryGetValue("From", sourceValue) Then Throw New InvalidOperationException("Every Switch/Extra row must include a source budget dimension.")
                    Dim source = TryCast(sourceValue, Dictionary(Of String, Object))
                    Dim sourceDimension As OTBSwitchBudgetGuard.BudgetDimension = ParseSubmittedDimension(source, "source")
                    affectedDimensions.Add(sourceDimension)

                    Dim submittedKey As String = functionName.ToUpperInvariant() & "|" & OTBSwitchBudgetGuard.BuildDetailLockResource(sourceDimension)
                    If functionName.Equals("Switch", StringComparison.OrdinalIgnoreCase) Then
                        Dim destinationValue As Object = Nothing
                        If Not submittedRow.TryGetValue("To", destinationValue) Then Throw New InvalidOperationException("Every Switch row must include a destination budget dimension.")
                        Dim destination = TryCast(destinationValue, Dictionary(Of String, Object))
                        If destination Is Nothing Then Throw New InvalidOperationException("Every Switch row must include a destination budget dimension.")
                        Dim destinationDimension As OTBSwitchBudgetGuard.BudgetDimension = ParseSubmittedDimension(destination, "destination")
                        If String.Equals(OTBSwitchBudgetGuard.BuildDetailLockResource(sourceDimension),
                                         OTBSwitchBudgetGuard.BuildDetailLockResource(destinationDimension),
                                         StringComparison.Ordinal) Then
                            Throw New InvalidOperationException("Source and destination budget dimensions must be different.")
                        End If
                        affectedDimensions.Add(destinationDimension)
                        submittedKey &= "|" & OTBSwitchBudgetGuard.BuildDetailLockResource(destinationDimension) & "|" & submittedAmount.ToString(CultureInfo.InvariantCulture)
                    End If
                    If Not submittedKeys.Add(submittedKey) Then
                        Throw New InvalidOperationException("Duplicate movement rows were submitted. SAP was not called.")
                    End If
                Next
                Using budgetLockHandle As IDisposable = OTBSwitchBudgetGuard.AcquireBudgetLocks(
                    budgetLockConn, OTBSwitchBudgetGuard.BuildGroupLockResources(affectedDimensions))
                    OTBSwitchBudgetGuard.EnsureValidMasterDimensions(budgetLockConn, affectedDimensions)
                    OTBSwitchBudgetGuard.EnsureNoActiveApprovalClaims(budgetLockConn, affectedDimensions)

                Dim liveCalculator As New OTBBudgetCalculator()
                Dim livePOValidator As New POValidate()
                Dim batchUsage As New Dictionary(Of String, Decimal)(StringComparer.OrdinalIgnoreCase)
                For Each submittedRow In rows
                    If submittedRow("Function").ToString().Equals("Switch", StringComparison.OrdinalIgnoreCase) Then
                        Dim source = TryCast(submittedRow("From"), Dictionary(Of String, Object))
                        Dim submittedAmount = share_class.ParseAndRoundAmount(Convert.ToString(submittedRow("Amount"), CultureInfo.InvariantCulture))
                        Dim sourceKey = OTBSwitchBudgetGuard.BuildSourceKey(
                            source("Year").ToString(), source("Month").ToString(), source("Company").ToString(),
                            source("Category").ToString(), source("Segment").ToString(), source("Brand").ToString(), source("Vendor").ToString())
                        Dim alreadyRequested = If(batchUsage.ContainsKey(sourceKey), batchUsage(sourceKey), 0D)
                        Dim budgetCheck = OTBSwitchBudgetGuard.Check(
                            source("Year").ToString(), source("Month").ToString(), source("Category").ToString(), source("Company").ToString(),
                            source("Segment").ToString(), source("Brand").ToString(), source("Vendor").ToString(), liveCalculator, livePOValidator)
                        budgetCheck.AvailableBudget -= alreadyRequested
                        OTBSwitchBudgetGuard.EnsureSufficient(budgetCheck, submittedAmount)
                        batchUsage(sourceKey) = alreadyRequested + submittedAmount
                    End If
                Next

            Dim sapRequest As New OtbSwitchRequest()
            sapRequest.TestMode = ""

            Dim pendingDbInserts As New List(Of Dictionary(Of String, Object))

            For Each row In rows
                Dim _func As String = row("Function").ToString()
                Dim amt As Decimal = share_class.ParseAndRoundAmount(Convert.ToString(row("Amount"), CultureInfo.InvariantCulture))
                Dim f = TryCast(row("From"), Dictionary(Of String, Object))
                Dim t = TryCast(row("To"), Dictionary(Of String, Object))

                ' --- Logic กำหนด TypeCode (SAP) และ Display Name ---
                Dim fromCode As String = "E" ' Default Extra
                Dim toCode As String = ""
                Dim typeNameDisplay As String = "Extra"

                If _func.Equals("Switch", StringComparison.OrdinalIgnoreCase) Then
                    fromCode = "D" : toCode = "C" : typeNameDisplay = "Switch"

                    Dim dFrom As New Date(Convert.ToInt32(f("Year")), Convert.ToInt32(f("Month")), 1)
                    Dim dTo As New Date(Convert.ToInt32(t("Year")), Convert.ToInt32(t("Month")), 1)

                    If dFrom > dTo Then
                        fromCode = "G" : toCode = "F" : typeNameDisplay = "Carry"
                    ElseIf dFrom < dTo Then
                        fromCode = "I" : toCode = "H" : typeNameDisplay = "Balance"
                    End If
                End If

                ' --- เตรียม SAP Item ---
                Dim item As New OtbSwitchItem With {
                    .DocYearFrom = f("Year").ToString(), .PeriodFrom = f("Month").ToString(), .FmAreaFrom = f("Company").ToString(),
                    .CatFrom = f("Category").ToString(), .SegmentFrom = f("Segment").ToString(), .BrandFrom = f("Brand").ToString(), .VendorFrom = f("Vendor").ToString(),
                    .Budget = share_class.FormatAmountForSap(amt), .TypeFrom = fromCode
                }

                ' --- Mapping ข้อมูลชื่อฝั่ง From (ต้องมีทุกประเภทรายการ) ---
                f("CompanyName") = GetName(dtCompany, "CompanyCode", "CompanyNameShort", f("Company").ToString())
                f("VendorName") = GetName(dtVendor, "VendorCode", "Vendor", f("Vendor").ToString())
                f("CategoryName") = GetName(dtCategory, "Cate", "Category", f("Category").ToString())
                f("SegmentName") = GetName(dtSegment, "SegmentCode", "SegmentName", f("Segment").ToString())
                f("BrandName") = GetName(dtBrand, "Brand Code", "Brand Name", f("Brand").ToString())

                ' --- จัดการฝั่ง To เฉพาะรายการ Switch เท่านั้น ---
                If _func.Equals("Switch", StringComparison.OrdinalIgnoreCase) AndAlso t IsNot Nothing Then
                    item.DocYearTo = t("Year").ToString() : item.PeriodTo = t("Month").ToString() : item.FmAreaTo = t("Company").ToString()
                    item.CatTo = t("Category").ToString() : item.SegmentTo = t("Segment").ToString() : item.BrandTo = t("Brand").ToString() : item.VendorTo = t("Vendor").ToString()
                    item.TypeTo = toCode

                    ' Mapping ชื่อฝั่ง To
                    t("CompanyName") = GetName(dtCompany, "CompanyCode", "CompanyNameShort", t("Company").ToString())
                    t("VendorName") = GetName(dtVendor, "VendorCode", "Vendor", t("Vendor").ToString())
                    t("CategoryName") = GetName(dtCategory, "Cate", "Category", t("Category").ToString())
                    t("SegmentName") = GetName(dtSegment, "SegmentCode", "SegmentName", t("Segment").ToString())
                    t("BrandName") = GetName(dtBrand, "Brand Code", "Brand Name", t("Brand").ToString())
                Else
                    ' ถ้าเป็น Extra ให้ฝั่ง To เป็น Nothing เพื่อความชัดเจนในการส่ง JSON
                    t = Nothing
                End If

                sapRequest.Data.Add(item)

                ' เก็บข้อมูลเพื่อบันทึกลง DB และส่งกลับ UI
                Dim dbRow As New Dictionary(Of String, Object) From {
                    {"f", f}, {"t", t}, {"amt", amt}, {"typeFrom", fromCode},
                    {"remark", If(row.ContainsKey("Remark"), Convert.ToString(row("Remark")), "")}, {"actionType", _func}, {"typeName", typeNameDisplay}
                }
                pendingDbInserts.Add(dbRow)
            Next

            ' 1. ส่ง SAP API
            Dim sapRes = Task.Run(Async Function() Await SapApiHelper.SwitchOtbPlanAsync(sapRequest)).Result
            If sapRes Is Nothing Then Throw New Exception("No response from SAP API.")
            If sapRes.Status Is Nothing Then Throw New Exception("SAP response did not include status.")
            If sapRes.Results Is Nothing Then Throw New Exception("SAP response did not include results.")

            ' 2. จัดการ Mapping ผลลัพธ์ส่งกลับไปหน้าจอ
            Dim processResults As New List(Of Object)
            Dim hasError As Boolean = False
            Dim sapSummaryMessage As String = ""

            If sapRes.Status.Total <> pendingDbInserts.Count OrElse
               sapRes.Status.Success <> pendingDbInserts.Count OrElse
               sapRes.Status.ErrorCount > 0 OrElse
               sapRes.Results.Count <> pendingDbInserts.Count Then
                hasError = True
                sapSummaryMessage = $"SAP processing failed or was incomplete. Total: {sapRes.Status.Total}, Success: {sapRes.Status.Success}, Error: {sapRes.Status.ErrorCount}. No records were saved to the database."
            End If

            For i As Integer = 0 To pendingDbInserts.Count - 1
                Dim status As String = "Success"
                Dim msg As String = "Success"

                If i < sapRes.Results.Count Then
                    Dim resItem = sapRes.Results(i)
                    Dim messageType As String = If(resItem.MessageType, "").Trim()
                    If Not messageType.Equals("S", StringComparison.OrdinalIgnoreCase) Then
                        status = "Error"
                        hasError = True
                    End If
                    msg = If(String.IsNullOrWhiteSpace(resItem.Message), $"SAP MessageType: {messageType}", resItem.Message)
                Else
                    status = "Error"
                    msg = "SAP did not return a result for this row."
                    hasError = True
                End If

                ' สร้างโครงสร้าง JSON ที่สมบูรณ์ส่งกลับไปให้ JavaScript
                processResults.Add(New With {
                    .status = status,
                    .message = msg,
                    .row = New With {
                        .Type = pendingDbInserts(i)("typeName"),
                        .Function = pendingDbInserts(i)("actionType"),
                        .Amount = pendingDbInserts(i)("amt"),
                        .From = pendingDbInserts(i)("f"),
                        .To = pendingDbInserts(i)("t") ' ถ้าเป็น Extra ค่านี้จะเป็น null โดยอัตโนมัติ
                    }
                })
            Next

            ' 3. บันทึกฐานข้อมูล (กรณีไม่มี Error จาก SAP เลย)
            If Not hasError Then
                Using conn As New SqlConnection(connectionString)
                    conn.Open()
                    Using dbTrans As SqlTransaction = conn.BeginTransaction()
                        Try
                            For Each dbItem In pendingDbInserts
                                InsertToDB(cmd:=New SqlCommand("", conn, dbTrans),
                                           f:=dbItem("f"), t:=dbItem("t"), amt:=dbItem("amt"),
                                           typeFrom:=dbItem("typeFrom"), user:=uploadBy, rmk:=dbItem("remark"),
                                           actionType:=dbItem("actionType"))
                            Next
                            dbTrans.Commit()
                        Catch ex As Exception
                            dbTrans.Rollback() : Throw New Exception("Database Insert Failed: " & ex.Message)
                        End Try
                    End Using
                End Using
            End If

            ' 4. ส่ง JSON Response
            context.Response.Write(New JavaScriptSerializer().Serialize(New With {
                .success = Not hasError,
                .message = If(hasError, If(String.IsNullOrEmpty(sapSummaryMessage), "Batch processed with some errors at SAP. No records were saved to the database.", sapSummaryMessage), "Upload & Save Completed Successfully."),
                .results = processResults
            }))
            End Using
            End Using

        Catch ex As Exception
            context.Response.Write(New JavaScriptSerializer().Serialize(New With {.success = False, .message = ex.Message}))
        End Try
    End Sub

    Private Sub InsertToDB(cmd As SqlCommand, f As Object, t As Object, amt As Decimal, typeFrom As String, user As String, rmk As String, actionType As String)

        cmd.CommandText = "INSERT INTO [dbo].[OTB_Switching_Transaction] " &
            "([Year], [Month], [Company], [Category], [Segment], [Brand], [Vendor], [From], [BudgetAmount], [Release], " &
            "[SwitchYear], [SwitchMonth], [SwitchCompany], [SwitchCategory], [SwitchSegment], [To], [SwitchBrand], [SwitchVendor], " &
            "[OTBStatus], [Batch], [Remark], [CreateBy], [CreateDT], [ActionBy]) " &
            "VALUES (@Y, @M, @Co, @Ca, @Se, @Br, @Ve, @TyF, @Amt, 0, @SY, @SM, @SCo, @SCa, @SSe, @TyT, @SBr, @SVe, 'Approved', NULL, @Rem, @User, GETDATE(), @User)"

        cmd.Parameters.Clear()
        cmd.Parameters.AddWithValue("@Y", f("Year"))
        cmd.Parameters.AddWithValue("@M", f("Month"))
        cmd.Parameters.AddWithValue("@Co", f("Company"))
        cmd.Parameters.AddWithValue("@Ca", f("Category"))
        cmd.Parameters.AddWithValue("@Se", f("Segment"))
        cmd.Parameters.AddWithValue("@Br", f("Brand"))
        cmd.Parameters.AddWithValue("@Ve", f("Vendor"))
        cmd.Parameters.AddWithValue("@TyF", typeFrom)
        cmd.Parameters.AddWithValue("@Amt", amt)
        cmd.Parameters.AddWithValue("@User", user)
        cmd.Parameters.AddWithValue("@Rem", If(String.IsNullOrEmpty(rmk), DBNull.Value, rmk))

        If actionType.Equals("Switch", StringComparison.OrdinalIgnoreCase) Then
            cmd.Parameters.AddWithValue("@SY", t("Year"))
            cmd.Parameters.AddWithValue("@SM", t("Month"))
            cmd.Parameters.AddWithValue("@SCo", t("Company"))
            cmd.Parameters.AddWithValue("@SCa", t("Category"))
            cmd.Parameters.AddWithValue("@SSe", t("Segment"))
            cmd.Parameters.AddWithValue("@SBr", t("Brand"))
            cmd.Parameters.AddWithValue("@SVe", t("Vendor"))

            Dim typeTo As String = ""
            If typeFrom = "D" Then typeTo = "C"
            If typeFrom = "G" Then typeTo = "F"
            If typeFrom = "I" Then typeTo = "H"
            cmd.Parameters.AddWithValue("@TyT", typeTo)
        Else
            cmd.Parameters.AddWithValue("@SY", DBNull.Value)
            cmd.Parameters.AddWithValue("@SM", DBNull.Value)
            cmd.Parameters.AddWithValue("@SCo", DBNull.Value)
            cmd.Parameters.AddWithValue("@SCa", DBNull.Value)
            cmd.Parameters.AddWithValue("@SSe", DBNull.Value)
            cmd.Parameters.AddWithValue("@SBr", DBNull.Value)
            cmd.Parameters.AddWithValue("@SVe", DBNull.Value)
            cmd.Parameters.AddWithValue("@TyT", DBNull.Value)
        End If
        cmd.ExecuteNonQuery()
    End Sub

    Private Function IsValidMaster(dt As DataTable, codeCol As String, codeVal As String) As Boolean
        If dt Is Nothing OrElse String.IsNullOrEmpty(codeVal) Then Return False
        Dim rows = dt.Select($"[{codeCol}] = '{codeVal.Replace("'", "''")}'")
        Return rows.Length > 0
    End Function

    Private Shared Function ParseSubmittedDimension(values As Dictionary(Of String, Object), label As String) As OTBSwitchBudgetGuard.BudgetDimension
        If values Is Nothing Then Throw New InvalidOperationException("Every row must include a " & label & " budget dimension.")
        Return OTBSwitchBudgetGuard.CreateDimension(
            GetRequiredSubmittedText(values, "Year", label),
            GetRequiredSubmittedText(values, "Month", label),
            GetRequiredSubmittedText(values, "Company", label),
            GetRequiredSubmittedText(values, "Category", label),
            GetRequiredSubmittedText(values, "Segment", label),
            GetRequiredSubmittedText(values, "Brand", label),
            GetRequiredSubmittedText(values, "Vendor", label))
    End Function

    Private Shared Function GetRequiredSubmittedText(values As Dictionary(Of String, Object), key As String, label As String) As String
        Dim raw As Object = Nothing
        If values Is Nothing OrElse Not values.TryGetValue(key, raw) Then
            Throw New InvalidOperationException(label & " " & key & " is required.")
        End If
        Dim text As String = Convert.ToString(raw, CultureInfo.InvariantCulture).Trim()
        If String.IsNullOrWhiteSpace(text) Then Throw New InvalidOperationException(label & " " & key & " is required.")
        Return text
    End Function

    Private Shared Sub EnsureAjaxMutationRequest(context As HttpContext)
        If context Is Nothing OrElse Not String.Equals(context.Request.HttpMethod, "POST", StringComparison.OrdinalIgnoreCase) Then
            Throw New UnauthorizedAccessException("This action must be submitted with a same-origin POST request.")
        End If
        If Not String.Equals(context.Request.Headers("X-Requested-With"), "XMLHttpRequest", StringComparison.OrdinalIgnoreCase) Then
            Throw New UnauthorizedAccessException("The request could not be verified. Refresh the page and try again.")
        End If
    End Sub

    Private Shared Function EnsureSwitchEditPermission(context As HttpContext) As String
        Dim currentUser As String = If(context Is Nothing OrElse context.Session Is Nothing,
                                       "", Convert.ToString(context.Session("user"))).Trim()
        If String.IsNullOrWhiteSpace(currentUser) OrElse currentUser.Length > 100 Then
            Throw New UnauthorizedAccessException("Your login session has expired. Please sign in again.")
        End If
        Dim roleName As Object = context.Session("UserRole")
        Dim rights As PermissionHelper.UserRights = PermissionHelper.GetPermission(currentUser, "createOTBswitching.aspx", roleName)
        If Not rights.CanView OrElse Not rights.CanEdit Then
            Throw New UnauthorizedAccessException("You do not have permission to create OTB movements.")
        End If
        Return currentUser
    End Function

    Private Shared Sub ValidateMovementPersistenceFields(amount As Decimal, remark As String)
        If amount <= 0D Then Throw New InvalidOperationException("Every amount must be greater than 0.")
        If amount > MaxMovementAmount Then Throw New InvalidOperationException("A movement amount exceeds the database limit.")
        If If(remark, "").Length > 500 Then Throw New InvalidOperationException("Remark must not exceed 500 characters.")
    End Sub

    ReadOnly Property IsReusable() As Boolean Implements IHttpHandler.IsReusable
        Get
            Return False
        End Get
    End Property
End Class
