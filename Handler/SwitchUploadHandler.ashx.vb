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

Public Class SwitchUploadHandler
    Implements IHttpHandler, IRequiresSessionState

    Private Shared connectionString As String = ConfigurationManager.ConnectionStrings("BMSConnectionString")?.ConnectionString

    ' ตัวแปรสำหรับเก็บ Master Data
    Private dtCompany As DataTable
    Private dtCategory As DataTable
    Private dtSegment As DataTable
    Private dtBrand As DataTable
    Private dtVendor As DataTable

    Public Sub ProcessRequest(ByVal context As HttpContext) Implements IHttpHandler.ProcessRequest
        context.Response.Clear()
        context.Response.ContentEncoding = Encoding.UTF8

        Dim action As String = context.Request("action")
        Dim uploadBy As String = If(context.Session("user") IsNot Nothing, context.Session("user").ToString(), "System")

        If action = "preview" Then
            HandlePreview(context)
        ElseIf action = "save" Then
            HandleSave(context, uploadBy)
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
        Dim allValid As Boolean = True

        sb.Append("<div class='table-responsive' style='max-height: 500px;'>")
        sb.Append("<table class='table table-bordered table-sm table-hover' id='tblBulkPreview' style='font-size: 0.85rem;'>")
        sb.Append("<thead class='table-dark sticky-top'><tr>")
        sb.Append("<th style='width:30px;'>No.</th><th style='width:60px;'>Type</th>")
        sb.Append("<th>From Details (Source)</th>")
        sb.Append("<th class='text-end' style='width:100px;'>Amount</th>")
        sb.Append("<th>To Details (Destination)</th>")
        sb.Append("<th>Remark</th>")
        sb.Append("<th style='width:80px;'>Status</th><th>Issue</th>")
        sb.Append("</tr></thead><tbody>")

        For i As Integer = 0 To dt.Rows.Count - 1
            Dim row As DataRow = dt.Rows(i)
            Dim errors As New List(Of String)()

            ' --- 1. Read Data ---
            Dim type As String = GetVal(row, "Type")
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
            If String.IsNullOrEmpty(type) Then errors.Add("Missing Type")
            If Not Decimal.TryParse(amtStr, amount) OrElse amount <= 0 Then errors.Add("Invalid Amount")
            If String.IsNullOrEmpty(fYear) OrElse String.IsNullOrEmpty(fMonth) OrElse String.IsNullOrEmpty(fVend) Then errors.Add("Missing Source")

            If type.Equals("Switch", StringComparison.OrdinalIgnoreCase) Then
                If String.IsNullOrEmpty(tYear) OrElse String.IsNullOrEmpty(tMonth) OrElse String.IsNullOrEmpty(tVend) Then errors.Add("Missing Dest")
                If fYear = tYear AndAlso fMonth = tMonth AndAlso fComp = tComp AndAlso fCate = tCate AndAlso fSeg = tSeg AndAlso fBrand = tBrand AndAlso fVend = tVend Then
                    errors.Add("Source=Dest")
                End If
            End If

            ' --- 3. Budget Check ---
            If errors.Count = 0 AndAlso amount > 0 Then
                Try
                    Dim key As String = $"{fYear}|{fMonth}|{fCate}|{fComp}|{fSeg}|{fBrand}|{fVend}"
                    Dim currentDbBudget As Decimal = budgetCalc.CalculateCurrentApprovedBudget(fYear, fMonth, fCate, fComp, fSeg, fBrand, fVend)
                    Dim usedInBatch As Decimal = If(pendingUsage.ContainsKey(key), pendingUsage(key), 0)
                    Dim available As Decimal = currentDbBudget - usedInBatch

                    If available < amount Then
                        errors.Add($"Over Budget (Avail: {available:N2})")
                    Else
                        If pendingUsage.ContainsKey(key) Then pendingUsage(key) += amount Else pendingUsage.Add(key, amount)
                    End If
                Catch ex As Exception
                    errors.Add("Budget Error")
                End Try
            End If

            Dim isError As Boolean = (errors.Count > 0)
            If isError Then allValid = False

            ' --- 4. Serialize Data for Save ---
            Dim rowDataObj As New Dictionary(Of String, Object) From {
                {"Type", type}, {"Amount", amount}, {"Remark", remark},
                {"From", New Dictionary(Of String, String) From {{"Year", fYear}, {"Month", fMonth}, {"Company", fComp}, {"Category", fCate}, {"Segment", fSeg}, {"Brand", fBrand}, {"Vendor", fVend}}},
                {"To", New Dictionary(Of String, String) From {{"Year", tYear}, {"Month", tMonth}, {"Company", tComp}, {"Category", tCate}, {"Segment", tSeg}, {"Brand", tBrand}, {"Vendor", tVend}}}
            }
            Dim rowJson As String = If(isError, "", HttpUtility.HtmlAttributeEncode(New JavaScriptSerializer().Serialize(rowDataObj)))

            ' --- 5. Prepare Display Strings (With Names) ---
            Dim fCompName As String = GetName(dtCompany, "CompanyCode", "CompanyNameShort", fComp)
            Dim fCateName As String = GetName(dtCategory, "Cate", "Category", fCate)
            Dim fSegName As String = GetName(dtSegment, "SegmentCode", "SegmentName", fSeg)
            Dim fBrandName As String = GetName(dtBrand, "Brand Code", "Brand Name", fBrand)
            Dim fVendName As String = GetName(dtVendor, "VendorCode", "Vendor", fVend)

            ' HTML Block for From
            Dim fHtml As String = $"<strong>{fYear}/{fMonth}</strong><br/>" &
                                  $"<small>Comp: {fComp}:{fCompName}<br/>" &
                                  $"Vend: {fVend}:{fVendName}<br/>" &
                                  $"Brand: {fBrand}:{fBrandName}<br/>" &
                                  $"Seg: {fSeg}:{fSegName} | Cat: {fCate}:{fCateName}</small>"

            Dim tHtml As String = "-"
            If type.Equals("Switch", StringComparison.OrdinalIgnoreCase) Then
                Dim tCompName As String = GetName(dtCompany, "CompanyCode", "CompanyNameShort", tComp)
                Dim tCateName As String = GetName(dtCategory, "Cate", "Category", tCate)
                Dim tSegName As String = GetName(dtSegment, "SegmentCode", "SegmentName", tSeg)
                Dim tBrandName As String = GetName(dtBrand, "Brand Code", "Brand Name", tBrand)
                Dim tVendName As String = GetName(dtVendor, "VendorCode", "Vendor", tVend)

                tHtml = $"<strong>{tYear}/{tMonth}</strong><br/>" &
                        $"<small>Comp: {tComp}:{tCompName}<br/>" &
                        $"Vend: {tVend}:{tVendName}<br/>" &
                        $"Brand: {tBrand}:{tBrandName}<br/>" &
                        $"Seg: {tSeg}:{tSegName} | Cat: {tCate}:{tCateName}</small>"
            End If

            ' --- 6. Render Row ---
            Dim trClass As String = If(isError, "table-danger", "")
            sb.AppendFormat("<tr class='{0}'>", trClass)
            sb.AppendFormat("<td class='text-center'>{0}<input type='hidden' class='row-data' value='{1}'></td>", i + 1, rowJson)
            sb.AppendFormat("<td class='text-center fw-bold'>{0}</td>", type)

            ' From Column
            sb.AppendFormat("<td>{0}</td>", fHtml)

            ' Amount Column
            sb.AppendFormat("<td class='text-end fw-bold'>{0:N2}</td>", amount)

            ' To Column
            sb.AppendFormat("<td>{0}</td>", tHtml)

            ' Remark Column
            sb.AppendFormat("<td><small>{0}</small></td>", HttpUtility.HtmlEncode(remark))

            ' Status & Error
            Dim status As String = If(isError, "<span class='badge bg-danger'>Fail</span>", "<span class='badge bg-success'>Pass</span>")
            Dim errText As String = If(errors.Count > 0, String.Join("<br/>", errors), "")
            sb.AppendFormat("<td class='text-center'>{0}</td><td class='text-danger small'>{1}</td>", status, errText)
            sb.Append("</tr>")
        Next
        sb.Append("</tbody></table></div>")

        If Not allValid Then
            sb.Append("<div class='alert alert-warning mt-2'><i class='bi bi-exclamation-triangle'></i> ข้อมูลบางรายการไม่ถูกต้อง กรุณาแก้ไขไฟล์แล้ว Upload ใหม่ (ต้องผ่านทุกรายการถึงจะบันทึกได้)</div>")
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
    ' 2. SAVE (BULK SAP & DB)
    ' ==========================================
    Private Sub HandleSave(context As HttpContext, uploadBy As String)
        ' ... (ใช้ Code เดิมจาก Version ก่อนหน้าได้เลยครับ Logic Save ไม่ได้เปลี่ยน) ...
        context.Response.ContentType = "application/json"

        Try
            Dim json As String = context.Request.Form("data")
            Dim rows As List(Of Dictionary(Of String, Object)) = New JavaScriptSerializer().Deserialize(Of List(Of Dictionary(Of String, Object)))(json)

            Dim sapRequest As New OtbSwitchRequest()
            sapRequest.TestMode = "X"

            Dim pendingDbInserts As New List(Of Dictionary(Of String, Object))

            For Each row In rows
                Dim type As String = row("Type").ToString()
                Dim amt As Decimal = Convert.ToDecimal(row("Amount"))
                Dim f = TryCast(row("From"), Dictionary(Of String, Object))
                Dim t = TryCast(row("To"), Dictionary(Of String, Object))

                Dim fromCode As String = "E"
                Dim toCode As String = ""

                If type.Equals("Switch", StringComparison.OrdinalIgnoreCase) Then
                    fromCode = "D" : toCode = "C"
                    Dim dFrom As New Date(Convert.ToInt32(f("Year")), Convert.ToInt32(f("Month")), 1)
                    Dim dTo As New Date(Convert.ToInt32(t("Year")), Convert.ToInt32(t("Month")), 1)
                    If dFrom > dTo Then fromCode = "G" : toCode = "F"
                    If dFrom < dTo Then fromCode = "I" : toCode = "H"
                End If

                Dim item As New OtbSwitchItem With {
                    .DocYearFrom = f("Year").ToString(), .PeriodFrom = f("Month").ToString(), .FmAreaFrom = f("Company").ToString(),
                    .CatFrom = f("Category").ToString(), .SegmentFrom = f("Segment").ToString(), .BrandFrom = f("Brand").ToString(), .VendorFrom = f("Vendor").ToString(),
                    .Budget = amt.ToString("F2"), .TypeFrom = fromCode
                }

                If type.Equals("Switch", StringComparison.OrdinalIgnoreCase) Then
                    item.DocYearTo = t("Year").ToString() : item.PeriodTo = t("Month").ToString() : item.FmAreaTo = t("Company").ToString()
                    item.CatTo = t("Category").ToString() : item.SegmentTo = t("Segment").ToString() : item.BrandTo = t("Brand").ToString() : item.VendorTo = t("Vendor").ToString()
                    item.TypeTo = toCode
                Else
                    item.DocYearTo = Nothing : item.PeriodTo = Nothing : item.FmAreaTo = Nothing
                    item.CatTo = Nothing : item.SegmentTo = Nothing : item.BrandTo = Nothing : item.VendorTo = Nothing
                    item.TypeTo = Nothing
                End If
                sapRequest.Data.Add(item)

                Dim dbRow As New Dictionary(Of String, Object) From {
                    {"f", f}, {"t", t}, {"amt", amt}, {"typeFrom", fromCode},
                    {"remark", row("Remark")}, {"actionType", type}
                }
                pendingDbInserts.Add(dbRow)
            Next

            Dim sapRes = Task.Run(Async Function() Await SapApiHelper.SwitchOtbPlanAsync(sapRequest)).Result

            If sapRes Is Nothing Then Throw New Exception("No response from SAP API.")

            Dim hasError As Boolean = False
            Dim errorMsg As String = ""

            If sapRes.Status.ErrorCount > 0 OrElse sapRes.Status.Total <> sapRes.Status.Success Then
                hasError = True
                errorMsg = "SAP reported errors in batch processing."
            End If

            If sapRes.Results IsNot Nothing Then
                For Each resItem In sapRes.Results
                    If resItem.MessageType = "E" Then
                        hasError = True
                        errorMsg = If(String.IsNullOrEmpty(resItem.Message), "SAP Error (No message)", resItem.Message)
                        Exit For
                    End If
                Next
            End If

            If hasError Then
                Throw New Exception("Batch Failed at SAP: " & errorMsg)
            End If

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
                        dbTrans.Rollback()
                        Throw New Exception("Database Insert Failed: " & ex.Message)
                    End Try
                End Using
            End Using

            context.Response.Write(New JavaScriptSerializer().Serialize(New With {
                .success = True,
                .message = "Successfully processed " & rows.Count & " transactions."
            }))

        Catch ex As Exception
            context.Response.Write(New JavaScriptSerializer().Serialize(New With {
                .success = False,
                .message = ex.Message
            }))
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

    ReadOnly Property IsReusable() As Boolean Implements IHttpHandler.IsReusable
        Get
            Return False
        End Get
    End Property
End Class