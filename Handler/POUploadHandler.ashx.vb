Imports System
Imports System.Web
Imports System.Data
Imports System.Data.SqlClient
Imports System.IO
Imports System.Collections.Generic
Imports Newtonsoft.Json
Imports ExcelDataReader
Imports System.Text
Imports BMS

Public Class POUploadHandler
    Implements IHttpHandler, IRequiresSessionState

    Private Shared connectionString As String = ConfigurationManager.ConnectionStrings("BMSConnectionString")?.ConnectionString

    Public Sub ProcessRequest(ByVal context As HttpContext) Implements IHttpHandler.ProcessRequest
        Dim action As String = If(context.Request("action"), "").ToLower()

        If action = "preview" Then
            HandlePreview(context)
        ElseIf action = "savepreview" Then
            HandleSavePreview(context)
        Else
            context.Response.ContentType = "text/plain"
            context.Response.Write("Invalid Action")
        End If
    End Sub

    ' =========================================================
    ' 1. PREVIEW ACTION (Upload Excel -> Return HTML Table)
    ' =========================================================
    Private Sub HandlePreview(context As HttpContext)
        context.Response.ContentType = "text/html"
        Dim tempPath As String = ""

        Try
            If context.Request.Files.Count = 0 Then
                context.Response.Write("<div class='alert alert-warning'>กรุณาเลือกไฟล์ Excel</div>")
                Return
            End If

            Dim file As HttpPostedFile = context.Request.Files(0)
            tempPath = Path.GetTempFileName()
            file.SaveAs(tempPath)

            ' 1. อ่าน Excel
            Dim dt As DataTable = ReadExcelToDataTable(tempPath)
            If dt Is Nothing OrElse dt.Rows.Count = 0 Then
                context.Response.Write("<div class='alert alert-warning'>ไม่พบข้อมูลในไฟล์ Excel</div>")
                Return
            End If

            ' 2. แปลงเป็น Object และ Validate
            Dim poList As List(Of POValidate.DraftPOItem) = ConvertDataTableToPOList(dt)
            Dim validator As New POValidate()
            Dim result As POValidate.ValidationResult = validator.ValidateBatch(poList)

            ' 3. สร้าง HTML Table ส่งกลับ
            Dim html As String = GeneratePreviewTableHtml(poList, result)
            context.Response.Write(html)

        Catch ex As Exception
            context.Response.Write($"<div class='alert alert-danger'>System Error: {ex.Message}</div>")
        Finally
            If File.Exists(tempPath) Then Try : File.Delete(tempPath) : Catch : End Try
        End Try
    End Sub

    ' =========================================================
    ' 2. SAVE ACTION (Receive JSON -> Save to DB)
    ' =========================================================
    Private Sub HandleSavePreview(context As HttpContext)
        context.Response.ContentType = "application/json"

        Try
            Dim jsonStr As String = context.Request.Form("selectedData")
            Dim uploadBy As String = If(context.Request.Form("uploadBy"), "System")

            If String.IsNullOrEmpty(jsonStr) Then
                Throw New Exception("No data received.")
            End If

            ' 1. แปลง JSON กลับเป็น List Object
            Dim selectedRows = JsonConvert.DeserializeObject(Of List(Of FrontendPORow))(jsonStr)
            Dim itemsToSave As New List(Of POValidate.DraftPOItem)

            For Each row In selectedRows
                itemsToSave.Add(New POValidate.DraftPOItem With {
                    .DraftPO_ID = 0,
                    .PO_Year = row.Year, .PO_Month = row.Month,
                    .Company_Code = row.Company, .Category_Code = row.Category,
                    .Segment_Code = row.Segment, .Brand_Code = row.Brand,
                    .Vendor_Code = row.Vendor, .PO_No = row.DraftPONo,
                    .Currency = row.CCY, .Amount_CCY = row.AmountCCY,
                    .ExchangeRate = row.ExRate, .Amount_THB = row.AmountTHB
                })
            Next

            ' 2. Re-Validate
            Dim validator As New POValidate()
            Dim valResult As POValidate.ValidationResult = validator.ValidateBatch(itemsToSave)

            If Not valResult.IsValid Then
                Dim errorMsg As String = valResult.GlobalError
                If valResult.RowErrors.Count > 0 Then errorMsg &= " (Data verification failed on server side)"
                SendJsonResponse(context, False, errorMsg)
                Return
            End If

            ' 3. Save to Database (ใช้ SqlTransaction แทน TransactionScope)
            Dim successCount As Integer = 0

            Using conn As New SqlConnection(connectionString)
                conn.Open()
                ' เริ่ม Transaction
                Using trans As SqlTransaction = conn.BeginTransaction()
                    Try
                        For Each item In itemsToSave
                            Dim query As String = "INSERT INTO [BMS].[dbo].[Draft_PO_Transaction] " &
                                                  "([DraftPO_No], [PO_Year], [PO_Month], [Company_Code], [Category_Code], [Segment_Code], [Brand_Code], [Vendor_Code], " &
                                                  "[CCY], [Exchange_Rate], [Amount_CCY], [Amount_THB], " &
                                                  "[PO_Type], [Status], [Status_Date], [Status_By], [Remark], [Created_By], [Created_Date]) " &
                                                  "VALUES " &
                                                  "(@pono, @year, @month, @company, @category, @segment, @brand, @vendor, " &
                                                  "@ccy, @exRate, @amtCCY, @amtTHB, " &
                                                  "'Upload', 'Draft', GETDATE(), @user, NULL, @user, GETDATE())"

                            ' ต้องส่ง trans เข้าไปใน SqlCommand ด้วย
                            Using cmd As New SqlCommand(query, conn, trans)
                                cmd.Parameters.AddWithValue("@pono", item.PO_No)
                                cmd.Parameters.AddWithValue("@year", item.PO_Year)
                                cmd.Parameters.AddWithValue("@month", item.PO_Month)
                                cmd.Parameters.AddWithValue("@company", item.Company_Code)
                                cmd.Parameters.AddWithValue("@category", item.Category_Code)
                                cmd.Parameters.AddWithValue("@segment", item.Segment_Code)
                                cmd.Parameters.AddWithValue("@brand", item.Brand_Code)
                                cmd.Parameters.AddWithValue("@vendor", item.Vendor_Code)
                                cmd.Parameters.AddWithValue("@ccy", item.Currency)
                                cmd.Parameters.AddWithValue("@exRate", item.ExchangeRate)
                                cmd.Parameters.AddWithValue("@amtCCY", item.Amount_CCY)
                                cmd.Parameters.AddWithValue("@amtTHB", item.Amount_THB)
                                cmd.Parameters.AddWithValue("@user", uploadBy)
                                cmd.ExecuteNonQuery()
                            End Using
                            successCount += 1
                        Next

                        ' ถ้าทุกอย่างผ่าน ให้ Commit
                        trans.Commit()

                    Catch ex As Exception
                        ' ถ้ามี Error ให้ Rollback ทั้งหมด
                        trans.Rollback()
                        Throw ' ส่ง Error ออกไปให้ Catch ด้านนอกจัดการต่อ
                    End Try
                End Using
            End Using

            ' ส่งผลลัพธ์กลับ
            Dim results As New List(Of Object)
            For Each item In itemsToSave
                results.Add(New With {.DraftPONo = item.PO_No, .Status = "Success", .Message = "Saved"})
            Next
            context.Response.Write(JsonConvert.SerializeObject(results))

        Catch ex As Exception
            context.Response.StatusCode = 200 ' ส่ง 200 แต่เป็น Error Format
            SendJsonResponse(context, False, "System Error: " & ex.Message)
        End Try
    End Sub

    ' =========================================================
    ' HELPERS
    ' =========================================================

    Private Function GeneratePreviewTableHtml(items As List(Of POValidate.DraftPOItem), result As POValidate.ValidationResult) As String
        Dim sb As New StringBuilder()
        sb.Append("<table class='table table-bordered table-hover' id='tblPreview'>")
        sb.Append("<thead class='table-light'><tr>")
        sb.Append("<th><input type='checkbox' id='selectAll' onclick='toggleAll(this)'></th>")
        sb.Append("<th>#</th><th>Status</th><th>Message</th>")
        sb.Append("<th>PO No.</th><th>Year</th><th>Month</th><th>Vendor</th>")
        sb.Append("<th>Amount (THB)</th>")
        sb.Append("</tr></thead><tbody>")

        For Each item In items
            Dim isError As Boolean = result.RowErrors.ContainsKey(item.RowIndex)
            Dim rowClass As String = If(isError, "table-danger", "")
            Dim statusIcon As String = If(isError, "<i class='bi bi-x-circle text-danger'></i>", "<i class='bi bi-check-circle text-success'></i>")
            Dim msg As String = If(isError, String.Join(", ", result.RowErrors(item.RowIndex)), "Ready")
            Dim disabled As String = If(isError, "disabled", "")

            Dim dataAttrs As String = $"data-pono='{item.PO_No}' data-year='{item.PO_Year}' data-month='{item.PO_Month}' " &
                                      $"data-company='{item.Company_Code}' data-category='{item.Category_Code}' " &
                                      $"data-segment='{item.Segment_Code}' data-brand='{item.Brand_Code}' " &
                                      $"data-vendor='{item.Vendor_Code}' data-ccy='{item.Currency}' " &
                                      $"data-exrate='{item.ExchangeRate}' data-amountccy='{item.Amount_CCY}' " &
                                      $"data-amountthb='{item.Amount_THB}'"

            sb.Append($"<tr class='{rowClass}'>")
            sb.Append($"<td class='text-center'><input type='checkbox' name='selectedRows' {disabled} {dataAttrs}></td>")
            sb.Append($"<td>{item.RowIndex}</td>")
            sb.Append($"<td class='text-center'>{statusIcon}</td>")
            sb.Append($"<td>{msg}</td>")
            sb.Append($"<td>{item.PO_No}</td>")
            sb.Append($"<td>{item.PO_Year}</td>")
            sb.Append($"<td>{item.PO_Month}</td>")
            sb.Append($"<td>{item.Vendor_Code}</td>")
            sb.Append($"<td class='text-end'>{item.Amount_THB:N2}</td>")
            sb.Append("</tr>")
        Next

        sb.Append("</tbody></table>")
        sb.Append("<script>function toggleAll(source) { checkboxes = document.getElementsByName('selectedRows'); for(var i=0, n=checkboxes.length;i<n;i++) { if(!checkboxes[i].disabled) checkboxes[i].checked = source.checked; } }</script>")
        Return sb.ToString()
    End Function

    Private Sub SendJsonResponse(context As HttpContext, success As Boolean, message As String)
        Dim resp = New With {.success = success, .message = message}
        context.Response.Write(JsonConvert.SerializeObject(resp))
    End Sub

    Public Class FrontendPORow
        Public Property DraftPONo As String
        Public Property Year As String
        Public Property Month As String
        Public Property Company As String
        Public Property Category As String
        Public Property Segment As String
        Public Property Brand As String
        Public Property Vendor As String
        Public Property CCY As String
        Public Property ExRate As Decimal
        Public Property AmountCCY As Decimal
        Public Property AmountTHB As Decimal
    End Class

    Private Function ReadExcelToDataTable(filePath As String) As DataTable
        Using stream As FileStream = File.Open(filePath, FileMode.Open, FileAccess.Read)
            Using reader As IExcelDataReader = ExcelReaderFactory.CreateReader(stream)
                Dim conf As New ExcelDataSetConfiguration() With {.ConfigureDataTable = Function(__) New ExcelDataTableConfiguration() With {.UseHeaderRow = True}}
                Dim ds As DataSet = reader.AsDataSet(conf)
                Return If(ds.Tables.Count > 0, ds.Tables(0), Nothing)
            End Using
        End Using
    End Function

    Private Function ConvertDataTableToPOList(dt As DataTable) As List(Of POValidate.DraftPOItem)
        Dim list As New List(Of POValidate.DraftPOItem)()
        For i As Integer = 0 To dt.Rows.Count - 1
            Dim row As DataRow = dt.Rows(i)
            Dim item As New POValidate.DraftPOItem With {
                .RowIndex = i + 1,
                .DraftPO_ID = 0,
                .PO_No = GetSafeString(row, "Draft PO no."),
                .PO_Year = GetSafeString(row, "Year"),
                .PO_Month = GetSafeString(row, "Month"),
                .Category_Code = GetSafeString(row, "Category"),
                .Company_Code = GetSafeString(row, "Company"),
                .Segment_Code = GetSafeString(row, "Segment"),
                .Brand_Code = GetSafeString(row, "Brand"),
                .Vendor_Code = GetSafeString(row, "Vendor"),
                .Currency = GetSafeString(row, "CCY"),
                .Amount_CCY = ParseDecimal(GetSafeString(row, "Amount (CCY)")),
                .ExchangeRate = ParseDecimal(GetSafeString(row, "Ex. Rate"))
            }
            item.Amount_THB = item.Amount_CCY * item.ExchangeRate
            list.Add(item)
        Next
        Return list
    End Function

    Private Function GetSafeString(row As DataRow, colName As String) As String
        Return If(row.Table.Columns.Contains(colName) AndAlso row(colName) IsNot DBNull.Value, row(colName).ToString().Trim(), "")
    End Function

    Private Function ParseDecimal(val As String) As Decimal
        Dim d As Decimal
        Decimal.TryParse(val.Replace(",", ""), Globalization.NumberStyles.Any, Globalization.CultureInfo.InvariantCulture, d)
        Return d
    End Function

    Public ReadOnly Property IsReusable() As Boolean Implements IHttpHandler.IsReusable
        Get
            Return False
        End Get
    End Property
End Class