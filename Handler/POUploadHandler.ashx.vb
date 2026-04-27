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
            If selectedRows Is Nothing OrElse selectedRows.Count = 0 Then
                Throw New Exception("No selected rows received.")
            End If

            Dim itemsToSave As New List(Of POValidate.DraftPOItem)

            For i As Integer = 0 To selectedRows.Count - 1
                Dim row = selectedRows(i)
                Dim amountTHB As Decimal = row.AmountCCY * row.ExRate

                itemsToSave.Add(New POValidate.DraftPOItem With {
                    .RowIndex = i + 1,
                    .DraftPO_ID = 0,
                    .PO_Year = row.Year, .PO_Month = row.Month,
                    .Company_Code = row.Company, .Category_Code = row.Category,
                    .Segment_Code = row.Segment, .Brand_Code = row.Brand,
                    .Vendor_Code = row.Vendor, .PO_No = row.DraftPONo,
                    .Currency = row.CCY, .Amount_CCY = row.AmountCCY,
                    .ExchangeRate = row.ExRate, .Amount_THB = amountTHB
                })
            Next

            ' 2. Re-Validate
            Dim validator As New POValidate()
            Dim valResult As POValidate.ValidationResult = validator.ValidateBatch(itemsToSave)

            If Not valResult.IsValid Then
                Dim errorMessages As New List(Of String)()
                If Not String.IsNullOrWhiteSpace(valResult.GlobalError) Then
                    errorMessages.Add(valResult.GlobalError)
                End If
                For Each rowError In valResult.RowErrors
                    errorMessages.Add($"Row {rowError.Key}: {String.Join(", ", rowError.Value)}")
                Next

                Dim errorMsg As String = If(errorMessages.Count > 0, String.Join(" | ", errorMessages), "Data verification failed on server side.")
                SendJsonResponse(context, False, errorMsg)
                Return
            End If

            ' 3. Save to Database (ใช้ SqlBulkCopy ภายใต้ SqlTransaction)
            Dim bulkTable As DataTable = BuildDraftPOBulkTable(itemsToSave, uploadBy)

            Using conn As New SqlConnection(connectionString)
                conn.Open()
                ' เริ่ม Transaction
                Using trans As SqlTransaction = conn.BeginTransaction()
                    Try
                        If bulkTable.Rows.Count > 0 Then
                            Using bulkCopy As New SqlBulkCopy(conn, SqlBulkCopyOptions.Default, trans)
                                bulkCopy.DestinationTableName = "[dbo].[Draft_PO_Transaction]"
                                bulkCopy.BatchSize = Math.Min(bulkTable.Rows.Count, 1000)
                                bulkCopy.BulkCopyTimeout = 300

                                For Each col As DataColumn In bulkTable.Columns
                                    bulkCopy.ColumnMappings.Add(col.ColumnName, col.ColumnName)
                                Next

                                bulkCopy.WriteToServer(bulkTable)
                            End Using
                        End If
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
        sb.Append("<th><input type='checkbox' id='selectAll' onclick='toggleAll(this)' checked ></th>")
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
            Dim exRateText As String = item.ExchangeRate.ToString(Globalization.CultureInfo.InvariantCulture)
            Dim amountCCYText As String = item.Amount_CCY.ToString(Globalization.CultureInfo.InvariantCulture)
            Dim amountTHBText As String = item.Amount_THB.ToString(Globalization.CultureInfo.InvariantCulture)

            Dim dataAttrs As String = $"data-pono=""{EncodeAttribute(item.PO_No)}"" data-year=""{EncodeAttribute(item.PO_Year)}"" data-month=""{EncodeAttribute(item.PO_Month)}"" " &
                                      $"data-company=""{EncodeAttribute(item.Company_Code)}"" data-category=""{EncodeAttribute(item.Category_Code)}"" " &
                                      $"data-segment=""{EncodeAttribute(item.Segment_Code)}"" data-brand=""{EncodeAttribute(item.Brand_Code)}"" " &
                                      $"data-vendor=""{EncodeAttribute(item.Vendor_Code)}"" data-ccy=""{EncodeAttribute(item.Currency)}"" " &
                                      $"data-exrate=""{EncodeAttribute(exRateText)}"" data-amountccy=""{EncodeAttribute(amountCCYText)}"" " &
                                      $"data-amountthb=""{EncodeAttribute(amountTHBText)}"""

            sb.Append($"<tr class='{rowClass}'>")
            sb.Append($"<td class='text-center'><input type='checkbox' name='selectedRows' {disabled} {dataAttrs} checked ></td>")
            sb.Append($"<td>{item.RowIndex}</td>")
            sb.Append($"<td class='text-center'>{statusIcon}</td>")
            sb.Append($"<td>{EncodeHtml(msg)}</td>")
            sb.Append($"<td>{EncodeHtml(item.PO_No)}</td>")
            sb.Append($"<td>{EncodeHtml(item.PO_Year)}</td>")
            sb.Append($"<td>{EncodeHtml(item.PO_Month)}</td>")
            sb.Append($"<td>{EncodeHtml(item.Vendor_Code)}</td>")
            sb.Append($"<td class='text-end'>{item.Amount_THB:N2}</td>")
            sb.Append("</tr>")
        Next

        sb.Append("</tbody></table>")
        sb.Append("<script>function toggleAll(source) { checkboxes = document.getElementsByName('selectedRows'); for(var i=0, n=checkboxes.length;i<n;i++) { if(!checkboxes[i].disabled) checkboxes[i].checked = source.checked; } }</script>")
        Return sb.ToString()
    End Function

    Private Function EncodeHtml(value As Object) As String
        Return HttpUtility.HtmlEncode(If(value, "").ToString())
    End Function

    Private Function EncodeAttribute(value As Object) As String
        Return HttpUtility.HtmlAttributeEncode(If(value, "").ToString())
    End Function

    Private Sub SendJsonResponse(context As HttpContext, success As Boolean, message As String)
        Dim resp = New With {.success = success, .message = message}
        context.Response.Write(JsonConvert.SerializeObject(resp))
    End Sub

    Private Function BuildDraftPOBulkTable(items As List(Of POValidate.DraftPOItem), uploadBy As String) As DataTable
        Dim table As New DataTable()
        table.Columns.Add("DraftPO_No", GetType(String))
        table.Columns.Add("PO_Year", GetType(Integer))
        table.Columns.Add("PO_Month", GetType(Integer))
        table.Columns.Add("Company_Code", GetType(String))
        table.Columns.Add("Category_Code", GetType(String))
        table.Columns.Add("Segment_Code", GetType(String))
        table.Columns.Add("Brand_Code", GetType(String))
        table.Columns.Add("Vendor_Code", GetType(String))
        table.Columns.Add("CCY", GetType(String))
        table.Columns.Add("Exchange_Rate", GetType(Decimal))
        table.Columns.Add("Amount_CCY", GetType(Decimal))
        table.Columns.Add("Amount_THB", GetType(Decimal))
        table.Columns.Add("PO_Type", GetType(String))
        table.Columns.Add("Status", GetType(String))
        table.Columns.Add("Status_Date", GetType(DateTime))
        table.Columns.Add("Status_By", GetType(String))
        table.Columns.Add("Remark", GetType(String))
        table.Columns.Add("Created_By", GetType(String))
        table.Columns.Add("Created_Date", GetType(DateTime))

        Dim nowValue As DateTime = DateTime.Now
        For Each item In items
            Dim row As DataRow = table.NewRow()
            row("DraftPO_No") = item.PO_No.Replace(" ", "")
            row("PO_Year") = Convert.ToInt32(item.PO_Year)
            row("PO_Month") = Convert.ToInt32(item.PO_Month)
            row("Company_Code") = item.Company_Code
            row("Category_Code") = item.Category_Code
            row("Segment_Code") = item.Segment_Code
            row("Brand_Code") = item.Brand_Code
            row("Vendor_Code") = item.Vendor_Code
            row("CCY") = item.Currency
            row("Exchange_Rate") = item.ExchangeRate
            row("Amount_CCY") = item.Amount_CCY
            row("Amount_THB") = item.Amount_THB
            row("PO_Type") = "Upload"
            row("Status") = "Draft"
            row("Status_Date") = nowValue
            row("Status_By") = uploadBy
            row("Remark") = DBNull.Value
            row("Created_By") = uploadBy
            row("Created_Date") = nowValue
            table.Rows.Add(row)
        Next

        Return table
    End Function

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
                .PO_No = GetSafeString(row, "Draft PO no.").Replace(" ", ""),
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
