Imports System
Imports System.ComponentModel.DataAnnotations
Imports System.Data
Imports System.Data.SqlClient
Imports System.Globalization
Imports System.IO
Imports System.Text
Imports System.Web
Imports ExcelDataReader


Public Class POUploadHandler
    Implements System.Web.IHttpHandler

    Public Shared connectionString93 As String = "Data Source=10.3.0.93;Initial Catalog=BMS;Persist Security Info=True;User ID=sa;Password=sql2014"

    Sub ProcessRequest(ByVal context As HttpContext) Implements IHttpHandler.ProcessRequest
        context.Response.Clear()
        context.Response.ContentType = "text/html"
        context.Response.ContentEncoding = Encoding.UTF8

        Dim uploadBy As String = context.Request.Form("uploadBy")
        If String.IsNullOrEmpty(uploadBy) Then uploadBy = "unknown"

        If context.Request.Files.Count = 0 Then
            context.Response.Write("No file uploaded.")
            Return
        End If

        Dim postedFile As HttpPostedFile = context.Request.Files(0)
        Dim tempPath As String = Path.GetTempFileName()
        postedFile.SaveAs(tempPath)

        Try
            Dim dt As DataTable = Nothing
            If Path.GetExtension(postedFile.FileName).ToLower() = ".csv" Then
                dt = ReadCsv(tempPath)
            Else
                dt = ReadExcel(tempPath)
            End If

            If context.Request("action") = "preview" Then
                context.Response.Write(GenerateHtmlTable(dt))
            ElseIf context.Request("action") = "save" Then
                SaveToDatabase(dt, uploadBy, context) ' 
                context.Response.Write("OK")
            End If

        Catch ex As Exception
            context.Response.StatusCode = 500
            context.Response.Write($"<div class='alert alert-danger'>Error: {HttpUtility.HtmlEncode(ex.Message)}</div>")
        Finally
            If File.Exists(tempPath) Then File.Delete(tempPath)
        End Try

    End Sub


    Private Function ReadExcel(filePath As String) As DataTable
        Dim result As DataTable
        Using stream = File.Open(filePath, FileMode.Open, FileAccess.Read)
            Using reader = ExcelReaderFactory.CreateReader(stream)
                Dim conf As New ExcelDataSetConfiguration()
                conf.ConfigureDataTable = Function(tableReader)
                                              Return New ExcelDataTableConfiguration() With {
                                             .UseHeaderRow = True
                                         }
                                          End Function
                Dim ds = reader.AsDataSet(conf)
                result = ds.Tables(0)
            End Using
        End Using
        Return result
    End Function

    Private Function ReadCsv(filePath As String) As DataTable
        Dim dt As New DataTable()
        Using reader As New StreamReader(filePath, Encoding.UTF8)
            Dim headers As String() = reader.ReadLine().Split(","c)
            For Each header In headers
                dt.Columns.Add(header.Trim())
            Next

            While Not reader.EndOfStream
                Dim line As String = reader.ReadLine()
                Dim values As String() = SplitCsvLine(line)
                dt.Rows.Add(values)
            End While
        End Using
        Return dt
    End Function

    Private Function SplitCsvLine(line As String) As String()
        Dim fields As New List(Of String)
        Dim inQuotes As Boolean = False
        Dim current As New StringBuilder()

        For i As Integer = 0 To line.Length - 1
            Dim c As Char = line(i)
            If c = """"c Then
                inQuotes = Not inQuotes
            ElseIf c = ","c AndAlso Not inQuotes Then
                fields.Add(current.ToString().Trim().Replace("""", ""))
                current.Clear()
            Else
                current.Append(c)
            End If
        Next
        fields.Add(current.ToString().Trim().Replace("""", ""))
        Return fields.ToArray()
    End Function

    Private Function GenerateHtmlTable(dt As DataTable) As String
        Dim validator As OTBValidate = Nothing
        Try
            validator = New OTBValidate()
        Catch ex As Exception
            ' ถ้า validator มี error ให้ return error message
            Return $"<div class='alert alert-danger'>
                    <strong>Error creating validator:</strong><br/>
                    {HttpUtility.HtmlEncode(ex.Message)}<br/>
                    <small>Stack Trace: {HttpUtility.HtmlEncode(ex.StackTrace)}</small>
                 </div>"
        End Try

        Dim sb As New StringBuilder()
        ' CSS Style
        sb.Append("<style>")
        sb.Append(".table-responsive { border: 1px solid #dee2e6; }")
        sb.Append(".sticky-header { position: sticky; top: 0; z-index: 10; }")
        sb.Append(".text-truncate-custom { max-width: 150px; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }")
        sb.Append("</style>")

        sb.Append("<div class='table-responsive' style='max-height:600px; overflow:auto;'>")
        sb.Append("<table id='previewTable' class='table table-bordered table-striped table-sm table-hover'>")

        ' Header
        sb.Append("<thead class='table-primary sticky-header'><tr>")
        sb.Append("<th class='text-center' style='width:60px;'>Select</th>")
        sb.Append("<th class='text-center' style='width:50px;'>Draft PO No.</th>")
        sb.Append("<th class='text-center' style='width:70px;'>Year</th>")
        sb.Append("<th class='text-center' style='width:70px;'>Month</th>")
        sb.Append("<th style='width:100px;'>Category</th>")
        sb.Append("<th class='text-center' style='width:100px;'>Category name</th>")
        sb.Append("<th class='text-center' style='width:90px;'>Company</th>")
        sb.Append("<th class='text-center' style='width:90px;'>Segment</th>")
        sb.Append("<th style='width:120px;'>Segment name</th>")
        sb.Append("<th class='text-center' style='width:80px;'>Brand</th>")
        sb.Append("<th style='width:120px;'>Brand name</th>")
        sb.Append("<th class='text-center' style='width:90px;'>Vendor</th>")
        sb.Append("<th style='width:150px;'>Vendor name</th>")
        sb.Append("<th class='text-end' style='width:130px;'> Amount (TH)</th>")
        sb.Append("<th class='text-end' style='width:130px;'> Amount (CCY)</th>")
        sb.Append("<th class='text-end' style='width:130px;'> CCY </th>")
        sb.Append("<th class='text-end' style='width:150px;'>Ex.Rate </th>")
        sb.Append("<th style='width:100px;'>Remark</th>")
        sb.Append("<th class='text-danger' style='min-width:250px;'>Error</th>")
        sb.Append("</tr></thead>")

        sb.Append("<tbody>")

        Dim validCount As Integer = 0
        Dim errorCount As Integer = 0
        Dim updateableCount As Integer = 0
        Dim duplicateInExcelChecker As New Dictionary(Of String, Integer)  ' เช็คซ้ำภายใน Excel file เอง


        For i As Integer = 0 To dt.Rows.Count - 1
            Dim row As DataRow = dt.Rows(i)

            ' ดึงค่าจาก Excel
            Dim PONOValue As String = If(row("Draft PO no. ") IsNot DBNull.Value, row("Draft PO no. ").ToString().Trim(), "")
            Dim yearValue As String = If(row("Year") IsNot DBNull.Value, row("Year").ToString().Trim(), "")
            Dim monthValue As String = If(row("Month") IsNot DBNull.Value, row("Month").ToString().Trim(), "")
            Dim categoryValue As String = If(row("Category") IsNot DBNull.Value, row("Category").ToString().Trim(), "")
            Dim companyValue As String = If(row("Company") IsNot DBNull.Value, row("Company").ToString().Trim(), "")
            Dim segmentValue As String = If(row("Segment") IsNot DBNull.Value, row("Segment").ToString().Trim(), "")
            Dim brandValue As String = If(row("Brand") IsNot DBNull.Value, row("Brand").ToString().Trim(), "")
            Dim vendorValue As String = If(row("Vendor") IsNot DBNull.Value, row("Vendor").ToString().Trim(), "")
            Dim amountTHBValue As String = If(row("AmountTHB") IsNot DBNull.Value, row("AmountTHB").ToString().Trim(), "")
            Dim amountCCYValue As String = If(row("AmountCCY") IsNot DBNull.Value, row("AmountCCY").ToString().Trim(), "")
            Dim ccyValue As String = If(row("CCY") IsNot DBNull.Value, row("CCY").ToString().Trim(), "")
            Dim exRateValue As String = If(row("ExRate") IsNot DBNull.Value, row("ExRate").ToString().Trim(), "")
            Dim RemarkValue As String = If(row("Remark") IsNot DBNull.Value, row("Remark").ToString().Trim(), "")



            ' Validate แต่ละแถว
            Dim errorMessages As New List(Of String)
            Dim isValid As Boolean = True
            Dim canUpdate As Boolean = False

            Try
                'Dim allErrors As String = validator.ValidateAllWithDuplicateCheck(, yearValue, monthValue,
                '                                                             categoryValue, companyValue, segmentValue,
                '                                                             brandValue, vendorValue, amountValue, canUpdate)


                ' แยก error messages
                'If Not String.IsNullOrWhiteSpace(allErrors) Then
                '    Dim errors() As String = allErrors.Split(New Char() {" "c}, StringSplitOptions.RemoveEmptyEntries)
                '    For Each errMsg As String In errors
                '        Dim trimmed As String = errMsg.Trim()
                '        If Not String.IsNullOrEmpty(trimmed) Then
                '            errorMessages.Add(trimmed)
                '        End If
                '    Next
                'End If

                'Dim uniqueKey As String = $"{typeValue}|{yearValue}|{monthValue}|{categoryValue}|{companyValue}|{segmentValue}|{brandValue}|{vendorValue}"
                'If duplicateInExcelChecker.ContainsKey(uniqueKey) Then
                '    errorMessages.Add("Duplicated_Draft OTB_Excel")
                '    isValid = False
                'Else
                '    duplicateInExcelChecker.Add(uniqueKey, i)
                'End If
                ' ตรวจสอบว่า valid หรือไม่
                ' ถ้ามี error ที่ร้ายแรง (ไม่ใช่ Duplicate) = invalid
                Dim seriousErrors As Integer = 0
                For Each err As String In errorMessages
                    If Not err.Contains("Duplicated_Draft OTB") AndAlso
                   Not err.Contains("(Will Update)") Then
                        seriousErrors += 1
                    End If
                Next

                isValid = seriousErrors = 0

            Catch ex As Exception
                errorMessages.Add("Data format error")
                isValid = False
            End Try

            ' นับสถิติ
            If isValid Then
                validCount += 1
                If canUpdate Then updateableCount += 1
            Else
                errorCount += 1
            End If

            ' สร้างแถว
            Dim rowClass As String = ""
            If Not isValid Then
                rowClass = "table-danger"
            ElseIf canUpdate Then
                rowClass = "table-warning" ' สีเหลืองสำหรับ Update
            End If
            sb.AppendFormat("<tr class='{0}' data-row-index='{1}'>", rowClass, i)

            ' Checkbox Column
            If isValid OrElse canUpdate Then
                ' Valid หรือ CanUpdate = checkbox enabled
                Dim checkboxClass As String = If(canUpdate, "update-checkbox", "row-checkbox")
                sb.AppendFormat("<td class='text-center'><input type='checkbox' name='selectedRows' class='form-check-input {0}' value='{1}' checked data-type='{2}' data-year='{3}' data-month='{4}' data-category='{5}' data-company='{6}' data-segment='{7}' data-brand='{8}' data-vendor='{9}' data-amount='{10}' data-can-update='{11}'></td>",
                          checkboxClass,
                          i,
                          HttpUtility.HtmlAttributeEncode(yearValue),
                          HttpUtility.HtmlAttributeEncode(monthValue),
                          HttpUtility.HtmlAttributeEncode(categoryValue),
                          HttpUtility.HtmlAttributeEncode(companyValue),
                          HttpUtility.HtmlAttributeEncode(segmentValue),
                          HttpUtility.HtmlAttributeEncode(brandValue),
                          HttpUtility.HtmlAttributeEncode(vendorValue),
                          HttpUtility.HtmlAttributeEncode(RemarkValue),
                          canUpdate.ToString().ToLower())
            Else
                ' Invalid = checkbox disabled
                sb.Append("<td class='text-center'><input type='checkbox' class='form-check-input' disabled></td>")
            End If

            ' No. Column
            sb.AppendFormat("<td class='text-center'>{0}</td>", i + 1)

            ' Type Column (Original = ดำ, Revise = แดง)
            'Dim typeClass As String = If(typeValue.Equals("Original", StringComparison.OrdinalIgnoreCase), "", "text-danger fw-bold")
            'sb.AppendFormat("<td class='text-center {0}'>{1}</td>", typeClass, HttpUtility.HtmlEncode(typeValue))


            ' Year Column
            sb.AppendFormat("<td class='text-center'>{0}</td>", HttpUtility.HtmlEncode(yearValue))


            ' Month Column
            Dim monthDisplay As String = yearValue
            Select Case monthValue
                Case "1" : monthDisplay = "Jan"
                Case "2" : monthDisplay = "Feb"
                Case "3" : monthDisplay = "Mar"
                Case "4" : monthDisplay = "Apr"
                Case "5" : monthDisplay = "May"
                Case "6" : monthDisplay = "Jun"
                Case "7" : monthDisplay = "Jul"
                Case "8" : monthDisplay = "Aug"
                Case "9" : monthDisplay = "Sep"
                Case "10" : monthDisplay = "Oct"
                Case "11" : monthDisplay = "Nov"
                Case "12" : monthDisplay = "Dec"
            End Select
            sb.AppendFormat("<td class='text-center'>{0}</td>", HttpUtility.HtmlEncode(monthDisplay))

            ' Category Code
            sb.AppendFormat("<td class='text-center'>{0}</td>", HttpUtility.HtmlEncode(categoryValue))

            ' Category Name (ดึงจาก Master - TODO: implement)
            sb.Append("<td class='text-center'>-</td>")

            ' Company
            sb.AppendFormat("<td class='text-center'>{0}</td>", HttpUtility.HtmlEncode(companyValue))

            ' Segment Code
            sb.AppendFormat("<td class='text-center'>{0}</td>", HttpUtility.HtmlEncode(segmentValue))

            ' Segment Name (TODO: ดึงจาก Master)
            sb.Append("<td>-</td>")

            ' Brand Code
            sb.AppendFormat("<td class='text-center'>{0}</td>", HttpUtility.HtmlEncode(brandValue))

            ' Brand Name (TODO: ดึงจาก Master)
            sb.Append("<td>-</td>")

            ' Vendor Code
            sb.AppendFormat("<td class='text-center'>{0}</td>", HttpUtility.HtmlEncode(vendorValue))

            ' Vendor Name (TODO: ดึงจาก Master)
            sb.Append("<td>-</td>")

            ' Amount
            Try
                'Dim amountDec As Decimal = Convert.ToDecimal(amountValue)
                'sb.AppendFormat("<td class='text-end'>{0}</td>", amountDec.ToString("N2"))
            Catch
                'sb.AppendFormat("<td class='text-end'>{0}</td>", HttpUtility.HtmlEncode(amountValue))
            End Try

            ' Current Budget (ยังไม่มีข้อมูล)
            sb.Append("<td class='text-end'>0.00</td>")

            ' Remark
            sb.AppendFormat("<td>{0}</td>", HttpUtility.HtmlEncode(remarkValue))

            ' Error Column
            If errorMessages.Count > 0 Then
                sb.AppendFormat("<td class='text-danger small'>{0}</td>", HttpUtility.HtmlEncode(String.Join(" ** ", errorMessages)))
            Else
                sb.Append("<td></td>")
            End If

            sb.Append("</tr>")
        Next

        sb.Append("</tbody></table></div>")

        ' Summary และปุ่ม Submit
        sb.Append("<div class='p-3 bg-light border-top'>")
        sb.AppendFormat("<div class='alert alert-info mb-0'>Total: <strong>{0}</strong> rows | Valid: <strong class='text-success'>{1}</strong> | Error: <strong class='text-danger'>{2}</strong> | <strong class='text-warning'>Will Update: {3}</strong></div>",
                   dt.Rows.Count, validCount, errorCount, updateableCount)
        sb.Append("</div>")
        sb.Append("<div class='col-md-4 text-end'>")


        ' JavaScript สำหรับจัดการ Checkbox
        sb.Append("<script>")
        sb.Append("$(document).ready(function() {")
        sb.Append("  // Select All checkbox functionality")
        sb.Append("  $('#selectAllCheckbox').on('change', function() {")
        sb.Append("    $('.row-checkbox:not(:disabled)').prop('checked', this.checked);")
        sb.Append("  });")
        sb.Append("  // Update select all when individual checkbox changes")
        sb.Append("  $('.row-checkbox').on('change', function() {")
        sb.Append("    var total = $('.row-checkbox:not(:disabled)').length;")
        sb.Append("    var checked = $('.row-checkbox:checked').length;")
        sb.Append("    $('#selectAllCheckbox').prop('checked', total === checked);")
        sb.Append("  });")
        sb.Append("});")
        sb.Append("</script>")

        Return sb.ToString()
    End Function

    Private Sub SaveToDatabase(dt As DataTable, uploadBy As String, context As HttpContext)
        ' === 1. ตรวจสอบคอลัมน์ที่จำเป็น ===
        Dim requiredColumns As String() = {"Type", "Year", "Month", "Category", "Company", "Segment", "Brand", "Vendor", "Amount"}
        For Each colName In requiredColumns
            If Not dt.Columns.Contains(colName) Then
                Throw New Exception($"Missing required column: {colName}")
            End If
        Next

        ' === 2. สร้าง Validator (สำหรับตรวจสอบข้อมูล) ===
        Dim validator As New OTBValidate()

        ' === 3. ดึง Batch ใหม่ ===
        Dim newBatch As String = ""
        Dim newBatchInt As Integer = Convert.ToInt32(newBatch)
        Dim createDT As DateTime = DateTime.Now

        ' === 4. แยกข้อมูลเป็น INSERT และ UPDATE ===
        Dim insertTable As New DataTable()
        insertTable.Columns.Add("Type", GetType(String))
        insertTable.Columns.Add("Year", GetType(String))
        insertTable.Columns.Add("Month", GetType(String))
        insertTable.Columns.Add("Category", GetType(String))
        insertTable.Columns.Add("Company", GetType(String))
        insertTable.Columns.Add("Segment", GetType(String))
        insertTable.Columns.Add("Brand", GetType(String))
        insertTable.Columns.Add("Vendor", GetType(String))
        insertTable.Columns.Add("Amount", GetType(String))
        insertTable.Columns.Add("Version", GetType(String))
        insertTable.Columns.Add("UploadBy", GetType(String))
        insertTable.Columns.Add("Batch", GetType(String))
        insertTable.Columns.Add("CreateDT", GetType(DateTime))

        Dim updateList As New List(Of Dictionary(Of String, Object))

        Dim savedCount As Integer = 0
        Dim updatedCount As Integer = 0

        For i As Integer = 0 To dt.Rows.Count - 1
            Dim row As DataRow = dt.Rows(i)

            Try
                ' ดึงค่า
                Dim typeValue As String = If(row("Type") IsNot DBNull.Value, row("Type").ToString().Trim(), "")
                Dim yearValue As String = If(row("Year") IsNot DBNull.Value, row("Year").ToString().Trim(), "")
                Dim monthValue As String = If(row("Month") IsNot DBNull.Value, row("Month").ToString().Trim(), "")
                Dim categoryValue As String = If(row("Category") IsNot DBNull.Value, row("Category").ToString().Trim(), "")
                Dim companyValue As String = If(row("Company") IsNot DBNull.Value, row("Company").ToString().Trim(), "")
                Dim segmentValue As String = If(row("Segment") IsNot DBNull.Value, row("Segment").ToString().Trim(), "")
                Dim brandValue As String = If(row("Brand") IsNot DBNull.Value, row("Brand").ToString().Trim(), "")
                Dim vendorValue As String = If(row("Vendor") IsNot DBNull.Value, row("Vendor").ToString().Trim(), "")
                Dim amountValue As String = If(row("Amount") IsNot DBNull.Value, row("Amount").ToString().Trim(), "")

                ' Validate
                Dim canUpdate As Boolean = False
                Dim errorMsg As String = validator.ValidateAllWithDuplicateCheck(typeValue, yearValue, monthValue,
                                                                            categoryValue, companyValue, segmentValue,
                                                                            brandValue, vendorValue, amountValue, canUpdate)

                ' ตรวจสอบ serious errors
                Dim hasSeriousError As Boolean = False
                If Not String.IsNullOrEmpty(errorMsg) Then
                    If errorMsg.Contains("No Original found") OrElse
                   errorMsg.Contains("Data format error") OrElse
                   errorMsg.Contains("is required") OrElse
                   errorMsg.Contains("Not found") Then
                        If Not errorMsg.Contains("(Will Update)") Then
                            hasSeriousError = True
                        End If
                    End If
                End If

                ' บันทึกเฉพาะแถวที่ valid หรือ canUpdate
                If Not hasSeriousError Then
                    Dim yearInt As Integer = Convert.ToInt32(yearValue)
                    Dim monthShort As Short = Convert.ToInt16(monthValue)
                    Dim amountDec As Decimal = Convert.ToDecimal(amountValue)

                    ' คำนวณ Version


                    If canUpdate Then
                        ' === UPDATE Case ===
                        Dim updateData As New Dictionary(Of String, Object)
                        updateData.Add("Type", typeValue)
                        updateData.Add("Year", yearValue)
                        updateData.Add("Month", monthValue)
                        updateData.Add("Category", categoryValue)
                        updateData.Add("Company", companyValue)
                        updateData.Add("Segment", segmentValue)
                        updateData.Add("Brand", brandValue)
                        updateData.Add("Vendor", vendorValue)
                        updateData.Add("Amount", amountDec.ToString("0.00"))
                        updateData.Add("UploadBy", uploadBy)
                        updateData.Add("Batch", newBatch)
                        updateData.Add("UpdateDT", createDT)

                        updateList.Add(updateData)
                        updatedCount += 1
                    Else
                        ' === INSERT Case ===
                        Dim newRow As DataRow = insertTable.NewRow()
                        newRow("Type") = typeValue
                        newRow("Year") = yearValue
                        newRow("Month") = monthValue
                        newRow("Category") = categoryValue
                        newRow("Company") = companyValue
                        newRow("Segment") = segmentValue
                        newRow("Brand") = brandValue
                        newRow("Vendor") = vendorValue
                        newRow("Amount") = amountDec.ToString("0.00")
                        newRow("UploadBy") = uploadBy
                        newRow("Batch") = newBatch
                        newRow("CreateDT") = createDT
                        insertTable.Rows.Add(newRow)
                        savedCount += 1
                    End If
                End If

            Catch ex As Exception
                Continue For
            End Try
        Next

        ' === 5. Execute INSERT และ UPDATE ===
        Using conn As New SqlConnection(connectionString93)
            conn.Open()
            Using transaction As SqlTransaction = conn.BeginTransaction()
                Try
                    ' Bulk Insert
                    If insertTable.Rows.Count > 0 Then
                        Using bulkCopy As New SqlBulkCopy(conn, SqlBulkCopyOptions.Default, transaction)
                            bulkCopy.DestinationTableName = "[dbo].[Template_Upload_Draft_OTB]"
                            bulkCopy.BatchSize = insertTable.Rows.Count
                            bulkCopy.BulkCopyTimeout = 300

                            For Each col As DataColumn In insertTable.Columns
                                bulkCopy.ColumnMappings.Add(col.ColumnName, col.ColumnName)
                            Next

                            bulkCopy.WriteToServer(insertTable)
                        End Using
                    End If

                    ' UPDATE แต่ละแถว
                    For Each updateData As Dictionary(Of String, Object) In updateList
                        Dim updateQuery As String = "UPDATE [dbo].[Template_Upload_Draft_OTB]
                                                 SET [Amount] = @Amount,
                                                     [UploadBy] = @UploadBy,
                                                     [Batch] = @Batch,
                                                     [UpdateDT] = @UpdateDT
                                                 WHERE [Type] = @Type
                                                   AND [Year] = @Year
                                                   AND [Month] = @Month
                                                   AND [Category] = @Category
                                                   AND [Company] = @Company
                                                   AND [Segment] = @Segment
                                                   AND [Brand] = @Brand
                                                   AND [Vendor] = @Vendor
                                                   AND (OTBStatus IS NULL OR OTBStatus = 'Draft')"

                        Using cmd As New SqlCommand(updateQuery, conn, transaction)
                            cmd.Parameters.AddWithValue("@Type", updateData("Type"))
                            cmd.Parameters.AddWithValue("@Year", updateData("Year"))
                            cmd.Parameters.AddWithValue("@Month", updateData("Month"))
                            cmd.Parameters.AddWithValue("@Category", updateData("Category"))
                            cmd.Parameters.AddWithValue("@Company", updateData("Company"))
                            cmd.Parameters.AddWithValue("@Segment", updateData("Segment"))
                            cmd.Parameters.AddWithValue("@Brand", updateData("Brand"))
                            cmd.Parameters.AddWithValue("@Vendor", updateData("Vendor"))
                            cmd.Parameters.AddWithValue("@Amount", updateData("Amount"))
                            cmd.Parameters.AddWithValue("@UploadBy", updateData("UploadBy"))
                            cmd.Parameters.AddWithValue("@Batch", updateData("Batch"))
                            cmd.Parameters.AddWithValue("@UpdateDT", updateData("UpdateDT"))

                            cmd.ExecuteNonQuery()
                        End Using
                    Next

                    transaction.Commit()

                Catch ex As Exception
                    transaction.Rollback()
                    Throw New Exception("Error saving data: " & ex.Message)
                End Try
            End Using
        End Using

        context.Response.Write($"Successfully saved {savedCount} new rows and updated {updatedCount} rows to Draft OTB (Batch: {newBatch})")
    End Sub

    ReadOnly Property IsReusable() As Boolean Implements IHttpHandler.IsReusable
        Get
            Return False
        End Get
    End Property

End Class