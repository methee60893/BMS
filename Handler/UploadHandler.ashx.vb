Imports System
Imports System.ComponentModel.DataAnnotations
Imports System.Data
Imports System.Data.SqlClient
Imports System.Globalization
Imports System.IO
Imports System.Text
Imports System.Web
Imports System.Web.Script.Serialization
Imports ExcelDataReader

Public Class UploadHandler : Implements IHttpHandler

    Private Shared connectionString As String = ConfigurationManager.ConnectionStrings("BMSConnectionString")?.ConnectionString

    Public Sub ProcessRequest(ByVal context As HttpContext) Implements IHttpHandler.ProcessRequest
        context.Response.Clear()
        context.Response.ContentType = "text/html"
        context.Response.ContentEncoding = Encoding.UTF8

        '  รับค่า uploadBy จาก form data
        Dim uploadBy As String = context.Request.Form("uploadBy")
        If String.IsNullOrEmpty(uploadBy) Then uploadBy = "unknown"

        Dim action As String = context.Request("action")
        If action = "savePreview" Then
            Try
                Dim jsonData As String = context.Request.Form("selectedData")
                SaveFromPreview(jsonData, uploadBy, context) ' เรียก Method ใหม่
            Catch ex As Exception
                context.Response.StatusCode = 500
                context.Response.Write($"<div class='alert alert-danger'>Error: {HttpUtility.HtmlEncode(ex.Message)}</div>")
            End Try
            Return ' ออกจากการทำงานทันที
        End If

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
                SaveToDatabase(dt, uploadBy, context) '  ส่ง uploadBy ไปด้วย
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

    ' ฟังก์ชันแยก CSV ที่มี comma อยู่ในข้อความ เช่น "abc, def", ghi
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
        sb.Append("<th class='text-center' style='width:50px;'>No.</th>")
        sb.Append("<th style='width:80px;'>Type</th>")
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
        sb.Append("<th class='text-end' style='width:130px;'>T0-BE Amount (TH)</th>")
        sb.Append("<th class='text-end' style='width:150px;'>Current total approved budget</th>")
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
            Dim typeValue As String = If(row("Type") IsNot DBNull.Value, row("Type").ToString().Trim(), "")
            Dim yearValue As String = If(row("Year") IsNot DBNull.Value, row("Year").ToString().Trim(), "")
            Dim monthValue As String = If(row("Month") IsNot DBNull.Value, row("Month").ToString().Trim(), "")
            Dim categoryValue As String = If(row("Category") IsNot DBNull.Value, row("Category").ToString().Trim(), "")
            Dim companyValue As String = If(row("Company") IsNot DBNull.Value, row("Company").ToString().Trim(), "")
            Dim segmentValue As String = If(row("Segment") IsNot DBNull.Value, row("Segment").ToString().Trim(), "")
            Dim brandValue As String = If(row("Brand") IsNot DBNull.Value, row("Brand").ToString().Trim(), "")
            Dim vendorValue As String = If(row("Vendor") IsNot DBNull.Value, row("Vendor").ToString().Trim(), "")
            Dim amountValue As String = If(row("Amount") IsNot DBNull.Value, row("Amount").ToString().Trim(), "")
            Dim remarkValue As String = If(row("Remark") IsNot DBNull.Value, row("Remark").ToString().Trim(), "")

            ' Validate แต่ละแถว
            Dim errorMessages As New List(Of String)
            Dim isValid As Boolean = True
            Dim canUpdate As Boolean = False

            Try
                Dim allErrors As String = validator.ValidateAllWithDuplicateCheck(typeValue, yearValue, monthValue,
                                                                             categoryValue, companyValue, segmentValue,
                                                                             brandValue, vendorValue, amountValue, canUpdate)


                ' แยก error messages
                If Not String.IsNullOrWhiteSpace(allErrors) Then
                    Dim errors() As String = allErrors.Split(New Char() {" "c}, StringSplitOptions.RemoveEmptyEntries)
                    For Each errMsg As String In errors
                        Dim trimmed As String = errMsg.Trim()
                        If Not String.IsNullOrEmpty(trimmed) Then
                            errorMessages.Add(trimmed)
                        End If
                    Next
                End If

                Dim uniqueKey As String = $"{typeValue}|{yearValue}|{monthValue}|{categoryValue}|{companyValue}|{segmentValue}|{brandValue}|{vendorValue}"
                If duplicateInExcelChecker.ContainsKey(uniqueKey) Then
                    errorMessages.Add("Duplicated_Draft OTB_Excel")
                    isValid = False
                Else
                    duplicateInExcelChecker.Add(uniqueKey, i)
                End If

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
                          HttpUtility.HtmlAttributeEncode(typeValue),
                          HttpUtility.HtmlAttributeEncode(yearValue),
                          HttpUtility.HtmlAttributeEncode(monthValue),
                          HttpUtility.HtmlAttributeEncode(categoryValue),
                          HttpUtility.HtmlAttributeEncode(companyValue),
                          HttpUtility.HtmlAttributeEncode(segmentValue),
                          HttpUtility.HtmlAttributeEncode(brandValue),
                          HttpUtility.HtmlAttributeEncode(vendorValue),
                          HttpUtility.HtmlAttributeEncode(amountValue),
                          HttpUtility.HtmlAttributeEncode(remarkValue),
                          canUpdate.ToString().ToLower())
            Else
                ' Invalid = checkbox disabled
                sb.Append("<td class='text-center'><input type='checkbox' class='form-check-input' disabled></td>")
            End If

            ' No. Column
            sb.AppendFormat("<td class='text-center'>{0}</td>", i + 1)

            ' Type Column (Original = ดำ, Revise = แดง)
            Dim typeClass As String = If(typeValue.Equals("Original", StringComparison.OrdinalIgnoreCase), "", "text-danger fw-bold")
            sb.AppendFormat("<td class='text-center {0}'>{1}</td>", typeClass, HttpUtility.HtmlEncode(typeValue))


            ' Year Column
            sb.AppendFormat("<td class='text-center'>{0}</td>", HttpUtility.HtmlEncode(yearValue))


            ' Month Column
            Dim monthDisplay As String = monthValue
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
                Dim amountDec As Decimal = Convert.ToDecimal(amountValue)
                sb.AppendFormat("<td class='text-end'>{0}</td>", amountDec.ToString("N2"))
            Catch
                sb.AppendFormat("<td class='text-end'>{0}</td>", HttpUtility.HtmlEncode(amountValue))
            End Try



            ' Current Budget (ยังไม่มีข้อมูล)
            sb.AppendFormat("<td class='text-end'>{0}</td>", HttpUtility.HtmlEncode(0.00))

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

        'If validCount > 0 Then
        '    sb.Append("<button id='submitBtn' onclick='submitData()' class='btn btn-success btn-lg'>")
        '    sb.Append("<i class='bi bi-check-circle'></i> Submit")
        '    sb.Append("</button>")
        'Else
        '    sb.Append("<button class='btn btn-secondary btn-lg' disabled>")
        '    sb.Append("<i class='bi bi-x-circle'></i> No Valid Data to Submit")
        '    sb.Append("</button>")
        'End If

        'sb.Append("</div>")
        'sb.Append("</div>")

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
        Dim newBatch As String = GetNextBatchNumber()
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
                    Dim versionValue As String = CalculateVersionFromHistory(typeValue, yearValue, monthValue,
                                                                         categoryValue, companyValue, segmentValue,
                                                                         brandValue, vendorValue)

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
                        newRow("Version") = versionValue
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
        Using conn As New SqlConnection(connectionString)
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

    Private Sub SaveFromPreview(jsonData As String, uploadBy As String, context As HttpContext)
        ' === 0. แปลง JSON ===
        If String.IsNullOrEmpty(jsonData) Then
            Throw New Exception("No data selected.")
        End If

        Dim serializer As New JavaScriptSerializer()
        Dim selectedRows As List(Of OTBUploadPreviewRow) = serializer.Deserialize(Of List(Of OTBUploadPreviewRow))(jsonData)

        If selectedRows.Count = 0 Then
            ' แม้ JS จะเช็คแล้ว แต่ Server ก็ควรเช็คด้วย
            context.Response.Write("No rows were selected to save.")
            Return
        End If

        ' === 1. สร้าง Validator (สำหรับตรวจสอบข้อมูล) ===
        Dim validator As New OTBValidate()

        ' === 2. ดึง Batch ใหม่ ===
        Dim newBatch As String = GetNextBatchNumber()
        Dim createDT As DateTime = DateTime.Now

        ' === 3. แยกข้อมูลเป็น INSERT และ UPDATE ===
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
        insertTable.Columns.Add("Remark", GetType(String))
        insertTable.Columns.Add("CreateDT", GetType(DateTime))

        Dim updateList As New List(Of Dictionary(Of String, Object))

        ' --- (START) NEW DE-DUPLICATION LOGIC ---
        ' 3.5. สร้าง Dictionary เพื่อเก็บแถวที่ไม่ซ้ำกัน (ยึดแถวสุดท้าย)
        Dim uniqueRowsToProcess As New Dictionary(Of String, OTBUploadPreviewRow)(StringComparer.OrdinalIgnoreCase)
        Dim totalSelected As Integer = selectedRows.Count
        Dim duplicateInBatchCount As Integer = 0

        For Each row As OTBUploadPreviewRow In selectedRows
            ' สร้าง Key จากทุกฟิลด์ *ยกเว้น* Amount
            Dim compositeKey As String = String.Join("|", New String() {
                row.Type, row.Year, row.Month, row.Category, row.Company,
                row.Segment, row.Brand, row.Vendor,
                If(row.Remark, "") ' (เพิ่ม Remark เข้าไปใน Key ด้วย)
            })

            If uniqueRowsToProcess.ContainsKey(compositeKey) Then
                duplicateInBatchCount += 1
            End If
            ' Add or Overwrite: การทำแบบนี้จะทำให้ Dictionary เก็บเฉพาะแถว "สุดท้าย" ที่มี Key นี้
            uniqueRowsToProcess(compositeKey) = row
        Next
        ' --- (END) NEW DE-DUPLICATION LOGIC ---

        Dim savedCount As Integer = 0
        Dim updatedCount As Integer = 0

        ' === 4. วนลูปข้อมูลที่ส่งมาจาก Preview (นี่คือส่วนที่เปลี่ยน) ===
        For Each row As OTBUploadPreviewRow In uniqueRowsToProcess.Values
            Try
                ' ดึงค่าจาก Object (ไม่ใช่ DataRow)
                Dim typeValue As String = row.Type
                Dim yearValue As String = row.Year
                Dim monthValue As String = row.Month
                Dim categoryValue As String = row.Category
                Dim companyValue As String = row.Company
                Dim segmentValue As String = row.Segment
                Dim brandValue As String = row.Brand
                Dim vendorValue As String = row.Vendor
                Dim amountValue As String = row.Amount
                Dim remarkValue As String = row.Remark
                ' Validate (เหมือนเดิม)
                Dim canUpdate As Boolean = False
                Dim errorMsg As String = validator.ValidateAllWithDuplicateCheck(typeValue, yearValue, monthValue,
                                                                            categoryValue, companyValue, segmentValue,
                                                                            brandValue, vendorValue, amountValue, canUpdate)

                ' ตรวจสอบ serious errors (เหมือนเดิม)
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

                ' บันทึกเฉพาะแถวที่ valid หรือ canUpdate (เหมือนเดิม) 
                If Not hasSeriousError Then
                    Dim amountDec As Decimal = Convert.ToDecimal(amountValue)
                    Dim versionValue As String = CalculateVersionFromHistory(typeValue, yearValue, monthValue,
                                                                         categoryValue, companyValue, segmentValue,
                                                                         brandValue, vendorValue)

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
                        updateData.Add("Remark", If(String.IsNullOrEmpty(remarkValue), DBNull.Value, remarkValue))
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
                        newRow("Version") = versionValue
                        newRow("UploadBy") = uploadBy
                        newRow("Batch") = newBatch
                        newRow("Remark") = If(String.IsNullOrEmpty(remarkValue), DBNull.Value, remarkValue) ' (เพิ่ม Remark)
                        newRow("CreateDT") = createDT
                        insertTable.Rows.Add(newRow)
                        savedCount += 1
                    End If
                End If

            Catch ex As Exception
                ' ข้ามแถวที่มีปัญหา
                Continue For
            End Try
        Next

        ' === 5. Execute INSERT และ UPDATE (เหมือนเดิม) ===
        Using conn As New SqlConnection(connectionString)
            conn.Open()
            Using transaction As SqlTransaction = conn.BeginTransaction()
                Try
                    ' Bulk Insert
                    If insertTable.Rows.Count > 0 Then
                        Using bulkCopy As New SqlBulkCopy(conn, SqlBulkCopyOptions.Default, transaction)
                            bulkCopy.DestinationTableName = "[dbo].[Template_Upload_Draft_OTB]"
                            ' ... (Column Mappings) ... 
                            For Each col As DataColumn In insertTable.Columns
                                bulkCopy.ColumnMappings.Add(col.ColumnName, col.ColumnName)
                            Next
                            bulkCopy.WriteToServer(insertTable)
                        End Using
                    End If

                    ' UPDATE แต่ละแถว
                    For Each updateData As Dictionary(Of String, Object) In updateList
                        ' ... (Query และ Parameters เหมือนเดิม) ...
                        Dim updateQuery As String = "UPDATE [dbo].[Template_Upload_Draft_OTB] SET [Amount] = @Amount, [UploadBy] = @UploadBy, [Batch] = @Batch, [UpdateDT] = @UpdateDT WHERE [Type] = @Type AND [Year] = @Year AND [Month] = @Month AND [Category] = @Category AND [Company] = @Company AND [Segment] = @Segment AND [Brand] = @Brand AND [Vendor] = @Vendor AND (OTBStatus IS NULL OR OTBStatus = 'Draft')"
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
                            cmd.Parameters.AddWithValue("@Remark", updateData("Remark")) ' (เพิ่ม Parameter)
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

        ' === 6. ส่งผลลัพธ์กลับ ===
        Dim duplicateMessage As String = ""
        If duplicateInBatchCount > 0 Then
            duplicateMessage = $" ({duplicateInBatchCount} duplicate rows in the file were consolidated based on the last row.)"
        End If
        context.Response.Write($"Successfully saved {savedCount} new rows and updated {updatedCount} existing DB rows (Batch: {newBatch}).{duplicateMessage}")
    End Sub

    Public ReadOnly Property IsReusable() As Boolean Implements IHttpHandler.IsReusable
        Get
            Return False
        End Get
    End Property

    Private Function GetNextBatchNumber() As String

        Dim currentMax As Integer = 0

        Using conn As New SqlConnection(connectionString)
            conn.Open()
            Using cmd As New SqlCommand("SELECT ISNULL(MAX(CAST(Batch AS INT)), 0) FROM [dbo].[Template_Upload_Draft_OTB]", conn)
                Dim result = cmd.ExecuteScalar()
                If result IsNot Nothing AndAlso Not IsDBNull(result) Then
                    currentMax = Convert.ToInt32(result)
                End If
            End Using
        End Using

        Return (currentMax + 1).ToString()
    End Function

    ''' <summary>
    ''' คำนวณ Version โดยดูจาก History ของ Key นี้
    ''' - ถ้ายังไม่เคยมีข้อมูล Key นี้ → A1
    ''' - ถ้าเคยมี Original แล้ว และ Type=Revise → R1, R2, R3...
    ''' </summary>
    Private Function CalculateVersionFromHistory(type As String, year As String, month As String,
                                                 category As String, company As String, segment As String,
                                                 brand As String, vendor As String) As String
        Try
            ' ถ้าเป็น Original → Version = A1 เสมอ
            If type.Equals("Original", StringComparison.OrdinalIgnoreCase) Then
                Return "A1"
            End If

            ' ถ้าเป็น Revise → ต้องหา Version ล่าสุดของ Key นี้
            If type.Equals("Revise", StringComparison.OrdinalIgnoreCase) Then
                Using conn As New SqlConnection(connectionString)
                    conn.Open()

                    ' Query หา Version ล่าสุดของ Key เดียวกัน
                    Dim query As String = "SELECT TOP 1 [Version]
                                      FROM [dbo].[Template_Upload_Draft_OTB]
                                      WHERE [Year] = @Year
                                        AND [Month] = @Month
                                        AND [Category] = @Category
                                        AND [Company] = @Company
                                        AND [Segment] = @Segment
                                        AND [Brand] = @Brand
                                        AND [Vendor] = @Vendor
                                      ORDER BY [CreateDT] DESC, [Batch] DESC"

                    Using cmd As New SqlCommand(query, conn)
                        cmd.Parameters.AddWithValue("@Year", year)
                        cmd.Parameters.AddWithValue("@Month", month)
                        cmd.Parameters.AddWithValue("@Category", category)
                        cmd.Parameters.AddWithValue("@Company", company)
                        cmd.Parameters.AddWithValue("@Segment", segment)
                        cmd.Parameters.AddWithValue("@Brand", brand)
                        cmd.Parameters.AddWithValue("@Vendor", vendor)

                        Dim lastVersion As Object = cmd.ExecuteScalar()

                        If lastVersion IsNot Nothing AndAlso Not IsDBNull(lastVersion) Then
                            Dim lastVersionStr As String = lastVersion.ToString()

                            ' แยก Version number
                            ' A1 → ไม่ควรมาถึงตรงนี้ (เพราะ Type=Revise)
                            ' R1 → R2
                            ' R2 → R3
                            If lastVersionStr.StartsWith("R") Then
                                Dim numPart As String = lastVersionStr.Substring(1)
                                Dim reviseNum As Integer
                                If Integer.TryParse(numPart, reviseNum) Then
                                    Return $"R{reviseNum + 1}"
                                End If
                            ElseIf lastVersionStr.StartsWith("A") Then
                                ' ถ้า Version ล่าสุดเป็น A1 แสดงว่านี่คือ Revise แรก
                                Return "R1"
                            End If
                        Else
                            ' ไม่พบข้อมูลเก่า แต่ Type=Revise → Error (แต่ให้ default R1)
                            Return "R1"
                        End If
                    End Using
                End Using
            End If

            ' Default
            Return "A1"

        Catch ex As Exception
            ' ถ้า error ให้ return default
            Return "A1"
        End Try
    End Function

End Class