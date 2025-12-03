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
Imports System.Linq

Public Class UploadHandler : Implements IHttpHandler

    Private Shared connectionString As String = ConfigurationManager.ConnectionStrings("BMSConnectionString")?.ConnectionString

    Public Sub ProcessRequest(ByVal context As HttpContext) Implements IHttpHandler.ProcessRequest
        context.Response.Clear()
        context.Response.ContentType = "text/html"
        context.Response.ContentEncoding = Encoding.UTF8

        '  รับค่า uploadBy จาก form data
        Dim uploadBy As String = context.Request.Form("uploadBy")
        If String.IsNullOrEmpty(uploadBy) Then uploadBy = "unknown"


        Dim util As New MasterDataUtil()

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
                context.Response.Write(GenerateHtmlTable(dt, util))
            ElseIf context.Request("action") = "save" Then
                SaveToDatabase(dt, uploadBy, context) '  ส่ง uploadBy ไปด้วย
                context.Response.Write("OK")
            End If

        Catch ex As Exception
            context.Response.StatusCode = 200
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

    Private Function GenerateHtmlTable(dt As DataTable, util As MasterDataUtil) As String
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

        Dim multiYearErrorMsg As String = ""
        Try
            ' ตรวจสอบว่ามีคอลัมน์ "Year" หรือไม่
            If Not dt.Columns.Contains("Year") Then
                multiYearErrorMsg = "SERIOUS_ERROR: ไม่พบคอลัมน์ 'Year' ในไฟล์ Excel"
            Else
                ' ดึงค่า Year ที่ไม่ซ้ำกันทั้งหมด (ไม่รวมค่าว่าง)
                Dim distinctYears = dt.AsEnumerable().
                                    Select(Function(r)
                                               If r.IsNull("Year") Then
                                                   Return Nothing
                                               Else
                                                   Return r("Year").ToString().Trim()
                                               End If
                                           End Function).
                                    Where(Function(y) Not String.IsNullOrEmpty(y)).
                                    Distinct().
                                    ToList()

                ' ถ้าพบ Year มากกว่า 1 ค่า ให้สร้างข้อความ Error
                If distinctYears.Count > 1 Then
                    multiYearErrorMsg = $"SERIOUS_ERROR: พบปีงบประมาณหลายค่า ({String.Join(", ", distinctYears)})"
                ElseIf distinctYears.Count = 0 Then
                    multiYearErrorMsg = "SERIOUS_ERROR: ไม่พบข้อมูล 'Year' ที่ถูกต้องในไฟล์ Excel"
                End If
            End If
        Catch ex As Exception
            multiYearErrorMsg = $"SERIOUS_ERROR: ไม่สามารถตรวจสอบข้อมูล Year ได้ ({ex.Message})"
        End Try

        Dim sb As New StringBuilder()
        ' ... (โค้ด CSS Style และ Table Header เหมือนเดิม) ...
        sb.Append("<div class='table-responsive' style='max-height:600px; overflow:auto;'>")
        sb.Append("<table id='previewTable' class='table table-bordered table-striped table-sm table-hover'>")
        sb.Append("<thead class='table-primary sticky-header'><tr>")
        sb.Append("<th class='text-center' style='width:60px;'>Select</th>")
        sb.Append("<th class='text-center' style='width:50px;'>Row No.</th>")
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
                    ' --- [BMS Gem MODIFICATION 6 START] ---
                    ' (เปลี่ยนตัวคั่นจาก " "c เป็น "|"c)
                    Dim errors() As String = allErrors.Split(New Char() {"|"c}, StringSplitOptions.RemoveEmptyEntries)
                    ' --- [BMS Gem MODIFICATION 6 END] ---
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

                If Not String.IsNullOrEmpty(multiYearErrorMsg) Then
                    errorMessages.Add(multiYearErrorMsg)
                End If

                ' ตรวจสอบว่า valid หรือไม่
                ' ถ้ามี error ที่ร้ายแรง (ไม่ใช่ Warning) = invalid
                Dim seriousErrors As Integer = 0
                For Each err As String In errorMessages
                    Dim isWarning As Boolean = False
                    'If err.Contains("Duplicated_Draft OTB") Then isWarning = True ' (Rule 3)
                    If err.Contains("(Will Update)") Then isWarning = True
                    If err.Contains("Duplicate_Approved_Warn") Then isWarning = True ' (Rule 2)
                    If err.Contains("(Will Revise)") Then isWarning = True
                    If err.Contains("Decimal_places_exceeded") Then isWarning = True ' (CSV Warning)

                    If Not isWarning Then
                        seriousErrors += 1 ' นี่คือ Error ร้ายแรงจริง
                    End If
                Next

                isValid = (seriousErrors = 0)


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
                rowClass = "table-danger" ' (สีแดง: Error ร้ายแรง)
            ElseIf canUpdate OrElse errorMessages.Count > 0 Then
                rowClass = "table-warning" ' (สีเหลือง: มี Warning แต่ Save ได้)
            End If
            sb.AppendFormat("<tr class='{0}' data-row-index='{1}'>", rowClass, i)

            ' Checkbox Column
            If isValid Then
                ' (isValid = ไม่มี Error ร้ายแรง) -> Checkbox Enabled
                Dim checkboxClass As String = If(canUpdate, "update-checkbox", "row-checkbox")
                sb.AppendFormat("<td class='text-center'><input type='checkbox' name='selectedRows' class='form-check-input {0}' value='{1}' checked data-type='{2}' data-year='{3}' data-month='{4}' data-category='{5}' data-company='{6}' data-segment='{7}' data-brand='{8}' data-vendor='{9}' data-amount='{10}' data-remark='{11}' data-can-update='{12}'></td>",
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

            ' ... (โค้ดสร้าง Cell ที่เหลือทั้งหมด เหมือนเดิม) ...
            sb.AppendFormat("<td class='text-center'>{0}</td>", i + 2)
            Dim typeClass As String = If(typeValue.Equals("Original", StringComparison.OrdinalIgnoreCase), "", "text-danger fw-bold")
            sb.AppendFormat("<td class='text-center {0}'>{1}</td>", typeClass, HttpUtility.HtmlEncode(typeValue))
            sb.AppendFormat("<td class='text-center'>{0}</td>", HttpUtility.HtmlEncode(yearValue))
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
            sb.AppendFormat("<td class='text-center'>{0}</td>", HttpUtility.HtmlEncode(categoryValue))
            sb.AppendFormat("<td>{0}</td>", HttpUtility.HtmlEncode(util.GetCategoryName(categoryValue)))
            sb.AppendFormat("<td class='text-center'>{0}</td>", HttpUtility.HtmlEncode(companyValue))
            sb.AppendFormat("<td class='text-center'>{0}</td>", HttpUtility.HtmlEncode(segmentValue))
            sb.AppendFormat("<td>{0}</td>", HttpUtility.HtmlEncode(util.GetSegmentName(segmentValue)))
            sb.AppendFormat("<td class='text-center'>{0}</td>", HttpUtility.HtmlEncode(brandValue))
            sb.AppendFormat("<td>{0}</td>", HttpUtility.HtmlEncode(util.GetBrandName(brandValue)))
            sb.AppendFormat("<td class='text-center'>{0}</td>", HttpUtility.HtmlEncode(vendorValue))
            sb.AppendFormat("<td>{0}</td>", HttpUtility.HtmlEncode(util.GetVendorName(vendorValue)))
            Try
                Dim amountDec As Decimal = Convert.ToDecimal(amountValue)
                sb.AppendFormat("<td class='text-end'>{0}</td>", amountDec.ToString("N2"))
            Catch
                sb.AppendFormat("<td class='text-end'>{0}</td>", HttpUtility.HtmlEncode(amountValue))
            End Try
            sb.AppendFormat("<td>{0}</td>", HttpUtility.HtmlEncode(remarkValue))
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
        sb.AppendFormat("<div class='alert alert-info mb-0'>Total: <strong>{0}</strong> rows |
 Valid: <strong class='text-success'>{1}</strong> | Error: <strong class='text-danger'>{2}</strong> | <strong class='text-warning'>Will Update: {3}</strong></div>",
                   dt.Rows.Count, validCount, errorCount, updateableCount)
        sb.Append("</div>")
        ' ... (โค้ด JavaScript ที่เหลือเหมือนเดิม) ...
        sb.Append("<script>")
        sb.Append("$(document).ready(function() {")
        sb.Append("  $('#selectAllCheckbox').on('change', function() {")
        sb.Append("    $('.row-checkbox:not(:disabled)').prop('checked', this.checked);")
        sb.Append("  });")
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
        insertTable.Columns.Add("Remark", GetType(String))
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
                Dim remarkValue As String = If(row("Remark") IsNot DBNull.Value, row("Remark").ToString().Trim(), "")
                ' Validate
                Dim canUpdate As Boolean = False
                Dim errorMsg As String = validator.ValidateAllWithDuplicateCheck(typeValue, yearValue, monthValue,
                                                                            categoryValue, companyValue, segmentValue,
                                                                            brandValue, vendorValue, amountValue, canUpdate)

                ' ตรวจสอบ serious errors
                Dim hasSeriousError As Boolean = False
                If Not String.IsNullOrEmpty(errorMsg) Then
                    ' --- [BMS Gem MODIFICATION 6 START] ---
                    ' (ใช้ Logic เดียวกับ GenerateHtmlTable)
                    ' (เปลี่ยนตัวคั่นจาก " "c เป็น "|"c)
                    Dim errors() As String = errorMsg.Split(New Char() {"|"c}, StringSplitOptions.RemoveEmptyEntries)

                    For Each err As String In errors
                        Dim isWarning As Boolean = False
                        If err.Contains("Duplicated_Draft OTB") Then isWarning = True ' (Rule 3)
                        If err.Contains("(Will Update)") Then isWarning = True
                        If err.Contains("Duplicate_Approved_Warn") Then isWarning = True ' (Rule 2)
                        If err.Contains("(Will Revise)") Then isWarning = True
                        If err.Contains("Decimal_places_exceeded") Then isWarning = True ' (CSV Warning)

                        If Not isWarning Then
                            hasSeriousError = True ' นี่คือ Error ร้ายแรงจริง
                            Exit For ' (เจอ Error ร้ายแรง 1 อันก็พอแล้ว)
                        End If
                    Next
                    ' --- [BMS Gem MODIFICATION 6 END] ---
                End If
                Dim versionValue As String = CalculateVersionFromHistory(typeValue, yearValue, monthValue,
                                                                       categoryValue, companyValue, segmentValue,
                                                                       brandValue, vendorValue)


                ' บันทึกเฉพาะแถวที่ valid หรือ canUpdate
                If Not hasSeriousError Then
                    Dim yearInt As Integer = Convert.ToInt32(yearValue)
                    Dim monthShort As Short = Convert.ToInt16(monthValue)
                    Dim amountDec As Decimal = Convert.ToDecimal(amountValue)


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
                        updateData.Add("Remark", If(String.IsNullOrEmpty(remarkValue), DBNull.Value, remarkValue))
                        updateData.Add("Version", versionValue)
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
                        newRow("Remark") = If(String.IsNullOrEmpty(remarkValue), DBNull.Value, remarkValue)
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
                                                     [UpdateDT] = @UpdateDT,
                                                     [Version] = @Version
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
                            cmd.Parameters.AddWithValue("@Version", updateData("Version"))

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


                ' 1. Validate (Logic นี้ถูกแก้ไขใน OTBValidate.vb แล้ว)
                Dim canUpdate As Boolean = False
                Dim errorMsg As String = validator.ValidateAllWithDuplicateCheck(typeValue, yearValue, monthValue,
                                                                                 categoryValue, companyValue, segmentValue,
                                                                                 brandValue, vendorValue, amountValue, canUpdate)

                ' 2. ตรวจสอบ serious errors (ไม่สนใจ Warning)
                Dim hasSeriousError As Boolean = False
                If Not String.IsNullOrEmpty(errorMsg) Then
                    ' --- [BMS Gem MODIFICATION 6 START] ---
                    ' (ใช้ Logic เดียวกับ GenerateHtmlTable)
                    ' (เปลี่ยนตัวคั่นจาก " "c เป็น "|"c)
                    Dim errors() As String = errorMsg.Split(New Char() {"|"c}, StringSplitOptions.RemoveEmptyEntries)

                    For Each err As String In errors
                        Dim isWarning As Boolean = False
                        If err.Contains("Duplicated_Draft OTB") Then isWarning = True ' (Rule 3)
                        If err.Contains("(Will Update)") Then isWarning = True
                        If err.Contains("Duplicate_Approved_Warn") Then isWarning = True ' (Rule 2)
                        If err.Contains("(Will Revise)") Then isWarning = True
                        If err.Contains("Decimal_places_exceeded") Then isWarning = True ' (CSV Warning)

                        If Not isWarning Then
                            hasSeriousError = True ' นี่คือ Error ร้ายแรงจริง
                            Exit For ' (เจอ Error ร้ายแรง 1 อันก็พอแล้ว)
                        End If
                    Next
                    ' --- [BMS Gem MODIFICATION 6 END] ---
                End If

                Dim versionValue As String = CalculateVersionFromHistory(typeValue, yearValue, monthValue,
                                                                           categoryValue, companyValue, segmentValue,
                                                                           brandValue, vendorValue)


                ' บันทึกเฉพาะแถวที่ valid หรือ canUpdate (เหมือนเดิม) 
                If Not hasSeriousError Then
                    Dim amountDec As Decimal = Convert.ToDecimal(amountValue)

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
                        updateData.Add("Version", versionValue)
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
                        newRow("Remark") = If(String.IsNullOrEmpty(remarkValue), DBNull.Value, remarkValue)
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
                        Dim updateQuery As String = "UPDATE [dbo].[Template_Upload_Draft_OTB] SET [Amount] = @Amount, [UploadBy] = @UploadBy, [Batch] = @Batch, [UpdateDT] = @UpdateDT, [Remark] = @Remark WHERE [Type] = @Type AND [Year] = @Year AND [Month] = @Month AND [Category] = @Category AND [Company] = @Company AND [Segment] = @Segment AND [Brand] = @Brand AND [Vendor] = @Vendor AND (OTBStatus IS NULL OR OTBStatus = 'Draft')"
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
                            cmd.Parameters.AddWithValue("@Version", updateData("Version")) ' <--- (BMS Gem) เพิ่ม Parameter
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
    ''' (NEW LOGIC) Calculates the next version (A1, R1...R15) for a key.
    ''' Logic: Checks ONLY OTB_Transaction table by querying the DB.
    ''' - If key not in OTB_Transaction -> returns "A1".
    ''' - If key in OTB_Transaction (max is A1) -> returns "R1".
    ''' - If key in OTB_Transaction (max is R(n)) -> returns "R(n+1)".
    ''' - Throws exception if next version > R15.
    ''' </summary>
    ''' <returns>The next version string (e.g., "A1", "R2")</returns>
    Private Function CalculateVersionFromHistory(type As String, year As String, month As String,
                                                 category As String, company As String, segment As String,
                                                 brand As String, vendor As String) As String
        Try
            ' --- [BMS Gem MODIFICATION LOGIC START] ---
            ' This function NO LONGER uses the 'type' parameter.
            ' It calculates the version based *only* on OTB_Transaction.

            Dim latestVersionNum As Integer = -1 ' A1 = 0, R1 = 1, R2 = 2

            ' Helper function to parse version string
            Dim getVersionNum = Function(v As String)
                                    If String.IsNullOrEmpty(v) Then Return -1
                                    If v.Equals("A1", StringComparison.OrdinalIgnoreCase) Then Return 0
                                    If v.StartsWith("R", StringComparison.OrdinalIgnoreCase) Then
                                        Dim numPart As Integer
                                        If Integer.TryParse(v.Substring(1), numPart) Then
                                            Return numPart ' R1=1, R2=2
                                        End If
                                    End If
                                    Return -1 ' Unknown format
                                End Function

            Using conn As New SqlConnection(connectionString)
                conn.Open()

                ' 1. Queryหา Version ล่าสุดจาก OTB_Transaction (Approved History) เท่านั้น
                Dim queryApproved As String = "
                    SELECT [Version]
                    FROM [dbo].[OTB_Transaction]
                    WHERE [Year] = @Year
                    ORDER BY ISNULL([ApprovedDate], [CreateDate]) DESC, [ID] DESC" ' อ้างอิงจากโค้ดเดิมในไฟล์เดียวกัน

                Using cmd As New SqlCommand(queryApproved, conn)
                    cmd.Parameters.AddWithValue("@Year", year)


                    ' เราแค่ต้องการค่าสูงสุดค่าเดียว
                    Dim lastApprovedVersion As Object = cmd.ExecuteScalar()
                    If lastApprovedVersion IsNot Nothing AndAlso Not IsDBNull(lastApprovedVersion) Then
                        latestVersionNum = getVersionNum(lastApprovedVersion.ToString())
                    End If
                End Using
            End Using

            ' 3. Calculate next version
            Dim nextVersionNum As Integer
            If latestVersionNum = -1 Then
                ' Rule 1: Not found in Approved, this is the first (A1)
                nextVersionNum = 0 ' A1
            Else
                ' Found in Approved, this is a Revise (R1, R2...)
                nextVersionNum = latestVersionNum + 1
            End If

            ' 4. Enforce R15 limit
            If nextVersionNum > 15 Then
                Throw New Exception($"Cannot save. The next version (R{nextVersionNum}) would exceed the R15 limit.")
            End If

            If nextVersionNum = 0 Then
                Return "A1"
            Else
                Return $"R{nextVersionNum}"
            End If
            ' --- [BMS Gem MODIFICATION LOGIC END] ---

        Catch ex As Exception
            ' ถ้า error ให้โยน Exception เพื่อหยุดการ Save
            Throw ex
        End Try
    End Function


End Class