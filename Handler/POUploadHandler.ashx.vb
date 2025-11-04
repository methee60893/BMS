Imports System
Imports System.ComponentModel.DataAnnotations
Imports System.Data
Imports System.Data.SqlClient
Imports System.Globalization
Imports System.IO
Imports System.Text
Imports System.Web
Imports ExcelDataReader
Imports System.Web.Script.Serialization
Imports System.Web.SessionState

' Class สำหรับรับ JSON
Public Class POPreviewRow
    Public Property DraftPONo As String
    Public Property Year As String
    Public Property Month As String
    Public Property Category As String
    Public Property Company As String
    Public Property Segment As String
    Public Property Brand As String
    Public Property Vendor As String
    Public Property AmountTHB As String
    Public Property AmountCCY As String
    Public Property CCY As String
    Public Property ExRate As String
    Public Property Remark As String
End Class

Public Class POUploadHandler
    Implements System.Web.IHttpHandler, IRequiresSessionState

    Private Shared connectionString As String = ConfigurationManager.ConnectionStrings("BMSConnectionString")?.ConnectionString

    Sub ProcessRequest(ByVal context As HttpContext) Implements IHttpHandler.ProcessRequest
        context.Response.Clear()
        context.Response.ContentType = "text/html"
        context.Response.ContentEncoding = Encoding.UTF8

        Dim uploadBy As String = context.Request.Form("uploadBy")
        If String.IsNullOrEmpty(uploadBy) Then uploadBy = "unknown"

        ' ตรวจสอบ Action ใหม่ก่อน
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
                context.Response.Write("OK (Legacy Save)")
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

    Private Function GetExistingPOs(poNos As List(Of String)) As HashSet(Of String)
        ' === START MODIFICATION: 1 of 1 ===
        ' สร้าง HashSet แบบไม่สนใจตัวพิมพ์เล็ก/ใหญ่ (Case-Insensitive)
        ' This makes the C# check (HashSet.Contains) behave like the SQL check (JOIN)
        Dim existingPOs As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
        ' === END MODIFICATION ===

        If poNos.Count = 0 Then
            Return existingPOs
        End If

        ' 1. สร้าง DataTable สำหรับ Bulk Check
        Dim poCheckTable As New DataTable()
        poCheckTable.Columns.Add("DraftPO_No", GetType(String))

        ' === START MODIFICATION 2: ป้องกันการใส่ PO ซ้ำในตารางชั่วคราว ===
        ' ใช้ HashSet ชั่วคราวเพื่อกรอง PO ที่ซ้ำกันจากไฟล์ Excel
        ' ซึ่งจะป้องกัน Error "Violation of PRIMARY KEY" ใน #TempCheckPOs
        Dim uniquePOsForBulkCheck As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)

        For Each po In poNos
            If Not String.IsNullOrEmpty(po) Then
                ' .Add 
                If uniquePOsForBulkCheck.Add(po) Then
                    ' ถ้า po นี้ยังไม่เคยถูกเพิ่มลง HashSet ให้เพิ่มลงใน poCheckTable
                    poCheckTable.Rows.Add(po)
                End If
            End If
        Next
        ' === END MODIFICATION 2 ===

        ' If no valid POs were added, return the empty set
        If poCheckTable.Rows.Count = 0 Then
            Return existingPOs
        End If

        Using conn As New SqlConnection(connectionString)
            conn.Open()
            ' ใช้ Transaction เพื่อจัดการ Temp Table
            Using transaction As SqlTransaction = conn.BeginTransaction()
                Try
                    ' 2. สร้าง Temp Table
                    Using cmdCreateTemp As New SqlCommand("CREATE TABLE #TempCheckPOs (DraftPO_No VARCHAR(100) PRIMARY KEY)", conn, transaction)
                        cmdCreateTemp.ExecuteNonQuery()
                    End Using

                    ' 3. Bulk Insert POs to Temp Table
                    Using bulkCheck As New SqlBulkCopy(conn, SqlBulkCopyOptions.Default, transaction)
                        bulkCheck.DestinationTableName = "#TempCheckPOs"
                        bulkCheck.ColumnMappings.Add("DraftPO_No", "DraftPO_No")
                        bulkCheck.WriteToServer(poCheckTable)
                    End Using

                    ' 4. Select duplicates (อัปเดตชื่อตาราง)
                    Dim checkQuery As String = "SELECT T.DraftPO_No 
                                              FROM [dbo].[Draft_PO_Transaction] T 
                                              JOIN #TempCheckPOs TT ON T.DraftPO_No = TT.DraftPO_No"

                    Using cmdCheck As New SqlCommand(checkQuery, conn, transaction)
                        Using reader As SqlDataReader = cmdCheck.ExecuteReader()
                            While reader.Read()
                                existingPOs.Add(reader("DraftPO_No").ToString().Trim()) ' Trim data from DB just in case
                            End While
                        End Using
                    End Using

                    transaction.Commit() ' Commit เพื่อยืนยันการอ่าน (Temp table จะถูก drop เมื่อปิด conn)
                Catch ex As Exception
                    transaction.Rollback()
                    ' ถ้ามีข้อผิดพลาดในการตรวจสอบ ให้ส่ง Set ว่างกลับไป (เพื่อไม่ให้หน้า Preview พัง)
                    ' (ควร Log Error ไว้)
                    System.Diagnostics.Debug.WriteLine("Error in GetExistingPOs: " & ex.Message)
                End Try
            End Using
        End Using

        Return existingPOs
    End Function



    Private Function GenerateHtmlTable(dt As DataTable) As String
        Dim validator As POValidate = Nothing
        Try
            validator = New POValidate()
        Catch ex As Exception
            ' ถ้า validator มี error ให้ return error message
            Return $"<div class='alert alert-danger'>
                    <strong>Error creating validator:</strong><br/>
                    {HttpUtility.HtmlEncode(ex.Message)}<br/>
                    <small>Stack Trace: {HttpUtility.HtmlEncode(ex.StackTrace)}</small>
                 </div>"
        End Try

        ' ***  (ขั้นตอนที่ 1) PRE-CHECK: ดึง PO ทั้งหมดจาก Excel ***
        Dim poNosFromExcel As New List(Of String)
        If dt.Columns.Contains("Draft PO no.") Then
            For Each row As DataRow In dt.Rows
                If row("Draft PO no.") IsNot DBNull.Value Then
                    Dim po As String = row("Draft PO no.").ToString().Trim()
                    If Not String.IsNullOrEmpty(po) Then
                        poNosFromExcel.Add(po)
                    End If
                End If
            Next
        Else
            Return $"<div class='alert alert-danger'>Error: Missing required column 'Draft PO no.' in the file.</div>"
        End If


        Dim existingDbPOs As HashSet(Of String) = GetExistingPOs(poNosFromExcel)
        Dim sb As New StringBuilder()


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
        Dim duplicateInExcelChecker As New Dictionary(Of String, Integer)

        For i As Integer = 0 To dt.Rows.Count - 1
            Dim row As DataRow = dt.Rows(i)
            Dim errorMessages As New List(Of String)
            Dim isValid As Boolean = True

            ' ดึงค่าจาก Excel
            Dim PONOValue As String = If(row("Draft PO no.") IsNot DBNull.Value, row("Draft PO no.").ToString().Trim(), "")
            Dim yearValue As String = If(row("Year") IsNot DBNull.Value, row("Year").ToString().Trim(), "")
            Dim monthValue As String = If(row("Month") IsNot DBNull.Value, row("Month").ToString().Trim(), "")
            Dim categoryValue As String = If(row("Category") IsNot DBNull.Value, row("Category").ToString().Trim(), "")
            Dim companyValue As String = If(row("Company") IsNot DBNull.Value, row("Company").ToString().Trim(), "")
            Dim segmentValue As String = If(row("Segment") IsNot DBNull.Value, row("Segment").ToString().Trim(), "")
            Dim brandValue As String = If(row("Brand") IsNot DBNull.Value, row("Brand").ToString().Trim(), "")
            Dim vendorValue As String = If(row("Vendor") IsNot DBNull.Value, row("Vendor").ToString().Trim(), "")
            Dim amountTHBValue As String = If(row("Amount (THB)") IsNot DBNull.Value, row("Amount (THB)").ToString().Trim(), "")
            Dim amountCCYValue As String = If(row("Amount (CCY)") IsNot DBNull.Value, row("Amount (CCY)").ToString().Trim(), "")
            Dim ccyValue As String = If(row("CCY") IsNot DBNull.Value, row("CCY").ToString().Trim(), "")
            Dim exRateValue As String = If(row("Ex. Rate") IsNot DBNull.Value, row("Ex. Rate").ToString().Trim(), "")
            Dim RemarkValue As String = If(row("Remark") IsNot DBNull.Value, row("Remark").ToString().Trim(), "")


            If String.IsNullOrEmpty(PONOValue) Then
                errorMessages.Add("Draft PO no. is required")
                isValid = False
            ElseIf duplicateInExcelChecker.ContainsKey(PONOValue) Then
                errorMessages.Add("Duplicated_Draft PO_Excel")
                isValid = False
            ElseIf existingDbPOs.Contains(PONOValue) Then
                errorMessages.Add("Duplicated_Draft PO_Database")
                isValid = False
            Else
                duplicateInExcelChecker.Add(PONOValue, i)
            End If

            If isValid Then validCount += 1 Else errorCount += 1




            ' สร้างแถว
            Dim rowClass As String = ""
            If Not isValid Then
                rowClass = "table-danger" ' 
            End If
            sb.AppendFormat("<tr class='{0}' data-row-index='{1}'>", rowClass, i)

            ' Checkbox Column
            If isValid Then
                ' (โค้ดสำหรับ Checkbox ที่มี data-attributes)
                sb.AppendFormat("<td class='text-center'><input type='checkbox' name='selectedRows' class='form-check-input row-checkbox' value='{0}' " &
                    "data-pono='{1}' data-year='{2}' data-month='{3}' data-category='{4}' data-company='{5}' " &
                    "data-segment='{6}' data-brand='{7}' data-vendor='{8}' data-amountthb='{9}' data-amountccy='{10}' " &
                    "data-ccy='{11}' data-exrate='{12}' data-remark='{13}' checked></td>",
                    i,
                    HttpUtility.HtmlAttributeEncode(PONOValue),
                    HttpUtility.HtmlAttributeEncode(yearValue),
                    HttpUtility.HtmlAttributeEncode(monthValue),
                    HttpUtility.HtmlAttributeEncode(categoryValue),
                    HttpUtility.HtmlAttributeEncode(companyValue),
                    HttpUtility.HtmlAttributeEncode(segmentValue),
                    HttpUtility.HtmlAttributeEncode(brandValue),
                    HttpUtility.HtmlAttributeEncode(vendorValue),
                    HttpUtility.HtmlAttributeEncode(amountTHBValue),
                    HttpUtility.HtmlAttributeEncode(amountCCYValue),
                    HttpUtility.HtmlAttributeEncode(ccyValue),
                    HttpUtility.HtmlAttributeEncode(exRateValue),
                    HttpUtility.HtmlAttributeEncode(RemarkValue)
                )
            Else
                sb.Append("<td class='text-center'><input type='checkbox' class='form-check-input' disabled></td>")
            End If

            ' No. Column
            sb.AppendFormat("<td class='text-center'>{0}</td>", HttpUtility.HtmlEncode(PONOValue))

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

            ' Amount (TH)
            Try
                Dim amountTHBDec As Decimal = Convert.ToDecimal(amountTHBValue)
                sb.AppendFormat("<td class='text-end'>{0}</td>", amountTHBDec.ToString("N2"))
            Catch
                sb.AppendFormat("<td class='text-end'>{0}</td>", HttpUtility.HtmlEncode(amountTHBValue))
            End Try

            ' Amount (CCY)
            Try
                Dim amountCCYDec As Decimal = Convert.ToDecimal(amountCCYValue)
                sb.AppendFormat("<td class='text-end'>{0}</td>", amountCCYDec.ToString("N2"))
            Catch
                sb.AppendFormat("<td class='text-end'>{0}</td>", HttpUtility.HtmlEncode(amountCCYValue))
            End Try

            ' CCY
            sb.AppendFormat("<td class='text-end'>{0}</td>", HttpUtility.HtmlEncode(ccyValue))

            'Ex.Rate
            Try
                Dim exRateDec As Decimal = Convert.ToDecimal(exRateValue)
                sb.AppendFormat("<td class='text-end'>{0}</td>", exRateDec.ToString("N2")) ' (ปรับทศนิยมถ้าต้องการ)
            Catch
                sb.AppendFormat("<td class='text-end'>{0}</td>", HttpUtility.HtmlEncode(exRateValue))
            End Try

            ' Remark
            sb.AppendFormat("<td>{0}</td>", HttpUtility.HtmlEncode(RemarkValue))

            ' Error Column (เหมือนเดิม)
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



        ' JavaScript สำหรับจัดการ Checkbox
        sb.Append("<script>")
        sb.Append("$(document).ready(function() {")
        sb.Append("  /* Select All checkbox functionality */")
        sb.Append("  $('#selectAllCheckbox').on('change', function() {")
        sb.Append("    $('.row-checkbox:not(:disabled)').prop('checked', this.checked);")
        sb.Append("  });")
        sb.Append("  /* Update select all when individual checkbox changes */")
        sb.Append("  $('.row-checkbox').on('change', function() {")
        sb.Append("    var total = $('.row-checkbox:not(:disabled)').length;")
        sb.Append("    var checked = $('.row-checkbox:checked').length;")
        sb.Append("    $('#selectAllCheckbox').prop('checked', total === checked);")
        sb.Append("  });")
        sb.Append("});")
        sb.Append("</script>")

        Return sb.ToString()
    End Function

    Private Sub SaveFromPreview(jsonData As String, uploadBy As String, context As HttpContext)
        ' === 0. แปลง JSON ===
        If String.IsNullOrEmpty(jsonData) Then Throw New Exception("No data selected.")

        Dim serializer As New JavaScriptSerializer()
        Dim selectedRows As List(Of POPreviewRow) = serializer.Deserialize(Of List(Of POPreviewRow))(jsonData)

        If selectedRows.Count = 0 Then
            context.Response.Write("No rows were selected to save.")
            Return
        End If

        ' === 1. (Validator - ถ้ามี) ===
        Dim createDT As DateTime = DateTime.Now
        Dim Validator As New POValidate()

        ' === 3. เตรียม DataTables ===
        Dim insertTable As New DataTable()
        ' (คอลัมน์ทั้งหมดสำหรับ Insert)
        insertTable.Columns.Add("DraftPO_No", GetType(String))
        insertTable.Columns.Add("PO_Year", GetType(String)) ' Changed to Integer
        insertTable.Columns.Add("PO_Month", GetType(String)) ' Changed to Integer
        insertTable.Columns.Add("Company_Code", GetType(String))
        insertTable.Columns.Add("Category_Code", GetType(String))
        insertTable.Columns.Add("Segment_Code", GetType(String))
        insertTable.Columns.Add("Brand_Code", GetType(String))
        insertTable.Columns.Add("Vendor_Code", GetType(String))
        insertTable.Columns.Add("CCY", GetType(String))
        insertTable.Columns.Add("Exchange_Rate", GetType(Decimal))
        insertTable.Columns.Add("Amount_CCY", GetType(Decimal))
        insertTable.Columns.Add("Amount_THB", GetType(Decimal))
        insertTable.Columns.Add("PO_Type", GetType(String)) ' Added
        insertTable.Columns.Add("Status", GetType(String))  ' Added
        insertTable.Columns.Add("Remark", GetType(String))
        insertTable.Columns.Add("Created_By", GetType(String))
        insertTable.Columns.Add("Created_Date", GetType(DateTime))

        ' ตารางสำหรับตรวจสอบ PO No ที่ซ้ำ
        Dim poCheckTable As New DataTable()
        poCheckTable.Columns.Add("DraftPO_No", GetType(String))

        Dim savedCount As Integer = 0
        Dim errorList As New List(Of String)

        ' === 4. วนลูปข้อมูลที่ส่งมาจาก Preview ===
        For Each row As POPreviewRow In selectedRows
            Try
                ' === 4.0 Re-Validate Data ===
                Dim validationErrors As Dictionary(Of String, String) = Validator.ValidateDraftPO(
                    row.Year, row.Month, row.Company, row.Category, row.Segment, row.Brand, row.Vendor,
                    row.DraftPONo, row.AmountCCY, row.CCY, row.ExRate, row.AmountTHB,
                    checkDuplicate:=False ' We do duplicate check manually in bulk
                )

                If validationErrors.Count > 0 Then
                    errorList.Add($"{row.DraftPONo}: {String.Join(", ", validationErrors.Values)}")
                    Continue For ' Skip this row
                End If

                ' === 4.1 เพิ่มข้อมูลลงตาราง Insert ===
                Dim newRow As DataRow = insertTable.NewRow()
                newRow("DraftPO_No") = row.DraftPONo
                newRow("PO_Year") = Convert.ToInt32(row.Year)
                newRow("PO_Month") = Convert.ToInt32(row.Month)
                newRow("Category_Code") = row.Category
                newRow("Company_Code") = row.Company
                newRow("Segment_Code") = row.Segment
                newRow("Brand_Code") = row.Brand
                newRow("Vendor_Code") = row.Vendor
                newRow("CCY") = row.CCY
                newRow("Exchange_Rate") = Convert.ToDecimal(row.ExRate)
                newRow("Amount_CCY") = Convert.ToDecimal(row.AmountCCY)
                newRow("Amount_THB") = Convert.ToDecimal(row.AmountTHB)
                newRow("Remark") = row.Remark
                newRow("Created_By") = uploadBy
                newRow("Created_Date") = createDT
                newRow("PO_Type") = "Draft" ' Default value
                newRow("Status") = "Draft"   ' Default value

                insertTable.Rows.Add(newRow)

                ' === 4.2 เพิ่มข้อมูลลงตาราง Check ===
                poCheckTable.Rows.Add(row.DraftPONo)

            Catch ex As Exception
                ' Catch conversion errors etc.
                errorList.Add($"{row.DraftPONo}: {ex.Message}")
                Continue For
            End Try
        Next

        ' If all rows failed validation before DB check
        If insertTable.Rows.Count = 0 Then
            If errorList.Count > 0 Then
                context.Response.StatusCode = 400 ' Bad Request
                context.Response.Write("Data validation failed for all rows: " & String.Join("; ", errorList))
            Else
                context.Response.Write("No valid rows were selected to save.")
            End If
            Return
        End If


        ' === 5. Execute INSERT (ย้าย Transaction ขึ้นมาคลุมทั้งหมด) ===
        Using conn As New SqlConnection(connectionString)
            conn.Open()
            Using transaction As SqlTransaction = conn.BeginTransaction()
                Try
                    ' === 5A. สร้างตารางชั่วคราว ===
                    Using cmdCreateTemp As New SqlCommand("CREATE TABLE #TempCheckPOs (DraftPO_No VARCHAR(100) PRIMARY KEY)", conn, transaction)
                        cmdCreateTemp.ExecuteNonQuery()
                    End Using

                    ' === 5B. Bulk Insert POs ลงตารางชั่วคราว ===
                    Using bulkCheck As New SqlBulkCopy(conn, SqlBulkCopyOptions.Default, transaction)
                        bulkCheck.DestinationTableName = "#TempCheckPOs"
                        bulkCheck.ColumnMappings.Add("DraftPO_No", "DraftPO_No")
                        bulkCheck.WriteToServer(poCheckTable)
                    End Using

                    ' === 5C. ตรวจสอบข้อมูลซ้ำ (อัปเดตชื่อตาราง) ===
                    Dim duplicates As New List(Of String)
                    Dim checkQuery As String = "SELECT T.DraftPO_No 
                                              FROM [dbo].[Draft_PO_Transaction] T 
                                              JOIN #TempCheckPOs TT ON T.DraftPO_No = TT.DraftPO_No"

                    Using cmdCheck As New SqlCommand(checkQuery, conn, transaction)
                        Using reader As SqlDataReader = cmdCheck.ExecuteReader()
                            While reader.Read()
                                duplicates.Add(reader("DraftPO_No").ToString())
                            End While
                        End Using
                    End Using

                    ' === 5D. ถ้าพบข้อมูลซ้ำ ให้ Rollback และแจ้งเตือน ===
                    If duplicates.Count > 0 Then
                        transaction.Rollback()

                        Dim errorMsg As String = $"Error: Cannot save. The following Draft PO number(s) already exist in the database: {String.Join(", ", duplicates.Take(5))}"
                        If duplicates.Count > 5 Then errorMsg &= "..."

                        context.Response.StatusCode = 409 ' 409 Conflict
                        context.Response.Write(errorMsg)
                        Return
                    End If

                    ' === 5E. ถ้าไม่ซ้ำ ให้บันทึกข้อมูลจริง (อัปเดตชื่อตาราง) ===
                    If insertTable.Rows.Count > 0 Then
                        Using bulkCopy As New SqlBulkCopy(conn, SqlBulkCopyOptions.Default, transaction)
                            bulkCopy.DestinationTableName = "[dbo].[Draft_PO_Transaction]" ' 👈 อัปเดตชื่อตาราง
                            bulkCopy.BatchSize = insertTable.Rows.Count
                            bulkCopy.BulkCopyTimeout = 300

                            ' (อัปเดต Column Mappings)
                            For Each col As DataColumn In insertTable.Columns
                                bulkCopy.ColumnMappings.Add(col.ColumnName, col.ColumnName)
                            Next

                            bulkCopy.WriteToServer(insertTable)
                            savedCount = insertTable.Rows.Count ' Count successful rows
                        End Using
                    End If

                    ' === 5F. Commit Transaction ===
                    transaction.Commit()

                Catch ex As Exception
                    transaction.Rollback()
                    Throw New Exception("Error saving data: " & ex.Message)
                End Try
            End Using
        End Using

        Dim finalMessage As String = $"Successfully saved {savedCount} new rows to Draft PO."
        If errorList.Count > 0 Then
            finalMessage &= $" {errorList.Count} rows were skipped due to validation errors."
        End If

        context.Response.Write(finalMessage)
    End Sub

    ReadOnly Property IsReusable() As Boolean Implements IHttpHandler.IsReusable
        Get
            Return False
        End Get
    End Property

End Class