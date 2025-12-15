Imports System
Imports System.Data
Imports System.Data.SqlClient
Imports System.Globalization
Imports System.IO
Imports System.Linq
Imports System.Text
Imports System.Web
Imports System.Web.Script.Serialization
Imports BMS
Imports ExcelDataReader
Imports Newtonsoft.Json

' Class สำหรับรับ-ส่งข้อมูล JSON
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

Public Class RowSaveResult
    Public Property DraftPONo As String
    Public Property Status As String ' "Success" or "Error"
    Public Property Message As String
End Class

Public Class POUploadHandler
    Implements IHttpHandler, IRequiresSessionState

    Private Shared connectionString As String = ConfigurationManager.ConnectionStrings("BMSConnectionString")?.ConnectionString

    Sub ProcessRequest(ByVal context As HttpContext) Implements IHttpHandler.ProcessRequest
        context.Response.Clear()
        context.Response.ContentType = "text/html"
        context.Response.ContentEncoding = Encoding.UTF8

        Dim uploadBy As String = context.Request.Form("uploadBy")
        If String.IsNullOrEmpty(uploadBy) Then uploadBy = "unknown"

        Dim action As String = context.Request("action")

        ' --- ACTION: Save Data ---
        If action = "savePreview" Then
            Try
                Dim jsonData As String = context.Request.Form("selectedData")
                SaveFromPreview(jsonData, uploadBy, context)
            Catch ex As Exception
                context.Response.StatusCode = 200
                context.Response.Write($"<div class='alert alert-danger'>Error: {HttpUtility.HtmlEncode(ex.Message)}</div>")
            End Try
            Return
        End If

        ' --- ACTION: Upload & Preview ---
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

            If dt Is Nothing OrElse dt.Rows.Count = 0 Then
                context.Response.Write("<div class='alert alert-warning'>Uploaded file contains no data.</div>")
                Return
            End If

            If context.Request("action") = "preview" Then
                context.Response.Write(GenerateHtmlTable(dt))
            End If

        Catch ex As Exception
            context.Response.StatusCode = 500
            context.Response.Write($"<div class='alert alert-danger'>Error processing file: {HttpUtility.HtmlEncode(ex.Message)}</div>")
        Finally
            If File.Exists(tempPath) Then File.Delete(tempPath)
        End Try
    End Sub

    ' --- Helper: อ่านค่าจาก DataRow อย่างปลอดภัย (ป้องกัน Crash) ---
    Private Function GetSafeString(row As DataRow, colName As String) As String
        ' 1. เช็คว่ามีคอลัมน์นี้จริงหรือไม่
        If Not row.Table.Columns.Contains(colName) Then
            Return "" ' ถ้าไม่มี ให้คืนค่าว่าง (ระบบจะไม่ Crash แต่จะไปติด Validate สีแดงที่หน้าจอแทน)
        End If
        ' 2. ถ้ามี ให้ดึงค่า
        If row(colName) IsNot DBNull.Value Then
            Return row(colName).ToString().Trim()
        End If
        Return ""
    End Function

    ' --- Main Logic: สร้างตาราง HTML ---
    Private Function GenerateHtmlTable(dt As DataTable) As String
        Dim validator As POValidate = Nothing
        Try
            validator = New POValidate()
        Catch ex As Exception
            Return $"<div class='alert alert-danger'>Error loading validator: {ex.Message}</div>"
        End Try

        ' 1. เตรียมข้อมูล
        Dim previewList As New List(Of POPreviewRow)
        Dim poNosInFile As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
        Dim budgetGroups As New Dictionary(Of POMatchKey, Decimal)

        For Each row As DataRow In dt.Rows
            ' ใช้ GetSafeString แทนการเรียก row("...") ตรงๆ เพื่อป้องกัน Error Column Missing
            Dim item As New POPreviewRow With {
                .DraftPONo = GetSafeString(row, "Draft PO no."),
                .Year = GetSafeString(row, "Year"),
                .Month = GetSafeString(row, "Month"),
                .Category = GetSafeString(row, "Category"),
                .Company = GetSafeString(row, "Company"),
                .Segment = GetSafeString(row, "Segment"),
                .Brand = GetSafeString(row, "Brand"),
                .Vendor = GetSafeString(row, "Vendor"),
                .AmountCCY = GetSafeString(row, "Amount (CCY)"),
                .CCY = GetSafeString(row, "CCY"),
                .ExRate = GetSafeString(row, "Ex. Rate"),
                .Remark = GetSafeString(row, "Remark")
            }

            ' คำนวณ THB
            Dim amtCCY As Decimal = 0
            Dim exRate As Decimal = 0
            Decimal.TryParse(item.AmountCCY.Replace(",", ""), NumberStyles.Any, CultureInfo.InvariantCulture, amtCCY)
            Decimal.TryParse(item.ExRate.Replace(",", ""), NumberStyles.Any, CultureInfo.InvariantCulture, exRate)
            Dim amtTHB As Decimal = amtCCY * exRate
            item.AmountTHB = amtTHB.ToString("0.00")

            ' รวมกลุ่ม Budget
            Dim key As New POMatchKey With {
                .Year = item.Year, .Month = item.Month, .Company = item.Company,
                .Category = item.Category, .Segment = item.Segment, .Brand = item.Brand, .Vendor = item.Vendor
            }
            If budgetGroups.ContainsKey(key) Then
                budgetGroups(key) += amtTHB
            Else
                budgetGroups.Add(key, amtTHB)
            End If

            If Not String.IsNullOrEmpty(item.DraftPONo) Then poNosInFile.Add(item.DraftPONo)
            previewList.Add(item)
        Next

        ' 2. ตรวจสอบข้อมูล (Validate)
        'Dim existingDbPOs As HashSet(Of String) = GetExistingPOs(poNosInFile.ToList())
        Dim existingDbKeys As HashSet(Of String) = GetExistingCompositeKeys(previewList)
        Dim budgetFailedKeys As New HashSet(Of POMatchKey)

        ' เช็ค Budget รวม
        For Each kvp In budgetGroups
            Dim remaining As Decimal = validator.GetRemainingBudget(kvp.Key)
            If kvp.Value > remaining Then
                budgetFailedKeys.Add(kvp.Key)
            End If
        Next

        ' 3. สร้าง HTML Table
        Dim sb As New StringBuilder()
        sb.Append("<div class='table-responsive' style='max-height:600px; overflow:auto;'>")
        sb.Append("<table id='previewTable' class='table table-bordered table-striped table-sm table-hover'>")
        sb.Append("<thead class='table-primary sticky-header'><tr>")
        sb.Append("<th class='text-center' style='width:50px;'>Select</th>")
        sb.Append("<th>Draft PO No.</th><th>Year</th><th>Month</th><th>Category</th><th>Company</th><th>Segment</th><th>Brand</th><th>Vendor</th>")
        sb.Append("<th class='text-end'>Amount (THB)</th><th class='text-end'>Amount (CCY)</th><th class='text-center'>CCY</th><th class='text-end'>Ex. Rate</th>")
        sb.Append("<th>Remark</th><th class='text-danger' style='min-width:200px;'>Error</th></tr></thead><tbody>")

        Dim totalErrorCount As Integer = 0
        Dim duplicateInFileCount As New Dictionary(Of String, Integer)

        For i As Integer = 0 To previewList.Count - 1
            Dim item = previewList(i)
            Dim errors As New List(Of String)
            Dim rowStyle As String = ""
            Dim isDuplicate As Boolean = False
            Dim isBudgetFail As Boolean = False

            Dim currentKey As String = String.Format("{0}|{1}|{2}|{3}|{4}|{5}|{6}|{7}",
                                              item.DraftPONo.Trim(), item.Year.Trim(), item.Month.Trim(),
                                              item.Company.Trim(), item.Category.Trim(), item.Segment.Trim(),
                                              item.Brand.Trim(), item.Vendor.Trim())

            '' Check Duplicate PO (Req 2)
            'If duplicateInFileCount.ContainsKey(item.DraftPONo) Then
            '    duplicateInFileCount(item.DraftPONo) += 1
            '    isDuplicate = True
            '    errors.Add("Duplicate PO No. in file")
            'Else
            '    duplicateInFileCount.Add(item.DraftPONo, 1)
            'End If
            'If existingDbPOs.Contains(item.DraftPONo) Then
            '    isDuplicate = True
            '    errors.Add("PO No. already exists in Database")
            'End If

            ' Check Duplicate In File (เช็คซ้ำในไฟล์เดียวกันเอง)
            If duplicateInFileCount.ContainsKey(currentKey) Then
                duplicateInFileCount(currentKey) += 1
                isDuplicate = True
                errors.Add("Duplicate Data in file")
            Else
                duplicateInFileCount.Add(currentKey, 1)
            End If

            ' Check Duplicate In Database (เช็คซ้ำกับ DB โดยใช้ Composite Key)
            If existingDbKeys.Contains(currentKey) Then
                isDuplicate = True
                errors.Add("Data already exists in Database") ' เปลี่ยนข้อความให้สื่อความหมาย
            End If

            ' Check Budget (Req 3 & 4)
            Dim key As New POMatchKey With {
                .Year = item.Year, .Month = item.Month, .Company = item.Company,
                .Category = item.Category, .Segment = item.Segment, .Brand = item.Brand, .Vendor = item.Vendor
            }
            If budgetFailedKeys.Contains(key) Then
                isBudgetFail = True
                errors.Add("Insufficient Budget")
            End If

            ' Check Fields (Req 5 & 7)
            Dim amtCCY As Decimal = 0
            Dim exRate As Decimal = 0
            If Not Decimal.TryParse(item.AmountCCY.Replace(",", ""), NumberStyles.Any, CultureInfo.InvariantCulture, amtCCY) Then
                errors.Add("Invalid Amount format")
            End If
            If Not Decimal.TryParse(item.ExRate.Replace(",", ""), NumberStyles.Any, CultureInfo.InvariantCulture, exRate) Then
                errors.Add("Invalid Ex.Rate format")
            End If

            If amtCCY < 0 Then errors.Add("Amount (CCY) cannot be negative")
            If exRate < 0 Then errors.Add("Ex. Rate cannot be negative")

            If String.IsNullOrEmpty(item.CCY) Then
                errors.Add("CCY is required")
            ElseIf Not validator.ValidateVendorCCY(item.Vendor, item.CCY) Then
                errors.Add("Not found CCY in master")
            End If

            If String.IsNullOrEmpty(item.DraftPONo) Then errors.Add("PO No. is required")

            If errors.Count > 0 Then totalErrorCount += 1

            ' Determine Row Style
            If isBudgetFail Then
                rowStyle = "table-danger"
            ElseIf isDuplicate Then
                rowStyle = "table-warning"
            ElseIf errors.Count > 0 Then
                rowStyle = "table-danger"
            End If

            ' Render Row
            sb.Append($"<tr class='{rowStyle}' data-row-index='{i}'>")

            If errors.Count > 0 Then
                sb.Append("<td class='text-center'><input type='checkbox' disabled></td>")
            Else
                sb.AppendFormat("<td class='text-center'><input type='checkbox' name='selectedRows' class='form-check-input row-checkbox' value='{0}' " &
                   "data-pono='{1}' data-year='{2}' data-month='{3}' data-category='{4}' data-company='{5}' " &
                   "data-segment='{6}' data-brand='{7}' data-vendor='{8}' data-amountthb='{9}' data-amountccy='{10}' " &
                   "data-ccy='{11}' data-exrate='{12}' data-remark='{13}' checked></td>",
                   i,
                   HttpUtility.HtmlAttributeEncode(item.DraftPONo),
                   HttpUtility.HtmlAttributeEncode(item.Year),
                   HttpUtility.HtmlAttributeEncode(item.Month),
                   HttpUtility.HtmlAttributeEncode(item.Category),
                   HttpUtility.HtmlAttributeEncode(item.Company),
                   HttpUtility.HtmlAttributeEncode(item.Segment),
                   HttpUtility.HtmlAttributeEncode(item.Brand),
                   HttpUtility.HtmlAttributeEncode(item.Vendor),
                   HttpUtility.HtmlAttributeEncode(item.AmountTHB),
                   HttpUtility.HtmlAttributeEncode(item.AmountCCY),
                   HttpUtility.HtmlAttributeEncode(item.CCY),
                   HttpUtility.HtmlAttributeEncode(item.ExRate),
                   HttpUtility.HtmlAttributeEncode(item.Remark))
            End If

            sb.Append($"<td>{item.DraftPONo}</td>")
            sb.Append($"<td>{item.Year}</td>")
            sb.Append($"<td>{item.Month}</td>")
            sb.Append($"<td>{item.Category}</td>")
            sb.Append($"<td>{item.Company}</td>")
            sb.Append($"<td>{item.Segment}</td>")
            sb.Append($"<td>{item.Brand}</td>")
            sb.Append($"<td>{item.Vendor}</td>")
            sb.Append($"<td class='text-end'>{Decimal.Parse(item.AmountTHB).ToString("N2")}</td>")
            sb.Append($"<td class='text-end'>{Decimal.Parse(item.AmountCCY.Replace(",", "")).ToString("N2")}</td>")
            sb.Append($"<td class='text-center'>{item.CCY}</td>")
            sb.Append($"<td class='text-end'>{Decimal.Parse(item.ExRate.Replace(",", "")).ToString("N4")}</td>")
            sb.Append($"<td>{item.Remark}</td>")

            Dim errorText As String = If(errors.Count > 0, String.Join(", ", errors), "")
            sb.Append($"<td class='text-danger small'>{errorText}</td>")
            sb.Append("</tr>")
        Next

        sb.Append("</tbody></table></div>")

        ' Control Submit Button Script
        sb.Append("<script>")
        sb.Append("$(document).ready(function() {")
        sb.AppendFormat("  $('#btnSubmitData').prop('disabled', {0});", If(totalErrorCount > 0, "true", "false"))
        sb.Append("});")
        sb.Append("</script>")

        Return sb.ToString()
    End Function

    ' --- Function อ่าน Excel ---
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
            Dim headerLine As String = reader.ReadLine()
            If headerLine IsNot Nothing Then
                Dim headers As String() = headerLine.Split(","c)
                For Each header In headers
                    dt.Columns.Add(header.Trim())
                Next
                While Not reader.EndOfStream
                    Dim line As String = reader.ReadLine()
                    Dim values As String() = SplitCsvLine(line)
                    dt.Rows.Add(values)
                End While
            End If
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

    ' --- [BMS Gem MODIFICATION] ---
    ' เปลี่ยนจากการเช็คแค่ PO No เป็นการเช็ค Composite Key (PO + 7 Keys)
    Private Function GetExistingCompositeKeys(previewRows As List(Of POPreviewRow)) As HashSet(Of String)
        Dim existingKeys As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
        If previewRows.Count = 0 Then Return existingKeys

        ' 1. สร้าง DataTable สำหรับ Bulk Copy ไปยัง Temp Table
        Dim dtCheck As New DataTable()
        dtCheck.Columns.Add("DraftPO_No", GetType(String))
        dtCheck.Columns.Add("PO_Year", GetType(String))
        dtCheck.Columns.Add("PO_Month", GetType(String))
        dtCheck.Columns.Add("Company_Code", GetType(String))
        dtCheck.Columns.Add("Category_Code", GetType(String))
        dtCheck.Columns.Add("Segment_Code", GetType(String))
        dtCheck.Columns.Add("Brand_Code", GetType(String))
        dtCheck.Columns.Add("Vendor_Code", GetType(String))

        ' เติมข้อมูลลง DataTable (Distinct เพื่อลดปริมาณข้อมูล)
        ' ใช้ Composite String เพื่อทำ Distinct ใน Memory ก่อนส่งไป DB
        Dim uniqueRows As New HashSet(Of String)

        For Each row In previewRows
            ' Normalize ข้อมูล (Trim) เพื่อความแม่นยำ
            Dim key As String = String.Format("{0}|{1}|{2}|{3}|{4}|{5}|{6}|{7}",
                                              row.DraftPONo.Trim(), row.Year.Trim(), row.Month.Trim(),
                                              row.Company.Trim(), row.Category.Trim(), row.Segment.Trim(),
                                              row.Brand.Trim(), row.Vendor.Trim())

            If Not uniqueRows.Contains(key) Then
                uniqueRows.Add(key)
                dtCheck.Rows.Add(row.DraftPONo.Trim(), row.Year.Trim(), row.Month.Trim(),
                                 row.Company.Trim(), row.Category.Trim(), row.Segment.Trim(),
                                 row.Brand.Trim(), row.Vendor.Trim())
            End If
        Next

        Using conn As New SqlConnection(connectionString)
            conn.Open()
            Using transaction As SqlTransaction = conn.BeginTransaction()
                Try
                    ' 2. สร้าง Temp Table ที่มี structure ครบทุก Key
                    Dim createSql As String = "CREATE TABLE #TempCheckFull (
                                                [DraftPO_No] NVARCHAR(50),
                                                [PO_Year] NVARCHAR(10),
                                                [PO_Month] NVARCHAR(10),
                                                [Company_Code] NVARCHAR(20),
                                                [Category_Code] NVARCHAR(20),
                                                [Segment_Code] NVARCHAR(20),
                                                [Brand_Code] NVARCHAR(20),
                                                [Vendor_Code] NVARCHAR(20)
                                              )"
                    Using cmdCreate As New SqlCommand(createSql, conn, transaction)
                        cmdCreate.ExecuteNonQuery()
                    End Using

                    ' 3. Bulk Copy ข้อมูลที่ Upload เข้า Temp Table
                    Using bulkCopy As New SqlBulkCopy(conn, SqlBulkCopyOptions.Default, transaction)
                        bulkCopy.DestinationTableName = "#TempCheckFull"
                        For Each col As DataColumn In dtCheck.Columns
                            bulkCopy.ColumnMappings.Add(col.ColumnName, col.ColumnName)
                        Next
                        bulkCopy.WriteToServer(dtCheck)
                    End Using

                    ' 4. Query Join เพื่อหาตัวที่ซ้ำจริงๆ (Composite Key Match)
                    ' ต้องไม่รวม Status = 'Cancelled'
                    Dim query As String = "SELECT 
                                                T.DraftPO_No, T.PO_Year, T.PO_Month, T.Company_Code, 
                                                T.Category_Code, T.Segment_Code, T.Brand_Code, T.Vendor_Code
                                           FROM [BMS].[dbo].[Draft_PO_Transaction] T 
                                           INNER JOIN #TempCheckFull TMP 
                                           ON T.DraftPO_No = TMP.DraftPO_No
                                              AND T.PO_Year = TMP.PO_Year
                                              AND T.PO_Month = TMP.PO_Month
                                              AND T.Company_Code = TMP.Company_Code
                                              AND T.Category_Code = TMP.Category_Code
                                              AND T.Segment_Code = TMP.Segment_Code
                                              AND T.Brand_Code = TMP.Brand_Code
                                              AND T.Vendor_Code = TMP.Vendor_Code
                                           WHERE ISNULL(T.Status, '') <> 'Cancelled'"

                    Using cmd As New SqlCommand(query, conn, transaction)
                        Using reader As SqlDataReader = cmd.ExecuteReader()
                            While reader.Read()
                                ' สร้าง Key จากผลลัพธ์ที่เจอใน DB กลับมาเพื่อเช็ค
                                Dim dbKey As String = String.Format("{0}|{1}|{2}|{3}|{4}|{5}|{6}|{7}",
                                              reader("DraftPO_No").ToString().Trim(),
                                              reader("PO_Year").ToString().Trim(),
                                              reader("PO_Month").ToString().Trim(),
                                              reader("Company_Code").ToString().Trim(),
                                              reader("Category_Code").ToString().Trim(),
                                              reader("Segment_Code").ToString().Trim(),
                                              reader("Brand_Code").ToString().Trim(),
                                              reader("Vendor_Code").ToString().Trim())

                                existingKeys.Add(dbKey)
                            End While
                        End Using
                    End Using
                    transaction.Commit()
                Catch
                    transaction.Rollback()
                    Throw
                End Try
            End Using
        End Using

        Return existingKeys
    End Function
    Private Function GetExistingPOs(poNos As List(Of String)) As HashSet(Of String)
        Dim existingPOs As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
        If poNos.Count = 0 Then Return existingPOs

        Dim uniquePOs As New HashSet(Of String)(poNos, StringComparer.OrdinalIgnoreCase)
        Dim poCheckTable As New DataTable()
        poCheckTable.Columns.Add("DraftPO_No", GetType(String))
        For Each po In uniquePOs
            poCheckTable.Rows.Add(po)
        Next

        Using conn As New SqlConnection(connectionString)
            conn.Open()
            Using transaction As SqlTransaction = conn.BeginTransaction()
                Try
                    Using cmdCreate As New SqlCommand("CREATE TABLE #TempCheckPOs (DraftPO_No NVARCHAR(100))", conn, transaction)
                        cmdCreate.ExecuteNonQuery()
                    End Using
                    Using bulkCopy As New SqlBulkCopy(conn, SqlBulkCopyOptions.Default, transaction)
                        bulkCopy.DestinationTableName = "#TempCheckPOs"
                        bulkCopy.WriteToServer(poCheckTable)
                    End Using
                    Dim query As String = "SELECT T.DraftPO_No FROM [BMS].[dbo].[Draft_PO_Transaction] T INNER JOIN #TempCheckPOs TMP ON T.DraftPO_No = TMP.DraftPO_No"
                    Using cmd As New SqlCommand(query, conn, transaction)
                        Using reader As SqlDataReader = cmd.ExecuteReader()
                            While reader.Read()
                                existingPOs.Add(reader("DraftPO_No").ToString())
                            End While
                        End Using
                    End Using
                    transaction.Commit()
                Catch
                    transaction.Rollback()
                End Try
            End Using
        End Using
        Return existingPOs
    End Function

    Private Sub SaveFromPreview(jsonData As String, uploadBy As String, context As HttpContext)
        If String.IsNullOrEmpty(jsonData) Then Throw New Exception("No data selected.")

        context.Response.ContentType = "application/json" ' เปลี่ยน ContentType เป็น JSON
        Dim results As New List(Of RowSaveResult)

        Dim serializer As New JavaScriptSerializer()
        Dim selectedRows As List(Of POPreviewRow) = serializer.Deserialize(Of List(Of POPreviewRow))(jsonData)

        If selectedRows.Count = 0 Then
            context.Response.Write("No rows to save.")
            Return
        End If

        Dim Validator As New POValidate()
        For Each row As POPreviewRow In selectedRows
            Dim result As New RowSaveResult With {.DraftPONo = row.DraftPONo}

            Try
                Dim amtCCY As Decimal = Decimal.Parse(row.AmountCCY)
                Dim exRate As Decimal = Decimal.Parse(row.ExRate)
                Dim amtTHB As Decimal = Decimal.Parse(row.AmountTHB)

                ' 1. Validate
                Dim errors As Dictionary(Of String, String) = Validator.ValidateDraftPOCreation(
                    row.Year, row.Month, row.Company, row.Category, row.Segment, row.Brand, row.Vendor,
                    row.DraftPONo, amtCCY, row.CCY, exRate, amtTHB
                )

                If errors.Count > 0 Then
                    result.Status = "Error"
                    result.Message = String.Join(", ", errors.Values)
                Else
                    ' 2. Insert ลง DB (ถ้าผ่าน)
                    Using conn As New SqlConnection(connectionString)
                        conn.Open()
                        Dim query As String = "INSERT INTO [BMS].[dbo].[Draft_PO_Transaction] " &
                                            "([DraftPO_No], [PO_Year], [PO_Month], [Company_Code], [Category_Code], [Segment_Code], [Brand_Code], [Vendor_Code], " &
                                            "[CCY], [Exchange_Rate], [Amount_CCY], [Amount_THB], [PO_Type], [Status], [Status_Date], [Status_By], [Remark], [Created_By], [Created_Date]) " &
                                            "VALUES (@PoNo, @Year, @Month, @Company, @Category, @Segment, @Brand, @Vendor, " &
                                            "@CCY, @ExRate, @AmtCCY, @AmtTHB, 'Upload', 'Draft', GETDATE(), @User, @Remark, @User, GETDATE())"

                        Using cmd As New SqlCommand(query, conn)
                            cmd.Parameters.AddWithValue("@PoNo", row.DraftPONo)
                            cmd.Parameters.AddWithValue("@Year", row.Year)
                            cmd.Parameters.AddWithValue("@Month", row.Month)
                            cmd.Parameters.AddWithValue("@Company", row.Company)
                            cmd.Parameters.AddWithValue("@Category", row.Category)
                            cmd.Parameters.AddWithValue("@Segment", row.Segment)
                            cmd.Parameters.AddWithValue("@Brand", row.Brand)
                            cmd.Parameters.AddWithValue("@Vendor", row.Vendor)
                            cmd.Parameters.AddWithValue("@CCY", row.CCY)
                            cmd.Parameters.AddWithValue("@ExRate", Decimal.Parse(row.ExRate))
                            cmd.Parameters.AddWithValue("@AmtCCY", Decimal.Parse(row.AmountCCY))
                            cmd.Parameters.AddWithValue("@AmtTHB", Decimal.Parse(row.AmountTHB))
                            cmd.Parameters.AddWithValue("@User", uploadBy)
                            cmd.Parameters.AddWithValue("@Remark", row.Remark)
                            cmd.ExecuteNonQuery()
                        End Using
                    End Using

                    result.Status = "Success"
                    result.Message = "Saved successfully"
                End If

            Catch ex As Exception
                result.Status = "Error"
                result.Message = ex.Message
            End Try

            results.Add(result)
        Next

        ' ส่งผลลัพธ์กลับเป็น JSON List
        context.Response.Write(JsonConvert.SerializeObject(results))

    End Sub

    ReadOnly Property IsReusable() As Boolean Implements IHttpHandler.IsReusable
        Get
            Return False
        End Get
    End Property

End Class