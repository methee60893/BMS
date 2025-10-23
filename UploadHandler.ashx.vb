Imports System
Imports System.Web
Imports System.Data
Imports System.IO
Imports System.Data.SqlClient
Imports System.Text
Imports ExcelDataReader
Imports System.Globalization

Public Class UploadHandler : Implements IHttpHandler

    Public Shared connectionString93 As String = "Data Source=10.3.0.93;Initial Catalog=BMS;Persist Security Info=True;User ID=sa;Password=sql2014"

    Public Sub ProcessRequest(ByVal context As HttpContext) Implements IHttpHandler.ProcessRequest
        context.Response.Clear()
        context.Response.ContentType = "text/html"
        context.Response.ContentEncoding = Encoding.UTF8

        '  รับค่า uploadBy จาก form data
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
        Dim sb As New StringBuilder()
        sb.Append("<div style='max-height:500px; overflow:auto;'>")
        sb.Append("<table class='table table-bordered table-striped'>")
        sb.Append("<thead><tr>")
        For Each col As DataColumn In dt.Columns
            sb.AppendFormat("<th>{0}</th>", HttpUtility.HtmlEncode(col.ColumnName))
        Next
        sb.Append("</tr></thead><tbody>")

        Dim maxRows As Integer = Math.Min(dt.Rows.Count, 100) ' แสดงแค่ 100 แถวแรก
        For i As Integer = 0 To maxRows - 1
            sb.Append("<tr>")
            For j As Integer = 0 To dt.Columns.Count - 1
                Dim val As String = If(dt.Rows(i)(j) IsNot DBNull.Value, dt.Rows(i)(j).ToString(), "")
                sb.AppendFormat("<td>{0}</td>", HttpUtility.HtmlEncode(val))
            Next
            sb.Append("</tr>")
        Next

        sb.Append("</tbody></table>")
        If dt.Rows.Count > 100 Then
            sb.AppendFormat("<p class='text-muted'>Showing first 100 of {0} rows.</p>", dt.Rows.Count)
        End If
        sb.Append("</div>")
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

        ' === 2. ดึง Batch ใหม่ ===
        Dim newBatch As String = GetNextBatchNumber()
        Dim createDT As DateTime = DateTime.Now

        ' === 3. สร้าง DataTable สำหรับ Bulk Insert ===
        Dim uploadTable As New DataTable()
        uploadTable.Columns.Add("Type", GetType(String))
        uploadTable.Columns.Add("Year", GetType(String))
        uploadTable.Columns.Add("Month", GetType(String))
        uploadTable.Columns.Add("Category", GetType(String))
        uploadTable.Columns.Add("Company", GetType(String))
        uploadTable.Columns.Add("Segment", GetType(String))
        uploadTable.Columns.Add("Brand", GetType(String))
        uploadTable.Columns.Add("Vendor", GetType(String))
        uploadTable.Columns.Add("Amount", GetType(String))
        uploadTable.Columns.Add("UploadBy", GetType(String))
        uploadTable.Columns.Add("Batch", GetType(String))      '  เพิ่ม Batch
        uploadTable.Columns.Add("CreateDT", GetType(DateTime)) '  เพิ่ม CreateDT

        ' === 4. เติมข้อมูลทุกแถว ===
        For Each row As DataRow In dt.Rows
            Dim newRow As DataRow = uploadTable.NewRow()
            newRow("Type") = If(row("Type") IsNot DBNull.Value, row("Type").ToString(), "")
            newRow("Year") = If(row("Year") IsNot DBNull.Value, row("Year").ToString(), "")
            newRow("Month") = If(row("Month") IsNot DBNull.Value, row("Month").ToString(), "")
            newRow("Category") = If(row("Category") IsNot DBNull.Value, row("Category").ToString(), "")
            newRow("Company") = If(row("Company") IsNot DBNull.Value, row("Company").ToString(), "")
            newRow("Segment") = If(row("Segment") IsNot DBNull.Value, row("Segment").ToString(), "")
            newRow("Brand") = If(row("Brand") IsNot DBNull.Value, row("Brand").ToString(), "")
            newRow("Vendor") = If(row("Vendor") IsNot DBNull.Value, row("Vendor").ToString(), "")
            newRow("Amount") = If(row("Amount") IsNot DBNull.Value, row("Amount").ToString(), "")
            newRow("UploadBy") = uploadBy
            newRow("Batch") = newBatch          '  ทุกแถวได้ Batch เดียวกัน
            newRow("CreateDT") = createDT       '  เวลาเดียวกันทั้ง batch
            uploadTable.Rows.Add(newRow)
        Next

        ' === 5. Bulk Insert ===
        Using conn As New SqlConnection(connectionString93)
            conn.Open()
            Using bulkCopy As New SqlBulkCopy(conn, SqlBulkCopyOptions.CheckConstraints, Nothing)
                bulkCopy.DestinationTableName = "[dbo].[Template_Upload_Draft_OTB]"
                bulkCopy.BatchSize = uploadTable.Rows.Count
                bulkCopy.BulkCopyTimeout = 300

                ' แมปคอลัมน์
                For Each col As DataColumn In uploadTable.Columns
                    bulkCopy.ColumnMappings.Add(col.ColumnName, col.ColumnName)
                Next

                bulkCopy.WriteToServer(uploadTable)
            End Using
        End Using
    End Sub

    Public ReadOnly Property IsReusable() As Boolean Implements IHttpHandler.IsReusable
        Get
            Return False
        End Get
    End Property

    Private Function GetNextBatchNumber() As String
        Dim connectionString As String = connectionString93
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

End Class