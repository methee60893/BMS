Imports System.Data
Imports System.Data.SqlClient
Imports System.Text

Public Class ApprovedOTBManager
    Private Shared connectionString As String = ConfigurationManager.ConnectionStrings("BMSConnectionString")?.ConnectionString
    ''' <summary>
    ''' Approve Draft OTB โดยส่ง IDs
    ''' </summary>
    Public Shared Function ApproveDraftOTB(draftIDs As List(Of Integer), approvedBy As String,
                                          Optional remark As String = Nothing) As Dictionary(Of String, Object)
        Dim result As New Dictionary(Of String, Object)

        Try
            ' แปลง List to comma-separated string
            Dim idsString As String = String.Join(",", draftIDs)

            Using conn As New SqlConnection(connectionString)
                conn.Open()

                Using cmd As New SqlCommand("SP_Approve_Draft_OTB", conn)
                    cmd.CommandType = CommandType.StoredProcedure
                    cmd.CommandTimeout = 300

                    cmd.Parameters.AddWithValue("@DraftIDs", idsString)
                    cmd.Parameters.AddWithValue("@ApprovedBy", approvedBy)
                    cmd.Parameters.AddWithValue("@Remark", If(String.IsNullOrEmpty(remark), DBNull.Value, remark))

                    Using reader As SqlDataReader = cmd.ExecuteReader()
                        If reader.Read() Then
                            result.Add("ApprovedCount", If(reader("ApprovedCount") IsNot DBNull.Value, Convert.ToInt32(reader("ApprovedCount")), 0))
                            result.Add("Status", If(reader("Status") IsNot DBNull.Value, reader("Status").ToString(), ""))

                            If reader.FieldCount > 2 AndAlso reader("Status").ToString() = "Success" Then
                                result.Add("SAPDate", If(reader("SAPDate") IsNot DBNull.Value, Convert.ToDateTime(reader("SAPDate")), DateTime.Now))
                            ElseIf reader("Status").ToString() = "Error" Then
                                result.Add("ErrorMessage", If(reader("ErrorMessage") IsNot DBNull.Value, reader("ErrorMessage").ToString(), "Unknown error"))
                            End If
                        End If
                    End Using
                End Using
            End Using

        Catch ex As Exception
            result.Add("Status", "Error")
            result.Add("ErrorMessage", ex.Message)
            result.Add("ApprovedCount", 0)
        End Try

        Return result
    End Function

    Public Shared Function DeleteDraftOTB(runNos As List(Of Integer)) As Dictionary(Of String, Object)
        Dim result As New Dictionary(Of String, Object)

        Try
            ' แปลง List to comma-separated string
            Dim idsString As String = String.Join(",", runNos)

            Using conn As New SqlConnection(connectionString)
                conn.Open()

                Using cmd As New SqlCommand("SP_Deleted_Draft_OTB", conn)
                    cmd.CommandType = CommandType.StoredProcedure
                    cmd.CommandTimeout = 300

                    cmd.Parameters.AddWithValue("@runNos", idsString)


                    Using reader As SqlDataReader = cmd.ExecuteReader()
                        If reader.Read() Then
                            result.Add("DeletedCount", If(reader("DeletedCount") IsNot DBNull.Value, Convert.ToInt32(reader("DeletedCount")), 0))
                            result.Add("Status", If(reader("Status") IsNot DBNull.Value, reader("Status").ToString(), ""))
                            If reader("Status").ToString() = "Error" Then
                                result.Add("Message", If(reader("Message") IsNot DBNull.Value, reader("Message").ToString(), "Unknown error"))
                            End If
                        End If
                    End Using
                End Using
            End Using

        Catch ex As Exception
            result.Add("Status", "Error")
            result.Add("ErrorMessage", ex.Message)
            result.Add("DeletedCount", 0)
        End Try

        Return result
    End Function

    ''' <summary>
    ''' Search Approved OTB
    ''' </summary>
    Public Shared Function SearchApprovedOTB(
        Optional type As String = Nothing,
        Optional year As String = Nothing,
        Optional month As String = Nothing,
        Optional company As String = Nothing,
        Optional category As String = Nothing,
        Optional segment As String = Nothing,
        Optional brand As String = Nothing,
        Optional vendor As String = Nothing,
        Optional version As String = Nothing) As DataTable

        Dim dateFrom As DateTime? = Nothing
        Dim dateTo As DateTime? = Nothing
        Dim status As String = "Approved"
        Dim dt As New DataTable()

        Try
            Using conn As New SqlConnection(connectionString)
                conn.Open()

                Using cmd As New SqlCommand("SP_Search_Approved_OTB", conn)
                    cmd.CommandType = CommandType.StoredProcedure
                    cmd.CommandTimeout = 300

                    cmd.Parameters.AddWithValue("@Type", If(String.IsNullOrEmpty(type), DBNull.Value, type))
                    cmd.Parameters.AddWithValue("@Year", If(String.IsNullOrEmpty(year), DBNull.Value, year))
                    cmd.Parameters.AddWithValue("@Month", If(String.IsNullOrEmpty(month), DBNull.Value, month))
                    cmd.Parameters.AddWithValue("@Company", If(String.IsNullOrEmpty(company), DBNull.Value, company))
                    cmd.Parameters.AddWithValue("@Category", If(String.IsNullOrEmpty(category), DBNull.Value, category))
                    cmd.Parameters.AddWithValue("@Segment", If(String.IsNullOrEmpty(segment), DBNull.Value, segment))
                    cmd.Parameters.AddWithValue("@Brand", If(String.IsNullOrEmpty(brand), DBNull.Value, brand))
                    cmd.Parameters.AddWithValue("@Vendor", If(String.IsNullOrEmpty(vendor), DBNull.Value, vendor))
                    cmd.Parameters.AddWithValue("@Status", If(String.IsNullOrEmpty(status), DBNull.Value, status))
                    cmd.Parameters.AddWithValue("@DateFrom", If(dateFrom.HasValue, dateFrom.Value, DBNull.Value))
                    cmd.Parameters.AddWithValue("@DateTo", If(dateTo.HasValue, dateTo.Value, DBNull.Value))
                    cmd.Parameters.AddWithValue("@Version", If(String.IsNullOrEmpty(version), DBNull.Value, version))

                    Using adapter As New SqlDataAdapter(cmd)
                        adapter.Fill(dt)
                    End Using
                End Using
            End Using

        Catch ex As Exception
            Throw New Exception("Error searching approved OTB: " & ex.Message)
        End Try

        Return dt
    End Function

    ''' <summary>
    ''' Search Approved OTB
    ''' </summary>
    Public Shared Function SearchSwitchOTB(
        Optional type As String = Nothing,
        Optional year As String = Nothing,
        Optional month As String = Nothing,
        Optional company As String = Nothing,
        Optional category As String = Nothing,
        Optional segment As String = Nothing,
        Optional brand As String = Nothing,
        Optional vendor As String = Nothing) As DataTable

        Dim dt As New DataTable()

        Try
            Using conn As New SqlConnection(connectionString)
                conn.Open()

                Using cmd As New SqlCommand("SP_Search_SWitch_OTB", conn)
                    cmd.CommandType = CommandType.StoredProcedure
                    cmd.CommandTimeout = 300

                    cmd.Parameters.AddWithValue("@Type", If(String.IsNullOrEmpty(type), DBNull.Value, type))
                    cmd.Parameters.AddWithValue("@Year", If(String.IsNullOrEmpty(year), DBNull.Value, year))
                    cmd.Parameters.AddWithValue("@Month", If(String.IsNullOrEmpty(month), DBNull.Value, month))
                    cmd.Parameters.AddWithValue("@Company", If(String.IsNullOrEmpty(company), DBNull.Value, company))
                    cmd.Parameters.AddWithValue("@Category", If(String.IsNullOrEmpty(category), DBNull.Value, category))
                    cmd.Parameters.AddWithValue("@Segment", If(String.IsNullOrEmpty(segment), DBNull.Value, segment))
                    cmd.Parameters.AddWithValue("@Brand", If(String.IsNullOrEmpty(brand), DBNull.Value, brand))
                    cmd.Parameters.AddWithValue("@Vendor", If(String.IsNullOrEmpty(vendor), DBNull.Value, vendor))
                    Using adapter As New SqlDataAdapter(cmd)
                        adapter.Fill(dt)
                    End Using
                End Using
            End Using

        Catch ex As Exception
            Throw New Exception("Error searching switch OTB: " & ex.Message)
        End Try

        Return dt
    End Function


    ''' <summary>
    ''' Export Approved OTB to DataTable (สำหรับ Excel Export)
    ''' </summary>
    Public Shared Function ExportToDataTable(dt As DataTable) As DataTable
        Dim exportDt As New DataTable()

        ' สร้าง columns ตามที่ต้องการ export
        exportDt.Columns.Add("Create date", GetType(String))
        exportDt.Columns.Add("Version", GetType(String))
        exportDt.Columns.Add("Type", GetType(String))
        exportDt.Columns.Add("Year", GetType(Integer))
        exportDt.Columns.Add("Month", GetType(String))
        exportDt.Columns.Add("Category", GetType(String))
        exportDt.Columns.Add("Category name", GetType(String))
        exportDt.Columns.Add("Company", GetType(String))
        exportDt.Columns.Add("Segment", GetType(String))
        exportDt.Columns.Add("Segment name", GetType(String))
        exportDt.Columns.Add("Brand", GetType(String))
        exportDt.Columns.Add("Brand name", GetType(String))
        exportDt.Columns.Add("Vendor", GetType(String))
        exportDt.Columns.Add("Vendor name", GetType(String))
        exportDt.Columns.Add("Amount (THB)", GetType(String))
        exportDt.Columns.Add("Revised Diff", GetType(String))
        exportDt.Columns.Add("Remark", GetType(String))
        exportDt.Columns.Add("Status", GetType(String))
        exportDt.Columns.Add("Approved date", GetType(String))
        exportDt.Columns.Add("SAP date", GetType(String))
        exportDt.Columns.Add("Action by", GetType(String))

        For Each row As DataRow In dt.Rows
            Dim newRow As DataRow = exportDt.NewRow()

            newRow("Create date") = If(row("CreateDate") IsNot DBNull.Value, Convert.ToDateTime(row("CreateDate")).ToString("dd/MM/yyyy HH:mm"), "")
            newRow("Version") = If(row("Version") IsNot DBNull.Value, row("Version").ToString(), "")
            newRow("Type") = If(row("Type") IsNot DBNull.Value, row("Type").ToString(), "")
            newRow("Year") = If(row("Year") IsNot DBNull.Value, row("Year"), 0)
            newRow("Month") = GetMonthName(row("Month"))
            newRow("Category") = If(row("Category") IsNot DBNull.Value, row("Category").ToString(), "")
            newRow("Category name") = If(row("CategoryName") IsNot DBNull.Value, row("CategoryName").ToString(), "")
            newRow("Company") = If(row("Company") IsNot DBNull.Value, row("Company").ToString(), "")
            newRow("Segment") = If(row("Segment") IsNot DBNull.Value, row("Segment").ToString(), "")
            newRow("Segment name") = If(row("SegmentName") IsNot DBNull.Value, row("SegmentName").ToString(), "")
            newRow("Brand") = If(row("Brand") IsNot DBNull.Value, row("Brand").ToString(), "")
            newRow("Brand name") = If(row("BrandName") IsNot DBNull.Value, row("BrandName").ToString(), "")
            newRow("Vendor") = If(row("Vendor") IsNot DBNull.Value, row("Vendor").ToString(), "")
            newRow("Vendor name") = If(row("VendorName") IsNot DBNull.Value, row("VendorName").ToString(), "")

            Dim amount As Decimal = If(row("Amount") IsNot DBNull.Value, Convert.ToDecimal(row("Amount")), 0)
            newRow("Amount (THB)") = amount.ToString("N2")

            If row("RevisedDiff") IsNot DBNull.Value Then
                Dim revDiff As Decimal = Convert.ToDecimal(row("RevisedDiff"))
                newRow("Revised Diff") = revDiff.ToString("N2")
            Else
                newRow("Revised Diff") = ""
            End If

            newRow("Remark") = If(row("Remark") IsNot DBNull.Value, row("Remark").ToString(), "")
            newRow("Status") = If(row("OTBStatus") IsNot DBNull.Value, row("OTBStatus").ToString(), "")
            newRow("Approved date") = If(row("ApprovedDate") IsNot DBNull.Value, Convert.ToDateTime(row("ApprovedDate")).ToString("dd/MM/yyyy HH:mm"), "")
            newRow("SAP date") = If(row("SAPDate") IsNot DBNull.Value, Convert.ToDateTime(row("SAPDate")).ToString("dd/MM/yyyy HH:mm"), "")
            newRow("Action by") = If(row("ActionBy") IsNot DBNull.Value, row("ActionBy").ToString(), "")

            exportDt.Rows.Add(newRow)
        Next

        Return exportDt
    End Function

    Private Shared Function GetMonthName(month As Object) As String
        If month Is DBNull.Value Then Return ""

        Dim monthInt As Integer = Convert.ToInt32(month)
        Select Case monthInt
            Case 1 : Return "Jan"
            Case 2 : Return "Feb"
            Case 3 : Return "Mar"
            Case 4 : Return "Apr"
            Case 5 : Return "May"
            Case 6 : Return "Jun"
            Case 7 : Return "Jul"
            Case 8 : Return "Aug"
            Case 9 : Return "Sep"
            Case 10 : Return "Oct"
            Case 11 : Return "Nov"
            Case 12 : Return "Dec"
            Case Else : Return monthInt.ToString()
        End Select
    End Function

End Class