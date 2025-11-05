Imports System.Web
Imports System.Data
Imports System.Data.SqlClient
Imports System.Text
Imports Newtonsoft.Json ' ต้องมี Newtonsoft.Json ในโปรเจกต์

Public Class DataOTBRemainingHandler
    Implements IHttpHandler

    Private Shared connectionString As String = ConfigurationManager.ConnectionStrings("BMSConnectionString")?.ConnectionString

    Sub ProcessRequest(ByVal context As HttpContext) Implements IHttpHandler.ProcessRequest
        context.Response.ContentType = "application/json"
        context.Response.ContentEncoding = Encoding.UTF8

        Try
            ' 1. รับค่า Parameters จาก Form
            Dim year As Integer = 0
            Integer.TryParse(context.Request.Form("OTByear"), year)
            Dim month As Integer = 0
            Integer.TryParse(context.Request.Form("OTBmonth"), month)
            Dim company As String = If(context.Request.Form("OTBCompany"), "")
            Dim category As String = If(context.Request.Form("OTBCategory"), "")
            Dim segment As String = If(context.Request.Form("OTBSegment"), "")
            Dim brand As String = If(context.Request.Form("OTBBrand"), "")
            Dim vendor As String = If(context.Request.Form("OTBVendor"), "")

            ' 2. ตรวจสอบว่ามี Parameter ครบหรือไม่
            If year = 0 OrElse month = 0 OrElse String.IsNullOrEmpty(company) OrElse String.IsNullOrEmpty(category) OrElse String.IsNullOrEmpty(segment) OrElse String.IsNullOrEmpty(brand) OrElse String.IsNullOrEmpty(vendor) Then
                Throw New Exception("All filter fields are required to view the report.")
            End If

            Dim ds As New DataSet()
            Using conn As New SqlConnection(connectionString)
                conn.Open()
                ' 3. เรียก Stored Procedure
                Using cmd As New SqlCommand("SP_Get_OTB_Remaining_Report", conn)
                    cmd.CommandType = CommandType.StoredProcedure
                    cmd.Parameters.AddWithValue("@Year", year)
                    cmd.Parameters.AddWithValue("@Month", month)
                    cmd.Parameters.AddWithValue("@Company", company)
                    cmd.Parameters.AddWithValue("@Category", category)
                    cmd.Parameters.AddWithValue("@Segment", segment)
                    cmd.Parameters.AddWithValue("@Brand", brand)
                    cmd.Parameters.AddWithValue("@Vendor", vendor)

                    Using adapter As New SqlDataAdapter(cmd)
                        ' 4. ดึงข้อมูลทั้ง 2 Result Sets
                        adapter.Fill(ds)
                    End Using
                End Using
            End Using

            ' 5. ตรวจสอบว่าได้ข้อมูล 2 ตารางกลับมา
            If ds.Tables.Count < 2 Then
                Throw New Exception("Stored procedure did not return the expected data (Detail and OtherRemaining).")
            End If

            ' 6. ตั้งชื่อตารางเพื่อง่ายต่อการ Serialize
            ds.Tables(0).TableName = "detail"
            ds.Tables(1).TableName = "otherRemaining"

            ' 7. Serialize DataSet ทั้งหมด (ที่มี 2 ตาราง) กลับไปเป็น JSON
            Dim jsonResult As String = JsonConvert.SerializeObject(ds, Formatting.None)
            context.Response.Write(jsonResult)

        Catch ex As Exception
            context.Response.StatusCode = 500
            ' ส่ง Error กลับไปเป็น JSON
            Dim errorResponse As New With {
                .success = False,
                .message = ex.Message
            }
            context.Response.Write(JsonConvert.SerializeObject(errorResponse))
        End Try
    End Sub

    ReadOnly Property IsReusable() As Boolean Implements IHttpHandler.IsReusable
        Get
            Return False
        End Get
    End Property

End Class